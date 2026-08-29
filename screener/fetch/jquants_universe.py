"""screener/fetch/jquants_universe.py - ユニバース構築 (仕様書 §2-3).

2つの入力源を持つ。どちらを使ったかは companies.source に必ず残る
(暗黙の切り替えはしない —— リポジトリ規約「サイレントフォールバック禁止」)。

  --source jquants   J-Quants API。上場一覧 + 株価/売買代金 + 発行済株式数。
                     時価総額と20日平均売買代金が取れるので **ユニバース条件を
                     完全に判定できる**。無料プランは12週遅延だが、規模・流動性の
                     判定には仕様書 §2-3 のとおり許容。
                     認証は .env の JQUANTS_MAIL / JQUANTS_PASSWORD から
                     refreshToken → idToken を自動更新する(トークンは
                     data/cache/ にキャッシュ。refresh 1週間 / id 24時間)。
                     貼り付け済みの JQUANTS_REFRESH_TOKEN があればそれも使えるが、
                     1週間で失効し自動更新できないため補助扱い。

  --source jpx       JPX「東証上場銘柄一覧」(data_j.xls)。コード・銘柄名・市場区分・
                     33業種が取れる。**時価総額と売買代金は無い**ので、規模・流動性の
                     判定はできない。該当銘柄は universe_flag=0 のまま
                     exclude_reason='規模・流動性データ未取得' が入る ——
                     「条件を満たさない」ではなく「まだ判定していない」を区別する。

Usage
    python -m screener.fetch.jquants_universe --source jquants --build
    python -m screener.fetch.jquants_universe --source jpx --build
    python -m screener.fetch.jquants_universe --report
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sys
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

JQ = "https://api.jquants.com/v1"
JPX_XLS = ("https://www.jpx.co.jp/markets/statistics-equities/misc/"
           "tvdivq0000001vg2-att/data_j.xls")
SOURCE = "jquants"


# ------------------------------------------------------------------ jquants
TOKEN_CACHE = os.path.join(C.CACHE_DIR, "jquants_token.json")

# J-Quants の有効期間。refresh token 1週間 / id token 24時間。
# 週次運用でトークンを手で貼り替え続けるのは続かないので、mail+password から
# 自動更新する経路を主とし、貼り付け済みの refresh token は補助に回す。
REFRESH_TTL_H = 24 * 6          # 7日の手前で取り直す
ID_TTL_H = 20                   # 24時間の手前で取り直す


def _cache_read() -> dict:
    import json
    try:
        with open(TOKEN_CACHE, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def _cache_write(d: dict) -> None:
    """Tokens are credentials. They live under screener/data/cache/, which is
    git-ignored, and the file is created 0600 where the OS honours it."""
    import json
    os.makedirs(C.CACHE_DIR, exist_ok=True)
    tmp = TOKEN_CACHE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(d, fh)
    try:
        os.chmod(tmp, 0o600)
    except OSError:
        pass
    os.replace(tmp, TOKEN_CACHE)


def _age_hours(iso: str | None) -> float:
    from datetime import datetime
    if not iso:
        return 1e9
    try:
        return (datetime.now() - datetime.fromisoformat(iso)).total_seconds() / 3600
    except ValueError:
        return 1e9


def jq_refresh_token(force: bool = False) -> str:
    """mail+password → refreshToken, cached.

    Falls back to a hand-pasted JQUANTS_REFRESH_TOKEN only when no credentials
    are configured — and says so, rather than silently using a stale token.
    """
    import requests
    C.load_env()
    cache = _cache_read()
    if not force and cache.get("refresh_token") and \
            _age_hours(cache.get("refresh_at")) < REFRESH_TTL_H:
        return cache["refresh_token"]

    mail = os.environ.get("JQUANTS_MAIL")
    password = os.environ.get("JQUANTS_PASSWORD")
    if mail and password:
        r = requests.post(f"{JQ}/token/auth_user",
                          json={"mailaddress": mail, "password": password},
                          timeout=60)
        if r.status_code != 200:
            raise SystemExit(
                f"ERROR: J-Quants auth_user returned HTTP {r.status_code} "
                f"({r.text[:160]}).\n"
                f"  JQUANTS_MAIL / JQUANTS_PASSWORD in .env were rejected. "
                f"Check them at https://application.jpx-jquants.com/ .")
        token = r.json()["refreshToken"]
        from datetime import datetime
        cache.update({"refresh_token": token,
                      "refresh_at": datetime.now().isoformat()})
        _cache_write(cache)
        C.log("J-Quants: refresh token renewed from JQUANTS_MAIL/PASSWORD")
        return token

    token = os.environ.get("JQUANTS_REFRESH_TOKEN")
    if token:
        C.log("J-Quants: using the pasted JQUANTS_REFRESH_TOKEN "
              "(JQUANTS_MAIL/JQUANTS_PASSWORD are not set, so it cannot be "
              "auto-renewed and will stop working within a week)")
        return token

    raise SystemExit(
        "ERROR: no J-Quants credentials.\n"
        "  Preferred (auto-renewing) — add to .env:\n"
        "      JQUANTS_MAIL=you@example.com\n"
        "      JQUANTS_PASSWORD=...\n"
        "  Or paste a refresh token (expires after one week):\n"
        "      JQUANTS_REFRESH_TOKEN=...\n"
        "  Sign-up / credentials: https://application.jpx-jquants.com/")


def jq_id_token(fetcher: "C.Fetcher | None" = None, force: bool = False) -> str:
    """refreshToken → idToken, cached for ID_TTL_H hours.

    A 403 on auth_refresh means the refresh token is dead; with mail+password
    configured we take one automatic retry with a freshly minted refresh token
    before giving up, because a week-old token expiring is the normal case, not
    an exceptional one.
    """
    import requests
    from datetime import datetime

    cache = _cache_read()
    if not force and cache.get("id_token") and \
            _age_hours(cache.get("id_at")) < ID_TTL_H:
        return cache["id_token"]

    for attempt in (1, 2):
        token = jq_refresh_token(force=(attempt == 2))
        r = requests.post(f"{JQ}/token/auth_refresh",
                          params={"refreshtoken": token}, timeout=60)
        if r.status_code == 200:
            idt = r.json()["idToken"]
            cache.update({"id_token": idt, "id_at": datetime.now().isoformat()})
            _cache_write(cache)
            return idt
        if attempt == 1 and os.environ.get("JQUANTS_MAIL"):
            C.log(f"J-Quants: auth_refresh HTTP {r.status_code} — the cached "
                  f"refresh token is stale; re-authenticating with "
                  f"JQUANTS_MAIL/PASSWORD")
            continue
        raise SystemExit(
            f"ERROR: J-Quants auth_refresh returned HTTP {r.status_code} "
            f"({r.text[:160]}).\n"
            f"  The refresh token is invalid or expired (they last one week).\n"
            f"  Set JQUANTS_MAIL / JQUANTS_PASSWORD in .env so the token can be "
            f"renewed automatically, or paste a fresh JQUANTS_REFRESH_TOKEN.")
    raise SystemExit("unreachable")


def jq_get(fetcher, path: str, token: str, **params):
    """Paginated GET. J-Quants returns `pagination_key` when more data exists;
    ignoring it silently truncates the universe, which is exactly the kind of
    quiet loss this project refuses."""
    out, key = [], None
    while True:
        p = dict(params)
        if key:
            p["pagination_key"] = key
        r = fetcher.get(f"{JQ}{path}", params=p,
                        allow_status=(200, 400, 401, 403, 404))
        if r.status_code != 200:
            raise RuntimeError(f"{path}: HTTP {r.status_code} {r.text[:160]}")
        d = r.json()
        body = next((v for k, v in d.items() if isinstance(v, list)), [])
        out.extend(body)
        key = d.get("pagination_key")
        if not key:
            return out


def build_from_jquants(con, fetcher, price_days: int = 40) -> dict:
    token = jq_id_token()
    fetcher.s.headers["Authorization"] = f"Bearer {token}"

    C.log("J-Quants: /listed/info")
    info = jq_get(fetcher, "/listed/info", token)
    C.log(f"  {len(info)} listed rows")

    # 直近の取得可能日から遡って price_days 営業日ぶんの日次four本値を集める。
    # 無料プランは12週遅延なので、今日から12週+α遡った日を起点にする。
    end = date.today() - timedelta(weeks=12) - timedelta(days=2)
    quotes: dict[str, list] = {}
    d, got = end, 0
    while got < price_days and (end - d).days < price_days * 2 + 30:
        if d.weekday() < 5:
            rows = jq_get(fetcher, "/prices/daily_quotes", token, date=d.isoformat())
            if rows:
                got += 1
                for r in rows:
                    quotes.setdefault(r["Code"], []).append(r)
        d -= timedelta(days=1)
    C.log(f"  {got} trading day(s) of quotes for {len(quotes)} code(s) "
          f"(ending {end.isoformat()}, free-plan 12-week delay)")

    rules = C.load_yaml("universe_rules.yaml")
    n = 0
    for row in info:
        code = C.normalise_code(row.get("Code"))
        qs = quotes.get(row.get("Code"), [])
        closes = [q for q in qs if q.get("Close")]
        adv = None
        if len(closes) >= rules["liquidity"]["min_observations"]:
            vals = [(q.get("TurnoverValue") or 0) for q in closes]
            adv = sum(vals) / len(vals) / 1_000_000        # JPY -> JPY mn
        close = closes[0].get("Close") if closes else None
        upsert_company(con, code, row.get("CompanyName"),
                       row.get("MarketCodeName"), row.get("Sector33CodeName"),
                       sector17=row.get("Sector17CodeName"),
                       scale=row.get("ScaleCategory"),
                       mktcap=None, adv20=adv, source="jquants")
        if close and adv is not None:
            con.execute(
                "INSERT OR REPLACE INTO prices (code, date, close, volume, adv20) "
                "VALUES (?,?,?,?,?)",
                (code, end.isoformat(), close, closes[0].get("Volume"), adv))
        n += 1
    con.commit()
    return {"rows": n, "source": "jquants"}


# ---------------------------------------------------------------------- jpx
def build_from_jpx(con, fetcher) -> dict:
    """東証上場銘柄一覧(data_j.xls)。時価総額・売買代金は含まれない。"""
    import pandas as pd
    C.log(f"JPX: {JPX_XLS}")
    r = fetcher.get(JPX_XLS)
    if r.status_code != 200:
        raise SystemExit(f"ERROR: JPX list returned HTTP {r.status_code}")
    df = pd.read_excel(io.BytesIO(r.content))
    asof = str(df["日付"].iloc[0]) if "日付" in df.columns else "?"
    C.log(f"  {len(df)} rows, as of {asof}")

    n = 0
    for _, row in df.iterrows():
        code = C.normalise_code(row.get("コード"))
        if not code:
            continue
        upsert_company(con, code, row.get("銘柄名"), row.get("市場・商品区分"),
                       row.get("33業種区分"), sector17=row.get("17業種区分"),
                       scale=row.get("規模区分"), mktcap=None, adv20=None,
                       source="jpx")
        n += 1
    con.commit()
    return {"rows": n, "source": "jpx", "asof": asof}


# ------------------------------------------------------------------- shared
def upsert_company(con, code, name, market, sector, *, sector17, scale,
                   mktcap, adv20, source) -> None:
    """COALESCE everywhere: the TDnet archiver already created rows with a name
    only, and a later J-Quants run must not blank out a field it happens not to
    carry."""
    con.execute(
        "INSERT INTO companies (code, name, market, sector, mktcap, adv20, "
        " sector17, scale_category, source, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?) "
        "ON CONFLICT(code) DO UPDATE SET "
        " name=COALESCE(excluded.name, companies.name), "
        " market=COALESCE(excluded.market, companies.market), "
        " sector=COALESCE(excluded.sector, companies.sector), "
        " mktcap=COALESCE(excluded.mktcap, companies.mktcap), "
        " adv20=COALESCE(excluded.adv20, companies.adv20), "
        " sector17=COALESCE(excluded.sector17, companies.sector17), "
        " scale_category=COALESCE(excluded.scale_category, companies.scale_category), "
        " source=excluded.source, updated_at=excluded.updated_at",
        (code, _s(name), _s(market), _s(sector), mktcap, adv20, _s(sector17),
         _s(scale), source, C.utcnow()))


def _s(v):
    if v is None:
        return None
    s = str(v).strip()
    return None if s in ("", "-", "nan") else s


def apply_universe_rules(con) -> dict:
    """universe_flag と exclude_reason を評価する。

    「条件を満たさない」と「判定に必要なデータがまだ無い」を必ず区別する。
    後者を除外扱いにすると、J-Quantsが繋がった日に何社増えるのかが分からなくなる。
    """
    rules = C.load_yaml("universe_rules.yaml")
    size, liq, exc = rules["size"], rules["liquidity"], rules["exclude_sectors"]
    bad_sectors = set(exc.get("by_name") or [])
    bad_markets = tuple(exc.get("by_market_name") or [])
    patterns = [re.compile(p) for p in (rules.get("exclude_code_patterns") or [])]

    counts = {"universe": 0, "excluded": 0, "pending": 0}
    reasons: dict[str, int] = {}
    for row in con.execute("SELECT code, market, sector, mktcap, adv20 FROM companies"):
        code, market = row["code"], row["market"] or ""
        reason = None
        if any(m in market for m in bad_markets):
            reason = f"市場区分除外({market})"
        elif (row["sector"] or "") in bad_sectors:
            reason = f"業種除外({row['sector']})"
        elif not market and any(p.match(code) for p in patterns):
            reason = "コード帯除外(市場区分不明のETF/REIT帯)"
        elif row["mktcap"] is None and row["adv20"] is None:
            reason = "規模・流動性データ未取得"          # ← 未判定。除外ではない
        elif row["mktcap"] is not None and not (
                size["mktcap_min_mn"] <= row["mktcap"] <= size["mktcap_max_mn"]):
            reason = "時価総額レンジ外"
        elif row["adv20"] is not None and row["adv20"] < liq["adv20_min_mn"]:
            reason = "流動性不足"
        elif row["mktcap"] is None:
            reason = "時価総額未取得"

        flag = 0 if reason else 1
        con.execute("UPDATE companies SET universe_flag=?, exclude_reason=? "
                    "WHERE code=?", (flag, reason, code))
        if flag:
            counts["universe"] += 1
        elif reason and ("未取得" in reason):
            counts["pending"] += 1
            reasons[reason] = reasons.get(reason, 0) + 1
        else:
            counts["excluded"] += 1
            reasons[reason] = reasons.get(reason, 0) + 1
    con.commit()
    counts["reasons"] = reasons
    return counts


def report(con) -> None:
    tot = con.execute("SELECT COUNT(*) c FROM companies").fetchone()["c"]
    uni = con.execute("SELECT COUNT(*) c FROM companies WHERE universe_flag=1"
                      ).fetchone()["c"]
    with_cap = con.execute("SELECT COUNT(*) c FROM companies WHERE mktcap IS NOT NULL"
                           ).fetchone()["c"]
    with_adv = con.execute("SELECT COUNT(*) c FROM companies WHERE adv20 IS NOT NULL"
                           ).fetchone()["c"]
    C.log(f"companies: {tot} rows / universe_flag=1: {uni} "
          f"/ mktcap present: {with_cap} / adv20 present: {with_adv}")
    C.log("by source:")
    for r in con.execute("SELECT source, COUNT(*) c FROM companies "
                         "GROUP BY source ORDER BY c DESC"):
        C.log(f"  {str(r['source']):<12} {r['c']:>6}")
    C.log("exclude_reason:")
    for r in con.execute("SELECT exclude_reason, COUNT(*) c FROM companies "
                         "WHERE universe_flag=0 GROUP BY exclude_reason "
                         "ORDER BY c DESC LIMIT 15"):
        C.log(f"  {str(r['exclude_reason']):<40} {r['c']:>6}")
    C.log("market breakdown (universe candidates before size/liquidity):")
    for r in con.execute(
            "SELECT market, COUNT(*) c FROM companies "
            "WHERE exclude_reason IS NULL OR exclude_reason LIKE '%未取得%' "
            "GROUP BY market ORDER BY c DESC LIMIT 10"):
        C.log(f"  {str(r['market']):<32} {r['c']:>6}")


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="universe builder (仕様書 §2-3)")
    p.add_argument("--source", choices=("jquants", "jpx"), default="jquants")
    p.add_argument("--build", action="store_true")
    p.add_argument("--report", action="store_true")
    p.add_argument("--price-days", type=int, default=40)
    p.add_argument("--check-auth", action="store_true",
                   help="J-Quants の認証だけ試して終了する")
    a = p.parse_args(argv)

    con = C.init_db()
    if a.check_auth:
        tok = jq_id_token(force=True)
        C.log(f"J-Quants auth OK (idToken {len(tok)} chars)")
        return 0
    if a.build:
        run_id = C.start_run(con, f"universe_{a.source}", date.today().isoformat())
        fetcher = C.Fetcher(min_interval=0.4)
        try:
            res = (build_from_jquants(con, fetcher, a.price_days)
                   if a.source == "jquants" else build_from_jpx(con, fetcher))
        except SystemExit as e:
            C.finish_run(con, run_id, "failed", error=str(e)[:400])
            C.log(f"universe build FAILED and was recorded in fetch_runs: {e}")
            raise
        counts = apply_universe_rules(con)
        C.finish_run(con, run_id, "ok", n_listed=res["rows"],
                     n_saved=counts["universe"],
                     note=f"source={res['source']} pending={counts['pending']}")
        C.log(f"universe: {counts['universe']} in / {counts['excluded']} excluded "
              f"/ {counts['pending']} pending (data not yet fetched)")

    if a.report or a.build:
        report(con)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
