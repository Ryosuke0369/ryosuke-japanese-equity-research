"""screener/fetch/jquants_universe.py - ユニバース構築 (仕様書 §2-3).

2つの入力源を持つ。どちらを使ったかは companies.source に必ず残る
(暗黙の切り替えはしない —— リポジトリ規約「サイレントフォールバック禁止」)。

  --source jquants   J-Quants API **V2**。上場一覧 + 株価/売買代金 + 時価総額。
                     時価総額と平均売買代金が取れるので **ユニバース条件を
                     完全に判定できる**。無料プランは遅延配信だが、規模・流動性の
                     判定には仕様書 §2-3 のとおり許容。
                     認証は .env の JQUANTS_API_KEY を x-api-key ヘッダーで送る。

  --source jpx       JPX「東証上場銘柄一覧」(data_j.xls)。コード・銘柄名・市場区分・
                     33業種が取れる。**時価総額と売買代金は無い**ので、規模・流動性の
                     判定はできない。該当銘柄は universe_flag=0 のまま
                     exclude_reason='規模・流動性データ未取得' が入る ——
                     「条件を満たさない」ではなく「まだ判定していない」を区別する。

Usage
    python -m screener.fetch.jquants_universe --check-auth
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
import time
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

# V1 (https://api.jquants.com/v1) は 2026-06-01 に終了した。全エンドポイントが
# 403 を返していたのはそのため。V2 はダッシュボード発行の API キーを x-api-key
# で送る方式で、refreshToken/idToken の二段認証もトークンキャッシュも無い。
JQ = "https://api.jquants.com/v2"
EP_MASTER = "/equities/master"          # V1: /listed/info
EP_BARS_DAILY = "/equities/bars/daily"  # V1: /prices/daily_quotes
JPX_XLS = ("https://www.jpx.co.jp/markets/statistics-equities/misc/"
           "tvdivq0000001vg2-att/data_j.xls")
SOURCE = "jquants"

# 商品区分コード (ProdCat)。V1 の MarketCodeName は「プライム（内国株式）」
# 「ETF・ETN」のように市場区分と商品種別の合成だったが、V2 は MktNm が
# 「プライム」等の純粋な市場名になり、商品種別は ProdCat に分離された。
# universe_rules.yaml の exclude_sectors.by_market_name は市場名で除外するので、
# 内国株券以外は商品種別を市場名に併記して復元する（REIT/ETF を取りこぼさない）。
PRODUCT_CATEGORY = {
    "011": "内国株券",
    "012": "優先出資証券",
    "013": "REIT",
    "014": "ETF",
    "021": "外国株券",
    "022": "外国REIT",
    "023": "外国ETF",
    "024": "外国株預託証券",
}


# ------------------------------------------------------------------ jquants
def jq_api_key() -> str:
    """.env の JQUANTS_API_KEY。V2 はこれ一本で、更新も失効管理も要らない。

    ダッシュボードからのコピペで前後の空白や折り返しの改行が混ざることがある。
    そのまま送ると 403 になり「キーが間違っている」と誤診するので、空白類は
    すべて落としてから使う（API キーに空白は含まれない）。
    """
    C.load_env()
    raw = os.environ.get("JQUANTS_API_KEY") or ""
    key = "".join(raw.split())
    if not key:
        raise SystemExit(
            "ERROR: no J-Quants API key.\n"
            "  J-Quants API は V2 へ移行し、V1 (mail+password -> refreshToken ->\n"
            "  idToken) は 2026-06-01 に終了しました。ダッシュボードで API キーを\n"
            "  発行し、.env に次の 1 行を追加してください:\n"
            "      JQUANTS_API_KEY=...\n"
            "  発行元: https://jpx-jquants.com/ja/dashboard")
    if key != raw.strip():
        C.log("J-Quants: API キーに空白/改行が含まれていたため除去しました")
    return key


def jq_auth(fetcher) -> None:
    """セッションに x-api-key を載せる。

    Authorization ヘッダーとの同時送信は不可（V1 の Bearer が残っていると V2 は
    401/403 を返す）。前段の処理が付けている可能性を考えて明示的に落とす。
    """
    fetcher.s.headers["x-api-key"] = jq_api_key()
    for stale in ("Authorization", "authorization"):
        fetcher.s.headers.pop(stale, None)


def _jq_diagnose(path: str, r) -> str:
    """HTTP ステータスごとに原因を切り分けたメッセージを組み立てる。

    403 を一律「キーが不正」と出すと、実際にはプラン外のエンドポイントを叩いて
    いるだけのときに .env を疑って時間を溶かす —— 今回まさにそれをやった。
    """
    try:
        msg = (r.json() or {}).get("message") or r.text[:200]
    except ValueError:
        msg = r.text[:200]
    hint = {
        400: ("パラメータ誤り —— 必須パラメータの欠落か書式違反。"
              "V2 の日付は YYYYMMDD もしくは YYYY-MM-DD。"),
        401: ("認証ヘッダーを解釈できない —— V1 の Authorization: Bearer が"
              "残っていないか確認する。V2 は x-api-key のみ。"),
        403: ("キー不正 or 権限外 —— .env の JQUANTS_API_KEY をダッシュボードの"
              "発行値と照合する。キーが正しい場合は、契約プランで許可されていない"
              "エンドポイント/期間か、パスの綴り違い。"),
        429: ("レート制限超過 —— 無料プランは 5 req/min（Light 60 / Standard 120"
              " / Premium 500）。--rpm を実際のプランに合わせる。"),
        500: "J-Quants 側の一時障害。時間を置いて再実行する。",
    }.get(r.status_code, "想定外のステータス。")
    return f"{path}: HTTP {r.status_code} — {hint} (message: {msg})"


def jq_get(fetcher, path: str, **params):
    """Paginated GET. J-Quants returns `pagination_key` when more data exists;
    ignoring it silently truncates the universe, which is exactly the kind of
    quiet loss this project refuses.

    V2 は本体を必ず `data` 配列で返す（V1 はエンドポイント毎に別名だった）。
    429 だけは待てば通るので、その場で数回まで待ち直す。
    """
    out, key, waits = [], None, 0
    while True:
        p = dict(params)
        if key:
            p["pagination_key"] = key
        r = fetcher.get(f"{JQ}{path}", params=p,
                        allow_status=(200, 400, 401, 403, 404, 429))
        if r.status_code == 429 and waits < 8:
            waits += 1
            # 待つだけでは同じ壁に当たり続ける（1分窓を滑らせているので、
            # 公称レートぎりぎりの間隔だと窓の境界で必ず弾かれる）。待つのと
            # 同時に以後のペースそのものを落とし、実測に合わせて自動収束させる。
            fetcher.throttle.min_interval *= 1.3
            C.log(f"  {_jq_diagnose(path, r)}")
            C.log(f"  65 秒待機し、以後の間隔を "
                  f"{fetcher.throttle.min_interval:.1f} 秒へ広げて再試行 ({waits}/8)")
            time.sleep(65)
            continue
        if r.status_code != 200:
            raise RuntimeError(_jq_diagnose(path, r))
        d = r.json()
        out.extend(d.get("data") or [])
        key = d.get("pagination_key")
        if not key:
            return out


def _market_label(row) -> str | None:
    """MktNm と ProdCat から、除外ルールが読む市場区分名を組み立てる。"""
    market = _s(row.get("MktNm"))
    product = PRODUCT_CATEGORY.get(_s(row.get("ProdCat")) or "")
    if product in (None, "内国株券"):
        return market
    return f"{market}({product})" if market else product


def build_from_jquants(con, fetcher, price_days: int = 25) -> dict:
    jq_auth(fetcher)

    C.log(f"J-Quants V2: {EP_MASTER}")
    info = jq_get(fetcher, EP_MASTER)
    C.log(f"  {len(info)} listed rows")

    # 直近の取得可能日から遡って price_days 営業日ぶんの日次四本値を集める。
    # 無料プランは遅延配信なので、今日から12週+α遡った日を起点にする。
    end = date.today() - timedelta(weeks=12) - timedelta(days=2)
    quotes: dict[str, list] = {}
    d, got = end, 0
    C.log(f"J-Quants V2: {EP_BARS_DAILY} ({price_days} 営業日ぶん、"
          f"{end.isoformat()} から遡行)")
    while got < price_days and (end - d).days < price_days * 2 + 30:
        if d.weekday() < 5:
            rows = jq_get(fetcher, EP_BARS_DAILY, date=d.strftime("%Y%m%d"))
            if rows:
                got += 1
                for r in rows:
                    quotes.setdefault(r["Code"], []).append(r)
                C.log(f"  [{got}/{price_days}] {d.isoformat()} {len(rows)} rows "
                      f"(interval {fetcher.throttle.min_interval:.1f}s, "
                      f"{fetcher.n_requests} req)")
            else:
                C.log(f"  [{got}/{price_days}] {d.isoformat()} 休場")
        d -= timedelta(days=1)
    C.log(f"  {got} trading day(s) of quotes for {len(quotes)} code(s) "
          f"(ending {end.isoformat()}, delayed feed)")

    rules = C.load_yaml("universe_rules.yaml")
    liq = rules["liquidity"]
    # ADV は「20日平均」。price_days は休場日で取りこぼさないための余分な遡行なので、
    # 平均そのものは新しい方から adv_window_days 日ぶんだけで取る。
    window = int(liq.get("adv_window_days") or 20)
    n = 0
    for row in info:
        code = C.normalise_code(row.get("Code"))
        qs = quotes.get(row.get("Code"), [])
        closes = [q for q in qs if q.get("C")]          # V1: Close
        adv = None
        if len(closes) >= liq["min_observations"]:
            vals = [(q.get("Va") or 0) for q in closes[:window]]  # V1: TurnoverValue
            adv = sum(vals) / len(vals) / 1_000_000     # JPY -> JPY mn
        latest = closes[0] if closes else None
        close = latest.get("C") if latest else None
        # MktCap は V2 で日次バーに入った項目 (JPY mn)。V1 では取得できず、
        # 全銘柄が exclude_reason='時価総額未取得' で滞留していた。
        mktcap = latest.get("MktCap") if latest else None
        upsert_company(con, code, row.get("CoName"),
                       _market_label(row), row.get("S33Nm"),
                       sector17=row.get("S17Nm"),
                       scale=row.get("ScaleCat"),
                       mktcap=mktcap, adv20=adv, source="jquants")
        if close and adv is not None:
            con.execute(
                "INSERT OR REPLACE INTO prices (code, date, close, volume, adv20) "
                "VALUES (?,?,?,?,?)",
                (code, end.isoformat(), close, latest.get("Vo"), adv))
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


def is_non_common_share(code, name, share_rules) -> bool:
    """普通株以外の株式種別（優先株式等）か。規則は universe_rules.yaml の
    exclude_share_classes。

    証券コード5桁目は株式の種類で 0 が普通株。normalise_code は末尾0だけを
    落とすので、正規化後も5桁のコードは普通株ではない。
    """
    if share_rules.get("non_common_code") and len(code or "") == 5:
        return True
    return any(p in (name or "") for p in (share_rules.get("by_name_pattern") or []))


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
    share = rules.get("exclude_share_classes") or {}

    counts = {"universe": 0, "excluded": 0, "pending": 0}
    reasons: dict[str, int] = {}
    for row in con.execute("SELECT code, name, market, sector, mktcap, adv20 "
                           "FROM companies"):
        code, market = row["code"], row["market"] or ""
        reason = None
        # 大文字化して照合する。同じ市場が 'TOKYO PRO MARKET' と 'PRO Market'
        # の両方の綴りで DB に入っており、素の in では後者を取りこぼして
        # 「除外」ではなく「未判定」に落ちる —— この2状態の取り違えが一番まずい。
        market_u = market.upper()
        if any(m.upper() in market_u for m in bad_markets):
            reason = f"市場区分除外({market})"
        elif is_non_common_share(code, row["name"], share):
            # 商品種別除外と同じ扱い。優先株式は MktNm が「プライム」のままで
            # 市場名の除外に掛からない（2026-09-19、25935 伊藤園で判明）。
            reason = "株式種別除外(優先株式等)"
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
    # ADV の窓は liquidity.adv_window_days (20日)。ここはその 20 日を確実に
    # 埋めるための遡行幅で、平均の窓ではない。V2 でリクエストが希少資源に
    # なった（無料プラン 5 req/min = 1日1リクエスト）ので、40 → 25 に絞る。
    # 25 営業日あれば新しい方から 20 日を取っても余りが出る。
    p.add_argument("--price-days", type=int, default=25)
    p.add_argument("--rpm", type=int, default=5,
                   help="J-Quants のレートリミット (req/min)。"
                        "Free 5 / Light 60 / Standard 120 / Premium 500")
    p.add_argument("--check-auth", action="store_true",
                   help="J-Quants の認証だけ試して終了する")
    a = p.parse_args(argv)

    # V2 のレートリミットは分あたり。無料プランの 5 req/min で 0.4 秒間隔を
    # 続ければ即 429 になるので、間隔はプランから逆算する。
    # マージンは +30%。1分窓が滑る実装なので公称値ちょうど(12.0s)や +10%(13.2s)
    # では窓の境界で必ず弾かれ、429 ごとに 65 秒失って実効速度がかえって落ちる
    # ——実測で確認済み。足りなければ jq_get が 429 を見て自動で更に広げる。
    interval = 60.0 / max(a.rpm, 1) * 1.3

    con = C.init_db()
    if a.check_auth:
        fetcher = C.Fetcher(min_interval=interval)
        jq_auth(fetcher)
        rows = jq_get(fetcher, EP_MASTER, code="86970")
        if not rows:
            raise SystemExit("ERROR: 認証は通ったが /equities/master が 0 件を"
                             "返した。契約プランを確認する。")
        r = rows[0]
        C.log(f"J-Quants V2 auth OK — {EP_MASTER} code=86970 -> "
              f"{r.get('Code')} {r.get('CoName')} / {r.get('MktNm')} / "
              f"{r.get('S33Nm')} (as of {r.get('Date')})")
        return 0
    if a.build:
        run_id = C.start_run(con, f"universe_{a.source}", date.today().isoformat())
        fetcher = C.Fetcher(min_interval=interval if a.source == "jquants" else 0.4)
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
