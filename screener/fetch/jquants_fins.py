"""screener/fetch/jquants_fins.py — J-Quants 決算サマリー（/fins/summary）を5年ぶん取り込む。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウI-v2」I2-2。

なぜ要るのか
------------
進捗率（累計 ÷ 通期着地）を**過去5年ぶん**作るには、四半期の累計と通期の着地が要る。
本体の `quarterly_standalone_all` は営業利益の入っている行が46%しかなく、通期着地が
5年そろう銘柄は6社しかない（calibration_backlog §50）。決算サマリーは1行に
「その期の累計（`OP`）」「通期予想（`FOP`）」「開示日時」を持ち、Light の5年ローリングで取れる。

既存の `cache/jquants_fins_summary/`（7列だけのキャッシュ・`earnings_window` が使う）には
**触らない**。こちらは値を持つ別系統として本体DBの `fins_summary` に入れる。

    python -m screener.fetch.jquants_fins --years 5
    python -m screener.fetch.jquants_fins --from 2026-01-01 --to 2026-09-18
    python -m screener.fetch.jquants_fins --report
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.fetch.jquants_universe import JQ, jq_auth

SOURCE = "jquants_fins"
EP = "/fins/summary"

# 保存する列。**全111列は保存しない**（5年で1GB 超になる）。
TEXT_COLS = ("Code", "DiscDate", "DiscTime", "DocType", "CurPerType",
             "CurPerSt", "CurPerEn", "CurFYEn")
NUM_COLS = ("Sales", "OP", "OdP", "NP",          # 期首からの累計（実績）
            "FSales", "FOP",                      # 通期予想（連結）
            "FNCSales", "FNCOP",                  # 通期予想（単体）
            "NxFSales", "NxFOP")                  # 翌期予想

DDL = """
CREATE TABLE IF NOT EXISTS fins_summary (
    code        TEXT NOT NULL,        -- 正規化後（4桁 or 3桁+英字）
    disc_date   TEXT NOT NULL,
    doc_type    TEXT NOT NULL,
    disc_time   TEXT,
    per_type    TEXT,                 -- 1Q / 2Q / 3Q / FY
    per_start   TEXT,
    per_end     TEXT,
    fy_end      TEXT,
    sales REAL, op REAL, odp REAL, np REAL,
    f_sales REAL, f_op REAL,
    fnc_sales REAL, fnc_op REAL,
    nxf_sales REAL, nxf_op REAL,
    fetched_at  TEXT NOT NULL,
    PRIMARY KEY (code, disc_date, doc_type)
);
CREATE INDEX IF NOT EXISTS ix_fins_code_fy ON fins_summary (code, fy_end, per_type);
CREATE INDEX IF NOT EXISTS ix_fins_date ON fins_summary (disc_date);
"""


def _num(v):
    """空文字は None。**0 と欠測を混ぜない。**"""
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def to_row(r: dict, fetched_at: str):
    code = C.normalise_code(r.get("Code"))
    if not code or not r.get("DiscDate") or not r.get("DocType"):
        return None
    vals = [code, r["DiscDate"], r["DocType"], r.get("DiscTime"), r.get("CurPerType"),
            r.get("CurPerSt"), r.get("CurPerEn"), r.get("CurFYEn")]
    vals += [_num(r.get(k)) for k in NUM_COLS]
    vals.append(fetched_at)
    return tuple(vals)


def covered_days(con) -> set:
    return {r[0] for r in con.execute("SELECT DISTINCT disc_date FROM fins_summary")}


def fetch_day(fetcher, d: date) -> list:
    """1営業日ぶん。ページングは pagination_key で辿る。"""
    out, pk = [], None
    while True:
        params = {"date": d.isoformat()}
        if pk:
            params["pagination_key"] = pk
        for _ in range(4):
            r = fetcher.get(JQ + EP, params=params, allow_status=(200, 400, 403, 429))
            if r.status_code != 429:
                break
            time.sleep(20)
        if r.status_code != 200:
            raise RuntimeError("HTTP %s %s" % (r.status_code, (r.text or "")[:120]))
        j = r.json()
        out += j.get("data") or []
        pk = j.get("pagination_key")
        if not pk:
            return out


def backfill(con, fetcher, start: date, end: date, force=False) -> dict:
    con.executescript(DDL)
    have = set() if force else covered_days(con)
    days = []
    d = start
    while d <= end:
        if d.weekday() < 5 and d.isoformat() not in have:
            days.append(d)
        d += timedelta(days=1)
    C.log("決算サマリー %s..%s: 対象 %d 営業日（取得済み %d 日はスキップ）"
          % (start, end, len(days), len(have)))
    n_rows = n_days = 0
    now = C.utcnow()
    for i, d in enumerate(days, 1):
        try:
            raw = fetch_day(fetcher, d)
        except RuntimeError as e:
            C.log("  ! %s: %s" % (d, e))
            if "not available on your subscription" in str(e):
                C.log("  契約範囲外に到達したとみなして中断する")
                break
            continue
        rows = [x for x in (to_row(r, now) for r in raw) if x]
        con.executemany(
            "INSERT OR REPLACE INTO fins_summary (code, disc_date, doc_type, disc_time,"
            " per_type, per_start, per_end, fy_end, sales, op, odp, np, f_sales, f_op,"
            " fnc_sales, fnc_op, nxf_sales, nxf_op, fetched_at)"
            " VALUES (%s)" % ",".join("?" * 19), rows)
        n_rows += len(rows)
        n_days += 1
        if i % 50 == 0 or i == len(days):
            con.commit()
            C.log("  [%d/%d] %s  累計 %d 日 / %s 行" % (i, len(days), d, n_days, f"{n_rows:,}"))
    con.commit()
    return {"days": n_days, "rows": n_rows}


def report(con) -> None:
    con.executescript(DDL)
    r = con.execute("SELECT COUNT(*), COUNT(DISTINCT code), COUNT(DISTINCT disc_date),"
                    " MIN(disc_date), MAX(disc_date) FROM fins_summary").fetchone()
    C.log("fins_summary: %s 行 / %s 銘柄 / %s 営業日 / %s .. %s"
          % (f"{r[0]:,}", r[1], r[2], r[3], r[4]))
    for pt, n, with_op in con.execute(
            "SELECT per_type, COUNT(*), SUM(op IS NOT NULL) FROM fins_summary "
            "GROUP BY per_type ORDER BY 2 DESC"):
        C.log("  %-4s %6d 行（累計OPあり %d）" % (pt, n, with_op or 0))


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="J-Quants 決算サマリーの取り込み（進捗率の素材）")
    p.add_argument("--years", type=float, help="今日から何年ぶん遡るか（Light は5年）")
    p.add_argument("--from", dest="dfrom")
    p.add_argument("--to", dest="dto")
    p.add_argument("--force", action="store_true", help="取得済みの日も取り直す")
    p.add_argument("--rpm", type=int, default=60)
    p.add_argument("--report", action="store_true")
    a = p.parse_args(argv)

    con = C.init_db()
    if a.report:
        report(con)
        return 0
    end = C.parse_date_arg(a.dto) if a.dto else date.today()
    if a.years:
        start = end - timedelta(days=int(365.25 * a.years))
    elif a.dfrom:
        start = C.parse_date_arg(a.dfrom)
    else:
        p.error("--years か --from を指定する")
    fetcher = C.Fetcher(min_interval=60.0 / max(a.rpm, 1) * 1.3)
    jq_auth(fetcher)
    run_id = C.start_run(con, SOURCE, end.isoformat())
    try:
        st = backfill(con, fetcher, start, end, a.force)
        C.finish_run(con, run_id, "ok", n_saved=st["rows"],
                     note="fins summary %d days" % st["days"])
    except Exception as e:
        C.finish_run(con, run_id, "failed", error=str(e)[:400])
        raise
    report(con)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
