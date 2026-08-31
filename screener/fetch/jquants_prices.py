"""screener/fetch/jquants_prices.py — 日次株価と TOPIX のバックフィル (J-Quants V2)。

バックテストは (1) 銘柄の日次終値 (2) 市場指数 の2つを必要とする。仕様書 §5 の
prices テーブルと、超過リターン評価用の market_index テーブルを埋める。

## プランと遡及範囲

無料プランは2年ローリング(直近12週を除く)で、EDINET四半期報告書から作れる
FY2023〜FY2024 の窓と1日も重ならなかった。Light 加入により 2021-08-31 まで
遡れるようになった(2026-08-31 実測。API が message で範囲を明示する)。

範囲外の日付は HTTP 400 を返し、message に契約範囲が入る。**推測せず、
API が言う範囲を信じて止まる** —— 400 を握りつぶして空で埋めると、
「データが無い日」と「契約範囲外の日」が区別できなくなる。

## レート制限

分あたり(Light 60 req/min)。jquants_universe と同じく 60/rpm × 1.3 の
マージンを取り、429 を見たら間隔を1.3倍する自動減速をかける。

Usage
    python -m screener.fetch.jquants_prices --prices --from 2021-08-31 --rpm 60
    python -m screener.fetch.jquants_prices --topix --from 2021-08-31 --rpm 60
    python -m screener.fetch.jquants_prices --report
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

from screener.fetch.jquants_universe import JQ, jq_auth, jq_get, _jq_diagnose

EP_BARS = "/equities/bars/daily"
EP_TOPIX = "/indices/bars/daily/topix"
SOURCE = "jquants_prices"


def _days(start: date, end: date):
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


def backfill_prices(con, fetcher, start: date, end: date, codes: set | None) -> dict:
    """日次四本値を1日1リクエストで埋める。既に入っている日は飛ばす。"""
    # 既に取得済みでも、調整後株価などの新しい列が NULL の日は取り直す。
    # 「行がある」と「必要な列が埋まっている」は違う。
    have = {r["date"] for r in con.execute(
        "SELECT date FROM prices GROUP BY date "
        "HAVING SUM(CASE WHEN adj_close IS NULL THEN 1 ELSE 0 END) = 0")}
    todo = [d for d in _days(start, end) if d.isoformat() not in have]
    C.log(f"株価バックフィル {start}..{end}: 対象 {len(todo)} 営業日 "
          f"(取得済み {len(have)} 日はスキップ)")
    n_rows = n_days = 0
    for i, d in enumerate(todo, 1):
        try:
            rows = jq_get(fetcher, EP_BARS, date=d.strftime("%Y%m%d"))
        except RuntimeError as e:
            if "HTTP 400" in str(e):
                C.log(f"  {d}: {e}")
                C.log("  契約範囲外に到達したとみなして中断する（推測で埋めない）")
                break
            raise
        if not rows:
            continue
        buf = []
        for r in rows:
            code = C.normalise_code(r.get("Code"))
            if codes is not None and code not in codes:
                continue
            c = r.get("C")
            if c is None:
                continue
            # API が返すものは全部取る。5年ローリングで窓から落ちた日付は
            # 二度と取得できないので、「今は使わない」列も落とさない。
            buf.append((code, d.isoformat(), c, r.get("Vo"), r.get("Va"),
                        r.get("O"), r.get("H"), r.get("L"),
                        r.get("AdjFactor"), r.get("AdjC"), r.get("AdjVo"),
                        r.get("MktCap")))
        con.executemany(
            "INSERT OR REPLACE INTO prices (code, date, close, volume, "
            " turnover_value, open, high, low, adj_factor, adj_close, "
            " adj_volume, mktcap) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", buf)
        n_rows += len(buf)
        n_days += 1
        if i % 25 == 0 or i == len(todo):
            con.commit()
            C.log(f"  [{i}/{len(todo)}] {d}  累計 {n_days} 日 / {n_rows:,} 行 "
                  f"(間隔 {fetcher.throttle.min_interval:.1f}s, {fetcher.n_requests} req)")
    con.commit()
    return {"days": n_days, "rows": n_rows}


def backfill_topix(con, fetcher, start: date, end: date) -> dict:
    """TOPIX は from/to の範囲指定が効くので、月単位でまとめて取る。"""
    C.log(f"TOPIX バックフィル {start}..{end}")
    n = 0
    cur = start
    while cur <= end:
        nxt = min(date(cur.year + (cur.month // 12), cur.month % 12 + 1, 1)
                  - timedelta(days=1), end)
        try:
            rows = jq_get(fetcher, EP_TOPIX,
                          **{"from": cur.strftime("%Y%m%d"),
                             "to": nxt.strftime("%Y%m%d")})
        except RuntimeError as e:
            if "HTTP 400" in str(e):
                C.log(f"  {cur}..{nxt}: {e}")
                C.log("  契約範囲外に到達したとみなして中断する")
                break
            raise
        con.executemany(
            "INSERT OR REPLACE INTO market_index (date, close) VALUES (?,?)",
            [(r["Date"], r["C"]) for r in rows if r.get("C") is not None])
        n += len(rows)
        con.commit()
        C.log(f"  {cur:%Y-%m}  {len(rows):>3} 日  累計 {n:,}")
        cur = nxt + timedelta(days=1)
    return {"rows": n}


def report(con) -> None:
    p = con.execute("SELECT COUNT(*) c, COUNT(DISTINCT code) n, COUNT(DISTINCT date) d, "
                    "MIN(date) a, MAX(date) b FROM prices").fetchone()
    C.log(f"prices: {p['c']:,} 行 / {p['n']} 銘柄 / {p['d']} 営業日 / "
          f"{p['a']} .. {p['b']}")
    m = con.execute("SELECT COUNT(*) c, MIN(date) a, MAX(date) b "
                    "FROM market_index").fetchone()
    C.log(f"market_index: {m['c']:,} 行 / {m['a']} .. {m['b']}")


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="日次株価/TOPIX バックフィル")
    p.add_argument("--prices", action="store_true")
    p.add_argument("--topix", action="store_true")
    p.add_argument("--report", action="store_true")
    p.add_argument("--from", dest="dfrom", default="2021-08-31")
    p.add_argument("--to", dest="dto")
    p.add_argument("--rpm", type=int, default=60,
                   help="契約プランのレート上限。Free 5 / Light 60 / Standard 120")
    p.add_argument("--all-codes", action="store_true",
                   help="ユニバース外も保存する（既定はユニバース候補∪検証8銘柄）")
    a = p.parse_args(argv)

    con = C.init_db()
    if a.report:
        report(con)
        return 0
    start = C.parse_date_arg(a.dfrom)
    end = C.parse_date_arg(a.dto) if a.dto else date.today()
    # 公称値ぎりぎりは1分窓の境界で429になる。+30%のマージンを取る。
    fetcher = C.Fetcher(min_interval=60.0 / max(a.rpm, 1) * 1.3)
    jq_auth(fetcher)

    codes = None
    if not a.all_codes:
        from screener.fetch.edinet_bulk import universe_codes
        codes = universe_codes(con)
        C.log(f"保存対象: {len(codes)} 銘柄（ユニバース候補 ∪ 検証8銘柄）")

    run_id = C.start_run(con, SOURCE, end.isoformat())
    try:
        if a.prices:
            r = backfill_prices(con, fetcher, start, end, codes)
            C.finish_run(con, run_id, "ok", n_saved=r["rows"],
                         note=f"prices {r['days']} days")
        if a.topix:
            r = backfill_topix(con, fetcher, start, end)
            C.finish_run(con, run_id, "ok", n_saved=r["rows"], note="topix")
    except Exception as e:
        C.finish_run(con, run_id, "failed", error=str(e)[:400])
        raise
    report(con)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
