"""screener/report/tdnet_completeness.py — TDnet 収集の件数突合（全営業日）と修復。

なぜ要るか
----------
fetch_runs の 'ok' は「その回に一覧が読めた」しか意味していなかった。朝に
走った回は当日分の一部しか無くても 'ok' になり、backfill は欠損を見ない
（2026-09-13 に 9/02,03,04,10,11 の5日で発見）。**記録ではなく一次ソースの
件数と突き合わせる**のがこのスクリプト。

1営業日ごとに
  actual_total   TDnet 一覧の「全N件」（今の時点で読み直す）
  actual_scope   そのうち保存対象（短信・予想修正・配当修正・説明資料）
  db_present     対象書類のうち filings に行があり、ファイルが実在する件数
  recorded_*     fetch_runs の最新行（status / n_listed / n_target / n_saved）
を並べ、actual_scope != db_present または status が ok/empty でない日を
「不一致」とする。一覧が 404 の日は保持期間外で照合不能（unreachable）。

    python -m screener.report.tdnet_completeness --from 2026-07-23 --to 2026-09-11
    python -m screener.report.tdnet_completeness --from 2026-07-23 --to 2026-09-11 --repair \
        --csv C:/screener_data/tdnet_completeness_20260913.csv
"""
from __future__ import annotations

import argparse
import csv
import os
import sys
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.fetch import tdnet_archiver as T


def reachable(fetcher, d: date) -> bool:
    r = fetcher.get(T.LIST_URL.format(page=1, ymd=d.strftime("%Y%m%d")))
    return r.status_code == 200


def check_day(con, fetcher, d: date) -> dict:
    iso = d.isoformat()
    rec = con.execute(
        "SELECT status, n_listed, n_target, n_saved, attempt, started_at FROM fetch_runs "
        "WHERE source='tdnet' AND target_date=? ORDER BY id DESC LIMIT 1", (iso,)).fetchone()
    out = {"date": iso,
           "recorded_status": rec["status"] if rec else "",
           "recorded_listed": rec["n_listed"] if rec else "",
           "recorded_in_scope": rec["n_target"] if rec else "",
           "recorded_saved": rec["n_saved"] if rec else "",
           "recorded_attempt": rec["attempt"] if rec else "",
           "db_filings_on_date": con.execute(
               "SELECT COUNT(*) FROM filings WHERE source='tdnet' AND date=?", (iso,)).fetchone()[0]}
    if not reachable(fetcher, d):
        out.update({"reach": "unreachable", "actual_total": "", "actual_rows": "",
                    "actual_in_scope": "", "db_present": "", "verdict": "照合不能(保持期間外)"})
        return out
    rows, total = T.fetch_day_index(fetcher, d)
    targets = [r for r in rows if T.classify(r["title"])[0]]
    present = 0
    for t in targets:
        f = con.execute("SELECT path FROM filings WHERE source='tdnet' AND doc_id=?",
                        (T._doc_id_of(t),)).fetchone()
        if f and f["path"] and os.path.exists(C.full_path(f["path"])):
            present += 1
    ok_status = out["recorded_status"] in ("ok", "empty")
    complete = (total == len(rows)) and present == len(targets)
    out.update({"reach": "ok", "actual_total": total, "actual_rows": len(rows),
                "actual_in_scope": len(targets), "db_present": present,
                "verdict": "一致" if (complete and ok_status and
                                     str(out["recorded_listed"]) == str(total)) else "不一致"})
    return out


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--from", dest="dfrom", required=True)
    p.add_argument("--to", dest="dto", required=True)
    p.add_argument("--repair", action="store_true", help="不一致の日を archive_day で取り直す")
    p.add_argument("--csv")
    a = p.parse_args(argv)
    start, end = date.fromisoformat(a.dfrom), date.fromisoformat(a.dto)
    con = C.init_db()
    fetcher = C.Fetcher(min_interval=0.7)

    results = []
    d = start
    while d <= end:
        if d.weekday() < 5:
            before = check_day(con, fetcher, d)
            before["added"] = ""
            before["after_status"] = ""
            before["after_present"] = ""
            if a.repair and before["verdict"] == "不一致":
                res = T.archive_day(con, fetcher, d)
                after = check_day(con, fetcher, d)
                before["added"] = res["saved"]
                before["after_status"] = res["status"]
                before["after_present"] = after["db_present"]
                before["after_verdict"] = after["verdict"]
            results.append(before)
            C.log("  %s %-18s rec=%s/%s/%s act=%s/%s db=%s %s%s" % (
                before["date"], before["verdict"], before["recorded_status"],
                before["recorded_listed"], before["recorded_in_scope"],
                before["actual_total"], before["actual_in_scope"], before["db_present"],
                ("added=%s after=%s/%s" % (before["added"], before["after_status"],
                                           before["after_present"])) if before["added"] != "" else "",
                ""))
        d += timedelta(days=1)

    from collections import Counter
    C.log("verdict: %s" % dict(Counter(r["verdict"] for r in results)))
    if a.repair:
        C.log("repaired days: %d / added file sets: %d / still mismatched: %d" % (
            sum(1 for r in results if r["added"] != ""),
            sum(int(r["added"] or 0) for r in results),
            sum(1 for r in results if r.get("after_verdict") == "不一致")))
    if a.csv:
        keys = []
        for r in results:
            keys += [k for k in r if k not in keys]
        with open(a.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=keys)
            w.writeheader()
            w.writerows(results)
        C.log("CSV: %s" % a.csv)
    bad = [r for r in results if (r.get("after_verdict") or r["verdict"]) == "不一致"]
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
