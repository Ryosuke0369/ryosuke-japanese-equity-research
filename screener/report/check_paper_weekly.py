"""screener/report/check_paper_weekly.py — 週次バッチが走ったかの健全性チェック。

火曜 06:00 に起動され、「今週の月曜の週次バッチが成功しているか」を
fetch_runs(source='paper_weekly') で確認する。無ければ ERROR を出して
終了コード2で落ちる。次にセッションを開いたとき、ログを見れば気づける。

**このスクリプトは何も直さない。** 直すのは人間の仕事で、ここの仕事は
「黙って壊れている状態」を作らないこと。バッチが起動すらしなかった場合、
バッチ自身は何も書けない。**起動しなかったことを検知できるのは外側だけ**
なので、別プロセス・別トリガーで見る。

    python -m screener.report.check_paper_weekly
    python -m screener.report.check_paper_weekly --week 2026-09-07
"""
from __future__ import annotations

import argparse
import datetime
import os
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C


def check(con, monday):
    row = con.execute(
        "SELECT status, finished_at, note FROM fetch_runs "
        "WHERE source='paper_weekly' AND target_date=? "
        "ORDER BY id DESC LIMIT 1", (monday.isoformat(),)).fetchone()
    if row is None:
        return 2, ("ERROR: %s 週の paper_weekly の実行記録が無い"
                   "（バッチが起動していない）" % monday)
    if row["status"] != "ok":
        return 2, ("ERROR: %s 週の paper_weekly が status=%s で終わっている: %s"
                   % (monday, row["status"], row["note"] or ""))
    return 0, ("OK: %s 週の paper_weekly は成功 (%s) %s"
               % (monday, row["finished_at"], row["note"] or ""))


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--week", help="確認する週の月曜 YYYY-MM-DD（既定は今週）")
    a = p.parse_args(argv)

    if a.week:
        monday = datetime.date.fromisoformat(a.week)
    else:
        today = datetime.date.today()
        monday = today - datetime.timedelta(days=today.weekday())

    con = C.init_db()
    rc, msg = check(con, monday)
    C.log(msg)
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
