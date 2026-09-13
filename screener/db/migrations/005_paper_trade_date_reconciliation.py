"""005_paper_trade_date_reconciliation.py — paper_trades に発表日の推定/実績列を足す。

schema.sql は CREATE TABLE IF NOT EXISTS なので、**既に作られたテーブルには
新しい列が入らない**。2026-09-02、ペーパートレードの初回凍結で
`no column named event_date_estimated` を踏んだ。

足す列:
  event_date_estimated  エントリー起点にした推定発表日
  event_date_actual     EDINET/TDnet で捕捉した実際の発表日
  date_error_bdays      実績 - 推定（営業日）

    python -m screener.db.migrations.005_paper_trade_date_reconciliation
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))

from screener import common as C

COLS = (("event_date_estimated", "TEXT"),
        ("event_date_actual", "TEXT"),
        ("date_error_bdays", "INTEGER"))


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    con = C.init_db()
    have = {r[1] for r in con.execute("PRAGMA table_info(paper_trades)")}
    added = 0
    for name, typ in COLS:
        if name in have:
            C.log(f"  {name}: 既にある")
            continue
        con.execute(f"ALTER TABLE paper_trades ADD COLUMN {name} {typ}")
        C.log(f"  {name}: 追加")
        added += 1
    con.commit()
    after = [r[1] for r in con.execute("PRAGMA table_info(paper_trades)")]
    C.log(f"paper_trades の列: {after}")
    C.log(f"追加 {added} 列")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
