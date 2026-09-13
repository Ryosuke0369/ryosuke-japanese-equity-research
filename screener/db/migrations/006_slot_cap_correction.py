"""006_slot_cap_correction.py — 枠制約バグによる初週凍結の訂正。

2026-09-02、初回凍結で同時保有上限10に対し **13件** 建った。原因は
週次バッチが1回の実行で複数日ぶんを決めるのに、空き枠をDBだけで数えて
いたこと（書き込みはループ後なので日跨ぎの累積が見えない）。

訂正の方針（追記専用を壊さない）:
  - 既存の凍結13件は**削除しない**
  - `invalidated_by` / `invalidated_reason` を持つ訂正列を足し、
    無効化フラグを立てる
  - 再凍結は**元の判定日と同じ as_of** で行う。今日のデータで
    再計算すると判定日以降の情報が混入して PIT 違反になる

    python -m screener.db.migrations.006_slot_cap_correction
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))

from screener import common as C

REASON = "slot_cap_bug_20260902"
COLS = {
    "forecast_snapshots": (("invalidated", "INTEGER DEFAULT 0"),
                           ("invalidated_reason", "TEXT"),
                           ("invalidated_at", "TEXT")),
    "paper_trades": (("invalidated", "INTEGER DEFAULT 0"),
                     ("invalidated_reason", "TEXT"),
                     ("invalidated_at", "TEXT")),
    "shadow_snapshots": (("invalidated", "INTEGER DEFAULT 0"),
                         ("invalidated_reason", "TEXT"),
                         ("invalidated_at", "TEXT")),
    "shadow_trades": (("invalidated", "INTEGER DEFAULT 0"),
                      ("invalidated_reason", "TEXT"),
                      ("invalidated_at", "TEXT")),
}


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    con = C.init_db()
    for table, cols in COLS.items():
        have = {r[1] for r in con.execute("PRAGMA table_info(%s)" % table)}
        for name, typ in cols:
            if name not in have:
                con.execute("ALTER TABLE %s ADD COLUMN %s %s" % (table, name, typ))
                C.log("  %s.%s 追加" % (table, name))
    con.commit()

    now = C.utcnow()
    n = 0
    for table in ("forecast_snapshots", "paper_trades",
                  "shadow_snapshots", "shadow_trades"):
        cur = con.execute(
            "UPDATE %s SET invalidated=1, invalidated_reason=?, invalidated_at=? "
            "WHERE COALESCE(invalidated,0)=0" % table, (REASON, now))
        C.log("  %s: %d 行を無効化" % (table, cur.rowcount))
        n += cur.rowcount
    con.commit()
    C.log("合計 %d 行に訂正フラグ（理由: %s）" % (n, REASON))
    C.log("※ 行は削除していない。再凍結は同一 as_of で行うこと。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
