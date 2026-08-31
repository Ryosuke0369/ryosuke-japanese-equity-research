r"""003 - prices に始値・調整後株価・時価総額を足す。

`CREATE TABLE IF NOT EXISTS` は既存テーブルに列を足さないので ALTER で埋める。

なぜ急ぐか: J-Quants のカバレッジは5年ローリングで、**窓から落ちた日付は
二度と取得できない**。取れるうちに API が返す全項目を保存する。

  open       出口ルール「翌日寄り指値」の再現に必須。終値だけでは執行を模擬できない
  adj_close  株式分割の調整。未調整だと分割が偽の -50% リターンとしてバックテストに
             入る(86970 の 2022-01-04 は C=2518.5 に対し AdjC=1259.3)
  mktcap     日次時価総額。ユニバース条件を過去時点で再現するのに使う

    python -m screener.db.migrations.003_prices_ohlc_adjusted [--apply]
"""
from __future__ import annotations

import argparse
import os
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))))))
    from screener import common as C

COLS = ("open", "high", "low", "adj_factor", "adj_close", "adj_volume", "mktcap")


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="prices に OHLC/調整後株価を追加")
    p.add_argument("--apply", action="store_true")
    a = p.parse_args(argv)

    con = C.connect()
    have = {r[1] for r in con.execute("PRAGMA table_info(prices)")}
    todo = [c for c in COLS if c not in have]
    if not todo:
        C.log("追加すべき列は無い")
        return 0
    n = con.execute("SELECT COUNT(*) c FROM prices").fetchone()["c"]
    C.log(f"prices {n} 行に列を追加: {', '.join(todo)}")
    if not a.apply:
        C.log("ドライラン。実際に変更するには --apply")
        return 0
    for c in todo:
        con.execute(f"ALTER TABLE prices ADD COLUMN {c} REAL")
    con.commit()
    C.log("追加した。既存行の新列は NULL のままなので、取り直しが要る")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
