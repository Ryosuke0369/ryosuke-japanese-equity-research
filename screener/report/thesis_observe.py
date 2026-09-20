"""screener/report/thesis_observe.py — テーゼ破綻の観測ログ（前向き・記録専用）。

正本は docs/backtest_acceptance_criteria.md「シャドウE 前向き記録の範囲」（2026-09-20）。
**売買は一切変えない。** v2 / シャドウA〜D の採否・出口・記録には触れない。
variant も増やさない。ここが書くのは観測テーブル `thesis_observations` だけ。

なぜ記録するのか
----------------
インサンプルの測定（calibration_backlog §39）で、テーゼ破綻（分岐2）は
**出口としては損益をほとんど動かさない**が、**建玉の良し悪しの識別力はある**と出た
（健全 +0.031〜+0.042 / 破綻 −0.016〜+0.002・3本とも同符号）。
E4 の形で前向きに回す意味は無いが、識別力だけは前向きに貯める価値がある。

判定は出口の分岐2 と同じ規則（`exit_sim.thesis_break`）。
**OR（売上 or 営業利益が前年同期比 ≤ 0）を本線、AND を併記。**

    python -m screener.report.thesis_observe            # 決算が出た建玉を記録
    python -m screener.report.thesis_observe --report   # 貯まった符号を見る
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import exit_sim as X

DDL = """
CREATE TABLE IF NOT EXISTS thesis_observations (
    trade_id        INTEGER PRIMARY KEY,   -- paper_trades.trade_id（v2 の建玉）
    code            TEXT NOT NULL,
    event_date      TEXT NOT NULL,         -- 判定に使った開示日
    observed_at     TEXT NOT NULL,
    evaluable       INTEGER NOT NULL,      -- 前年同期が取れたか
    break_or        INTEGER,               -- 本線: 売上 or 営業利益が前年割れ
    break_and       INTEGER,               -- 併記: 両方とも前年割れ
    sales_yoy_diff  REAL,                  -- run-rate の差（前年が負でも壊れない）
    op_yoy_diff     REAL,
    ret_net         REAL                   -- その建玉の実現リターン（出たあとに埋まる）
);
"""


def observe(mcon, pcon, as_of=None) -> dict:
    """決算を通過した v2 の建玉について、テーゼ破綻を記録する（冪等）。"""
    mcon.executescript(DDL)
    rows = mcon.execute(
        "SELECT trade_id, code, event_date, ret_net FROM paper_trades "
        "WHERE event_date IS NOT NULL").fetchall()
    n_new = n_upd = 0
    for trade_id, code, event_date, ret_net in rows:
        if as_of and event_date > as_of:
            continue                      # まだ決算が来ていない建玉は書かない
        th = X.thesis_break(pcon, code, event_date)
        cur = mcon.execute("SELECT 1 FROM thesis_observations WHERE trade_id=?",
                           (trade_id,)).fetchone()
        mcon.execute(
            "INSERT INTO thesis_observations (trade_id, code, event_date, observed_at,"
            " evaluable, break_or, break_and, sales_yoy_diff, op_yoy_diff, ret_net)"
            " VALUES (?,?,?,?,?,?,?,?,?,?)"
            " ON CONFLICT(trade_id) DO UPDATE SET ret_net=excluded.ret_net,"
            " observed_at=excluded.observed_at",
            (trade_id, code, event_date, C.utcnow(), int(th["evaluable"]),
             int(th["break_or"]), int(th["break_and"]),
             th["sales_yoy"], th["op_yoy"], ret_net))
        n_upd += 1 if cur else 0
        n_new += 0 if cur else 1
    mcon.commit()
    return {"new": n_new, "updated": n_upd, "trades": len(rows)}


def report(mcon) -> str:
    mcon.executescript(DDL)
    L = ["# テーゼ破綻の観測ログ（記録専用・売買は変えていない）", ""]
    tot = mcon.execute("SELECT COUNT(*) FROM thesis_observations").fetchone()[0]
    L.append("観測 %d 件" % tot)
    if not tot:
        L.append("（まだ決算を通過した建玉が無い）")
        return "\n".join(L)
    L.append("")
    L.append("| 群 | 件数 | 実現リターン(純)の平均 | 勝率 |")
    L.append("|---|---|---|---|")
    for label, where in (("テーゼ健全", "evaluable=1 AND break_or=0"),
                         ("テーゼ破綻(OR)", "evaluable=1 AND break_or=1"),
                         ("（参考）AND該当", "evaluable=1 AND break_and=1"),
                         ("判定不能", "evaluable=0")):
        r = mcon.execute(
            "SELECT COUNT(*), AVG(ret_net), AVG(CASE WHEN ret_net>0 THEN 1.0 ELSE 0.0 END) "
            "FROM thesis_observations WHERE %s AND ret_net IS NOT NULL" % where).fetchone()
        L.append("| %s | %d | %s | %s |" % (
            label, r[0], "–" if r[1] is None else "%+.4f" % r[1],
            "–" if r[2] is None else "%.0f%%" % (100 * r[2])))
    L.append("")
    L.append("**評価は 4〜6四半期後（v3-4 と同じ）。それまでは件数と符号の記録のみ。**")
    return "\n".join(L)


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--report", action="store_true", help="記録を読むだけ")
    p.add_argument("--as-of", help="この日までに開示された分だけ記録する")
    p.add_argument("--projection", default=os.path.join(C.DATA_DIR, "projection.db"))
    a = p.parse_args(argv)
    mcon = C.connect()
    if a.report:
        print(report(mcon))
        return 0
    pcon = sqlite3.connect("file:%s?mode=ro" % a.projection.replace("\\", "/"), uri=True)
    st = observe(mcon, pcon, a.as_of)
    C.log("観測ログ: 新規 %d / 更新 %d（建玉 %d）" % (st["new"], st["updated"], st["trades"]))
    print(report(mcon))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
