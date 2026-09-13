"""screener/report/exit_rule_check.py — B3: 4分岐エグジットの設計検証（1回のみ）。

4分岐ルールはモック設計段階（データを見る前）に定めた既存設計なので、
**インサンプルでの「設計の確認」として1回だけ**評価する。

  これはチューニングではない。結果を見て 4分岐のパラメータ
  （MISS_LINE 95% / INLINE_BAND 105% 等）をいじることは禁止。
  変更したくなったら新規に事前登録が要る。

比較するもの
------------
v2 と同じ建玉セット（T-15 エントリー・10枠制約・市場フィルター）に対し、

  固定    : T+2 の終値で全売却（v2 の実装そのもの）
  4分岐   : 発表を見てから decide_exit() の判断で売却

決済タイミングの扱い（規則が定めていない部分は明示する）:
  STOP_LOSS / TP_ALL / EXIT は全売却 → 発表日の翌営業日終値
  TP50_TRAIL は「50%利確 + トレール」だが、トレール幅を規則が定めていない。
  **ここで勝手に決めるとチューニングになる**ので、50% を翌営業日、
  残り 50% を T+5 とする代理実装で計算し、その旨を明示して報告する。

    python -m screener.report.exit_rule_check
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys
from collections import Counter
from datetime import date

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1
from screener.report import backtest_v2 as V2

COST = V1.COST_ROUND_TRIP


def _px_after(con, ticker, d, n):
    """d より後の n 営業日目の終値。無ければ最後に取れた日の終値。"""
    rows = con.execute(
        "SELECT date, close FROM daily_prices WHERE ticker=? AND date>? "
        "ORDER BY date LIMIT ?", (ticker, d, n)).fetchall()
    return (rows[-1][1], rows[-1][0]) if rows else (None, None)


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    argparse.ArgumentParser(description=__doc__).parse_args(argv)

    root = V1._external_root()
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from module_c.exit_engine import decide_exit
    from module_b.forecaster import forecast

    db = V1._db_path(0)
    con = sqlite3.connect("file:%s?mode=ro" % db.replace("\\", "/"), uri=True)
    con.row_factory = sqlite3.Row

    C.log("=== B3 4分岐エグジットの設計検証（1回のみ）===")
    _res, trades = V1.run_one(0)
    sim = V2.simulate(trades, con, V1.MAIN_ENTRY, V1.MAIN_EXIT)
    taken = sim["trades"]
    C.log("  v2 と同じ建玉セット: %d 件" % len(taken))

    branch = Counter()
    fixed_rets, rule_rets = [], []
    n_eval = n_nopred = n_noactual = n_error = 0
    err_sample = []
    reasons = Counter()

    for t in taken:
        code, ev = t["ticker"], t["event_date"]
        try:
            f = forecast(con, code, date.fromisoformat(t["entry_date"]))
        except Exception as e:
            # 例外を available=False に混ぜない。配線ミスとデータ不足は別物で、
            # 混ぜると「データが無い」ように見えて原因にたどり着けなくなる。
            n_error += 1
            if not err_sample:
                err_sample.append("%s: %s" % (type(e).__name__, e))
            continue
        if not f.get("available"):
            n_nopred += 1
            reasons[f.get("reason", "?")] += 1
            continue
        row = con.execute(
            "SELECT sales, operating_profit FROM quarterly_standalone "
            "WHERE ticker=? AND is_valid=1 AND period_end=("
            "  SELECT period_end FROM filings WHERE ticker=? AND filing_date=? LIMIT 1)",
            (code, code, ev)).fetchone()
        if not row or row["sales"] is None or row["operating_profit"] is None:
            n_noactual += 1
            continue
        g = con.execute(
            "SELECT forecast_op FROM company_forecasts WHERE ticker=? AND source_date<=? "
            "ORDER BY source_date DESC LIMIT 2", (code, ev)).fetchall()
        new_g = g[0]["forecast_op"] if g else None
        prev_g = g[1]["forecast_op"] if len(g) > 1 else None

        d = decide_exit(f.get("pred_sales") or 0, f.get("pred_op") or 0,
                        row["sales"], row["operating_profit"],
                        prev_g, new_g, None)
        branch[d["action"]] += 1
        n_eval += 1

        entry_px = con.execute(
            "SELECT close FROM daily_prices WHERE ticker=? AND date=?",
            (code, t["entry_date"])).fetchone()
        if not entry_px:
            continue
        ep = entry_px[0]
        p1, _ = _px_after(con, code, ev, 1)
        p5, _ = _px_after(con, code, ev, 5)
        if p1 is None:
            continue
        if d["action"] == "TP50_TRAIL" and p5 is not None:
            rule = 0.5 * (p1 / ep - 1) + 0.5 * (p5 / ep - 1) - COST
        else:
            rule = p1 / ep - 1 - COST
        fixed_rets.append(t["net"])
        rule_rets.append(rule)

    C.log("  予測できず(データ不足) %d / 実績が取れず %d / 例外 %d / 評価できた %d"
          % (n_nopred, n_noactual, n_error, n_eval))
    if err_sample:
        C.log("  例外の例: %s" % err_sample[0])
    for r, n in reasons.most_common(4):
        C.log("    データ不足の理由: %s = %d 件" % (r, n))
    if not rule_rets:
        C.log("  → 比較できるトレードが無い。設計検証は成立しない。")
        return 0

    def stat(rs):
        w = [r for r in rs if r > 0]
        return (len(rs), sum(rs) / len(rs), len(w) / len(rs))

    nf, ef, wf = stat(fixed_rets)
    nr, er, wr = stat(rule_rets)
    C.log("")
    C.log("=== 損益比較（コスト往復0.4%控除後・同一건玉）===".replace("건", "建"))
    C.log("  %-12s %6s %12s %8s" % ("方式", "件数", "期待値", "勝率"))
    C.log("  %-12s %6d %+11.4f %7.1f%%" % ("固定 T+2", nf, ef, wf * 100))
    C.log("  %-12s %6d %+11.4f %7.1f%%" % ("4分岐", nr, er, wr * 100))
    C.log("  差分（4分岐 - 固定）: 期待値 %+.4f / 勝率 %+.1f pt"
          % (er - ef, (wr - wf) * 100))
    C.log("")
    C.log("=== 分岐の内訳 ===")
    for k, v in branch.most_common():
        C.log("  %-12s %4d 件 (%.0f%%)" % (k, v, v / max(n_eval, 1) * 100))
    C.log("")
    C.log("  ※ TP50_TRAIL のトレール幅は規則が定めていない。ここでは")
    C.log("    50%%を翌営業日・残り50%%を T+5 とする代理実装で計算した。")
    C.log("  ※ これは設計の確認であり、結果を見てパラメータを変えない。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
