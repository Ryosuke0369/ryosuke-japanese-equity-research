"""screener/report/s5c_check.py — シャドウI: S5c の実現可能性と、上方修正の先行力。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウI」（2026-09-21 登録）。
**このスクリプトは定義を実装するだけで、閾値や窓を決めない。**

I-4 の順に出す:
  1. 実現可能性 —— S5c が計算できる銘柄数・判定日数と、計算できない理由の内訳
  2. 符号 —— シャドウG と同じ枠組み（90日窓・同じベースレート・同じ群分け）で
     S5c 発火の up率を比べる。**S5 と S5c を並べる**
  3. 母集団を「通期予想がある観測」に揃えた比較も同時に出す（§46 の診断を受けて登録済み）

**S5c は 0点・表示のみ。** スコア・採否・出口には接続しない。

    python -m screener.report.s5c_check
"""
from __future__ import annotations

import argparse
import csv
import os
import sqlite3
import sys
from collections import Counter, defaultdict
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import s5_revision as G
from screener.signals import s5c_landing as S5C
from screener.signals.span_runner import score_ticker

GROUPS = ("S5c発火(乖離30%以上)", "S5c非発火", "S5c計算不能")


def collect(mcon, pcon, dates, codes, revs):
    """(判定日 × 銘柄) に S5 と S5c の両方を付ける。"""
    scorers = G.scorers_all()
    last_day = G.LAST_DATE.isoformat()
    rows, reasons = [], Counter()
    for as_of in dates:
        horizon = (date.fromisoformat(as_of) + timedelta(days=G.WINDOW_DAYS)).isoformat()
        for code in codes:
            try:
                res = score_ticker(pcon, code, scorers, as_of=as_of, policy=G.POLICY)
            except Exception:
                res = {}
            s5_group, s5_score, dead = G.s5_state(res)
            c = S5C.evaluate(pcon, code, as_of=as_of)
            if not c["available"]:
                reasons[c["evidence"].split("（")[0][:40]] += 1
                group = "S5c計算不能"
            else:
                group = GROUPS[0] if c["fired"] else GROUPS[1]
            rows.append({
                "as_of": as_of, "code": code,
                "s5_group": s5_group, "s5_score": s5_score, "guidance_dead": int(dead),
                "s5c_group": group, "s5c_gap_op": c["gap_op"],
                "s5c_gap_sales": c["gap_sales"], "s5c_sales_fired": int(c["sales_fired"]),
                "s5c_elapsed_q": c["elapsed_q"], "s5c_progress_prior": c["progress_prior"],
                "has_forecast": int(c["forecast_op"] is not None),
                "outcome": G.outcome_for(revs, code, as_of, min(horizon, last_day)),
                "window_days": (date.fromisoformat(last_day)
                                - date.fromisoformat(as_of)).days,
            })
    return rows, reasons


def _rate(sub):
    return (sum(r["outcome"] == "up" for r in sub) / len(sub)) if sub else None


def _line(label, sub, base):
    if not sub:
        return "| %s | 0 | – | – | – |" % label
    p = _rate(sub)
    d, lo, hi = G.diff_ci(p, len(sub), _rate(base), len(base))
    codes = {r["code"] for r in sub}
    return "| %s | %d | %d | %.2f%% | %s |" % (
        label, len(sub), len({r["code"] for r in sub if r["outcome"] == "up"}),
        100 * p, "–" if d is None else "%+.2fpt [%+.2f, %+.2f]" % (100 * d, 100 * lo, 100 * hi))


def render(rows, reasons, dates, codes):
    L = ["# シャドウI: S5c（着地推定 × 開示義務ライン）の実現可能性と先行力", ""]
    L.append("- 判定日 %d 日（%s 〜 %s）/ ユニバース %d 銘柄 / のべ %d 観測"
             % (len(dates), dates[0], dates[-1], len(codes), len(rows)))
    L.append("- 追跡窓 %d 暦日（**満了していない。実窓 %d〜%d 日**）"
             % (G.WINDOW_DAYS, min(r["window_days"] for r in rows),
                max(r["window_days"] for r in rows)))
    L.append("- S5c は **0点・表示のみ**。スコアには接続していない")
    L.append("")
    L.append("## 1. 実現可能性（I-4-1）")
    L.append("")
    ok = [r for r in rows if r["s5c_group"] != "S5c計算不能"]
    L.append("計算できた観測 **%d / %d（%.0f%%）** / 銘柄ユニーク **%d / %d**"
             % (len(ok), len(rows), 100 * len(ok) / len(rows),
                len({r["code"] for r in ok}), len(codes)))
    L.append("")
    L.append("| 計算できない理由 | 件数 |")
    L.append("|---|---|")
    for why, n in reasons.most_common():
        L.append("| %s | %d |" % (why, n))
    L.append("")
    eq = Counter(r["s5c_elapsed_q"] for r in ok)
    L.append("経過四半期の内訳: " + " / ".join("Q%s %d" % (k, v) for k, v in sorted(eq.items())))
    L.append("")
    L.append("## 2. 上方修正の先行力（I-4-2・G と同じ枠組み）")
    L.append("")
    L.append("| 群 | のべn | up銘柄 | up率 | ベースレート差 [95%CI] |")
    L.append("|---|---|---|---|---|")
    for g in GROUPS:
        L.append(_line(g, [r for r in rows if r["s5c_group"] == g], rows))
    L.append(_line("（参考）S5発火", [r for r in rows if r["s5_group"] == "S5発火(進捗超過)"], rows))
    L.append(_line("ユニバース全体", rows, rows))
    L.append("")
    L.append("## 3. 母集団を「通期予想がある観測」に揃えた比較（I-4-3）")
    L.append("")
    base = [r for r in rows if r["has_forecast"]]
    L.append("揃えた母集団 %d 観測（銘柄 %d）" % (len(base), len({r["code"] for r in base})))
    L.append("")
    L.append("| 群 | のべn | up銘柄 | up率 | 揃えた全体との差 [95%CI] |")
    L.append("|---|---|---|---|---|")
    for g in GROUPS[:2]:
        L.append(_line(g, [r for r in base if r["s5c_group"] == g], base))
    L.append(_line("（参考）S5発火", [r for r in base if r["s5_group"] == "S5発火(進捗超過)"], base))
    L.append(_line("揃えた全体", base, base))
    L.append("")
    L.append("## 4. 乖離の分布（発火ラインの位置を見る）")
    L.append("")
    gaps = sorted(r["s5c_gap_op"] for r in ok if r["s5c_gap_op"] is not None)
    if gaps:
        from screener.technical.event_response import percentile as pct
        L.append("| 分位 | 乖離 |")
        L.append("|---|---|")
        for q in (0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99):
            L.append("| %.0f%% | %+.1f%% |" % (100 * q, 100 * pct(gaps, q)))
        L.append("")
        L.append("+30% 以上 %d 件（%.1f%%）/ 売上10%%ライン発火 %d 件"
                 % (sum(1 for g in gaps if g >= S5C.FIRE_OP_PCT),
                    100 * sum(1 for g in gaps if g >= S5C.FIRE_OP_PCT) / len(gaps),
                    sum(r["s5c_sales_fired"] for r in ok)))
    L.append("")
    L.append("## 5. S5 と S5c の重なり")
    L.append("")
    L.append("| | S5c発火 | S5c非発火 | S5c計算不能 |")
    L.append("|---|---|---|---|")
    for s5g in ("S5発火(進捗超過)", "S5非発火", "S5評価不能"):
        cells = [sum(1 for r in rows if r["s5_group"] == s5g and r["s5c_group"] == g)
                 for g in GROUPS]
        L.append("| %s | %d | %d | %d |" % (s5g, *cells))
    L.append("")
    L.append("**窓が満了していないので、up率の絶対水準は 90日ぶんの確率ではない。**"
             "同じ判定日の中での群間比較として読む（G-5 と同じ）。")
    return "\n".join(L) + "\n"


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--db", default=C.DB_PATH)
    p.add_argument("--projection", default=os.path.join(C.DATA_DIR, "projection.db"))
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "s5c_report.md"))
    p.add_argument("--csv-out", default=os.path.join(C.DATA_DIR, "s5c_observations.csv"))
    a = p.parse_args(argv)

    mcon = sqlite3.connect("file:%s?mode=ro" % a.db.replace("\\", "/"), uri=True)
    pcon = sqlite3.connect("file:%s?mode=ro" % a.projection.replace("\\", "/"), uri=True)
    pcon.row_factory = sqlite3.Row
    cal = [r[0] for r in mcon.execute("SELECT date FROM market_index ORDER BY date")]
    dates = G.judgment_dates(cal)
    codes = [r[0] for r in mcon.execute(
        "SELECT code FROM companies WHERE universe_flag=1 ORDER BY code")]
    C.log("シャドウI: 判定日 %d × 銘柄 %d" % (len(dates), len(codes)))
    rows, reasons = collect(mcon, pcon, dates, codes, G.revisions_by_code(mcon))
    with open(a.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    text = render(rows, reasons, dates, codes)
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    C.log("出力: %s / %s" % (a.out, a.csv_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
