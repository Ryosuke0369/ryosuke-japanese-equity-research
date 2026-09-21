"""screener/report/progress_profile.py — シャドウI-v2 の分布（I2-4 の 1・2・4）。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウI-v2」。
**このスクリプトは定義を実装するだけで、閾値を決めない。記録専用（0点・表示のみ）。**

出すもの:
  1. 過去中央値進捗率そのものの分布（四半期別）と、**100%超の銘柄**
     （= 期の途中で通期予想に届いてしまう＝常習的に保守的な予想を出す会社の代理変数）
  2. R（当期進捗率 ÷ 過去中央値進捗率）の分布と、R ≥ 1.30 の件数・銘柄
  4. **生の進捗率で並べた上位20 と、R で並べた上位20 を並べて出す**（誤読の実例を残す）

    python -m screener.report.progress_profile --as-of 2026-09-18
"""
from __future__ import annotations

import argparse
import csv
import os
import sqlite3
import statistics
import sys
from collections import Counter

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.signals import progress_history as P
from screener.technical.event_response import percentile as pct


def build(con, codes, as_of):
    """銘柄ごとに R と、四半期別の過去中央値進捗率を作る。"""
    rows, reasons = [], Counter()
    med_by_q = {q: {} for q in P.QUARTERS}
    for code in codes:
        raw = P.load_rows(con, code, as_of)
        series = P.progress_series(raw)
        for q in P.QUARTERS:
            med, years = P.median_progress(series, q)
            if med is not None:
                med_by_q[q][code] = (med, len(years))
        r = P.evaluate(con, code, as_of)
        if not r["available"]:
            reasons[r["evidence"].split("（")[0][:44]] += 1
            continue
        rows.append({"code": code, "period": r["period"], "fy_end": r["fy_end"],
                     "disc_date": r["disc_date"], "progress": r["progress"],
                     "median_progress": r["median_progress"],
                     "n_prior_years": len(r["prior_years"]), "r": r["r"],
                     "r_sales": r["r_sales"], "fired": int(r["fired"]),
                     "sales_fired": int(r["sales_fired"]),
                     "cum_op": r["cum_op"], "forecast_op": r["forecast_op"],
                     "landing_op": r["landing_op"], "evidence": r["evidence"]})
    return rows, reasons, med_by_q


def _dist(vals, fmt="%.1f%%", scale=100):
    xs = sorted(vals)
    if not xs:
        return "–"
    return " / ".join(fmt % (scale * pct(xs, q)) for q in (0.10, 0.25, 0.50, 0.75, 0.90))


def render(rows, reasons, med_by_q, names, as_of, n_codes):
    L = ["# シャドウI-v2: 進捗率の季節性と R の分布（記録専用・0点）", ""]
    L.append("- as_of %s / 対象 %d 銘柄 / R を計算できた **%d 銘柄**" % (as_of, n_codes, len(rows)))
    L.append("- R = 当期進捗率 ÷ 過去中央値進捗率 = 推定着地OP ÷ 会社予想OP。発火は **R ≥ %.2f**"
             % P.FIRE_R)
    L.append("")
    L.append("## 計算できなかった理由")
    L.append("")
    L.append("| 理由 | 銘柄数 |")
    L.append("|---|---|")
    for why, n in reasons.most_common():
        L.append("| %s | %d |" % (why, n))
    L.append("")
    L.append("## 1. 過去中央値進捗率の分布（四半期別・10/25/50/75/90 パーセンタイル）")
    L.append("")
    L.append("| 四半期 | 銘柄 | 分布 | 100%超の銘柄 |")
    L.append("|---|---|---|---|")
    for q in P.QUARTERS:
        d = med_by_q[q]
        over = [c for c, (m, _n) in d.items() if m > 1.0]
        L.append("| %s | %d | %s | **%d（%.1f%%）** |"
                 % (q, len(d), _dist([m for m, _n in d.values()]), len(over),
                    100 * len(over) / len(d) if d else 0))
    L.append("")
    L.append("### 100%超の銘柄（＝期の途中で通期予想に届く＝保守的な予想の代理変数）")
    L.append("")
    for q in P.QUARTERS:
        over = sorted(((m, c) for c, (m, _n) in med_by_q[q].items() if m > 1.0),
                      reverse=True)
        if not over:
            continue
        L.append("**%s（%d銘柄）**: " % (q, len(over)) + " / ".join(
            "%s %s %.0f%%" % (c, names.get(c, ""), 100 * m) for m, c in over[:15])
            + (" …" if len(over) > 15 else ""))
        L.append("")
    L.append("## 2. R の分布")
    L.append("")
    L.append("| 量 | 10% | 25% | 50% | 75% | 90% | 95% | 99% |")
    L.append("|---|---|---|---|---|---|---|---|")
    rs = sorted(r["r"] for r in rows)
    L.append("| R（営業利益） | " + " | ".join(
        "%.2f" % pct(rs, q) for q in (0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99)) + " |")
    ps = sorted(r["progress"] for r in rows)
    L.append("| 生の進捗率 | " + " | ".join(
        "%.0f%%" % (100 * pct(ps, q)) for q in (0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99)) + " |")
    L.append("")
    fired = [r for r in rows if r["fired"]]
    L.append("**R ≥ %.2f: %d 銘柄（%.1f%%）** / 売上ライン（R_sales ≥ %.2f）: %d 銘柄"
             % (P.FIRE_R, len(fired), 100 * len(fired) / len(rows) if rows else 0,
                P.NOTE_R_SALES, sum(r["sales_fired"] for r in rows)))
    L.append("")
    L.append("## 4. 生の進捗率で並べた上位20 と、R で並べた上位20")
    L.append("")
    L.append("| # | 進捗率で並べる | 進捗 | 過去中央値 | R | | R で並べる | 進捗 | 過去中央値 | R |")
    L.append("|---|---|---|---|---|---|---|---|---|---|")
    by_p = sorted(rows, key=lambda r: -r["progress"])[:20]
    by_r = sorted(rows, key=lambda r: -r["r"])[:20]
    for i in range(20):
        a = by_p[i] if i < len(by_p) else None
        b = by_r[i] if i < len(by_r) else None
        def cell(x):
            if not x:
                return " | | | "
            return "%s %s | %.0f%% | %.0f%% | %.2f" % (
                x["code"], names.get(x["code"], "")[:8], 100 * x["progress"],
                100 * x["median_progress"], x["r"])
        L.append("| %d | %s | | %s |" % (i + 1, cell(a), cell(b)))
    L.append("")
    L.append("**同じ銘柄が両方に出るとは限らない。** 生の進捗率の上位は季節性の強い会社が占める。")
    return "\n".join(L) + "\n"


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--db", default=C.DB_PATH)
    p.add_argument("--as-of", default=None, help="この日までの開示だけを見る（PIT）")
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "progress_profile.md"))
    p.add_argument("--csv-out", default=os.path.join(C.DATA_DIR, "progress_profile.csv"))
    a = p.parse_args(argv)

    con = sqlite3.connect("file:%s?mode=ro" % a.db.replace("\\", "/"), uri=True)
    codes = [r[0] for r in con.execute(
        "SELECT code FROM companies WHERE universe_flag=1 ORDER BY code")]
    names = dict(con.execute("SELECT code, name FROM companies"))
    C.log("進捗率の分布: %d 銘柄（as_of %s）" % (len(codes), a.as_of or "指定なし"))
    rows, reasons, med_by_q = build(con, codes, a.as_of)
    if rows:
        with open(a.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
    text = render(rows, reasons, med_by_q, names, a.as_of or "指定なし", len(codes))
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    C.log("出力: %s / %s" % (a.out, a.csv_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
