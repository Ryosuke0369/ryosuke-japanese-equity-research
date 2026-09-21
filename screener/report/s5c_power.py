"""screener/report/s5c_power.py — シャドウI-v2 追補: 5年サンプルでの予測力。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウI-v2 追補」（2026-09-21・測定前に確定）。
**このスクリプトは定義を実装するだけで、閾値・窓・群分けを決めない。**

  J-1 修正イベント = `fins_summary` の通期営業利益予想 `f_op` が前の行から変わった開示
  J-2 予測力      = 判定日（各月の最初の営業日）ごとに R を計算し、90日窓の up 発生率を比べる
  J-3 R の水準別  = precision（層の中の up率）と recall（up 全体に占める割合）。**診断のみ**
  J-4 保守的な会社 = 過去中央値進捗率 100%超 / 90-100% / 90%未満

**S5c は 0点・表示のみ。** スコア・採否・出口には接続しない。

    python -m screener.report.s5c_power
"""
from __future__ import annotations

import argparse
import csv
import math
import os
import sqlite3
import statistics
import sys
from collections import Counter, defaultdict
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.signals import progress_history as P

# ---- J-2 の値。動かさない -----------------------------------------------
WINDOW_DAYS = 90
FIRST_MONTH = (2021, 10)
LAST_MONTH = (2026, 6)
LAST_DATA = "2026-09-18"           # 収集の最終日
R_BINS = ((None, 1.0, "R<1.0"), (1.0, 1.1, "1.0-1.1"), (1.1, 1.3, "1.1-1.3"),
          (1.3, 1.5, "1.3-1.5"), (1.5, 2.0, "1.5-2.0"), (2.0, None, "2.0+"))
MIN_N = 100                        # §4: 1群100件未満は判定不能


def judgment_dates(cal: list[str]) -> list[str]:
    """各月の最初の営業日（J-2）。"""
    seen, out = set(), []
    for d in cal:
        ym = (int(d[:4]), int(d[5:7]))
        if ym < FIRST_MONTH or ym > LAST_MONTH or ym in seen:
            continue
        seen.add(ym)
        out.append(d)
    return out


def revision_events(con) -> dict:
    """J-1: code -> [(開示日, 方向, 規模)]。`f_op` が前の行から変わった開示。"""
    out = defaultdict(list)
    last = {}                       # (code, fy_end) -> 直近の f_op
    skipped = Counter()
    for code, d, fy, f_op in con.execute(
            "SELECT code, disc_date, fy_end, f_op FROM fins_summary "
            "WHERE f_op IS NOT NULL AND fy_end IS NOT NULL ORDER BY disc_date, code"):
        key = (code, fy)
        prev = last.get(key)
        last[key] = f_op
        if prev is None:
            skipped["initial（前の予想が無い）"] += 1
            continue
        if prev == 0:
            skipped["前回予想が0"] += 1
            continue
        if f_op == prev:
            skipped["据置"] += 1
            continue
        out[code].append((d, "up" if f_op > prev else "down", (f_op - prev) / abs(prev)))
    return out, skipped


def outcome(evs, code, as_of, end):
    for d, direction, _mag in evs.get(code, ()):
        if as_of < d <= end:
            return direction
    return "none"


def collect(con, dates, codes, evs):
    rows, reasons = [], Counter()
    for as_of in dates:
        horizon = (date.fromisoformat(as_of) + timedelta(days=WINDOW_DAYS)).isoformat()
        truncated = horizon > LAST_DATA
        for code in codes:
            r = P.evaluate(con, code, as_of)
            if not r["available"]:
                reasons[r["evidence"].split("（")[0][:40]] += 1
                continue
            rows.append({
                "as_of": as_of, "code": code, "period": r["period"],
                "r": r["r"], "progress": r["progress"],
                "median_progress": r["median_progress"],
                "n_prior_years": len(r["prior_years"]),
                "fired": int(r["fired"]), "truncated": int(truncated),
                "outcome": outcome(evs, code, as_of, min(horizon, LAST_DATA)),
            })
    return rows, reasons


# ------------------------------------------------------------------ 集計
def up_rate(rows):
    return (sum(r["outcome"] == "up" for r in rows) / len(rows)) if rows else None


def diff_ci(p1, n1, p0, n0):
    if not n1 or not n0 or p1 is None or p0 is None:
        return None, None, None
    se = math.sqrt(p1 * (1 - p1) / n1 + p0 * (1 - p0) / n0)
    return p1 - p0, (p1 - p0) - 1.96 * se, (p1 - p0) + 1.96 * se


def clustered(rows):
    """判定日ごとの up率の平均と t（J-2）。"""
    by = defaultdict(list)
    for r in rows:
        by[r["as_of"]].append(r["outcome"] == "up")
    ps = [sum(v) / len(v) for v in by.values()]
    if len(ps) < 2:
        return None, None, len(ps)
    m = statistics.mean(ps)
    sd = statistics.stdev(ps)
    return m, (m / (sd / math.sqrt(len(ps))) if sd > 0 else None), len(ps)


def _row(label, sub, base):
    if not sub:
        return "| %s | 0 | – | – | – | – |" % label
    p = up_rate(sub)
    d, lo, hi = diff_ci(p, len(sub), up_rate(base), len(base))
    m, t, nd = clustered(sub)
    note = " ※件数不足" if len(sub) < MIN_N else ""
    return "| %s | %d%s | %d | %.2f%% | %s | %s |" % (
        label, len(sub), note, len({r["code"] for r in sub}), 100 * p,
        "–" if d is None else "%+.2fpt [%+.2f, %+.2f]" % (100 * d, 100 * lo, 100 * hi),
        "%.2f%% / t %s (%d日)" % (100 * m, "–" if t is None else "%.2f" % t, nd))


def render(rows, reasons, skipped, dates, evs):
    L = ["# シャドウI-v2 追補: 5年サンプルでの予測力（記録専用・0点）", ""]
    n_ev = sum(len(v) for v in evs.values())
    ups = sum(1 for v in evs.values() for e in v if e[1] == "up")
    L.append("- 判定日 %d（%s 〜 %s）/ 観測 %s（R が計算できたものだけ）"
             % (len(dates), dates[0], dates[-1], f"{len(rows):,}"))
    L.append("- 修正イベント **%s 件**（up %s / down %s）。据置 %s / initial %s / 前回0 %s"
             % (f"{n_ev:,}", f"{ups:,}", f"{n_ev - ups:,}",
                f"{skipped['据置']:,}", f"{skipped['initial（前の予想が無い）']:,}",
                f"{skipped['前回予想が0']:,}"))
    L.append("")
    L.append("## 修正イベントの内訳（J-1）")
    L.append("")
    by_year = Counter()
    for v in evs.values():
        for d, direction, _m in v:
            by_year[(d[:4], direction)] += 1
    L.append("| 年 | up | down | 計 |")
    L.append("|---|---|---|---|")
    for y in sorted({y for y, _ in by_year}):
        L.append("| %s | %d | %d | %d |"
                 % (y, by_year[(y, "up")], by_year[(y, "down")],
                    by_year[(y, "up")] + by_year[(y, "down")]))
    L.append("")
    L.append("## R を計算できなかった観測の理由")
    L.append("")
    L.append("| 理由 | のべ件数 |")
    L.append("|---|---|")
    for why, n in reasons.most_common():
        L.append("| %s | %s |" % (why, f"{n:,}"))
    L.append("")
    L.append("## J-2. 予測力（母集団 = R を計算できた観測）")
    L.append("")
    hdr = ("| 群 | のべn | 銘柄 | up率 | 母集団との差 [95%CI] | 判定日クラスタ |\n"
           "|---|---|---|---|---|---|")
    L.append(hdr)
    fired = [r for r in rows if r["fired"]]
    notf = [r for r in rows if not r["fired"]]
    L.append(_row("R ≥ 1.30", fired, rows))
    L.append(_row("R < 1.30", notf, rows))
    L.append(_row("母集団（全体）", rows, rows))
    L.append("")
    full = [r for r in rows if not r["truncated"]]
    if full and len(full) != len(rows):
        L.append("### 窓が満了した観測だけ（%s 件）" % f"{len(full):,}")
        L.append("")
        L.append(hdr)
        L.append(_row("R ≥ 1.30", [r for r in full if r["fired"]], full))
        L.append(_row("R < 1.30", [r for r in full if not r["fired"]], full))
        L.append(_row("母集団（全体）", full, full))
        L.append("")
    L.append("## J-3. R の水準別（**診断のみ・閾値は動かさない**）")
    L.append("")
    L.append("| R の帯 | のべn | 銘柄 | up率（precision） | up 全体に占める割合（recall） |")
    L.append("|---|---|---|---|---|")
    n_up_all = sum(1 for r in rows if r["outcome"] == "up")
    for lo, hi, label in R_BINS:
        sub = [r for r in rows
               if (lo is None or r["r"] >= lo) and (hi is None or r["r"] < hi)]
        if not sub:
            continue
        n_up = sum(1 for r in sub if r["outcome"] == "up")
        L.append("| %s | %s | %d | %.2f%% | %.1f%% |"
                 % (label, f"{len(sub):,}", len({r["code"] for r in sub}),
                    100 * n_up / len(sub), 100 * n_up / n_up_all if n_up_all else 0))
    L.append("")
    L.append("## J-4. 保守的な会社（過去中央値進捗率で分ける）")
    L.append("")
    L.append(hdr)
    for label, lo, hi in (("100%超", 1.0, None), ("90-100%", 0.9, 1.0),
                          ("90%未満", None, 0.9)):
        sub = [r for r in rows
               if (lo is None or r["median_progress"] > lo)
               and (hi is None or r["median_progress"] <= hi)]
        L.append(_row(label, sub, rows))
    L.append("")
    L.append("**S5c は 0点・表示のみ。この節の数値でスコア・採否・出口を変えない。**")
    return "\n".join(L) + "\n"


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--db", default=C.DB_PATH)
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "s5c_power_report.md"))
    p.add_argument("--csv-out", default=os.path.join(C.DATA_DIR, "s5c_power.csv"))
    a = p.parse_args(argv)

    con = sqlite3.connect("file:%s?mode=ro" % a.db.replace("\\", "/"), uri=True)
    cal = [r[0] for r in con.execute("SELECT date FROM market_index ORDER BY date")]
    dates = judgment_dates(cal)
    codes = [r[0] for r in con.execute(
        "SELECT code FROM companies WHERE universe_flag=1 ORDER BY code")]
    evs, skipped = revision_events(con)
    C.log("修正イベント %d 件 / 判定日 %d × 銘柄 %d"
          % (sum(len(v) for v in evs.values()), len(dates), len(codes)))
    rows, reasons = collect(con, dates, codes, evs)
    C.log("R を計算できた観測 %d" % len(rows))
    with open(a.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    text = render(rows, reasons, skipped, dates, evs)
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    C.log("出力: %s / %s" % (a.out, a.csv_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
