"""screener/report/rescore_diff.py — 全銘柄再スコア CSV の要約と、過去の週次 CSV との差分。

    python -m screener.report.rescore_diff --new C:/screener_data/weekly_20260913_full_rescored.csv \
        --old C:/screener_data/weekly_20260902_v3.csv C:/screener_data/weekly_20260913_earnings_window.csv \
        --floor 0.10

出すもの（数値のみ・判断はしない）
  1. シグナル別の発火件数・発火率（分母 = 採点対象銘柄数）。S1/S2/S4/S12
  2. 合成スコア上位30（コード/名/スコア/発火/doc_period/doc_date/根拠期が直前四半期か/信頼性フラグ）
  3. 旧CSVとの差分: 消えた / 新規 / 0.2 以上動いた
  4. スコア未付与・閾値未満の理由内訳
  5. サニティ: 3475・2776 が消えるか警告付き、3441 の S1 が負側で残るか、5136 が健全な根拠で残るか
"""
from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter

SIGNALS = ("S1", "S2", "S4", "S12")


def _f(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def load(path):
    with open(path, encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


def fire_rates(rows):
    n = len(rows)
    out = {}
    for sig in SIGNALS:
        k = sum(1 for r in rows if sig in (r.get("fired") or "").split(","))
        out[sig] = (k, round(100.0 * k / n, 2) if n else 0.0)
    return n, out


def diff(new, old, floor):
    nn = {r["code"]: r for r in new if (_f(r.get("score")) or -9) >= floor}
    oo = {r["code"]: r for r in old if (_f(r.get("score")) or -9) >= floor}
    allnew = {r["code"]: r for r in new}
    gone = sorted(set(oo) - set(nn))
    added = sorted(set(nn) - set(oo))
    moved = []
    for c in set(oo) & set(allnew):
        a, b = _f(oo[c].get("score")), _f(allnew[c].get("score"))
        if a is not None and b is not None and abs(b - a) >= 0.2:
            moved.append((c, a, b))
    return gone, added, sorted(moved, key=lambda x: x[2] - x[1]), allnew


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--new", required=True)
    p.add_argument("--old", nargs="*", default=[])
    p.add_argument("--floor", type=float, default=0.10)
    p.add_argument("--top", type=int, default=30)
    a = p.parse_args(argv)
    new = load(a.new)

    n, rates = fire_rates(new)
    print("== 1. シグナル別発火（分母 %d 銘柄）==" % n)
    for sig, (k, pct) in rates.items():
        verdict = "50%超" if pct > 50 else ("0.1%未満" if pct < 0.1 else "")
        print("  %-4s 発火 %5d 件  %6.2f%% %s" % (sig, k, pct, verdict))
    scored = [r for r in new if _f(r.get("score")) is not None]
    print("  スコア付与 %d / 閾値 %.2f 以上 %d" % (len(scored), a.floor,
                                             sum(1 for r in scored if _f(r["score"]) >= a.floor)))

    print("== 2. 合成スコア上位%d ==" % a.top)
    top = sorted(scored, key=lambda r: -_f(r["score"]))[:a.top]
    print("  code | name | score | fired | doc_period | doc_date | 直前四半期(lag) | 信頼性・注意")
    for r in top:
        rel = [x for x in (r.get("reliability_flag"),
                           "GC" if r.get("going_concern") == "1" else "",
                           "会計変更" if r.get("accounting_change") == "1" else "",
                           ("注記:" + r["strict_notes"]) if r.get("strict_notes") else "") if x]
        recent = {"1": "はい", "0": "いいえ"}.get(str(r.get("evidence_recent")), "-")
        print("  %s | %s | %s | %s | %s | %s | %s(%s) | %s" % (
            r["code"], r["name"][:14], r["score"], r["fired"], r.get("doc_period") or "-",
            r.get("doc_date") or "-", recent, r.get("evidence_lag_q") or "-", " / ".join(rel) or "-"))

    for path in a.old:
        old = load(path)
        gone, added, moved, allnew = diff(new, old, a.floor)
        print("== 3. 差分 vs %s（旧 閾値以上 %d 銘柄）==" % (path, sum(1 for r in old if (_f(r.get("score")) or -9) >= a.floor)))
        print("  消えた %d: %s" % (len(gone), ", ".join(
            "%s(%s→%s %s)" % (c, next(x["score"] for x in old if x["code"] == c),
                              allnew.get(c, {}).get("score", "対象外"),
                              allnew.get(c, {}).get("no_score_reason") or allnew.get(c, {}).get("false_positive_flags", ""))
            for c in gone)))
        print("  新規 %d: %s" % (len(added), ", ".join("%s(%s)" % (c, allnew[c]["score"]) for c in added[:40])))
        print("  0.2以上動いた %d: %s" % (len(moved), ", ".join("%s %.3f→%.3f" % m for m in moved[:40])))

    print("== 4. スコア未付与・閾値未満の理由 ==")
    reasons = Counter((r.get("no_score_reason") or "").split("(")[0] or "(閾値以上)" for r in new)
    for k, v in reasons.most_common():
        print("  %-40s %d" % (k, v))
    detail = Counter()
    for r in new:
        nsr = r.get("no_score_reason") or ""
        if nsr.startswith("証拠不採用("):
            for f in nsr[len("証拠不採用("):-1].split(","):
                detail[f] += 1
    if detail:
        print("  証拠不採用の内訳（銘柄×種別）: %s" % dict(detail))

    print("== 5. サニティ ==")
    by = {r["code"]: r for r in new}
    for c, expect in (("3475", "消える/警告"), ("2776", "消える/警告"), ("3441", "S1 負側で減点"),
                      ("5136", "健全な根拠で残る")):
        r = by.get(c)
        if not r:
            print("  %s: 出力に無い" % c)
            continue
        print("  %s [%s] score=%s fired=%s doc_period=%s doc_date=%s lag=%s flags=%s reason=%s" % (
            c, expect, r["score"], r["fired"] or "-", r.get("doc_period") or "-", r.get("doc_date") or "-",
            r.get("evidence_lag_q") or "-", r.get("false_positive_flags") or "-", r.get("no_score_reason") or "-"))
        if c == "3441":
            ev = [x for x in (r.get("evidence") or "").split(" ｜ ") if x.startswith("S1")]
            print("      S1 evidence(発火時のみ出る): %s" % (ev or "非発火"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
