"""screener/report/policy_shadow_compare.py — 採点方針の前後を同じ投影DBで突き合わせる（シャドウ比較）。

入力は evidence_audit.py の出力2組（シグナル単位CSV と _tickers.csv）。
既定の切替（calibration_backlog §31）の前に、影響銘柄数と代表例を出すためのもの。

    python -m screener.report.policy_shadow_compare \
        --before C:/screener_data/evidence_audit_20260913_shadow_prefer_span.csv \
        --after  C:/screener_data/evidence_audit_20260913_shadow_evidence_strict.csv \
        --out    C:/screener_data/policy_shadow_20260913.csv
"""
from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter

FLOOR = 0.10


def _load(path):
    with open(path, encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


def _f(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def compare(before_sig, after_sig, before_tick, after_tick, watch=("3475", "2776", "5136", "3441")):
    bt = {r["code"]: r for r in before_tick}
    at = {r["code"]: r for r in after_tick}
    rows = []
    for code in sorted(set(bt) | set(at)):
        b, a = bt.get(code, {}), at.get(code, {})
        sb, sa = _f(b.get("evidence_score")), _f(a.get("evidence_score"))
        hb, ha = sb is not None and sb >= FLOOR, sa is not None and sa >= FLOOR
        if hb and not ha:
            change = "閾値以上→消えた"
        elif ha and not hb:
            change = "新規に閾値以上"
        elif hb and ha:
            change = "閾値以上のまま"
        else:
            change = "閾値未満のまま"
        rows.append({"code": code, "name": b.get("name") or a.get("name", ""),
                     "score_before": "" if sb is None else sb, "score_after": "" if sa is None else sa,
                     "delta": "" if (sb is None or sa is None) else round(sa - sb, 3),
                     "fired_before": b.get("fired", ""), "fired_after": a.get("fired", ""),
                     "change": change})
    flags = Counter()
    flag_codes = {}
    for r in after_sig:
        for f in (r.get("strict_flags") or "").split(";"):
            if f:
                k = f.split("(")[0]
                k = "stale_lag2+" if k.startswith("stale_lag") else k
                flags[(r["signal"], k)] += 1
                flag_codes.setdefault(k, set()).add(r["code"])
    excl_fired = Counter()
    bsig = {(r["code"], r["signal"]): r for r in before_sig}
    for r in after_sig:
        b = bsig.get((r["code"], r["signal"]))
        if b and b["fired"] == "1" and r["fired"] != "1":
            excl_fired[r["signal"]] += 1
    watch_rows = {c: next((x for x in rows if x["code"] == c), None) for c in watch}
    return rows, flags, flag_codes, excl_fired, watch_rows


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--before", required=True)
    p.add_argument("--after", required=True)
    p.add_argument("--out", required=True)
    a = p.parse_args(argv)
    bs, as_ = _load(a.before), _load(a.after)
    bt, at = _load(a.before.replace(".csv", "_tickers.csv")), _load(a.after.replace(".csv", "_tickers.csv"))
    rows, flags, flag_codes, excl_fired, watch = compare(bs, as_, bt, at)

    ch = Counter(r["change"] for r in rows)
    print("銘柄数 %d / 変化: %s" % (len(rows), dict(ch)))
    moved = [r for r in rows if r["delta"] != "" and abs(r["delta"]) >= 0.2]
    print("スコアが 0.2 以上動いた銘柄: %d" % len(moved))
    print("スコア付き→無評価: %d / 無評価→スコア付き: %d" % (
        sum(1 for r in rows if r["score_before"] != "" and r["score_after"] == ""),
        sum(1 for r in rows if r["score_before"] == "" and r["score_after"] != "")))
    print("不採用フラグ（シグナル×種別、available だったシグナル単位）:")
    for (sig, k), n in sorted(flags.items()):
        print("  %s %-24s %d" % (sig, k, n))
    print("不採用フラグ種別ごとの銘柄数: %s" % {k: len(v) for k, v in flag_codes.items()})
    print("修正前に発火→修正後に非発火（シグナル単位）: %s" % dict(excl_fired))
    for label in ("閾値以上→消えた", "新規に閾値以上"):
        sel = sorted([r for r in rows if r["change"] == label],
                     key=lambda r: -(_f(r["score_before"]) or _f(r["score_after"]) or 0))
        print("--- %s（上位10）---" % label)
        for r in sel[:10]:
            print("  %s %-14s %s → %s  fired %s → %s" % (r["code"], r["name"][:14], r["score_before"],
                                                        r["score_after"], r["fired_before"], r["fired_after"]))
    print("--- 監視銘柄 ---")
    for c, r in watch.items():
        print("  %s %s" % (c, r))
    with open(a.out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print("CSV: %s" % a.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
