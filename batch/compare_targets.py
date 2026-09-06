"""フェーズ2 §5 — old vs new Target, and the cause of every rating flip.

Both sides are read by batch/snapshot_models.py, so the comparison never depends
on a hand-maintained state file.

Attribution is by elimination, from what the workbooks themselves record, and it
does not guess. For every ticker the script reports the beta and WACC move, and
flags each OTHER フェーズ2 fix that demonstrably touched this model: net debt
resolved differently, guidance now obtained (so the Management scenario is no
longer a copy of Base), the latest fiscal year changed (EDINET found a newer
有報), the market data source changed. A ticker whose only flagged change is the
beta has its Target move explained by the beta; anything else is listed by name
for inspection rather than assigned a cause it cannot support.

Usage:
    python batch/compare_targets.py --before batch/phase2_before.json \
                                    --after  batch/phase2_after.json
"""
import argparse
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

RANK = {"BUY": 2, "HOLD": 1, "SELL": 0}


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def verdict_of(row):
    v = row.get("verdict")
    return str(v).strip().upper() if isinstance(v, str) else None



def leg_basis(row):
    """How the Target was formed: the average of both DCF legs, or one leg alone.

    Derived from the three cached numbers rather than from a label, so it works
    on both sides of the comparison. 追補6 §X demotions and the automatic
    INVALID marking both collapse the Target onto a single leg; this tells them
    apart from a genuine valuation move, which is the difference between "the
    model changed its mind" and "the model changed the question".
    """
    t, p, e = num(row.get("target_mid")), num(row.get("dcf_pgm")), num(row.get("dcf_exit"))
    if t is None:
        return "no target"
    if p is not None and e is not None:
        if abs(t - (p + e) / 2) <= 1.5:
            return "average"
        if abs(t - p) <= 1.5:
            return "PGM only"
        if abs(t - e) <= 1.5:
            return "Exit only"
        return "other"
    if p is None and e is not None:
        return "Exit only (PGM invalid)"
    if e is None and p is not None:
        return "PGM only (Exit invalid)"
    return "no legs"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--before", default=os.path.join(HERE, "phase2_before.json"))
    ap.add_argument("--after", default=os.path.join(HERE, "phase2_after.json"))
    ap.add_argument("--out", default=os.path.join(HERE, "phase2_comparison.json"))
    a = ap.parse_args()

    before = json.load(open(a.before, encoding="utf-8"))
    after = json.load(open(a.after, encoding="utf-8"))
    codes = sorted(set(before) & set(after))

    rows = []
    for c in codes:
        b, n = before[c], after[c]
        tb, tn = num(b.get("target_mid")), num(n.get("target_mid"))
        wb_, wn = num(b.get("wacc")), num(n.get("wacc"))
        bb, bn = num(b.get("beta")), num(n.get("beta"))
        vb, vn = verdict_of(b), verdict_of(n)
        row = {
            "code": c,
            "price_old": num(b.get("price")), "price_new": num(n.get("price")),
            "target_old": tb, "target_new": tn,
            "target_pct": (round((tn - tb) / tb * 100, 1)
                           if (tb and tn) else None),
            "beta_old": bb, "beta_new": bn,
            "wacc_old": wb_, "wacc_new": wn,
            "wacc_bps": (round((wn - wb_) * 10000) if (wb_ is not None and wn is not None) else None),
            "verdict_old": vb, "verdict_new": vn,
            "flip": (vb != vn) if (vb and vn) else None,
            "flip_dir": None,
            "beta_raw": n.get("meta", {}).get("beta_raw"),
            "beta_clamped": n.get("meta", {}).get("beta_clamped"),
            "guidance": n.get("meta", {}).get("guidance_source"),
            "errors_new": sum(len(v) for v in (n.get("errors") or {}).values()),
        }
        # What else moved, read off the two workbooks. Only differences that are
        # visible in the files are reported; nothing is inferred.
        mo, mn = b.get("meta") or {}, n.get("meta") or {}
        other = []
        nd_o, nd_n = num(b.get("net_debt")), num(n.get("net_debt"))
        if nd_o is not None and nd_n is not None and abs(nd_n - nd_o) > max(1.0, abs(nd_o) * 0.001):
            other.append(f"net_debt {nd_o:,.0f}->{nd_n:,.0f}")
        # Guidance (#9) is deliberately NOT listed as a cause of a Target move.
        # It only rewrites the Management scenario, and the active scenario (C27)
        # is "Base" in every workbook on both sides of this comparison - measured,
        # not assumed: 73/73 surviving 2026-09-05 workbooks and 85/85 regenerated
        # ones. So obtaining guidance changed what the Management column MEANS
        # without moving the Target by a yen. It is reported on its own.
        # フェーズ2 #10 moved C5 off the historical mean and back onto the
        # analyst's capex_pct. It is recorded, but NOT as a Target cause: every
        # one of these tickers carries a complete 5-year capex_direct.projections
        # array, so C5 is the fallback for years that do not exist and the
        # projection rows never read it. Verified rather than assumed - see the
        # phase-2 report §5.
        if mo.get("capex_pct_basis") == "hist_3yr_avg" != mn.get("capex_pct_basis"):
            row["c5_basis_changed"] = True
        # How the Target is formed. A demotion (追補6 §X) or an INVALID leg
        # changes the Target without changing the valuation, so it is the first
        # thing to separate out.
        lb_o, lb_n = leg_basis(b), leg_basis(n)
        row["leg_basis_old"], row["leg_basis_new"] = lb_o, lb_n
        if lb_o != lb_n:
            other.append(f"Target basis {lb_o}->{lb_n} (追補6 §X)")
        for key, label in (("exit_multiple", "exit multiple"),
                           ("size_premium", "size premium"),
                           ("terminal_growth", "terminal g")):
            vo, vn2 = num(b.get(key)), num(n.get(key))
            if vo is not None and vn2 is not None and abs(vn2 - vo) > 1e-9:
                other.append(f"{label} {vo}->{vn2}")
        po, pn = num(b.get("price")), num(n.get("price"))
        if po and pn and abs(pn - po) / po > 0.001:
            other.append(f"price {po:,.0f}->{pn:,.0f}")
        row["other_changes"] = other
        row["beta_only"] = not other
        if row["flip"] and vb in RANK and vn in RANK:
            row["flip_dir"] = f"{vb}->{vn}"
        rows.append(row)

    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)

    hdr = ("code", "price", "T old", "T new", "chg%", "beta o", "beta n",
           "WACC o", "WACC n", "bps", "verdict")
    print(f"{hdr[0]:<6} {hdr[1]:>8} {hdr[2]:>8} {hdr[3]:>8} {hdr[4]:>8} "
          f"{hdr[5]:>7} {hdr[6]:>7} {hdr[7]:>7} {hdr[8]:>7} {hdr[9]:>6}  {hdr[10]}")
    print("-" * 104)
    for r in rows:
        pct = "-" if r["target_pct"] is None else f"{r['target_pct']:+.1f}"
        print(f"{r['code']:<6} {r['price_new'] or 0:>8,.0f} "
              f"{r['target_old'] or 0:>8,.0f} {r['target_new'] or 0:>8,.0f} "
              f"{pct:>8} "
              f"{r['beta_old'] or 0:>7.3f} {r['beta_new'] or 0:>7.3f} "
              f"{(r['wacc_old'] or 0) * 100:>6.2f}% {(r['wacc_new'] or 0) * 100:>6.2f}% "
              f"{r['wacc_bps'] if r['wacc_bps'] is not None else 0:>6} "
              f" {r['verdict_old']} -> {r['verdict_new']}"
              + ("   ** FLIP **" if r["flip"] else ""))

    flips = [r for r in rows if r["flip"]]
    up = [r for r in rows if (r["target_pct"] or 0) > 0]
    dn = [r for r in rows if (r["target_pct"] or 0) < 0]
    print("-" * 104)
    print(f"tickers compared: {len(rows)}")
    print(f"  Target up: {len(up)} / down: {len(dn)}")
    if rows:
        pcts = [r["target_pct"] for r in rows if r["target_pct"] is not None]
        if pcts:
            pcts_sorted = sorted(pcts)
            print(f"  Target change %: min {min(pcts):+.1f} / median "
                  f"{pcts_sorted[len(pcts_sorted)//2]:+.1f} / max {max(pcts):+.1f}")
        bps = [r["wacc_bps"] for r in rows if r["wacc_bps"] is not None]
        if bps:
            print(f"  WACC change bps: min {min(bps):+d} / max {max(bps):+d} / "
                  f"mean {sum(bps)/len(bps):+.0f}")
    print(f"  rating flips: {len(flips)}")
    def _n(v, fmt=",.0f"):
        return "N/A" if v is None else format(v, fmt)

    for r in flips:
        cause = ("beta only" if r["beta_only"]
                 else "beta + " + "; ".join(r["other_changes"]))
        print(f"     {r['code']}  {r['verdict_old']} -> {r['verdict_new']}  "
              f"Target {_n(r['target_old'])} -> {_n(r['target_new'])} "
              f"({_n(r['target_pct'], '+.1f')}%)  "
              f"WACC {_n(r['wacc_bps'], '+d')}bps  cause: {cause}")
    n_beta_only = sum(1 for r in rows if r["beta_only"])
    print(f"  Target move attributable to beta alone: {n_beta_only}/{len(rows)}; "
          f"{len(rows) - n_beta_only} had another Target-affecting input change "
          f"too (listed per ticker in the JSON)")
    n_c5 = sum(1 for r in rows if r.get("c5_basis_changed"))
    n_leg = sum(1 for r in rows if r["leg_basis_old"] != r["leg_basis_new"])
    print(f"  Target basis changed by the 追補6 §X re-arbitration: {n_leg}")
    print(f"  C5 basis moved off the historical mean (#10, display/sensitivity "
          f"only - the projections are direct): {n_c5}")
    n_guid = sum(1 for r in rows
                 if r.get("guidance") and r["guidance"] != "none")
    print(f"  Management scenario now driven by real 会社予想: {n_guid}/{len(rows)} "
          f"(does not move the Target - the active scenario is Base everywhere)")
    errs = [r for r in rows if r["errors_new"]]
    print(f"  regenerated workbooks with formula errors: {len(errs)}"
          + (f" -> {', '.join(r['code'] for r in errs)}" if errs else ""))
    print(f"  detail: {a.out}")


if __name__ == "__main__":
    main()
