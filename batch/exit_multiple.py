"""Derive exit_multiple from the ticker's own comps CSV (no cross-ticker reuse).

    exit_multiple = median(peer EV/EBITDA)      rounded to 1 dp

Deliberately NOT margin-adjusted. An earlier version scaled the peer median by
half the subject's operating-margin premium; it was rejected because EV/EBITDA
already embeds capital intensity, so an OPM-based uplift is directionally wrong
for capital-heavy businesses - it produced 16.3x for 9022 (41% OPM) against the
7.4x that name actually trades at. The plain peer median is uniform, has no
invented coefficients, and never references the subject's own share price.

Type B names are capped at the peer median by 4-4; with this rule that is the
same number, so the cap never binds.
"""
import sys, csv, os, statistics as st
sys.stdout.reconfigure(encoding="utf-8")

def derive(code, quiet=False):
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "data", "comps", f"{code}_comps.csv")
    rows = list(csv.DictReader(open(path, encoding="utf-8")))
    def num(r, k):
        v = r[k].strip()
        return float(v) if v else None
    def ev_eb(r):
        e, m, n = num(r, "EBITDA"), num(r, "Market_Cap"), num(r, "Net_Debt")
        return None if (not e or m is None or n is None) else (m + n) / e
    subj, peers = rows[0], rows[1:]
    vals = [(r["Ticker"], ev_eb(r)) for r in peers]
    good = [v for _, v in vals if v]
    med = st.median(good)
    if not quiet:
        for t, v in vals:
            print(f"  {t:<9}{'EV/EBITDA n/a' if not v else f'EV/EBITDA {v:6.2f}x'}")
        print(f"{code}: peer median EV/EBITDA = {med:.2f}x (n={len(good)}) "
              f"-> exit_multiple {round(med,1)}x   [subject own {ev_eb(subj):.2f}x, 算式には不使用]")
    return round(med, 1)

if __name__ == "__main__":
    for c in [a for a in sys.argv[1:] if not a.startswith("--")]:
        derive(c)
