"""Build data/comps/<code>_comps.csv from the yfinance cache.

Usage: python batch/build_comps.py <subject> <peer> <peer> ...
Self row first. EBITDA = Operating Income + D&A (cash-flow D&A) for every row;
a row whose D&A is unavailable is written with a blank EBITDA (contract).
"""
import sys, os, json, csv
sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "cache")
OUT = os.path.join(os.path.dirname(HERE), "data", "comps")

def load(c):
    return json.load(open(os.path.join(CACHE, f"{c}.json"), encoding="utf-8"))

def latest_year(d):
    ks = set()
    for blk in ("income", "balance", "cashflow"):
        for v in d[blk].values():
            ks |= set(v.keys())
    return sorted(ks)[-1]

def pick(d, blk, names, y, back=2, note=None):
    """Value for fiscal year `y`; falls back to the most recent earlier year.

    A block may not carry year `y` at all (yfinance often lags a year on the cash
    flow statement). Whenever the value actually used comes from a DIFFERENT year
    than `y`, the year is appended to `note` so the row never mixes fiscal years
    silently - the caller prints it.
    """
    ys = sorted({k for v in d[blk].values() for k in v})
    if not ys:
        return None
    cand = [k for k in ys if k <= y][-(back + 1):] or ys[:1]
    for k in reversed(cand):
        for n in names:
            v = d[blk].get(n, {}).get(k)
            if v is not None:
                if k != y and note is not None:
                    note.append(f"{names[0]}<-{k[:7]}")
                return v
    return None


def row(c):
    d = load(c)
    y = latest_year(d)
    mix = []
    rev = pick(d, "income", ["Total Revenue", "Operating Revenue"], y, note=mix)
    oi  = pick(d, "income", ["Operating Income", "Total Operating Income As Reported"], y, note=mix)
    ni  = pick(d, "income", ["Net Income Common Stockholders", "Net Income"], y, note=mix)
    da  = pick(d, "cashflow", ["Depreciation And Amortization",
                               "Depreciation Amortization Depletion"], y, note=mix) \
          or pick(d, "income", ["Reconciled Depreciation"], y, note=mix)
    bv  = pick(d, "balance", ["Stockholders Equity"], y, note=mix)
    cash = pick(d, "balance", ["Cash And Cash Equivalents",
                               "Cash Cash Equivalents And Short Term Investments"], y, note=mix)
    debt = pick(d, "balance", ["Total Debt"], y, note=mix)
    mc  = (d.get("market_cap") or 0) / 1e6
    ebitda = (oi + da) if (oi is not None and da is not None) else None
    # A company with cash but no Total Debt line is debt-free (yfinance omits the
    # row rather than reporting zero). Treat it as zero and surface it in the log.
    if debt is None and cash is not None:
        debt, dbg = 0.0, " [debt-free: Total Debt line absent -> 0]"
    else:
        dbg = ""
    nd = (debt - cash) if (debt is not None and cash is not None) else None
    return {
        "Ticker": f"{c}.T", "Name": d["name"],
        "Revenue": None if rev is None else round(rev),
        "EBITDA": None if ebitda is None else round(ebitda),
        "Operating_Income": None if oi is None else round(oi),
        "Net_Income": None if ni is None else round(ni),
        "Book_Value": None if bv is None else round(bv),
        "Net_Debt": None if nd is None else round(nd),
        "Market_Cap": round(mc),
        "_fy": y, "_da": None if da is None else round(da), "_dbg": dbg + (f" [MIXED-YEAR: {', '.join(mix)}]" if mix else ""),
    }

if __name__ == "__main__":
    subject, peers = sys.argv[1], sys.argv[2:]
    rows = [row(subject)] + [row(p) for p in peers]
    cols = ["Ticker", "Name", "Revenue", "EBITDA", "Operating_Income",
            "Net_Income", "Book_Value", "Net_Debt", "Market_Cap"]
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, f"{subject}_comps.csv")
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({k: ("" if r[k] is None else r[k]) for k in cols})
    print(f"wrote {path}")
    for r in rows:
        miss = [k for k in cols if r[k] is None]
        print(f"  {r['Ticker']:<8}{r['Name'][:30]:<32} FY={r['_fy']} D&A={r['_da']} "
              f"{'MISSING:' + ','.join(miss) if miss else 'ok'}{r['_dbg']}")
