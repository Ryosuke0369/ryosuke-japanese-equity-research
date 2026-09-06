"""Summarize a cached ticker into the inputs the overrides design needs.

Generic: ticker codes come from argv. Prints FY series, margin distribution,
NWC days, capex/D&A ratios, leverage and the size-premium band.
"""
import sys, os, json, statistics as st
sys.stdout.reconfigure(encoding="utf-8")
CACHE = os.path.join(os.path.dirname(__file__), "cache")

def g(d, key):
    return d.get(key, {})

def yrs(d):
    ks = set()
    for blk in ("income", "balance", "cashflow"):
        for v in d[blk].values():
            ks |= set(v.keys())
    return sorted(ks)

def pick(d, blk, names, y):
    for n in names:
        v = d[blk].get(n, {}).get(y)
        if v is not None:
            return v
    return None

def band(mcap_oku):
    if mcap_oku < 100: return 0.050
    if mcap_oku < 300: return 0.040
    if mcap_oku < 1000: return 0.030
    if mcap_oku < 3000: return 0.020
    return 0.010

def main(code):
    d = json.load(open(os.path.join(CACHE, f"{code}.json"), encoding="utf-8"))
    ys = yrs(d)
    print(f"=== {code} {d['name']} | {d.get('sector')} / {d.get('industry')}")
    mc = (d.get("market_cap") or 0) / 1e8
    print(f"price={d['price']} ({d.get('price_date')}) shares={d.get('shares_outstanding'):,} "
          f"mcap={mc:,.0f}oku beta={d.get('beta')} -> size_premium band {band(mc):.1%}")
    print(f"{'FY':<12}{'Rev':>12}{'OP':>11}{'OPM':>8}{'NI':>11}{'COGS%':>8}{'SGA%':>8}"
          f"{'D&A':>10}{'D&A%':>7}{'Capex':>10}{'Cpx%':>7}{'OCF':>11}")
    rows = {}
    for y in ys:
        rev = pick(d,"income",["Total Revenue","Operating Revenue"],y)
        op  = pick(d,"income",["Operating Income","Total Operating Income As Reported"],y)
        ni  = pick(d,"income",["Net Income Common Stockholders","Net Income"],y)
        cogs= pick(d,"income",["Cost Of Revenue"],y)
        sga = pick(d,"income",["Selling General And Administration"],y)
        da  = pick(d,"cashflow",["Depreciation And Amortization","Depreciation Amortization Depletion"],y) \
              or pick(d,"income",["Reconciled Depreciation"],y)
        cpx = pick(d,"cashflow",["Capital Expenditure"],y)
        ocf = pick(d,"cashflow",["Operating Cash Flow"],y)
        pre = pick(d,"income",["Pretax Income"],y)
        rows[y] = dict(rev=rev,op=op,ni=ni,cogs=cogs,sga=sga,da=da,cpx=cpx,ocf=ocf,pre=pre)
        f = lambda v: f"{v:,.0f}" if isinstance(v,(int,float)) else "-"
        p = lambda a,b: f"{a/b:.1%}" if (a is not None and b) else "-"
        print(f"{y:<12}{f(rev):>12}{f(op):>11}{p(op,rev):>8}{f(ni):>11}{p(cogs,rev):>8}"
              f"{p(sga,rev):>8}{f(da):>10}{p(da,rev):>7}{f(abs(cpx) if cpx else None):>10}"
              f"{p(abs(cpx) if cpx else None,rev):>7}{f(ocf):>11}")
    # distributions
    opms = [r['op']/r['rev'] for r in rows.values() if r['op'] and r['rev']]
    if opms:
        opms_s = sorted(opms)
        q = lambda p: opms_s[0] if len(opms_s)==1 else st.quantiles(opms_s,n=4,method='inclusive')[p]
        print(f"OPM  min={min(opms):.2%} p25={q(0):.2%} med={st.median(opms):.2%} "
              f"p75={q(2):.2%} max={max(opms):.2%}")
    revs = [rows[y]['rev'] for y in ys if rows[y]['rev']]
    if len(revs) >= 2:
        n = len(revs)-1
        print(f"Revenue CAGR {n}y: {((revs[-1]/revs[0])**(1/n)-1):.2%}  "
              f"YoY latest: {(revs[-1]/revs[-2]-1):.2%}")
    # pretax vs OP (equity-method / non-operating dominance proxy)
    ly = ys[-1]
    r = rows[ly]
    if r['pre'] and r['op']:
        print(f"NonOp/OP proxy (Pretax-OP)/OP @{ly}: {(r['pre']-r['op'])/abs(r['op']):+.1%}")
    # balance sheet latest
    cash = pick(d,"balance",["Cash And Cash Equivalents","Cash Cash Equivalents And Short Term Investments"],ly)
    debt = pick(d,"balance",["Total Debt"],ly)
    mi   = pick(d,"balance",["Minority Interest"],ly)
    eq   = pick(d,"balance",["Stockholders Equity"],ly)
    ar   = pick(d,"balance",["Accounts Receivable","Receivables"],ly)
    inv  = pick(d,"balance",["Inventory"],ly)
    ap   = pick(d,"balance",["Accounts Payable","Payables"],ly)
    ta   = pick(d,"balance",["Total Assets"],ly)
    print(f"BS@{ly}: cash={cash:,.0f} debt={debt:,.0f} MI={mi if mi is None else format(mi,',.0f')} "
          f"equity={eq:,.0f} TA={ta:,.0f}" if cash and debt and eq and ta else f"BS@{ly}: partial")
    if cash is not None and debt is not None:
        print(f"  net_debt(type A) = {debt:,.0f} - {cash:,.0f} = {debt-cash:,.0f}")
    if debt and rows[ly]['rev']:
        print(f"  Debt/Revenue = {debt/rows[ly]['rev']:.2f}")
    rev = rows[ly]['rev']
    if rev:
        cogs = rows[ly]['cogs']
        if ar: print(f"  DSO = {ar/rev*365:.0f}d", end="")
        if inv and cogs: print(f"  DIH = {inv/cogs*365:.0f}d", end="")
        if ap and cogs: print(f"  DPO = {ap/cogs*365:.0f}d", end="")
        print()
    # hist arrays for overrides
    print("hist_years    =", [f"FY{y[:4]}" for y in ys])
    for k,blk,names in [("hist_revenue","income",["Total Revenue","Operating Revenue"]),
                        ("hist_operating_income","income",["Operating Income"]),
                        ("hist_net_income","income",["Net Income Common Stockholders","Net Income"]),
                        ("hist_cogs","income",["Cost Of Revenue"]),
                        ("hist_sga","income",["Selling General And Administration"]),
                        ("hist_ocf","cashflow",["Operating Cash Flow"]),
                        ("hist_capex","cashflow",["Capital Expenditure"]),
                        ("hist_depreciation","cashflow",["Depreciation And Amortization","Depreciation Amortization Depletion"]),
                        ("hist_cash","balance",["Cash And Cash Equivalents","Cash Cash Equivalents And Short Term Investments"]),
                        ("hist_debt","balance",["Total Debt"])]:
        vals=[pick(d,blk,names,y) for y in ys]
        if k=="hist_capex": vals=[abs(v) if v is not None else None for v in vals]
        print(f"{k:<22}=", [None if v is None else round(v,1) for v in vals])
    ie = pick(d,"income",["Interest Expense","Interest Expense Non Operating"],ly)
    print("interest_expense@%s = %s" % (ly, ie))

if __name__ == "__main__":
    for c in sys.argv[1:]:
        main(c); print()
