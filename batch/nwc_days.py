import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")
def pick(d, blk, names, y):
    for n in names:
        v = d[blk].get(n, {}).get(y)
        if v is not None: return v
    return None
for c in sys.argv[1:]:
    d = json.load(open(os.path.join(CACHE, f"{c}.json"), encoding="utf-8"))
    ys = sorted({k for blk in ("income","balance","cashflow") for v in d[blk].values() for k in v})
    print(f"--- {c} {d['name']}")
    acc = {"dso":[], "dih":[], "dpo":[]}
    for y in ys:
        rev = pick(d,"income",["Total Revenue","Operating Revenue"],y)
        cogs= pick(d,"income",["Cost Of Revenue"],y)
        ar  = pick(d,"balance",["Accounts Receivable","Receivables"],y)
        inv = pick(d,"balance",["Inventory"],y)
        ap  = pick(d,"balance",["Accounts Payable","Payables"],y)
        dso = ar/rev*365 if ar and rev else None
        dih = inv/cogs*365 if inv and cogs else None
        dpo = ap/cogs*365 if ap and cogs else None
        for k,v in (("dso",dso),("dih",dih),("dpo",dpo)):
            if v: acc[k].append(v)
        f=lambda v: f"{v:.0f}" if v else "-"
        print(f"  {y}  DSO={f(dso):>4}  DIH={f(dih):>4}  DPO={f(dpo):>4}")
    print("  AVG   DSO=%s  DIH=%s  DPO=%s" % tuple(
        f"{sum(acc[k])/len(acc[k]):.0f}" if acc[k] else "-" for k in ("dso","dih","dpo")))
