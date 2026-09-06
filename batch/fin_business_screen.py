"""Screen a cached EDINET XBRL for consolidated financial-business line items.

Catches the 9433 KDDI case: a group that consolidates a bank / credit business
shows dedicated balance-sheet elements the ordinary type-A net_debt definition
must not absorb. Run after the pipeline has downloaded a ticker's XBRL.

Usage: python batch/fin_business_screen.py <docID> [...]
       python batch/fin_business_screen.py --latest    (scan every cached doc)
"""
import sys, os, glob, re
sys.stdout.reconfigure(encoding="utf-8")
from bs4 import BeautifulSoup

ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "tmp", "edinet_data")
PAT = re.compile(r"(ForFinancialBusiness|BankingBusiness|CallLoan|CallMoney|"
                 r"DepositsFromCustomers|LoansAndBillsDiscounted|InsuranceContract|"
                 r"PolicyReserve|InstallmentReceivable|LeaseReceivable)")

def scan(doc, return_hits=False):
    files = glob.glob(os.path.join(ROOT, doc, "XBRL", "PublicDoc", "*.xbrl"))
    if not files:
        print(f"{doc}: no PublicDoc xbrl cached")
        return {} if return_hits else None
    soup = BeautifulSoup(open(files[0], encoding="utf-8").read(), "xml")
    hits, total = {}, None
    for el in soup.find_all():
        n = el.name.split(":")[-1]
        ctx = el.get("contextRef", "")
        if "CurrentYearInstant" not in ctx or "Member" in ctx:
            continue
        try:
            v = float(el.text)
        except (TypeError, ValueError):
            continue
        if n in ("AssetsIFRS", "Assets"):
            total = v
        if PAT.search(n):
            hits[n] = v
    if not hits:
        print(f"{doc}: CLEAN - no financial-business balance-sheet elements")
        return {} if return_hits else None
    print(f"{doc}: *** FINANCIAL-BUSINESS ELEMENTS FOUND ***"
          f"{'' if not total else f'  (total assets {total/1e6:,.0f} mn)'}")
    for n, v in sorted(hits.items(), key=lambda kv: -kv[1]):
        share = f"  = {v/total:.1%} of assets" if total else ""
        print(f"   {n:<58}{v/1e6:>15,.0f} mn{share}")
    if return_hits:
        return {n: round(v / 1e6) for n, v in hits.items()}

if __name__ == "__main__":
    args = sys.argv[1:]
    docs = ([os.path.basename(p) for p in glob.glob(os.path.join(ROOT, "S*"))]
            if args == ["--latest"] else args)
    for d in docs:
        scan(d)
