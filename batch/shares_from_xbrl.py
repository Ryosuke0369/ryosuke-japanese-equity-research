"""Extract 発行済株式総数 / 自己株式数 from a cached EDINET XBRL (generic).

Usage: python batch/shares_from_xbrl.py <docID> [...]
"""
import sys, os, glob, re
sys.stdout.reconfigure(encoding="utf-8")
from bs4 import BeautifulSoup
ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "tmp", "edinet_data")
TAGS = {
 "issued_total": ["TotalNumberOfIssuedSharesSummaryOfBusinessResults",
                  "NumberOfIssuedSharesAsOfFilingDateIssuedSharesTotalNumberOfSharesEtc",
                  "TotalNumberOfIssuedShares"],
 "treasury": ["NumberOfTreasuryStock", "TreasurySharesAtEndOfPeriod",
              "NumberOfTreasuryStockAtTheEndOfCurrentPeriod"],
}
def scan(doc):
    for p in glob.glob(os.path.join(ROOT, doc, "XBRL", "PublicDoc", "*.xbrl")):
        soup = BeautifulSoup(open(p, encoding="utf-8").read(), "xml")
        print(f"  file {os.path.basename(p)}")
        hits = {}
        for el in soup.find_all():
            n = el.name.split(":")[-1]
            if re.search(r"(IssuedShares|TreasuryStock|TreasuryShares|NumberOfShares)", n):
                ctx = el.get("contextRef", "")
                v = (el.text or "").strip()
                if v and re.fullmatch(r"-?\d+(\.\d+)?", v):
                    hits.setdefault(n, []).append((ctx, v))
        for n, vs in sorted(hits.items()):
            for ctx, v in vs[:3]:
                print(f"    {n:<62} {ctx:<34} {float(v):,.0f}")
if __name__ == "__main__":
    for d in sys.argv[1:]:
        print(f"=== {d}"); scan(d)
