"""Catch silent yfinance market-data fallbacks (price 1,000 / shares 10,000,000).

When get_live_market_data() throws, generate_dcf.py falls back to the template's
placeholder price and share count WITHOUT failing validation. 4568 shipped a
Target of 294,427 yen that way (market cap 10,000 mn instead of 5,084,100 mn).
This check compares the workbook against the cached market data.

Usage: python batch/check_market_data.py [<xlsx> ...]   (default: all 20260905 models)
"""
import sys, os, json, glob, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from stale_check import assert_fresh, assert_recalculated  # 追補6 §Z / 追補10 §AO

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLACEHOLDER = {"price": 1000.0, "shares": 10000000}


def check(x):
    code = os.path.basename(x)[:4]
    ws = openpyxl.load_workbook(x, data_only=True)["DCF Model"]
    price = shares = None
    for r in range(1, 40):
        lab = ws.cell(r, 2).value
        if isinstance(lab, str):
            if lab.startswith("Fully Diluted Shares"):
                shares = ws.cell(r, 3).value
    es = openpyxl.load_workbook(x, data_only=True)["Executive Summary"]
    for r in range(1, 20):
        if es.cell(r, 2).value == "Current Price":
            price = es.cell(r, 3).value
    cache = os.path.join(ROOT, "batch", "cache", "%s.json" % code)
    ref_p = ref_s = None
    if os.path.isfile(cache):
        d = json.load(open(cache, encoding="utf-8"))
        ref_p, ref_s = d.get("price"), d.get("shares_outstanding")
    ph = (price == PLACEHOLDER["price"] and shares == PLACEHOLDER["shares"])
    off = (ref_p and price and abs(price - ref_p) / ref_p > 0.02) or \
          (ref_s and shares and abs(shares - ref_s) / ref_s > 0.02)
    status = "*** PLACEHOLDER FALLBACK ***" if ph else ("*** MISMATCH ***" if off else "OK ")
    print(f"{status} {code}: price={price!r} (cache {ref_p!r})  shares={shares!r} (cache {ref_s!r})")
    return not (ph or off)


if __name__ == "__main__":
    paths = sys.argv[1:] or sorted(glob.glob(os.path.join(ROOT, "models", "*_DCF_Model_2026090[56].xlsx")))
    stale = [p for p in paths if not assert_fresh(p)]      # 追補6 §Z
    stale += [p for p in paths if not assert_recalculated(p)]   # 追補10 §AO
    bad = list(dict.fromkeys([p for p in paths if not check(p)] + stale))
    print("\n%d/%d clean" % (len(paths) - len(bad), len(paths)))
