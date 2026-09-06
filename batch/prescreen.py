"""Pre-screen a ticker for consolidated financial-business line items (追補1 §B).

Downloads ONLY the most recent annual report XBRL (num_years=1) and runs the
same element-name check as fin_business_screen.py. Cheap enough to run before
the type triage, which is where it has to sit: 9433 KDDI looks like an ordinary
telecom on every financial ratio and is only identifiable from the XBRL.

Usage: python batch/prescreen.py <ticker> [<ticker> ...]
Exit output per ticker: "CLEAN" or "FINANCIAL-BUSINESS" plus the elements found.
"""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from scripts.edinet_fetcher import get_document_ids, download_and_extract_xbrl
from batch.fin_business_screen import scan

STATE = os.path.join(HERE, "prescreen_results.json")


def load():
    return json.load(open(STATE, encoding="utf-8")) if os.path.isfile(STATE) else {}


def run(code, fy_month=None):
    res = load()
    if code in res:
        print(f"{code}: cached -> {res[code]['verdict']} ({res[code]['doc_id']})")
        return res[code]
    docs = get_document_ids(code, num_years=1, fiscal_year_end_month=fy_month)
    if not docs:
        print(f"{code}: NO ANNUAL REPORT FOUND")
        res[code] = {"verdict": "NO_DOC", "doc_id": None}
    else:
        d = docs[0]
        doc_id = d["doc_id"] if isinstance(d, dict) else d
        download_and_extract_xbrl(doc_id)
        print(f"--- {code} (docID {doc_id})")
        hits = scan(doc_id, return_hits=True)
        res[code] = {"verdict": "FINANCIAL-BUSINESS" if hits else "CLEAN",
                     "doc_id": doc_id, "hits": hits or {}}
    json.dump(res, open(STATE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return res[code]


if __name__ == "__main__":
    import csv
    fy = {r["ticker"]: int(r["fiscal_year_end_month"])
          for r in csv.DictReader(open(os.path.join(HERE, "tickers.csv"), encoding="utf-8"))}
    for c in sys.argv[1:]:
        try:
            run(c, fy.get(c))
        except Exception as e:
            print(f"{c}: PRESCREEN FAILED {type(e).__name__}: {e}")
