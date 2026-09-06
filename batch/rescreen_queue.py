"""フェーズ2 §6 — re-screen the queued tickers whose blocker was a data-acquisition one.

A queued ticker is not re-modelled here. Modelling one needs overrides and a
comps CSV that do not exist for any of these twenty, and inventing them would be
exactly the guessing the batch rules forbid. What this does is measure whether
the BLOCKER is still there, using the repaired acquisition path (#7 archive
docIDs, #11 FY-end inference, #1 missing-value handling, #9 guidance):

  * which fiscal years EDINET now returns, and how recent the newest one is
  * whether depreciation / operating income / net debt are present for the
    latest year — the three absences that sent tickers to the queue
  * whether 会社予想 is now obtainable

Tickers queued for a STRUCTURAL reason (型C/D/E, a negative mid-cycle OPM, a peer
set that does not exist) are not touched: no pipeline fix changes what kind of
company they are. They are listed as-is so the report can carry them forward.

Usage:
    python batch/rescreen_queue.py                  # the data-blocker candidates
    python batch/rescreen_queue.py --all            # every queued ticker
"""
import argparse
import json
import os
import sys
import warnings

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from scripts.edinet_fetcher import (   # noqa: E402
    fetch_and_parse_multi_year, docs_from_screener_archive,
    infer_fiscal_year_end_month,
)
from scripts.guidance_fetcher import get_guidance   # noqa: E402

# Blockers that a pipeline fix could plausibly have removed.
DATA_BLOCKED = {
    "4519": "D&A が EDINET有報XBRL・yfinance のいずれからも取得できず core_ebitda ゲート不通過",
    "6594": "基準年が17か月古い（EDINET最新有報が FY2024/3 止まりと報告されていた）",
    "7741": "yfinance の損益計算書が構造的に誤マッピング（COGS 誤取得）",
    "7752": "D&A 取得不能",
}
# Blockers that describe what the company IS, not what the pipeline could fetch.
STRUCTURAL = {
    "2914": "Comps 不成立（有意な比較対象が海外たばこのみ。§5 で海外Peerはバッチ対象外）",
    "3401": "3期連続営業赤字。観測窓にミッドサイクルが存在しない（型B）",
    "4689": "型D 相当（金融事業を連結）",
    "6301": "型C（Komatsu Financial）",
    "6326": "型C（Kubota Credit）",
    "6971": "判定困難（KDDI 株主導）",
    "7731": "減損除外後も直近赤字（型B）",
    "8001": "型E（オリコ連結）",
    "8002": "型C/E 相当（持分法主導）",
    "8031": "型C/E 相当（持分法主導）",
    "8053": "型C/E 相当（持分法主導＋金融）",
    "8058": "型C/E 相当（持分法主導）",
    "9101": "判定困難（ONE 持分法主導）",
    "9202": "Peer セット不成立",
    "9433": "型E（auじぶん銀行・au損保を連結）",
    "9434": "型E（PayPay銀行連結）",
}

FIELDS = ("revenue", "operating_income", "depreciation", "net_debt",
          "total_debt", "cash", "cogs", "sga", "net_income")


def probe(code):
    out = {"code": code}
    fy_m = infer_fiscal_year_end_month(code)
    out["fy_end_month"] = fy_m
    arch = docs_from_screener_archive(code, 5, fy_m)
    out["archive_docs"] = [(d["period_end"], d["doc_id"]) for d in arch]
    try:
        info, merged = fetch_and_parse_multi_year(code, 5, fiscal_year_end_month=fy_m)
    except Exception as e:
        out["edinet_error"] = f"{type(e).__name__}: {e}"
        return out
    out["company"] = info.get("company_name")
    fys = sorted(k for k in merged if str(k).startswith("FY"))
    out["fiscal_years"] = fys
    out["latest_fy"] = fys[-1] if fys else None
    if fys:
        latest = merged[fys[-1]]
        out["latest_fields"] = {f: latest.get(f) for f in FIELDS}
        out["missing_latest"] = [f for f in FIELDS if latest.get(f) is None]
    fy_year = max((int(k[2:6]) for k in fys if k[2:6].isdigit()), default=None)
    fd, note, src = get_guidance(code, min_fy_year=fy_year,
                                xbrl_paths=(merged.get("_meta") or {}).get("xbrl_paths"))
    out["guidance_source"] = src
    out["guidance_note"] = note
    out["guidance"] = fd
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--only", nargs="*")
    ap.add_argument("--out", default=os.path.join(HERE, "phase2_rescreen.json"))
    a = ap.parse_args()

    codes = a.only or (sorted(set(DATA_BLOCKED) | set(STRUCTURAL)) if a.all
                       else sorted(DATA_BLOCKED))
    res = {}
    for code in codes:
        print(f"\n=== {code}  ({DATA_BLOCKED.get(code) or STRUCTURAL.get(code, '?')})")
        r = probe(code)
        res[code] = r
        if r.get("edinet_error"):
            print(f"  EDINET: {r['edinet_error']}")
            continue
        print(f"  {r.get('company')}  FY-end month {r.get('fy_end_month')}")
        print(f"  archive docIDs: "
              + ", ".join(f"{p}={d}" for p, d in r["archive_docs"][:5]))
        print(f"  fiscal years  : {', '.join(r.get('fiscal_years') or [])}")
        lf = r.get("latest_fields") or {}
        print(f"  latest FY {r.get('latest_fy')}: "
              + ", ".join(f"{k}={'-' if v is None else format(v, ',.0f')}"
                          for k, v in lf.items()))
        if r.get("missing_latest"):
            print(f"  STILL MISSING: {', '.join(r['missing_latest'])}")
        else:
            print("  all probed fields present")
        print(f"  guidance: [{r['guidance_source']}] {r['guidance_note'][:120]}")

    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)
    print(f"\ndetail: {a.out}")
    print("\nstructural blockers (not re-attempted - no pipeline fix applies):")
    for c, why in sorted(STRUCTURAL.items()):
        print(f"  {c}  {why}")


if __name__ == "__main__":
    main()
