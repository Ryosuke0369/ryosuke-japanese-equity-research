"""Emit the mandatory preemptive EDINET-completion block (追補2 §F-2).

Root cause (generate_dcf.py:66) — `_val(data_dict, key, default=0)` turns a
missing EDINET value into 0 rather than None — makes every hist_* series unsafe.
The confirmed workaround is to replace hist_years with the last 4 complete fiscal
years and supply the five P/L series ourselves, while leaving the CF/BS series to
Step 4.6, which re-maps them onto the final hist_years by fiscal-year key.

追補2 §F-2 makes this mandatory for every remaining ticker rather than a reaction
to a failure. This file is generic; ticker codes come from argv.

Also emits:
  * `core_ebitda` (= latest OI + latest cash-flow D&A) so the Comps subject block
    matches the peer definition — the silent failure caught by
    batch/check_core_ebitda.py (追補2 §G);
  * `current_price` / `shares_outstanding`, because a failed live yfinance call at
    generation time falls back to the template placeholders (price 1,000 /
    shares 10,000,000) with no validation failure — 4568 shipped a 294,427 yen
    Target that way. Supplying them also makes a run reproducible.

Usage: python batch/preemptive_block.py <code> [<code> ...]
Writes batch/draft/<code>_preemptive.json
"""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CACHE, DRAFT = os.path.join(HERE, "cache"), os.path.join(HERE, "draft")
os.makedirs(DRAFT, exist_ok=True)

IS = {"rev": ["Total Revenue", "Operating Revenue"],
      "oi": ["Operating Income", "Total Operating Income As Reported"],
      "ni": ["Net Income Common Stockholders", "Net Income"],
      "cogs": ["Cost Of Revenue"]}


def declared_type(code):
    """overrides の company_type を読む(無ければ None)。

    ここで型を知りたい理由は一つだけ: 型D/E の net_debt を出さないため。
    overrides がまだ無い新規銘柄では None が返り、従来どおりの動作になる。
    """
    p = os.path.join(ROOT, "data", "overrides", "%s_overrides.json" % code)
    if not os.path.exists(p):
        return None
    try:
        d = json.load(open(p, encoding="utf-8"))
    except (ValueError, OSError):
        return None
    t = str(d.get("company_type", "")).strip().upper()
    return t or None


def pick(d, blk, names, y):
    for n in names:
        v = d[blk].get(n, {}).get(y)
        if v is not None:
            return v
    return None


def build(code, n=4):
    d = json.load(open(os.path.join(CACHE, "%s.json" % code), encoding="utf-8"))
    years = sorted({k for blk in ("income", "balance", "cashflow")
                    for v in d[blk].values() for k in v})
    # keep only years where the four P/L lines are all present
    complete = [y for y in years if all(pick(d, "income", IS[k], y) is not None for k in IS)]
    ys = complete[-n:]
    if len(ys) < 3:
        print("%s: *** only %d complete P/L years — needs manual handling ***" % (code, len(ys)))
        return None
    rev = [pick(d, "income", IS["rev"], y) for y in ys]
    oi = [pick(d, "income", IS["oi"], y) for y in ys]
    ni = [pick(d, "income", IS["ni"], y) for y in ys]
    cogs = [pick(d, "income", IS["cogs"], y) for y in ys]
    # SGA is back-solved so that OP = Revenue - COGS - SGA holds in the workbook
    sga = [round(r - c - o, 1) for r, c, o in zip(rev, cogs, oi)]
    ly = ys[-1]
    da = (pick(d, "cashflow", ["Depreciation And Amortization",
                               "Depreciation Amortization Depletion"], ly)
          or pick(d, "income", ["Reconciled Depreciation"], ly))
    out = {
        "_preemptive_years": [y[:7] for y in ys],
        "current_price": d.get("price"),
        "shares_outstanding": d.get("shares_outstanding"),
        "hist_years": ["FY%s" % y[:4] for y in ys],
        "hist_revenue": [round(v, 1) for v in rev],
        "hist_operating_income": [round(v, 1) for v in oi],
        "hist_net_income": [round(v, 1) for v in ni],
        "hist_cogs": [round(v, 1) for v in cogs],
        "hist_sga": sga,
        "base_year_revenue": round(rev[-1], 1),
        "base_year_cogs": round(cogs[-1], 1),
        "base_year_ar": pick(d, "balance", ["Accounts Receivable", "Receivables"], ly),
        "base_year_inv": pick(d, "balance", ["Inventory"], ly),
        "base_year_ap": pick(d, "balance", ["Accounts Payable", "Payables"], ly),
        "book_value": pick(d, "balance", ["Stockholders Equity"], ly),
        "core_ebitda": None if da is None else round(oi[-1] + da, 1),
        "_da_latest": da,
        "_net_debt_yf": None,
    }
    ctype = declared_type(code)
    if ctype in ("D", "E"):
        # 型D/E: 連結の Total Debt には銀行の資金調達が、Cash には預け金が入る。
        # 値を出さず、どこから取るかを書く(サイレントに壊れた数字を渡さない)。
        out["_net_debt_yf"] = None
        out["_net_debt_note"] = (
            "型%s のため yfinance ベースの net_debt は出力しない。連結の Total Debt は"
            "銀行の資金調達を、Cash は預け金を含み、貸出金は資産側に残るため、"
            "この差額を net_debt にすると預金を有利子負債として割り引くことになる。"
            "有報の連結BS から非金融ベース(銀行預金・貸出金・コールローン/コールマネーを"
            "除外し、非支配株主持分を加算)で作成し、overrides の net_debt に明示すること。"
            % ctype)
        print("%s: [型%s] net_debt は出力しない — %s"
              % (code, ctype, "一次資料から非金融ベースで作成すること"))
    else:
        debt = pick(d, "balance", ["Total Debt"], ly)
        cash = pick(d, "balance", ["Cash And Cash Equivalents",
                                   "Cash Cash Equivalents And Short Term Investments"], ly)
        if cash is not None:
            out["_net_debt_yf"] = round((debt or 0) - cash, 1)
            out["_debt_line_absent"] = debt is None
    out = {k: v for k, v in out.items() if v is not None or k.startswith("_")}
    p = os.path.join(DRAFT, "%s_preemptive.json" % code)
    json.dump(out, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("%s: %s  px=%s shares=%s  core_ebitda=%s  net_debt(yf)=%s%s"
          % (code, out["hist_years"], out.get("current_price"), out.get("shares_outstanding"),
             out.get("core_ebitda"), out.get("_net_debt_yf"),
             "  [debt line absent -> treated as 0]" if out.get("_debt_line_absent") else ""))
    return out


if __name__ == "__main__":
    for c in sys.argv[1:]:
        build(c)
