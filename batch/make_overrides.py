"""Emit the data-driven half of an overrides JSON from the yfinance cache.

Generic (ticker codes come from argv). Writes batch/draft/<code>_data.json holding
ONLY the factual blocks: hist_* arrays, base-year values, WACC inputs derived by
the batch rules, and the NWC day averages. Scenario design, thesis/risks and the
exit multiple are added by hand per ticker - this file never invents them.
"""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
CACHE, DRAFT = os.path.join(HERE, "cache"), os.path.join(HERE, "draft")
os.makedirs(DRAFT, exist_ok=True)

RF, ERP, TAX, TG = 0.0297, 0.065, 0.25, 0.010
MARKET, LOOKBACK, INTERVAL, MIN_OBS = "1306.T", "2y", "1wk", 60


def topix_beta(code):
    """Raw beta from a 2y weekly OLS regression on TOPIX (1306.T), or None.

    The same measurement batch/rederive_beta.py applied to the 85, so a ticker
    added later sits on the same basis as the rest of the batch.
    """
    try:
        import warnings
        warnings.filterwarnings("ignore")
        import numpy as np, pandas as pd, yfinance as yf
        px = yf.download([f"{code}.T", MARKET], period=LOOKBACK, interval=INTERVAL,
                         auto_adjust=True, progress=False)["Close"]
        r = px.pct_change()
        d = pd.concat([r[f"{code}.T"], r[MARKET]], axis=1).dropna()
        if len(d) < MIN_OBS:
            return None
        y, x = d.iloc[:, 0].values, d.iloc[:, 1].values
        v = float(np.var(x, ddof=1))
        return round(float(np.cov(y, x, ddof=1)[0, 1] / v), 4) if v > 0 else None
    except Exception:
        return None

def band(mcap_mn):
    oku = mcap_mn / 100.0
    return 0.050 if oku < 100 else 0.040 if oku < 300 else 0.030 if oku < 1000 \
        else 0.020 if oku < 3000 else 0.010

def build(code, n=5):
    d = json.load(open(os.path.join(CACHE, f"{code}.json"), encoding="utf-8"))
    ys = sorted({k for blk in ("income","balance","cashflow") for v in d[blk].values() for k in v})
    ys = ys[-n:]
    def s(blk, names):
        out = []
        for y in ys:
            v = None
            for nm in names:
                if d[blk].get(nm, {}).get(y) is not None:
                    v = d[blk][nm][y]; break
            out.append(None if v is None else round(v, 1))
        return out
    rev  = s("income", ["Total Revenue", "Operating Revenue"])
    oi   = s("income", ["Operating Income", "Total Operating Income As Reported"])
    ni   = s("income", ["Net Income Common Stockholders", "Net Income"])
    cogs = s("income", ["Cost Of Revenue"])
    sga  = s("income", ["Selling General And Administration"])
    ocf  = s("cashflow", ["Operating Cash Flow"])
    cpx  = [None if v is None else abs(v) for v in s("cashflow", ["Capital Expenditure"])]
    da   = s("cashflow", ["Depreciation And Amortization", "Depreciation Amortization Depletion"])
    da   = [x if x is not None else y for x, y in zip(da, s("income", ["Reconciled Depreciation"]))]
    cash = s("balance", ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"])
    debt = s("balance", ["Total Debt"])
    ar   = s("balance", ["Accounts Receivable", "Receivables"])
    inv  = s("balance", ["Inventory"])
    ap   = s("balance", ["Accounts Payable", "Payables"])
    ie   = s("income", ["Interest Expense", "Interest Expense Non Operating"])

    mcap = (d.get("market_cap") or 0) / 1e6
    # フェーズ2 #6 / §2-4: `beta` in overrides is the RAW (regression) beta - the
    # template applies the Blume shrink and the [0.3, 2.0] clamp itself. This
    # used to emit max(0.6, min(1.75, raw)), i.e. the OLD clamped value, which
    # would have quietly put a new ticker back on the floor that 57 of the 85
    # were stuck at. The raw beta is measured against TOPIX rather than taken
    # from yfinance's `beta` field, which returns -0.165 for NTT and -0.201 for
    # 大阪ガス and is not a measurement of Japanese equity risk.
    raw_beta = topix_beta(code)
    beta_src = f"TOPIX {LOOKBACK} {INTERVAL} OLS"
    if raw_beta is None:
        raw_beta = d.get("beta")
        beta_src = "yfinance beta field (TOPIX regression unavailable)"
    beta = raw_beta
    # NWC day averages over the years where the inputs exist
    def avg(vals):
        vals = [v for v in vals if v]
        return round(sum(vals)/len(vals)) if vals else None
    dso = avg([a/r*365 for a, r in zip(ar, rev) if a and r])
    dih = avg([i/c*365 for i, c in zip(inv, cogs) if i and c])
    dpo = avg([p/c*365 for p, c in zip(ap, cogs) if p and c])

    out = {
      "_data_note": f"yfinance連結財務諸表({ys[0][:7]}〜{ys[-1][:7]})から機械生成。EDINET欠損年の補完用。",
      "_raw_beta": raw_beta, "_market_cap_mn": round(mcap), "_price": d.get("price"),
      "_price_date": d.get("price_date"), "_shares": d.get("shares_outstanding"),
      "_opm": [None if (o is None or not r) else round(o/r, 4) for o, r in zip(oi, rev)],
      "_capex_pct_hist": [None if (c is None or not r) else round(c/r, 4) for c, r in zip(cpx, rev)],
      "_da_pct_hist": [None if (x is None or not r) else round(x/r, 4) for x, r in zip(da, rev)],
      "_nwc_days_avg": {"dso": dso, "dih": dih, "dpo": dpo},
      "_interest_expense_latest": ie[-1],
      "_net_debt_calc": None if (debt[-1] is None or cash[-1] is None) else round(debt[-1]-cash[-1]),
      "risk_free": RF, "erp": ERP, "tax_rate": TAX, "terminal_growth": TG,
      "size_premium": band(mcap), "beta": beta,
      "hist_years": [f"FY{y[:4]}" for y in ys],
      "hist_revenue": rev, "hist_operating_income": oi, "hist_net_income": ni,
      "hist_cogs": cogs, "hist_sga": sga, "hist_ocf": ocf, "hist_capex": cpx,
      "hist_depreciation": da, "hist_cash": cash, "hist_debt": debt,
      "base_year_revenue": rev[-1], "base_year_cogs": cogs[-1],
      "base_year_ar": ar[-1], "base_year_inv": inv[-1], "base_year_ap": ap[-1],
      "nwc_method": "days",
    }
    p = os.path.join(DRAFT, f"{code}_data.json")
    json.dump(out, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"{code}: beta raw={raw_beta} ({beta_src})  size_prem={out['size_premium']:.1%}  "
          f"DSO/DIH/DPO={dso}/{dih}/{dpo}  net_debt={out['_net_debt_calc']}  -> {p}")

if __name__ == "__main__":
    for c in sys.argv[1:]:
        build(c)
