"""Generic yfinance fundamentals cache for the batch DCF run.

No ticker codes in this file: pass codes on the command line.
Writes batch/cache/<code>.json (JPY mn for financial statement lines).
"""
import sys, os, json, time, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
import yfinance as yf

CACHE = os.path.join(os.path.dirname(__file__), "cache")
os.makedirs(CACHE, exist_ok=True)
MN = 1_000_000.0

IS_KEYS = ["Total Revenue", "Operating Revenue", "Cost Of Revenue",
           "Selling General And Administration", "Operating Income",
           "Total Operating Income As Reported", "EBITDA", "EBIT",
           "Reconciled Depreciation", "Net Income Common Stockholders",
           "Net Income", "Pretax Income", "Tax Provision",
           "Interest Expense", "Interest Expense Non Operating"]
BS_KEYS = ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments",
           "Other Short Term Investments", "Total Debt", "Current Debt", "Long Term Debt",
           "Stockholders Equity", "Minority Interest", "Total Assets",
           "Total Liabilities Net Minority Interest", "Accounts Receivable",
           "Receivables", "Inventory", "Accounts Payable", "Payables",
           "Working Capital", "Net PPE"]
CF_KEYS = ["Operating Cash Flow", "Capital Expenditure",
           "Depreciation And Amortization", "Depreciation Amortization Depletion"]


def series(df, keys):
    out = {}
    if df is None or df.empty:
        return out
    cols = [str(c)[:10] for c in df.columns]
    for k in keys:
        if k in df.index:
            vals = []
            for c in df.columns:
                v = df.loc[k, c]
                try:
                    vals.append(None if v is None or v != v else round(float(v) / MN, 1))
                except Exception:
                    vals.append(None)
            out[k] = dict(zip(cols, vals))
    return out


def grab(code):
    t = yf.Ticker(f"{code}.T")
    info = {}
    try:
        info = t.info or {}
    except Exception as e:
        print(f"  info failed: {e}")
    fi = {}
    try:
        f = t.fast_info
        fi = {"lastPrice": f.get("lastPrice"), "marketCap": f.get("marketCap"),
              "shares": f.get("shares")}
    except Exception:
        pass
    px_date = None
    try:
        h = t.history(period="10d")
        if len(h):
            px_date = str(h.index[-1])[:10]
            fi["lastPrice"] = float(h["Close"].iloc[-1])
    except Exception:
        pass
    d = {
        "code": code,
        "name": info.get("longName") or info.get("shortName"),
        "sector": info.get("sector"), "industry": info.get("industry"),
        "price": fi.get("lastPrice") or info.get("currentPrice"),
        "price_date": px_date,
        "shares_outstanding": info.get("sharesOutstanding"),
        "implied_shares": info.get("impliedSharesOutstanding"),
        "market_cap": info.get("marketCap") or fi.get("marketCap"),
        "beta": info.get("beta"),
        "trailingPE": info.get("trailingPE"),
        "income": series(t.financials, IS_KEYS),
        "balance": series(t.balance_sheet, BS_KEYS),
        "cashflow": series(t.cashflow, CF_KEYS),
        "q_income": series(t.quarterly_financials, ["Total Revenue", "Operating Income"]),
    }
    return d


if __name__ == "__main__":
    codes = sys.argv[1:]
    for i, code in enumerate(codes):
        p = os.path.join(CACHE, f"{code}.json")
        if os.path.isfile(p):
            print(f"[{i+1}/{len(codes)}] {code} cached")
            continue
        try:
            d = grab(code)
            with open(p, "w", encoding="utf-8") as fh:
                json.dump(d, fh, ensure_ascii=False, indent=1)
            print(f"[{i+1}/{len(codes)}] {code} {d['name']} px={d['price']} "
                  f"mcap={(d['market_cap'] or 0)/1e8:,.0f}oku beta={d['beta']}")
        except Exception as e:
            print(f"[{i+1}/{len(codes)}] {code} FAILED {type(e).__name__}: {e}")
        time.sleep(1.2)
