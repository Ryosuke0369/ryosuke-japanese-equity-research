"""
comps_fetcher.py - Fetch comparable company data from CSV + yfinance.

Reads financial data from a CSV file and supplements with live market cap
from yfinance. Returns a list of dicts compatible with config["comps"].
"""

import csv
import io
import logging

logger = logging.getLogger(__name__)

try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False


def _fetch_market_cap(ticker_str):
    """Fetch market cap for a single ticker via yfinance.

    Returns market cap in JPY millions, or None on failure.
    Fallback: currentPrice * sharesOutstanding.
    """
    if not YFINANCE_AVAILABLE:
        logger.warning("yfinance not installed. Cannot fetch market cap for %s.", ticker_str)
        return None

    try:
        tkr = yf.Ticker(ticker_str)
        info = tkr.info

        mkt_cap = info.get("marketCap")
        if mkt_cap and mkt_cap > 0:
            return mkt_cap / 1_000_000  # Convert to JPY millions

        # Fallback: currentPrice * sharesOutstanding
        price = info.get("currentPrice") or info.get("regularMarketPrice")
        shares = info.get("sharesOutstanding")
        if price and shares:
            return (price * shares) / 1_000_000

        logger.warning("Could not determine market cap for %s from yfinance data.", ticker_str)
        return None
    except Exception as e:
        logger.warning("Failed to fetch market cap for %s: %s", ticker_str, e)
        return None


def get_comps_data(csv_path):
    """Load comparable company data from CSV, enrich with yfinance market cap.

    Args:
        csv_path: Path to UTF-8 comma-delimited CSV with columns:
                  Ticker, Name, Revenue, EBITDA, Operating_Income,
                  Net_Income, Book_Value, Net_Debt

    Returns:
        List of dicts with keys: name, ticker, mkt_cap, ev, revenue,
        ebitda, op_income, net_income, pbr, roe
    """
    comps = []

    with open(csv_path, encoding="utf-8") as f:
        # Sanitize: strip trailing whitespace from each line before parsing.
        # Trailing tabs corrupt delimiter auto-detection and DictReader fields.
        clean_lines = [line.rstrip() for line in f]

    clean_content = "\n".join(clean_lines)
    with io.StringIO(clean_content) as f_clean:
        # Auto-detect delimiter (handles both comma and tab-separated files)
        sample = clean_lines[0] if clean_lines else ""
        delimiter = "\t" if "\t" in sample else ","
        reader = csv.DictReader(f_clean, delimiter=delimiter)
        for row in reader:
            ticker = row["Ticker"].strip()
            name = row["Name"].strip()

            # Normalize column names: strip whitespace from keys
            row = {k.strip(): v.strip() for k, v in row.items()}

            revenue = float(row["Revenue"])
            ebitda = float(row["EBITDA"])
            op_income = float(row.get("Operating_Income") or row.get("Operating Income", "0"))
            net_income = float(row.get("Net_Income") or row.get("Net Income", "0"))
            book_value = float(row.get("Book_Value") or row.get("Book Value", "0"))
            net_debt = float(row.get("Net_Debt") or row.get("Net Debt", "0"))

            # Fetch market cap from yfinance
            mkt_cap = _fetch_market_cap(ticker)

            # Derived values
            if mkt_cap is not None:
                ev = mkt_cap + net_debt
                pbr = mkt_cap / book_value if book_value != 0 else None
            else:
                ev = None
                pbr = None

            roe = net_income / book_value if book_value != 0 else None

            comps.append({
                "name": name,
                "ticker": ticker,
                "mkt_cap": mkt_cap,
                "ev": ev,
                "revenue": revenue,
                "ebitda": ebitda,
                "op_income": op_income,
                "net_income": net_income,
                "pbr": pbr,
                "roe": roe,
            })

            status = f"mkt_cap={mkt_cap}" if mkt_cap is not None else "mkt_cap=N/A"
            print(f"  [Comps] {ticker} ({name}): {status}")

    print(f"[Comps] Loaded {len(comps)} comparable companies from {csv_path}")
    return comps
