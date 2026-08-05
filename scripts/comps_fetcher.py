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


def _normalize_ticker(ticker):
    """Append the Tokyo Stock Exchange suffix `.T` to a bare 4-digit code.

    yfinance returns HTTP 404 for bare codes (e.g. "5038"); it needs "5038.T".
    Tickers that already carry a suffix (".T", ".JP", etc.) are returned as-is.
    """
    t = ticker.strip()
    if "." not in t and t[:4].isdigit():
        return f"{t}.T"
    return t


def _fetch_market_cap(ticker_str):
    """Fetch market cap for a single ticker via yfinance.

    Returns market cap in JPY millions, or None on failure.
    Fallback: currentPrice * sharesOutstanding.
    """
    if not YFINANCE_AVAILABLE:
        logger.warning("yfinance not installed. Cannot fetch market cap for %s.", ticker_str)
        return None

    ticker_str = _normalize_ticker(ticker_str)
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
    """Load comparable company data from CSV.

    Args:
        csv_path: Path to UTF-8 comma-delimited CSV with columns:
                  Ticker, Name, Revenue, EBITDA, Operating_Income,
                  Net_Income, Book_Value, Net_Debt
                  Optional column Market_Cap (JPY mn): when present it must be
                  filled for ALL rows (a partial column raises ValueError) and
                  is used as-is with no yfinance call; when absent, market caps
                  are fetched live from yfinance with a reproducibility warning.

    Returns:
        List of dicts with keys: name, ticker, mkt_cap, ev, revenue,
        ebitda, op_income, net_income, pbr, roe
    """
    comps = []

    # Read bytes and decode defensively: files exported from Excel/PowerShell
    # are often UTF-16 with a BOM, which would garble a plain utf-8 open().
    with open(csv_path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "utf-16"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")

    # Sanitize: strip trailing whitespace from each line before parsing.
    # Trailing tabs corrupt delimiter auto-detection and DictReader fields.
    clean_lines = [line.rstrip() for line in text.splitlines()]

    clean_content = "\n".join(clean_lines)
    with io.StringIO(clean_content) as f_clean:
        # Auto-detect delimiter (handles both comma and tab-separated files)
        sample = clean_lines[0] if clean_lines else ""
        delimiter = "\t" if "\t" in sample else ","
        reader = csv.DictReader(f_clean, delimiter=delimiter)

        # Market_Cap column contract: all rows filled, or the column omitted
        # entirely. A partially-filled column would silently mix static and
        # live-fetched market caps in one comps set, so it is a hard error.
        fieldnames = [fn.strip() for fn in (reader.fieldnames or [])]
        has_mkt_cap_col = ("Market_Cap" in fieldnames) or ("Market Cap" in fieldnames)
        if has_mkt_cap_col:
            print("[Comps] Market cap source: CSV (static)")
        else:
            print("[Comps] WARNING: Market cap source: yfinance live - 出力は実行時点で"
                  "変動する（再現性が必要なら Market_Cap 列を記入）")

        for row in reader:
            ticker = row["Ticker"].strip()
            name = row["Name"].strip()

            # Normalize column names: strip whitespace from keys
            row = {k.strip(): v.strip() for k, v in row.items()}

            revenue = float(row["Revenue"])
            # EBITDA may be left blank when D&A is unavailable for that peer —
            # the contract is "blank, never EBIT", because a copied EBIT silently
            # becomes an EV/EBIT multiple inside the EV/EBITDA median.
            _ebitda_raw = (row.get("EBITDA") or "").strip()
            ebitda = float(_ebitda_raw) if _ebitda_raw else None
            op_income = float(row.get("Operating_Income") or row.get("Operating Income", "0"))
            net_income = float(row.get("Net_Income") or row.get("Net Income", "0"))
            book_value = float(row.get("Book_Value") or row.get("Book Value", "0"))
            net_debt = float(row.get("Net_Debt") or row.get("Net Debt", "0"))

            # Market cap (JPY mn): static from the Market_Cap column when the
            # column exists (required for delisted/TOB names, e.g. LightWorks
            # 4267, and for reproducible outputs); yfinance live otherwise.
            if has_mkt_cap_col:
                mkt_cap_manual = (row.get("Market_Cap") or row.get("Market Cap") or "").strip()
                if not mkt_cap_manual:
                    raise ValueError(
                        f"Market_Cap column exists but is blank for {ticker} ({name}). "
                        f"Fill Market_Cap (JPY mn) for ALL rows or remove the column "
                        f"entirely — a static/live mix is not allowed."
                    )
                mkt_cap = float(mkt_cap_manual)
            else:
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
                # Book value is carried through so the template can write PBR and
                # ROE as formulas off a visible Book Value column instead of
                # freezing the ratios computed here.
                "book_value": book_value if book_value != 0 else None,
                "pbr": pbr,
                "roe": roe,
            })

            if ebitda is not None and ebitda == op_income and ebitda > 0:
                print(f"  [Comps] WARNING: {ticker} ({name}) has EBITDA == Operating "
                      f"Income - D&A was not added back. It will be excluded from "
                      f"EV/EBITDA statistics. (EBITDA = 営業利益 + 減価償却費; leave "
                      f"the cell blank if D&A is unavailable.)")
            if not book_value:
                print(f"  [Comps] WARNING: {ticker} ({name}) has no Book_Value - "
                      f"PBR/ROE will be blank for this row.")

            status = f"mkt_cap={mkt_cap}" if mkt_cap is not None else "mkt_cap=N/A"
            print(f"  [Comps] {ticker} ({name}): {status}")

    print(f"[Comps] Loaded {len(comps)} comparable companies from {csv_path}")
    return comps
