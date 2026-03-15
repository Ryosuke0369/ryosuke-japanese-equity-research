"""
edinet_fetcher.py - EDINET API v2 client for downloading annual securities reports (XBRL).

Fetches 有価証券報告書 (Annual Securities Report) from the Financial Services Agency's
EDINET system using the official API v2, then extracts XBRL/iXBRL files for parsing.

Supports fetching up to 5 years of annual reports for multi-year financial analysis.

Usage:
    export EDINET_API_KEY="your-subscription-key-here"
    python scripts/edinet_fetcher.py 2359           # Fetch 5 years & display merged data
    python scripts/edinet_fetcher.py 2359 --years 3  # Fetch 3 years

API Reference:
    - Documents List: GET https://api.edinet-fsa.go.jp/api/v2/documents.json
    - Document Download: GET https://api.edinet-fsa.go.jp/api/v2/documents/{docID}
"""

import os
import sys
import time
import zipfile
import logging
import requests
from collections import OrderedDict
from datetime import date, timedelta
from dotenv import load_dotenv

# Load environment variables from .env file automatically
load_dotenv()

logger = logging.getLogger(__name__)

# =====================================================================
# CONSTANTS
# =====================================================================
BASE_URL = "https://api.edinet-fsa.go.jp/api/v2"
DOCUMENTS_LIST_URL = f"{BASE_URL}/documents.json"
DOCUMENT_DOWNLOAD_URL = f"{BASE_URL}/documents"

# docTypeCode for 有価証券報告書 (Annual Securities Report)
DOC_TYPE_ANNUAL_REPORT = "120"
# docTypeCode for 四半期報告書 (Quarterly Securities Report, abolished Apr 2024)
DOC_TYPE_QUARTERLY_REPORT = "140"
# docTypeCode for 半期報告書 (Semi-Annual Securities Report, old system)
DOC_TYPE_SEMIANNUAL_REPORT = "130"
# docTypeCode for 半期報告書 (Semi-Annual Securities Report, NEW system post-Apr 2024)
DOC_TYPE_SEMIANNUAL_REPORT_NEW = "160"

# EDINET secCode is 5 digits (ticker + trailing "0"), e.g. 2359 -> "23590"
SEC_CODE_SUFFIX = "0"

# Rate limiting: EDINET API has a limit of roughly 1-2 requests per second
REQUEST_DELAY_SEC = 0.5


# =====================================================================
# EXCEPTIONS
# =====================================================================
class EdinetApiError(Exception):
    """Base exception for EDINET API errors."""


class EdinetApiKeyMissing(EdinetApiError):
    """Raised when EDINET_API_KEY environment variable is not set."""


class EdinetDocumentNotFound(EdinetApiError):
    """Raised when no matching document is found for the given criteria."""


class EdinetRateLimitError(EdinetApiError):
    """Raised when API rate limit (HTTP 429) is hit."""


# =====================================================================
# API KEY
# =====================================================================
def _get_api_key():
    """Read EDINET API Subscription Key from environment variable.

    Returns:
        str: The API key.

    Raises:
        EdinetApiKeyMissing: If EDINET_API_KEY is not set.
    """
    key = os.environ.get("EDINET_API_KEY")
    if not key:
        raise EdinetApiKeyMissing(
            "EDINET_API_KEY environment variable is not set.\n"
            "Get your free Subscription Key at: https://disclosure2dl.edinet-fsa.go.jp/guide/static/disclosure/WZEK0110.html\n"
            "Then set it:\n"
            "  export EDINET_API_KEY='your-key-here'  (Linux/Mac)\n"
            '  set EDINET_API_KEY=your-key-here       (Windows CMD)\n'
            "  $env:EDINET_API_KEY='your-key-here'    (PowerShell)"
        )
    return key


# =====================================================================
# CORE FUNCTIONS
# =====================================================================
def _search_single_date(api_key, target_date, sec_code, doc_type_code=DOC_TYPE_ANNUAL_REPORT):
    """Query EDINET documents.json for a single date and return matching documents.

    Returns:
        list[dict]: Matching documents (may be empty).
        None: If the request should be retried (rate limit).

    Raises:
        EdinetApiError: On auth failure.
    """
    date_str = target_date.strftime("%Y-%m-%d")
    params = {"date": date_str, "type": 2, "Subscription-Key": api_key}

    try:
        resp = requests.get(DOCUMENTS_LIST_URL, params=params, timeout=30)
    except requests.RequestException as e:
        logger.warning("Network error on %s: %s", date_str, e)
        return []

    if resp.status_code == 401:
        raise EdinetApiError(
            "HTTP 401 Unauthorized: API key is invalid or expired. "
            "Please verify your EDINET_API_KEY."
        )
    if resp.status_code == 429:
        logger.warning("Rate limit hit on %s. Waiting 5s...", date_str)
        time.sleep(5)
        return None  # signal retry
    if resp.status_code != 200:
        logger.warning("HTTP %d on %s, skipping.", resp.status_code, date_str)
        return []

    matches = []
    for doc in resp.json().get("results", []):
        if (doc.get("secCode") == sec_code
                and (doc_type_code is None or doc.get("docTypeCode") == doc_type_code)):
            matches.append({
                "doc_id": doc["docID"],
                "filer_name": doc.get("filerName", ""),
                "doc_description": doc.get("docDescription", ""),
                "submit_date": doc.get("submitDateTime", ""),
                "period_end": doc.get("periodEnd", ""),
                "edinet_code": doc.get("edinetCode", ""),
                "doc_type_code_raw": doc.get("docTypeCode", ""),
            })
    return matches


def get_document_ids(ticker_code, num_years=5):
    """Find annual report docIDs for the past `num_years` years.

    Two-phase targeted search:
      Phase 1: Search peak filing dates (20th-28th) in the most common filing
               month (June) for each of the last num_years+1 calendar years.
               If the first report is found, immediately switch to Phase 2.
      Phase 2: Adaptive — use the discovered submit_date pattern to predict
               and search ±5 days around expected filing dates for all remaining
               years. Typically finds each report in 1-3 API calls.
      Fallback: If Phase 1 (June) yields nothing, try other filing seasons
               (March, September, December) with the same strategy.

    Expected API calls: ~15-40 for March FY companies (~8-20 seconds).

    Args:
        ticker_code: Stock ticker code (e.g. "2359" or 2359).
        num_years: Number of annual reports to find (default: 5, max: 5).

    Returns:
        list[dict]: Document metadata sorted newest-first.
    """
    api_key = _get_api_key()
    sec_code = str(ticker_code).strip() + SEC_CODE_SUFFIX
    today = date.today()
    current_year = today.year

    found_docs = []
    seen_period_ends = set()
    searched_dates = set()
    api_calls = 0

    def search_date(d):
        """Search a single date, returns True if new doc(s) found."""
        nonlocal api_calls
        if d in searched_dates or d > today or d.weekday() >= 5:
            return False
        searched_dates.add(d)

        result = _search_single_date(api_key, d, sec_code)
        api_calls += 1

        if result is None:  # rate limited
            time.sleep(5)
            result = _search_single_date(api_key, d, sec_code)
            api_calls += 1
            if result is None:
                return False

        found_new = False
        for doc in result:
            pe = doc["period_end"]
            if pe not in seen_period_ends:
                seen_period_ends.add(pe)
                found_docs.append(doc)
                found_new = True
                logger.info("Found [%d/%d]: docID=%s, period=%s (API calls: %d)",
                            len(found_docs), num_years, doc["doc_id"], pe, api_calls)

        time.sleep(REQUEST_DELAY_SEC)
        return found_new

    def search_window(year, month_start, day_start, month_end, day_end):
        """Search a date range for a given year. Returns True if doc found."""
        end_year = year + 1 if month_end < month_start else year
        try:
            d = date(year, month_start, day_start)
            end = date(end_year, month_end, day_end)
        except ValueError:
            return False
        while d <= end:
            if len(found_docs) >= num_years:
                return True
            if search_date(d):
                return True
            d += timedelta(days=1)
        return False

    def adaptive_search(reference_submit_date):
        """Phase 2: Use known filing date to predict and find remaining years."""
        ref_month = reference_submit_date.month
        ref_day = reference_submit_date.day
        logger.info("Phase 2 (adaptive): filing pattern = month %d, day ~%d",
                     ref_month, ref_day)

        for year in range(current_year, current_year - num_years - 2, -1):
            if len(found_docs) >= num_years:
                break
            try:
                predicted = date(year, ref_month, ref_day)
            except ValueError:
                predicted = date(year, ref_month, 28)
            # Search ±5 days around predicted date, closest first
            for delta in [0, -1, 1, -2, 2, -3, 3, -4, 4, -5, 5]:
                if len(found_docs) >= num_years:
                    break
                search_date(predicted + timedelta(days=delta))

    logger.info("Searching for %d annual reports: secCode=%s (ticker=%s)",
                num_years, sec_code, ticker_code)

    # Filing seasons: (peak_start_month, peak_start_day, peak_end_month, peak_end_day)
    # Ordered by frequency in Japanese market
    SEASONS = [
        (6, 18, 7, 2),    # March FY → June-July (~70% of companies)
        (3, 18, 4, 2),    # December FY → March-April
        (9, 18, 10, 2),   # June FY → Sep-Oct
        (12, 18, 1, 8),   # September FY → Dec-Jan
    ]

    for season_idx, (ms, ds, me, de) in enumerate(SEASONS):
        if len(found_docs) >= num_years:
            break

        logger.info("Trying filing season %d/%d (month %d-%d)...",
                     season_idx + 1, len(SEASONS), ms, me)

        # Search this season's peak window for each year (newest first)
        first_found_in_season = False
        for year in range(current_year, current_year - num_years - 2, -1):
            if len(found_docs) >= num_years:
                break
            if search_window(year, ms, ds, me, de):
                first_found_in_season = True
                break  # found one — switch to adaptive search

        # If we found a report in this season, use adaptive search for the rest
        if first_found_in_season and len(found_docs) < num_years:
            # Extract the submit date from the most recently found doc
            last_submit = found_docs[-1].get("submit_date", "")
            if len(last_submit) >= 10:
                try:
                    ref_date = date.fromisoformat(last_submit[:10])
                    adaptive_search(ref_date)
                except ValueError:
                    pass

    if not found_docs:
        raise EdinetDocumentNotFound(
            f"No 有価証券報告書 (docTypeCode={DOC_TYPE_ANNUAL_REPORT}) found "
            f"for secCode={sec_code} (ticker={ticker_code}) "
            f"in any filing season window within the last {num_years + 1} years."
        )

    found_docs.sort(key=lambda d: d["period_end"], reverse=True)
    logger.info("Done: %d report(s) for ticker=%s in %d API calls",
                len(found_docs), ticker_code, api_calls)
    return found_docs[:num_years]


def get_latest_document_id(ticker_code, search_days=400):
    """Find the latest annual report docID. (Backward-compatible wrapper.)"""
    docs = get_document_ids(ticker_code, num_years=1)
    return docs[0]


def get_latest_interim_id(ticker_code, fiscal_year_end=None):
    """Find the most recent interim report for a ticker.

    Searches for:
      - docTypeCode=160 (半期報告書, new semi-annual post-Apr 2024)
      - docTypeCode=140 (四半期報告書, legacy Q1/Q2/Q3, abolished Apr 2024)
      - docTypeCode=130 (半期報告書, old semi-annual)
      - Broad secCode-based search for post-2024 Q1/Q3 (any doc type)

    If fiscal_year_end is provided (e.g. "2024-03-31"), uses Adaptive Search:
      1. Compute the 3 quarter-end dates from FY end
      2. Estimate filing date = quarter_end + 45 days
      3. Search ±5 days around each estimated filing date
    Otherwise falls back to window-based search.

    Returns the newest one found (by period_end), or None.
    """
    api_key = _get_api_key()
    sec_code = str(ticker_code).strip() + SEC_CODE_SUFFIX
    today = date.today()
    searched_dates = set()
    api_calls = 0

    logger.info("Searching for latest interim report (160/140/130): secCode=%s", sec_code)

    doc_types = [DOC_TYPE_SEMIANNUAL_REPORT_NEW, DOC_TYPE_SEMIANNUAL_REPORT, DOC_TYPE_QUARTERLY_REPORT]
    MAX_API_CALLS = 50  # safety budget across all search phases

    # Interim doc types that may contain financial data (for broad search)
    INTERIM_DOC_TYPES = {
        DOC_TYPE_QUARTERLY_REPORT, DOC_TYPE_SEMIANNUAL_REPORT,
        DOC_TYPE_SEMIANNUAL_REPORT_NEW,
    }

    def _try_date(d, doc_types_to_search):
        """Search a single date for interim reports. Returns (doc, calls) or (None, calls).

        If doc_types_to_search contains None, does a broad search (all doc types)
        and filters results to INTERIM_DOC_TYPES.
        """
        calls = 0
        if d > today or d.weekday() >= 5 or d in searched_dates:
            return None, 0
        searched_dates.add(d)
        for doc_type in doc_types_to_search:
            result = _search_single_date(api_key, d, sec_code, doc_type)
            calls += 1
            if result is None:
                # 429 rate limit — _search_single_date already slept 5s, retry once
                result = _search_single_date(api_key, d, sec_code, doc_type)
                calls += 1
            if result is None:
                # Still rate-limited after retry — give up on this date
                logger.warning("Persistent rate limit on %s, skipping.", d)
                return None, calls
            if result:
                # If broad search (doc_type=None), filter to interim doc types
                if doc_type is None:
                    result = [r for r in result
                              if r.get("doc_type_code_raw") in INTERIM_DOC_TYPES]
                    if not result:
                        time.sleep(REQUEST_DELAY_SEC)
                        continue
                doc = result[0]
                doc["doc_type_code"] = doc.get("doc_type_code_raw", doc_type)
                logger.info("Found interim: docType=%s, docID=%s, period=%s, desc=%s",
                            doc["doc_type_code"], doc["doc_id"], doc["period_end"],
                            doc["doc_description"])
                return doc, calls
            time.sleep(REQUEST_DELAY_SEC)
        return None, calls

    # --- Adaptive Search: use fiscal_year_end to predict filing dates ---
    if fiscal_year_end:
        try:
            fy_month = int(fiscal_year_end[5:7])
        except (ValueError, IndexError):
            fy_month = None

        if fy_month:
            import calendar
            # Compute quarter-end months from FY start.
            # FY start = fy_month + 1 (e.g. March FY -> April start)
            # Q1 end = end of month FY_start+2, Q2 = +5, Q3 = +8
            fy_start_month = (fy_month % 12) + 1
            quarter_months = []
            for q_offset in [2, 5, 8]:
                qm = ((fy_start_month - 1 + q_offset) % 12) + 1
                quarter_months.append(qm)

            # Build candidate filing dates, newest quarter first
            # quarter_months = [Q1_end, Q2_end, Q3_end]
            candidate_dates = []
            fy_end_year = int(fiscal_year_end[:4])
            for year_offset in [1, 0, -1]:
                base_year = fy_end_year + year_offset
                # Iterate Q3, Q2, Q1 (newest first)
                for qi, qm in enumerate(reversed(quarter_months)):
                    q_label = ["Q3", "Q2", "Q1"][qi]
                    q_year = base_year - 1 if qm > fy_month else base_year
                    q_day = calendar.monthrange(q_year, qm)[1]
                    q_end = date(q_year, qm, q_day)
                    est_filing = q_end + timedelta(days=45)
                    # Skip future or too old (>18 months)
                    if est_filing > today or (today - est_filing).days > 540:
                        continue
                    # Post-Apr 2024 reform:
                    #   Q2: semi-annual (160) filed to EDINET
                    #   Q1/Q3: typically TDNet only, but do broad search as fallback
                    # Pre-Apr 2024: search old quarterly (140) and semi-annual (130)
                    if q_end >= date(2024, 4, 1):
                        if q_label == "Q2":
                            q_doc_types = [DOC_TYPE_SEMIANNUAL_REPORT_NEW]
                        else:
                            # Broad search: None = match any doc type for this secCode
                            q_doc_types = [None]
                    else:
                        q_doc_types = [DOC_TYPE_QUARTERLY_REPORT, DOC_TYPE_SEMIANNUAL_REPORT]
                    candidate_dates.append((q_end, est_filing, q_doc_types))

            logger.info("Adaptive interim search: %d candidate quarter-ends", len(candidate_dates))

            # Build spiral offsets: 0, -1, +1, -2, +2, ...±20
            offsets = [0]
            for i in range(1, 21):
                offsets.extend([-i, i])

            for q_end, est_filing, q_doc_types in candidate_dates:
                if api_calls >= MAX_API_CALLS:
                    logger.info("Adaptive search: hit API call budget (%d), stopping.", api_calls)
                    break
                # Try each doc type separately across all offsets
                # (1 API call per date instead of 2)
                for doc_type in q_doc_types:
                    found = False
                    for offset in offsets:
                        if api_calls >= MAX_API_CALLS:
                            break
                        d = est_filing + timedelta(days=offset)
                        doc, calls = _try_date(d, [doc_type])
                        api_calls += calls
                        if doc:
                            logger.info("Adaptive interim search: found in %d API calls", api_calls)
                            return doc
                    if found:
                        break

            if api_calls < MAX_API_CALLS:
                logger.info("Adaptive interim search exhausted (%d API calls), "
                            "falling back to window search", api_calls)
            else:
                logger.info("API call budget exhausted (%d calls), no interim found.", api_calls)
                return None

    # --- Fallback: window-based search ---
    current_year = today.year

    INTERIM_WINDOWS = [
        (2, 1, 2, 28),    # Q3 filing (Jan-Feb) or semi-annual for Sep FY
        (1, 15, 1, 31),   # Q3 filing (late Jan)
        (11, 1, 11, 30),  # Q2/semi-annual filing (Nov) for Mar FY
        (10, 15, 10, 31), # Q2/semi-annual filing (late Oct)
        (8, 1, 8, 31),    # Q1 filing (Aug) — legacy only
        (7, 15, 7, 31),   # Q1 filing (late Jul) — legacy only
    ]

    for ms, ds, me, de in INTERIM_WINDOWS:
        if api_calls >= MAX_API_CALLS:
            break
        for year in range(current_year, current_year - 2, -1):
            if api_calls >= MAX_API_CALLS:
                break
            try:
                d = date(year, ms, ds)
                end = date(year, me, de)
            except ValueError:
                continue
            while d <= end:
                if api_calls >= MAX_API_CALLS:
                    break
                doc, calls = _try_date(d, doc_types)
                api_calls += calls
                if doc:
                    logger.info("Window interim search: found in %d API calls", api_calls)
                    return doc
                d += timedelta(days=1)

    logger.info("No interim report found for secCode=%s (%d API calls)", sec_code, api_calls)
    return None


def fetch_tanshin(securities_code, edinet_api_key=None):
    """Fetch the latest 決算短信 (earnings summary) containing forecast data.

    Searches for docTypeCode="140" (四半期報告書/決算短信) and downloads the XBRL.
    Then extracts forecast/guidance data (業績予想) from the XBRL.

    Args:
        securities_code: Stock ticker code (e.g. "7974" or 7974).
        edinet_api_key: Optional API key override. If None, reads from environment.

    Returns:
        dict with keys:
            'forecast_data': {forecast_revenue, forecast_operating_income, ...} in JPY mn
            'doc_id': EDINET document ID
            'period_end': Period end date string
        Returns None if no tanshin found or no forecast data extracted.
    """
    if edinet_api_key:
        os.environ["EDINET_API_KEY"] = edinet_api_key
    api_key = _get_api_key()
    sec_code = str(securities_code).strip() + SEC_CODE_SUFFIX
    today = date.today()

    logger.info("Searching for 決算短信 (docTypeCode=140): secCode=%s", sec_code)

    # Search recent dates for docTypeCode="140"
    # Also try annual report docs (120) as they often contain forecast contexts
    DOC_TYPES_TO_TRY = [DOC_TYPE_QUARTERLY_REPORT, DOC_TYPE_ANNUAL_REPORT]
    found_doc = None
    api_calls = 0

    for doc_type in DOC_TYPES_TO_TRY:
        if found_doc:
            break
        # Search backwards from today, check peak filing windows
        for months_back in range(0, 18):
            if found_doc or api_calls >= 30:
                break
            search_date_obj = today - timedelta(days=months_back * 15)
            # Search a 5-day window around each candidate date
            for delta in range(0, 10):
                d = search_date_obj - timedelta(days=delta)
                if d > today or d.weekday() >= 5:
                    continue
                result = _search_single_date(api_key, d, sec_code, doc_type)
                api_calls += 1
                if result is None:
                    time.sleep(3)
                    continue
                if result:
                    found_doc = result[0]
                    logger.info("Found tanshin candidate: docID=%s, period=%s, type=%s",
                                found_doc["doc_id"], found_doc["period_end"], doc_type)
                    break
                time.sleep(REQUEST_DELAY_SEC)

    if not found_doc:
        logger.info("No 決算短信 found for secCode=%s (%d API calls)", sec_code, api_calls)
        return None

    # Download and extract XBRL
    try:
        dl_result = download_and_extract_xbrl(found_doc["doc_id"])
    except EdinetApiError as e:
        logger.warning("Failed to download tanshin docID=%s: %s", found_doc["doc_id"], e)
        return None

    xbrl_files = [f for f in dl_result["xbrl_files"] if f.lower().endswith(".xbrl")]
    if not xbrl_files:
        logger.warning("No XBRL files in tanshin docID=%s", found_doc["doc_id"])
        return None

    # Parse and extract forecast data
    try:
        from scripts.edinet_parser import parse_xbrl_file, extract_forecast_data
    except ImportError:
        from edinet_parser import parse_xbrl_file, extract_forecast_data

    soup = parse_xbrl_file(xbrl_files[0])
    forecast_data = extract_forecast_data(soup)

    if not forecast_data:
        logger.info("No forecast data found in tanshin docID=%s", found_doc["doc_id"])
        return None

    return {
        "forecast_data": forecast_data,
        "doc_id": found_doc["doc_id"],
        "period_end": found_doc["period_end"],
    }


def download_and_extract_xbrl(doc_id, output_dir=None):
    """Download a full disclosure ZIP from EDINET and extract XBRL files.

    Args:
        doc_id: EDINET document ID string (e.g. "S100XXXX").
        output_dir: Directory to extract files into. Defaults to a temp directory
                    under the script's parent folder.

    Returns:
        dict with keys:
            'extract_dir': Path to the extracted root directory.
            'xbrl_files': List of absolute paths to .xbrl and iXBRL (.htm) files
                          found in XBRL/PublicDoc/.
            'doc_id': The document ID used.

    Raises:
        EdinetApiKeyMissing: If API key is not configured.
        EdinetApiError: On download or extraction failure.
    """
    api_key = _get_api_key()

    if output_dir is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, "..", "tmp", "edinet_data")

    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # Download ZIP (type=1: full submission package including XBRL)
    download_url = f"{DOCUMENT_DOWNLOAD_URL}/{doc_id}"
    params = {
        "type": 1,
        "Subscription-Key": api_key,
    }

    logger.info("Downloading ZIP for docID=%s ...", doc_id)

    try:
        resp = requests.get(download_url, params=params, timeout=120, stream=True)
    except requests.RequestException as e:
        raise EdinetApiError(f"Network error downloading docID={doc_id}: {e}") from e

    if resp.status_code == 401:
        raise EdinetApiError(
            "HTTP 401 Unauthorized: API key is invalid or expired."
        )
    if resp.status_code == 404:
        raise EdinetApiError(
            f"HTTP 404 Not Found: Document docID={doc_id} does not exist on EDINET."
        )
    if resp.status_code == 429:
        raise EdinetRateLimitError(
            "HTTP 429 Too Many Requests: EDINET API rate limit exceeded. "
            "Please wait and try again."
        )
    if resp.status_code != 200:
        raise EdinetApiError(
            f"HTTP {resp.status_code} error downloading docID={doc_id}: "
            f"{resp.text[:200]}"
        )

    # Verify we received a ZIP file
    content_type = resp.headers.get("Content-Type", "")
    if "zip" not in content_type and "octet-stream" not in content_type:
        raise EdinetApiError(
            f"Unexpected Content-Type '{content_type}' for docID={doc_id}. "
            "Expected a ZIP file (type=1). The API may have returned an error page."
        )

    # Save ZIP
    zip_path = os.path.join(output_dir, f"{doc_id}.zip")
    total_bytes = 0
    with open(zip_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)
            total_bytes += len(chunk)

    logger.info("Downloaded %s (%.1f MB)", zip_path, total_bytes / 1_048_576)

    # Extract ZIP
    extract_dir = os.path.join(output_dir, doc_id)
    if os.path.exists(extract_dir):
        import shutil
        shutil.rmtree(extract_dir)

    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(extract_dir)
    except zipfile.BadZipFile as e:
        raise EdinetApiError(
            f"Downloaded file for docID={doc_id} is not a valid ZIP: {e}"
        ) from e

    logger.info("Extracted to: %s", extract_dir)

    # Find XBRL/iXBRL files in XBRL/PublicDoc/
    xbrl_files = _find_xbrl_files(extract_dir)

    if not xbrl_files:
        logger.warning(
            "No .xbrl or .htm files found in XBRL/PublicDoc/. "
            "Searching entire extracted directory..."
        )
        xbrl_files = _find_xbrl_files(extract_dir, search_all=True)

    logger.info("Found %d XBRL/iXBRL file(s):", len(xbrl_files))
    for f in xbrl_files:
        logger.info("  %s", f)

    # Clean up ZIP file (keep extracted data)
    os.remove(zip_path)
    logger.info("Cleaned up ZIP file: %s", zip_path)

    return {
        "extract_dir": extract_dir,
        "xbrl_files": xbrl_files,
        "doc_id": doc_id,
    }


def _find_xbrl_files(extract_dir, search_all=False):
    """Locate XBRL/iXBRL files within the extracted directory.

    EDINET ZIP structure (type=1):
        {docID}/
            XBRL/
                PublicDoc/          <- Main disclosure documents
                    *.xbrl          <- XBRL instance documents
                    *.htm           <- Inline XBRL (iXBRL) documents
                    *.xsd           <- Schema files
                AuditDoc/           <- Audit-related documents
            ...

    Args:
        extract_dir: Root of extracted ZIP.
        search_all: If True, search the entire directory tree (fallback).

    Returns:
        List of absolute paths to .xbrl and .htm files (excluding schemas/stylesheets).
    """
    xbrl_files = []

    if not search_all:
        # Standard path: XBRL/PublicDoc/
        public_doc_dir = os.path.join(extract_dir, "XBRL", "PublicDoc")
        if not os.path.isdir(public_doc_dir):
            return []

        for fname in os.listdir(public_doc_dir):
            fpath = os.path.join(public_doc_dir, fname)
            if not os.path.isfile(fpath):
                continue
            lower = fname.lower()
            # Include XBRL instance documents and Inline XBRL (htm) files
            # Exclude: .xsd (schema), .xml (linkbase), .css, .js
            if lower.endswith(".xbrl") or (lower.endswith(".htm") and "ixbrl" not in lower):
                xbrl_files.append(os.path.abspath(fpath))
    else:
        # Fallback: search entire tree
        for dirpath, _dirs, filenames in os.walk(extract_dir):
            for fname in filenames:
                lower = fname.lower()
                if lower.endswith(".xbrl") or lower.endswith(".htm"):
                    xbrl_files.append(os.path.abspath(os.path.join(dirpath, fname)))

    # Sort for deterministic output (instance docs first, then htm)
    xbrl_files.sort(key=lambda p: (not p.lower().endswith(".xbrl"), p))
    return xbrl_files


def fetch_and_parse_multi_year(ticker_code, num_years=5, output_dir=None):
    """Fetch multiple years of annual reports + latest quarterly, return merged data with LTM.

    Downloads up to `num_years` annual reports and the latest quarterly report,
    parses each XBRL file, merges annual data, and computes LTM if quarterly
    data is available.

    Args:
        ticker_code: Stock ticker code (e.g. "2359").
        num_years: Number of years to fetch (default: 5).
        output_dir: Directory for downloaded files (default: tmp/edinet_data).

    Returns:
        tuple: (company_info, merged_data) where merged_data is an OrderedDict
               with LTM column (if available) followed by FY columns.
    """
    try:
        from scripts.edinet_parser import (
            parse_xbrl_file, identify_clean_contexts, extract_financial_data,
            extract_company_info, merge_multi_year_data,
            identify_quarterly_contexts, extract_quarterly_data, calculate_ltm,
        )
    except ImportError:
        from edinet_parser import (
            parse_xbrl_file, identify_clean_contexts, extract_financial_data,
            extract_company_info, merge_multi_year_data,
            identify_quarterly_contexts, extract_quarterly_data, calculate_ltm,
        )

    # Step 1: Find annual report document IDs
    doc_infos = get_document_ids(ticker_code, num_years=num_years)

    print(f"\nFound {len(doc_infos)} annual report(s):")
    for i, d in enumerate(doc_infos, 1):
        print(f"  [{i}] docID={d['doc_id']}  period={d['period_end']}  filer={d['filer_name']}")

    # Step 2: Download and extract annual reports
    xbrl_paths_by_period = []
    for doc_info in doc_infos:
        doc_id = doc_info["doc_id"]
        period_end = doc_info["period_end"]

        try:
            result = download_and_extract_xbrl(doc_id, output_dir)
        except EdinetApiError as e:
            logger.warning("Failed to download docID=%s: %s", doc_id, e)
            continue

        xbrl_files = [f for f in result["xbrl_files"] if f.lower().endswith(".xbrl")]
        if xbrl_files:
            xbrl_paths_by_period.append((period_end, xbrl_files[0]))
            print(f"  Downloaded: {doc_id} -> {os.path.basename(xbrl_files[0])}")

        time.sleep(REQUEST_DELAY_SEC)

    if not xbrl_paths_by_period:
        raise EdinetApiError("No XBRL files could be downloaded.")

    xbrl_paths_by_period.sort(key=lambda x: x[0], reverse=True)

    # Step 3: Parse annual XBRL files
    all_year_data = []
    company_info = None

    for period_end, xbrl_path in xbrl_paths_by_period:
        soup = parse_xbrl_file(xbrl_path)
        if company_info is None:
            company_info = extract_company_info(soup)
        contexts = identify_clean_contexts(soup)
        data = extract_financial_data(soup, contexts)
        all_year_data.append((period_end, data))

    # Step 4: Merge annual data
    merged = merge_multi_year_data(all_year_data)

    # Step 5: Search for latest quarterly report and compute LTM
    print("\nSearching for latest interim report (quarterly/semi-annual)...")
    fiscal_year_end = doc_infos[0]["period_end"] if doc_infos else None
    quarterly_doc = get_latest_interim_id(ticker_code, fiscal_year_end=fiscal_year_end)

    if quarterly_doc:
        print(f"  Found: docID={quarterly_doc['doc_id']}  "
              f"period={quarterly_doc['period_end']}  "
              f"desc={quarterly_doc['doc_description']}")

        try:
            q_result = download_and_extract_xbrl(quarterly_doc["doc_id"], output_dir)
            q_xbrl_files = [f for f in q_result["xbrl_files"] if f.lower().endswith(".xbrl")]

            if q_xbrl_files:
                q_soup = parse_xbrl_file(q_xbrl_files[0])
                q_contexts = identify_quarterly_contexts(q_soup)

                if q_contexts:
                    q_data = extract_quarterly_data(q_soup, q_contexts)

                    # Get the latest FY data for LTM calculation
                    fy_keys = [k for k in merged if k.startswith("FY")]
                    if fy_keys:
                        latest_fy_key = fy_keys[0]
                        latest_fy_data = merged[latest_fy_key]

                        ltm_data, ltm_label = calculate_ltm(
                            latest_fy_data, q_data,
                            quarterly_doc["period_end"]
                        )

                        if ltm_data:
                            # Insert LTM as first column
                            new_merged = OrderedDict()
                            new_merged[ltm_label] = ltm_data
                            for k, v in merged.items():
                                if k != "_meta":
                                    new_merged[k] = v
                            new_merged["_meta"] = merged.get("_meta", {})
                            merged = new_merged
                            print(f"  LTM computed: {ltm_label}")
                else:
                    logger.warning("No quarterly contexts found in XBRL.")
        except EdinetApiError as e:
            logger.warning("Failed to process quarterly report: %s", e)
    else:
        print("  No quarterly report found (may already be latest FY).")

    return company_info, merged


# =====================================================================
# CLI ENTRY POINT
# =====================================================================
def main():
    """CLI interface for edinet_fetcher."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    # Check API key first
    try:
        _get_api_key()
    except EdinetApiKeyMissing as e:
        print(f"\n{e}\n")
        print("=" * 60)
        print("EDINET Fetcher - Multi-Year Annual Report Downloader")
        print("=" * 60)
        print()
        print("Usage:")
        print("  1. Set your EDINET API key:")
        print("       export EDINET_API_KEY='your-subscription-key'")
        print()
        print("  2. Run with a ticker code:")
        print("       python scripts/edinet_fetcher.py 2359")
        print("       python scripts/edinet_fetcher.py 2359 --years 3")
        print()
        print("  The script will:")
        print("    - Search EDINET for up to 5 years of 有価証券報告書")
        print("    - Download and parse each XBRL file")
        print("    - Display merged multi-year financial data")
        print()
        print("Get your free API key at:")
        print("  https://disclosure2dl.edinet-fsa.go.jp/guide/static/disclosure/WZEK0110.html")
        sys.exit(1)

    # Parse arguments
    import argparse
    parser = argparse.ArgumentParser(description="EDINET multi-year fetcher & parser")
    parser.add_argument("ticker", help="Stock ticker code (e.g. 2359)")
    parser.add_argument("--years", type=int, default=5, help="Number of years to fetch (default: 5)")
    parser.add_argument("--output-dir", default=None, help="Output directory for downloads")
    args = parser.parse_args()

    ticker_code = args.ticker.strip()
    num_years = min(args.years, 5)

    print(f"\n{'=' * 60}")
    print(f"EDINET Fetcher - {num_years}-Year Report for ticker: {ticker_code}")
    print(f"{'=' * 60}\n")

    try:
        company_info, merged_data = fetch_and_parse_multi_year(
            ticker_code, num_years=num_years, output_dir=args.output_dir
        )
    except EdinetDocumentNotFound as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
    except EdinetApiError as e:
        print(f"\nAPI ERROR: {e}")
        sys.exit(1)

    # Display merged results
    try:
        from scripts.edinet_parser import print_results
    except ImportError:
        from edinet_parser import print_results
    print_results(company_info, merged_data)

    print(f"\n{'=' * 60}")
    print("Done! Multi-year financial data extracted successfully.")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
