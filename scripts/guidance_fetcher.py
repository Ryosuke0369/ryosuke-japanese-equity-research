"""guidance_fetcher.py - 会社予想(業績予想) for the DCF pipeline.

Why this module exists
----------------------
The Management scenario is supposed to be "what the company itself says next
year looks like". In the 2026-09-05 batch it was identical to Base in **all 85
tickers**, because guidance was never obtained even once. Two independent
reasons, both structural:

1. **決算短信 is a TDnet document, not an EDINET one.** `fetch_tanshin()`
   searched EDINET for `docTypeCode=140` (四半期報告書) — a document class the
   FSA abolished in April 2024, and one that never carried 業績予想 in the first
   place. The search could not succeed; it just took ~50 seconds to fail.

2. **The 有報 scan was not scoped to the ticker.** Step 3 of generate_dcf.py
   globbed `tmp/edinet_data/**/*.xbrl` — every document ever downloaded, for
   every company (510 directories by 2026-09-06) — and took the first file that
   yielded a forecast. It found nothing (有価証券報告書 does not carry 業績予想),
   which cost 200 seconds per ticker; had it found one, it would have applied
   **another company's guidance** to this model.

Where the numbers actually live
-------------------------------
`screener/fetch/tdnet_archiver.py` already archives TDnet 決算短信 daily, and
`screener/extract/xbrl_parser.py` already parses 業績予想 out of the Summary
inline-XBRL into the `guidance` table. That table is the primary source here:
2,984 codes as of 2026-09-06, covering 78 of the 85 batch tickers.

The remaining 7 were checked one by one and are genuine absences, not a bug in
this module (6861 キーエンス publishes no guidance at all; 4901 / 7751 file
US-GAAP 短信 whose forecast facts the screener's mapping does not yet cover;
6506 / 9983 / 3086 have no 短信 inside the archive's retention window; 3863
issued its revision as a PDF-only 業績予想修正). For those the caller keeps the
existing behaviour — Management falls back to the CAGR path — and the reason is
recorded rather than passed over in silence.

Units: the `guidance` table stores yen. The DCF config wants JPY mn.
"""
import os
import re
import sqlite3
import sys

# guidance.item -> the key generate_dcf.py / dcf_comps_template.py expect.
# 'ordinary_income' and 'eps' are deliberately not mapped: nothing downstream
# reads them, and a key the template does not consume is a key that can drift.
ITEM_MAP = {
    "revenue": "forecast_revenue",
    "operating_income": "forecast_operating_income",
    "net_income_parent": "forecast_net_income",
}

YEN_TO_MN = 1_000_000.0


def _fy_year(label):
    """'FY2027' / 'FY2027/3' / 'FY2027(E)' -> 2027. None when unparseable."""
    m = re.search(r"(\d{4})", str(label or ""))
    return int(m.group(1)) if m else None


PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def _data_root_from_dotenv():
    """DATA_ROOT out of the repo .env, without importing the screener package."""
    p = os.path.join(PROJECT_ROOT, ".env")
    if not os.path.isfile(p):
        return None
    for line in open(p, encoding="utf-8", errors="replace"):
        line = line.strip()
        if line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        if k.strip() in ("SCREENER_DATA_ROOT", "DATA_ROOT"):
            return v.strip().strip('"').strip("'")
    return None


def screener_db_path():
    """Resolve the screener SQLite path, or None when it is not configured.

    Resolved defensively and in this order: an explicit SCREENER_DB, the
    screener package's own resolution, then DATA_ROOT from the process
    environment or the repo .env. The DCF pipeline must keep working on a
    machine where the screener has never been set up, so every step is
    optional and failure returns None rather than raising.
    """
    explicit = os.environ.get("SCREENER_DB")
    if explicit and os.path.isfile(explicit):
        return explicit
    try:
        if PROJECT_ROOT not in sys.path:
            sys.path.insert(0, PROJECT_ROOT)
        from screener import common as C
        if os.path.isfile(C.DB_PATH):
            return C.DB_PATH
    except Exception:
        pass
    root = (os.environ.get("SCREENER_DATA_ROOT")
            or os.environ.get("DATA_ROOT")
            or _data_root_from_dotenv())
    if root:
        p = os.path.join(root, "screener.db")
        if os.path.isfile(p):
            return p
    return None


def from_screener_db(ticker_code, min_fy_year=None, db_path=None):
    """Newest 会社予想 for `ticker_code` from the TDnet-derived guidance table.

    Args:
        ticker_code: 4-character TSE code, as stored in `guidance.code`.
        min_fy_year: If given, the guidance FY must be strictly newer than this
            (the latest actual FY). A 短信 for a year already closed is history,
            not guidance, and feeding it to the Management scenario would model
            the past as the future.
        db_path: Override for tests.

    Returns:
        (forecast_data, note). forecast_data is None when nothing usable was
        found; `note` always explains what happened, for the Adjustments Log.
    """
    path = db_path or screener_db_path()
    if not path:
        return None, "screener DB not configured (DATA_ROOT unset or screener.db absent)"

    code = str(ticker_code).strip()
    try:
        con = sqlite3.connect(f"file:{path}?mode=ro", uri=True, timeout=20)
    except sqlite3.Error as e:
        return None, f"screener DB unreadable: {type(e).__name__}: {e}"

    try:
        rows = con.execute(
            "SELECT date, fy, item, value, revision_direction "
            "FROM guidance WHERE code = ? ORDER BY date DESC, fy DESC",
            (code,),
        ).fetchall()
    except sqlite3.Error as e:
        return None, f"screener DB query failed: {type(e).__name__}: {e}"
    finally:
        con.close()

    if not rows:
        return None, f"no guidance row for code={code} in the TDnet archive"

    # Newest disclosure date wins; within it, the newest FY. A single 短信 files
    # both the interim and the full-year forecast, so restricting to one (date,
    # fy) pair is what keeps the full-year numbers from mixing with the interim.
    top_date = rows[0][0]
    fy_candidates = sorted(
        {r[1] for r in rows if r[0] == top_date},
        key=lambda f: (_fy_year(f) or 0),
        reverse=True,
    )
    if not fy_candidates:
        return None, f"guidance rows for code={code} carry no FY label"
    fy = fy_candidates[0]
    fy_year = _fy_year(fy)

    if min_fy_year is not None and fy_year is not None and fy_year <= min_fy_year:
        return None, (f"newest guidance is {fy} (disclosed {top_date}) but the "
                      f"latest actual FY is FY{min_fy_year} - that is history, "
                      f"not guidance; ignored")

    out = {}
    for date_s, fy_s, item, value, _rev in rows:
        if date_s != top_date or fy_s != fy:
            continue
        key = ITEM_MAP.get(item)
        if key and value is not None and key not in out:
            out[key] = float(value) / YEN_TO_MN

    if not out:
        return None, (f"guidance rows exist for code={code} ({fy} @ {top_date}) "
                      f"but none map to revenue / operating income / net income")

    got = ", ".join(f"{k.replace('forecast_', '')}={v:,.0f}mn" for k, v in sorted(out.items()))
    return out, f"TDnet 決算短信 via screener guidance table: {fy} disclosed {top_date} ({got})"


def from_cached_xbrl(xbrl_paths):
    """Scan THIS ticker's already-downloaded EDINET XBRL for forecast contexts.

    Kept as a second tier because a 有価証券報告書 occasionally does carry a
    ForecastMember context. The important part is `xbrl_paths`: the caller
    supplies the paths for this ticker only. The previous implementation globbed
    the whole cache and would have applied another company's forecast.
    """
    if not xbrl_paths:
        return None, "no cached EDINET XBRL for this ticker"
    try:
        from scripts.edinet_parser import parse_xbrl_file, extract_forecast_data
    except ImportError:
        from edinet_parser import parse_xbrl_file, extract_forecast_data

    for p in xbrl_paths:
        if not p or not os.path.isfile(p):
            continue
        try:
            fd = extract_forecast_data(parse_xbrl_file(p))
        except Exception:
            continue
        if fd and fd.get("forecast_revenue"):
            return fd, f"EDINET XBRL (this ticker): {os.path.basename(p)}"
    return None, (f"{len(xbrl_paths)} cached EDINET XBRL file(s) for this ticker "
                  f"carry no ForecastMember context")


def from_edinet_tanshin(ticker_code):
    """Legacy EDINET 短信 search. Off by default - see the module docstring."""
    try:
        from scripts.edinet_fetcher import fetch_tanshin
    except ImportError:
        from edinet_fetcher import fetch_tanshin
    try:
        r = fetch_tanshin(ticker_code)
    except Exception as e:
        return None, f"EDINET tanshin search failed: {type(e).__name__}: {e}"
    if not r or not r.get("forecast_data"):
        return None, "EDINET tanshin search found nothing (expected: 決算短信 is TDnet, not EDINET)"
    return r["forecast_data"], f"EDINET docID={r.get('doc_id')}"


def get_guidance(ticker_code, min_fy_year=None, xbrl_paths=None,
                 allow_edinet_tanshin=False):
    """Resolve 会社予想 through the source ladder. Never raises.

    Returns:
        (forecast_data | None, note, source_tag)
    """
    attempts = []

    fd, note = from_screener_db(ticker_code, min_fy_year=min_fy_year)
    if fd:
        return fd, note, "tdnet_screener_db"
    attempts.append(f"screener DB: {note}")

    fd, note = from_cached_xbrl(xbrl_paths)
    if fd:
        return fd, note, "edinet_xbrl"
    attempts.append(f"EDINET XBRL: {note}")

    if allow_edinet_tanshin:
        fd, note = from_edinet_tanshin(ticker_code)
        if fd:
            return fd, note, "edinet_tanshin"
        attempts.append(f"EDINET tanshin: {note}")
    else:
        attempts.append("EDINET tanshin: skipped (--tanshin-fallback not given; "
                        "EDINET does not host 決算短信)")

    return None, " | ".join(attempts), "none"


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print("usage: python scripts/guidance_fetcher.py <ticker> [<ticker> ...]")
        raise SystemExit(2)
    print(f"screener DB: {screener_db_path()}")
    for t in sys.argv[1:]:
        fd, note, tag = get_guidance(t)
        print(f"\n=== {t}  [{tag}]")
        print(f"  {note}")
        if fd:
            for k, v in sorted(fd.items()):
                print(f"    {k} = {v:,.1f} mn")
