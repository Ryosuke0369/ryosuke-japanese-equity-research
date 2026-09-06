"""screener_link.py - read-only access to the screener's archive from the DCF pipeline.

The screener (`screener/`) archives TDnet and EDINET filings every weekday into
<DATA_ROOT>/screener.db. The DCF pipeline is a separate program, but it keeps
re-deriving things the archive already holds exactly: which 有報 docIDs a company
has filed, and what guidance its latest 決算短信 carried. Re-deriving them by
scanning EDINET date-by-date is slower and, as the 2026-09-05 batch showed,
occasionally wrong (フェーズ2 #7 / #9).

This module is the one place that knows how to find that database. Everything
here degrades to None rather than raising: the DCF pipeline must keep working on
a machine where the screener has never been set up.
"""
import os
import sqlite3
import sys

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
    """Path to screener.db, or None when it is not configured on this machine.

    Order: an explicit SCREENER_DB, the screener package's own resolution, then
    DATA_ROOT from the process environment or the repo .env.
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


def connect_ro():
    """Read-only connection to screener.db, or None. Never raises."""
    p = screener_db_path()
    if not p:
        return None
    try:
        return sqlite3.connect(f"file:{p}?mode=ro", uri=True, timeout=20)
    except sqlite3.Error:
        return None
