"""screener/common.py - paths, config, DB and a throttled HTTP session.

Everything in screener/ goes through here so that rate limiting, retries and the
"never fail silently" rule are implemented once.

Design rules inherited from the repo (docs/DCFパイプライン標準運用手順書.md §8):
  * no securities code in a module name — per-ticker/per-run values come from
    config files or CLI arguments;
  * a fetch that fails is RECORDED, never swallowed (see fetch_runs);
  * anything the code could not map is kept (unknown_tags), not dropped.
"""
from __future__ import annotations

import os
import sqlite3
import sys
import time
from datetime import datetime, date, timedelta

import requests

try:                                                  # optional at import time
    import yaml
except ImportError:                                   # pragma: no cover
    yaml = None

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PKG = os.path.join(ROOT, "screener")
DATA_DIR = os.path.join(PKG, "data")
RAW_DIR = os.path.join(DATA_DIR, "raw")
CACHE_DIR = os.path.join(DATA_DIR, "cache")
CONFIG_DIR = os.path.join(PKG, "config")
DB_DIR = os.path.join(PKG, "db")
SCHEMA_PATH = os.path.join(DB_DIR, "schema.sql")
DB_PATH = os.environ.get("SCREENER_DB") or os.path.join(DATA_DIR, "screener.db")
LOG_DIR = os.path.join(DATA_DIR, "logs")

USER_AGENT = (
    "ryosuke-japanese-equity-research screener/1.0 "
    "(personal research archiver; contact via repository owner)"
)


# --------------------------------------------------------------------- utils
def utcnow() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_dirs() -> None:
    for d in (DATA_DIR, RAW_DIR, CACHE_DIR, LOG_DIR):
        os.makedirs(d, exist_ok=True)


def log(msg: str) -> None:
    line = f"[{utcnow()}] {msg}"
    print(line, flush=True)
    ensure_dirs()
    path = os.path.join(LOG_DIR, f"screener_{date.today():%Y%m}.log")
    with open(path, "a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def load_env() -> None:
    """Load .env from the repository root. Explicit path: python-dotenv's
    find_dotenv() walks the caller's frame and blows up under `python -c`."""
    try:
        from dotenv import load_dotenv
    except ImportError:                                # pragma: no cover
        log("WARNING: python-dotenv not installed - relying on process env only")
        return
    load_dotenv(os.path.join(ROOT, ".env"))


def require_env(name: str) -> str:
    """Return an env var or abort. A missing credential must stop the run, not
    quietly degrade it into a half-populated database."""
    load_env()
    v = os.environ.get(name)
    if not v:
        raise SystemExit(
            f"ERROR: {name} is not set. Put it in {os.path.join(ROOT, '.env')} "
            f"as `{name}=...` (the file is git-ignored)."
        )
    return v


def normalise_code(raw) -> str | None:
    """TDnet/EDINET give 5-character codes ('44770', '398A0'); the repo and
    J-Quants use the 4-character form ('4477', '398A'). Trailing '0' is the
    padding digit, so it is only stripped from a 5-character code."""
    if raw is None:
        return None
    s = str(raw).strip().upper()
    if not s:
        return None
    if len(s) == 5 and s.endswith("0"):
        return s[:4]
    return s


# ----------------------------------------------------------------- yaml/config
def load_yaml(name: str) -> dict:
    if yaml is None:                                   # pragma: no cover
        raise SystemExit("ERROR: pyyaml is required. pip install pyyaml")
    path = name if os.path.isabs(name) else os.path.join(CONFIG_DIR, name)
    if not os.path.isfile(path):
        raise SystemExit(f"ERROR: config not found: {path}")
    with open(path, encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


# ------------------------------------------------------------------------ db
def connect(db_path: str | None = None) -> sqlite3.Connection:
    ensure_dirs()
    con = sqlite3.connect(db_path or DB_PATH, timeout=60)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con


def init_db(db_path: str | None = None) -> sqlite3.Connection:
    con = connect(db_path)
    with open(SCHEMA_PATH, encoding="utf-8") as fh:
        con.executescript(fh.read())
    con.commit()
    return con


def start_run(con, source: str, target_date: str) -> int:
    """Open a fetch_runs row up-front so a crash still leaves evidence that the
    day was attempted. attempt increments across re-runs of the same day."""
    prev = con.execute(
        "SELECT MAX(attempt) AS a FROM fetch_runs WHERE source=? AND target_date=?",
        (source, target_date),
    ).fetchone()
    attempt = (prev["a"] or 0) + 1
    cur = con.execute(
        "INSERT INTO fetch_runs (source, target_date, status, started_at, attempt) "
        "VALUES (?,?,?,?,?)",
        (source, target_date, "running", utcnow(), attempt),
    )
    con.commit()
    return cur.lastrowid


def finish_run(con, run_id: int, status: str, **kw) -> None:
    cols = ("n_listed", "n_target", "n_saved", "n_failed", "error", "note")
    sets = ", ".join(f"{c}=?" for c in cols if c in kw)
    vals = [kw[c] for c in cols if c in kw]
    sql = "UPDATE fetch_runs SET status=?, finished_at=?"
    if sets:
        sql += ", " + sets
    sql += " WHERE id=?"
    con.execute(sql, [status, utcnow(), *vals, run_id])
    con.commit()


def missing_days(con, source: str, start: date, end: date,
                 weekdays_only: bool = True) -> list[str]:
    """Days in [start, end] with no successful fetch_runs row.

    This is the read side of the §2-1 requirement: a day is 'covered' only when
    a run finished ok/empty. A day whose only row is failed/running still shows
    up here, which is what makes the retry loop honest.
    """
    have = {
        r["target_date"]
        for r in con.execute(
            "SELECT DISTINCT target_date FROM fetch_runs "
            "WHERE source=? AND status IN ('ok','empty','partial')", (source,)
        )
    }
    out, d = [], start
    while d <= end:
        if not (weekdays_only and d.weekday() >= 5):
            s = d.isoformat()
            if s not in have:
                out.append(s)
        d += timedelta(days=1)
    return out


# ---------------------------------------------------------------------- http
class Throttle:
    """Minimum interval between requests to one host. The public archives we
    read (TDnet, EDINET, J-Quants) are courtesy resources; a screener that
    hammers them gets the whole project blocked."""

    def __init__(self, min_interval: float):
        self.min_interval = float(min_interval)
        self._last = 0.0

    def wait(self) -> None:
        gap = time.monotonic() - self._last
        if gap < self.min_interval:
            time.sleep(self.min_interval - gap)
        self._last = time.monotonic()


class Fetcher:
    """requests.Session + throttle + bounded retry with exponential backoff.

    Returns the Response for the caller to interpret. Raises only when every
    attempt failed, so callers can record the failure rather than guess.
    """

    def __init__(self, min_interval: float = 0.7, retries: int = 3,
                 timeout: int = 60, headers: dict | None = None):
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": USER_AGENT})
        if headers:
            self.s.headers.update(headers)
        self.throttle = Throttle(min_interval)
        self.retries = retries
        self.timeout = timeout
        self.n_requests = 0

    def get(self, url: str, *, params=None, expect_binary: bool = False,
            allow_status=(200,)):
        last = None
        for attempt in range(1, self.retries + 1):
            self.throttle.wait()
            try:
                r = self.s.get(url, params=params, timeout=self.timeout)
                self.n_requests += 1
                if r.status_code in allow_status:
                    return r
                # 4xx other than 429 will not fix itself by retrying
                if 400 <= r.status_code < 500 and r.status_code != 429:
                    return r
                last = f"HTTP {r.status_code}"
            except requests.RequestException as e:
                last = f"{type(e).__name__}: {e}"
            if attempt < self.retries:
                time.sleep(min(2 ** attempt, 20))
        raise RuntimeError(f"GET failed after {self.retries} attempts ({last}): {url}")

    def download(self, url: str, dest: str) -> int:
        """Write to a .part file then rename, so an interrupted run never leaves
        a truncated file that looks complete on the next pass."""
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        r = self.get(url, expect_binary=True)
        if r.status_code != 200:
            raise RuntimeError(f"HTTP {r.status_code} for {url}")
        tmp = dest + ".part"
        with open(tmp, "wb") as fh:
            fh.write(r.content)
        os.replace(tmp, dest)
        return len(r.content)


def business_days_back(n: int, end: date | None = None) -> list[date]:
    """The n most recent weekdays ending at `end` (inclusive if a weekday).

    Weekday-only, not a real exchange calendar: Japanese public holidays still
    produce a listing page, it is simply empty, and an empty day is recorded as
    'empty' rather than missing. Treating holidays as candidates is therefore
    safe and keeps the calendar out of the dependency list.
    """
    end = end or date.today()
    out, d = [], end
    while len(out) < n:
        if d.weekday() < 5:
            out.append(d)
        d -= timedelta(days=1)
    return list(reversed(out))


def parse_date_arg(s: str) -> date:
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise SystemExit(f"ERROR: cannot parse date {s!r} (use YYYYMMDD or YYYY-MM-DD)")


if __name__ == "__main__":                             # smoke test
    sys.stdout.reconfigure(encoding="utf-8")
    con = init_db()
    tables = [r[0] for r in con.execute(
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")]
    log(f"DB {DB_PATH}")
    log(f"tables: {', '.join(tables)}")
