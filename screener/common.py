"""screener/common.py - paths, config, DB and a throttled HTTP session.

Everything in screener/ goes through here so that rate limiting, retries and the
"never fail silently" rule are implemented once.

Design rules inherited from the repo (docs/DCFパイプライン標準運用手順書.md §8):
  * no securities code in a module name — per-ticker/per-run values come from
    config files or CLI arguments;
  * a fetch that fails is RECORDED, never swallowed (see fetch_runs);
  * anything the code could not map is kept (unknown_tags), not dropped;
  * every filesystem path is derived from DATA_DIR, which DATA_ROOT in .env
    can point at another drive (see _resolve_data_root).
"""
from __future__ import annotations

import os
import subprocess
import json
import contextlib
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
CONFIG_DIR = os.path.join(PKG, "config")
DB_DIR = os.path.join(PKG, "db")
SCHEMA_PATH = os.path.join(DB_DIR, "schema.sql")
ENV_PATH = os.path.join(ROOT, ".env")

_ENV_LOADED = False
_DOTENV_MISSING = False


def load_env() -> None:
    """Load .env from the repository root. Explicit path: python-dotenv's
    find_dotenv() walks the caller's frame and blows up under `python -c`.

    Called at import time (below) because DATA_ROOT has to be resolved before
    any path constant is computed. load_dotenv does not overwrite variables
    already present in the process environment, so a shell override still wins
    over the file, and calling this again later is a no-op.
    """
    global _ENV_LOADED, _DOTENV_MISSING
    if _ENV_LOADED:
        return
    _ENV_LOADED = True
    try:
        from dotenv import load_dotenv
    except ImportError:                                # pragma: no cover
        _DOTENV_MISSING = True                         # surfaced by log() below
        return
    load_dotenv(ENV_PATH)


load_env()


def _resolve_data_root() -> str:
    """The single place that decides where the screener writes bytes.

    Everything the screener produces -- raw/, cache/, logs/, the SQLite DB --
    hangs off this one directory, so the whole tree can be moved to another
    drive by setting DATA_ROOT in the repo .env. That is not hypothetical: C:
    on this machine is down to ~5 GB while the full 3-year EDINET pull alone is
    estimated at ~12.6 GB (README, P1 現状), so the data lives on D:.

    Resolution order, all explicit -- no silent fallback to a half-full drive:
      SCREENER_DATA_ROOT   process env or .env; wins, for one-off overrides
      DATA_ROOT            the normal setting, kept in .env
      screener/data        the historical in-repo default

    Read via load_env() rather than the process environment, so the Task
    Scheduler job picks up the same root without the task definition carrying
    it. No other module may join a data path from PKG -- if a path is not
    derived from DATA_DIR it is a bug.
    """
    raw = (os.environ.get("SCREENER_DATA_ROOT")
           or os.environ.get("DATA_ROOT")
           or "").strip().strip('"').strip("'")
    if not raw:
        return os.path.join(PKG, "data")
    return os.path.abspath(os.path.expandvars(os.path.expanduser(raw)))


DATA_DIR = _resolve_data_root()
RAW_DIR = os.path.join(DATA_DIR, "raw")
CACHE_DIR = os.path.join(DATA_DIR, "cache")
LOG_DIR = os.path.join(DATA_DIR, "logs")
DB_PATH = os.environ.get("SCREENER_DB") or os.path.join(DATA_DIR, "screener.db")

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


_DOTENV_WARNED = False


def log(msg: str) -> None:
    global _DOTENV_WARNED
    if _DOTENV_MISSING and not _DOTENV_WARNED:
        # Deferred from import time: without dotenv, .env is never read, so a
        # DATA_ROOT set there is silently ignored and the run would fill C:.
        _DOTENV_WARNED = True
        log(f"WARNING: python-dotenv not installed - {ENV_PATH} was NOT read; "
            f"DATA_ROOT there has no effect (data root = {DATA_DIR})")
    line = f"[{utcnow()}] {msg}"
    print(line, flush=True)
    ensure_dirs()
    path = os.path.join(LOG_DIR, f"screener_{date.today():%Y%m}.log")
    with open(path, "a", encoding="utf-8") as fh:
        fh.write(line + "\n")


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


def store_path(abs_path: str) -> str:
    """Absolute path -> the form stored in filings.path / pdf_path / xbrl_path.

    Stored relative to DATA_DIR, deliberately not to the repo and never
    absolute. The archive has already moved once (C: ran out of room, it now
    lives on D:) and moves again with the next machine; a path anchored to the
    data root survives both without a DB rewrite. See migration
    db/migrations/001_paths_relative_to_data_root.py.
    """
    return os.path.relpath(abs_path, DATA_DIR)


def full_path(stored: str) -> str:
    """Inverse of store_path. An absolute value is returned untouched so a
    hand-inserted row is not silently mangled into a wrong relative path."""
    if not stored:
        return stored
    return stored if os.path.isabs(stored) else os.path.join(DATA_DIR, stored)


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
class WriterBusy(RuntimeError):
    """他のジョブが書き込みロックを持っている。"""


LOCK_PATH = os.path.join(DATA_DIR, "writer.lock")


def _pid_alive(pid: int) -> bool:
    if os.name != "nt":                                  # pragma: no cover
        try:
            os.kill(pid, 0); return True
        except OSError:
            return False
    out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"],
                         capture_output=True, text=True, timeout=20).stdout
    return str(pid) in out


@contextlib.contextmanager
def writer_lock(purpose: str, wait_seconds: float = 0.0,
                stale_after: float = 12 * 3600):
    """DBに長時間書くジョブ同士を直列化する助言ロック。

    なぜ要るか —— 2026-09-01、手動の全件再解析(2.5時間)が2回とも
    `database is locked` で落ちた。真因はウイルス対策でも WAL でもなく、
    **タスクスケジューラの日次ジョブ ScreenerTdnetArchiver が 19:00 に
    起動して同じ DB に書き始めたこと**。SQLite の busy_timeout(60秒)は
    「数十分書き続ける別プロセス」には効かない。競合を待つのではなく、
    そもそも同時に走らせないのが正しい。

    ロックは DATA_ROOT の writer.lock。中身は PID と用途と取得時刻。
    保持者が死んでいれば奪う（クラッシュしたジョブのロックで永久に
    止まるほうが困る）。stale_after を過ぎたロックも奪う。

    wait_seconds=0 なら即座に WriterBusy を上げる。日次ジョブのように
    「今回は諦めて次回に拾えばよい」側が使う。
    """
    deadline = time.time() + wait_seconds
    while True:
        holder = None
        try:
            with open(LOCK_PATH, encoding="utf-8") as fh:
                holder = json.load(fh)
        except (OSError, ValueError):
            holder = None
        if holder:
            age = time.time() - float(holder.get("at", 0))
            if age > stale_after or not _pid_alive(int(holder.get("pid", -1))):
                log(f"writer.lock: 死んだ保持者を掃除 {holder}")
                holder = None
        if holder is None:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(LOCK_PATH, "w", encoding="utf-8") as fh:
                json.dump({"pid": os.getpid(), "purpose": purpose,
                           "at": time.time()}, fh, ensure_ascii=False)
            break
        if time.time() >= deadline:
            raise WriterBusy(
                f"別のジョブが書き込み中: {holder.get('purpose')} "
                f"(pid={holder.get('pid')})。同時に走らせると "
                f"database is locked で両方が壊れる")
        time.sleep(5)
    try:
        yield
    finally:
        try:
            with open(LOCK_PATH, encoding="utf-8") as fh:
                cur = json.load(fh)
            if int(cur.get("pid", -1)) == os.getpid():
                os.remove(LOCK_PATH)
        except (OSError, ValueError):                    # pragma: no cover
            pass


def connect(db_path: str | None = None) -> sqlite3.Connection:
    ensure_dirs()
    con = sqlite3.connect(db_path or DB_PATH, timeout=60)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    # WAL: 読み手が書き手をブロックしない。取得(EDINET/株価)と解析と問い合わせが
    # 同時に走るので、既定の rollback journal では database is locked で落ちる
    # —— 実際に EDINET バックフィルが索引 400/1109 日目で落ちた(2026-08-31)。
    # busy_timeout は connect(timeout=) と同じ 60 秒を明示しておく。
    try:
        con.execute("PRAGMA journal_mode = WAL")
        con.execute("PRAGMA busy_timeout = 60000")
        con.execute("PRAGMA synchronous = NORMAL")
    except sqlite3.DatabaseError:                      # pragma: no cover
        pass                                           # WAL 不可な環境でも動かす
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

    2026-09-13: 'partial' を covered から外した。加えて tdnet_archiver は
    後条件ゲート（一覧の総件数 = 取得行数、対象書類がすべて保存済み、
    対象日が終わってから取得）を満たさない回を 'incomplete' / 'partial' /
    'provisional' と記録する。いずれも covered ではないので次の backfill が
    取り直す。朝に走った回が当日分の一部だけで 'ok' と記録され、backfill が
    「欠損なし」と言い続けた事故（9/02,03,04,10,11）の再発防止。
    """
    have = {
        r["target_date"]
        for r in con.execute(
            "SELECT DISTINCT target_date FROM fetch_runs "
            "WHERE source=? AND status IN ('ok','empty')", (source,)
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
    log(f"data root {DATA_DIR}")
    log(f"DB {DB_PATH}")
    log(f"tables: {', '.join(tables)}")
