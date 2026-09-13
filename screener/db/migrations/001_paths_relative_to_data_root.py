r"""001 - re-anchor filings paths from the repo root to DATA_ROOT.

Until 2026-08-31 filings.path / pdf_path / xbrl_path were stored relative to the
repository root ("screener\data\raw\tdnet\20260826\x.pdf"). That was fine only
while the archive lived inside the repo. It no longer does: C: was down to ~5 GB
and the tree moved to D:\screener_data, and it will move again with the next
machine.

Paths are therefore now stored relative to DATA_DIR ("raw\tdnet\20260826\x.pdf"),
so moving the data root becomes a .env edit and nothing else. This migration
strips the old "screener/data/" prefix from rows written before the change.

Idempotent: rows already in the new form are left alone, so it is safe to re-run.
Absolute paths are reported and left untouched rather than guessed at.

    python -m screener.db.migrations.001_paths_relative_to_data_root [--apply]

Without --apply it only reports (dry run). --apply refuses to write if any
rewritten path does not resolve to a file, because a rewrite that points at
nothing would turn a working archive into a silently empty one.
"""
from __future__ import annotations

import argparse
import os
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))))))
    from screener import common as C

COLS = ("path", "pdf_path", "xbrl_path")
OLD_PREFIXES = ("screener\\data\\", "screener/data/")


def rewrite(value):
    """Old repo-relative form -> DATA_DIR-relative. None/new/absolute pass through."""
    if not value:
        return value
    for pre in OLD_PREFIXES:
        if value.startswith(pre):
            return value[len(pre):]
    return value


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="migration 001")
    ap.add_argument("--apply", action="store_true",
                    help="write the changes (default: dry run)")
    a = ap.parse_args(argv)

    con = C.connect()
    rows = con.execute("SELECT id, path, pdf_path, xbrl_path FROM filings").fetchall()
    C.log(f"migration 001: data root = {C.DATA_DIR}")
    C.log(f"  {len(rows)} filings rows")

    changed, absolute, missing = [], [], []
    for r in rows:
        new = {c: rewrite(r[c]) for c in COLS}
        for c in COLS:
            if r[c] and os.path.isabs(r[c]):
                absolute.append((r["id"], c, r[c]))
        if any(new[c] != r[c] for c in COLS):
            changed.append((r["id"], new))
        for c in COLS:                       # does the new path actually resolve?
            if new[c] and not os.path.exists(C.full_path(new[c])):
                missing.append((r["id"], c, new[c]))

    C.log(f"  rows needing rewrite            : {len(changed)}")
    C.log(f"  absolute paths (left untouched) : {len(absolute)}")
    C.log(f"  paths not on disk after rewrite : {len(missing)}")
    for i, (rid, col, v) in enumerate(missing):
        if i >= 10:
            C.log(f"    ... and {len(missing) - 10} more")
            break
        C.log(f"    ! filing {rid} {col}={v}")

    if not a.apply:
        C.log("  dry run - nothing written (pass --apply)")
        return 1 if missing else 0

    if missing:
        C.log("ERROR: refusing to apply while paths do not resolve. "
              "Check that DATA_ROOT points at the moved tree.")
        return 2

    for rid, new in changed:
        con.execute("UPDATE filings SET path=?, pdf_path=?, xbrl_path=? WHERE id=?",
                    (new["path"], new["pdf_path"], new["xbrl_path"], rid))
    con.commit()
    C.log(f"  applied: {len(changed)} rows rewritten")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
