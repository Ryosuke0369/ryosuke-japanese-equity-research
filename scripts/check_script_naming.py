"""check_script_naming.py - fail if a ticker code has leaked into a filename.

Why this exists: scripts/ started growing per-ticker copies of working code
(add_segment_bridge_3110.py, fill_adjustments_log_5726.py). Each copy forked the
logic, so a fix made in one never reached the others, and the next ticker got a
third copy. The rule is now: **one generic script, per-ticker data in a config
file** (data/overrides/, data/segments/, data/adjustments/, data/comps/).

Run standalone (exit 1 on a violation):

    python scripts/check_script_naming.py

generate_dcf.py also calls `check(warn_only=True)` at startup, so a new
violation is visible in the normal workflow rather than only in a check nobody
remembers to run.

The GRANDFATHERED set below is closed. Do not add to it — if you find yourself
wanting to, the answer is a config file.
"""
import os
import re
import sys

# A securities code in a filename: 4 digits (7203), or 3 digits + a letter
# (215A / 285A / 325A — the newer TSE format).
TICKER_RE = re.compile(r"(?:^|_)(\d{4}|\d{3}[A-Z])(?:_|\.|$)")

SCANNED_DIRS = ("scripts", "templates")

# Pre-2026-08-26 per-ticker runners. These are thin drivers (one ticker's inputs
# passed to a generic template), not forked logic, so they were left in place
# rather than rewritten. THIS LIST IS CLOSED — new files must not be added.
GRANDFATHERED = {
    "scripts/run_215A_dcf.py",
    "scripts/run_215A_market.py",
    "scripts/run_325A_market.py",
    "scripts/run_market_analysis_4192.py",
    "scripts/run_market_analysis_4192_v2.py",
    "scripts/run_market_analysis_5246.py",
    "scripts/run_market_analysis_6365.py",
    "scripts/run_market_analysis_6521.py",
    "scripts/run_market_analysis_7013.py",
    "scripts/run_narrative_4192.py",
}

ADVICE = {
    "segment": "data/segments/<ticker>_segments.json + scripts/add_segment_bridge.py",
    "adjustment": "data/adjustments/<ticker>_adjustments.json + scripts/fill_adjustments_log.py",
    "log": "data/adjustments/<ticker>_adjustments.json + scripts/fill_adjustments_log.py",
    "dcf": "data/overrides/<ticker>_overrides.json + scripts/generate_dcf.py",
    "comps": "data/comps/<ticker>_comps.csv",
}


def _advice_for(name):
    for word, how in ADVICE.items():
        if word in name.lower():
            return how
    return ("a config file under data/ consumed by the existing generic script "
            "(see docs/DCFフォーマット標準メモ §3-5)")


def find_violations(root=None):
    root = root or os.path.abspath(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    bad = []
    for d in SCANNED_DIRS:
        full = os.path.join(root, d)
        if not os.path.isdir(full):
            continue
        for name in sorted(os.listdir(full)):
            if not name.endswith(".py"):
                continue
            rel = "%s/%s" % (d, name)
            if rel in GRANDFATHERED:
                continue
            if TICKER_RE.search(os.path.splitext(name)[0]):
                bad.append(rel)
    return bad


def check(warn_only=False, root=None):
    bad = find_violations(root)
    if not bad:
        return True
    head = "WARNING" if warn_only else "ERROR"
    print("%s: ticker code(s) in script filename(s) — per-ticker logic must live "
          "in a config file, not a forked script:" % head)
    for rel in bad:
        print("  - %s  ->  %s" % (rel, _advice_for(os.path.basename(rel))))
    print("  See docs/DCFフォーマット標準メモ §3-5 / CLAUDE.md.")
    return False


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ok = check()
    print("OK: no ticker code in any script filename." if ok else "")
    sys.exit(0 if ok else 1)
