#!/bin/sh
# Regenerate one ticker with --force and validate the result.
#
# The date in the filename is the ANALYSIS BASIS date, not the wall-clock run
# date. It used to be stamped from datetime.now(), so a batch that ran past
# midnight produced a second file per ticker and left the first behind as a
# stale twin; 追補6 §Z worked around that here with a rename step driven by
# TARGET_DATE. フェーズ2 #4 made --date a first-class argument of
# generate_dcf.py, so the rename dance is gone and the generator writes the
# right filename in the first place.
#
# フェーズ2 #3: the generator's exit status is honoured. It used to be swallowed
# by a pipe into grep, so a run that generated NOTHING (exit 3: output file
# already present and --force not given) still fell through to validate — which
# then printed VERDICT: PASS for the STALE workbook.
# Exit codes: 0 ok / 1 validate FAIL / 2 bad invocation / 3 skipped.
t="$1"; shift
TARGET="${TARGET_DATE:-20260906}"
log="$(mktemp)"
python -u scripts/generate_dcf.py "$t" --force --date "$TARGET" "$@" > "$log" 2>&1
gen_rc=$?
grep -E 'VERDICT|FAIL [0-9]|\[WARN\]|WACC:|Recalculated|ERROR|Traceback' "$log"
rm -f "$log"
if [ "$gen_rc" -ne 0 ]; then
  echo "  ERROR: ${t}: generate_dcf.py exited ${gen_rc} - NOT validating (no fresh workbook)."
  exit "$gen_rc"
fi
