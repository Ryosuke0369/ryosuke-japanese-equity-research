#!/bin/sh
# 追補6 §Z: regenerate with --force, then restore the batch's target filename.
#
# generate_dcf.py stamps the RUN date into the filename; the batch's name carries
# the ANALYSIS BASIS date. Keeping one file per ticker matters more than the run
# date, which the Pipeline Metadata block records inside the workbook anyway.
#
# フェーズ2 #3: the generator's exit status is now honoured. It used to be
# swallowed by the pipe into grep, so a run that generated NOTHING (exit 3:
# output file already present and --force not given) still fell through to the
# rename + validate block below, which then printed VERDICT: PASS for the STALE
# workbook. Exit codes: 0 ok / 1 validate FAIL / 2 bad invocation / 3 skipped.
t="$1"; shift
log="$(mktemp)"
python -u scripts/generate_dcf.py "$t" --force "$@" > "$log" 2>&1
gen_rc=$?
grep -E 'VERDICT|FAIL [0-9]|\[WARN\]|WACC:|Recalculated|ERROR|Traceback' "$log"
rm -f "$log"
if [ "$gen_rc" -ne 0 ]; then
  echo "  ERROR: ${t}: generate_dcf.py exited ${gen_rc} - NOT validating (no fresh workbook)."
  exit "$gen_rc"
fi

TARGET="${TARGET_DATE:-20260905}"
for d in 20260906 20260907 20260908; do
  [ "$d" = "$TARGET" ] && continue
  if [ -f "models/${t}_DCF_Model_${d}.xlsx" ]; then
    mv -f "models/${t}_DCF_Model_${d}.xlsx" "models/${t}_DCF_Model_${TARGET}.xlsx"
    rm -f "models/${t}_DCF_Model_${d}_validation.txt"
    python scripts/validate_output.py "models/${t}_DCF_Model_${TARGET}.xlsx" 2>&1 \
      | grep -E 'FAIL [0-9]|VERDICT'
    echo "  (renamed ${d} -> ${TARGET})"
  fi
done
