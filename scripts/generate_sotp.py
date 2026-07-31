#!/usr/bin/env python3
"""
SOTP Valuation Model Generator
Usage: python scripts/generate_sotp.py <ticker>
Example: python scripts/generate_sotp.py 7013

Reads the `sotp` section from data/overrides/<ticker>_overrides.json
and generates a 6-sheet Excel workbook in models/<ticker>_SOTP_Model.xlsx.

Requires: openpyxl
"""

import sys
import json
import os
import glob
import subprocess


def find_latest_dcf_model(project_root, ticker):
    """Find the latest DCF model file for a ticker in models/ and output/ dirs."""
    candidates = []
    for directory in ["models", "output"]:
        pattern = os.path.join(project_root, directory, f"{ticker}_DCF_Model_*.xlsx")
        candidates.extend(glob.glob(pattern))
    if not candidates:
        return None
    # Sort by filename (date suffix) descending to get latest
    candidates.sort(reverse=True)
    return candidates[0]


def read_dcf_crosscheck(dcf_path):
    """Read fair values from a DCF model's Executive Summary sheet.

    Uses win32com so formulas are evaluated even when the workbook was never
    recalculated (openpyxl alone would see empty cells).

    Rows are matched by their column-B LABELS, not by position: rows 16-19 are
    only the plain-DCF layout, and a bank-type model (8410: DDM / Residual
    Income) rebuilds its Valuation Summary — a positional read would import one
    method's number under another method's name. Falls back gracefully when
    Excel is unavailable.
    """
    result = {}
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(dcf_path), ReadOnly=True)
        wb.Application.CalculateFull()
        ws = wb.Sheets("Executive Summary")

        matchers = [
            ("exit_fair_value", lambda s: "exit" in s),
            ("pgm_fair_value", lambda s: "perpetuity" in s or "pgm" in s),
            ("comps_ev_ebitda", lambda s: "ev/ebitda" in s or "ev/sales" in s),
            ("comps_per", lambda s: "per" in s and "comps" in s),
        ]
        taken = set()
        for row in range(1, 61):
            label = ws.Cells(row, 2).Value
            if not isinstance(label, str) or not label.strip():
                continue
            s = label.strip().lower()
            for key, pred in matchers:
                if key in taken or not pred(s):
                    continue
                result[key] = ws.Cells(row, 3).Value
                taken.add(key)
                break
        wb.Close(SaveChanges=False)
        excel.Quit()
        missing = [k for k, _ in matchers if k not in taken]
        if missing:
            print(f"  WARNING: no Executive Summary row matched {', '.join(missing)}")
        # Convert float values to int for clean display
        for k, v in result.items():
            if isinstance(v, float):
                result[k] = int(round(v))
    except Exception as e:
        print(f"  WARNING: Could not read DCF cross-check values: {e}")
    return result


def main():
    if len(sys.argv) < 2:
        print("Usage: python scripts/generate_sotp.py <ticker>")
        print("Example: python scripts/generate_sotp.py 7013")
        sys.exit(1)

    ticker = sys.argv[1]
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    overrides_path = os.path.join(project_root, "data", "overrides", f"{ticker}_overrides.json")

    if not os.path.exists(overrides_path):
        print(f"Overrides file not found: {overrides_path}")
        sys.exit(1)

    with open(overrides_path, encoding="utf-8") as f:
        overrides = json.load(f)

    if "sotp" not in overrides:
        print(f"No 'sotp' section in {overrides_path}. Skipping SOTP generation.")
        sys.exit(0)

    sotp = overrides["sotp"]

    # Inject shares from single source of truth (overrides["shares"])
    if "shares" in overrides:
        fd_shares = overrides["shares"]["fully_diluted_shares"]
        fd_thousands = round(fd_shares / 1000)
        orig = sotp.get("consolidated", {}).get("shares_outstanding")
        if orig and abs(orig - fd_thousands) > 1:
            print(f"  WARNING: shares_outstanding corrected: {orig} -> {fd_thousands} (thousands)")
        sotp["consolidated"]["shares_outstanding"] = fd_thousands
        print(f"  Shares: {fd_thousands:,} thousand (from overrides.shares.fully_diluted_shares={fd_shares:,})")

    # Inject net_debt from top-level override if present
    if "net_debt" in overrides:
        sotp["consolidated"]["net_debt"] = overrides["net_debt"]
        print(f"  Net Debt: {overrides['net_debt']:,} (from overrides.net_debt)")

    # Read cross-check values from latest DCF model
    dcf_path = find_latest_dcf_model(project_root, ticker)
    if dcf_path:
        print(f"  DCF model found: {os.path.basename(dcf_path)}")
        xcheck = read_dcf_crosscheck(dcf_path)
        # Only inject non-None values
        xcheck_clean = {k: v for k, v in xcheck.items() if v is not None}
        if xcheck_clean:
            sotp["dcf_crosscheck"] = xcheck_clean
            print(f"  Cross-check values: {list(xcheck_clean.keys())}")
        else:
            print("  WARNING: DCF model has no cached values (formulas not yet calculated)")
    else:
        print(f"  No DCF model found for {ticker}, cross-check will be empty")

    output_path = os.path.join(project_root, "models", f"{ticker}_SOTP_Model.xlsx")

    # Import and run template engine
    sys.path.insert(0, os.path.join(project_root, "templates"))
    from sotp_template import generate_sotp_excel
    generate_sotp_excel(sotp, output_path, dcf_excel_path=dcf_path)

    # Run recalc.py for formula validation
    recalc_path = os.path.join(project_root, "scripts", "recalc.py")
    if os.path.exists(recalc_path):
        print(f"\nRunning formula validation...")
        result = subprocess.run(
            [sys.executable, recalc_path, output_path],
            capture_output=True, text=True
        )
        print(result.stdout)
        if result.stderr:
            print(result.stderr)
        if result.returncode != 0:
            print(f"WARNING: recalc.py reported errors (exit code {result.returncode})")
    else:
        print(f"recalc.py not found at {recalc_path}, skipping validation.")

    print(f"\nDone: {output_path}")


if __name__ == "__main__":
    main()
