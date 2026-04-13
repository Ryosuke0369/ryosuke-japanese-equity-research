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
import subprocess


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

    output_path = os.path.join(project_root, "models", f"{ticker}_SOTP_Model.xlsx")

    # Import and run template engine
    sys.path.insert(0, os.path.join(project_root, "templates"))
    from sotp_template import generate_sotp_excel
    generate_sotp_excel(sotp, output_path)

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
