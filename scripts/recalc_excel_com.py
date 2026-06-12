"""Recalculate an Excel workbook in place via Excel COM (Windows).

openpyxl writes formulas but never computes them, so freshly generated DCF
models have no cached formula values. The market-analysis reverse DCF reads
computed cells (WACC, D&A/Capex/dNWC, PGM target); this script opens the
workbook in Excel, forces a full rebuild, and saves so those caches exist.
"""
import os
import sys
import win32com.client as win32

def recalc(path):
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(path)
        excel.CalculateFullRebuild()
        wb.Save()
        wb.Close(SaveChanges=True)
        print(f"Recalculated and saved: {path}")
    finally:
        excel.Quit()

if __name__ == "__main__":
    recalc(sys.argv[1])
