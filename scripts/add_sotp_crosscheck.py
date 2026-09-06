# -*- coding: utf-8 -*-
"""
add_sotp_crosscheck.py - Adds a listed-stake SOTP floor sheet to a consolidated
DCF workbook.

A holding company for many listed and unlisted subsidiaries cannot be fully
described by a consolidated DCF; SOTP is the theoretically correct lens. When
only some of the subsidiaries are still listed, the sum of those stakes is a
FLOOR, not a valuation - and this sheet says so on its face, deliberately
refusing to price the unlisted units rather than guessing an EBITDA split.

Company-specific values (listed stakes, ownership %, parent net debt, NCI book
value, reference price, the notes) come from

    data/sotp/<ticker>_sotp_crosscheck.json

as required by CLAUDE.md ("銘柄コード入りスクリプトを作らない"). They used to be
compiled into this file, which meant running it on a second company valued that
company with the first one's subsidiaries - the same failure that retired
scripts/add_bank_valuation.py.

    python scripts/add_sotp_crosscheck.py models/<code>_DCF_Model_<date>.xlsx

The ticker is read from the workbook filename; pass --config PATH to override.

Note the difference from the other two SOTP tools:
  * this script  - a floor from LISTED STAKES, appended to an existing DCF workbook
  * scripts/generate_sotp.py - a full segment SOTP (EV/EBITDA per segment, and for
    型E a PBR-valued financial segment) as a standalone 6-sheet workbook
"""
import argparse
import io
import json
import os
import re
import sys

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

BLACK_FONT  = Font(name="Arial", size=10, color="000000")
BLUE_FONT   = Font(name="Arial", size=10, color="0000FF")
BOLD_FONT   = Font(name="Arial", size=10, bold=True)
TITLE_FONT  = Font(name="Arial", size=14, bold=True)
SUB_FONT    = Font(name="Arial", size=11, bold=True)
GREY_FONT   = Font(name="Arial", size=9, italic=True, color="808080")
HEADER_FONT = Font(name="Arial", size=11, bold=True, color="FFFFFF")
RED_FONT    = Font(name="Arial", size=10, bold=True, color="C00000")

HEADER_FILL   = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
LIGHT_GREEN   = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
LIGHT_YELLOW  = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
LIGHT_RED     = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

THIN_BORDER = Border(left=Side(style="thin"), right=Side(style="thin"),
                      top=Side(style="thin"), bottom=Side(style="thin"))
_GRAY_SIDE = Side(style="thin", color="B0B0B0")
INPUT_BORDER = Border(left=_GRAY_SIDE, right=_GRAY_SIDE, top=_GRAY_SIDE, bottom=_GRAY_SIDE)
TOP_BOTTOM = Border(top=Side(style="thin"), bottom=Side(style="double"))
SECTION_BOTTOM = Border(bottom=Side(style="thin"))

FMT_YEN   = '#,##0;(#,##0)'
FMT_PCT   = '0.0%;(0.0%)'
FMT_RATIO = '0.00"x"'
FMT_INT   = '#,##0'


def set_cell(ws, row, col, value, font=None, fmt=None, fill=None, border=None, alignment=None):
    c = ws.cell(row=row, column=col, value=value)
    if font: c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if border: c.border = border
    if alignment: c.alignment = alignment
    return c


def header_row(ws, row, col_start, labels):
    for i, lbl in enumerate(labels):
        c = ws.cell(row=row, column=col_start + i, value=lbl)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = SECTION_BOTTOM


def section_title(ws, row, col, text, font=SUB_FONT):
    c = ws.cell(row=row, column=col, value=text)
    c.font = font
    return c


def col_letter(n):
    return get_column_letter(n)


def load_config(xlsx_path, explicit=None):
    """Read data/sotp/<ticker>_sotp_crosscheck.json for the workbook's ticker.

    Error-stops when the file is missing. A silent fallback here would produce a
    sheet full of another company's subsidiaries, which is worse than no sheet.
    """
    if explicit:
        path = explicit
    else:
        base = os.path.basename(xlsx_path)
        m = re.match(r"([0-9A-Z]{4})_", base)
        if not m:
            raise SystemExit(
                f"ERROR: cannot read a ticker from {base!r}. Expected "
                f"<ticker>_DCF_Model_<date>.xlsx, or pass --config PATH.")
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(root, "data", "sotp", f"{m.group(1)}_sotp_crosscheck.json")
    if not os.path.exists(path):
        raise SystemExit(
            f"ERROR: {path} not found.\n"
            f"       銘柄固有の値(上場子会社の時価総額・持分比率、親会社純有利子負債、"
            f"NCI 簿価、参照株価)は data/sotp/ の設定ファイルから供給します。")
    cfg = json.load(io.open(path, encoding="utf-8"))
    missing = [k for k in ("listed_stakes", "parent_net_debt_mn",
                           "minority_interest_book_mn", "reference_price")
               if cfg.get(k) is None]
    if missing:
        raise SystemExit(f"ERROR: {path} is missing required key(s): {', '.join(missing)}")
    if not cfg["listed_stakes"]:
        raise SystemExit(f"ERROR: {path} lists no listed stakes - nothing to sum.")
    return cfg, path


def build_sotp_sheet(wb, cfg):
    ws = wb.create_sheet("SOTP Cross-Check")
    ws.sheet_properties.tabColor = "C00000"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 15

    company = cfg.get("company", "This company")
    set_cell(ws, 2, 2, "SOTP Cross-Check (intended PRIMARY framework -- see note)", font=TITLE_FONT)
    r = 3
    set_cell(ws, r, 2, cfg.get("intro_note", "").format(company=company), font=GREY_FONT)
    r += 2

    delisted = cfg.get("delisted_stakes", [])
    if delisted:
        set_cell(ws, r, 2, cfg.get("critical_finding", ""), font=BOLD_FONT, fill=LIGHT_RED)
        r += 1
        for d in delisted:
            set_cell(ws, r, 2, f"{d['name']} ({d['ticker']})", font=BOLD_FONT)
            set_cell(ws, r, 3, d.get("note", ""), font=GREY_FONT)
            r += 1
        r += 1
    set_cell(ws, r, 2, cfg.get("floor_note", ""), font=GREY_FONT)
    ws.row_dimensions[r].height = 45
    r += 3

    # ── Listed subsidiary stakes ──
    section_title(ws, r, 2, "Listed Subsidiary / Affiliate Stakes")
    r += 1
    header_row(ws, r, 2, ["Company", "Ticker", "Market Cap (JPY mn)", "Ownership %", "Implied Stake Value (JPY mn)"])
    hdr_row = r
    r += 1
    first_data_row = r
    for sub_ in cfg["listed_stakes"]:
        set_cell(ws, r, 2, sub_["name"], font=BLACK_FONT)
        set_cell(ws, r, 3, sub_["ticker"], font=BLACK_FONT)
        set_cell(ws, r, 4, sub_["market_cap_mn"], font=BLUE_FONT, fmt=FMT_YEN, border=INPUT_BORDER)
        set_cell(ws, r, 5, sub_["ownership"], font=BLUE_FONT, fmt=FMT_PCT, border=INPUT_BORDER)
        set_cell(ws, r, 6, f"=D{r}*E{r}", font=BLACK_FONT, fmt=FMT_YEN, border=THIN_BORDER)
        r += 1
    last_data_row = r - 1
    set_cell(ws, r, 2, "Sum of Listed Stakes (JPY mn)", font=BOLD_FONT, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    set_cell(ws, r, 6, f"=SUM(F{first_data_row}:F{last_data_row})", font=BOLD_FONT, fmt=FMT_YEN, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    sum_stakes_row = r
    r += 1
    set_cell(ws, r, 2, cfg.get("stakes_note", ""), font=GREY_FONT)
    r += 2

    set_cell(ws, r, 2, cfg.get("unpriced_label", "Non-listed / wholly-owned businesses"), font=BOLD_FONT)
    r += 1
    set_cell(ws, r, 2, "NOT PRICED -- no reliable per-segment EBITDA breakdown available; not guessed (per instruction).", font=GREY_FONT, fill=LIGHT_YELLOW)
    r += 2

    # ── Bridge to equity value ──
    section_title(ws, r, 2, "Bridge to SOTP Equity Value (FLOOR -- listed stakes only)")
    r += 1
    set_cell(ws, r, 2, "Sum of Listed Stakes (JPY mn)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=F{sum_stakes_row}", font=BLACK_FONT, fmt=FMT_YEN)
    bridge_stakes_row = r
    r += 1
    set_cell(ws, r, 2, "Less: Parent (unconsolidated) net interest-bearing debt", font=BOLD_FONT)
    set_cell(ws, r, 3, cfg["parent_net_debt_mn"], font=BLUE_FONT, fmt=FMT_YEN, border=INPUT_BORDER)
    parent_debt_row = r
    r += 1
    set_cell(ws, r, 2, "  Note: " + cfg.get("parent_net_debt_note", ""), font=GREY_FONT)
    ws.row_dimensions[r].height = 30
    r += 1
    set_cell(ws, r, 2, "SOTP Equity Value - FLOOR (JPY mn)", font=BOLD_FONT, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    set_cell(ws, r, 3, f"=C{bridge_stakes_row}-C{parent_debt_row}", font=BOLD_FONT, fmt=FMT_YEN, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    floor_equity_row = r
    r += 1
    set_cell(ws, r, 2, "Shares Outstanding", font=BOLD_FONT)
    set_cell(ws, r, 3, "='DCF Model'!C15", font=BLACK_FONT, fmt=FMT_INT)
    shares_row = r
    r += 1
    set_cell(ws, r, 2, "SOTP Floor Value per Share (JPY, pre-discount)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=ROUND(C{floor_equity_row}/C{shares_row}*1000000,0)", font=BOLD_FONT, fmt=FMT_YEN)
    floor_per_share_row = r
    r += 2
    set_cell(ws, r, 2,
             "This floor is likely NEGATIVE or small: it deducts group-level debt against only a small slice of group value "
             "(the 6 remaining minority-held listed stakes), deliberately excluding the wholly-owned businesses that generate "
             "most of consolidated EBITDA. A negative or very low floor does NOT mean the company is worth that little -- it means "
             "this floor construction cannot see most of the company. Treat as a data-availability limitation, not a valuation conclusion.",
             font=GREY_FONT)
    ws.row_dimensions[r].height = 45
    r += 3

    # ── Holding company discount sensitivity ──
    section_title(ws, r, 2, "Holding Company Discount Sensitivity (applied to floor equity value)")
    r += 1
    ref_price = cfg["reference_price"]
    header_row(ws, r, 2, ["Discount %", "SOTP Equity Value (JPY mn)", "Per Share (JPY)",
                          f"vs Current Price ({ref_price:,})"])
    disc_hdr_row = r
    r += 1
    discounts = cfg.get("discount_range", [0.0, 0.10, 0.20, 0.30, 0.40])
    disc_first_row = r
    for d in discounts:
        set_cell(ws, r, 2, d, font=BLUE_FONT, fmt=FMT_PCT, border=INPUT_BORDER)
        set_cell(ws, r, 3, f"=$C${floor_equity_row}*(1-B{r})", font=BLACK_FONT, fmt=FMT_YEN)
        set_cell(ws, r, 4, f"=ROUND(C{r}/$C${shares_row}*1000000,0)", font=BLACK_FONT, fmt=FMT_YEN)
        set_cell(ws, r, 5, f"=D{r}/{ref_price}-1", font=BLACK_FONT, fmt=FMT_PCT)
        r += 1
    r += 1
    set_cell(ws, r, 2,
             f"Compare each row against the current price of {ref_price:,}. Where no discount level brings the floor near "
             f"the market price, the floor is not a usable standalone valuation: it shows only that the remaining minority "
             f"listed stakes are a small fraction of group equity value, the balance being wholly-owned businesses this "
             f"sheet cannot independently price.",
             font=GREY_FONT)
    ws.row_dimensions[r].height = 45
    r += 3

    # ── MI comparison ──
    section_title(ws, r, 2, "Cross-check: Sum of Listed Stakes vs Non-Controlling Interests (book)")
    r += 1
    set_cell(ws, r, 2, "Sum of Listed Stakes (JPY mn, parent's share)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=F{sum_stakes_row}", font=BLACK_FONT, fmt=FMT_YEN)
    mi_compare_row1 = r
    r += 1
    set_cell(ws, r, 2, "Non-Controlling Interests, book value (JPY mn)", font=BOLD_FONT)
    set_cell(ws, r, 3, cfg["minority_interest_book_mn"], font=BLUE_FONT, fmt=FMT_YEN, border=INPUT_BORDER)
    mi_compare_row2 = r
    r += 1
    set_cell(ws, r, 2, "Difference (JPY mn)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=C{mi_compare_row1}-C{mi_compare_row2}", font=BLACK_FONT, fmt=FMT_YEN)
    diff_row = r
    r += 1
    set_cell(ws, r, 2, "Difference per share (JPY)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=ROUND(C{diff_row}/C{shares_row}*1000000,0)", font=BLACK_FONT, fmt=FMT_YEN)
    r += 2
    set_cell(ws, r, 2, cfg.get("mi_note", ""), font=GREY_FONT)
    ws.row_dimensions[r].height = 45

    ws.freeze_panes = "C6"
    return {"floor_per_share_row": floor_per_share_row}


def patch_comps_exclude_self(wb):
    if "Comps Analysis" not in wb.sheetnames:
        return False
    ws = wb["Comps Analysis"]
    last_comp_row = 4
    row = 5
    while ws.cell(row=row, column=2).value not in (None, ""):
        last_comp_row = row
        row += 1
    if last_comp_row <= 5:
        return False

    stat_rows = {"25th Percentile": None, "Median (50th)": None, "75th Percentile": None}
    for r in range(1, 40):
        label = ws.cell(row=r, column=2).value
        if label in stat_rows:
            stat_rows[label] = r

    stat_col_map = [(4, 10), (5, 11), (6, 12), (7, 13), (8, 14), (9, 15)]
    patched = 0
    for label, r in stat_rows.items():
        if r is None:
            continue
        for dst_col, src_col in stat_col_map:
            src_letter = col_letter(src_col)
            rng = f"{src_letter}6:{src_letter}{last_comp_row}"
            if label == "25th Percentile":
                formula = f"=PERCENTILE({rng},0.25)"
            elif label == "Median (50th)":
                formula = f"=MEDIAN({rng})"
            else:
                formula = f"=PERCENTILE({rng},0.75)"
            ws.cell(row=r, column=dst_col, value=formula)
            patched += 1
    print(f"  [add_sotp_crosscheck] Patched {patched} Comps Analysis stat formulas to exclude subject row 5 "
          f"(stats now over rows 6:{last_comp_row}).")
    return True


def main():
    ap = argparse.ArgumentParser(
        description="Append a listed-stake SOTP floor sheet to a DCF workbook.")
    ap.add_argument("xlsx", help="models/<ticker>_DCF_Model_<date>.xlsx")
    ap.add_argument("--config", help="path to <ticker>_sotp_crosscheck.json "
                                     "(default: data/sotp/<ticker>_sotp_crosscheck.json)")
    args = ap.parse_args()

    if not os.path.exists(args.xlsx):
        print(f"ERROR: file not found: {args.xlsx}")
        sys.exit(1)

    cfg, cfg_path = load_config(args.xlsx, args.config)
    print(f"  [add_sotp_crosscheck] inputs: {cfg_path}")
    print(f"  [add_sotp_crosscheck] {len(cfg['listed_stakes'])} listed stake(s), "
          f"reference price {cfg['reference_price']:,}")

    wb = openpyxl.load_workbook(args.xlsx)
    if "SOTP Cross-Check" in wb.sheetnames:
        del wb["SOTP Cross-Check"]

    build_sotp_sheet(wb, cfg)
    patch_comps_exclude_self(wb)

    wb.save(args.xlsx)
    print(f"Saved SOTP Cross-Check sheet into: {args.xlsx}")
    print("Run scripts/recalc_excel_com.py on this file next to compute formula values.")


if __name__ == "__main__":
    main()
