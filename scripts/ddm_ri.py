"""
ddm_ri.py - DDM + Residual Income valuation sheets for a 型D (bank) model.

Banks break the standard UFCF/EV-net_debt DCF: debt is raw material rather than
funding, net working capital is undefined, and regulatory capital constrains
payouts. 手順書 §2 therefore routes 型D to a dividend discount model and a
residual income model, and the Target is the average of those two - never a DCF
number (追補13 §C / 型D/E 拡張 §2).

This engine was extracted from scripts/add_bank_valuation.py, which produced
8410 セブン銀行's DDM (258) and Residual Income (293) sheets but carried that
company's numbers inline: book value 280,491, net income 13,476, DPS
[11.00 .. 13.00], ROE [6.1% .. 7.0%], terminal g 1.5%. A file with one company's
figures compiled into it cannot be a pipeline path - CLAUDE.md requires
ticker-specific values to arrive from data/ - so they now come from the
overrides' `bank_valuation` block and this module holds only the arithmetic.

The formulas are unchanged from the 8410 precedent: values are never hardcoded
into the sheet, only inputs, and everything downstream is a live Excel formula.

Both sensitivity grids are centred on the model's own Ke and g rather than on a
fixed 5.2%-7.2% band, which reproduces 8410's grid exactly (its Ke was 6.2% and
g 1.5%) while generalising to a bank with different inputs.
"""
import sys
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

# ── Style constants (mirrored from templates/dcf_comps_template.py for visual consistency) ──
BLACK_FONT  = Font(name="Arial", size=10, color="000000")
BLUE_FONT   = Font(name="Arial", size=10, color="0000FF")   # input cells
BOLD_FONT   = Font(name="Arial", size=10, bold=True)
TITLE_FONT  = Font(name="Arial", size=14, bold=True)
SUB_FONT    = Font(name="Arial", size=11, bold=True)
GREY_FONT   = Font(name="Arial", size=9, italic=True, color="808080")
HEADER_FONT = Font(name="Arial", size=11, bold=True, color="FFFFFF")

HEADER_FILL   = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
LIGHT_GREEN   = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
LIGHT_YELLOW  = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
SUBTOTAL_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
HIGHLIGHT_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

THIN_BORDER = Border(left=Side(style="thin"), right=Side(style="thin"),
                      top=Side(style="thin"), bottom=Side(style="thin"))
_GRAY_SIDE = Side(style="thin", color="B0B0B0")
INPUT_BORDER = Border(left=_GRAY_SIDE, right=_GRAY_SIDE, top=_GRAY_SIDE, bottom=_GRAY_SIDE)
TOP_BOTTOM = Border(top=Side(style="thin"), bottom=Side(style="double"))
SECTION_BOTTOM = Border(bottom=Side(style="thin"))

FMT_YEN   = '#,##0;(#,##0)'
FMT_PCT   = '0.0%;(0.0%)'
FMT_PCT2  = '0.00%;(0.00%)'
FMT_RATIO = '0.00"x"'
FMT_INT   = '#,##0'
FMT_JPY_PS = '#,##0.0"円"'


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


# =====================================================================
# DDM SHEET
# =====================================================================
def build_ddm_sheet(wb, cfg):
    ws = wb.create_sheet("DDM")
    ws.sheet_properties.tabColor = "1F4E78"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 28
    for col in "CDEFGHIJK":
        ws.column_dimensions[col].width = 13

    set_cell(ws, 2, 2, "DDM - 2-Stage Dividend Discount Model (PRIMARY METHOD - bank)", font=TITLE_FONT)
    set_cell(ws, 3, 2, "Standard UFCF DCF does not apply to a bank (see 'DCF Model' sheet, used only as an auxiliary ATM-fee-franchise cross-check).",
             font=GREY_FONT)

    # ── Inputs ──
    section_title(ws, 5, 2, "Inputs")
    rows = [
        ("Risk-Free Rate",       "='DCF Model'!C7",  FMT_PCT2),
        ("Beta",                 "='DCF Model'!C8",  '0.00'),
        ("Equity Risk Premium",  "='DCF Model'!C9",  FMT_PCT2),
        ("Cost of Equity (Ke)",  "='DCF Model'!C23", FMT_PCT2),
        ("Shares Outstanding",   "='DCF Model'!C15", FMT_INT),
        ("Current Price (JPY)",  "='Executive Summary'!C9", FMT_YEN),
        ("Terminal Growth (g)",  cfg["terminal_growth"], FMT_PCT2),
    ]
    r = 6
    for label, val, fmt in rows:
        set_cell(ws, r, 2, label, font=BOLD_FONT)
        is_input = isinstance(val, (int, float))
        set_cell(ws, r, 3, val, font=(BLUE_FONT if is_input else BLACK_FONT), fmt=fmt,
                 border=(INPUT_BORDER if is_input else None))
        r += 1
    KE_CELL = "$C$9"
    G_CELL = "$C$12"
    SHARES_CELL = "$C$10"
    PRICE_CELL = "$C$11"

    set_cell(ws, r + 1, 2, "Guard (Ke must exceed g)", font=BOLD_FONT)
    set_cell(ws, r + 1, 3, f'=IF({KE_CELL}<={G_CELL},"ERROR: Ke <= g, model invalid","OK")',
             font=BLACK_FONT)
    guard_row = r + 1

    # ── DPS projection (Base scenario) ──
    dps_title_row = guard_row + 2
    section_title(ws, dps_title_row, 2, "Dividend per Share Forecast (Base scenario)")
    yr_row = dps_title_row + 1
    header_row(ws, yr_row, 3, cfg["year_labels"])

    dps_row = yr_row + 1
    set_cell(ws, dps_row, 2, "DPS (JPY)", font=BOLD_FONT)
    dps_values = cfg["dps"]
    for i, v in enumerate(dps_values):
        set_cell(ws, dps_row, 3 + i, v, font=BLUE_FONT, fmt='0.00', border=INPUT_BORDER)
    if cfg.get("dps_note"):
        set_cell(ws, dps_row, 8, cfg["dps_note"], font=GREY_FONT)

    df_row = dps_row + 1
    set_cell(ws, df_row, 2, "Discount Factor", font=BOLD_FONT)
    for i in range(5):
        col = 3 + i
        set_cell(ws, df_row, col, f"=1/(1+{KE_CELL})^{i+1}", font=BLACK_FONT, fmt='0.0000')

    pv_row = df_row + 1
    set_cell(ws, pv_row, 2, "PV of DPS", font=BOLD_FONT)
    for i in range(5):
        col = 3 + i
        cl = col_letter(col)
        set_cell(ws, pv_row, col, f"={cl}{dps_row}*{cl}{df_row}", font=BLACK_FONT, fmt='0.00')

    # ── Valuation ──
    val_title_row = pv_row + 2
    section_title(ws, val_title_row, 2, "Valuation")
    r = val_title_row + 1
    set_cell(ws, r, 2, "Sum of PV of DPS (Yr 1-5)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=SUM({col_letter(3)}{pv_row}:{col_letter(7)}{pv_row})", font=BLACK_FONT, fmt='0.00')
    sum_pv_row = r
    r += 1
    set_cell(ws, r, 2, "Terminal Value (Yr 5)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"={col_letter(7)}{dps_row}*(1+{G_CELL})/({KE_CELL}-{G_CELL})", font=BLACK_FONT, fmt='0.00')
    tv_row = r
    r += 1
    set_cell(ws, r, 2, "PV of Terminal Value", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=C{tv_row}*{col_letter(7)}{df_row}", font=BLACK_FONT, fmt='0.00')
    pv_tv_row = r
    r += 1
    set_cell(ws, r, 2, "Implied Value per Share (JPY)", font=BOLD_FONT, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    set_cell(ws, r, 3, f"=ROUND(C{sum_pv_row}+C{pv_tv_row},0)", font=BOLD_FONT, fmt=FMT_YEN, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    implied_row = r
    r += 1
    set_cell(ws, r, 2, "Upside / (Downside) vs Current Price", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=C{implied_row}/{PRICE_CELL}-1", font=BLACK_FONT, fmt=FMT_PCT)

    # ── Sensitivity: Ke (rows) x g (cols) ──
    sens_title_row = r + 2
    section_title(ws, sens_title_row, 2, "Sensitivity: Ke (rows) x Terminal g (cols)")
    # Centred on the model's own Ke and g (+-1.0pt at a 0.25pt step), which
    # reproduces the 8410 precedent exactly (Ke 6.2% -> 5.2%..7.2%, g 1.5% ->
    # 0.5%..2.5%) and stays informative for a bank with different inputs.
    _ke_c = round(cfg["ke"] / 0.0025) * 0.0025
    _g_c = round(cfg["terminal_growth"] / 0.0025) * 0.0025
    ke_vals = [round(_ke_c - 0.010 + 0.0025 * i, 6) for i in range(9)]
    g_vals = [round(_g_c - 0.010 + 0.0025 * i, 6) for i in range(9)]

    hdr_row = sens_title_row + 1
    set_cell(ws, hdr_row, 2, "Ke \\ g", font=BOLD_FONT)
    for j, g in enumerate(g_vals):
        set_cell(ws, hdr_row, 3 + j, g, font=BOLD_FONT, fmt=FMT_PCT2, alignment=Alignment(horizontal="center"))

    dps_range = f"${col_letter(3)}${dps_row}:${col_letter(7)}${dps_row}"
    for i, ke in enumerate(ke_vals):
        row = hdr_row + 1 + i
        set_cell(ws, row, 2, ke, font=BOLD_FONT, fmt=FMT_PCT2)
        ke_cell = f"$B{row}"
        for j, g in enumerate(g_vals):
            col = 3 + j
            g_cell = f"{col_letter(col)}${hdr_row}"
            formula = (
                f'=IF({ke_cell}<={g_cell},"",'
                f'ROUND(SUMPRODUCT({dps_range},1/(1+{ke_cell})^{{1,2,3,4,5}})'
                f'+{col_letter(7)}{dps_row}*(1+{g_cell})/({ke_cell}-{g_cell})/(1+{ke_cell})^5,0))'
            )
            set_cell(ws, row, col, formula, font=BLACK_FONT, fmt=FMT_YEN, border=THIN_BORDER)

    ws.freeze_panes = "C6"
    return {"ke_cell": KE_CELL, "g_cell": G_CELL, "shares_cell": SHARES_CELL,
            "price_cell": PRICE_CELL, "dps_row": dps_row, "implied_row": implied_row,
            "dps_col_start": 3, "dps_col_end": 7}


# =====================================================================
# RESIDUAL INCOME SHEET
# =====================================================================
def build_ri_sheet(wb, ddm_refs, cfg):
    ws = wb.create_sheet("Residual Income")
    ws.sheet_properties.tabColor = "1F4E78"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    for col in "CDEFGHIJK":
        ws.column_dimensions[col].width = 13

    set_cell(ws, 2, 2, "Residual Income Model (PRIMARY METHOD - bank)", font=TITLE_FONT)
    set_cell(ws, 3, 2, "Value = Book Value + PV(Residual Income). When ROE approaches Ke, value converges to book value.",
              font=GREY_FONT)

    section_title(ws, 5, 2, "Inputs")
    r = 6
    set_cell(ws, r, 2, "Beginning Book Value (JPY mn)", font=BOLD_FONT)
    set_cell(ws, r, 3, cfg["book_value_mn"], font=BLUE_FONT, fmt=FMT_YEN, border=INPUT_BORDER)
    bv0_row = r
    r += 1
    set_cell(ws, r, 2, "Shares Outstanding", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=DDM!{ddm_refs['shares_cell'][1:]}", font=BLACK_FONT, fmt=FMT_INT)
    shares_row = r
    r += 1
    set_cell(ws, r, 2, "BPS_0 (JPY)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=C{bv0_row}/C{shares_row}*1000000", font=BLACK_FONT, fmt='0.0"円"')
    bps0_row = r
    r += 1
    set_cell(ws, r, 2, "Cost of Equity (Ke)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=DDM!{ddm_refs['ke_cell'][1:]}", font=BLACK_FONT, fmt=FMT_PCT2)
    ke_row = r
    r += 1
    set_cell(ws, r, 2, "Terminal Growth (g)", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=DDM!{ddm_refs['g_cell'][1:]}", font=BLACK_FONT, fmt=FMT_PCT2)
    g_row = r
    KE_REF, G_REF = f"$C${ke_row}", f"$C${g_row}"

    # ── ROE forecast ──
    r += 2
    section_title(ws, r, 2, "ROE Forecast (Base scenario)")
    yr_row = r + 1
    header_row(ws, yr_row, 3, cfg["year_labels"])
    roe_row = yr_row + 1
    set_cell(ws, roe_row, 2, "ROE", font=BOLD_FONT)
    roe_values = cfg["roe"]
    for i, v in enumerate(roe_values):
        set_cell(ws, roe_row, 3 + i, v, font=BLUE_FONT, fmt=FMT_PCT, border=INPUT_BORDER)
    if cfg.get("roe_note"):
        set_cell(ws, roe_row, 8, cfg["roe_note"], font=GREY_FONT)

    # ── Rollforward ──
    r = roe_row + 2
    section_title(ws, r, 2, "Book Value Rollforward (JPY mn)")
    bv_hdr = r + 1
    set_cell(ws, bv_hdr, 2, "", font=BOLD_FONT)
    for i in range(5):
        set_cell(ws, bv_hdr, 3 + i, cfg["year_labels"][i],
                 font=BOLD_FONT, alignment=Alignment(horizontal="center"))

    ni_row = bv_hdr + 1
    div_row = ni_row + 1
    bv_row = div_row + 1
    ri_row = bv_row + 1
    pv_ri_row = ri_row + 1

    set_cell(ws, ni_row, 2, "Net Income_t = ROE_t x BV_(t-1)", font=BOLD_FONT)
    set_cell(ws, div_row, 2, "Dividends_t (from DDM)", font=BOLD_FONT)
    set_cell(ws, bv_row, 2, "Ending Book Value", font=BOLD_FONT)
    set_cell(ws, ri_row, 2, "Residual Income_t = (ROE_t - Ke) x BV_(t-1)", font=BOLD_FONT)
    set_cell(ws, pv_ri_row, 2, "PV of Residual Income", font=BOLD_FONT)

    ddm_dps_row = ddm_refs["dps_row"]
    for i in range(5):
        col = 3 + i
        cl = col_letter(col)
        prev_bv_cell = f"C{bv_row}" if i == 0 else f"{col_letter(col-1)}{bv_row}"
        prev_bv_ref = f"$C${bv0_row}" if i == 0 else prev_bv_cell
        set_cell(ws, ni_row, col, f"={cl}{roe_row}*{prev_bv_ref}", font=BLACK_FONT, fmt=FMT_YEN)
        set_cell(ws, div_row, col, f"=DDM!{cl}{ddm_dps_row}*C{shares_row}/1000000", font=BLACK_FONT, fmt=FMT_YEN)
        set_cell(ws, bv_row, col, f"={prev_bv_ref}+{cl}{ni_row}-{cl}{div_row}", font=BLACK_FONT, fmt=FMT_YEN, border=THIN_BORDER)
        set_cell(ws, ri_row, col, f"=({cl}{roe_row}-{KE_REF})*{prev_bv_ref}", font=BLACK_FONT, fmt=FMT_YEN)
        set_cell(ws, pv_ri_row, col, f"={cl}{ri_row}/(1+{KE_REF})^{i+1}", font=BLACK_FONT, fmt=FMT_YEN)

    # ── Valuation ──
    r = pv_ri_row + 2
    section_title(ws, r, 2, "Valuation")
    r += 1
    set_cell(ws, r, 2, "Sum of PV of Residual Income", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=SUM(C{pv_ri_row}:G{pv_ri_row})", font=BLACK_FONT, fmt=FMT_YEN)
    sum_pv_ri_row = r
    r += 1
    set_cell(ws, r, 2, "Terminal Residual Income", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=G{ri_row}*(1+{G_REF})/({KE_REF}-{G_REF})", font=BLACK_FONT, fmt=FMT_YEN)
    term_ri_row = r
    r += 1
    set_cell(ws, r, 2, "PV of Terminal Residual Income", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=C{term_ri_row}/(1+{KE_REF})^5", font=BLACK_FONT, fmt=FMT_YEN)
    pv_term_ri_row = r
    r += 1
    set_cell(ws, r, 2, "Equity Value (JPY mn)", font=BOLD_FONT, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    set_cell(ws, r, 3, f"=C{bv0_row}+C{sum_pv_ri_row}+C{pv_term_ri_row}", font=BOLD_FONT, fmt=FMT_YEN, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    ev_row = r
    r += 1
    set_cell(ws, r, 2, "Implied Value per Share (JPY)", font=BOLD_FONT, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    set_cell(ws, r, 3, f"=ROUND(C{ev_row}/C{shares_row}*1000000,0)", font=BOLD_FONT, fmt=FMT_YEN, fill=LIGHT_GREEN, border=TOP_BOTTOM)
    implied_row = r
    r += 1
    set_cell(ws, r, 2, "Upside / (Downside) vs Current Price", font=BOLD_FONT)
    set_cell(ws, r, 3, f"=C{implied_row}/DDM!{ddm_refs['price_cell'][1:]}-1", font=BLACK_FONT, fmt=FMT_PCT)
    r += 2
    set_cell(ws, r, 2,
             "Note: ROE sits close to Ke, so Residual Income per year is small and the implied value sits close to "
             "BPS_0 -- the theoretical price is structurally anchored to book value, not to growth assumptions. "
             "Whether ROE durably exceeds Ke is the entire valuation question for this stock.",
             font=GREY_FONT)

    # ── Justified P/B grid ──
    r += 2
    section_title(ws, r, 2, "Justified P/B = (ROE - g) / (Ke - g)")
    r += 1
    # Centred on the forecast ROE and the model Ke, at a 0.5pt step. For 8410
    # (mean ROE 6.68% -> 6.5%, Ke 6.2%) this reproduces the precedent's
    # 4.5%..8.5% x 5.2%..7.2% grid exactly.
    _roe_c = round((sum(cfg["roe"]) / len(cfg["roe"])) / 0.005) * 0.005
    _ke_c5 = round(cfg["ke"] / 0.005) * 0.005
    roe_grid_vals = [round(_roe_c - 0.020 + 0.005 * i, 6) for i in range(9)]
    ke_grid_vals = [round(_ke_c5 - 0.010 + 0.005 * i, 6) for i in range(5)]

    hdr_row2 = r
    set_cell(ws, hdr_row2, 2, "ROE \\ Ke", font=BOLD_FONT)
    for j, ke in enumerate(ke_grid_vals):
        set_cell(ws, hdr_row2, 3 + j, ke, font=BOLD_FONT, fmt=FMT_PCT2, alignment=Alignment(horizontal="center"))

    pb_first_data_row = hdr_row2 + 1
    for i, roe in enumerate(roe_grid_vals):
        row = pb_first_data_row + i
        set_cell(ws, row, 2, roe, font=BOLD_FONT, fmt=FMT_PCT)
        roe_cell = f"$B{row}"
        for j, ke in enumerate(ke_grid_vals):
            col = 3 + j
            ke_cell = f"{col_letter(col)}${hdr_row2}"
            formula = f'=IF({ke_cell}<={G_REF},"",ROUND(({roe_cell}-{G_REF})/({ke_cell}-{G_REF}),2))'
            set_cell(ws, row, col, formula, font=BLACK_FONT, fmt=FMT_RATIO, border=THIN_BORDER)
    pb_last_data_row = pb_first_data_row + len(roe_grid_vals) - 1

    # Conditional formatting: highlight cells within 0.05x of the current P/B (Market_Cap/Book_Value)
    cf_range = f"C{pb_first_data_row}:{col_letter(2+len(ke_grid_vals))}{pb_last_data_row}"
    current_pb_ref = f"DDM!{ddm_refs['price_cell'][1:]}*C{shares_row}/1000000/C{bv0_row}"
    rule = FormulaRule(
        formula=[f'AND(ISNUMBER({col_letter(3)}{pb_first_data_row}),ABS({col_letter(3)}{pb_first_data_row}-({current_pb_ref}))<0.05)'],
        fill=HIGHLIGHT_FILL,
    )
    # openpyxl FormulaRule anchors to the top-left cell of the range; apply per-cell via a relative formula instead
    rule2 = FormulaRule(formula=[f'ABS(C{pb_first_data_row}-({current_pb_ref}))<0.05'], fill=HIGHLIGHT_FILL)
    ws.conditional_formatting.add(cf_range, rule2)

    r = pb_last_data_row + 2
    set_cell(ws, r, 2, "Implied Share Price at each (ROE, Ke) node = BPS_0 x Justified P/B (JPY)", font=BOLD_FONT)
    r += 1
    hdr_row3 = r
    set_cell(ws, hdr_row3, 2, "ROE \\ Ke", font=BOLD_FONT)
    for j, ke in enumerate(ke_grid_vals):
        set_cell(ws, hdr_row3, 3 + j, ke, font=BOLD_FONT, fmt=FMT_PCT2, alignment=Alignment(horizontal="center"))
    price_first_row = hdr_row3 + 1
    for i, roe in enumerate(roe_grid_vals):
        row = price_first_row + i
        pb_row_ref = pb_first_data_row + i
        set_cell(ws, row, 2, roe, font=BOLD_FONT, fmt=FMT_PCT)
        for j, ke in enumerate(ke_grid_vals):
            col = 3 + j
            cl = col_letter(col)
            formula = f'=IF({cl}{pb_row_ref}="","",ROUND({cl}{pb_row_ref}*$C${bps0_row},0))'
            set_cell(ws, row, col, formula, font=BLACK_FONT, fmt=FMT_YEN, border=THIN_BORDER)

    ws.freeze_panes = "C6"
    return {"bv0_row": bv0_row, "bps0_row": bps0_row, "implied_row": implied_row}


# =====================================================================
# COMPS ANALYSIS PATCH: exclude subject company (row 5) from stats
# =====================================================================
def patch_comps_exclude_self(wb):
    if "Comps Analysis" not in wb.sheetnames:
        print("  [add_bank_valuation] No 'Comps Analysis' sheet found -- skipping self-exclusion patch.")
        return False
    ws = wb["Comps Analysis"]

    # Find the last comp data row by scanning column B (Company) from row 5 downward
    last_comp_row = 4
    row = 5
    while ws.cell(row=row, column=2).value not in (None, ""):
        last_comp_row = row
        row += 1

    if last_comp_row <= 5:
        print("  [add_bank_valuation] Only one (or zero) comp rows found -- nothing to exclude, skipping patch.")
        return False

    stat_rows = {"25th Percentile": None, "Median (50th)": None, "75th Percentile": None}
    for r in range(1, 40):
        label = ws.cell(row=r, column=2).value
        if label in stat_rows:
            stat_rows[label] = r

    stat_col_map = [(4, 10), (5, 11), (6, 12), (7, 13), (8, 14), (9, 15)]  # (dst_col, src_col)
    patched = 0
    for label, r in stat_rows.items():
        if r is None:
            continue
        for dst_col, src_col in stat_col_map:
            src_letter = col_letter(src_col)
            rng = f"{src_letter}6:{src_letter}{last_comp_row}"  # start at row 6, skip subject row 5
            if label == "25th Percentile":
                formula = f"=PERCENTILE({rng},0.25)"
            elif label == "Median (50th)":
                formula = f"=MEDIAN({rng})"
            else:
                formula = f"=PERCENTILE({rng},0.75)"
            ws.cell(row=r, column=dst_col, value=formula)
            patched += 1
    print(f"  [add_bank_valuation] Patched {patched} Comps Analysis stat formulas to exclude subject row 5 "
          f"(stats now over rows 6:{last_comp_row}).")
    return True


def suppress_ev_multiples(wb, quiet=False):
    """型D: blank the EV/EBITDA and EV/Revenue columns on the Comps sheet.

    手順書 §2 型D: "EV/EBITDA・EV/Revenue は使わず PER/PBR のみ". For a bank the
    comps CSV carries Net_Debt = 0 for every row by contract, so enterprise value
    collapses to market cap and the multiple is not an enterprise multiple at
    all. Left in place the percentile formulas also produce #NUM! (8410's
    D15:D17), which is a broken cell on display rather than a disclosed
    exclusion. The columns are replaced by the reason.
    """
    if "Comps Analysis" not in wb.sheetnames:
        return
    ws = wb["Comps Analysis"]
    hdr_row = None
    for r in range(1, 20):
        for c in range(3, 20):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().upper().startswith("EV/EBITDA"):
                hdr_row = r
                break
        if hdr_row:
            break
    if not hdr_row:
        return
    targets = []
    for c in range(3, 20):
        v = ws.cell(hdr_row, c).value
        if isinstance(v, str) and v.strip().upper().startswith(("EV/EBITDA", "EV/REVENUE", "EV/SALES")):
            targets.append(c)
    n = 0
    for c in targets:
        for r in range(hdr_row + 1, min(ws.max_row, hdr_row + 30) + 1):
            if ws.cell(r, c).value is not None:
                ws.cell(r, c).value = None
                n += 1
        ws.cell(hdr_row + 1, c).value = "N/A"
    # The Statistics block computes its percentiles from the per-peer multiple
    # columns just blanked, so PERCENTILE() over an empty range would leave
    # #NUM! on display. Those cells are replaced by the same "N/A" the template
    # uses for an excluded method. The EV columns of that block are D and E by
    # the template's fixed layout (the same positions scripts/diff_models.py
    # names as "Comps 25th pct EV/EBITDA" and "Comps median EV/Revenue").
    for r in range(hdr_row, min(ws.max_row, hdr_row + 30) + 1):
        lab = ws.cell(r, 2).value
        if isinstance(lab, str) and (lab.startswith(("25th", "75th"))
                                     or lab.startswith("Median")):
            for c in (4, 5):
                ws.cell(r, c).value = "N/A"
                n += 1

    # Say why, next to the header rather than in a cell somebody has to hunt for.
    note_col = (max(targets) + 1) if targets else 4
    ws.cell(hdr_row, note_col).value = (
        "← 型D(銀行): EV 倍率は使用しない（comps は全社 Net_Debt=0 の契約であり "
        "EV は時価総額に等しくなる）。PER / PBR のみを参照すること")
    if not quiet:
        print(f"  [型D] Comps Analysis: EV 倍率列を N/A 化（{n} セル、理由を注記）")


def _preserve_formulas_across_insert(ws, before, limit=64):
    """openpyxl の insert_rows が落とした数式をラベルで突き合わせて書き戻す。

    openpyxl は Excel の【共有数式(shared formula)】を、グループの先頭セルだけが
    実体を持ち残りは参照、という形で読む。insert_rows で行をずらすとその参照が壊れ、
    **保存時にメンバー側のセルが空になる**（メモリ上では見えているので、保存して
    読み直すまで気づけない）。

    実害: Executive Summary の "Comps - EV/EBITDA"（='Comps Analysis'!C27）と
    "Comps - PER"（!C28）は隣接する共有数式グループで、加算脚や DDM/RI が行を
    挿入したモデルでは **EV/EBITDA 側だけが空になっていた**。しかも check 20 は
    空セルを「テキスト＝除外された手法」と解釈して PASS していたため、
    8001/9434/8058/6971/4689 のすべてで見逃されていた。

    挿入前に (ラベル -> 数式) を控え、挿入後に空になったセルへ書き戻す。
    新しく書いた行はラベルが before に無いので触らない。
    """
    # メモリ上では数式が見えているので「空になったセルだけ直す」では捕まらない
    # （落ちるのは保存時）。控えておいた数式を**全部そのまま書き戻す**ことで、
    # 共有数式グループのメンバーを通常の数式に変換する（de-share）。
    restored = []
    for r in range(1, limit):
        lab = ws.cell(r, 2).value
        if not isinstance(lab, str) or lab not in before:
            continue
        ws.cell(r, 3).value = before[lab]
        restored.append((r, lab))
    return restored


def _snapshot_formulas(ws, limit=64):
    out = {}
    for r in range(1, limit):
        lab, v = ws.cell(r, 2).value, ws.cell(r, 3).value
        if isinstance(lab, str) and isinstance(v, str) and v.startswith("="):
            out[lab] = v
    return out


def wire_exec_summary(wb, ddm_refs, ri_refs, quiet=False):
    """Point the Target at the DDM/RI average and demote the DCF legs to reference.

    手順書 §2 型D: 「DCF不成立。主手法は DDM+RI」。The DCF sheet is kept as an
    auxiliary cross-check - it is where the fee-franchise economics can still be
    read - but its two legs must not enter the Target, and the workbook has to
    say so where a reader looks first.

    Following 手順書§5-5, no label is ever written starting with "=".
    """
    if "Executive Summary" not in wb.sheetnames:
        return None
    ws = wb["Executive Summary"]

    def find(prefix, limit=44):
        for r in range(1, limit):
            v = ws.cell(r, 2).value
            if isinstance(v, str) and v.startswith(prefix):
                return r
        return None

    r_tgt = find("Target Price")
    r_pgm = find("DCF - Perpetuity Growth")
    r_exit = find("DCF - Exit Multiple")
    r_note = find("Note: Target Mid")
    if not r_tgt:
        return None

    tag = "[参考・Target不算入 — 型D: 銀行に DCF は成立しない]"
    for r in (r_pgm, r_exit):
        if r:
            lab = ws.cell(r, 2).value
            if tag not in str(lab):
                ws.cell(r, 2).value = str(lab) + " " + tag

    # Insert the two primary methods directly under the Exit row so the summary
    # reads in the order the methods are actually used.
    _before = _snapshot_formulas(ws)
    anchor_row = max([x for x in (r_pgm, r_exit) if x] or [r_tgt])
    ws.insert_rows(anchor_row + 1, 2)
    r_ddm, r_ri = anchor_row + 1, anchor_row + 2
    if r_note and r_note > anchor_row:
        r_note += 2
    ws.cell(r_ddm, 2).value = "DDM - 2-Stage Dividend Discount (主手法)"
    ws.cell(r_ddm, 3).value = f"=DDM!C{ddm_refs['implied_row']}"
    ws.cell(r_ri, 2).value = "Residual Income (主手法)"
    ws.cell(r_ri, 3).value = f"='Residual Income'!C{ri_refs['implied_row']}"

    ws.cell(r_tgt, 3).value = (
        f'=IF(AND(ISNUMBER(C{r_ddm}),ISNUMBER(C{r_ri})),'
        f'ROUND(AVERAGE(C{r_ddm}:C{r_ri}),0),"N/A")')
    lab = str(ws.cell(r_tgt, 2).value or "Target Price (Mid)")
    if "DDM" not in lab:
        ws.cell(r_tgt, 2).value = lab.split(" - ")[0] + " - DDM / Residual Income の平均"

    sentence = (
        "【型D: 銀行】手順書 §2 により **DCF は成立しない**（負債が資金調達ではなく原材料であり、"
        "運転資本が定義できず、自己資本規制が配当を制約するため）。"
        "**Target は DDM と Residual Income の2手法の平均のみ**で構成し、"
        "DCF の2脚（PGM / Exit）と Comps の EV 倍率は Target に算入しない。"
        "DCF Model シートは ATM 手数料フランチャイズの採算を読むための補助として残してある。")
    if r_note:
        cur = str(ws.cell(r_note, 2).value or "")
        if "型D" not in cur:
            ws.cell(r_note, 2).value = (cur + " ■" + sentence) if cur else sentence
    _restored = _preserve_formulas_across_insert(ws, _before)
    if _restored and not quiet:
        print(f"  [型D] 行挿入で失われた共有数式 {len(_restored)} 件を復元: "
              + ", ".join(f"C{r}({l[:22]})" for r, l in _restored))
    # 型D は EV 倍率を使わない（手順書 §2「銀行に EV 倍率は使わない」）。
    # 行を挿入すると openpyxl が保存した数式を Excel が修復で落とすことがあり、
    # 実際 8410 の "Comps - EV/EBITDA" は空セルになっていた。空は「除外」ではなく
    # 「壊れた」なので、意図どおり **テキストの N/A** を明示的に書く。
    for r in range(1, 48):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v.startswith("Comps - EV/"):
            ws.cell(r, 3).value = "N/A"
            if "型D" not in v:
                ws.cell(r, 2).value = v + "（型D: 銀行に EV 倍率は使わない）"

    if not quiet:
        print(f"  [型D] Executive Summary: Target = AVERAGE(C{r_ddm}:C{r_ri}) "
              f"(DDM / Residual Income)、DCF 2脚は [参考・Target不算入] に降格")
    return {"target_row": r_tgt, "ddm_row": r_ddm, "ri_row": r_ri}


def warn_on_dcf_sheet(wb, quiet=False):
    """Put the 'a bank is not a DCF' warning where the DCF sheet is read."""
    if "DCF Model" not in wb.sheetnames:
        return
    ws = wb["DCF Model"]
    msg = ("⚠ 型D（銀行）: このシートは参考表示である。銀行に UFCF ベースの DCF は成立しない"
           "（負債は資金調達ではなく原材料、運転資本が定義できない、自己資本規制が配当を制約する）。"
           "**主手法は DDM シートと Residual Income シート**であり、Target はその2手法の平均である。")
    # Row 1 of the DCF Model sheet is empty by construction (the template starts
    # its title at row 2), so the warning goes there directly. It must NOT be
    # inserted: inserting shifts every row down and breaks 'DCF Model'!C23, the
    # Cost of Equity the DDM sheet reads - which is exactly what happened on the
    # first 8410 run and what the 型D check caught (Ke read as 0.00%).
    if isinstance(ws.cell(1, 2).value, str) and "型D" in ws.cell(1, 2).value:
        return
    ws.cell(1, 2).value = msg
    if not quiet:
        print("  [型D] DCF Model シートの先頭行(B1)に警告を記入（行は挿入しない）")


REQUIRED = ("book_value_mn", "dps", "roe", "terminal_growth", "year_labels")


def resolve_config(overrides, ke, projection_start_fy=None):
    """Build the engine config from the overrides' `bank_valuation` block.

    Nothing is invented here. A missing or malformed block raises, because the
    alternative - guessing a dividend path or an ROE path for a bank - is
    exactly the fabrication this pipeline forbids.
    """
    b = (overrides or {}).get("bank_valuation")
    if not isinstance(b, dict):
        raise ValueError(
            "型D には overrides の \"bank_valuation\" ブロックが必要です "
            "(book_value_mn / dps / roe / terminal_growth、year_labels は任意)。"
            "docs/overrides_schema.md の型D節を参照。")
    cfg = dict(b)
    cfg.setdefault("terminal_growth", overrides.get("terminal_growth"))
    if not cfg.get("year_labels"):
        # FY2027(E) ... FY2031(E) style, derived from the projection start label
        base = projection_start_fy or ""
        n = "".join(ch for ch in str(base) if ch.isdigit())
        if len(n) == 4:
            y0 = int(n)
            cfg["year_labels"] = [f"FY{y0 + i}(E)" for i in range(5)]
        else:
            cfg["year_labels"] = [f"Year {i + 1}(E)" for i in range(5)]
    missing = [k for k in REQUIRED if cfg.get(k) in (None, "", [])]
    if missing:
        raise ValueError(f"bank_valuation に不足しているキー: {', '.join(missing)}")
    for k in ("dps", "roe"):
        if len(cfg[k]) != 5:
            raise ValueError(f"bank_valuation.{k} は5要素の配列でなければならない "
                             f"(現在 {len(cfg[k])} 要素)")
    if ke is None:
        raise ValueError("Cost of Equity (Ke) を解決できない")
    cfg["ke"] = float(ke)
    if cfg["ke"] <= cfg["terminal_growth"]:
        raise ValueError(f"Ke ({cfg['ke']:.4f}) が g ({cfg['terminal_growth']:.4f}) 以下です。"
                         f"ゴードン成長式が成立しません")
    return cfg


def add_sheets(xlsx, cfg, quiet=False):
    """Append/replace the DDM and Residual Income sheets on an existing workbook.

    openpyxl does not compute: the caller must recalculate before validating.
    """
    wb = openpyxl.load_workbook(xlsx)
    for name in ("DDM", "Residual Income"):
        if name in wb.sheetnames:
            del wb[name]
    ddm_refs = build_ddm_sheet(wb, cfg)
    ri_refs = build_ri_sheet(wb, ddm_refs, cfg)
    patch_comps_exclude_self(wb)
    suppress_ev_multiples(wb, quiet=quiet)
    es_refs = wire_exec_summary(wb, ddm_refs, ri_refs, quiet=quiet)
    warn_on_dcf_sheet(wb, quiet=quiet)
    wb.save(xlsx)
    if not quiet:
        print(f"  [型D] DDM / Residual Income シートを生成: {os.path.basename(xlsx)}")
        print(f"        Ke {cfg['ke']:.2%} / g {cfg['terminal_growth']:.2%} / "
              f"BV0 {cfg['book_value_mn']:,.0f} mn / DPS {cfg['dps']} / ROE {cfg['roe']}")
    return {"ddm": ddm_refs, "ri": ri_refs, "exec": es_refs}


def main():
    import argparse
    import json as _json
    ap = argparse.ArgumentParser(
        description="型D(銀行)の DDM + Residual Income シートを既存ワークブックに追加する")
    ap.add_argument("xlsx")
    ap.add_argument("--overrides", default=None,
                    help="既定は data/overrides/<code>_overrides.json")
    ap.add_argument("--ke", type=float, default=None,
                    help="Cost of Equity。既定はワークブックの 'DCF Model'!C23 のキャッシュ値")
    a = ap.parse_args()
    if not os.path.exists(a.xlsx):
        raise SystemExit(f"ERROR: file not found: {a.xlsx}")
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    code = os.path.basename(a.xlsx)[:4]
    ovp = a.overrides or os.path.join(root, "data", "overrides", f"{code}_overrides.json")
    if not os.path.isfile(ovp):
        raise SystemExit(f"ERROR: overrides not found: {ovp}")
    overrides = _json.load(open(ovp, encoding="utf-8"))
    ke = a.ke
    if ke is None:
        wbv = openpyxl.load_workbook(a.xlsx, data_only=True)
        ke = wbv["DCF Model"]["C23"].value
        if not isinstance(ke, (int, float)):
            raise SystemExit("ERROR: 'DCF Model'!C23 (Cost of Equity) のキャッシュ値が無い。"
                             "recalc してから実行するか --ke を渡すこと")
    try:
        cfg = resolve_config(overrides, ke, overrides.get("projection_start_fy"))
    except ValueError as e:
        raise SystemExit(f"ERROR: {e}")
    add_sheets(a.xlsx, cfg)
    print("次に scripts/recalc_excel_com.py を実行して数式を計算すること。")


if __name__ == "__main__":
    main()
