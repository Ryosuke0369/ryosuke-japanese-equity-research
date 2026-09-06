"""
dcf_comps_build_v3.py - DCF / Comps Equity Research Excel Generator (V3)

Can be used standalone (with PDF extraction) or imported as a library:
  - Standalone: python templates/dcf_comps_template.py
  - Library: from dcf_comps_template import generate_dcf_workbook
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import subprocess, sys, os
import re
import datetime as _dt
import json as _json

# =====================================================================
# NARRATIVE TOKENS (investment_thesis / key_risks)
# =====================================================================
# A thesis that hardcodes "target 644 yen (+15%)" is wrong the moment the price
# or share count is refreshed. These tokens are rewritten into a live
# ="..."&TEXT(<cell>,"<fmt>")&"..." formula so the prose tracks the model.
# Keep in sync with docs/overrides_schema.md and overrides_validator.py.
NARRATIVE_TOKEN_FORMATS = {
    "price":        "#,##0",
    "target_price": "#,##0",
    "upside_pct":   "+0.0%;-0.0%",
    "pb":           '0.00"x"',
    "per":          '0.0"x"',
    "wacc":         "0.00%",
}
_TOKEN_RE = re.compile(r"\{([a-zA-Z_][a-zA-Z0-9_]*)\}")
# Excel's hard limit on a single formula is 8,192 characters.
_MAX_FORMULA_LEN = 8192


def _excel_str_literal(text):
    return '"' + text.replace('"', '""') + '"'


def _render_narrative_line(line, token_refs, warn):
    """Render one thesis/risk line into 1+ cell values.

    Lines without tokens are returned unchanged (plain text — the pre-token
    behaviour, so existing overrides are untouched). Lines with tokens become
    string-concatenation formulas; a line whose formula would exceed Excel's
    8,192-character limit is split across consecutive cells.
    """
    if not isinstance(line, str) or not _TOKEN_RE.search(line):
        return [line]

    def build(fragment):
        pieces, last = [], 0
        for m in _TOKEN_RE.finditer(fragment):
            name = m.group(1)
            ref = token_refs.get(name)
            if ref is None:
                # Unknown tokens are rejected by the validator; a known token
                # with no reference (e.g. {pb} with no comps) degrades to text.
                warn(name)
                continue
            if m.start() > last:
                pieces.append(_excel_str_literal(fragment[last:m.start()]))
            fmt = NARRATIVE_TOKEN_FORMATS[name]
            pieces.append(f'TEXT({ref},{_excel_str_literal(fmt)})')
            last = m.end()
        if last == 0:
            return None  # no token resolved -> keep as plain text
        if last < len(fragment):
            pieces.append(_excel_str_literal(fragment[last:]))
        return "=" + "&".join(pieces)

    formula = build(line)
    if formula is None:
        return [line]
    if len(formula) <= _MAX_FORMULA_LEN:
        return [formula]

    # Too long for one cell: split on sentence boundaries and lay the halves out
    # in consecutive cells rather than truncating or dropping the tokens.
    chunks, buf = [], ""
    for sentence in re.split(r"(?<=[。.!?])\s*", line):
        candidate = buf + sentence
        if buf and len(build(candidate) or candidate) > _MAX_FORMULA_LEN - 200:
            chunks.append(buf)
            buf = sentence
        else:
            buf = candidate
    if buf:
        chunks.append(buf)
    return [build(ch) or ch for ch in chunks]

try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False

# =====================================================================
# V3: ROW NUMBERS — Full waterfall, no SGA_OFFSET toggle
# =====================================================================
R_DRV_GROWTH   = 30  # driver row: Revenue Growth (YoY)
R_DRV_COGS     = 31  # driver row: COGS % of Revenue
R_DRV_SGA      = 32  # driver row: SGA Expense
R_REVENUE      = 33
R_COGS         = 34
R_GROSS_PROFIT = 35
R_GROSS_MARGIN = 36
R_SGA          = 37
R_OP_M_IMPL   = 38
R_EBIT         = 39
R_TAX          = 40
R_NOPAT        = 41
R_DA           = 42
R_CAPEX        = 43
R_CHG_NWC      = 44  # Change in NWC (linked from NWC Schedule)
R_UFCF         = 45
R_DISC         = 46
R_PV_FCF       = 47
# Stub period assumption rows
R_STUB_FRACTION = 19
R_LTM_REVENUE   = 20

# PGM section: R_PV_FCF + 2 gap
R_PGM_SEC    = R_PV_FCF + 2
R_SUM_PV     = R_PGM_SEC + 1
R_TV_PGM     = R_SUM_PV + 1
R_PV_TV_PGM  = R_TV_PGM + 1
R_EV_PGM     = R_PV_TV_PGM + 1
R_EQ_PGM     = R_EV_PGM + 1
R_PRICE_PGM  = R_EQ_PGM + 1
# Exit section: R_PRICE_PGM + 2 gap
R_EXIT_SEC   = R_PRICE_PGM + 2
R_SUM_PV_EX  = R_EXIT_SEC + 1
R_YR5_EBITDA = R_SUM_PV_EX + 1
R_TV_EXIT    = R_YR5_EBITDA + 1
R_PV_TV_EXIT = R_TV_EXIT + 1
R_EV_EXIT    = R_PV_TV_EXIT + 1
R_EQ_EXIT    = R_EV_EXIT + 1
R_PRICE_EXIT = R_EQ_EXIT + 1

# Scenario Matrix Section (below Exit valuation)
SCENARIO_NAMES = ["Base", "Upside", "Management", "Downside 1", "Downside 2"]
NUM_SCENARIOS  = 5

R_SCEN_SEC        = R_PRICE_EXIT + 2       # Section header
R_SCEN_YEARS      = R_SCEN_SEC + 1         # Year 1-5 column headers

# Each block: sub-header 1 row + 5 scenario rows + 1 blank = 7 rows
R_SCEN_BLK_GROWTH = R_SCEN_YEARS + 1
R_SCEN_BLK_COGS   = R_SCEN_BLK_GROWTH + 7
R_SCEN_BLK_SGA    = R_SCEN_BLK_COGS + 7

# ── NWC Schedule Row Numbers ──
NWC_R_DSO      = 5
NWC_R_DIH      = 6
NWC_R_DPO      = 7
NWC_R_REV      = 9
NWC_R_COGS     = 10
NWC_R_AR       = 12
NWC_R_INV      = 13
NWC_R_CA       = 14
NWC_R_AP       = 15
NWC_R_CL       = 16
NWC_R_NWC      = 18
NWC_R_CHG_NWC  = 19

NWC_R_SCEN_SEC      = 22
NWC_R_SCEN_YEARS    = 23
NWC_R_SCEN_BLK_DSO  = 24
NWC_R_SCEN_BLK_DIH  = NWC_R_SCEN_BLK_DSO + 7
NWC_R_SCEN_BLK_DPO  = NWC_R_SCEN_BLK_DIH + 7

# =====================================================================
# STYLE CONSTANTS
# =====================================================================
BLUE_FONT   = Font(name="Arial", size=10, color="000000", bold=False)  # unified to black
BLACK_FONT  = Font(name="Arial", size=10, color="000000")
GREEN_FONT  = Font(name="Arial", size=10, color="006600")
BOLD_FONT   = Font(name="Arial", size=10, bold=True)
HEADER_FONT = Font(name="Arial", size=11, bold=True, color="FFFFFF")
TITLE_FONT  = Font(name="Arial", size=14, bold=True)
SUB_FONT    = Font(name="Arial", size=11, bold=True)
GREY_FONT   = Font(name="Arial", size=9, italic=True, color="808080")

HEADER_FILL    = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
LIGHT_FILL     = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
LIGHT_GREEN    = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
LIGHT_YELLOW   = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
SUBTOTAL_FILL  = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

# Borders
THIN_BORDER     = Border(left=Side(style="thin"), right=Side(style="thin"),
                          top=Side(style="thin"), bottom=Side(style="thin"))
_GRAY_SIDE      = Side(style="thin", color="B0B0B0")
LIGHT_GRAY_BORDER = Border(bottom=Side(style="dotted", color="B0B0B0"))
_GRAY_HAIR       = Side(style="hair", color="B0B0B0")
NWC_DATA_BORDER  = Border(left=_GRAY_HAIR, right=_GRAY_HAIR,
                           top=_GRAY_HAIR, bottom=_GRAY_HAIR)
SECTION_BOTTOM  = Border(bottom=Side(style="thin"))
SUBTOTAL_BORDER = Border(top=Side(style="thin"), bottom=Side(style="thin"))
TOP_BOTTOM      = Border(top=Side(style="thin"), bottom=Side(style="double"))
INPUT_BORDER    = Border(left=_GRAY_SIDE, right=_GRAY_SIDE,
                          top=_GRAY_SIDE, bottom=_GRAY_SIDE)

FMT_YEN     = '#,##0;(#,##0)'
FMT_YEN_DEC = '#,##0.0;(#,##0.0)'
FMT_PCT     = '0.0%;(0.0%)'
FMT_PCT2    = '0.00%;(0.00%)'
FMT_RATIO   = '0.00"x"'
FMT_INT     = '#,##0'
FMT_EPS     = '#,##0.0;(#,##0.0)'
FMT_DAYS    = '#,##0'

# =====================================================================
# HELPER FUNCTIONS
# =====================================================================
def set_cell(ws, row, col, value, font=None, fmt=None, fill=None, border=None, alignment=None):
    c = ws.cell(row=row, column=col, value=value)
    if font:      c.font = font
    if fmt:       c.number_format = fmt
    if fill:      c.fill = fill
    if border:    c.border = border
    if alignment: c.alignment = alignment
    return c

def header_row(ws, row, col_start, col_end, labels, fill=HEADER_FILL, font=HEADER_FONT):
    for i, lbl in enumerate(labels):
        c = ws.cell(row=row, column=col_start + i, value=lbl)
        c.font = font
        c.fill = fill
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = SECTION_BOTTOM

def section_title(ws, row, col, text, font=SUB_FONT):
    c = ws.cell(row=row, column=col, value=text)
    c.font = font
    return c

def col_letter(col_num):
    return get_column_letter(col_num)

def choose_formula(block_start, cl):
    """Generate CHOOSE formula referencing 5 scenario rows in the matrix."""
    refs = [f"{cl}{block_start + 1 + s}" for s in range(NUM_SCENARIOS)]
    return f"=CHOOSE($D$27,{','.join(refs)})"

def nwc_choose_formula(block_start, cl):
    """Generate CHOOSE formula for NWC Schedule referencing DCF Model scenario index."""
    refs = [f"{cl}{block_start + 1 + s}" for s in range(NUM_SCENARIOS)]
    return f"=CHOOSE('DCF Model'!$D$27,{','.join(refs)})"

def seg_choose_formula(scenario_rows, cl):
    """Generate CHOOSE formula for Segment Analysis referencing DCF Model scenario index.

    Args:
        scenario_rows: list of 5 row numbers (one per scenario: Base, Upside, Mgmt, DS1, DS2)
        cl: column letter
    Returns:
        Excel CHOOSE formula string
    """
    refs = [f"{cl}{r}" for r in scenario_rows]
    return f"=CHOOSE('DCF Model'!$D$27,{','.join(refs)})"

# =====================================================================
# SENSITIVITY ANALYSIS HELPERS (V3: full waterfall)
# =====================================================================
def calc_dcf_pgm(rev_growth, gross_margin, wacc, tg, cfg):
    """Calculate implied share price using Perpetuity Growth Method.
    V3: Uses gross_margin (GM%) and absolute SGA to compute EBIT.
    """
    n = cfg["projection_years"]
    capex_pct = cfg["capex_pct"]
    da_pct = cfg["da_pct"]
    tax = cfg["tax_rate"]
    net_debt = cfg["net_debt"]
    shares = cfg["shares_outstanding"]
    sga_pct_list = cfg["sga_pct"]
    _capex_method = cfg.get("capex_method", "revenue_pct")
    _da_method = cfg.get("da_method", "revenue_pct")
    _capex_direct = cfg.get("capex_direct", {}).get("projections", []) if _capex_method == "direct" else []
    _da_direct = cfg.get("da_direct", {}).get("projections", []) if _da_method == "direct" else []

    revenues = []
    rev = cfg["base_year_revenue"]
    for _ in range(n):
        rev = rev * (1 + rev_growth)
        revenues.append(rev)

    sum_pv_fcf = 0
    last_fcf = 0
    for yr_idx, rev in enumerate(revenues):
        cogs = rev * (1 - gross_margin)
        gp = rev - cogs
        sga = rev * sga_pct_list[yr_idx]
        ebit = gp - sga
        nopat = ebit * (1 - tax) if ebit > 0 else ebit
        if _da_method == "direct" and yr_idx < len(_da_direct) and _da_direct[yr_idx] is not None:
            da = _da_direct[yr_idx]
        else:
            da = rev * da_pct
        if _capex_method == "direct" and yr_idx < len(_capex_direct) and _capex_direct[yr_idx] is not None:
            capex = _capex_direct[yr_idx]
        else:
            capex = rev * capex_pct
        fcf = nopat + da - capex
        df = 1 / (1 + wacc) ** (yr_idx + 1)
        sum_pv_fcf += fcf * df
        last_fcf = fcf

    tv = last_fcf * (1 + tg) / (wacc - tg)
    pv_tv = tv / (1 + wacc) ** n
    ev = sum_pv_fcf + pv_tv
    equity = ev - net_debt
    price = round(equity * 1_000_000 / shares)
    return price

def calc_dcf_exit(rev_growth, gross_margin, wacc, exit_mult, cfg):
    """Calculate implied share price using Exit Multiple Method."""
    n = cfg["projection_years"]
    capex_pct = cfg["capex_pct"]
    da_pct = cfg["da_pct"]
    tax = cfg["tax_rate"]
    net_debt = cfg["net_debt"]
    shares = cfg["shares_outstanding"]
    sga_pct_list = cfg["sga_pct"]
    _capex_method = cfg.get("capex_method", "revenue_pct")
    _da_method = cfg.get("da_method", "revenue_pct")
    _capex_direct = cfg.get("capex_direct", {}).get("projections", []) if _capex_method == "direct" else []
    _da_direct = cfg.get("da_direct", {}).get("projections", []) if _da_method == "direct" else []

    revenues = []
    rev = cfg["base_year_revenue"]
    for _ in range(n):
        rev = rev * (1 + rev_growth)
        revenues.append(rev)

    sum_pv_fcf = 0
    last_ebit = 0
    for yr_idx, rev in enumerate(revenues):
        cogs = rev * (1 - gross_margin)
        gp = rev - cogs
        sga = rev * sga_pct_list[yr_idx]
        ebit = gp - sga
        nopat = ebit * (1 - tax) if ebit > 0 else ebit
        if _da_method == "direct" and yr_idx < len(_da_direct) and _da_direct[yr_idx] is not None:
            da = _da_direct[yr_idx]
        else:
            da = rev * da_pct
        if _capex_method == "direct" and yr_idx < len(_capex_direct) and _capex_direct[yr_idx] is not None:
            capex = _capex_direct[yr_idx]
        else:
            capex = rev * capex_pct
        fcf = nopat + da - capex
        df = 1 / (1 + wacc) ** (yr_idx + 1)
        sum_pv_fcf += fcf * df
        last_ebit = ebit

    if _da_method == "direct" and len(_da_direct) >= n and _da_direct[n - 1] is not None:
        yr5_da = _da_direct[n - 1]
    else:
        yr5_da = revenues[-1] * da_pct
    yr5_ebitda = last_ebit + yr5_da
    tv = yr5_ebitda * exit_mult
    pv_tv = tv / (1 + wacc) ** n
    ev = sum_pv_fcf + pv_tv
    equity = ev - net_debt
    price = round(equity * 1_000_000 / shares)
    return price

# =====================================================================
# Derived values for WACC
# =====================================================================
def calc_wacc(cfg):
    ke = cfg["risk_free"] + cfg["beta"] * cfg["erp"] + cfg["size_premium"]
    we = 1 / (1 + cfg["de_ratio"])
    wd = cfg["de_ratio"] / (1 + cfg["de_ratio"])
    return ke * we + cfg["cost_of_debt_at"] * wd

# =====================================================================
# BUILD WORKBOOK
# =====================================================================

# V2: DYNAMIC STOCK DATA FETCHING
# =====================================================================
def get_live_market_data(ticker_str, fallback_price, fallback_shares):
    if not YFINANCE_AVAILABLE:
        print("yfinance not installed. Using fallback market data.")
        return fallback_price, fallback_shares, 1.0

    try:
        print(f"Fetching live data for {ticker_str} via yfinance...")
        tkr = yf.Ticker(ticker_str)
        info = tkr.info
        live_price = info.get("currentPrice") or info.get("regularMarketPrice") or fallback_price
        live_shares = info.get("sharesOutstanding") or fallback_shares
        raw_beta = info.get("beta")
        if raw_beta and 0.6 <= raw_beta <= 1.5:
            live_beta = raw_beta
        else:
            live_beta = 1.0  # sector-standard fallback
            print(f"  Beta {raw_beta} outside [0.6, 1.5] range - using fallback 1.0")
        print(f"Successfully fetched: Price={live_price}, Shares={live_shares}, Beta={live_beta}")
        return float(live_price), int(live_shares), float(live_beta)
    except Exception as e:
        print(f"Warning: Failed to fetch live data ({str(e).encode('ascii', 'replace').decode()}). Using fallback market data.")
        return fallback_price, fallback_shares, 1.0


# =====================================================================
# SEGMENT ANALYSIS SHEET
# =====================================================================
def _create_segment_sheet(wb, C, segments, proj_years, year_labels):
    """Generate Segment Analysis sheet from overrides JSON segments data.

    v2: Revenue uses YoY growth rates (not absolute values).
    Revenue display = Base Year × (1+growth) chain via CHOOSE formulas.
    Includes Consolidated Inputs section (SGA%, NWC%) after segment blocks.

    Returns:
        dict with total_rev_row, total_op_row, n_hist,
              sga_scenario_rows, nwc_scenario_rows for cross-sheet references.
    """
    ws = wb.create_sheet("Segment Analysis")
    ws.sheet_properties.tabColor = "8B0000"  # Dark red

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 34

    # Determine historical years from segment data
    n_hist = 0
    for seg in segments:
        hist_rev = seg.get("historical", {}).get("revenue", [])
        if len(hist_rev) > n_hist:
            n_hist = len(hist_rev)

    n_data_cols = n_hist + proj_years
    for ci in range(n_data_cols):
        ws.column_dimensions[col_letter(3 + ci)].width = 16

    # Check if segments have revenue_growth (v2) or revenue only (v1 fallback)
    _has_growth = any(
        seg.get("projections", {}).get("revenue_growth")
        for seg in segments
    )
    if not _has_growth:
        print("WARNING: segments use absolute 'revenue' without 'revenue_growth'. "
              "Falling back to v1 absolute revenue mode.")

    # Check for consolidated-level scenario data
    _scenarios = C.get("scenarios", {})
    _has_nwc_pct = (
        C.get("nwc_method") == "revenue_pct"
        and all("nwc_pct" in _scenarios[sn] for sn in SCENARIO_NAMES if sn in _scenarios)
    )

    # Title
    set_cell(ws, 2, 2, f'Segment Analysis - {C["company_name"]}', font=TITLE_FONT)

    # Header row: historical FY labels + projection year labels
    hist_labels = []
    if n_hist >= 3:
        hist_labels = ["FY-2", "FY-1", "FY0 (Base)"]
    elif n_hist == 2:
        hist_labels = ["FY-1", "FY0 (Base)"]
    elif n_hist == 1:
        hist_labels = ["FY0 (Base)"]

    # If projection_start_fy is available, use actual FY labels
    if C.get("projection_start_fy"):
        import re as _re
        _m = _re.search(r"FY(\d+)", C["projection_start_fy"])
        if _m:
            _base_fy = int(_m.group(1))
            hist_labels = [f"FY{_base_fy - n_hist + 1 + i}" for i in range(n_hist)]

    all_headers = hist_labels + year_labels
    header_row(ws, 4, 3, 3 + len(all_headers) - 1, all_headers)

    # ─────────────────────────────────────────────────────────────────
    # FIRST PASS: Build Segment Scenario Input Matrix to know row numbers
    # before writing display area (display area CHOOSE formulas need them).
    # ─────────────────────────────────────────────────────────────────

    # Pre-calculate display area row count to determine matrix start row.
    # Per segment: 1 header + 1 rev + 1 op + 1 opm + 1 blank = 5 rows
    # Total block: 1 header + 1 rev + 1 op + 1 opm + 1 blank = 5 rows
    # Reconciliation: 1 header + 3 rows + 1 blank = 5 rows
    display_rows = len(segments) * 5 + 5 + 5
    matrix_start = 5 + display_rows + 1  # +1 for extra gap

    # Build scenario input matrix row map
    # v2: Revenue Growth (%) + OP Margin blocks per segment
    # Each segment block: header(1) + 5 scenario rows + blank(1) = 7 rows for Revenue Growth
    #                     header(1) + 5 scenario rows + blank(1) = 7 rows for OPM
    # Total per segment: 14 rows
    seg_matrix_info = []
    matrix_cur = matrix_start + 2  # +2 for section header + year headers

    for seg_idx, seg in enumerate(segments):
        # Revenue Growth block (v2)
        rev_sub_header = matrix_cur
        rev_scenario_rows = [matrix_cur + 1 + s for s in range(NUM_SCENARIOS)]
        matrix_cur += 1 + NUM_SCENARIOS + 1  # sub-header + 5 rows + blank

        # OP Margin block
        opm_sub_header = matrix_cur
        opm_scenario_rows = [matrix_cur + 1 + s for s in range(NUM_SCENARIOS)]
        matrix_cur += 1 + NUM_SCENARIOS + 1  # sub-header + 5 rows + blank

        seg_matrix_info.append({
            "rev_sub_header": rev_sub_header,
            "rev_scenario_rows": rev_scenario_rows,
            "opm_sub_header": opm_sub_header,
            "opm_scenario_rows": opm_scenario_rows,
        })

    # ── Consolidated Inputs section (COGS%, SGA%, NWC%) after segment blocks ──
    consolidated_start = matrix_cur + 1  # gap row

    # Check if cogs_pct exists in scenarios
    _has_cogs_pct = any("cogs_pct" in _scenarios.get(sn, {}) for sn in SCENARIO_NAMES)

    # COGS% block (if present): sub-header(1) + 5 rows + blank(1) = 7 rows
    if _has_cogs_pct:
        cogs_sub_header_row = consolidated_start + 2  # after section header + year headers
        cogs_scenario_rows = [cogs_sub_header_row + 1 + s for s in range(NUM_SCENARIOS)]
        sga_sub_header_row = cogs_sub_header_row + 1 + NUM_SCENARIOS + 1  # after COGS block + blank
    else:
        cogs_sub_header_row = None
        cogs_scenario_rows = None
        sga_sub_header_row = consolidated_start + 2  # original position

    sga_scenario_rows = [sga_sub_header_row + 1 + s for s in range(NUM_SCENARIOS)]
    nwc_scenario_rows = None
    if _has_nwc_pct:
        nwc_sub_header_row = sga_sub_header_row + 1 + NUM_SCENARIOS + 1  # after SGA block + blank
        nwc_scenario_rows = [nwc_sub_header_row + 1 + s for s in range(NUM_SCENARIOS)]

    # ─────────────────────────────────────────────────────────────────
    # DISPLAY AREA: Per-segment blocks with CHOOSE formulas
    # ─────────────────────────────────────────────────────────────────
    cur_row = 5  # Start of segment blocks

    seg_rev_rows = []   # Track revenue row for each segment (for Total SUM)
    seg_op_rows = []    # Track OP row for each segment

    for seg_idx, seg in enumerate(segments):
        seg_name = seg.get("name", f"Segment {seg_idx + 1}")
        seg_name_jp = seg.get("name_jp", "")
        display_name = f"{seg_name}" + (f" ({seg_name_jp})" if seg_name_jp else "")

        hist = seg.get("historical", {})
        hist_rev = hist.get("revenue", [None] * n_hist)
        hist_op = hist.get("op", [None] * n_hist)

        # Pad historical arrays to n_hist
        while len(hist_rev) < n_hist:
            hist_rev.insert(0, None)
        while len(hist_op) < n_hist:
            hist_op.insert(0, None)

        mi = seg_matrix_info[seg_idx]

        # ── Segment header ──
        c = section_title(ws, cur_row, 2, display_name)
        c.fill = LIGHT_FILL
        for ci in range(3, 3 + n_data_cols):
            ws.cell(row=cur_row, column=ci).fill = LIGHT_FILL
        cur_row += 1

        # ── Revenue row ──
        rev_row = cur_row
        seg_rev_rows.append(rev_row)
        set_cell(ws, rev_row, 2, "Revenue", font=BOLD_FONT)

        # Historical revenue
        for i in range(n_hist):
            col = 3 + i
            val = hist_rev[i]
            if val is not None:
                set_cell(ws, rev_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                         border=NWC_DATA_BORDER)
            else:
                set_cell(ws, rev_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)

        # Projected revenue — v2: Base Year × (1+growth) chain via CHOOSE
        if _has_growth:
            base_year_col = col_letter(3 + n_hist - 1)  # last historical column
            for yr in range(proj_years):
                col = 3 + n_hist + yr
                cl = col_letter(col)
                # Growth rate CHOOSE references from scenario matrix
                growth_refs = [f"{cl}{r}" for r in mi["rev_scenario_rows"]]
                growth_choose = f"CHOOSE('DCF Model'!$D$27,{','.join(growth_refs)})"
                if yr == 0:
                    # Y1: =BaseYearRev × (1 + CHOOSE(growth))
                    formula = f"={base_year_col}{rev_row}*(1+{growth_choose})"
                else:
                    prev_cl = col_letter(col - 1)
                    # Y2+: =PrevYearRev × (1 + CHOOSE(growth))
                    formula = f"={prev_cl}{rev_row}*(1+{growth_choose})"
                set_cell(ws, rev_row, col, formula, font=BLACK_FONT, fmt=FMT_YEN,
                         border=NWC_DATA_BORDER)
        else:
            # v1 fallback: absolute revenue CHOOSE
            for yr in range(proj_years):
                col = 3 + n_hist + yr
                cl = col_letter(col)
                formula = seg_choose_formula(mi["rev_scenario_rows"], cl)
                set_cell(ws, rev_row, col, formula, font=BLACK_FONT, fmt=FMT_YEN,
                         border=NWC_DATA_BORDER)
        cur_row += 1

        # ── Operating Profit row ──
        op_row = cur_row
        seg_op_rows.append(op_row)
        set_cell(ws, op_row, 2, "Operating Profit", font=BOLD_FONT)

        # Historical OP
        for i in range(n_hist):
            col = 3 + i
            val = hist_op[i]
            if val is not None:
                set_cell(ws, op_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                         border=NWC_DATA_BORDER)
            else:
                set_cell(ws, op_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)

        # Projected OP = Revenue × CHOOSE(OP Margin from scenario matrix)
        for yr in range(proj_years):
            col = 3 + n_hist + yr
            cl = col_letter(col)
            opm_refs = [f"{cl}{r}" for r in mi["opm_scenario_rows"]]
            margin_choose = f"CHOOSE('DCF Model'!$D$27,{','.join(opm_refs)})"
            set_cell(ws, op_row, col,
                     f"={cl}{rev_row}*{margin_choose}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=NWC_DATA_BORDER)
        cur_row += 1

        # ── OP Margin row ──
        opm_row = cur_row
        set_cell(ws, opm_row, 2, "OP Margin", font=Font(name="Arial", size=10,
                 italic=True, color="808080"))

        for i in range(n_hist):
            col = 3 + i
            cl = col_letter(col)
            if hist_rev[i] is not None and hist_op[i] is not None:
                set_cell(ws, opm_row, col, f"={cl}{op_row}/{cl}{rev_row}",
                         font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
            else:
                set_cell(ws, opm_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)

        # Projected OPM — CHOOSE from scenario matrix (display only)
        for yr in range(proj_years):
            col = 3 + n_hist + yr
            cl = col_letter(col)
            formula = seg_choose_formula(mi["opm_scenario_rows"], cl)
            set_cell(ws, opm_row, col, formula,
                     font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
        cur_row += 2  # blank row between segments

    # ═══════════════════════════════════════════════════════════════
    # TOTAL BLOCK
    # ═══════════════════════════════════════════════════════════════
    c = section_title(ws, cur_row, 2, "Consolidated Total")
    c.fill = LIGHT_GREEN
    for ci in range(3, 3 + n_data_cols):
        ws.cell(row=cur_row, column=ci).fill = LIGHT_GREEN
    cur_row += 1

    # Total Revenue
    total_rev_row = cur_row
    set_cell(ws, total_rev_row, 2, "Total Revenue", font=BOLD_FONT)
    for ci in range(n_data_cols):
        col = 3 + ci
        cl = col_letter(col)
        refs = [f"{cl}{r}" for r in seg_rev_rows]
        formula = f"={'+'.join(refs)}"
        set_cell(ws, total_rev_row, col, formula, font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    # Total OP
    total_op_row = cur_row
    set_cell(ws, total_op_row, 2, "Total Operating Profit", font=BOLD_FONT)
    for ci in range(n_data_cols):
        col = 3 + ci
        cl = col_letter(col)
        refs = [f"{cl}{r}" for r in seg_op_rows]
        formula = f"={'+'.join(refs)}"
        set_cell(ws, total_op_row, col, formula, font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    # Total OP Margin
    total_opm_row = cur_row
    set_cell(ws, total_opm_row, 2, "Total OP Margin", font=Font(name="Arial", size=10,
             italic=True, color="808080"))
    for ci in range(n_data_cols):
        col = 3 + ci
        cl = col_letter(col)
        set_cell(ws, total_opm_row, col,
                 f"=IFERROR({cl}{total_op_row}/{cl}{total_rev_row},\"—\")",
                 font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
    cur_row += 2

    # ═══════════════════════════════════════════════════════════════
    # RECONCILIATION CHECK (projected years only)
    # ═══════════════════════════════════════════════════════════════
    c = section_title(ws, cur_row, 2, "Reconciliation vs DCF Model")
    c.fill = LIGHT_YELLOW
    for ci in range(3, 3 + n_data_cols):
        ws.cell(row=cur_row, column=ci).fill = LIGHT_YELLOW
    cur_row += 1

    set_cell(ws, cur_row, 2, "DCF Model Revenue", font=BOLD_FONT)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        dcf_col_letter = col_letter(3 + yr)
        set_cell(ws, cur_row, col,
                 f"='DCF Model'!{dcf_col_letter}{R_REVENUE}",
                 font=GREEN_FONT, fmt=FMT_YEN, border=NWC_DATA_BORDER)
    cur_row += 1

    set_cell(ws, cur_row, 2, "Segment Total Revenue", font=BOLD_FONT)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, cur_row, col,
                 f"={cl}{total_rev_row}",
                 font=BLACK_FONT, fmt=FMT_YEN, border=NWC_DATA_BORDER)
    cur_row += 1

    recon_diff_row = cur_row
    set_cell(ws, recon_diff_row, 2, "Difference", font=BOLD_FONT)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, recon_diff_row, col,
                 f"={cl}{cur_row - 1}-{cl}{cur_row - 2}",
                 font=BLACK_FONT, fmt=FMT_YEN, border=TOP_BOTTOM)
    cur_row += 2

    # ═══════════════════════════════════════════════════════════════
    # SEGMENT SCENARIO INPUT MATRIX
    # ═══════════════════════════════════════════════════════════════
    # Section header
    c = section_title(ws, matrix_start, 2, "Segment Scenario Input Matrix")
    c.fill = PatternFill(start_color="E6CCE6", end_color="E6CCE6", fill_type="solid")
    for ci in range(3, 3 + n_data_cols):
        ws.cell(row=matrix_start, column=ci).fill = PatternFill(
            start_color="E6CCE6", end_color="E6CCE6", fill_type="solid")

    # Year headers (projection columns only)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        set_cell(ws, matrix_start + 1, col, year_labels[yr],
                 font=HEADER_FONT, fill=HEADER_FILL,
                 alignment=Alignment(horizontal="center"))

    # Per-segment scenario blocks
    for seg_idx, seg in enumerate(segments):
        seg_name = seg.get("name", f"Segment {seg_idx + 1}")
        proj = seg.get("projections", {})
        scenario_proj = seg.get("scenario_projections", {})

        mi = seg_matrix_info[seg_idx]

        # ── Revenue Growth block (v2) or Revenue absolute block (v1 fallback) ──
        if _has_growth:
            section_title(ws, mi["rev_sub_header"], 2,
                          f"{seg_name} - Revenue Growth (YoY)")
            for s, scen_name in enumerate(SCENARIO_NAMES):
                r = mi["rev_scenario_rows"][s]
                set_cell(ws, r, 2, scen_name, font=BOLD_FONT)

                # Get scenario-specific revenue_growth; fallback to Base projections
                if (scen_name in scenario_proj
                        and "revenue_growth" in scenario_proj[scen_name]):
                    scen_data = scenario_proj[scen_name]["revenue_growth"]
                else:
                    scen_data = proj.get("revenue_growth", [])

                for yr in range(proj_years):
                    col = 3 + n_hist + yr
                    val = scen_data[yr] if yr < len(scen_data) else None
                    if val is not None:
                        set_cell(ws, r, col, val, font=BLUE_FONT, fmt=FMT_PCT,
                                 border=INPUT_BORDER)
                    else:
                        set_cell(ws, r, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
        else:
            # v1 fallback: absolute revenue
            section_title(ws, mi["rev_sub_header"], 2, f"{seg_name} - Revenue")
            for s, scen_name in enumerate(SCENARIO_NAMES):
                r = mi["rev_scenario_rows"][s]
                set_cell(ws, r, 2, scen_name, font=BOLD_FONT)

                if scen_name in scenario_proj and "revenue" in scenario_proj[scen_name]:
                    scen_rev = scenario_proj[scen_name]["revenue"]
                else:
                    scen_rev = proj.get("revenue", [])

                for yr in range(proj_years):
                    col = 3 + n_hist + yr
                    val = scen_rev[yr] if yr < len(scen_rev) else None
                    if val is not None:
                        set_cell(ws, r, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                                 border=INPUT_BORDER)
                    else:
                        set_cell(ws, r, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)

        # ── OP Margin block (unchanged) ──
        section_title(ws, mi["opm_sub_header"], 2, f"{seg_name} - OP Margin")
        for s, scen_name in enumerate(SCENARIO_NAMES):
            r = mi["opm_scenario_rows"][s]
            set_cell(ws, r, 2, scen_name, font=BOLD_FONT)

            if scen_name in scenario_proj and "op_margin" in scenario_proj[scen_name]:
                scen_opm = scenario_proj[scen_name]["op_margin"]
            else:
                scen_opm = proj.get("op_margin", [])

            for yr in range(proj_years):
                col = 3 + n_hist + yr
                val = scen_opm[yr] if yr < len(scen_opm) else None
                if val is not None:
                    set_cell(ws, r, col, val, font=BLUE_FONT, fmt=FMT_PCT,
                             border=INPUT_BORDER)
                else:
                    set_cell(ws, r, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)

    # ═══════════════════════════════════════════════════════════════
    # CONSOLIDATED INPUTS (SGA%, NWC%) — v2 addition
    # ═══════════════════════════════════════════════════════════════
    CONSOL_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

    c = section_title(ws, consolidated_start, 2, "Consolidated Inputs (連結レベル)")
    c.fill = CONSOL_FILL
    for ci in range(3, 3 + n_data_cols):
        ws.cell(row=consolidated_start, column=ci).fill = CONSOL_FILL

    # Year headers
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        set_cell(ws, consolidated_start + 1, col, year_labels[yr],
                 font=HEADER_FONT, fill=HEADER_FILL,
                 alignment=Alignment(horizontal="center"))

    # ── COGS % of Revenue (only when cogs_pct exists in scenarios) ──
    if _has_cogs_pct and cogs_scenario_rows is not None:
        section_title(ws, cogs_sub_header_row, 2, "COGS % of Revenue")
        for s, scen_name in enumerate(SCENARIO_NAMES):
            r = cogs_scenario_rows[s]
            set_cell(ws, r, 2, scen_name, font=BOLD_FONT)
            scen_data = _scenarios.get(scen_name, {}).get("cogs_pct", [])
            for yr in range(proj_years):
                col = 3 + n_hist + yr
                val = scen_data[yr] if yr < len(scen_data) else None
                if val is not None:
                    set_cell(ws, r, col, val, font=BLUE_FONT, fmt=FMT_PCT,
                             border=INPUT_BORDER)
                else:
                    set_cell(ws, r, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)

    # ── SGA % of Revenue ──
    section_title(ws, sga_sub_header_row, 2, "SGA % of Revenue")
    for s, scen_name in enumerate(SCENARIO_NAMES):
        r = sga_scenario_rows[s]
        set_cell(ws, r, 2, scen_name, font=BOLD_FONT)
        scen_data = _scenarios.get(scen_name, {}).get("sga_pct", [])
        for yr in range(proj_years):
            col = 3 + n_hist + yr
            val = scen_data[yr] if yr < len(scen_data) else None
            if val is not None:
                set_cell(ws, r, col, val, font=BLUE_FONT, fmt=FMT_PCT,
                         border=INPUT_BORDER)
            else:
                set_cell(ws, r, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)

    # ── NWC % of Revenue (only when nwc_pct exists in scenarios) ──
    if _has_nwc_pct and nwc_scenario_rows is not None:
        section_title(ws, nwc_sub_header_row, 2, "NWC % of Revenue")
        for s, scen_name in enumerate(SCENARIO_NAMES):
            r = nwc_scenario_rows[s]
            set_cell(ws, r, 2, scen_name, font=BOLD_FONT)
            scen_data = _scenarios.get(scen_name, {}).get("nwc_pct", [])
            for yr in range(proj_years):
                col = 3 + n_hist + yr
                val = scen_data[yr] if yr < len(scen_data) else None
                if val is not None:
                    set_cell(ws, r, col, val, font=BLUE_FONT, fmt=FMT_PCT,
                             border=INPUT_BORDER)
                else:
                    set_cell(ws, r, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)

    # Freeze pane
    ws.freeze_panes = "C5"

    return {
        "total_rev_row": total_rev_row,
        "total_op_row": total_op_row,
        "n_hist": n_hist,
        "cogs_scenario_rows": cogs_scenario_rows,
        "sga_scenario_rows": sga_scenario_rows,
        "nwc_scenario_rows": nwc_scenario_rows,
    }


# =====================================================================
# DRIVER ANALYSIS SHEET
# =====================================================================
def _create_driver_sheet(wb, C, segments, proj_years, year_labels):
    """Generate Driver Analysis sheet with driver_type-specific sections.

    Supported driver_types:
      - 'backlog': Order backlog → revenue recognition (equipment makers)
      - 'manmonth': Headcount × utilization × unit price (IT services)
      - 'growth_rate': Revenue growth rate based (generic)
      - 'manual': Direct revenue input (no driver decomposition)
    """
    ws = wb.create_sheet("Driver Analysis")
    ws.sheet_properties.tabColor = "4B0082"  # Indigo

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 34

    # Determine historical years (same logic as segment sheet)
    n_hist = 0
    for seg in segments:
        hist_rev = seg.get("historical", {}).get("revenue", [])
        if len(hist_rev) > n_hist:
            n_hist = len(hist_rev)

    n_data_cols = n_hist + proj_years
    for ci in range(n_data_cols):
        ws.column_dimensions[col_letter(3 + ci)].width = 16

    set_cell(ws, 2, 2, f'Driver Analysis - {C["company_name"]}', font=TITLE_FONT)

    # Header row
    hist_labels = []
    if n_hist >= 3:
        hist_labels = ["FY-2", "FY-1", "FY0 (Base)"]
    elif n_hist == 2:
        hist_labels = ["FY-1", "FY0 (Base)"]
    elif n_hist == 1:
        hist_labels = ["FY0 (Base)"]

    if C.get("projection_start_fy"):
        import re as _re
        _m = _re.search(r"FY(\d+)", C["projection_start_fy"])
        if _m:
            _base_fy = int(_m.group(1))
            hist_labels = [f"FY{_base_fy - n_hist + 1 + i}" for i in range(n_hist)]

    all_headers = hist_labels + year_labels
    header_row(ws, 4, 3, 3 + len(all_headers) - 1, all_headers)

    cur_row = 5

    for seg_idx, seg in enumerate(segments):
        seg_name = seg.get("name", f"Segment {seg_idx + 1}")
        seg_name_jp = seg.get("name_jp", "")
        display_name = f"{seg_name}" + (f" ({seg_name_jp})" if seg_name_jp else "")
        driver_type = seg.get("driver_type", "manual")

        hist = seg.get("historical", {})
        proj = seg.get("projections", {})

        # ── Segment header ──
        c = section_title(ws, cur_row, 2, f"{display_name} [{driver_type}]")
        c.fill = LIGHT_FILL
        for ci in range(3, 3 + n_data_cols):
            ws.cell(row=cur_row, column=ci).fill = LIGHT_FILL
        cur_row += 1

        # ══════════════════════════════════════════════════════════
        if driver_type == "backlog":
            cur_row = _driver_backlog(ws, seg, hist, proj, n_hist, proj_years, cur_row)

        elif driver_type == "manmonth":
            cur_row = _driver_manmonth(ws, seg, hist, proj, n_hist, proj_years, cur_row)

        elif driver_type == "growth_rate":
            cur_row = _driver_growth_rate(ws, seg, hist, proj, n_hist, proj_years, cur_row)

        elif driver_type == "retail":
            cur_row = _driver_retail(ws, seg, hist, proj, n_hist, proj_years, cur_row)

        elif driver_type == "subscription":
            cur_row = _driver_subscription(ws, seg, hist, proj, n_hist, proj_years, cur_row)

        else:  # "manual" or unknown
            cur_row = _driver_manual(ws, seg, hist, proj, n_hist, proj_years, cur_row)

        cur_row += 1  # blank row between segments

    # Freeze pane
    ws.freeze_panes = "C5"


# ── Driver sub-functions ──

def _driver_backlog(ws, seg, hist, proj, n_hist, proj_years, cur_row):
    """Backlog-based driver: Beginning Backlog + Orders - Revenue = Ending Backlog."""

    hist_rev = hist.get("revenue", [None] * n_hist)
    hist_orders = hist.get("orders", [None] * n_hist)
    hist_backlog = hist.get("backlog_end", [None] * n_hist)
    proj_rev = proj.get("revenue", [None] * proj_years)
    proj_orders = proj.get("orders", [None] * proj_years)

    # Pad to n_hist
    while len(hist_rev) < n_hist:
        hist_rev.insert(0, None)
    while len(hist_orders) < n_hist:
        hist_orders.insert(0, None)
    while len(hist_backlog) < n_hist:
        hist_backlog.insert(0, None)

    n_data_cols = n_hist + proj_years

    # Row: Beginning Backlog
    bb_row = cur_row
    set_cell(ws, bb_row, 2, "Beginning Backlog", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        if i > 0 and hist_backlog[i - 1] is not None:
            set_cell(ws, bb_row, col, hist_backlog[i - 1], font=BLUE_FONT,
                     fmt=FMT_YEN, border=NWC_DATA_BORDER)
        else:
            set_cell(ws, bb_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        prev_cl = col_letter(col - 1)
        # Beginning Backlog = previous period's Ending Backlog
        eb_row = bb_row + 3  # Ending backlog row (calculated below)
        if yr == 0 and hist_backlog[-1] is not None:
            set_cell(ws, bb_row, col, hist_backlog[-1], font=BLUE_FONT,
                     fmt=FMT_YEN, border=NWC_DATA_BORDER)
        else:
            set_cell(ws, bb_row, col, f"={prev_cl}{eb_row}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=NWC_DATA_BORDER)
    cur_row += 1

    # Row: + New Orders
    ord_row = cur_row
    set_cell(ws, ord_row, 2, "+ New Orders", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_orders[i]
        if val is not None:
            set_cell(ws, ord_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=NWC_DATA_BORDER)
        else:
            set_cell(ws, ord_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_orders[yr] if yr < len(proj_orders) else None
        if val is not None:
            set_cell(ws, ord_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, ord_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # Row: - Revenue (Recognized)
    rev_row = cur_row
    set_cell(ws, rev_row, 2, "- Revenue (Recognized)", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_rev[i]
        if val is not None:
            set_cell(ws, rev_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=NWC_DATA_BORDER)
        else:
            set_cell(ws, rev_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_rev[yr] if yr < len(proj_rev) else None
        if val is not None:
            set_cell(ws, rev_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, rev_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # Row: = Ending Backlog (formula: BB + Orders - Revenue)
    eb_row = cur_row
    set_cell(ws, eb_row, 2, "Ending Backlog", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_backlog[i]
        if val is not None:
            set_cell(ws, eb_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        else:
            set_cell(ws, eb_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, eb_row, col,
                 f"={cl}{bb_row}+{cl}{ord_row}-{cl}{rev_row}",
                 font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    # Row: Book-to-Bill ratio
    btb_row = cur_row
    set_cell(ws, btb_row, 2, "Book-to-Bill",
             font=Font(name="Arial", size=10, italic=True, color="808080"))
    for ci in range(n_data_cols):
        col = 3 + ci
        cl = col_letter(col)
        set_cell(ws, btb_row, col,
                 f"=IFERROR({cl}{ord_row}/{cl}{rev_row},\"—\")",
                 font=BLACK_FONT, fmt=FMT_RATIO, border=NWC_DATA_BORDER)
    cur_row += 1

    return cur_row


def _driver_manmonth(ws, seg, hist, proj, n_hist, proj_years, cur_row):
    """Man-month driver: HC × Utilization × Unit Price × 12 = Layer 1, + Layer 2."""

    hist_rev = hist.get("revenue", [None] * n_hist)
    hist_l2 = hist.get("layer2_revenue", [None] * n_hist)
    proj_hc = proj.get("headcount", [None] * proj_years)
    proj_util = proj.get("utilization", [None] * proj_years)
    proj_price = proj.get("unit_price_monthly", [None] * proj_years)
    proj_l2 = proj.get("layer2_revenue", [None] * proj_years)

    while len(hist_rev) < n_hist:
        hist_rev.insert(0, None)
    while len(hist_l2) < n_hist:
        hist_l2.insert(0, None)

    n_data_cols = n_hist + proj_years

    # Headcount
    hc_row = cur_row
    set_cell(ws, hc_row, 2, "Headcount", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, hc_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_hc[yr] if yr < len(proj_hc) else None
        if val is not None:
            set_cell(ws, hc_row, col, val, font=BLUE_FONT, fmt=FMT_INT,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, hc_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # Utilization
    util_row = cur_row
    set_cell(ws, util_row, 2, "× Utilization", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, util_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_util[yr] if yr < len(proj_util) else None
        if val is not None:
            set_cell(ws, util_row, col, val, font=BLUE_FONT, fmt=FMT_PCT,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, util_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # Unit Price (M/month)
    price_row = cur_row
    set_cell(ws, price_row, 2, "× Unit Price (JPY mn/month)", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, price_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_price[yr] if yr < len(proj_price) else None
        if val is not None:
            set_cell(ws, price_row, col, val, font=BLUE_FONT, fmt=FMT_YEN_DEC,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, price_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # Layer 1 Revenue = HC × Util × Price × 12
    l1_row = cur_row
    set_cell(ws, l1_row, 2, "Layer 1 Revenue", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, l1_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, l1_row, col,
                 f"={cl}{hc_row}*{cl}{util_row}*{cl}{price_row}*12",
                 font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    # Layer 2 (Solution) Revenue
    l2_row = cur_row
    set_cell(ws, l2_row, 2, "+ Layer 2 (Solution) Revenue", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_l2[i]
        if val is not None:
            set_cell(ws, l2_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=NWC_DATA_BORDER)
        else:
            set_cell(ws, l2_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_l2[yr] if yr < len(proj_l2) else None
        if val is not None:
            set_cell(ws, l2_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, l2_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # Segment Revenue = L1 + L2
    seg_rev_row = cur_row
    set_cell(ws, seg_rev_row, 2, "Segment Revenue", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_rev[i]
        if val is not None:
            set_cell(ws, seg_rev_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        else:
            set_cell(ws, seg_rev_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, seg_rev_row, col,
                 f"={cl}{l1_row}+{cl}{l2_row}",
                 font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    # Layer 2 Mix %
    mix_row = cur_row
    set_cell(ws, mix_row, 2, "Layer 2 Mix %",
             font=Font(name="Arial", size=10, italic=True, color="808080"))
    for ci in range(n_data_cols):
        col = 3 + ci
        cl = col_letter(col)
        set_cell(ws, mix_row, col,
                 f"=IFERROR({cl}{l2_row}/{cl}{seg_rev_row},\"—\")",
                 font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
    cur_row += 1

    return cur_row


def _driver_growth_rate(ws, seg, hist, proj, n_hist, proj_years, cur_row):
    """Growth rate driver: Revenue = Prior × (1 + g)."""

    hist_rev = hist.get("revenue", [None] * n_hist)
    proj_growth = proj.get("revenue_growth", [None] * proj_years)

    while len(hist_rev) < n_hist:
        hist_rev.insert(0, None)

    n_data_cols = n_hist + proj_years

    # Revenue Growth row
    g_row = cur_row
    set_cell(ws, g_row, 2, "Revenue Growth (YoY)", font=BOLD_FONT)
    # Historical: compute from data
    for i in range(n_hist):
        col = 3 + i
        cl = col_letter(col)
        if i == 0 or hist_rev[i] is None or hist_rev[i - 1] is None:
            set_cell(ws, g_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
        else:
            prev_cl = col_letter(col - 1)
            rev_row = cur_row + 1  # revenue row is next
            set_cell(ws, g_row, col,
                     f"=({cl}{rev_row}-{prev_cl}{rev_row})/{prev_cl}{rev_row}",
                     font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_growth[yr] if yr < len(proj_growth) else None
        if val is not None:
            set_cell(ws, g_row, col, val, font=BLUE_FONT, fmt=FMT_PCT,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, g_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # Revenue row
    rev_row = cur_row
    set_cell(ws, rev_row, 2, "Revenue", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_rev[i]
        if val is not None:
            set_cell(ws, rev_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=NWC_DATA_BORDER)
        else:
            set_cell(ws, rev_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        prev_cl = col_letter(col - 1)
        set_cell(ws, rev_row, col,
                 f"={prev_cl}{rev_row}*(1+{cl}{g_row})",
                 font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    return cur_row


def _driver_manual(ws, seg, hist, proj, n_hist, proj_years, cur_row):
    """Manual driver: Revenue directly input, no decomposition."""

    hist_rev = hist.get("revenue", [None] * n_hist)
    proj_rev = proj.get("revenue", [None] * proj_years)

    while len(hist_rev) < n_hist:
        hist_rev.insert(0, None)

    n_data_cols = n_hist + proj_years

    # Revenue row (direct input)
    rev_row = cur_row
    set_cell(ws, rev_row, 2, "Revenue (Direct Input)", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_rev[i]
        if val is not None:
            set_cell(ws, rev_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=NWC_DATA_BORDER)
        else:
            set_cell(ws, rev_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_rev[yr] if yr < len(proj_rev) else None
        if val is not None:
            set_cell(ws, rev_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, rev_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # YoY Growth (computed)
    g_row = cur_row
    set_cell(ws, g_row, 2, "YoY Growth",
             font=Font(name="Arial", size=10, italic=True, color="808080"))
    for ci in range(n_data_cols):
        col = 3 + ci
        cl = col_letter(col)
        if ci == 0:
            set_cell(ws, g_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
        else:
            prev_cl = col_letter(col - 1)
            set_cell(ws, g_row, col,
                     f"=IFERROR(({cl}{rev_row}-{prev_cl}{rev_row})/{prev_cl}{rev_row},\"—\")",
                     font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
    cur_row += 1

    return cur_row


def _driver_retail(ws, seg, hist, proj, n_hist, proj_years, cur_row):
    """Retail driver: Store count rollforward + SSSG + new store contribution."""

    hist_rev = hist.get("revenue", [None] * n_hist)
    hist_store = hist.get("store_count", [None] * n_hist)
    proj_new_stores = proj.get("new_stores", [None] * proj_years)
    proj_closures = proj.get("closures", [None] * proj_years)
    proj_sssg = proj.get("sssg", [None] * proj_years)
    new_store_months = proj.get("new_store_months", 6)

    while len(hist_rev) < n_hist:
        hist_rev.insert(0, None)
    while len(hist_store) < n_hist:
        hist_store.insert(0, None)

    n_data_cols = n_hist + proj_years

    # R1: Beginning Store Count
    beg_row = cur_row
    set_cell(ws, beg_row, 2, "Beginning Store Count", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        if i > 0 and hist_store[i - 1] is not None:
            set_cell(ws, beg_row, col, hist_store[i - 1], font=BLUE_FONT,
                     fmt=FMT_INT, border=NWC_DATA_BORDER)
        else:
            set_cell(ws, beg_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        prev_cl = col_letter(col - 1)
        end_row = beg_row + 3  # Ending Store Count row
        if yr == 0 and hist_store[-1] is not None:
            set_cell(ws, beg_row, col, hist_store[-1], font=BLUE_FONT,
                     fmt=FMT_INT, border=NWC_DATA_BORDER)
        else:
            set_cell(ws, beg_row, col, f"={prev_cl}{end_row}",
                     font=BLACK_FONT, fmt=FMT_INT, border=NWC_DATA_BORDER)
    cur_row += 1

    # R2: + New Stores
    new_row = cur_row
    set_cell(ws, new_row, 2, "+ New Stores", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, new_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_new_stores[yr] if yr < len(proj_new_stores) else None
        if val is not None:
            set_cell(ws, new_row, col, val, font=BLUE_FONT, fmt=FMT_INT,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, new_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # R3: - Closures
    close_row = cur_row
    set_cell(ws, close_row, 2, "- Closures", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, close_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_closures[yr] if yr < len(proj_closures) else None
        if val is not None:
            set_cell(ws, close_row, col, val, font=BLUE_FONT, fmt=FMT_INT,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, close_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # R4: Ending Store Count = Beginning + New - Closures
    end_sc_row = cur_row
    set_cell(ws, end_sc_row, 2, "Ending Store Count", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_store[i]
        if val is not None:
            set_cell(ws, end_sc_row, col, val, font=BLUE_FONT, fmt=FMT_INT,
                     border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        else:
            set_cell(ws, end_sc_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, end_sc_row, col,
                 f"={cl}{beg_row}+{cl}{new_row}-{cl}{close_row}",
                 font=BLACK_FONT, fmt=FMT_INT,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    # R5: SSSG
    sssg_row = cur_row
    set_cell(ws, sssg_row, 2, "SSSG", font=BOLD_FONT)
    # Historical: reverse-calc from per-store revenue YoY
    avg_row = beg_row + 8  # Avg Store Count row (R9)
    seg_rev_row = beg_row + 7  # Segment Revenue row (R8)
    for i in range(n_hist):
        col = 3 + i
        cl = col_letter(col)
        prev_cl = col_letter(col - 1)
        if i == 0:
            set_cell(ws, sssg_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
        else:
            set_cell(ws, sssg_row, col,
                     f"=IFERROR(({cl}{seg_rev_row}/{cl}{avg_row})/({prev_cl}{seg_rev_row}/{prev_cl}{avg_row})-1,\"—\")",
                     font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_sssg[yr] if yr < len(proj_sssg) else None
        if val is not None:
            set_cell(ws, sssg_row, col, val, font=BLUE_FONT, fmt=FMT_PCT,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, sssg_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # R6: Existing Store Revenue = Prior Segment Revenue × (1 + SSSG)
    exist_row = cur_row
    set_cell(ws, exist_row, 2, "Existing Store Revenue", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, exist_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        prev_cl = col_letter(col - 1)
        set_cell(ws, exist_row, col,
                 f"={prev_cl}{seg_rev_row}*(1+{cl}{sssg_row})",
                 font=BLACK_FONT, fmt=FMT_YEN, border=NWC_DATA_BORDER)
    cur_row += 1

    # R7: + New Store Revenue = New Stores × (Prior Rev / Prior Ending SC) × (months/12)
    new_rev_row = cur_row
    set_cell(ws, new_rev_row, 2, "+ New Store Revenue", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, new_rev_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    month_frac = round(new_store_months / 12, 6)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        prev_cl = col_letter(col - 1)
        set_cell(ws, new_rev_row, col,
                 f"={cl}{new_row}*({prev_cl}{seg_rev_row}/{prev_cl}{end_sc_row})*{month_frac}",
                 font=BLACK_FONT, fmt=FMT_YEN, border=NWC_DATA_BORDER)
    cur_row += 1

    # R8: Segment Revenue
    seg_rev_row_actual = cur_row
    assert seg_rev_row_actual == seg_rev_row, "Row layout mismatch for Segment Revenue"
    set_cell(ws, seg_rev_row, 2, "Segment Revenue", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_rev[i]
        if val is not None:
            set_cell(ws, seg_rev_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        else:
            set_cell(ws, seg_rev_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, seg_rev_row, col,
                 f"={cl}{exist_row}+{cl}{new_rev_row}",
                 font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    # R9: Avg Store Count (KPI)
    avg_row_actual = cur_row
    assert avg_row_actual == avg_row, "Row layout mismatch for Avg Store Count"
    set_cell(ws, avg_row, 2, "Avg Store Count",
             font=Font(name="Arial", size=10, italic=True, color="808080"))
    for ci in range(n_data_cols):
        col = 3 + ci
        cl = col_letter(col)
        set_cell(ws, avg_row, col,
                 f"=IFERROR(({cl}{beg_row}+{cl}{end_sc_row})/2,\"—\")",
                 font=BLACK_FONT, fmt=FMT_INT, border=NWC_DATA_BORDER)
    cur_row += 1

    return cur_row


def _driver_subscription(ws, seg, hist, proj, n_hist, proj_years, cur_row):
    """Subscription driver: ARR bridge (churn/expansion/new) with NRR fallback."""

    hist_rev = hist.get("revenue", [None] * n_hist)
    hist_arr = hist.get("arr_end", [None] * n_hist)
    proj_nrr = proj.get("nrr", [None] * proj_years)
    proj_churn_rate = proj.get("churn_rate", None)
    proj_new_arr = proj.get("new_arr", [None] * proj_years)

    while len(hist_rev) < n_hist:
        hist_rev.insert(0, None)
    while len(hist_arr) < n_hist:
        hist_arr.insert(0, None)

    n_data_cols = n_hist + proj_years

    # ── NRR fallback: pre-compute churn/expansion values ──
    proj_beg_arr = []
    proj_churned = []
    proj_expansion = []
    proj_end_arr = []

    for yr in range(proj_years):
        if yr == 0:
            beg = ([v for v in hist_arr if v is not None] or [0])[-1]
        else:
            beg = proj_end_arr[yr - 1]
        proj_beg_arr.append(beg)

        nrr = proj_nrr[yr] if yr < len(proj_nrr) and proj_nrr[yr] is not None else 1.0
        if proj_churn_rate and yr < len(proj_churn_rate) and proj_churn_rate[yr] is not None:
            churn_r = proj_churn_rate[yr]
        else:
            churn_r = 0.0
        expansion_r = nrr - (1 - churn_r)

        churned = beg * churn_r
        expansion = beg * expansion_r
        new_arr_val = proj_new_arr[yr] if yr < len(proj_new_arr) and proj_new_arr[yr] is not None else 0

        proj_churned.append(churned)
        proj_expansion.append(expansion)
        proj_end_arr.append(beg - churned + expansion + new_arr_val)

    # S1: Beginning ARR
    beg_arr_row = cur_row
    set_cell(ws, beg_arr_row, 2, "Beginning ARR", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        if i > 0 and hist_arr[i - 1] is not None:
            set_cell(ws, beg_arr_row, col, hist_arr[i - 1], font=BLUE_FONT,
                     fmt=FMT_YEN, border=NWC_DATA_BORDER)
        else:
            set_cell(ws, beg_arr_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    end_arr_row = beg_arr_row + 4  # S5: Ending ARR
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        prev_cl = col_letter(col - 1)
        if yr == 0 and hist_arr[-1] is not None:
            set_cell(ws, beg_arr_row, col, hist_arr[-1], font=BLUE_FONT,
                     fmt=FMT_YEN, border=NWC_DATA_BORDER)
        else:
            set_cell(ws, beg_arr_row, col, f"={prev_cl}{end_arr_row}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=NWC_DATA_BORDER)
    cur_row += 1

    # S2: - Churned ARR (positive value, subtracted in S5 formula)
    churn_row = cur_row
    set_cell(ws, churn_row, 2, "- Churned ARR", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, churn_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        set_cell(ws, churn_row, col, proj_churned[yr], font=BLUE_FONT, fmt=FMT_YEN,
                 border=INPUT_BORDER)
    cur_row += 1

    # S3: + Expansion ARR
    exp_row = cur_row
    set_cell(ws, exp_row, 2, "+ Expansion ARR", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, exp_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        set_cell(ws, exp_row, col, proj_expansion[yr], font=BLUE_FONT, fmt=FMT_YEN,
                 border=INPUT_BORDER)
    cur_row += 1

    # S4: + New ARR
    new_arr_row = cur_row
    set_cell(ws, new_arr_row, 2, "+ New ARR", font=BOLD_FONT)
    for i in range(n_hist):
        set_cell(ws, new_arr_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        val = proj_new_arr[yr] if yr < len(proj_new_arr) and proj_new_arr[yr] is not None else None
        if val is not None:
            set_cell(ws, new_arr_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=INPUT_BORDER)
        else:
            set_cell(ws, new_arr_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    cur_row += 1

    # S5: Ending ARR = Beginning - Churned + Expansion + New
    end_arr_row_actual = cur_row
    assert end_arr_row_actual == end_arr_row, "Row layout mismatch for Ending ARR"
    set_cell(ws, end_arr_row, 2, "Ending ARR", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_arr[i]
        if val is not None:
            set_cell(ws, end_arr_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        else:
            set_cell(ws, end_arr_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, end_arr_row, col,
                 f"={cl}{beg_arr_row}-{cl}{churn_row}+{cl}{exp_row}+{cl}{new_arr_row}",
                 font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    # S6: Revenue = (Beginning ARR + Ending ARR) / 2
    rev_row = cur_row
    set_cell(ws, rev_row, 2, "Revenue", font=BOLD_FONT)
    for i in range(n_hist):
        col = 3 + i
        val = hist_rev[i]
        if val is not None:
            set_cell(ws, rev_row, col, val, font=BLUE_FONT, fmt=FMT_YEN,
                     border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        else:
            set_cell(ws, rev_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, rev_row, col,
                 f"=({cl}{beg_arr_row}+{cl}{end_arr_row})/2",
                 font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
    cur_row += 1

    # S7: NRR % (KPI)
    nrr_row = cur_row
    set_cell(ws, nrr_row, 2, "NRR %",
             font=Font(name="Arial", size=10, italic=True, color="808080"))
    for i in range(n_hist):
        set_cell(ws, nrr_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, nrr_row, col,
                 f"=IFERROR(({cl}{end_arr_row}-{cl}{new_arr_row})/{cl}{beg_arr_row},\"—\")",
                 font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
    cur_row += 1

    # S8: Gross Churn % (KPI)
    gc_row = cur_row
    set_cell(ws, gc_row, 2, "Gross Churn %",
             font=Font(name="Arial", size=10, italic=True, color="808080"))
    for i in range(n_hist):
        set_cell(ws, gc_row, 3 + i, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
    for yr in range(proj_years):
        col = 3 + n_hist + yr
        cl = col_letter(col)
        set_cell(ws, gc_row, col,
                 f"=IFERROR({cl}{churn_row}/{cl}{beg_arr_row},\"—\")",
                 font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
    cur_row += 1

    # S9: ARR YoY Growth (KPI)
    yoy_row = cur_row
    set_cell(ws, yoy_row, 2, "ARR YoY Growth",
             font=Font(name="Arial", size=10, italic=True, color="808080"))
    for ci in range(n_data_cols):
        col = 3 + ci
        cl = col_letter(col)
        if ci == 0:
            set_cell(ws, yoy_row, col, "—", font=GREY_FONT, border=NWC_DATA_BORDER)
        else:
            prev_cl = col_letter(col - 1)
            set_cell(ws, yoy_row, col,
                     f"=IFERROR({cl}{end_arr_row}/{prev_cl}{end_arr_row}-1,\"—\")",
                     font=BLACK_FONT, fmt=FMT_PCT, border=NWC_DATA_BORDER)
    cur_row += 1

    return cur_row


# =====================================================================
# BUILD WORKBOOK
# =====================================================================
def _reverse_dcf_api():
    """Import the Reverse DCF builder lazily (it imports styles back from here)."""
    try:
        from templates.reverse_dcf_sheet import (
            resolve_params, build_reverse_dcf_sheet, place_after)
    except ImportError:                                          # flat sys.path
        from reverse_dcf_sheet import (
            resolve_params, build_reverse_dcf_sheet, place_after)
    return resolve_params, build_reverse_dcf_sheet, place_after


def generate_dcf_workbook(config, output_path=None):
    """Generate DCF/Comps Excel workbook from config dict.

    Args:
        config: Dict with all financial data and assumptions.
        output_path: Output file path. If None, auto-generates.

    Returns:
        str: Path to saved Excel file.
    """
    C = config

    # ── Generation metadata ──
    # Everything a machine needs to re-check the workbook after the fact
    # (scripts/validate_output.py) is recorded here and written to the
    # 'Adjustments Log' sheet as a "Pipeline Metadata" block. Without it the
    # validator would have to re-derive the intent from the numbers, which is
    # exactly the guesswork that let year-shifted data through before.
    _meta = {}

    # ── Comps Analysis layout (needed by Executive Summary, built first) ──
    # The sheet keeps its historical row numbers (statistics 14-17, implied
    # valuation 19-28) so existing consumers of 'Comps Analysis'!C27/C28 keep
    # working; a comps set large enough to reach row 14 shifts the whole lower
    # block down instead of silently overwriting the statistics.
    _n_comps = len(C.get("comps") or [])
    _comps_row_shift = max(0, (5 + _n_comps - 1) + 3 - 14) if _n_comps else 0
    R_CMP_SHARES = 23 + _comps_row_shift
    R_CMP_NETDEBT = 24 + _comps_row_shift
    R_CMP_IMPL_MULT = 27 + _comps_row_shift   # EV/EBITDA (or EV/Sales) implied
    R_CMP_IMPL_PER = 28 + _comps_row_shift    # PER implied
    R_STAT_SEC = 14 + _comps_row_shift        # "Statistics" section header
    R_STAT_MEDIAN = R_STAT_SEC + 2            # 25th / MEDIAN / 75th

    # ── Locate the subject company's own row in the comps table ──────────
    # The comps CSV carries the subject as its first row (for display), but its
    # own multiple must never enter the peer statistics — a company cannot be
    # its own comparable. Matching is by ticker, so the row is found wherever
    # the analyst puts it (no "row 5" assumption anywhere below).
    def _tkr_key(t):
        return str(t or "").strip().upper().split(".")[0]

    _subject_key = _tkr_key(C.get("ticker"))
    subject_idx = None
    for _i, _comp in enumerate(C.get("comps") or []):
        if _subject_key and _tkr_key(_comp.get("ticker")) == _subject_key:
            subject_idx = _i
            break
    if subject_idx is None:
        _cname = str(C.get("company_name", "")).strip()
        for _i, _comp in enumerate(C.get("comps") or []):
            if _cname and str(_comp.get("name", "")).strip() == _cname:
                subject_idx = _i
                break
    R_CMP_SUBJECT = (5 + subject_idx) if subject_idx is not None else None

    # ── Normalize WACC inputs ──
    # Beta: clamp to [0.6, 1.75]; outside range → sector-standard 1.0
    # Upper bound is 1.75 (not 1.5) to admit high-beta growth names (e.g. ELEMENTS
    # 5246 at 1.6) without collapsing them to 1.0.
    raw_beta = C.get("beta", 1.0)
    if not raw_beta or raw_beta < 0.6 or raw_beta > 1.75:
        C["beta"] = 1.0
    # Size Premium: auto-determine from market cap (JPY mn) unless explicitly overridden
    if "size_premium" not in C.get("_override_keys", set()):
        try:
            mkt_cap = C["current_price"] * C["shares_outstanding"] / 1_000_000
        except (KeyError, TypeError):
            mkt_cap = 0
        if mkt_cap >= 1_000_000:        # >= 1 trillion JPY
            C["size_premium"] = 0.0
        elif mkt_cap >= 100_000:         # >= 100 billion JPY
            C["size_premium"] = 0.015
        else:
            C["size_premium"] = 0.03

    # ── Cost of debt: use the ACTUAL rate when the inputs are supplied ──
    # The template default (2.8% pre-tax) is a JGB-plus-spread guess. When the
    # filing gives interest expense and the debt balances, the realised rate is
    # a fact, not an estimate — 5726 moved from an assumed 2.80% to an actual
    # 0.85% pre-tax, which is 0.4pt of WACC. The average balance defaults to the
    # last two hist_debt years (opening + closing of the latest FY).
    _cod_default_at = C.get("cost_of_debt_at")
    _cod_actual = None
    if C.get("interest_expense") is not None:
        _interest = float(C["interest_expense"]) + float(C.get("loan_fees") or 0)
        _d_beg, _d_end = C.get("debt_beginning"), C.get("debt_ending")
        _basis = "debt_beginning/debt_ending"
        if _d_beg is None or _d_end is None:
            _hd = [d for d in (C.get("hist_debt") or [])
                   if isinstance(d, (int, float)) and not isinstance(d, bool)]
            if len(_hd) >= 2:
                _d_beg, _d_end = _hd[-2], _hd[-1]
                _basis = "hist_debt[-2:]"
        if _d_beg and _d_end and (_d_beg + _d_end) > 0:
            _avg_debt = (_d_beg + _d_end) / 2.0
            _kd_pre = _interest / _avg_debt
            # Rounded to 4dp for the same reason C5/C18 are: the sheet shows a
            # rate, and an unrounded tail invites false precision.
            _kd_at = round(_kd_pre * (1 - C["tax_rate"]), 4)
            C["cost_of_debt_at"] = _kd_at
            _cod_actual = {
                "interest": _interest,
                "interest_expense": float(C["interest_expense"]),
                "loan_fees": float(C.get("loan_fees") or 0),
                "avg_debt": _avg_debt,
                "debt_beginning": _d_beg,
                "debt_ending": _d_end,
                "basis": _basis,
                "kd_pretax": _kd_pre,
                "kd_after_tax": _kd_at,
                "default_after_tax": _cod_default_at,
            }
            print(f"  [WACC] Actual cost of debt: {_interest:,.0f} / "
                  f"{_avg_debt:,.0f} = {_kd_pre:.4%} pre-tax -> {_kd_at:.4%} "
                  f"after-tax (was {(_cod_default_at or 0):.4%}; basis {_basis})")
        else:
            print("  WARNING: interest_expense supplied but no usable debt "
                  "balances (set debt_beginning / debt_ending, or give at least "
                  "two hist_debt years) - falling back to cost_of_debt_at.")
    _meta["cost_of_debt_basis"] = "actual" if _cod_actual else "assumption"
    if _cod_actual:
        _meta["cost_of_debt_actual"] = (
            f"interest {_cod_actual['interest']:,.0f} "
            f"(expense {_cod_actual['interest_expense']:,.0f} + fees "
            f"{_cod_actual['loan_fees']:,.0f}) / avg debt "
            f"{_cod_actual['avg_debt']:,.0f} ({_cod_actual['basis']}) = "
            f"{_cod_actual['kd_pretax']:.6f} pre-tax -> "
            f"{_cod_actual['kd_after_tax']:.6f} after-tax"
        )
        # Structured copies so validate_output can re-derive the rate instead of
        # parsing prose out of the note above.
        _meta["cost_of_debt_interest"] = _cod_actual["interest"]
        _meta["cost_of_debt_avg_debt"] = _cod_actual["avg_debt"]
        _meta["cost_of_debt_pretax"] = _cod_actual["kd_pretax"]
    _meta["cost_of_debt_at_c11"] = C.get("cost_of_debt_at")
    _meta["tax_rate_c6"] = C.get("tax_rate")

    # Restore flat arrays from Base scenario
    _base = C["scenarios"]["Base"]
    if "revenue_growth" in _base:
        C["revenue_growth"] = _base["revenue_growth"]
    if "cogs_pct" in _base:
        C["cogs_pct"] = _base["cogs_pct"]
    C["sga_pct"] = _base["sga_pct"]

    # When segments exist without scenario-level revenue_growth/cogs_pct,
    # compute implied values from segment data for potential downstream use
    if C.get("segments") and "revenue_growth" not in C:
        _segs = C["segments"]
        _n_proj = C.get("projection_years", 5)
        _base_revs = []
        for yr in range(_n_proj):
            total = 0
            for seg in _segs:
                hist_rev = seg.get("historical", {}).get("revenue", [])
                base_rev = ([v for v in hist_rev if v is not None] or [0])[-1]
                growths = seg.get("projections", {}).get("revenue_growth", [])
                r = base_rev
                for y in range(yr + 1):
                    g = growths[y] if y < len(growths) else 0
                    r = r * (1 + g)
                total += r
            _base_revs.append(total)
        _by_rev = C.get("base_year_revenue", 1)
        C["revenue_growth"] = [(_base_revs[0] / _by_rev) - 1 if _by_rev else 0]
        for i in range(1, len(_base_revs)):
            C["revenue_growth"].append(
                (_base_revs[i] / _base_revs[i - 1]) - 1 if _base_revs[i - 1] else 0
            )
    if C.get("segments") and "cogs_pct" not in C:
        # Derive implied COGS% from segment OP margins and SGA%
        C["cogs_pct"] = [0.78] * C.get("projection_years", 5)  # reasonable default

    USE_EV_SALES = (C.get("primary_multiple", "EV/EBITDA") == "EV/Sales")

    # ── Meaningless-multiple guards ──
    # A median multiple applied to a negative (or missing) base metric yields a
    # meaningless implied price (e.g. PER × net loss) that would silently
    # contaminate the Target Mid average. Excluded methods are written as the
    # text "N/A" — AVERAGE/MIN/MAX skip text cells — and the exclusion is
    # disclosed on the Executive Summary (never dropped silently).
    _ni = C.get("core_net_income")
    PER_EXCLUDED = not (isinstance(_ni, (int, float)) and _ni > 0)
    _eb = C.get("core_ebitda")
    EBITDA_EXCLUDED = (not USE_EV_SALES) and not (isinstance(_eb, (int, float)) and _eb > 0)
    # EV/Sales exit: when primary is EV/Sales AND an exit_sales_multiple is given,
    # the DCF Exit-Multiple terminal value uses Year-5 Revenue × exit_sales_multiple
    # instead of Year-5 EBITDA × exit_multiple. Backward-compatible: off unless both
    # conditions hold, so EBITDA-exit tickers (e.g. 4192) are unaffected.
    USE_EV_SALES_EXIT = USE_EV_SALES and C.get("exit_sales_multiple") is not None

    # ── Pre-calculate Segment Analysis row numbers (needed by DCF Model) ──
    has_segments = bool(C.get("segments"))
    # Check that segments use revenue_growth (v2 format)
    if has_segments:
        _has_seg_growth = any(
            s.get("projections", {}).get("revenue_growth")
            for s in C["segments"]
        )
        if not _has_seg_growth:
            print("WARNING: segments lack 'revenue_growth' - falling back to legacy mode")
            has_segments = False  # disable segment-linked mode

    seg_info = None
    if has_segments:
        _segments = C["segments"]
        _n_hist_seg = 0
        for _seg in _segments:
            _hr = _seg.get("historical", {}).get("revenue", [])
            if len(_hr) > _n_hist_seg:
                _n_hist_seg = len(_hr)
        # Display area: per segment 5 rows (header+rev+op+opm+blank)
        # Total block: header(1) + rev(1) + op(1) + opm(1) + blank(1) = 5
        _seg_display_rows = len(_segments) * 5
        # total_rev_row = 5 + seg_display_rows + 1 (header) → first data row
        _total_rev_row = 5 + _seg_display_rows + 1
        _total_op_row = _total_rev_row + 1

        # Pre-calculate Consolidated Inputs row positions (SGA%, NWC%)
        # Matrix layout: section header(1) + year headers(1)
        #   per segment: rev_growth block(7) + opm block(7) = 14 rows
        _matrix_start = 5 + _seg_display_rows + 5 + 5 + 1  # display+total+recon+gap
        _matrix_cur = _matrix_start + 2  # +2 for section header + year headers
        _matrix_cur += len(_segments) * 14  # 14 rows per segment

        _consolidated_start = _matrix_cur + 1  # gap
        _scenarios_dict = C.get("scenarios", {})

        _has_cogs_pct = any("cogs_pct" in _scenarios_dict.get(sn, {}) for sn in SCENARIO_NAMES)
        if _has_cogs_pct:
            _cogs_sub_header = _consolidated_start + 2
            _cogs_scenario_rows = [_cogs_sub_header + 1 + s for s in range(NUM_SCENARIOS)]
            _sga_sub_header = _cogs_sub_header + 1 + NUM_SCENARIOS + 1
        else:
            _cogs_scenario_rows = None
            _sga_sub_header = _consolidated_start + 2
        _sga_scenario_rows = [_sga_sub_header + 1 + s for s in range(NUM_SCENARIOS)]
        _has_nwc_pct = (
            C.get("nwc_method") == "revenue_pct"
            and all(
                "nwc_pct" in _scenarios_dict[sn]
                for sn in SCENARIO_NAMES if sn in _scenarios_dict
            )
        )
        _nwc_scenario_rows = None
        if _has_nwc_pct:
            _nwc_sub_header = _sga_sub_header + 1 + NUM_SCENARIOS + 1
            _nwc_scenario_rows = [_nwc_sub_header + 1 + s for s in range(NUM_SCENARIOS)]

        seg_info = {
            "total_rev_row": _total_rev_row,
            "total_op_row": _total_op_row,
            "n_hist": _n_hist_seg,
            "cogs_scenario_rows": _cogs_scenario_rows,
            "sga_scenario_rows": _sga_scenario_rows,
            "nwc_scenario_rows": _nwc_scenario_rows,
        }

    (_rdcf_resolve_params, _build_reverse_dcf_sheet,
     _rdcf_place_after) = _reverse_dcf_api()

    wb = openpyxl.Workbook()

    # =====================================================================
    # SHEET 1: Executive Summary
    # =====================================================================
    ws1 = wb.active
    ws1.title = "Executive Summary"
    ws1.sheet_properties.tabColor = "003366"

    ws1.column_dimensions["A"].width = 3
    ws1.column_dimensions["B"].width = 30
    ws1.column_dimensions["C"].width = 22
    ws1.column_dimensions["D"].width = 22
    ws1.column_dimensions["E"].width = 22

    # Disclaimer
    set_cell(ws1, 1, 2,
        "DISCLAIMER: This is a sample analysis for demonstration purposes only. "
        "It does not constitute investment advice.",
        font=GREY_FONT)
    ws1.merge_cells("B1:E1")

    # Title
    set_cell(ws1, 3, 2, f'{C["company_name"]} ({C["ticker"]})', font=TITLE_FONT)
    ws1.merge_cells("B3:D3")
    set_cell(ws1, 4, 2, "Equity Research Report", font=SUB_FONT)

    # Company info
    set_cell(ws1, 6, 2, "Company", font=BOLD_FONT)
    set_cell(ws1, 6, 3, C["company_name"])
    set_cell(ws1, 7, 2, "Ticker", font=BOLD_FONT)
    set_cell(ws1, 7, 3, f'{C["ticker"]} ({C["exchange"]})')
    set_cell(ws1, 8, 2, "Sector", font=BOLD_FONT)
    set_cell(ws1, 8, 3, C["sector"])
    set_cell(ws1, 9, 2, "Current Price", font=BOLD_FONT)
    set_cell(ws1, 9, 3, C["current_price"], font=BLUE_FONT, fmt=FMT_YEN)

    # Target Price (Mid) = the average of the TWO DCF legs only (C16:C17).
    # Comps (C18:C19) are reference marks, never Target inputs: the standard
    # (docs/DCFフォーマット標準メモ §2, established on 3905 / 5726) is that a
    # peer median prices the market's mood, not the business, so mixing it into
    # the target silently turns a DCF into a half-comps blend. Excluded DCF legs
    # hold the text "INVALID ..." / "N/A", which AVERAGE skips; COUNT guards the
    # case where BOTH legs are excluded (otherwise AVERAGE returns #DIV/0!).
    set_cell(ws1, 10, 2, "Target Price (DCF Mid)", font=BOLD_FONT)
    set_cell(ws1, 10, 3, '=IF(COUNT(C16:C17)=0,"N/A",ROUND(AVERAGE(C16:C17),0))',
             font=BLACK_FONT, fmt=FMT_YEN)

    # Recommendation
    set_cell(ws1, 11, 2, "Recommendation", font=BOLD_FONT)
    set_cell(ws1, 11, 3,
             '=IF(NOT(ISNUMBER(C12)),"N/A",'
             'IF(C12>0.15,"BUY",IF(C12>0.05,"HOLD","SELL")))', font=BLACK_FONT)

    # Upside / Downside
    set_cell(ws1, 12, 2, "Upside / Downside", font=BOLD_FONT)
    set_cell(ws1, 12, 3, '=IF(ISNUMBER(C10),(C10-C9)/C9,"N/A")',
             font=BLACK_FONT, fmt=FMT_PCT)

    # Valuation Summary section
    c = set_cell(ws1, 14, 2, "Valuation Summary", font=SUB_FONT)
    c.fill = LIGHT_FILL
    for col_idx in range(3, 5):
        ws1.cell(row=14, column=col_idx).fill = LIGHT_FILL

    header_row(ws1, 15, 2, 4, ["Methodology", "Implied Value (JPY)", "vs Current Price"])

    # DCF - Perpetuity Growth
    # Guard: when the PGM enterprise value falls below net debt (+MI) the equity
    # value is negative and the "implied price" is an artefact, not a valuation
    # (8267's negative PGM). It is surfaced as the text "INVALID" so AVERAGE /
    # MIN / MAX below skip it — the method drops out of Target Price instead of
    # dragging the average to a meaningless number.
    set_cell(ws1, 16, 2, "DCF - Perpetuity Growth")
    set_cell(ws1, 16, 3,
             f"=IF('DCF Model'!C{R_EV_PGM}<'DCF Model'!C16,"
             f"\"INVALID (EV < net debt)\",'DCF Model'!C{R_PRICE_PGM})",
             font=GREEN_FONT, fmt=FMT_YEN)
    set_cell(ws1, 16, 4, '=IF(ISNUMBER(C16),(C16-C9)/C9,"N/A")', font=BLACK_FONT, fmt=FMT_PCT)

    # DCF - Exit Multiple
    # The PGM leg has carried the EV < net debt guard since 8267; the Exit leg
    # had none, so a Downside scenario whose exit EV falls below net debt kept a
    # NEGATIVE implied price as a plain number and dragged the Target average
    # down (5726 Downside 2: Exit -558 survived while PGM was already INVALID).
    set_cell(ws1, 17, 2, "DCF - Exit Multiple")
    set_cell(ws1, 17, 3,
             f"=IF('DCF Model'!C{R_EV_EXIT}<'DCF Model'!C16,"
             f"\"INVALID (EV < net debt)\",'DCF Model'!C{R_PRICE_EXIT})",
             font=GREEN_FONT, fmt=FMT_YEN)
    set_cell(ws1, 17, 4, '=IF(ISNUMBER(C17),(C17-C9)/C9,"N/A")',
             font=BLACK_FONT, fmt=FMT_PCT)

    # Comps — reference marks only. The label says so on the sheet so a reader
    # cannot mistake them for Target inputs.
    if USE_EV_SALES:
        set_cell(ws1, 18, 2, "Comps - EV/Sales Median [参考・Target不算入]")
    else:
        set_cell(ws1, 18, 2, "Comps - EV/EBITDA Median [参考・Target不算入]")
    set_cell(ws1, 18, 3, f"='Comps Analysis'!C{R_CMP_IMPL_MULT}", font=GREEN_FONT, fmt=FMT_YEN)
    set_cell(ws1, 18, 4, '=IF(ISNUMBER(C18),(C18-C9)/C9,"N/A")', font=BLACK_FONT, fmt=FMT_PCT)

    set_cell(ws1, 19, 2, "Comps - PER Median [参考・Target不算入]")
    set_cell(ws1, 19, 3, f"='Comps Analysis'!C{R_CMP_IMPL_PER}", font=GREEN_FONT, fmt=FMT_YEN)
    set_cell(ws1, 19, 4, '=IF(ISNUMBER(C19),(C19-C9)/C9,"N/A")', font=BLACK_FONT, fmt=FMT_PCT)

    # Disclose the Target's composition and any excluded method explicitly —
    # never drop one silently, and never leave the reader to guess whether the
    # comps rows fed the average.
    _excluded_methods = ["Target Mid は DCF 2法（PGM / Exit）の平均のみ。"
                         "Comps 2法は [参考] で Target 不算入"]
    if PER_EXCLUDED:
        _excluded_methods.append("PER法は赤字（純利益≦0）のため除外")
    if EBITDA_EXCLUDED:
        _excluded_methods.append("EV/EBITDA法はEBITDA≦0のため除外")
    set_cell(ws1, 20, 2,
             "Note: " + "；".join(_excluded_methods)
             + "（EV<ネットデットで株式価値が負になる手法は INVALID として平均から除外）",
             font=GREY_FONT)
    ws1.merge_cells("B20:E20")

    # Valuation Range — the SAME two methods the target averages. Spanning all
    # four put the excluded comps back in through the side door: 5726 read
    # "200 - 2,247" while the DCF said 200-558, which is not a range anyone
    # should quote. MIN/MAX over an all-text range returns 0, so the row reports
    # "N/A" rather than a fabricated "0 - 0".
    set_cell(ws1, 21, 2, "DCF Valuation Range (PGM - Exit)", font=BOLD_FONT)
    set_cell(ws1, 21, 3,
             '=IF(COUNT(C16:C17)=0,"N/A",MIN(C16:C17)&" - "&MAX(C16:C17))',
             font=BLACK_FONT)

    # ── Investment Thesis / Key Risks (token-aware) ──
    # Cell addresses are taken from where this template just wrote each value,
    # never hardcoded, so the prose cannot drift from the numbers it cites.
    _token_refs = {
        "price":        "C9",
        "target_price": "C10",
        "upside_pct":   "C12",
        "wacc":         "'DCF Model'!C26",
        "pb":           (f"'Comps Analysis'!M{R_CMP_SUBJECT}" if R_CMP_SUBJECT else None),
        "per":          (f"'Comps Analysis'!L{R_CMP_SUBJECT}" if R_CMP_SUBJECT else None),
    }
    _unresolved = set()

    def _warn_token(name):
        _unresolved.add(name)

    def _render_block(lines):
        out = []
        for line in lines:
            out.extend(_render_narrative_line(line, _token_refs, _warn_token))
        return out

    _thesis_cells = _render_block(C["investment_thesis"])
    _risk_cells = _render_block(C["key_risks"])
    if _unresolved:
        print(f"  WARNING: narrative token(s) {sorted(_unresolved)} have no cell to "
              f"reference in this model (no comps subject row?) - rendered as plain text.")

    c = set_cell(ws1, 23, 2, "Key Investment Thesis", font=SUB_FONT)
    c.fill = LIGHT_FILL
    for col_idx in range(3, 5):
        ws1.cell(row=23, column=col_idx).fill = LIGHT_FILL
    for i, line in enumerate(_thesis_cells):
        set_cell(ws1, 24 + i, 2, line)

    # Key Risks: row 28 as before; only pushed down when the thesis needs the room
    _risk_hdr_row = max(28, 24 + len(_thesis_cells) + 1)
    c = set_cell(ws1, _risk_hdr_row, 2, "Key Risks", font=SUB_FONT)
    c.fill = LIGHT_FILL
    for col_idx in range(3, 5):
        ws1.cell(row=_risk_hdr_row, column=col_idx).fill = LIGHT_FILL
    for i, line in enumerate(_risk_cells):
        set_cell(ws1, _risk_hdr_row + 1 + i, 2, line)

    _meta["narrative_tokens_used"] = ",".join(sorted(
        {m.group(1) for line in list(C["investment_thesis"]) + list(C["key_risks"])
         if isinstance(line, str) for m in _TOKEN_RE.finditer(line)}
    )) or "none"

    # =====================================================================
    # SHEET 2: Financial Statements (V3 — Full PL Waterfall + BS Highlights)
    # =====================================================================
    ws2 = wb.create_sheet("Financial Statements")
    ws2.sheet_properties.tabColor = "003366"

    ws2.column_dimensions["A"].width = 3
    ws2.column_dimensions["B"].width = 32
    n_hist = len(C["hist_years"])
    for _ci in range(n_hist):
        ws2.column_dimensions[col_letter(3 + _ci)].width = 18

    set_cell(ws2, 2, 2, f'{C["company_name"]} - Historical Financials (JPY mn)', font=TITLE_FONT)

    header_row(ws2, 4, 3, 3 + n_hist - 1, C["hist_years"])

    # ── Income Statement (V3 full waterfall) ──
    section_title(ws2, 5, 2, "Income Statement")

    set_cell(ws2, 6, 2, "Revenue", font=BOLD_FONT)
    set_cell(ws2, 7, 2, "COGS", font=BOLD_FONT)
    set_cell(ws2, 8, 2, "Gross Profit", font=BOLD_FONT)
    set_cell(ws2, 9, 2, "Gross Margin")
    set_cell(ws2, 10, 2, "SGA", font=BOLD_FONT)
    set_cell(ws2, 11, 2, "Operating Income", font=BOLD_FONT)
    set_cell(ws2, 12, 2, "Net Income", font=BOLD_FONT)
    set_cell(ws2, 13, 2, "Operating Margin")
    set_cell(ws2, 14, 2, "Net Margin")
    set_cell(ws2, 15, 2, "Revenue Growth (YoY)")
    set_cell(ws2, 16, 2, "Operating Income Growth (YoY)")

    # A derived P/L row is written only when the cells it divides by are real.
    # This is the same rule the cash-flow block below already followed (bug B1);
    # the income statement did not, so a fiscal year EDINET could not extract
    # produced "=(C11-B11)/B11" against a blank B11 and the sheet carried
    # #DIV/0! (8 tickers in the 2026-09-05 batch). A year with no data now reads
    # as missing rather than as a division by an imaginary zero.
    def _pl(series, i):
        s = C.get(series)
        v = s[i] if s and i < len(s) else None
        return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

    for i in range(n_hist):
        col = 3 + i
        cl = col_letter(col)

        rev_val = _pl("hist_revenue", i)
        cogs_val = _pl("hist_cogs", i)
        sga_val = _pl("hist_sga", i)
        oi_val = _pl("hist_operating_income", i)
        ni_val = _pl("hist_net_income", i)

        set_cell(ws2, 6, col, rev_val, font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws2, 7, col, cogs_val, font=BLUE_FONT, fmt=FMT_YEN)
        # Gross Profit = Revenue - COGS
        if rev_val is not None and cogs_val is not None:
            set_cell(ws2, 8, col, f"={cl}6-{cl}7", font=BLACK_FONT, fmt=FMT_YEN)
            if rev_val:
                set_cell(ws2, 9, col, f"={cl}8/{cl}6", font=BLACK_FONT, fmt=FMT_PCT)
        set_cell(ws2, 10, col, sga_val, font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws2, 11, col, oi_val, font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws2, 12, col, ni_val, font=BLUE_FONT, fmt=FMT_YEN)
        # Margins need a non-zero revenue denominator AND a numerator
        if rev_val:
            if oi_val is not None:
                set_cell(ws2, 13, col, f"={cl}11/{cl}6", font=BLACK_FONT, fmt=FMT_PCT)
            if ni_val is not None:
                set_cell(ws2, 14, col, f"={cl}12/{cl}6", font=BLACK_FONT, fmt=FMT_PCT)
        # YoY Growth
        if i == 0:
            set_cell(ws2, 15, col, "n/a")
            set_cell(ws2, 16, col, "n/a")
        else:
            prev_cl = col_letter(col - 1)
            prev_rev = _pl("hist_revenue", i - 1)
            prev_oi = _pl("hist_operating_income", i - 1)
            if rev_val is not None and prev_rev:
                set_cell(ws2, 15, col, f"=({cl}6-{prev_cl}6)/{prev_cl}6",
                         font=BLACK_FONT, fmt=FMT_PCT)
            else:
                set_cell(ws2, 15, col, "n/a")
            # A prior-year operating income of exactly zero is a legitimate
            # value but not a legitimate denominator, so it is excluded here too.
            if oi_val is not None and prev_oi:
                set_cell(ws2, 16, col, f"=({cl}11-{prev_cl}11)/{prev_cl}11",
                         font=BLACK_FONT, fmt=FMT_PCT)
            else:
                set_cell(ws2, 16, col, "n/a")

    # ── Cash Flow Statement ──
    section_title(ws2, 18, 2, "Cash Flow Statement")
    set_cell(ws2, 19, 2, "Operating Cash Flow", font=BOLD_FONT)
    set_cell(ws2, 20, 2, "Capex", font=BOLD_FONT)
    set_cell(ws2, 21, 2, "Free Cash Flow", font=BOLD_FONT)
    set_cell(ws2, 22, 2, "FCF Margin")
    set_cell(ws2, 23, 2, "Capex / Revenue")

    for i in range(n_hist):
        col = 3 + i
        cl = col_letter(col)
        ocf_val = C["hist_ocf"][i] if C["hist_ocf"] and i < len(C["hist_ocf"]) else None
        capex_val = C["hist_capex"][i] if C["hist_capex"] and i < len(C["hist_capex"]) else None
        set_cell(ws2, 19, col, ocf_val, font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws2, 20, col, capex_val, font=BLUE_FONT, fmt=FMT_YEN)
        # Derived rows are left blank when a source year is blank, so a missing
        # year reads as missing instead of as "FCF = -capex" (bug B1).
        if ocf_val is not None and capex_val is not None:
            set_cell(ws2, 21, col, f"={cl}19-{cl}20", font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws2, 22, col, f"={cl}21/{cl}6", font=BLACK_FONT, fmt=FMT_PCT)
        if capex_val is not None:
            set_cell(ws2, 23, col, f"={cl}20/{cl}6", font=BLACK_FONT, fmt=FMT_PCT)

    # ── Balance Sheet Highlights ──
    section_title(ws2, 25, 2, "Balance Sheet Highlights")
    set_cell(ws2, 26, 2, "Cash & Deposits", font=BOLD_FONT)
    # generate_dcf.py feeds this row from EDINET's `total_debt` (short-term +
    # long-term borrowings + bonds + lease obligations), never the short-term
    # line alone. The old "Short-term Debt" label understated the balance by the
    # whole long-term leg to anyone reading the sheet.
    set_cell(ws2, 27, 2, "Total Interest-bearing Debt (short + long)", font=BOLD_FONT)
    set_cell(ws2, 28, 2, "Net Debt (Cash)", font=BOLD_FONT)

    for i in range(n_hist):
        col = 3 + i
        cl = col_letter(col)
        cash_val = C["hist_cash"][i] if C["hist_cash"] and i < len(C["hist_cash"]) else None
        debt_val = C["hist_debt"][i] if C["hist_debt"] and i < len(C["hist_debt"]) else None
        set_cell(ws2, 26, col, cash_val, font=BLUE_FONT, fmt=FMT_YEN)
        # Blank stays blank: writing 0 for "no data" used to present an unknown
        # debt balance as a confirmed zero (bug B1).
        set_cell(ws2, 27, col, debt_val, font=BLUE_FONT, fmt=FMT_YEN)
        # Net Debt = Debt - Cash (negative = net cash)
        if cash_val is not None and debt_val is not None:
            set_cell(ws2, 28, col, f"={cl}27-{cl}26", font=BLACK_FONT, fmt=FMT_YEN)

    # Year-key coverage note for the CF/BS blocks (set by generate_dcf.py)
    if C.get("_fs_year_coverage"):
        set_cell(ws2, 30, 2,
                 f"Note: OCF / Cash / Debt coverage {C['_fs_year_coverage']} — blank "
                 f"cells are years with no year-key match in the source data "
                 f"(verify against 決算短信 or set hist_ocf / hist_cash / hist_debt "
                 f"in overrides).",
                 font=GREY_FONT)

    # ── Compute NWC Change row (needed before DCF Model sheet) ──
    _nwc_method = C.get("nwc_method", "days")
    if _nwc_method == "itemized":
        _nwc_items = C.get("nwc_items", [])
        _n_total = len(_nwc_items)
        _n_assets = sum(1 for it in _nwc_items if it["side"] == "asset")
        _n_liabs = sum(1 for it in _nwc_items if it["side"] == "liability")
        _itm_cogs = 5 + _n_total + 1 + 2  # ITM_PNL_HDR + 2
        _itm_ca_total = _itm_cogs + 2 + 1 + _n_assets  # after blank, WC header, assets
        _itm_cl_total = _itm_ca_total + 1 + _n_liabs
        _itm_nwc = _itm_cl_total + 2 + 1  # blank, header, NWC
        actual_nwc_chg_row = _itm_nwc + 1
    else:
        actual_nwc_chg_row = NWC_R_CHG_NWC  # 19

    # =====================================================================
    # SHEET 3: DCF Model (V3 — Full Waterfall)
    # =====================================================================
    ws3 = wb.create_sheet("DCF Model")
    ws3.sheet_properties.tabColor = "006600"

    ws3.column_dimensions["A"].width = 3
    ws3.column_dimensions["B"].width = 32
    for letter in ["C", "D", "E", "F", "G"]:
        ws3.column_dimensions[letter].width = 16

    set_cell(ws3, 2, 2, f'DCF Valuation Model - {C["company_name"]}', font=TITLE_FONT)

    # ── Assumptions ──
    c = section_title(ws3, 4, 2, "Assumptions")
    c.fill = LIGHT_FILL
    for col_idx in range(3, 8):
        ws3.cell(row=4, column=col_idx).fill = LIGHT_FILL

    # Determine effective Capex/D&A assumption values for display.
    #
    # Under the "direct" method these cells are the FALLBACK ratio: the per-year
    # projection arrays drive the FCF rows, and C5/C18 only step in for years the
    # arrays do not cover. They used to be back-solved from the projections
    # themselves (mean(projections) / base-year revenue), which made a
    # "fallback" that just echoed the forecast — circular, and wrong whenever the
    # forecast ramped. They are now the plain historical ratio: the mean of the
    # last three years of actual capex (D&A) over actual revenue.
    _capex_method = C.get("capex_method", "revenue_pct")
    _da_method = C.get("da_method", "revenue_pct")
    _capex_pct_display = C["capex_pct"]
    _da_pct_display = C["da_pct"]

    def _hist_ratio_3yr(num_key):
        """Mean of the last <=3 historical <num>/revenue ratios (None if no data)."""
        nums = C.get(num_key) or []
        revs = C.get("hist_revenue") or []
        ratios = [
            n / d for n, d in zip(nums, revs)
            if isinstance(n, (int, float)) and isinstance(d, (int, float))
            and d > 0 and n > 0
        ]
        ratios = ratios[-3:]
        return sum(ratios) / len(ratios) if ratios else None

    _capex_hist3 = _hist_ratio_3yr("hist_capex")
    _da_hist3 = _hist_ratio_3yr("hist_depreciation")
    _capex_label = "Capex / Revenue"
    _da_label = "D&A / Revenue"
    _capex_basis = "assumption"
    _da_basis = "assumption"

    if _capex_method == "direct":
        if _capex_hist3 is not None:
            _capex_pct_display = round(_capex_hist3, 4)
            _capex_label = "Capex / Revenue (hist 3yr avg, fallback)"
            _capex_basis = "hist_3yr_avg"
        else:
            _capex_label = "Capex / Revenue (fallback)"
            print("  WARNING: capex_method=direct but no historical capex/revenue "
                  "pairs - C5 falls back to the capex_pct assumption.")
    if _da_method == "direct":
        if _da_hist3 is not None:
            _da_pct_display = round(_da_hist3, 4)
            _da_label = "D&A / Revenue (hist 3yr avg, fallback)"
            _da_basis = "hist_3yr_avg"
        else:
            _da_label = "D&A / Revenue (fallback)"
            print("  WARNING: da_method=direct but no historical D&A/revenue "
                  "pairs - C18 falls back to the da_pct assumption.")

    # Keep the python-side sensitivity helpers on the same fallback ratio as the sheet
    C["capex_pct"] = _capex_pct_display
    C["da_pct"] = _da_pct_display

    _meta["capex_pct_c5"] = _capex_pct_display
    _meta["da_pct_c18"] = _da_pct_display
    _meta["capex_pct_basis"] = _capex_basis
    _meta["da_pct_basis"] = _da_basis
    _meta["capex_pct_hist3yr"] = round(_capex_hist3, 6) if _capex_hist3 is not None else "n/a"
    _meta["da_pct_hist3yr"] = round(_da_hist3, 6) if _da_hist3 is not None else "n/a"

    assumptions = [
        # Revenue Growth Rate and COGS % removed — now in per-year driver rows
        (_capex_label,                 _capex_pct_display,        FMT_PCT),      # C5
        ("Effective Tax Rate",         C["tax_rate"],             FMT_PCT),      # C6
        ("Risk-Free Rate",             C["risk_free"],            FMT_PCT),      # C7
        ("Beta",                       C["beta"],                 "0.00"),       # C8
        ("Equity Risk Premium",        C["erp"],                  FMT_PCT),      # C9
        ("Size Premium",               C["size_premium"],         FMT_PCT),      # C10
        ("After-tax Cost of Debt",     C["cost_of_debt_at"],      FMT_PCT),      # C11
        ("D/E Ratio",                  C["de_ratio"],             "0.000"),      # C12
        ("Terminal Growth Rate",       C["terminal_growth"],      FMT_PCT),      # C13
        (("Exit Multiple (EV/Sales)" if USE_EV_SALES_EXIT else "Exit Multiple (EV/EBITDA)"),
         (C["exit_sales_multiple"] if USE_EV_SALES_EXIT else C["exit_multiple"]),
         FMT_RATIO),    # C14
        ("Fully Diluted Shares",       C["shares_outstanding"],   FMT_INT),      # C15
        ("Net Debt (JPY mn)",          C["net_debt"],             FMT_YEN),      # C16
        ("Base Year Revenue (JPY mn)", C["base_year_revenue"],    FMT_YEN),      # C17
        (_da_label,                    _da_pct_display,           FMT_PCT),      # C18
        ("Stub Fraction (yr remaining)", C.get("stub_fraction", 1.0), "0.00"),  # C19
        (("LTM Revenue (JPY mn) (override)" if C.get("_ltm_revenue_overridden")
          else "LTM Revenue (JPY mn)"),
         C.get("ltm_revenue", C["base_year_revenue"]), FMT_YEN),  # C20
    ]
    _meta["ltm_revenue_c20"] = C.get("ltm_revenue", C["base_year_revenue"])
    _meta["ltm_revenue_source"] = ("override" if C.get("_ltm_revenue_overridden")
                                   else C.get("_ltm_revenue_source", "auto"))
    if C.get("_ltm_revenue_components"):
        _meta["ltm_revenue_components"] = C["_ltm_revenue_components"]
    for i, (label, val, fmt) in enumerate(assumptions):
        r = 5 + i
        set_cell(ws3, r, 2, label, font=BOLD_FONT)
        set_cell(ws3, r, 3, val, font=BLUE_FONT, fmt=fmt)

    # ── WACC Calculation ──
    c = section_title(ws3, 22, 2, "WACC Calculation")
    c.fill = LIGHT_FILL
    for col_idx in range(3, 8):
        ws3.cell(row=22, column=col_idx).fill = LIGHT_FILL

    set_cell(ws3, 23, 2, "Cost of Equity (Ke)", font=BOLD_FONT)
    set_cell(ws3, 23, 3, "=C7+C8*C9+C10", font=BLACK_FONT, fmt=FMT_PCT2)

    set_cell(ws3, 24, 2, "Weight of Equity", font=BOLD_FONT)
    set_cell(ws3, 24, 3, "=1/(1+C12)", font=BLACK_FONT, fmt=FMT_PCT2)

    set_cell(ws3, 25, 2, "Weight of Debt", font=BOLD_FONT)
    set_cell(ws3, 25, 3, "=C12/(1+C12)", font=BLACK_FONT, fmt=FMT_PCT2)

    set_cell(ws3, 26, 2, "WACC", font=BOLD_FONT)
    set_cell(ws3, 26, 3, "=C23*C24+C11*C25", font=BLACK_FONT, fmt=FMT_PCT2)

    # ── Active Scenario Selector (Row 27) ──
    set_cell(ws3, 27, 2, "Active Scenario", font=BOLD_FONT)
    set_cell(ws3, 27, 3, "Base", font=BLUE_FONT, border=INPUT_BORDER)
    # D27 = MATCH index (1-5) driving every CHOOSE() in the workbook.
    # ALWAYS a MATCH against a live range of the 5 scenario-name cells — never a
    # constant, and never an inline array constant (which cannot be audited and
    # silently decouples from the sheet if a scenario name is edited).
    # The range is derived from the rows the template itself writes the scenario
    # names into, so it follows any layout change automatically:
    #   - no segments  -> DCF Model's own Scenario Input Matrix (growth block)
    #   - segments     -> Segment Analysis's Consolidated Inputs SGA% block,
    #                     the one scenario-name column that always exists there.
    if has_segments:
        _scen_name_rows = seg_info["sga_scenario_rows"]
        _scen_range = (f"'Segment Analysis'!B{_scen_name_rows[0]}"
                       f":B{_scen_name_rows[-1]}")
    else:
        _scen_range = (f"B{R_SCEN_BLK_GROWTH + 1}"
                       f":B{R_SCEN_BLK_GROWTH + NUM_SCENARIOS}")
    set_cell(ws3, 27, 4, f"=MATCH(C27,{_scen_range},0)", font=BLACK_FONT)
    _meta["scenario_index_range"] = _scen_range

    dv_scenario = DataValidation(
        type="list",
        formula1='"Base,Upside,Management,Downside 1,Downside 2"',
        allow_blank=False,
        showDropDown=False,   # openpyxl quirk: False = show dropdown
    )
    dv_scenario.add("C27")
    ws3.add_data_validation(dv_scenario)

    # ── Projected FCF (V3 Full Waterfall) ──
    c = section_title(ws3, 28, 2, "Projected Free Cash Flow")
    c.fill = LIGHT_FILL
    for col_idx in range(3, 8):
        ws3.cell(row=28, column=col_idx).fill = LIGHT_FILL

    proj_years = C["projection_years"]
    year_labels = [f"Year {y}" for y in range(1, proj_years + 1)]
    # Use actual FY labels if projection_start_fy is set
    if C.get("projection_start_fy"):
        import re as _re
        _m = _re.search(r"FY(\d+)", C["projection_start_fy"])
        if _m:
            _base_fy = int(_m.group(1))
            year_labels = [f"FY{_base_fy + y}(E)" for y in range(proj_years)]
    header_row(ws3, 29, 3, 3 + proj_years - 1, year_labels)

    # Driver row labels
    row_labels_drv = [
        ("Revenue Growth (YoY)",          R_DRV_GROWTH),
        ("COGS % of Revenue",             R_DRV_COGS),
        ("SGA % of Revenue",              R_DRV_SGA),
    ]
    for label, r in row_labels_drv:
        set_cell(ws3, r, 2, label, font=BOLD_FONT)

    # V3: FCF row labels — full waterfall
    row_labels_fcf = [
        ("Revenue",                       R_REVENUE),
        ("COGS",                          R_COGS),
        ("Gross Profit",                  R_GROSS_PROFIT),
        ("Gross Margin",                  R_GROSS_MARGIN),
        ("SGA Expense",                   R_SGA),
        ("Implied Operating Margin",      R_OP_M_IMPL),
        ("Operating Income (EBIT)",       R_EBIT),
        ("Less: Tax",                     R_TAX),
        ("NOPAT",                         R_NOPAT),
        ("Plus: D&A",                     R_DA),
        ("Less: Capex",                   R_CAPEX),
        ("Change in NWC",                 R_CHG_NWC),
        ("Unlevered Free Cash Flow",      R_UFCF),
        ("Discount Factor",               R_DISC),
        ("PV of FCF",                     R_PV_FCF),
    ]
    for label, r in row_labels_fcf:
        set_cell(ws3, r, 2, label, font=BOLD_FONT)

    # V3: Year-by-year projection loop — full waterfall with driver rows
    for yr in range(proj_years):
        col = 3 + yr
        cl = col_letter(col)
        prev_cl = col_letter(col - 1) if yr > 0 else None

        if has_segments:
            # ── Segment-linked mode: Revenue & EBIT from Segment Analysis ──
            seg_cl = col_letter(3 + seg_info["n_hist"] + yr)  # Segment sheet column

            # Revenue Growth — back-calculated display
            if yr == 0:
                set_cell(ws3, R_DRV_GROWTH, col, f"=({cl}{R_REVENUE}/C17)-1",
                         font=BLACK_FONT, fmt=FMT_PCT, fill=LIGHT_FILL)
            else:
                set_cell(ws3, R_DRV_GROWTH, col, f"=({cl}{R_REVENUE}/{prev_cl}{R_REVENUE})-1",
                         font=BLACK_FONT, fmt=FMT_PCT, fill=LIGHT_FILL)

            # COGS% — CHOOSE from Segment Analysis Consolidated Inputs (or back-calc if no cogs override)
            if seg_info.get("cogs_scenario_rows"):
                cogs_refs = [f"'Segment Analysis'!{seg_cl}{r}" for r in seg_info["cogs_scenario_rows"]]
                set_cell(ws3, R_DRV_COGS, col,
                         f"=CHOOSE($D$27,{','.join(cogs_refs)})",
                         font=BLACK_FONT, fmt=FMT_PCT, fill=LIGHT_FILL)
            else:
                set_cell(ws3, R_DRV_COGS, col, f"=IFERROR({cl}{R_COGS}/{cl}{R_REVENUE},0)",
                         font=BLACK_FONT, fmt=FMT_PCT, fill=LIGHT_FILL)

            # SGA% — CHOOSE from Segment Analysis Consolidated Inputs
            sga_refs = [f"'Segment Analysis'!{seg_cl}{r}" for r in seg_info["sga_scenario_rows"]]
            set_cell(ws3, R_DRV_SGA, col,
                     f"=CHOOSE($D$27,{','.join(sga_refs)})",
                     font=BLACK_FONT, fmt=FMT_PCT, fill=LIGHT_FILL)

            # Revenue — from Segment Analysis total
            set_cell(ws3, R_REVENUE, col,
                     f"='Segment Analysis'!{seg_cl}{seg_info['total_rev_row']}",
                     font=BLACK_FONT, fmt=FMT_YEN)

            # COGS — from COGS% driver when override exists, otherwise back-calc
            if seg_info.get("cogs_scenario_rows"):
                set_cell(ws3, R_COGS, col, f"={cl}{R_REVENUE}*{cl}{R_DRV_COGS}",
                         font=BLACK_FONT, fmt=FMT_YEN)
            else:
                set_cell(ws3, R_COGS, col, f"={cl}{R_REVENUE}-{cl}{R_SGA}-{cl}{R_EBIT}",
                         font=BLACK_FONT, fmt=FMT_YEN)

            # Gross Profit = Revenue - COGS
            set_cell(ws3, R_GROSS_PROFIT, col, f"={cl}{R_REVENUE}-{cl}{R_COGS}", font=BLACK_FONT, fmt=FMT_YEN,
                     border=SUBTOTAL_BORDER)

            # Gross Margin = GP / Revenue
            set_cell(ws3, R_GROSS_MARGIN, col, f"={cl}{R_GROSS_PROFIT}/{cl}{R_REVENUE}", font=BLACK_FONT, fmt=FMT_PCT)

            # SGA Expense = Revenue * SGA% driver
            set_cell(ws3, R_SGA, col, f"={cl}{R_REVENUE}*{cl}{R_DRV_SGA}", font=BLACK_FONT, fmt=FMT_YEN)

            # EBIT — derived from COGS/SGA overrides when cogs_scenario_rows exists,
            # otherwise linked from Segment Analysis total OP
            if seg_info.get("cogs_scenario_rows"):
                set_cell(ws3, R_EBIT, col,
                         f"={cl}{R_GROSS_PROFIT}-{cl}{R_SGA}",
                         font=BLACK_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER)
            else:
                set_cell(ws3, R_EBIT, col,
                         f"='Segment Analysis'!{seg_cl}{seg_info['total_op_row']}",
                         font=BLACK_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER)

            # Implied Operating Margin = EBIT / Revenue
            set_cell(ws3, R_OP_M_IMPL, col,
                     f"={cl}{R_EBIT}/{cl}{R_REVENUE}",
                     font=BLACK_FONT, fmt=FMT_PCT)

        else:
            # ── Legacy mode: top-down Revenue Growth × Base Year ──

            # Driver rows (CHOOSE formulas referencing scenario matrix)
            set_cell(ws3, R_DRV_GROWTH, col, choose_formula(R_SCEN_BLK_GROWTH, cl),
                     font=BLACK_FONT, fmt=FMT_PCT, fill=LIGHT_FILL)
            set_cell(ws3, R_DRV_COGS, col, choose_formula(R_SCEN_BLK_COGS, cl),
                     font=BLACK_FONT, fmt=FMT_PCT, fill=LIGHT_FILL)
            set_cell(ws3, R_DRV_SGA, col, choose_formula(R_SCEN_BLK_SGA, cl),
                     font=BLACK_FONT, fmt=FMT_PCT, fill=LIGHT_FILL)

            # Revenue — Year 1 grows from Base Year Revenue (latest FY actual, C17)
            if yr == 0:
                set_cell(ws3, R_REVENUE, col, f"=C17*(1+{cl}{R_DRV_GROWTH})", font=BLACK_FONT, fmt=FMT_YEN)
            else:
                set_cell(ws3, R_REVENUE, col, f"={prev_cl}{R_REVENUE}*(1+{cl}{R_DRV_GROWTH})", font=BLACK_FONT, fmt=FMT_YEN)

            # COGS = Revenue * COGS% driver
            set_cell(ws3, R_COGS, col, f"={cl}{R_REVENUE}*{cl}{R_DRV_COGS}", font=BLACK_FONT, fmt=FMT_YEN)

            # Gross Profit = Revenue - COGS
            set_cell(ws3, R_GROSS_PROFIT, col, f"={cl}{R_REVENUE}-{cl}{R_COGS}", font=BLACK_FONT, fmt=FMT_YEN,
                     border=SUBTOTAL_BORDER)

            # Gross Margin = GP / Revenue
            set_cell(ws3, R_GROSS_MARGIN, col, f"={cl}{R_GROSS_PROFIT}/{cl}{R_REVENUE}", font=BLACK_FONT, fmt=FMT_PCT)

            # SGA Expense = Revenue * SGA% driver
            set_cell(ws3, R_SGA, col, f"={cl}{R_REVENUE}*{cl}{R_DRV_SGA}", font=BLACK_FONT, fmt=FMT_YEN)

            # Implied Operating Margin = (GP - SGA) / Revenue
            set_cell(ws3, R_OP_M_IMPL, col,
                     f"=({cl}{R_GROSS_PROFIT}-{cl}{R_SGA})/{cl}{R_REVENUE}",
                     font=BLACK_FONT, fmt=FMT_PCT)

            # EBIT = Gross Profit - SGA
            set_cell(ws3, R_EBIT, col, f"={cl}{R_GROSS_PROFIT}-{cl}{R_SGA}", font=BLACK_FONT, fmt=FMT_YEN,
                     border=SUBTOTAL_BORDER)

        # Tax with NOPAT floor (no tax benefit when EBIT < 0)
        set_cell(ws3, R_TAX, col, f"=MAX(0,{cl}{R_EBIT}*C6)", font=BLACK_FONT, fmt=FMT_YEN)

        # NOPAT
        set_cell(ws3, R_NOPAT, col, f"={cl}{R_EBIT}-{cl}{R_TAX}", font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER)
        # D&A
        if _da_method == "direct":
            _da_arr = C.get("da_direct", {}).get("projections", [])
            _da_val = _da_arr[yr] if yr < len(_da_arr) and _da_arr[yr] is not None else None
            if _da_val is not None:
                set_cell(ws3, R_DA, col, _da_val, font=BLUE_FONT, fmt=FMT_YEN)
            else:
                set_cell(ws3, R_DA, col, f"={cl}{R_REVENUE}*C18", font=BLACK_FONT, fmt=FMT_YEN)
        else:
            set_cell(ws3, R_DA, col, f"={cl}{R_REVENUE}*C18", font=BLACK_FONT, fmt=FMT_YEN)
        # Capex
        if _capex_method == "direct":
            _cx_arr = C.get("capex_direct", {}).get("projections", [])
            _cx_val = _cx_arr[yr] if yr < len(_cx_arr) and _cx_arr[yr] is not None else None
            if _cx_val is not None:
                set_cell(ws3, R_CAPEX, col, _cx_val, font=BLUE_FONT, fmt=FMT_YEN)
            else:
                set_cell(ws3, R_CAPEX, col, f"={cl}{R_REVENUE}*C5", font=BLACK_FONT, fmt=FMT_YEN)
        else:
            set_cell(ws3, R_CAPEX, col, f"={cl}{R_REVENUE}*C5", font=BLACK_FONT, fmt=FMT_YEN)
        # Change in NWC (linked from NWC Schedule; NWC col = DCF col + 1)
        nwc_col_letter = col_letter(col + 1)
        set_cell(ws3, R_CHG_NWC, col,
                 f"='NWC Schedule'!{nwc_col_letter}{actual_nwc_chg_row}",
                 font=BLACK_FONT, fmt=FMT_YEN)
        # UFCF = NOPAT + D&A - Capex - Change in NWC
        set_cell(ws3, R_UFCF, col, f"={cl}{R_NOPAT}+{cl}{R_DA}-{cl}{R_CAPEX}-{cl}{R_CHG_NWC}", font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        # Discount Factor
        set_cell(ws3, R_DISC, col, f"=1/(1+C26)^(C19+{yr})", font=BLACK_FONT, fmt="0.0000")
        # PV of FCF
        set_cell(ws3, R_PV_FCF, col, f"={cl}{R_UFCF}*{cl}{R_DISC}", font=BLACK_FONT, fmt=FMT_YEN,
                 border=SUBTOTAL_BORDER)

    # ── Valuation - Perpetuity Growth Method ──
    c = section_title(ws3, R_PGM_SEC, 2, "Valuation - Perpetuity Growth Method")
    c.fill = LIGHT_GREEN
    for col_idx in range(3, 8):
        ws3.cell(row=R_PGM_SEC, column=col_idx).fill = LIGHT_GREEN

    last_cl = col_letter(3 + proj_years - 1)  # G for 5 years

    set_cell(ws3, R_SUM_PV, 2, "Sum of PV of FCFs", font=BOLD_FONT)
    set_cell(ws3, R_SUM_PV, 3, f"=SUM(C{R_PV_FCF}:{last_cl}{R_PV_FCF})", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_TV_PGM, 2, "Terminal Value (PGM)", font=BOLD_FONT)
    set_cell(ws3, R_TV_PGM, 3, f"={last_cl}{R_UFCF}*(1+C13)/(C26-C13)", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_PV_TV_PGM, 2, "PV of Terminal Value", font=BOLD_FONT)
    set_cell(ws3, R_PV_TV_PGM, 3, f"=C{R_TV_PGM}*{last_cl}{R_DISC}", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_EV_PGM, 2, "Enterprise Value", font=BOLD_FONT)
    set_cell(ws3, R_EV_PGM, 3, f"=C{R_SUM_PV}+C{R_PV_TV_PGM}", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_EQ_PGM, 2, "Equity Value", font=BOLD_FONT)
    set_cell(ws3, R_EQ_PGM, 3, f"=C{R_EV_PGM}-C16", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_PRICE_PGM, 2, "Implied Share Price (PGM)", font=BOLD_FONT)
    set_cell(ws3, R_PRICE_PGM, 3, f"=ROUND(C{R_EQ_PGM}*1000000/C15,0)", font=BLACK_FONT, fmt=FMT_YEN,
             border=TOP_BOTTOM)

    # ── Valuation - Exit Multiple Method ──
    c = section_title(ws3, R_EXIT_SEC, 2, "Valuation - Exit Multiple Method")
    c.fill = LIGHT_GREEN
    for col_idx in range(3, 8):
        ws3.cell(row=R_EXIT_SEC, column=col_idx).fill = LIGHT_GREEN

    set_cell(ws3, R_SUM_PV_EX, 2, "Sum of PV of FCFs", font=BOLD_FONT)
    set_cell(ws3, R_SUM_PV_EX, 3, f"=C{R_SUM_PV}", font=BLACK_FONT, fmt=FMT_YEN)

    # Terminal-value metric: Year-5 Revenue (EV/Sales exit) or Year-5 EBITDA (EV/EBITDA).
    # C14 holds the corresponding multiple; TV = metric × C14 either way.
    if USE_EV_SALES_EXIT:
        set_cell(ws3, R_YR5_EBITDA, 2, "Year 5 Revenue", font=BOLD_FONT)
        set_cell(ws3, R_YR5_EBITDA, 3, f"={last_cl}{R_REVENUE}", font=BLACK_FONT, fmt=FMT_YEN)
    else:
        set_cell(ws3, R_YR5_EBITDA, 2, "Year 5 EBITDA", font=BOLD_FONT)
        set_cell(ws3, R_YR5_EBITDA, 3, f"={last_cl}{R_EBIT}+{last_cl}{R_DA}", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_TV_EXIT, 2, "Terminal Value (Exit Multiple)", font=BOLD_FONT)
    set_cell(ws3, R_TV_EXIT, 3, f"=C{R_YR5_EBITDA}*C14", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_PV_TV_EXIT, 2, "PV of Terminal Value", font=BOLD_FONT)
    set_cell(ws3, R_PV_TV_EXIT, 3, f"=C{R_TV_EXIT}*{last_cl}{R_DISC}", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_EV_EXIT, 2, "Enterprise Value", font=BOLD_FONT)
    set_cell(ws3, R_EV_EXIT, 3, f"=C{R_SUM_PV_EX}+C{R_PV_TV_EXIT}", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_EQ_EXIT, 2, "Equity Value", font=BOLD_FONT)
    set_cell(ws3, R_EQ_EXIT, 3, f"=C{R_EV_EXIT}-C16", font=BLACK_FONT, fmt=FMT_YEN)

    set_cell(ws3, R_PRICE_EXIT, 2, "Implied Share Price (Exit Multiple)", font=BOLD_FONT)
    set_cell(ws3, R_PRICE_EXIT, 3, f"=ROUND(C{R_EQ_EXIT}*1000000/C15,0)", font=BLACK_FONT, fmt=FMT_YEN,
             border=TOP_BOTTOM)

    # ── Scenario Input Matrix ──
    if has_segments:
        # All scenario inputs are in Segment Analysis sheet — show note only
        c = section_title(ws3, R_SCEN_SEC, 2,
                          "All scenario inputs are in Segment Analysis sheet.")
        c.font = GREY_FONT
    else:
        c = section_title(ws3, R_SCEN_SEC, 2, "Scenario Input Matrix")
        c.fill = LIGHT_GREEN
        for col_idx in range(3, 8):
            ws3.cell(row=R_SCEN_SEC, column=col_idx).fill = LIGHT_GREEN

        # Year headers for scenario matrix
        for yr in range(proj_years):
            set_cell(ws3, R_SCEN_YEARS, 3 + yr, f"Year {yr + 1}",
                     font=HEADER_FONT, fill=HEADER_FILL,
                     alignment=Alignment(horizontal="center"))

        driver_blocks = [
            ("Revenue Growth (YoY)", "revenue_growth",    FMT_PCT, R_SCEN_BLK_GROWTH),
            ("COGS % of Revenue",    "cogs_pct",          FMT_PCT, R_SCEN_BLK_COGS),
            ("SGA % of Revenue",     "sga_pct",           FMT_PCT, R_SCEN_BLK_SGA),
        ]

        for drv_label, drv_key, drv_fmt, blk_start in driver_blocks:
            # Sub-header row
            section_title(ws3, blk_start, 2, drv_label)
            # 5 scenario rows
            for s, scen_name in enumerate(SCENARIO_NAMES):
                r = blk_start + 1 + s
                set_cell(ws3, r, 2, scen_name, font=BOLD_FONT)
                scen_data = config["scenarios"][scen_name][drv_key]
                for yr in range(proj_years):
                    set_cell(ws3, r, 3 + yr, scen_data[yr],
                             font=BLUE_FONT, fmt=drv_fmt, border=INPUT_BORDER)

    # =====================================================================
    # SHEET 4: NWC Schedule (DSO/DIH/DPO)
    # =====================================================================
    ws_nwc = wb.create_sheet("NWC Schedule")
    ws_nwc.sheet_properties.tabColor = "CC6600"

    ws_nwc.column_dimensions["A"].width = 3
    ws_nwc.column_dimensions["B"].width = 28
    ws_nwc.column_dimensions["C"].width = 16
    for letter in ["D", "E", "F", "G", "H"]:
        ws_nwc.column_dimensions[letter].width = 16

    set_cell(ws_nwc, 2, 2, f'NWC Schedule - {C["company_name"]}', font=TITLE_FONT)

    # ── Headers: Base Year + FY labels (matching DCF Model sheet) ──
    nwc_proj_labels = [f"Year {y}" for y in range(1, proj_years + 1)]
    if C.get("projection_start_fy"):
        import re as _re
        _m = _re.search(r"FY(\d+)", C["projection_start_fy"])
        if _m:
            _base_fy = int(_m.group(1))
            nwc_proj_labels = [f"FY{_base_fy + y}(E)" for y in range(proj_years)]
    nwc_headers = ["Base Year"] + nwc_proj_labels
    header_row(ws_nwc, 4, 3, 3 + proj_years, nwc_headers)

    # ── Determine NWC method ──
    nwc_method = C.get("nwc_method", "days")

    if nwc_method == "revenue_pct":
        # ================================================================
        # REVENUE PERCENTAGE METHOD
        # Row layout is identical to days method; only cell content differs.
        # ================================================================

        # ── Working Capital Drivers (% of Revenue) ──
        c = section_title(ws_nwc, NWC_R_DSO - 1, 2, "Working Capital Drivers (% of Revenue)")
        c.fill = LIGHT_FILL
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=NWC_R_DSO - 1, column=col_idx).fill = LIGHT_FILL

        set_cell(ws_nwc, NWC_R_DSO, 2, "NWC % of Revenue", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_DIH, 2, "\u2014", font=GREY_FONT)
        set_cell(ws_nwc, NWC_R_DPO, 2, "\u2014", font=GREY_FONT)

        # Base Year NWC%: =C18/C9
        set_cell(ws_nwc, NWC_R_DSO, 3, f"=C{NWC_R_NWC}/C{NWC_R_REV}",
                 font=BLACK_FONT, fmt=FMT_PCT)
        set_cell(ws_nwc, NWC_R_DIH, 3, "\u2014", font=GREY_FONT)
        set_cell(ws_nwc, NWC_R_DPO, 3, "\u2014", font=GREY_FONT)

        # Projected NWC%
        for yr in range(proj_years):
            nwc_col = 4 + yr
            cl = col_letter(nwc_col)
            if has_segments and seg_info and seg_info.get("nwc_scenario_rows"):
                # Reference Segment Analysis Consolidated Inputs NWC% rows
                nwc_seg_cl = col_letter(3 + seg_info["n_hist"] + yr)
                nwc_refs = [f"'Segment Analysis'!{nwc_seg_cl}{r}"
                            for r in seg_info["nwc_scenario_rows"]]
                set_cell(ws_nwc, NWC_R_DSO, nwc_col,
                         f"=CHOOSE('DCF Model'!$D$27,{','.join(nwc_refs)})",
                         font=BLACK_FONT, fmt=FMT_PCT)
            else:
                set_cell(ws_nwc, NWC_R_DSO, nwc_col,
                         nwc_choose_formula(NWC_R_SCEN_BLK_DSO, cl),
                         font=BLACK_FONT, fmt=FMT_PCT)
            set_cell(ws_nwc, NWC_R_DIH, nwc_col, "\u2014", font=GREY_FONT)
            set_cell(ws_nwc, NWC_R_DPO, nwc_col, "\u2014", font=GREY_FONT)

        # ── Revenue & COGS ──
        c = section_title(ws_nwc, NWC_R_REV - 1, 2, "P&L Reference (JPY mn)")
        c.fill = LIGHT_FILL
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=NWC_R_REV - 1, column=col_idx).fill = LIGHT_FILL

        set_cell(ws_nwc, NWC_R_REV, 2, "Revenue", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_COGS, 2, "\u2014", font=GREY_FONT)

        # Base Year Revenue
        set_cell(ws_nwc, NWC_R_REV, 3, C["base_year_revenue"], font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws_nwc, NWC_R_COGS, 3, "\u2014", font=GREY_FONT)

        # Projected Revenue (linked to DCF Model)
        for yr in range(proj_years):
            nwc_col = 4 + yr
            dcf_col_letter = col_letter(3 + yr)
            set_cell(ws_nwc, NWC_R_REV, nwc_col,
                     f"='DCF Model'!{dcf_col_letter}{R_REVENUE}",
                     font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws_nwc, NWC_R_COGS, nwc_col, "\u2014", font=GREY_FONT)

        # ── Working Capital Items: all show "—" ──
        c = section_title(ws_nwc, NWC_R_AR - 1, 2, "Working Capital Items (JPY mn)")
        c.fill = LIGHT_FILL
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=NWC_R_AR - 1, column=col_idx).fill = LIGHT_FILL

        for _row in [NWC_R_AR, NWC_R_INV, NWC_R_CA, NWC_R_AP, NWC_R_CL]:
            set_cell(ws_nwc, _row, 2, "\u2014", font=GREY_FONT)
            for _ci in range(3, 3 + proj_years + 1):
                set_cell(ws_nwc, _row, _ci, "\u2014", font=GREY_FONT)

        # ── NWC Summary ──
        c = section_title(ws_nwc, NWC_R_NWC - 1, 2, "Net Working Capital (JPY mn)")
        c.fill = LIGHT_GREEN
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=NWC_R_NWC - 1, column=col_idx).fill = LIGHT_GREEN

        set_cell(ws_nwc, NWC_R_NWC, 2, "Net Working Capital", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_CHG_NWC, 2, "Change in NWC", font=BOLD_FONT)

        # Base Year NWC: hardcoded from computed value
        set_cell(ws_nwc, NWC_R_NWC, 3, C.get("base_year_nwc", 0),
                 font=BLUE_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        set_cell(ws_nwc, NWC_R_CHG_NWC, 3, "n/a", font=BLACK_FONT)

        # Projected NWC = Revenue × NWC%, Change = ΔNWC
        for yr in range(proj_years):
            nwc_col = 4 + yr
            cl = col_letter(nwc_col)
            prev_cl = col_letter(nwc_col - 1)
            set_cell(ws_nwc, NWC_R_NWC, nwc_col,
                     f"={cl}{NWC_R_REV}*{cl}{NWC_R_DSO}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
            set_cell(ws_nwc, NWC_R_CHG_NWC, nwc_col,
                     f"={cl}{NWC_R_NWC}-{prev_cl}{NWC_R_NWC}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=TOP_BOTTOM)

        # ── Apply consistent borders ──
        _nwc_data_rows = (
            list(range(NWC_R_DSO, NWC_R_DPO + 1))
            + list(range(NWC_R_REV, NWC_R_COGS + 1))
            + list(range(NWC_R_AR, NWC_R_CL + 1))
            + [NWC_R_NWC, NWC_R_CHG_NWC]
        )
        _nwc_col_end = 3 + proj_years
        for _r in _nwc_data_rows:
            for _ci in range(2, _nwc_col_end + 1):
                _cell = ws_nwc.cell(row=_r, column=_ci)
                _has_border = (_cell.border and any([
                    getattr(_cell.border.top, 'style', None),
                    getattr(_cell.border.bottom, 'style', None),
                    getattr(_cell.border.left, 'style', None),
                    getattr(_cell.border.right, 'style', None),
                ]))
                if not _has_border:
                    _cell.border = NWC_DATA_BORDER

        # ── Scenario Input Matrix: NWC % of Revenue ──
        if has_segments and seg_info and seg_info.get("nwc_scenario_rows"):
            # NWC inputs are in Segment Analysis Consolidated Inputs
            c = section_title(ws_nwc, NWC_R_SCEN_SEC, 2,
                              "NWC inputs are in Segment Analysis sheet.")
            c.font = GREY_FONT
        else:
            c = section_title(ws_nwc, NWC_R_SCEN_SEC, 2,
                              "Scenario Input Matrix (NWC % of Revenue)")
            c.fill = LIGHT_GREEN
            for col_idx in range(3, 3 + proj_years + 1):
                ws_nwc.cell(row=NWC_R_SCEN_SEC, column=col_idx).fill = LIGHT_GREEN

            for yr in range(proj_years):
                _scen_yr_label = nwc_proj_labels[yr] if yr < len(nwc_proj_labels) else f"Year {yr + 1}"
                set_cell(ws_nwc, NWC_R_SCEN_YEARS, 4 + yr, _scen_yr_label,
                         font=HEADER_FONT, fill=HEADER_FILL,
                         alignment=Alignment(horizontal="center"))

            # Single block: NWC % of Revenue at NWC_R_SCEN_BLK_DSO (row 24)
            section_title(ws_nwc, NWC_R_SCEN_BLK_DSO, 2, "NWC % of Revenue")
            for s, scen_name in enumerate(SCENARIO_NAMES):
                r = NWC_R_SCEN_BLK_DSO + 1 + s
                set_cell(ws_nwc, r, 2, scen_name, font=BOLD_FONT)
                scen_data = config["scenarios"][scen_name]["nwc_pct"]
                for yr in range(proj_years):
                    set_cell(ws_nwc, r, 4 + yr, scen_data[yr],
                             font=BLUE_FONT, fmt=FMT_PCT, border=INPUT_BORDER)

    elif nwc_method == "itemized":
        # ================================================================
        # ITEMIZED METHOD — 7+ individual BS items with turnover days
        # ================================================================
        nwc_items = C.get("nwc_items", [])
        asset_items = [it for it in nwc_items if it["side"] == "asset"]
        liab_items = [it for it in nwc_items if it["side"] == "liability"]
        n_total = len(nwc_items)
        n_assets = len(asset_items)
        n_liabs = len(liab_items)

        # ── Compute dynamic row positions ──
        ITM_DRV_START = 5
        ITM_PNL_HDR = ITM_DRV_START + n_total + 1
        ITM_REV = ITM_PNL_HDR + 1
        ITM_COGS = ITM_PNL_HDR + 2
        ITM_WC_HDR = ITM_COGS + 2
        ITM_FIRST_ASSET = ITM_WC_HDR + 1
        ITM_CA_TOTAL = ITM_FIRST_ASSET + n_assets
        ITM_FIRST_LIAB = ITM_CA_TOTAL + 1
        ITM_CL_TOTAL = ITM_FIRST_LIAB + n_liabs
        ITM_NWC_HDR = ITM_CL_TOTAL + 2
        ITM_NWC = ITM_NWC_HDR + 1
        ITM_CHG = ITM_NWC + 1
        actual_nwc_chg_row = ITM_CHG

        # Map each item to its driver row and value row
        drv_rows = {}  # scenario_key -> driver row
        val_rows = {}  # scenario_key -> value row
        for i, item in enumerate(nwc_items):
            drv_rows[item["scenario_key"]] = ITM_DRV_START + i

        asset_val_rows = []
        for i in range(n_assets):
            r = ITM_FIRST_ASSET + i
            asset_val_rows.append(r)
            val_rows[asset_items[i]["scenario_key"]] = r

        liab_val_rows = []
        for i in range(n_liabs):
            r = ITM_FIRST_LIAB + i
            liab_val_rows.append(r)
            val_rows[liab_items[i]["scenario_key"]] = r

        # ── Scenario block row positions ──
        ITM_SCEN_SEC = ITM_CHG + 3
        ITM_SCEN_YEARS = ITM_SCEN_SEC + 1
        scen_block_start = {}
        cur_blk = ITM_SCEN_YEARS + 1
        for item in nwc_items:
            scen_block_start[item["scenario_key"]] = cur_blk
            cur_blk += 7  # header(1) + 5 scenarios + blank(1)

        def itm_choose(block_start, cl):
            refs = [f"{cl}{block_start + 1 + s}" for s in range(NUM_SCENARIOS)]
            return f"=CHOOSE('DCF Model'!$D$27,{','.join(refs)})"

        # ── Driver section header ──
        c = section_title(ws_nwc, ITM_DRV_START - 1, 2, "Working Capital Drivers (Days)")
        c.fill = LIGHT_FILL
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=ITM_DRV_START - 1, column=col_idx).fill = LIGHT_FILL

        # ── Driver rows (turnover days) ──
        for i, item in enumerate(nwc_items):
            r = ITM_DRV_START + i
            set_cell(ws_nwc, r, 2, item["label"] + " Days", font=BOLD_FONT)
            denom_row = ITM_REV if item["denom"] == "revenue" else ITM_COGS
            # Base Year: = base_value / denominator * 365
            set_cell(ws_nwc, r, 3,
                     f"=C{val_rows[item['scenario_key']]}/C{denom_row}*365",
                     font=BLACK_FONT, fmt=FMT_DAYS)
            # Projected: CHOOSE from scenario matrix
            for yr in range(proj_years):
                nwc_col = 4 + yr
                cl = col_letter(nwc_col)
                set_cell(ws_nwc, r, nwc_col,
                         itm_choose(scen_block_start[item["scenario_key"]], cl),
                         font=BLACK_FONT, fmt=FMT_DAYS)

        # ── P&L Reference section ──
        c = section_title(ws_nwc, ITM_PNL_HDR, 2, "P&L Reference (JPY mn)")
        c.fill = LIGHT_FILL
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=ITM_PNL_HDR, column=col_idx).fill = LIGHT_FILL

        set_cell(ws_nwc, ITM_REV, 2, "Revenue", font=BOLD_FONT)
        set_cell(ws_nwc, ITM_COGS, 2, "COGS", font=BOLD_FONT)
        set_cell(ws_nwc, ITM_REV, 3, C["base_year_revenue"], font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws_nwc, ITM_COGS, 3, C["base_year_cogs"], font=BLUE_FONT, fmt=FMT_YEN)
        for yr in range(proj_years):
            nwc_col = 4 + yr
            dcf_cl = col_letter(3 + yr)
            set_cell(ws_nwc, ITM_REV, nwc_col,
                     f"='DCF Model'!{dcf_cl}{R_REVENUE}", font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws_nwc, ITM_COGS, nwc_col,
                     f"='DCF Model'!{dcf_cl}{R_COGS}", font=BLACK_FONT, fmt=FMT_YEN)

        # ── Working Capital Items section ──
        c = section_title(ws_nwc, ITM_WC_HDR, 2, "Working Capital Items (JPY mn)")
        c.fill = LIGHT_FILL
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=ITM_WC_HDR, column=col_idx).fill = LIGHT_FILL

        # Asset items
        for i, item in enumerate(asset_items):
            r = asset_val_rows[i]
            drv_r = drv_rows[item["scenario_key"]]
            denom_row = ITM_REV if item["denom"] == "revenue" else ITM_COGS
            set_cell(ws_nwc, r, 2, item["label"], font=BOLD_FONT)
            set_cell(ws_nwc, r, 3, item["base_value"], font=BLUE_FONT, fmt=FMT_YEN)
            for yr in range(proj_years):
                nwc_col = 4 + yr
                cl = col_letter(nwc_col)
                set_cell(ws_nwc, r, nwc_col,
                         f"={cl}{denom_row}*{cl}{drv_r}/365",
                         font=BLACK_FONT, fmt=FMT_YEN)

        # Total Current Assets
        set_cell(ws_nwc, ITM_CA_TOTAL, 2, "Total Current Assets", font=BOLD_FONT)
        first_a = asset_val_rows[0]
        last_a = asset_val_rows[-1]
        for ci in range(3, 3 + proj_years + 1):
            cl = col_letter(ci)
            set_cell(ws_nwc, ITM_CA_TOTAL, ci,
                     f"=SUM({cl}{first_a}:{cl}{last_a})",
                     font=BLACK_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER)

        # Liability items
        for i, item in enumerate(liab_items):
            r = liab_val_rows[i]
            drv_r = drv_rows[item["scenario_key"]]
            denom_row = ITM_REV if item["denom"] == "revenue" else ITM_COGS
            set_cell(ws_nwc, r, 2, item["label"], font=BOLD_FONT)
            set_cell(ws_nwc, r, 3, item["base_value"], font=BLUE_FONT, fmt=FMT_YEN)
            for yr in range(proj_years):
                nwc_col = 4 + yr
                cl = col_letter(nwc_col)
                set_cell(ws_nwc, r, nwc_col,
                         f"={cl}{denom_row}*{cl}{drv_r}/365",
                         font=BLACK_FONT, fmt=FMT_YEN)

        # Total Current Liabilities
        set_cell(ws_nwc, ITM_CL_TOTAL, 2, "Total Current Liabilities", font=BOLD_FONT)
        first_l = liab_val_rows[0]
        last_l = liab_val_rows[-1]
        for ci in range(3, 3 + proj_years + 1):
            cl = col_letter(ci)
            set_cell(ws_nwc, ITM_CL_TOTAL, ci,
                     f"=SUM({cl}{first_l}:{cl}{last_l})",
                     font=BLACK_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER)

        # ── NWC Summary ──
        c = section_title(ws_nwc, ITM_NWC_HDR, 2, "Net Working Capital (JPY mn)")
        c.fill = LIGHT_GREEN
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=ITM_NWC_HDR, column=col_idx).fill = LIGHT_GREEN

        set_cell(ws_nwc, ITM_NWC, 2, "Net Working Capital", font=BOLD_FONT)
        set_cell(ws_nwc, ITM_CHG, 2, "Change in NWC", font=BOLD_FONT)

        # Base Year NWC = CA - CL
        set_cell(ws_nwc, ITM_NWC, 3,
                 f"=C{ITM_CA_TOTAL}-C{ITM_CL_TOTAL}",
                 font=BLACK_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        set_cell(ws_nwc, ITM_CHG, 3, "n/a", font=BLACK_FONT)

        # Projected NWC & Change
        for yr in range(proj_years):
            nwc_col = 4 + yr
            cl = col_letter(nwc_col)
            prev_cl = col_letter(nwc_col - 1)
            set_cell(ws_nwc, ITM_NWC, nwc_col,
                     f"={cl}{ITM_CA_TOTAL}-{cl}{ITM_CL_TOTAL}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
            set_cell(ws_nwc, ITM_CHG, nwc_col,
                     f"={cl}{ITM_NWC}-{prev_cl}{ITM_NWC}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=TOP_BOTTOM)

        # ── Apply borders to all data rows ──
        _itm_data_rows = (
            list(range(ITM_DRV_START, ITM_DRV_START + n_total))
            + [ITM_REV, ITM_COGS]
            + asset_val_rows + [ITM_CA_TOTAL]
            + liab_val_rows + [ITM_CL_TOTAL]
            + [ITM_NWC, ITM_CHG]
        )
        for _r in _itm_data_rows:
            for _ci in range(2, 3 + proj_years + 1):
                _cell = ws_nwc.cell(row=_r, column=_ci)
                _has_border = (_cell.border and any([
                    getattr(_cell.border.top, 'style', None),
                    getattr(_cell.border.bottom, 'style', None),
                    getattr(_cell.border.left, 'style', None),
                    getattr(_cell.border.right, 'style', None),
                ]))
                if not _has_border:
                    _cell.border = NWC_DATA_BORDER

        # ── Scenario Input Matrix ──
        c = section_title(ws_nwc, ITM_SCEN_SEC, 2,
                          "Scenario Input Matrix (Working Capital Days)")
        c.fill = LIGHT_GREEN
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=ITM_SCEN_SEC, column=col_idx).fill = LIGHT_GREEN

        for yr in range(proj_years):
            _lbl = nwc_proj_labels[yr] if yr < len(nwc_proj_labels) else f"Year {yr + 1}"
            set_cell(ws_nwc, ITM_SCEN_YEARS, 4 + yr, _lbl,
                     font=HEADER_FONT, fill=HEADER_FILL,
                     alignment=Alignment(horizontal="center"))

        for item in nwc_items:
            blk = scen_block_start[item["scenario_key"]]
            section_title(ws_nwc, blk, 2, item["label"] + " Days")
            for s, scen_name in enumerate(SCENARIO_NAMES):
                r = blk + 1 + s
                set_cell(ws_nwc, r, 2, scen_name, font=BOLD_FONT)
                scen_data = config["scenarios"][scen_name].get(
                    item["scenario_key"], [0] * proj_years)
                for yr in range(proj_years):
                    set_cell(ws_nwc, r, 4 + yr, scen_data[yr],
                             font=BLUE_FONT, fmt=FMT_DAYS, border=INPUT_BORDER)

    else:
        # ================================================================
        # DAYS METHOD (default) — existing code, unchanged
        # ================================================================

        # ── Working Capital Drivers ──
        c = section_title(ws_nwc, NWC_R_DSO - 1, 2, "Working Capital Drivers (Days)")
        c.fill = LIGHT_FILL
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=NWC_R_DSO - 1, column=col_idx).fill = LIGHT_FILL

        set_cell(ws_nwc, NWC_R_DSO, 2, "DSO (Days Sales Outstanding)", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_DIH, 2, "DIH (Days Inventory Held)", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_DPO, 2, "DPO (Days Payable Outstanding)", font=BOLD_FONT)

        # Base Year DSO/DIH/DPO (computed from actuals)
        set_cell(ws_nwc, NWC_R_DSO, 3, f"=C{NWC_R_AR}/C{NWC_R_REV}*365",
                 font=BLACK_FONT, fmt=FMT_DAYS)
        set_cell(ws_nwc, NWC_R_DIH, 3, f"=C{NWC_R_INV}/C{NWC_R_COGS}*365",
                 font=BLACK_FONT, fmt=FMT_DAYS)
        set_cell(ws_nwc, NWC_R_DPO, 3, f"=C{NWC_R_AP}/C{NWC_R_COGS}*365",
                 font=BLACK_FONT, fmt=FMT_DAYS)

        # Projected DSO/DIH/DPO (CHOOSE from scenario matrix)
        for yr in range(proj_years):
            nwc_col = 4 + yr
            cl = col_letter(nwc_col)
            set_cell(ws_nwc, NWC_R_DSO, nwc_col,
                     nwc_choose_formula(NWC_R_SCEN_BLK_DSO, cl),
                     font=BLACK_FONT, fmt=FMT_DAYS)
            set_cell(ws_nwc, NWC_R_DIH, nwc_col,
                     nwc_choose_formula(NWC_R_SCEN_BLK_DIH, cl),
                     font=BLACK_FONT, fmt=FMT_DAYS)
            set_cell(ws_nwc, NWC_R_DPO, nwc_col,
                     nwc_choose_formula(NWC_R_SCEN_BLK_DPO, cl),
                     font=BLACK_FONT, fmt=FMT_DAYS)

        # ── Revenue & COGS (linked from DCF Model) ──
        c = section_title(ws_nwc, NWC_R_REV - 1, 2, "P&L Reference (JPY mn)")
        c.fill = LIGHT_FILL
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=NWC_R_REV - 1, column=col_idx).fill = LIGHT_FILL

        set_cell(ws_nwc, NWC_R_REV, 2, "Revenue", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_COGS, 2, "COGS", font=BOLD_FONT)

        # Base Year
        set_cell(ws_nwc, NWC_R_REV, 3, C["base_year_revenue"], font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws_nwc, NWC_R_COGS, 3, C["base_year_cogs"], font=BLUE_FONT, fmt=FMT_YEN)

        # Projected (linked to DCF Model; NWC col D = DCF col C, offset +1)
        for yr in range(proj_years):
            nwc_col = 4 + yr
            dcf_col_letter = col_letter(3 + yr)
            set_cell(ws_nwc, NWC_R_REV, nwc_col,
                     f"='DCF Model'!{dcf_col_letter}{R_REVENUE}",
                     font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws_nwc, NWC_R_COGS, nwc_col,
                     f"='DCF Model'!{dcf_col_letter}{R_COGS}",
                     font=BLACK_FONT, fmt=FMT_YEN)

        # ── Working Capital Items ──
        c = section_title(ws_nwc, NWC_R_AR - 1, 2, "Working Capital Items (JPY mn)")
        c.fill = LIGHT_FILL
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=NWC_R_AR - 1, column=col_idx).fill = LIGHT_FILL

        set_cell(ws_nwc, NWC_R_AR, 2, "Accounts Receivable", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_INV, 2, "Inventory", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_CA, 2, "Current Assets (AR + Inv)", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_AP, 2, "Accounts Payable", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_CL, 2, "Current Liabilities (AP)", font=BOLD_FONT)

        # Base Year actuals
        set_cell(ws_nwc, NWC_R_AR, 3, C["base_year_ar"], font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws_nwc, NWC_R_INV, 3, C["base_year_inv"], font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws_nwc, NWC_R_CA, 3, f"=C{NWC_R_AR}+C{NWC_R_INV}", font=BLACK_FONT, fmt=FMT_YEN)
        set_cell(ws_nwc, NWC_R_AP, 3, C["base_year_ap"], font=BLUE_FONT, fmt=FMT_YEN)
        set_cell(ws_nwc, NWC_R_CL, 3, f"=C{NWC_R_AP}", font=BLACK_FONT, fmt=FMT_YEN)

        # Projected WC items
        for yr in range(proj_years):
            nwc_col = 4 + yr
            cl = col_letter(nwc_col)
            set_cell(ws_nwc, NWC_R_AR, nwc_col,
                     f"={cl}{NWC_R_REV}*{cl}{NWC_R_DSO}/365", font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws_nwc, NWC_R_INV, nwc_col,
                     f"={cl}{NWC_R_COGS}*{cl}{NWC_R_DIH}/365", font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws_nwc, NWC_R_CA, nwc_col,
                     f"={cl}{NWC_R_AR}+{cl}{NWC_R_INV}", font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws_nwc, NWC_R_AP, nwc_col,
                     f"={cl}{NWC_R_COGS}*{cl}{NWC_R_DPO}/365", font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws_nwc, NWC_R_CL, nwc_col,
                     f"={cl}{NWC_R_AP}", font=BLACK_FONT, fmt=FMT_YEN)

        # ── NWC Summary ──
        c = section_title(ws_nwc, NWC_R_NWC - 1, 2, "Net Working Capital (JPY mn)")
        c.fill = LIGHT_GREEN
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=NWC_R_NWC - 1, column=col_idx).fill = LIGHT_GREEN

        set_cell(ws_nwc, NWC_R_NWC, 2, "Net Working Capital", font=BOLD_FONT)
        set_cell(ws_nwc, NWC_R_CHG_NWC, 2, "Change in NWC", font=BOLD_FONT)

        # Base Year NWC
        set_cell(ws_nwc, NWC_R_NWC, 3, f"=C{NWC_R_CA}-C{NWC_R_CL}",
                 font=BLACK_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
        set_cell(ws_nwc, NWC_R_CHG_NWC, 3, "n/a", font=BLACK_FONT)

        # Projected NWC & Change
        for yr in range(proj_years):
            nwc_col = 4 + yr
            cl = col_letter(nwc_col)
            prev_cl = col_letter(nwc_col - 1)
            set_cell(ws_nwc, NWC_R_NWC, nwc_col,
                     f"={cl}{NWC_R_CA}-{cl}{NWC_R_CL}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=SUBTOTAL_BORDER, fill=SUBTOTAL_FILL)
            set_cell(ws_nwc, NWC_R_CHG_NWC, nwc_col,
                     f"={cl}{NWC_R_NWC}-{prev_cl}{NWC_R_NWC}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=TOP_BOTTOM)

        # ── Apply consistent borders to all NWC data rows ──
        _nwc_data_rows = (
            list(range(NWC_R_DSO, NWC_R_DPO + 1))       # Drivers: DSO, DIH, DPO
            + list(range(NWC_R_REV, NWC_R_COGS + 1))     # P&L Reference: Revenue, COGS
            + list(range(NWC_R_AR, NWC_R_CL + 1))         # WC Items: AR, Inv, CA, AP, CL
            + [NWC_R_NWC, NWC_R_CHG_NWC]                  # NWC Summary
        )
        _nwc_col_end = 3 + proj_years  # last data column
        for _r in _nwc_data_rows:
            for _ci in range(2, _nwc_col_end + 1):
                _cell = ws_nwc.cell(row=_r, column=_ci)
                # Preserve existing meaningful borders (SUBTOTAL_BORDER, TOP_BOTTOM, etc.)
                _has_border = (_cell.border and any([
                    getattr(_cell.border.top, 'style', None),
                    getattr(_cell.border.bottom, 'style', None),
                    getattr(_cell.border.left, 'style', None),
                    getattr(_cell.border.right, 'style', None),
                ]))
                if not _has_border:
                    _cell.border = NWC_DATA_BORDER

        # ── Scenario Input Matrix (DSO, DIH, DPO) ──
        c = section_title(ws_nwc, NWC_R_SCEN_SEC, 2, "Scenario Input Matrix (Working Capital Days)")
        c.fill = LIGHT_GREEN
        for col_idx in range(3, 3 + proj_years + 1):
            ws_nwc.cell(row=NWC_R_SCEN_SEC, column=col_idx).fill = LIGHT_GREEN

        for yr in range(proj_years):
            _scen_yr_label = nwc_proj_labels[yr] if yr < len(nwc_proj_labels) else f"Year {yr + 1}"
            set_cell(ws_nwc, NWC_R_SCEN_YEARS, 4 + yr, _scen_yr_label,
                     font=HEADER_FONT, fill=HEADER_FILL,
                     alignment=Alignment(horizontal="center"))

        nwc_driver_blocks = [
            ("DSO (Days)", "dso_days", FMT_DAYS, NWC_R_SCEN_BLK_DSO),
            ("DIH (Days)", "dih_days", FMT_DAYS, NWC_R_SCEN_BLK_DIH),
            ("DPO (Days)", "dpo_days", FMT_DAYS, NWC_R_SCEN_BLK_DPO),
        ]

        for drv_label, drv_key, drv_fmt, blk_start in nwc_driver_blocks:
            section_title(ws_nwc, blk_start, 2, drv_label)
            for s, scen_name in enumerate(SCENARIO_NAMES):
                r = blk_start + 1 + s
                set_cell(ws_nwc, r, 2, scen_name, font=BOLD_FONT)
                scen_data = config["scenarios"][scen_name][drv_key]
                for yr in range(proj_years):
                    set_cell(ws_nwc, r, 4 + yr, scen_data[yr],
                             font=BLUE_FONT, fmt=drv_fmt, border=INPUT_BORDER)

    # =====================================================================
    # SHEET 5: Comps Analysis
    # =====================================================================
    ws4 = wb.create_sheet("Comps Analysis")
    ws4.sheet_properties.tabColor = "006600"

    ws4.column_dimensions["A"].width = 3
    ws4.column_dimensions["B"].width = 16
    ws4.column_dimensions["C"].width = 10
    for letter in ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]:
        ws4.column_dimensions[letter].width = 12
    ws4.column_dimensions["Q"].width = 26

    set_cell(ws4, 2, 2, "Comparable Company Analysis", font=TITLE_FONT)

    comps = C["comps"]

    # ── Peer data-quality screens (applied to peers only) ─────────────────
    # 1. D&A missing: a peer whose EBITDA equals its operating income has no
    #    depreciation added back. Its EV/EBITDA is really EV/EBIT and inflates
    #    the median (8410: implied roughly doubled). The row stays for display,
    #    but the EBITDA cell is blanked so every EV/EBITDA statistic skips it.
    # 2. Stale/delisted: flagged upstream by generate_dcf.py's peer freshness
    #    check (comp["exclude_from_stats"]) — excluded from ALL statistics.
    comp_flags = []
    for i, comp in enumerate(comps):
        is_subject = (i == subject_idx)
        eb = comp.get("ebitda")
        oi = comp.get("op_income")
        da_missing = not is_subject and (
            eb is None
            or (isinstance(eb, (int, float)) and isinstance(oi, (int, float))
                and eb == oi and eb > 0)
        )
        stale = bool(comp.get("exclude_from_stats")) and not is_subject
        notes = []
        if da_missing:
            notes.append("(D&A n/a)")
        if stale:
            notes.append(comp.get("exclude_reason") or "(stale/delisted)")
        comp_flags.append({
            "is_subject": is_subject,
            "da_missing": da_missing,
            "stale": stale,
            "note": " ".join(notes),
        })

    _da_missing_names = [c["name"] for c, f in zip(comps, comp_flags) if f["da_missing"]]
    _stale_names = [c["name"] for c, f in zip(comps, comp_flags) if f["stale"]]
    if _da_missing_names:
        print(f"  [Comps] WARNING: EBITDA == Operating Income (D&A not added back) for: "
              f"{', '.join(_da_missing_names)} - excluded from EV/EBITDA statistics.")
    if _stale_names:
        print(f"  [Comps] WARNING: excluded from all statistics (stale/delisted): "
              f"{', '.join(_stale_names)}")

    # As-of note for the static peer market caps (they are point-in-time values)
    _asof = _dt.datetime.now().strftime("%Y-%m-%d")
    set_cell(ws4, 3, 2,
             f"Peer prices as of {_asof} (Mkt Cap / EV of peer rows are static CSV "
             f"values; the subject row is live-linked to Executive Summary)",
             font=GREY_FONT)

    # Header row
    comp_headers = [
        "Company", "Ticker", "Mkt Cap\n(JPY mn)", "EV\n(JPY mn)",
        "Revenue\n(JPY mn)", "EBITDA\n(JPY mn)", "Op Income\n(JPY mn)",
        "Net Income\n(JPY mn)", "EV/EBITDA", "EV/Revenue", "PER",
        "PBR", "Op Margin", "ROE", "Book Value\n(JPY mn)", "Note"
    ]
    header_row(ws4, 4, 2, 17, comp_headers)

    _no_book_value = []

    # Company data rows (row 5 onward, one per CSV row)
    for i, comp in enumerate(comps):
        r = 5 + i
        flags = comp_flags[i]

        _na = lambda ws, r, c: set_cell(ws, r, c, "N/A", font=BLACK_FONT, border=THIN_BORDER,
                                         alignment=Alignment(horizontal="right"))

        set_cell(ws4, r, 2, comp["name"], font=BOLD_FONT)
        set_cell(ws4, r, 3, comp["ticker"])

        # Mkt Cap & EV.
        # Subject row: computed from the SAME price/share count the rest of the
        # model uses, so a price update flows through instead of leaving a stale
        # hardcoded market cap behind (285A/8410 used to be patched by hand).
        if flags["is_subject"]:
            set_cell(ws4, r, 4, f"='Executive Summary'!C9*C{R_CMP_SHARES}/1000000",
                     font=BLACK_FONT, fmt=FMT_YEN, border=THIN_BORDER)
            set_cell(ws4, r, 5, f"=D{r}+C{R_CMP_NETDEBT}",
                     font=BLACK_FONT, fmt=FMT_YEN, border=THIN_BORDER)
        else:
            if comp["mkt_cap"] is None:
                _na(ws4, r, 4)
            else:
                set_cell(ws4, r, 4, comp["mkt_cap"], font=BLUE_FONT, fmt=FMT_YEN, border=THIN_BORDER)
            if comp["ev"] is None:
                _na(ws4, r, 5)
            else:
                set_cell(ws4, r, 5, comp["ev"], font=BLUE_FONT, fmt=FMT_YEN, border=THIN_BORDER)

        set_cell(ws4, r, 6, comp["revenue"], font=BLUE_FONT, fmt=FMT_YEN, border=THIN_BORDER)
        if flags["da_missing"]:
            # Blank, not zero: MEDIAN/PERCENTILE skip empty cells but not zeros.
            set_cell(ws4, r, 7, None, border=THIN_BORDER)
        else:
            set_cell(ws4, r, 7, comp["ebitda"], font=BLUE_FONT, fmt=FMT_YEN, border=THIN_BORDER)
        set_cell(ws4, r, 8, comp["op_income"], font=BLUE_FONT, fmt=FMT_YEN, border=THIN_BORDER)
        set_cell(ws4, r, 9, comp["net_income"], font=BLUE_FONT, fmt=FMT_YEN, border=THIN_BORDER)

        # Book Value (JPY mn): subject prefers config["book_value"], else the CSV row
        bv = None
        if flags["is_subject"] and isinstance(C.get("book_value"), (int, float)):
            bv = C["book_value"]
        elif isinstance(comp.get("book_value"), (int, float)) and comp["book_value"] != 0:
            bv = comp["book_value"]
        if bv is None:
            _na(ws4, r, 16)
            _no_book_value.append(comp["name"])
        else:
            set_cell(ws4, r, 16, bv, font=BLUE_FONT, fmt=FMT_YEN, border=THIN_BORDER)

        # EV/EBITDA
        if flags["da_missing"] or comp["ev"] is None or not comp["ebitda"] or comp["ebitda"] <= 0:
            _na(ws4, r, 10)
        else:
            set_cell(ws4, r, 10, f"=IFERROR(E{r}/G{r},\"\")", font=BLACK_FONT, fmt=FMT_RATIO,
                     border=THIN_BORDER)

        # EV/Revenue
        if comp["ev"] is None or comp["revenue"] <= 0:
            _na(ws4, r, 11)
        else:
            set_cell(ws4, r, 11, f"=IFERROR(E{r}/F{r},\"\")", font=BLACK_FONT, fmt=FMT_RATIO,
                     border=THIN_BORDER)

        # PER
        if comp["mkt_cap"] is None or comp["net_income"] <= 0:
            _na(ws4, r, 12)
        else:
            set_cell(ws4, r, 12, f"=IFERROR(D{r}/I{r},\"\")", font=BLACK_FONT, fmt=FMT_RATIO,
                     border=THIN_BORDER)

        # PBR / ROE: formulas off the Book Value column, so they follow a
        # market-cap or book-value correction instead of freezing at load time.
        if bv is None:
            _na(ws4, r, 13)
            _na(ws4, r, 15)
        else:
            set_cell(ws4, r, 13, f"=IFERROR(D{r}/P{r},\"\")", font=BLACK_FONT, fmt=FMT_RATIO,
                     border=THIN_BORDER)
            set_cell(ws4, r, 15, f"=IFERROR(I{r}/P{r},\"\")", font=BLACK_FONT, fmt=FMT_PCT,
                     border=THIN_BORDER)

        set_cell(ws4, r, 14, f"=IFERROR(H{r}/F{r},\"\")", font=BLACK_FONT, fmt=FMT_PCT,
                 border=THIN_BORDER)

        if flags["note"]:
            set_cell(ws4, r, 17, flags["note"], font=GREY_FONT)

    if _no_book_value:
        print(f"  [Comps] WARNING: no Book Value for {', '.join(_no_book_value)} - "
              f"PBR/ROE left blank for those rows (add Book_Value to the comps CSV).")

    last_comp_row = 5 + len(comps) - 1

    # ── Statistics ──
    # Rows kept at 14-17 for backward compatibility (Executive Summary and
    # market_analysis_template read C27/C28 below); shifted down only when a
    # large comps set would otherwise overlap them.
    stat_sec_row = R_STAT_SEC
    section_title(ws4, stat_sec_row, 2, "Statistics")

    stat_labels = ["25th Percentile", "Median (50th)", "75th Percentile"]
    stat_rows = [stat_sec_row + 1, stat_sec_row + 2, stat_sec_row + 3]
    assert stat_rows[1] == R_STAT_MEDIAN

    stat_col_map = [
        (4, 10),  # EV/EBITDA
        (5, 11),  # EV/Revenue
        (6, 12),  # PER
        (7, 13),  # PBR
        (8, 14),  # Op Margin
        (9, 15),  # ROE
    ]

    # Peer rows = every data row except the subject's own and any row excluded
    # by the freshness screen. Derived from the comps list — no fixed addresses.
    peer_rows = [5 + i for i, f in enumerate(comp_flags)
                 if not f["is_subject"] and not f["stale"]]
    ev_ebitda_rows = [5 + i for i, f in enumerate(comp_flags)
                      if not f["is_subject"] and not f["stale"] and not f["da_missing"]
                      and comps[i].get("ev") is not None
                      and (comps[i].get("ebitda") or 0) > 0]
    n_ev_ebitda = len(ev_ebitda_rows)
    EV_EBITDA_INVALID = (not USE_EV_SALES) and n_ev_ebitda < 3
    if EV_EBITDA_INVALID:
        print(f"  [Comps] WARNING: only {n_ev_ebitda} peer(s) with usable EBITDA "
              f"(<3) - EV/EBITDA implied price marked INVALID and excluded from "
              f"the Target Price average.")

    def _rows_to_ref(letter, rows):
        """Collapse row numbers into an Excel reference (union when non-contiguous)."""
        if not rows:
            return None
        blocks, start, prev = [], rows[0], rows[0]
        for rr in rows[1:]:
            if rr == prev + 1:
                prev = rr
                continue
            blocks.append((start, prev))
            start = prev = rr
        blocks.append((start, prev))
        parts = [f"{letter}{a}:{letter}{b}" for a, b in blocks]
        return parts[0] if len(parts) == 1 else "(" + ",".join(parts) + ")"

    for stat_idx, (label, r) in enumerate(zip(stat_labels, stat_rows)):
        set_cell(ws4, r, 2, label, font=BOLD_FONT)

        for dst_col, src_col in stat_col_map:
            src_letter = col_letter(src_col)
            fmt = FMT_PCT if src_col in (14, 15) else FMT_RATIO

            rows_for_col = ev_ebitda_rows if src_col == 10 else peer_rows
            rng = _rows_to_ref(src_letter, rows_for_col)

            # No peer rows at all (zero comps, or the CSV holds only the subject,
            # or every peer was screened out): a range formula would return #NUM!,
            # so emit the text "N/A" that every downstream AVERAGE already skips.
            if rng is None:
                set_cell(ws4, r, dst_col, "N/A", font=BLACK_FONT, border=THIN_BORDER,
                         alignment=Alignment(horizontal="right"))
                continue

            if stat_idx == 0:
                formula = f"=PERCENTILE({rng},0.25)"
            elif stat_idx == 1:
                formula = f"=MEDIAN({rng})"
            else:
                formula = f"=PERCENTILE({rng},0.75)"

            set_cell(ws4, r, dst_col, formula, font=BLACK_FONT, fmt=fmt, border=THIN_BORDER)

    _meta["comps_subject_row"] = R_CMP_SUBJECT
    _meta["comps_peer_rows"] = ",".join(str(x) for x in peer_rows)
    _meta["comps_ev_ebitda_rows"] = ",".join(str(x) for x in ev_ebitda_rows)
    _meta["comps_da_missing"] = ", ".join(_da_missing_names) or "none"
    _meta["comps_stale_excluded"] = ", ".join(_stale_names) or "none"
    _meta["comps_stat_median_row"] = R_STAT_MEDIAN
    _meta["comps_asof"] = _asof

    # ── Implied Valuation ──
    _r_impl_sec = 19 + _comps_row_shift
    _r_fin_hdr = 20 + _comps_row_shift
    _r_metric = 21 + _comps_row_shift
    _r_ni = 22 + _comps_row_shift
    _r_med = R_STAT_MEDIAN

    c = section_title(ws4, _r_impl_sec, 2, f'Implied Valuation for {C["company_name"]}')
    c.fill = LIGHT_GREEN
    for col_idx in range(3, 10):
        ws4.cell(row=_r_impl_sec, column=col_idx).fill = LIGHT_GREEN

    section_title(ws4, _r_fin_hdr, 2, f'{C["company_name"]} Financials')

    if USE_EV_SALES:
        set_cell(ws4, _r_metric, 2, "Revenue (JPY mn)", font=BOLD_FONT)
        set_cell(ws4, _r_metric, 3, C["base_year_revenue"], font=BLUE_FONT, fmt=FMT_YEN)
    else:
        set_cell(ws4, _r_metric, 2, "EBITDA (JPY mn)", font=BOLD_FONT)
        set_cell(ws4, _r_metric, 3, C["core_ebitda"], font=BLUE_FONT, fmt=FMT_YEN)

    set_cell(ws4, _r_ni, 2, "Net Income (JPY mn)", font=BOLD_FONT)
    set_cell(ws4, _r_ni, 3, C["core_net_income"], font=BLUE_FONT, fmt=FMT_YEN)
    set_cell(ws4, R_CMP_SHARES, 2, "Shares Outstanding", font=BOLD_FONT)
    set_cell(ws4, R_CMP_SHARES, 3, C["shares_outstanding"], font=BLUE_FONT, fmt=FMT_INT)
    set_cell(ws4, R_CMP_NETDEBT, 2, "Net Debt (JPY mn)", font=BOLD_FONT)
    set_cell(ws4, R_CMP_NETDEBT, 3, C["net_debt"], font=BLUE_FONT, fmt=FMT_YEN)

    section_title(ws4, 26 + _comps_row_shift, 2, "Implied Share Price (Median Multiples)")

    _impl_na = lambda row, msg: (
        set_cell(ws4, row, 3, "N/A", font=BLACK_FONT, border=TOP_BOTTOM,
                 alignment=Alignment(horizontal="right")),
        set_cell(ws4, row, 4, msg, font=GREY_FONT),
    )

    if USE_EV_SALES:
        set_cell(ws4, R_CMP_IMPL_MULT, 2, "Via EV/Sales (Median)", font=BOLD_FONT)
        set_cell(ws4, R_CMP_IMPL_MULT, 3,
                 f"=ROUND((C{_r_metric}*E{_r_med}-C{R_CMP_NETDEBT})*1000000/C{R_CMP_SHARES},0)",
                 font=BLACK_FONT, fmt=FMT_YEN, border=TOP_BOTTOM)
    else:
        set_cell(ws4, R_CMP_IMPL_MULT, 2, "Via EV/EBITDA (Median)", font=BOLD_FONT)
        if EBITDA_EXCLUDED:
            _impl_na(R_CMP_IMPL_MULT, "EBITDA <= 0 — median multiple not meaningful")
        elif EV_EBITDA_INVALID:
            # Fewer than 3 peers with a usable (D&A-inclusive) EBITDA: the median
            # is not a statistic. Written as text so AVERAGE/MIN/MAX on the
            # Executive Summary skip it — the method is dropped, never averaged in.
            set_cell(ws4, R_CMP_IMPL_MULT, 3, f"INVALID (n<3)", font=BLACK_FONT,
                     border=TOP_BOTTOM, alignment=Alignment(horizontal="right"))
            set_cell(ws4, R_CMP_IMPL_MULT, 4,
                     f"Only {n_ev_ebitda} peer(s) with D&A-inclusive EBITDA — "
                     f"EV/EBITDA excluded from Target Price", font=GREY_FONT)
        else:
            set_cell(ws4, R_CMP_IMPL_MULT, 3,
                     f"=ROUND((C{_r_metric}*D{_r_med}-C{R_CMP_NETDEBT})*1000000/C{R_CMP_SHARES},0)",
                     font=BLACK_FONT, fmt=FMT_YEN, border=TOP_BOTTOM)

    set_cell(ws4, R_CMP_IMPL_PER, 2, "Via PER (Median)", font=BOLD_FONT)
    if PER_EXCLUDED:
        _impl_na(R_CMP_IMPL_PER, "Net income <= 0 — PER not meaningful")
    elif not peer_rows:
        _impl_na(R_CMP_IMPL_PER, "No peer rows — median PER not meaningful")
    else:
        set_cell(ws4, R_CMP_IMPL_PER, 3,
                 f"=ROUND(C{_r_ni}*F{_r_med}*1000000/C{R_CMP_SHARES},0)",
                 font=BLACK_FONT, fmt=FMT_YEN, border=TOP_BOTTOM)

    # ── Normalised net income reference rows (標準メモ §1: Comps は正常化純利益の
    # 参考行を持つ) ──
    # A denominator distorted by one-off charges (5726's JPY2.6bn of 特別損失)
    # makes the reported PER read as 39x when the underlying business is nearer
    # 23x. The rule is "残置 + 除外 + 理由記録", so the distorted PER stays as the
    # Target-side number and the normalised read sits next to it as [参考], never
    # in the Target average.
    _norm = C.get("normalized_net_income")
    if _norm is not None:
        _r_norm = R_CMP_IMPL_PER + 1
        _r_norm_price = _r_norm + 1
        if isinstance(_norm, dict):
            _pretax = _norm.get("pretax")
            _addbacks = _norm.get("addbacks") or 0
            _label = _norm.get("label")
            if _pretax is not None:
                # Live off the tax-rate cell so the row follows a tax change.
                _norm_formula = (f"=ROUND(({_pretax:g}+{_addbacks:g})"
                                 f"*(1-'DCF Model'!C6),0)")
                _label = _label or (f"Net Income normalized (pretax {_pretax:,.0f}"
                                    f" + addbacks {_addbacks:,.0f}, tax rate C6)")
            else:
                _norm_formula = _norm.get("value")
                _label = _label or "Net Income normalized (参考)"
        else:
            _norm_formula = _norm
            _label = "Net Income normalized (参考)"
        set_cell(ws4, _r_norm, 2, _label, font=BOLD_FONT)
        set_cell(ws4, _r_norm, 3, _norm_formula, font=BLACK_FONT, fmt=FMT_YEN)
        if not PER_EXCLUDED and peer_rows:
            set_cell(ws4, _r_norm_price, 2,
                     "Via PER (Median, normalized NI) [参考・Target不算入]",
                     font=BOLD_FONT)
            set_cell(ws4, _r_norm_price, 3,
                     f"=ROUND(C{_r_norm}*F{_r_med}*1000000/C{R_CMP_SHARES},0)",
                     font=BLACK_FONT, fmt=FMT_YEN)
            _meta["comps_normalized_per_row"] = _r_norm_price
        if isinstance(_norm, dict) and _norm.get("note"):
            set_cell(ws4, _r_norm, 4, _norm["note"], font=GREY_FONT)
        _meta["comps_normalized_ni_row"] = _r_norm

    _meta["comps_ev_ebitda_invalid"] = "yes" if EV_EBITDA_INVALID else "no"
    _meta["comps_impl_mult_row"] = R_CMP_IMPL_MULT
    _meta["comps_impl_per_row"] = R_CMP_IMPL_PER

    # =====================================================================
    # SHEET 6: Sensitivity Analysis (Dynamic Excel formulas)
    # =====================================================================
    ws5 = wb.create_sheet("Sensitivity Analysis")
    ws5.sheet_properties.tabColor = "996600"

    ws5.column_dimensions["A"].width = 3
    ws5.column_dimensions["B"].width = 24
    for letter in ["C", "D", "E", "F", "G", "H", "I"]:
        ws5.column_dimensions[letter].width = 14

    set_cell(ws5, 2, 2, "Sensitivity Analysis", font=TITLE_FONT)

    # ── Current values reference ──
    set_cell(ws5, 3, 2, "Current WACC:", font=BOLD_FONT)
    set_cell(ws5, 3, 3, "='DCF Model'!C26", font=BLACK_FONT, fmt=FMT_PCT2)
    set_cell(ws5, 3, 5, "Terminal g:", font=BOLD_FONT)
    set_cell(ws5, 3, 6, "='DCF Model'!C13", font=BLACK_FONT, fmt=FMT_PCT2)
    set_cell(ws5, 3, 8, "Exit Multiple:", font=BOLD_FONT)
    set_cell(ws5, 3, 9, "='DCF Model'!C14", font=BLACK_FONT, fmt=FMT_RATIO)

    # ── Dynamic formula builders ──
    _DCF = "'DCF Model'"
    _SHARES = f"{_DCF}!C15"
    _NET_DEBT = f"{_DCF}!C16"
    _last_cl = col_letter(3 + proj_years - 1)
    _ufcf_cells = [f"{_DCF}!{col_letter(3 + yr)}{R_UFCF}" for yr in range(proj_years)]
    _stub_ref = f"{_DCF}!C{R_STUB_FRACTION}"

    def _build_pgm_formula(wacc_ref, tg_ref):
        pv_parts = [f"{_ufcf_cells[yr]}/(1+{wacc_ref})^({_stub_ref}+{yr})" for yr in range(proj_years)]
        last_ufcf = _ufcf_cells[proj_years - 1]
        pv_tv = f"{last_ufcf}*(1+{tg_ref})/({wacc_ref}-{tg_ref})/(1+{wacc_ref})^({_stub_ref}+{proj_years-1})"
        return f'=IFERROR(ROUND(({"+".join(pv_parts)}+{pv_tv}-{_NET_DEBT})*1000000/{_SHARES},0),"")'

    def _build_exit_formula(wacc_ref, mult_ref):
        pv_parts = [f"{_ufcf_cells[yr]}/(1+{wacc_ref})^({_stub_ref}+{yr})" for yr in range(proj_years)]
        # Match the DCF Exit metric: Year-5 Revenue (EV/Sales) or Year-5 EBITDA (EV/EBITDA)
        if USE_EV_SALES_EXIT:
            yr5_metric = f"{_DCF}!{_last_cl}{R_REVENUE}"
        else:
            yr5_metric = f"({_DCF}!{_last_cl}{R_EBIT}+{_DCF}!{_last_cl}{R_DA})"
        pv_tv = f"{yr5_metric}*{mult_ref}/(1+{wacc_ref})^({_stub_ref}+{proj_years-1})"
        return f'=IFERROR(ROUND(({"+".join(pv_parts)}+{pv_tv}-{_NET_DEBT})*1000000/{_SHARES},0),"")'

    # ── Dynamic header helpers ──
    _N_GRID = 7
    _CENTER_IDX = 3  # 4th position (0-based)
    _WACC_STEP = 0.005
    _TG_STEP   = 0.0025
    _EXIT_STEP = 1.0
    _ANCHOR_WACC = "$C$3"
    _ANCHOR_TG   = "$F$3"
    _ANCHOR_EXIT = "$I$3"

    def _offset_formula(anchor, offset_val):
        if offset_val == 0:
            return f"={anchor}"
        elif offset_val > 0:
            return f"={anchor}+{offset_val}"
        else:
            return f"={anchor}-{abs(offset_val)}"

    # ── Table 1: WACC vs Terminal Growth Rate (PGM) ──
    T1_TITLE = 5
    T1_HDR = 6
    T1_DATA = 7

    section_title(ws5, T1_TITLE, 2,
                  "Table 1: WACC vs Terminal Growth Rate (PGM - Implied Share Price, JPY)")

    # Column headers: TG (dynamic, centered on F3)
    set_cell(ws5, T1_HDR, 2, "WACC \\ Terminal g", font=HEADER_FONT, fill=HEADER_FILL,
             alignment=Alignment(horizontal="center", wrap_text=True), border=THIN_BORDER)
    for j in range(_N_GRID):
        offset = round((j - _CENTER_IDX) * _TG_STEP, 6)
        set_cell(ws5, T1_HDR, 3 + j, _offset_formula(_ANCHOR_TG, offset),
                 font=HEADER_FONT, fill=HEADER_FILL, fmt=FMT_PCT,
                 alignment=Alignment(horizontal="center"), border=THIN_BORDER)

    # Row headers: WACC (dynamic, centered on C3) + data formulas
    for i in range(_N_GRID):
        r = T1_DATA + i
        offset = round((i - _CENTER_IDX) * _WACC_STEP, 6)
        set_cell(ws5, r, 2, _offset_formula(_ANCHOR_WACC, offset),
                 font=BLUE_FONT, fmt=FMT_PCT, border=THIN_BORDER)
        for j in range(_N_GRID):
            col = 3 + j
            cl = col_letter(col)
            formula = _build_pgm_formula(f"$B{r}", f"{cl}${T1_HDR}")
            set_cell(ws5, r, col, formula, font=BLACK_FONT, fmt=FMT_YEN, border=THIN_BORDER)

    # ── Table 2: WACC vs Exit Multiple ──
    T2_TITLE = T1_DATA + _N_GRID + 2
    T2_HDR = T2_TITLE + 1
    T2_DATA = T2_HDR + 1

    section_title(ws5, T2_TITLE, 2,
                  "Table 2: WACC vs Exit Multiple (Exit Multiple - Implied Share Price, JPY)")

    # Column headers: Exit Multiple (dynamic, centered on I3)
    set_cell(ws5, T2_HDR, 2, "WACC \\ Exit Multiple", font=HEADER_FONT, fill=HEADER_FILL,
             alignment=Alignment(horizontal="center", wrap_text=True), border=THIN_BORDER)
    for j in range(_N_GRID):
        offset = round((j - _CENTER_IDX) * _EXIT_STEP, 6)
        set_cell(ws5, T2_HDR, 3 + j, _offset_formula(_ANCHOR_EXIT, offset),
                 font=HEADER_FONT, fill=HEADER_FILL, fmt=FMT_RATIO,
                 alignment=Alignment(horizontal="center"), border=THIN_BORDER)

    # Row headers: WACC (dynamic) + data formulas
    for i in range(_N_GRID):
        r = T2_DATA + i
        offset = round((i - _CENTER_IDX) * _WACC_STEP, 6)
        set_cell(ws5, r, 2, _offset_formula(_ANCHOR_WACC, offset),
                 font=BLUE_FONT, fmt=FMT_PCT, border=THIN_BORDER)
        for j in range(_N_GRID):
            col = 3 + j
            cl = col_letter(col)
            formula = _build_exit_formula(f"$B{r}", f"{cl}${T2_HDR}")
            set_cell(ws5, r, col, formula, font=BLACK_FONT, fmt=FMT_YEN, border=THIN_BORDER)

    # ── Note ──
    _note_row = T2_DATA + _N_GRID + 1
    set_cell(ws5, _note_row, 2,
             "All values dynamically linked to DCF Model. "
             "Headers auto-center on current WACC / Terminal g / Exit Multiple.",
             font=GREY_FONT)
    ws5.merge_cells(start_row=_note_row, start_column=2,
                    end_row=_note_row, end_column=9)

    # ── Table 3: FX sensitivity (export-exposed names only) ──
    # Off by default: for a domestic name a currency grid is noise. Turned on
    # with fx_sensitivity.enabled, it answers the one question a yen move raises
    # for an exporter — how much operating profit moves per 1 JPY — by netting
    # the USD-linked share of revenue against the USD-linked share of COGS, so a
    # yen appreciation is not scored as pure downside for an importer of feed.
    _fx = C.get("fx_sensitivity") or {}
    _fx_rows = None
    if _fx.get("enabled"):
        _missing = [k for k in ("assumption_rate", "usd_revenue_ratio", "usd_cogs_ratio")
                    if _fx.get(k) is None]
        if _missing:
            print(f"  WARNING: fx_sensitivity.enabled but {', '.join(_missing)} "
                  f"missing - Table 3 skipped.")
            _meta["fx_sensitivity"] = f"skipped (missing {','.join(_missing)})"
        else:
            _pair = _fx.get("currency_pair", "USD/JPY")
            _cur = _pair.split("/")[0]
            _offsets = list(_fx.get("offsets") or (-20, -10, -5, 0, 5, 10, 20))
            _est = _fx.get("estimated", True)
            _est_tag = " (推定)" if _est else ""
            _yr = C.get("projection_start_fy") or "Year 1"

            T3 = _note_row + 2
            T3_RATE = T3 + 1
            T3_REV_R = T3 + 2
            T3_COGS_R = T3 + 3
            T3_REV = T3 + 4
            T3_COGS = T3 + 5
            T3_OP = T3 + 6
            T3_SENS = T3 + 7
            T3_HDR = T3 + 9
            T3_OP_ROW = T3 + 10
            T3_OPM = T3 + 11
            T3_NOTE = T3 + 12

            _title = f"Table 3: FX Sensitivity - {_yr} Operating Income"
            if _est:
                _title += f" ({_cur}連動比率は推定・会社開示なし・要確認)"
            c = section_title(ws5, T3, 2, _title)
            _src = _fx.get("assumption_source")
            set_cell(ws5, T3_RATE, 2,
                     f"Company FX assumption ({_pair}{', ' + _src if _src else ''})",
                     font=BOLD_FONT)
            set_cell(ws5, T3_RATE, 3, _fx["assumption_rate"], font=BLUE_FONT,
                     fmt="#,##0.00", border=INPUT_BORDER)
            set_cell(ws5, T3_REV_R, 2, f"{_cur}-linked revenue ratio{_est_tag}",
                     font=BOLD_FONT)
            set_cell(ws5, T3_REV_R, 3, _fx["usd_revenue_ratio"], font=BLUE_FONT,
                     fmt=FMT_PCT, border=INPUT_BORDER)
            set_cell(ws5, T3_COGS_R, 2, f"{_cur}-linked COGS ratio{_est_tag}",
                     font=BOLD_FONT)
            set_cell(ws5, T3_COGS_R, 3, _fx["usd_cogs_ratio"], font=BLUE_FONT,
                     fmt=FMT_PCT, border=INPUT_BORDER)

            # Live off the DCF Model's Year-1 column, so the table follows the
            # active scenario instead of freezing the Base case.
            set_cell(ws5, T3_REV, 2, f"{_yr} Revenue - active scenario", font=BOLD_FONT)
            set_cell(ws5, T3_REV, 3, f"='DCF Model'!C{R_REVENUE}", font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws5, T3_COGS, 2, f"{_yr} COGS - active scenario", font=BOLD_FONT)
            set_cell(ws5, T3_COGS, 3, f"='DCF Model'!C{R_COGS}", font=BLACK_FONT, fmt=FMT_YEN)
            set_cell(ws5, T3_OP, 2, f"{_yr} Operating Income - active scenario",
                     font=BOLD_FONT)
            set_cell(ws5, T3_OP, 3, f"='DCF Model'!C{R_EBIT}", font=BLACK_FONT, fmt=FMT_YEN)

            set_cell(ws5, T3_SENS, 2,
                     f"OP sensitivity per +/-1 JPY per {_cur} (JPY mn)", font=BOLD_FONT)
            set_cell(ws5, T3_SENS, 3,
                     f"=(C{T3_REV}*C{T3_REV_R}-C{T3_COGS}*C{T3_COGS_R})/C{T3_RATE}",
                     font=BLACK_FONT, fmt=FMT_YEN, fill=LIGHT_GREEN)

            set_cell(ws5, T3_HDR, 2, f"{_pair} rate", font=HEADER_FONT, fill=HEADER_FILL)
            set_cell(ws5, T3_OP_ROW, 2, f"Implied {_yr} OP (JPY mn)", font=BOLD_FONT)
            set_cell(ws5, T3_OPM, 2, "Implied OPM", font=BOLD_FONT)
            for j, off in enumerate(_offsets):
                col = 3 + j
                cl = col_letter(col)
                set_cell(ws5, T3_HDR, col,
                         f"=$C${T3_RATE}" if off == 0 else f"=$C${T3_RATE}{off:+g}",
                         font=HEADER_FONT, fmt="#,##0.0", fill=HEADER_FILL,
                         alignment=Alignment(horizontal="center"))
                set_cell(ws5, T3_OP_ROW, col,
                         f"=$C${T3_OP}+({cl}{T3_HDR}-$C${T3_RATE})*$C${T3_SENS}",
                         font=BLACK_FONT, fmt=FMT_YEN, border=THIN_BORDER)
                set_cell(ws5, T3_OPM, col,
                         f'=IFERROR({cl}{T3_OP_ROW}/$C${T3_REV},"N/A")',
                         font=BLACK_FONT, fmt=FMT_PCT, border=THIN_BORDER)

            _fx_note = _fx.get("note") or (
                f"Revenue is assumed {_fx['usd_revenue_ratio']:.0%} {_cur}-linked and "
                f"COGS {_fx['usd_cogs_ratio']:.0%}; the table shows the NET effect. "
                f"Volumes and prices are held at the active scenario."
            )
            set_cell(ws5, T3_NOTE, 2, _fx_note, font=GREY_FONT)
            ws5.merge_cells(start_row=T3_NOTE, start_column=2,
                            end_row=T3_NOTE, end_column=9)

            _fx_rows = {"title": T3, "rate": T3_RATE, "sens": T3_SENS,
                        "note": T3_NOTE, "estimated": _est}
            _meta["fx_sensitivity"] = (
                f"rows {T3}-{T3_NOTE}; {_pair} @ {_fx['assumption_rate']}; "
                f"rev_ratio={_fx['usd_revenue_ratio']} cogs_ratio={_fx['usd_cogs_ratio']}; "
                f"offsets={','.join(str(o) for o in _offsets)}; "
                f"{'estimated' if _est else 'disclosed'}"
            )
    else:
        _meta["fx_sensitivity"] = "not enabled"

    # =====================================================================
    # SHEET 4 (position): Reverse DCF
    # =====================================================================
    # Built last, placed 4th. It reads the other sheets, so it is written once
    # every row constant it cites is settled, then moved into place — the
    # standard 8-sheet order is Exec / FS / DCF Model / Reverse DCF / NWC /
    # Comps / Sensitivity / Adjustments Log (docs/DCFフォーマット標準メモ §1).
    _rdcf_params, _rdcf_skip = _rdcf_resolve_params(C)
    if _rdcf_params is None:
        _meta["reverse_dcf_sheet"] = f"skipped ({_rdcf_skip})"
        print(f"  WARNING: Reverse DCF sheet skipped - {_rdcf_skip}")
    else:
        _bench_row = None
        _bench_key = _tkr_key(_rdcf_params.get("benchmark_ticker"))
        if _bench_key:
            for _i, _comp in enumerate(C.get("comps") or []):
                if _tkr_key(_comp.get("ticker")) == _bench_key:
                    _bench_row = 5 + _i
                    break
            if _bench_row is None:
                print(f"  WARNING: reverse_dcf.benchmark_ticker "
                      f"{_rdcf_params['benchmark_ticker']!r} is not in the comps "
                      f"table - Block E (transaction benchmark) omitted.")
        _build_reverse_dcf_sheet(
            wb, C, _rdcf_params,
            {
                "proj_last_col": col_letter(3 + C["projection_years"] - 1),
                "r_revenue": R_REVENUE,
                "r_da": R_DA,
                "r_ev_pgm": R_EV_PGM,
                "cmp_subject_row": R_CMP_SUBJECT,
                "cmp_stat_median_row": R_STAT_MEDIAN,
                "cmp_benchmark_row": _bench_row,
            },
        )
        _rdcf_place_after(wb, "DCF Model")
        _meta["reverse_dcf_sheet"] = (
            f"op0={_rdcf_params['op0']:,.0f}({_rdcf_params['op0_label']}) "
            f"peak={_rdcf_params['peak_op']:,.0f}({_rdcf_params['peak_label']}) "
            f"peak_opm={_rdcf_params['peak_opm']:.4f} "
            f"N={'/'.join(str(n) for n in _rdcf_params['n_years'])}"
            + (f" benchmark_row={_bench_row}" if _bench_row else "")
        )

    # =====================================================================
    # SHEET 7 & 8: Segment / Driver Analysis (optional)
    # =====================================================================
    segments = C.get("segments")
    if segments:
        proj_years = C["projection_years"]
        _year_labels = [f"Year {y}" for y in range(1, proj_years + 1)]
        if C.get("projection_start_fy"):
            import re as _re
            _m = _re.search(r"FY(\d+)", C["projection_start_fy"])
            if _m:
                _base_fy = int(_m.group(1))
                _year_labels = [f"FY{_base_fy + y}(E)" for y in range(proj_years)]

        _create_segment_sheet(wb, C, segments, proj_years, _year_labels)
        _create_driver_sheet(wb, C, segments, proj_years, _year_labels)

    # =====================================================================
    # SHEET: Adjustments Log (manual-edit ledger + pipeline metadata)
    # =====================================================================
    # Every model eventually gets hand-patched. Shipping the ledger with the
    # model means the edit gets recorded in the model instead of in a chat log.
    # The metadata block below it is what scripts/validate_output.py checks the
    # workbook against, so the model carries its own ground truth.
    ws_log = wb.create_sheet("Adjustments Log")
    ws_log.sheet_properties.tabColor = "808080"
    ws_log.column_dimensions["A"].width = 3
    ws_log.column_dimensions["B"].width = 14
    ws_log.column_dimensions["C"].width = 16
    ws_log.column_dimensions["D"].width = 46
    ws_log.column_dimensions["E"].width = 46
    ws_log.column_dimensions["F"].width = 20
    ws_log.column_dimensions["G"].width = 12

    set_cell(ws_log, 2, 2, f'Adjustments Log - {C["company_name"]}', font=TITLE_FONT)
    header_row(ws_log, 4, 2, 8, ["日付", "セル", "変更内容", "理由", "元の値", "状態"])

    _template_rev = C.get("_template_rev") or _dt.date.today().strftime("%Y-%m-%d")
    set_cell(ws_log, 5, 2, _dt.datetime.now().strftime("%Y-%m-%d"))
    set_cell(ws_log, 5, 3, "-")
    set_cell(ws_log, 5, 4,
             f"Generated by pipeline on {_dt.datetime.now().strftime('%Y-%m-%d %H:%M')}, "
             f"template rev {_template_rev}")
    set_cell(ws_log, 5, 5, "初期生成（手修正なし）")
    set_cell(ws_log, 5, 6, "-")
    set_cell(ws_log, 5, 7, "generated")

    # ── Auto-recorded derivations (row 6 onward) ──
    # Anything the pipeline DERIVED rather than took at face value is written
    # here, so the reader sees the basis without opening the overrides. Manual
    # entries (scripts/fill_adjustments_log.py) are appended below these.
    _auto_log = []
    _today = _dt.datetime.now().strftime("%Y-%m-%d")
    if _cod_actual:
        _a = _cod_actual
        _auto_log.append((
            f"DCF Model!C11",
            f"After-tax Cost of Debt = {_a['kd_after_tax']:.2%} (実績ベース)",
            f"実績Kd = 支払利息{_a['interest_expense']:,.0f}"
            + (f" + 手数料{_a['loan_fees']:,.0f}" if _a['loan_fees'] else "")
            + f" ÷ 平均有利子負債{_a['avg_debt']:,.0f}"
            f"(期首{_a['debt_beginning']:,.0f}/期末{_a['debt_ending']:,.0f}、{_a['basis']}) "
            f"= 税引前{_a['kd_pretax']:.2%} × (1−{C['tax_rate']:.1%}) = {_a['kd_after_tax']:.2%}。"
            f"【マージナルコスト注記】これは既存借入の実績平均コストであり、"
            f"新規調達の限界コストではない。テンプレ既定(税引前2.8% = 税引後"
            f"{(_cod_default_at if _cod_default_at is not None else 0.0194):.2%})は"
            f"10年JGB+スプレッドの推定値で、金利上昇局面での借換え・増額調達には"
            f"既定値に近い水準を当てるべき。WACCは実績Kd採用で低下する（＝保守的でない方向）",
            (f"テンプレ既定/overrides {_cod_default_at:.2%}(税引後)"
             if _cod_default_at is not None else "テンプレ既定"),
            "確定(実績)",
        ))
    if _fx_rows:
        _auto_log.append((
            f"Sensitivity!C{_fx_rows['rate']}:C{_fx_rows['rate'] + 2}",
            f"為替感応度 Table 3 を生成（{C['fx_sensitivity'].get('currency_pair', 'USD/JPY')} "
            f"{C['fx_sensitivity']['assumption_rate']}円、売上連動比率"
            f"{C['fx_sensitivity']['usd_revenue_ratio']:.0%}、原価連動比率"
            f"{C['fx_sensitivity']['usd_cogs_ratio']:.0%}）",
            "輸出型フラグ(fx_sensitivity.enabled)がONのため生成。"
            + ("【推定・要確認】連動比率は会社開示が無く推定値。青字入力セルであり、"
               "開示または IR 確認が取れ次第上書きすること。感応度 = (売上×売上連動比率 "
               "− COGS×原価連動比率) ÷ 前提レート で、原価側の連動が円高メリットを"
               "一部相殺する構造を織り込んでいる。"
               if _fx_rows["estimated"] else "連動比率は会社開示値。"),
            "-",
            "推定・要確認" if _fx_rows["estimated"] else "確定",
        ))
    for _i, _rec in enumerate(_auto_log):
        _r = 6 + _i
        set_cell(ws_log, _r, 2, _today)
        set_cell(ws_log, _r, 3, _rec[0], font=BOLD_FONT)
        for _j, _v in enumerate(_rec[1:], start=4):
            set_cell(ws_log, _r, _j, _v)
        ws_log.row_dimensions[_r].height = 30
    _meta["auto_log_entries"] = len(_auto_log)

    _meta.setdefault("template_rev", _template_rev)
    _meta["generated_at"] = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _meta["company_name"] = C.get("company_name")
    _meta["ticker"] = C.get("ticker")
    # Audit trail recorded by generate_dcf.py: year-key alignment (bug B1)
    # and the 会社予想 source ladder (フェーズ2 #9).
    for _k in ("_fs_year_coverage", "_fs_year_map", "_fs_year_sources",
               "_guidance_source", "_guidance_note"):
        if C.get(_k) is not None:
            _meta[_k.lstrip("_")] = C[_k]

    _meta_start = max(12, 6 + len(_auto_log) + 2)
    c = section_title(ws_log, _meta_start, 2,
                      "Pipeline Metadata (do not edit — read by scripts/validate_output.py)")
    c.fill = LIGHT_FILL
    header_row(ws_log, _meta_start + 1, 2, 3, ["key", "value"])
    for _i, (_k, _v) in enumerate(sorted(_meta.items())):
        _r = _meta_start + 2 + _i
        set_cell(ws_log, _r, 2, str(_k), font=BOLD_FONT)
        _vs = "" if _v is None else str(_v)
        _c = set_cell(ws_log, _r, 3, _vs)
        # A value starting with "=" would be stored as a formula (and blow up as
        # #NAME?); metadata is always text.
        if _vs.startswith("="):
            _c.data_type = "s"

    # =====================================================================
    # SAVE & VERIFY
    # =====================================================================

    if output_path is None:
        ticker_safe = C["ticker"].replace(".", "")
        output_path = f"{ticker_safe}_Equity_Research_V3.xlsx"

    wb.save(output_path)
    print(f"\nSaved: {output_path}")

    # Run recalc.py for verification
    recalc_script = os.path.join("scripts", "recalc.py")
    if os.path.exists(recalc_script):
        print(f"\nRunning verification: python {recalc_script} {output_path}")
        result = subprocess.run([sys.executable, recalc_script, output_path],
                                capture_output=True, text=True)
        print(result.stdout)
        if result.stderr:
            print("STDERR:", result.stderr)

    return output_path


# =====================================================================
# STANDALONE EXECUTION (backward compatibility)
# =====================================================================
if __name__ == "__main__":
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "scripts")))
    from pdf_parser import extract_all_financials
    from comps_fetcher import get_comps_data

    _script_dir = os.path.dirname(os.path.abspath(__file__))
    _pdf_result = extract_all_financials(_script_dir)

    _FALLBACK = {
    "hist_revenue": [271, 332, 490, 517],
    "hist_operating_income": [-433, -598, -527, -800],
    "hist_net_income": [-2237, -413, -69, -801],
    "hist_cogs": [147, 156, 52, 177],
    "hist_sga": [558, 775, 966, 1141],
    "hist_ocf": [-515, -619, -491, -815],
    "hist_capex": [137, 20, 433, 162],
    "hist_cash": [604, 852, 1720, 2594],
    "hist_debt": [None, 200, 200, 200],
    "latest_net_debt": -2394,
    }

    if _pdf_result is None:
        print("FATAL: PDF extraction failed. Using fallback values.")
        _pdf_result = _FALLBACK

    def _get(key):
        val = _pdf_result.get(key)
        if val is None:
            return _FALLBACK.get(key)
        return val

    config = {
    # ── Company Info ──
    "company_name": "TEMPLATE COMPANY",
    "ticker": "0000.T",
    "exchange": "TSE Growth",
    "sector": "Information & Communication",
    "current_price": 1000,
    "shares_outstanding": 10_000_000,
    "net_debt": _get("latest_net_debt") or -1000,  # JPY mn (negative = net cash), from BS
    
    # ── Historical Financials (JPY mn) — all from PDFs ──
    "hist_years": ["FY2022 (Mar-22)", "FY2023 (Mar-23)", "FY2024 (Mar-24)", "FY2025 (Mar-25)"],
    "hist_revenue":          _get("hist_revenue"),
    "hist_operating_income": _get("hist_operating_income"),
    "hist_net_income":       _get("hist_net_income"),
    "hist_cogs":             _get("hist_cogs"),
    "hist_sga":              _get("hist_sga"),
    "hist_ocf":              _get("hist_ocf"),
    "hist_capex":            _get("hist_capex"),
    "hist_cash":             _get("hist_cash"),
    "hist_debt":             _get("hist_debt"),
    
    # ── DCF Assumptions — Future Projections ──
    "scenarios": {
    "Base":       {"revenue_growth": [0.10,0.08,0.07,0.06,0.05], "cogs_pct": [0.70,0.70,0.70,0.70,0.70], "sga_pct": [0.13,0.13,0.13,0.13,0.13], "dso_days": [60,60,60,60,60], "dih_days": [30,30,30,30,30], "dpo_days": [45,45,45,45,45]},
    "Upside":     {"revenue_growth": [0.10,0.12,0.18,0.20,0.15], "cogs_pct": [0.71,0.705,0.70,0.70,0.68], "sga_pct": [0.12,0.12,0.12,0.12,0.12], "dso_days": [55,55,55,55,55], "dih_days": [28,28,28,28,28], "dpo_days": [48,48,48,48,48]},
    "Management": {"revenue_growth": [0.10,0.10,0.10,0.10,0.10], "cogs_pct": [0.70,0.70,0.70,0.70,0.70], "sga_pct": [0.13,0.13,0.13,0.13,0.13], "dso_days": [60,60,60,60,60], "dih_days": [30,30,30,30,30], "dpo_days": [45,45,45,45,45]},
    "Downside 1": {"revenue_growth": [0.02,0.02,0.02,0.02,0.02], "cogs_pct": [0.73,0.73,0.74,0.75,0.73], "sga_pct": [0.14,0.14,0.14,0.14,0.14], "dso_days": [65,65,65,65,65], "dih_days": [33,33,33,33,33], "dpo_days": [42,42,42,42,42]},
    "Downside 2": {"revenue_growth": [0.00,0.00,0.00,0.00,0.00], "cogs_pct": [0.76,0.76,0.76,0.76,0.76], "sga_pct": [0.15,0.15,0.15,0.15,0.15], "dso_days": [70,70,70,70,70], "dih_days": [35,35,35,35,35], "dpo_days": [40,40,40,40,40]},
    },
    "capex_pct": 0.03,
    "da_pct": 0.015,
    "tax_rate": 0.30,
    "risk_free": 0.023,
    "beta": 1.75,
    "erp": 0.060,
    "size_premium": 0.050,
    "cost_of_debt_at": 0.015,
    "de_ratio": 0.05,
    "terminal_growth": 0.015,
    "exit_multiple": 15.0,
    "projection_years": 5,
    "base_year_revenue": (_get("hist_revenue") or [517])[-1],
    "base_year_cogs": (_get("hist_cogs") or [177])[-1],
    
    # ── NWC Base Year Actuals (JPY mn) — edit for each company ──
    "base_year_ar":   85,    # Accounts Receivable
    "base_year_inv":  14,    # Inventory
    "base_year_ap":   22,    # Accounts Payable
    
    # ── Comparable Companies (loaded dynamically from CSV) ──
    "comps": get_comps_data(os.path.join(_script_dir, "comps_input_template.csv")),
    
    # ── Kudan Comps Data (for implied valuation) ──
    "core_ebitda": (_get("hist_operating_income") or [-800])[-1] + 8,
    "core_net_income": (_get("hist_net_income") or [-801])[-1],
    
    # ── Investment Thesis & Risks ──
    "investment_thesis": [
    "1. Global leader in Artificial Perception (SLAM) technology",
    "2. High leverage on revenue growth due to fixed-cost intensive IP licensing model",
    "3. Transitioning from R&D phase to commercial scaling phase",
    ],
    "key_risks": [
    "1. Prolonged losses and negative free cash flow burning cash runway",
    "2. Long sales cycles converting PoC (Proof of Concept) to commercial licenses",
    "3. High WACC (17.7%) depressing present value heavily",
    ],
    
    # ── V3 Settings ──
    "primary_multiple": "EV/Sales",  # "EV/EBITDA" or "EV/Sales"
    }
    
    # Restore flat arrays from Base scenario (backward compatibility for sensitivity analysis)
    _base = config["scenarios"]["Base"]
    config["revenue_growth"]    = _base["revenue_growth"]
    config["cogs_pct"]          = _base["cogs_pct"]
    config["sga_pct"]           = _base["sga_pct"]
    
    # =====================================================================

    config["current_price"], config["shares_outstanding"] = get_live_market_data(
        config.get("ticker", ""),
        config.get("current_price", 0),
        config.get("shares_outstanding", 0)
    )

    generate_dcf_workbook(config)
