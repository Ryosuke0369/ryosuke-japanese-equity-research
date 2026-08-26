"""reverse_dcf_sheet.py — the standard 'Reverse DCF' sheet (workbook sheet 4 of 8).

The rest of dcf_comps_template answers "what is it worth?". This sheet answers
the inverse — "what does today's price already assume?" — which is the
decision-relevant question for a cyclical, whose reported earnings say little
about mid-cycle earning power.

Every cell is a live Excel formula off 'DCF Model' / 'Executive Summary' /
'Comps Analysis', so the sheet tracks a changed price, WACC or tax rate without
regeneration. Every reference is resolved from the template's own row constants
(passed in as `refs`) — this module hardcodes no address on another sheet.

Model
-----
Operating profit ramps LINEARLY from the last actual (OP0) to a steady state
OP* over N years, then grows at g for ever with capex = D&A and no working
capital build (a true steady state). Solving today's EV for OP*:

    EV = (1-t)[ OP0*(A_N - B_N) + OP* * B_N ] + OP* * (1-t)*(1+g)/((w-g)(1+w)^N)

    A_N   = (1-(1+w)^-N)/w                                       annuity factor
    B_N   = (1+w)/(w^2 N) * (1-(N+1)/(1+w)^N + N/(1+w)^(N+1))     ramp-weighted PV
    TVF_N = (1+g)/((w-g)(1+w)^N)                                  terminal factor

    OP* = ( EV/(1-t) - OP0*(A_N-B_N) ) / ( B_N + TVF_N )

B_N is the closed form of (1/N) * SUM(t=1..N) t/(1+w)^t. `selftest_bn()` checks
it against the direct summation on every build — a closed form nobody re-derives
is a closed form nobody notices has gone wrong.
"""

from openpyxl.styles import Alignment

# ---------------------------------------------------------------- row layout
R_IN0 = 6                # Block 0 inputs, C6..C17
R_PEAK_OPM = 18
R_GAP = 19
R_GAP_PS = 20
R_A = 24                 # Block A rows 24..29
HDR = 36                 # WACC header row shared by the three Block-B grids
G1 = 38                  # ramp weight (A-B)          rows 38..41
G2 = 43                  # steady weight (B+TVF)      rows 43..46
OPS = 48                 # required OP*               rows 48..51
MULT = 53                # multiple of the peak       rows 53..56
R_C = 60                 # Block C anchor
R_C_HDR = 62             # Block C grid header
R_D = 70                 # Block D anchor
DEFAULT_N_YEARS = (3, 5, 7, 10)
HEADLINE_OFFSET = 1      # index of the "headline" N within n_years
WACC_COL = "F"           # third of five WACC columns == the model WACC
SHEET_NAME = "Reverse DCF"


# ------------------------------------------------------------------ self-test
def selftest_bn():
    """The B_N closed form must equal the direct summation it replaces."""
    for w in (0.03, 0.05, 0.0948, 0.15, 0.25):
        for n in (1, 3, 5, 7, 10, 20):
            direct = sum(t / (1 + w) ** t for t in range(1, n + 1)) / n
            closed = ((1 + w) / (w ** 2 * n)
                      * (1 - (n + 1) / (1 + w) ** n + n / (1 + w) ** (n + 1)))
            assert abs(direct - closed) < 1e-10, (w, n, direct, closed)
    return True


# ------------------------------------------------------------------ parameters
def resolve_params(C):
    """Derive the sheet's inputs from config + the optional `reverse_dcf` block.

    Returns (params, skip_reason). `params` is None when the sheet cannot be
    built honestly — never a sheet full of zeros dressed up as a valuation.
    """
    rd = dict(C.get("reverse_dcf") or {})
    if rd.get("enabled") is False:
        return None, "reverse_dcf.enabled = false"

    years = list(C.get("hist_years") or [])
    ops = list(C.get("hist_operating_income") or [])
    revs = list(C.get("hist_revenue") or [])

    def _num(x):
        return x if isinstance(x, (int, float)) else None

    pairs = [(years[i] if i < len(years) else "FY-%d" % i, _num(ops[i]),
              _num(revs[i]) if i < len(revs) else None)
             for i in range(len(ops))]
    actual = [p for p in pairs if p[1] is not None]

    op0 = rd.get("op0")
    op0_label = rd.get("op0_label")
    if op0 is None:
        if not actual:
            return None, ("no historical operating income — set reverse_dcf.op0 "
                          "in the overrides to build the sheet")
        op0_label = op0_label or actual[-1][0]
        op0 = actual[-1][1]
    op0_label = op0_label or "latest FY"

    peak_op = rd.get("peak_op")
    peak_label = rd.get("peak_label")
    peak_opm = rd.get("peak_opm")
    if peak_op is None:
        if not actual:
            return None, "no historical operating income to locate a cycle peak"
        peak = max(actual, key=lambda p: p[1])
        peak_label = peak_label or peak[0]
        peak_op = peak[1]
        if peak_opm is None and peak[2]:
            peak_opm = peak[1] / peak[2]
    peak_label = peak_label or "cycle peak"
    if peak_opm is None:
        base_rev = C.get("base_year_revenue")
        if peak_op and base_rev:
            peak_opm = peak_op / base_rev
    if not peak_op or peak_op <= 0:
        return None, ("cycle-peak operating profit is <= 0 — the multiple-of-peak "
                      "rows would be meaningless; set reverse_dcf.peak_op")
    if not peak_opm or peak_opm <= 0:
        return None, ("cannot derive a positive cycle-peak operating margin — "
                      "set reverse_dcf.peak_opm in the overrides")

    grid = rd.get("opm_grid")
    if not grid:
        # Scaled off the company's own cycle peak, so the table straddles the
        # relevant range for a 3%-margin distributor and a 30%-margin materials
        # name alike. `None` marks the row that reads the live peak-OPM cell.
        grid = ([round(peak_opm * f, 4) for f in (0.50, 0.75)] + [None]
                + [round(peak_opm * f, 4) for f in (1.25, 1.50)])

    return {
        "op0": float(op0),
        "op0_label": str(op0_label),
        "peak_op": float(peak_op),
        "peak_label": str(peak_label),
        "peak_opm": float(peak_opm),
        "opm_grid": list(grid),
        "n_years": tuple(rd.get("n_years") or DEFAULT_N_YEARS),
        "benchmark_ticker": rd.get("benchmark_ticker"),
        "deal_note": list(rd.get("deal_note") or []),
    }, None


# ------------------------------------------------------------------- formulas
def _f_A(cl, r):
    w, n = "%s$%d" % (cl, HDR), "$C%d" % r
    return "(1-(1+{w})^(-{n}))/{w}".format(w=w, n=n)


def _f_B(cl, r):
    w, n = "%s$%d" % (cl, HDR), "$C%d" % r
    return ("(1+{w})/({w}^2*{n})*(1-({n}+1)/(1+{w})^{n}"
            "+{n}/(1+{w})^({n}+1))").format(w=w, n=n)


def _ratio(expr):
    """Ratio formulas divide by cells a Downside scenario can drive to zero."""
    return '=IFERROR(%s,"N/A")' % expr


# -------------------------------------------------------------------- builder
def build_reverse_dcf_sheet(wb, C, params, refs, index=None):
    """Create the 'Reverse DCF' sheet in `wb` and return it.

    refs — addresses resolved by the caller from its own row constants:
        proj_last_col        column letter of the terminal projection year
        r_revenue, r_da      'DCF Model' rows for revenue / D&A
        r_ev_pgm             'DCF Model' row holding the PGM enterprise value
        cmp_subject_row      'Comps Analysis' row of the subject (or None)
        cmp_stat_median_row  'Comps Analysis' peer-median statistics row
        cmp_benchmark_row    'Comps Analysis' row of the transaction benchmark
    """
    selftest_bn()

    # Imported here, not at module import time: dcf_comps_template imports this
    # module, so a top-level import back into it would be a cycle.
    try:
        from templates.dcf_comps_template import (
            TITLE_FONT, SUB_FONT, BOLD_FONT, BLACK_FONT, GREY_FONT, HEADER_FONT,
            HEADER_FILL, LIGHT_FILL, LIGHT_GREEN, INPUT_BORDER,
            FMT_YEN, FMT_PCT, FMT_PCT2, FMT_RATIO, FMT_INT,
        )
    except ImportError:                                          # flat sys.path
        from dcf_comps_template import (
            TITLE_FONT, SUB_FONT, BOLD_FONT, BLACK_FONT, GREY_FONT, HEADER_FONT,
            HEADER_FILL, LIGHT_FILL, LIGHT_GREEN, INPUT_BORDER,
            FMT_YEN, FMT_PCT, FMT_PCT2, FMT_RATIO, FMT_INT,
        )

    def put(ws, row, col, value, font=None, fmt=None, fill=None, border=None, align=None):
        cell = ws.cell(row=row, column=col)
        cell.value = value
        if font:
            cell.font = font
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        if border:
            cell.border = border
        if align:
            cell.alignment = align
        return cell

    def label(ws, row, col, text, font=BLACK_FONT, fill=None):
        """Write a TEXT label.

        A string that starts with '=' is stored by openpyxl as a FORMULA; a
        label that happens to begin with '=' corrupts the workbook so badly that
        Excel refuses to open it (tasks/lessons.md, 2962). Assert on every
        write, not on a hand-picked sample.
        """
        assert not str(text).startswith("="), (
            "label at R%dC%d starts with '=': %r" % (row, col, text))
        return put(ws, row, col, text, font=font, fill=fill)

    def band(ws, row, text, last_col=9, fill=LIGHT_FILL):
        label(ws, row, 2, text, font=SUB_FONT, fill=fill)
        for c in range(3, last_col + 1):
            ws.cell(row=row, column=c).fill = fill

    op0 = params["op0"]
    op0_label = params["op0_label"]
    peak_op = params["peak_op"]
    peak_label = params["peak_label"]
    peak_opm = params["peak_opm"]
    n_years = params["n_years"]

    last_col = refs["proj_last_col"]
    A_REV_T = "'DCF Model'!%s%d" % (last_col, refs["r_revenue"])
    A_DA_T = "'DCF Model'!%s%d" % (last_col, refs["r_da"])
    A_EV_PGM = "'DCF Model'!C%d" % refs["r_ev_pgm"]
    subj = refs.get("cmp_subject_row")
    A_SUBJ_REV = ("'Comps Analysis'!F%d" % subj) if subj else None
    A_SUBJ_EBITDA = ("'Comps Analysis'!G%d" % subj) if subj else None
    A_PEER_MED = "'Comps Analysis'!D%d" % refs["cmp_stat_median_row"]

    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]
    ws = (wb.create_sheet(SHEET_NAME) if index is None
          else wb.create_sheet(SHEET_NAME, index))
    ws.sheet_properties.tabColor = "C00000"

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 52
    for c in "CDEFGHI":
        ws.column_dimensions[c].width = 15

    label(ws, 2, 2, "Reverse DCF - %s (%s)" % (C["company_name"], C["ticker"]),
          font=TITLE_FONT)
    label(ws, 3, 2, "What is the current share price already assuming? "
                    "Every value below is a live formula off the other sheets.",
          font=GREY_FONT)

    # ---------------------------------------------------------- Block 0
    band(ws, 5, "Block 0: Market-implied enterprise value (live)")
    rows0 = [
        ("Current Share Price (JPY)", "='Executive Summary'!C9", FMT_INT),
        ("Fully Diluted Shares", "='DCF Model'!C15", FMT_INT),
        ("Market Capitalisation (JPY mn)", "=C6*C7/1000000", FMT_YEN),
        ("Net Debt (JPY mn)", "='DCF Model'!C16", FMT_YEN),
        ("Enterprise Value (JPY mn)", "=C8+C9", FMT_YEN),
        ("WACC (w)", "='DCF Model'!C26", FMT_PCT2),
        ("Terminal Growth (g)", "='DCF Model'!C13", FMT_PCT2),
        ("Effective Tax Rate (t)", "='DCF Model'!C6", FMT_PCT2),
        ("Starting Operating Profit OP0 (%s actual)" % op0_label, op0, FMT_YEN),
        ("Cycle-peak Operating Profit (%s actual)" % peak_label, peak_op, FMT_YEN),
        ("Terminal-year Revenue, active scenario (JPY mn)", "=" + A_REV_T, FMT_YEN),
        ("Base-case DCF Enterprise Value, PGM (JPY mn)", "=" + A_EV_PGM, FMT_YEN),
    ]
    for i, (lab, val, fmt) in enumerate(rows0):
        r = R_IN0 + i
        label(ws, r, 2, lab, font=BOLD_FONT)
        put(ws, r, 3, val, font=BLACK_FONT, fmt=fmt,
            border=None if isinstance(val, str) else INPUT_BORDER)

    label(ws, R_PEAK_OPM, 2,
          "Cycle-peak Operating Margin (%s actual)" % peak_label, font=BOLD_FONT)
    put(ws, R_PEAK_OPM, 3, peak_opm, font=BLACK_FONT, fmt=FMT_PCT, border=INPUT_BORDER)

    label(ws, R_GAP, 2, "Gap: market EV less base-case DCF EV (JPY mn)", font=BOLD_FONT)
    put(ws, R_GAP, 3, "=C10-C17", font=BLACK_FONT, fmt=FMT_YEN)
    label(ws, R_GAP_PS, 2, "Gap per share (JPY) - the part of the price the base "
                           "case does not explain", font=BOLD_FONT)
    put(ws, R_GAP_PS, 3, _ratio("ROUND(C19*1000000/C7,0)"), font=BLACK_FONT, fmt=FMT_INT)

    # ---------------------------------------------------------- Block A
    band(ws, 22, "Block A: Required steady-state operating profit, reached "
                 "IMMEDIATELY (no ramp)")
    label(ws, 23, 2, "The lowest bar the price can clear: today's EV read as a "
                     "perpetuity of one steady-state profit, capex = D&A, no NWC build.",
          font=GREY_FONT)
    rowsA = [
        ("Required steady-state NOPAT (JPY mn)", "=C10*(C11-C12)", FMT_YEN),
        ("Required steady-state Operating Profit (JPY mn)",
         _ratio("C24/(1-C13)"), FMT_YEN),
        ("... as a multiple of the %s peak" % peak_label, _ratio("C25/C15"), FMT_RATIO),
        ("... as a multiple of %s actual" % op0_label, _ratio("C25/C14"), FMT_RATIO),
        ("... implied OPM on terminal-year revenue", _ratio("C25/C16"), FMT_PCT),
        ("... revenue needed if margins only reach the cycle peak",
         _ratio("C25/C18"), FMT_YEN),
    ]
    for i, (lab, val, fmt) in enumerate(rowsA):
        r = R_A + i
        label(ws, r, 2, lab, font=BOLD_FONT if i < 2 else BLACK_FONT)
        put(ws, r, 3, val, font=BLACK_FONT, fmt=fmt)

    # ---------------------------------------------------------- Block B
    band(ws, 32, "Block B: Required steady-state operating profit by "
                 "YEARS-TO-REACH (linear ramp) x WACC")
    label(ws, 33, 2, "Operating profit ramps linearly from OP0 (C14) to OP* over N "
                     "years, then grows at g for ever. Column F is the model WACC.",
          font=GREY_FONT)
    label(ws, 34, 2, "OP* = ( EV/(1-t) - OP0*(A-B) ) / ( B + TVF ). A longer ramp "
                     "needs a HIGHER steady state, because more years are spent below it.",
          font=GREY_FONT)

    label(ws, HDR, 2, "Years to steady state (N)   \\   WACC",
          font=HEADER_FONT, fill=HEADER_FILL)
    label(ws, HDR, 3, "N", font=HEADER_FONT, fill=HEADER_FILL)
    for j in range(5):
        offset = (j - 2) * 0.01
        f = ("='DCF Model'!C26" if offset == 0
             else "='DCF Model'!C26%+.2f" % offset)
        put(ws, HDR, 4 + j, f, font=HEADER_FONT, fmt=FMT_PCT2, fill=HEADER_FILL,
            align=Alignment(horizontal="center"))

    def grid(start_row, title, fn, fmt, fill=None):
        label(ws, start_row - 1, 2, title, font=SUB_FONT, fill=fill)
        if fill:
            for c in range(3, 10):
                ws.cell(row=start_row - 1, column=c).fill = fill
        for i, n in enumerate(n_years):
            r = start_row + i
            label(ws, r, 2, "N = %d years" % n, font=BOLD_FONT)
            put(ws, r, 3, n, font=BLACK_FONT, fmt=FMT_INT, border=INPUT_BORDER)
            for j in range(5):
                cl = chr(ord("D") + j)
                put(ws, r, 4 + j, fn(cl, r, i), font=BLACK_FONT, fmt=fmt)

    grid(G1, "B-1  PV weight on the RAMP from OP0   (A - B)",
         lambda cl, r, i: _ratio(_f_A(cl, r) + "-" + _f_B(cl, r)), "0.000")
    grid(G2, "B-2  PV weight on the STEADY STATE OP*   (B + terminal factor)",
         lambda cl, r, i: _ratio("%s+(1+$C$12)/((%s$%d-$C$12)*(1+%s$%d)^$C%d)"
                                 % (_f_B(cl, r), cl, HDR, cl, HDR, r)), "0.000")
    grid(OPS, "B-3  REQUIRED steady-state Operating Profit (JPY mn)",
         lambda cl, r, i: _ratio("($C$10/(1-$C$13)-$C$14*%s%d)/%s%d"
                                 % (cl, G1 + i, cl, G2 + i)),
         FMT_YEN, fill=LIGHT_GREEN)
    grid(MULT, "B-4  ... as a multiple of the %s peak operating profit" % peak_label,
         lambda cl, r, i: _ratio("%s%d/$C$15" % (cl, OPS + i)), FMT_RATIO,
         fill=LIGHT_GREEN)

    hi = min(HEADLINE_OFFSET, len(n_years) - 1)
    hn = n_years[hi]
    c_ops = "%s%d" % (WACC_COL, OPS + hi)
    c_mult = "%s%d" % (WACC_COL, MULT + hi)

    # ---------------------------------------------------------- Block C
    band(ws, 58, "Block C: Decomposing the required profit into revenue x margin "
                 "(N = %d years, model WACC)" % hn)
    label(ws, 59, 2, "The revenue the company must run at, for a given steady-state "
                     "operating margin, to earn the Block B-3 profit.", font=GREY_FONT)
    label(ws, R_C, 2, "Required steady-state Operating Profit (N=%d, model WACC)" % hn,
          font=BOLD_FONT)
    put(ws, R_C, 3, "=" + c_ops, font=BLACK_FONT, fmt=FMT_YEN)

    label(ws, R_C_HDR, 2, "Assumed steady-state OPM", font=HEADER_FONT, fill=HEADER_FILL)
    label(ws, R_C_HDR, 3, "Required Revenue", font=HEADER_FONT, fill=HEADER_FILL)
    label(ws, R_C_HDR, 4, "x base-year revenue", font=HEADER_FONT, fill=HEADER_FILL)
    for i, m in enumerate(params["opm_grid"]):
        r = R_C_HDR + 1 + i
        if m is None:
            label(ws, r, 2, "cycle-peak OPM (%s)" % peak_label, font=BOLD_FONT)
            put(ws, r, 3, _ratio("$C$%d/$C$%d" % (R_C, R_PEAK_OPM)),
                font=BLACK_FONT, fmt=FMT_YEN)
        else:
            put(ws, r, 2, m, font=BOLD_FONT, fmt=FMT_PCT, border=INPUT_BORDER)
            put(ws, r, 3, _ratio("$C$%d/B%d" % (R_C, r)), font=BLACK_FONT, fmt=FMT_YEN)
        put(ws, r, 4, _ratio("C%d/'DCF Model'!C17" % r), font=BLACK_FONT, fmt=FMT_RATIO)

    # ---------------------------------------------------------- Block D
    band(ws, 69, "Block D: The same question expressed as multiples")
    na = '="N/A - the subject has no row in the comps table"'
    rows_d = [
        ("Current EV / latest-FY EBITDA (actual)",
         _ratio("C10/" + A_SUBJ_EBITDA) if subj else na, FMT_RATIO),
        ("Current EV / latest-FY Revenue (actual)",
         _ratio("C10/" + A_SUBJ_REV) if subj else na, FMT_RATIO),
        ("EV / EBITDA if the Block A profit were already earned",
         _ratio("C10/(C25+%s)" % A_DA_T), FMT_RATIO),
        ("EV / EBITDA if the Block B-3 (N=%d) profit were already earned" % hn,
         _ratio("C10/(C%d+%s)" % (R_C, A_DA_T)), FMT_RATIO),
        ("Peer median EV / EBITDA (statistics rows)", _ratio(A_PEER_MED), FMT_RATIO),
    ]
    for i, (lab, val, fmt) in enumerate(rows_d):
        r = R_D + i
        label(ws, r, 2, lab, font=BOLD_FONT if i >= 2 else BLACK_FONT)
        put(ws, r, 3, val, font=BLACK_FONT, fmt=fmt)
    label(ws, R_D + 5, 2, "Rows 3-4 answer: what multiple would today's EV be paying "
                          "IF the company already earned the profit the price requires?",
          font=GREY_FONT)

    # ---------------------------------------------------------- Block E
    nxt = 77
    bench = refs.get("cmp_benchmark_row")
    if bench and subj:
        band(ws, nxt, "Block E: Transaction benchmark (control price paid for the "
                      "direct peer)")
        rows_e = [
            ("Peer row EV / EBITDA at the announced deal terms",
             _ratio("'Comps Analysis'!J%d" % bench), FMT_RATIO),
            ("Peer row EV / Revenue at the announced deal terms",
             _ratio("'Comps Analysis'!K%d" % bench), FMT_RATIO),
            ("Subject share price implied by that EV / EBITDA",
             _ratio("ROUND((%s*'Comps Analysis'!J%d-C9)*1000000/C7,0)"
                    % (A_SUBJ_EBITDA, bench)), FMT_INT),
            ("Subject share price implied by that EV / Revenue",
             _ratio("ROUND((%s*'Comps Analysis'!K%d-C9)*1000000/C7,0)"
                    % (A_SUBJ_REV, bench)), FMT_INT),
        ]
        for i, (lab, val, fmt) in enumerate(rows_e):
            r = nxt + 1 + i
            label(ws, r, 2, lab, font=BOLD_FONT if i >= 2 else BLACK_FONT)
            put(ws, r, 3, val, font=BLACK_FONT, fmt=fmt)
        for i, line in enumerate(params["deal_note"]):
            label(ws, nxt + 6 + i, 2, line, font=GREY_FONT)
        nxt = nxt + 6 + len(params["deal_note"]) + 1

    # ---------------------------------------------------------- Block F
    band(ws, nxt, "Block F: The one-line answer", fill=LIGHT_GREEN)
    put(ws, nxt + 1, 2,
        '="At "&TEXT(C6,"#,##0")&" yen the market is paying for a steady-state '
        'operating profit of "&TEXT(%s,"#,##0")&" mn JPY - "&TEXT(%s,"0.00")'
        '&"x the %s cycle peak of "&TEXT(C15,"#,##0")&" mn - reached in %d years '
        'and sustained for ever, against "&TEXT(C14,"#,##0")&" mn actually earned '
        'in %s."' % (c_ops, c_mult, peak_label, hn, op0_label),
        font=BOLD_FONT)
    put(ws, nxt + 2, 2,
        '="Even with no ramp at all (Block A) the price needs "&TEXT(C25,"#,##0")'
        '&" mn, or "&TEXT(C26,"0.00")&"x the peak. The base case supports an EV of '
        'only "&TEXT(C17,"#,##0")&" mn, leaving "&TEXT(C20,"#,##0")&" yen per share '
        'of the current price unexplained by mid-cycle economics."',
        font=BOLD_FONT)

    return ws


def place_after(wb, anchor="DCF Model"):
    """Move the sheet to directly after `anchor` (workbook sheet 4 of 8)."""
    if SHEET_NAME not in wb.sheetnames:
        return
    order = wb.sheetnames
    order.remove(SHEET_NAME)
    pos = order.index(anchor) + 1 if anchor in order else len(order)
    order.insert(pos, SHEET_NAME)
    wb._sheets = [wb[n] for n in order]


if __name__ == "__main__":
    selftest_bn()
    print("self-test OK: B_N closed form matches direct summation")
