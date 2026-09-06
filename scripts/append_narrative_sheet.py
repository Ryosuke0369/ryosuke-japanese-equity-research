"""Append a 'Narrative Stage' sheet to an existing market-analysis workbook.

Reads a narrative JSON (produced by run_narrative_4192.py) and writes a
data-driven Block 1-5 sheet — 6-axis scores, stage judgment, earnings impact +
TAM sensitivity, and the fundamental x stage Verdict Matrix (both BUY and
CAUTION rows) — into the workbook, matching the styling of the other sheets.

Reusable across tickers:
    python scripts/append_narrative_sheet.py 4192
    python scripts/append_narrative_sheet.py 5258 --date 2026-06-03

Sheet order ends as: Implied Growth Analysis -> Market Scorecard -> Narrative Stage.
"""
import sys
import os
import json
import argparse

import openpyxl
from openpyxl.styles import Font, Alignment

# Reuse the exact style constants the other sheets use.
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates'))
from market_analysis_template import (  # noqa: E402
    TITLE_FONT, TITLE_FILL, HEADER_FONT, HEADER_FILL,
    SUBHEADER_FONT, SUBHEADER_FILL, INPUT_FILL, OUTPUT_FILL,
    HIGHLIGHT_FILL, VERDICT_FILL, WARNING_FILL, BORDER,
)

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

AXIS_DISPLAY = [
    ('1_label_status',          'Axis 1: Label / Spark'),
    ('2a_catalyst_potential',   'Axis 2A: Catalyst Potential (fuel)'),
    ('2b_catalyst_realization', 'Axis 2B: Catalyst Realization (点火)'),
    ('3_diffusion_stage',       'Axis 3: Diffusion / Reach'),
    ('4_earnings_materiality',  'Axis 4: Earnings Materiality'),
    ('5_narrative_durability',  'Axis 5: Narrative Durability'),
]

WRAP_TOP = Alignment(wrap_text=True, vertical='top')


def _ticker_disp(ticker):
    return ticker if ticker.endswith('.T') else f'{ticker}.T'


def build_narrative_sheet(wb, data):
    """Create/replace the 'Narrative Stage' sheet from narrative JSON `data`."""
    if 'Narrative Stage' in wb.sheetnames:
        del wb['Narrative Stage']
    ws = wb.create_sheet('Narrative Stage')  # appended last -> correct order

    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 34
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 70

    inp = data['input']
    res_buy = data['results']['BUY']
    res_caut = data['results']['CAUTION']
    common = res_buy  # stage / impact / headroom are identical across verdicts

    company = inp.get('company_name', '')
    ticker = _ticker_disp(inp.get('ticker', ''))

    # ── Title ──
    ws['B2'] = f"Narrative Stage Assessment - {company} ({ticker})"
    ws['B2'].font = TITLE_FONT
    ws['B2'].fill = TITLE_FILL
    ws.merge_cells('B2:E2')
    ws['B3'] = inp.get('date', '')
    ws['B3'].font = Font(name='Calibri', size=9, italic=True)

    def header(row, text, span='B{r}:E{r}'):
        ws[f'B{row}'] = text
        ws[f'B{row}'].font = HEADER_FONT
        ws[f'B{row}'].fill = HEADER_FILL
        ws.merge_cells(span.format(r=row))

    def subheaders(row, cols):
        for col, txt in cols:
            ws[f'{col}{row}'] = txt
            ws[f'{col}{row}'].font = SUBHEADER_FONT
            ws[f'{col}{row}'].fill = SUBHEADER_FILL

    # ════════ Block 1: 6-Axis Scores ════════
    header(5, 'Block 1: 6-Axis Scores（6軸スコア）')
    subheaders(6, [('B', 'Axis'), ('C', 'Score (0/1/2)'), ('E', 'Note')])
    rate_limit = common['rate_limiting_axis']
    r = 7
    for key, label in AXIS_DISPLAY:
        core = key.split('_')[0]  # '1','2a','2b','3','4','5'
        is_rl = f'axis_{core}' in rate_limit
        ws[f'B{r}'] = label + ('  【律速軸】' if is_rl else '')
        ws[f'C{r}'] = common['axis_scores'][key]
        ws[f'C{r}'].alignment = Alignment(horizontal='center')
        ws[f'C{r}'].fill = INPUT_FILL
        ws[f'E{r}'] = common['axis_notes'][key]
        ws[f'E{r}'].alignment = WRAP_TOP
        if is_rl:
            ws[f'B{r}'].fill = HIGHLIGHT_FILL
        for c in 'BCDE':
            ws[f'{c}{r}'].border = BORDER
        r += 1
    ws[f'B{r}'] = 'Total Score'
    ws[f'B{r}'].font = SUBHEADER_FONT
    ws[f'C{r}'] = f"{common['total_score']} / 12"
    ws[f'C{r}'].font = Font(name='Calibri', size=11, bold=True)
    ws[f'C{r}'].fill = HIGHLIGHT_FILL
    ws[f'C{r}'].alignment = Alignment(horizontal='center')
    for c in 'BCDE':
        ws[f'{c}{r}'].border = BORDER

    # ════════ Block 2: Stage 判定 ════════
    r += 2
    header(r, 'Block 2: Stage 判定')
    r += 1
    stage_rows = [
        ('Stage', f"{common['stage']}  →  {common['stage_label']}", OUTPUT_FILL),
        ('Rate-limiting axis（律速軸）', common['rate_limiting_axis'], HIGHLIGHT_FILL),
        ('理由 (Reason)', common['rate_limiting_reason'], None),
        ('Earnings cap applied (軸4==0頭打ち)',
         'Yes' if common['earnings_cap_applied'] else 'No', None),
        ('Headroom signal（軸1 vs 軸3）', common['headroom_signal'], OUTPUT_FILL),
        ('Headroom note', common['headroom_note'], None),
    ]
    for label, val, fill in stage_rows:
        ws[f'B{r}'] = label
        ws[f'B{r}'].font = SUBHEADER_FONT
        ws[f'C{r}'] = val
        ws.merge_cells(f'C{r}:E{r}')
        ws[f'C{r}'].alignment = WRAP_TOP
        if fill:
            ws[f'C{r}'].fill = fill
        for c in 'BCDE':
            ws[f'{c}{r}'].border = BORDER
        r += 1

    # ════════ Block 3: Earnings Impact ════════
    r += 1
    header(r, 'Block 3: Earnings Impact（利益インパクト = TAM × share × OPM）')
    r += 1
    impact = common.get('earnings_impact_oku')
    uplift = common.get('earnings_uplift_ratio')
    cur_op = inp.get('current_operating_profit_oku')
    impact_rows = [
        ('Mid case (TAM {:.0f}億 × {:.0f}% × OPM{:.0f}%)'.format(
            inp.get('tam_oku_jpy', 0), inp.get('expected_share_pct', 0),
            inp.get('segment_opm_pct', 0)),
         f"{impact:.1f} 億円" if impact is not None else 'n/a'),
        ('earnings_uplift_ratio（対現営利 {:.1f}億）'.format(cur_op or 0),
         f"{uplift:.1f} 倍" if uplift is not None else 'n/a'),
    ]
    for label, val in impact_rows:
        ws[f'B{r}'] = label
        ws[f'B{r}'].font = SUBHEADER_FONT
        ws[f'C{r}'] = val
        ws[f'C{r}'].fill = OUTPUT_FILL
        for c in 'BCDE':
            ws[f'{c}{r}'].border = BORDER
        r += 1

    # TAM sensitivity table
    r += 1
    ws[f'B{r}'] = 'TAM 感度 (impact = TAM × share × OPM)'
    ws[f'B{r}'].font = SUBHEADER_FONT
    r += 1
    subheaders(r, [('B', 'TAM (億円)'), ('C', 'Share'),
                   ('D', 'Impact (億)'), ('E', 'Uplift (×現営利)')])
    for c in 'BCDE':
        ws[f'{c}{r}'].border = BORDER
    r += 1
    for row in data.get('tam_sensitivity', []):
        ws[f'B{r}'] = row['tam_oku']
        ws[f'B{r}'].number_format = '#,##0'
        ws[f'C{r}'] = row['share_pct'] / 100.0
        ws[f'C{r}'].number_format = '0%'
        ws[f'D{r}'] = row['impact_oku']
        ws[f'D{r}'].number_format = '#,##0.0'
        ws[f'E{r}'] = (f"{row['uplift_ratio']:.0f}倍"
                       if row.get('uplift_ratio') is not None else 'n/a')
        for c in 'BCDE':
            ws[f'{c}{r}'].border = BORDER
        r += 1

    # ════════ Block 4: Verdict Matrix ════════
    r += 1
    header(r, 'Block 4: Verdict Matrix（2軸: ファンダ × ステージ）')
    r += 1
    subheaders(r, [('B', 'Fundamental Verdict'), ('C', 'Stage'),
                   ('D', 'Final Verdict'), ('E', 'Rationale')])
    for c in 'BCDE':
        ws[f'{c}{r}'].border = BORDER
    r += 1
    matrix_rows = [
        ('BUY (Comps/Exit=割安)', res_buy),
        ('CAUTION (逆算DCF/PGM=割高)', res_caut),
    ]
    for fund_label, res in matrix_rows:
        ws[f'B{r}'] = fund_label
        ws[f'C{r}'] = f"Stage {res['stage']}"
        ws[f'C{r}'].alignment = Alignment(horizontal='center')
        ws[f'D{r}'] = res['final_verdict']
        ws[f'D{r}'].font = Font(name='Calibri', size=11, bold=True)
        ws[f'D{r}'].fill = VERDICT_FILL
        ws[f'E{r}'] = res['verdict_rationale']
        ws[f'E{r}'].alignment = WRAP_TOP
        for c in 'BCDE':
            ws[f'{c}{r}'].border = BORDER
        r += 1

    # ════════ Block 5: Note (interpretation) ════════
    r += 1
    header(r, 'Block 5: Note（解釈）')
    r += 1
    notes = [
        'この銘柄はナラティブ段階(Stage1=未拡散)ではなく、どの評価レンズ(Comps/Exit vs '
        '逆算DCF/PGM)を採るかで最終判断が STRONG_BUY ⇔ AVOID に分かれる。',
        '軸2B(点火認定)が律速。黒字転換を点火と見て 2 に上げると Stage2 となり、'
        'verdict は BUY / CAUTION へ穏当化する。投資家本人の点火認定次第で感度が高い。',
        f'参考: Stage {common["stage"]} ({common["stage_label"]})、'
        f'Headroom = {common["headroom_signal"]}、'
        f'利益インパクト ≈ {impact:.1f}億円 ({uplift:.0f}倍化余地)。',
    ]
    for txt in notes:
        ws[f'B{r}'] = txt
        ws.merge_cells(f'B{r}:E{r}')
        ws[f'B{r}'].alignment = WRAP_TOP
        ws.row_dimensions[r].height = 30
        r += 1

    return ws


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('ticker', nargs='?', default='4192')
    ap.add_argument('--date', default='20260603')
    ap.add_argument('--xlsx', default=None)
    ap.add_argument('--json', default=None)
    args = ap.parse_args()

    xlsx_path = args.xlsx or os.path.join(
        PROJECT_ROOT, 'reports', f'{args.ticker}_market_analysis_{args.date}.xlsx')
    json_path = args.json or os.path.join(
        PROJECT_ROOT, 'data', 'narrative', f'{args.ticker}_narrative_{args.date}.json')

    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(xlsx_path)
    if not os.path.exists(json_path):
        raise FileNotFoundError(json_path)

    with open(json_path, encoding='utf-8') as f:
        data = json.load(f)

    wb = openpyxl.load_workbook(xlsx_path)
    build_narrative_sheet(wb, data)
    wb.save(xlsx_path)
    print(f"Sheets now: {wb.sheetnames}")
    print(f"Saved: {os.path.normpath(xlsx_path)}")


if __name__ == '__main__':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    main()
