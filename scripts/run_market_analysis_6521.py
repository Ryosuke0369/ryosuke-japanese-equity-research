"""Oxide (6521) Market Analysis — reverse DCF (implied growth/alpha) +
4-factor scorecard + 6-axis Narrative Stage.

Usage:
    python scripts/run_market_analysis_6521.py [--verdict CAUTION]

Input DCF: models/6521_DCF_Model_20260821.xlsx (user-edited final; NOT modified
by this script). 6521 is a single-segment optical/single-crystal maker with no
'Segment Analysis' sheet, so no segment_layout is passed — the template's
single-segment fallback reads the Base scenario from the DCF Model
'Scenario Input Matrix' and FY actuals from 'Financial Statements'.

--verdict sets narrative.fundamental_verdict. Per the sister-template SOP §4-3
this must be the Block 2-4 scorecard verdict, which is only known after a first
generation pass, so the script is run twice: once to read 'Market Scorecard'!G12,
then again with the observed verdict.
"""
import argparse
import os
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
sys.path.insert(0, os.path.join(PROJECT_ROOT, 'templates'))

from market_analysis_template import generate_market_analysis_excel

DCF_PATH = os.path.join(PROJECT_ROOT, 'models', '6521_DCF_Model_20260821.xlsx')
OUT_PATH = os.path.join(PROJECT_ROOT, 'models', '6521_market_analysis_20260821.xlsx')

parser = argparse.ArgumentParser()
parser.add_argument('--verdict', default='CAUTION', choices=['BUY', 'HOLD', 'CAUTION'],
                    help="narrative.fundamental_verdict (= Block 2-4 scorecard verdict)")
args = parser.parse_args()

# ── Pre-flight: the DCF must be recalculated and on the Base scenario ──
import openpyxl
import warnings
warnings.filterwarnings('ignore')

_wv = openpyxl.load_workbook(DCF_PATH, data_only=True)
_dcf, _exe = _wv['DCF Model'], _wv['Executive Summary']
_checks = {
    "'DCF Model'!C26 (WACC)": _dcf['C26'].value,
    "'Executive Summary'!C10 (Target Mid)": _exe['C10'].value,
    "'DCF Model'!C55 (PGM)": _dcf['C55'].value,
    "'DCF Model'!C64 (Exit)": _dcf['C64'].value,
}
_missing = [k for k, v in _checks.items() if v is None]
if _missing:
    sys.exit(
        "ERROR: the input DCF has no cached formula values for: "
        + ', '.join(_missing)
        + f"\n  {DCF_PATH}\n"
        "  Open it in Excel, press F9 (full recalculation), save, and re-run.\n"
        "  (Or run: python scripts/recalc_excel_com.py <path>)\n"
        "  Refusing to continue with placeholder values."
    )
_d27 = _dcf['D27'].value
if _d27 != 1:
    sys.exit(f"ERROR: 'DCF Model'!D27 (active scenario index) = {_d27!r}, expected 1 (Base). "
             f"Active scenario is {_dcf['C27'].value!r}. D&A / Capex / dNWC are read from the "
             f"active scenario rows, so market_analysis must be run on Base. "
             f"Switch the DCF dropdown back to Base, recalc, save, and re-run.")
print('Pre-flight OK — input DCF is recalculated and on the Base scenario.')
for k, v in _checks.items():
    print(f'  {k} = {v}')
print(f"  'DCF Model'!C19 (Stub Fraction) = {_dcf['C19'].value}")
print(f"  'Comps Analysis' EV/EBITDA implied = {_wv['Comps Analysis']['C28'].value}")
print(f"  narrative.fundamental_verdict = {args.verdict}")
print()

config = {
    'ticker': '6521.T',
    'company_name': 'Oxide Corporation',
    'current_price': 3350,               # 2026-08-21 (prompt-specified; yfinance previousClose)
    # No 'segment_layout' -> single-segment fallback (Base read from Scenario Input Matrix).
    # No 'entry_price'    -> not provided.
    # No 'price_targets'  -> actual trading levels are never committed to the repo.

    # ── Factor 2: Price Momentum ── yfinance 6521.T daily closes, fetched 2026-08-21
    'price_3m_ago': 5590,                # 2026-05-21 close
    'price_1m_ago': 3425,                # 2026-07-21 close
    'price_high':   7420,                # 52w intraday high, 2026-05-12
    'price_low':    1394,                # 52w intraday low,  2025-12-18

    # ── Factor 3: Margin trading (信用残) ──
    # Only 信用買残 was supplied. margin_sell / margin_sell_peak_6m are deliberately
    # omitted: the template scores Factor 3 as 0 (neutral) when any of the three is
    # blank, and dummy values are forbidden.
    'margin_buy': 699.4,                 # 千株

    # ── Factor 4: Forecast Gap (vs company) ──
    'company_rev_growth': -0.021,        # 会社予想 9,828 vs FY2026/2 実績 10,040
    'company_op_growth':   0.719,        # 会社予想 933   vs FY2026/2 実績 543

    # ── Reverse-comps forward multiple denominator ──
    'forward_revenue': 9828,

    # ── Block 5: Narrative Stage (6 axes) ──
    'narrative': {
        'date': '2026-08-21',
        'axis_1_label_status':          1,
        'axis_2a_catalyst_potential':   2,
        'axis_2b_catalyst_realization': 1,
        'axis_3_diffusion_stage':       1,
        'axis_4_earnings_materiality':  1,
        'axis_5_narrative_durability':  2,
        'axis_1_note':
            '市場の呼び名はまだ「半導体検査用単結晶メーカー」。量子・核融合向け光学材料への'
            '書き換えは一部で始まった段階。',
        'axis_2a_note':
            '量子コンピュータ向け波長変換結晶、核融合(レーザー)向け大型結晶、半導体検査の'
            '世代交代と、触媒の源泉は複数ある。',
        'axis_2b_note':
            'Raicol売却による財務改善とQ1黒字化は着弾済みだが、新領域の大型受注・提携発表は'
            '続報待ち。',
        'axis_3_note':
            '東証グロース小型(時価総額389億円)。個人投資家の一部に浸透、セルサイド・機関の'
            '継続カバレッジは限定的。',
        'axis_4_note':
            'ターンアラウンドは数字に出ている(営業利益率5.4%→Q1 9.1%)が、新領域(売上の30%)'
            'の利益寄与はまだ小さい。',
        'axis_5_note':
            '単結晶は参入障壁が高く、量子・核融合は10年単位のテーマ。',

        # 業績インパクト(楽観ケース、Stage判定とは独立)
        'tam_oku_jpy': 1000,
        'expected_share_pct': 10,
        'segment_opm_pct': 20,
        'current_operating_profit_oku': 5.4,
        '_tam_note': 'TAM 1,000億円は量子・核融合向け光学材料の推定値。'
                     '一次ソースの市場調査に基づく確定値ではない。',
        'fundamental_verdict': args.verdict,
    },
}

out = generate_market_analysis_excel(config, OUT_PATH, dcf_excel_path=DCF_PATH)
print('\nGenerated:', out)
