"""ELEMENTS (5246) Market Analysis — reverse DCF (implied growth/alpha) +
4-factor scorecard.

Usage:
    python scripts/run_market_analysis_5246.py [--dcf models/5246_DCF_Model_YYYYMMDD.xlsx]

The source DCF defaults to the LATEST dated model matching
models/5246_DCF_Model_<YYYYMMDD>.xlsx (models/archive/ and non-dated names are
ignored); pass --dcf to pin a specific file. The DCF must already be recalced
(WACC, D&A/Capex/dNWC, PGM) — see scripts/recalc_excel_com.py; an unrecalced
model is rejected by market_analysis_template with an error.

5246 is a single-business eKYC/認証 SaaS (no Segment Analysis sheet), so this
run relies on the single-segment fallback in
market_analysis_template.extract_dcf_data: the Base scenario (growth / OPM) is
read from the DCF Model 'Scenario Input Matrix' and FY actuals from
'Financial Statements'. No segment_layout is passed.
"""
import argparse
import glob
import os
import re
import sys

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
sys.path.insert(0, os.path.join(PROJECT_ROOT, 'templates'))

from market_analysis_template import generate_market_analysis_excel

TICKER = '5246'


def latest_dcf_model(ticker):
    """Return (path, YYYYMMDD) of the newest dated DCF model in models/."""
    dated = {}
    for p in glob.glob(os.path.join(PROJECT_ROOT, 'models', f'{ticker}_DCF_Model_*.xlsx')):
        m = re.fullmatch(rf'{ticker}_DCF_Model_(\d{{8}})\.xlsx', os.path.basename(p))
        if m:
            dated[m.group(1)] = p
    if not dated:
        sys.exit(f'ERROR: no dated DCF model models/{ticker}_DCF_Model_<YYYYMMDD>.xlsx found. '
                 f'Generate one with scripts/generate_dcf.py and recalc it first.')
    date_str = max(dated)
    return dated[date_str], date_str


parser = argparse.ArgumentParser(description=f'Market analysis for {TICKER}')
parser.add_argument('--dcf', help='source DCF xlsx (default: latest dated '
                                  f'models/{TICKER}_DCF_Model_*.xlsx)')
args = parser.parse_args()

if args.dcf:
    dcf_path = args.dcf
    m = re.search(r'_(\d{8})\.xlsx$', os.path.basename(dcf_path))
    date_str = m.group(1) if m else 'manual'
else:
    dcf_path, date_str = latest_dcf_model(TICKER)
print(f'Source DCF: {dcf_path}')

config = {
    'ticker': '5246.T',
    'company_name': 'ELEMENTS',
    'current_price': 557,            # 2026-06-12 終値 (yfinance, ユーザー承認 2026-06-12, overridesと同値)
    # No 'segment_layout' -> single-segment fallback (Base from Scenario Matrix).

    # ── Factor 2: Price Momentum ──
    # NOTE: placeholder = current price (no real 3M/1M history wired in yet);
    # yields 0% momentum -> neutral score 0. Replace with actual price history.
    'price_3m_ago': 557,
    'price_1m_ago': 557,
    'price_high':   557,
    'price_low':    557,

    # ── Factor 3: Margin trading (信用残) ──
    # NOTE: neutral placeholders (信用倍率 5x, 売残充実度 0.5) -> score 0.
    # Replace with real 信用残 data.
    'margin_buy':          1000,     # 信用買残 (千株)
    'margin_sell':          200,     # 信用売残 (千株)
    'margin_sell_peak_6m':  400,     # 過去6ヶ月 売残ピーク (千株)

    # ── Factor 4: Forecast Gap (vs company) ──
    # Revenue: our Base FY1 +33% = company FY2026/11 plan range mid (+33%) -> gap ~0.
    'company_rev_growth': 0.33,
    # OP growth %: undefined for a loss->profit transition (FY2025/11 OP = -215mn).
    # The template computes my_op_growth = base_FY1_OP / FY_actual_OP - 1
    #   = 103.607 / -215 - 1 = -1.4819 (a negative-base artifact, not a real -148%).
    # Set company_op_growth to the same value so the gap is ~0 (neutral); the
    # economically meaningful comparison is the revenue row above.
    'company_op_growth': -1.4819,
}

out_path = os.path.join(PROJECT_ROOT, 'reports', f'{TICKER}_market_analysis_{date_str}.xlsx')
if os.path.exists(out_path):
    sys.exit(f'ERROR: {out_path} already exists. reports/ files are never overwritten; '
             f'rename or remove the existing file deliberately, or pass --dcf to a model '
             f'with a different date.')

out = generate_market_analysis_excel(config, out_path, dcf_excel_path=dcf_path)
print('Generated:', out)
