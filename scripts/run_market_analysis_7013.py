"""IHI Market Analysis Script (2026/5/13)"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates'))

from market_analysis_template import generate_market_analysis_excel

config = {
    'ticker': '7013.T',
    'company_name': 'IHI Corporation',
    'current_price': 2824,
    'segment_layout': {
        'segments': [
            {'name': 'Resources Energy Environment',
             'dcf_fy26_cell': 'E6',
             'dcf_growth_base_row': 39,
             'dcf_opm_base_row': 46},
            {'name': 'Social Infrastructure',
             'dcf_fy26_cell': 'E11',
             'dcf_growth_base_row': 53,
             'dcf_opm_base_row': 60},
            {'name': 'Industrial Systems GPM',
             'dcf_fy26_cell': 'E16',
             'dcf_growth_base_row': 67,
             'dcf_opm_base_row': 74},
            {'name': 'Aero Engine Space Defense',
             'dcf_fy26_cell': 'E21',
             'dcf_growth_base_row': 81,
             'dcf_opm_base_row': 88},
        ]
    },
    'price_3m_ago': 2824,
    'price_1m_ago': 2824,
    'price_high': 2824,
    'price_low': 2824,
    'margin_buy': 22426,
    'margin_sell': 611,
    'margin_sell_peak_6m': 611,
    'company_op_growth': 0.10,
    'company_rev_growth': 0.04,
}

output = generate_market_analysis_excel(
    config,
    'reports/7013_market_analysis_20260513.xlsx',
    dcf_excel_path='models/7013_DCF_Model_20260328.xlsx'
)
print('Generated:', output)
