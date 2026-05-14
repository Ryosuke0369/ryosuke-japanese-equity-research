"""DMW Corporation (6365.T) Market Analysis Script"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates'))

from market_analysis_template import generate_market_analysis_excel

config = {
    'ticker': '6365.T',
    'company_name': 'DMW Corporation',
    'current_price': 5490,
    'segment_layout': {
        'segments': [
            {'name': 'Public Sector', 'dcf_fy26_cell': 'F6', 'dcf_growth_base_row': 34, 'dcf_opm_base_row': 41},
            {'name': 'Private Sector', 'dcf_fy26_cell': 'F11', 'dcf_growth_base_row': 48, 'dcf_opm_base_row': 55},
            {'name': 'Overseas Desalination', 'dcf_fy26_cell': 'F16', 'dcf_growth_base_row': 62, 'dcf_opm_base_row': 69},
        ]
    },
    'price_3m_ago': 5800,
    'price_1m_ago': 5700,
    'price_high': 6000,
    'price_low': 5400,
    'margin_buy': 9800,
    'margin_sell': 1,
    'margin_sell_peak_6m': 5000,
    'company_op_growth': 0.05,
    'company_rev_growth': 0.03,
}

output = generate_market_analysis_excel(
    config,
    'reports/6365_market_analysis_20260509.xlsx',
    dcf_excel_path='models/6365_DCF_Model_20260413.xlsx'
)
print('Generated:', output)
