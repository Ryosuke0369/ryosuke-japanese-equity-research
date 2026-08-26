"""SpiderPlus (4192) Market Analysis — narrative追加版（4ページ化）。
既存 4192_market_analysis_20260612.xlsx と衝突しないよう出力日付を変更。
入力DCF: models/4192_DCF_Model_20260612.xlsx

配置: scripts/run_market_analysis_4192_v2.py
実行: python scripts/run_market_analysis_4192_v2.py  （リポジトリのルートから）
"""
import os, sys

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
sys.path.insert(0, os.path.join(PROJECT_ROOT, 'templates'))
from market_analysis_template import generate_market_analysis_excel

dcf_path = os.path.join(PROJECT_ROOT, 'models', '4192_DCF_Model_20260612.xlsx')
out_path = os.path.join(PROJECT_ROOT, 'reports', '4192_market_analysis_20260618.xlsx')

config = {
    'ticker': '4192.T',
    'company_name': 'SpiderPlus & Co.',
    'current_price': 244,
    'price_3m_ago': 244, 'price_1m_ago': 244, 'price_high': 244, 'price_low': 244,
    'margin_buy': 1000, 'margin_sell': 200, 'margin_sell_peak_6m': 400,
    'company_rev_growth': 0.205,
    'company_op_growth': -4.204,
    # ── Narrative Stage（4ページ目）──
    # 業績インパクトの tam/share/opm/current_operating_profit は意図的に省略。
    # スパイダーは営業利益ほぼゼロのため、入れるとゼロ割りで「450倍化余地」等の
    # 無意味な値が出る（文書の教訓：ゼロ近傍を分母にした割り算の罠）。
    # 省くと「算出不可」と正しく表示され、Stage判定は出る。
    'narrative': {
        'date': '2026-06-12',
        'axis_1_label_status':        0,
        'axis_2a_catalyst_potential': 2,
        'axis_2b_catalyst_realization':1,
        'axis_3_diffusion_stage':     0,
        'axis_4_earnings_materiality':1,
        'axis_5_narrative_durability':2,
        'axis_1_note':  '赤字SaaSラベルのまま、黒字優良SaaSへの書き換え未開始',
        'axis_2a_note': '黒字化・DIS/インフォマート提携・鴻池組AI、潜在カタリスト豊富',
        'axis_2b_note': 'Q4黒字化は出たが提携寄与は中期以降、ARR+9%で再加速未実証',
        'axis_3_note':  '情報通の投資家すら知らない。まだ発見されていない最初期',
        'axis_4_note':  '黒字化目前だが営業利益ほぼゼロ、利益の寄与はまだ小さい',
        'axis_5_note':  'スイッチングコスト(解約0.9%)・データ蓄積で時間とともに強くなる堀',
        'fundamental_verdict': 'HOLD',
    },
}

if not os.path.exists(dcf_path):
    print('DCF Excel not found:', dcf_path); sys.exit(1)
if os.path.exists(out_path):
    print('既に存在:', out_path, '— 削除かリネームしてください'); sys.exit(1)

os.makedirs(os.path.dirname(out_path), exist_ok=True)
out = generate_market_analysis_excel(config, out_path, dcf_excel_path=dcf_path)
print('Generated:', out)
