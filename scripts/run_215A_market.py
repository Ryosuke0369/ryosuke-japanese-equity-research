"""
215A タイミー 逆算DCF・逆算Comps 実行スクリプト
market_analysis_template.py を使用。
前提：215A_DCF_Model_20260617.xlsx が models/ にあること。

配置: scripts/run_215A_market.py
テンプレ: templates/market_analysis_template.py
"""
import os, sys

_here = os.path.dirname(os.path.abspath(__file__))
_root = os.path.abspath(os.path.join(_here, '..'))
_templates = os.path.join(_root, 'templates')
sys.path.insert(0, _templates)

from market_analysis_template import generate_market_analysis_excel

project_root = _root

config_215A = {
    "ticker": "215A.T",
    "company_name": "株式会社タイミー",
    "current_price": 1459,            # 板で最新値に要更新（DCF Excelと同じ値推奨）
    "forward_revenue": 42861,         # base34,289×1.25=Base初年度。変則6ヶ月期を挟むためDCF Base初年度と揃える
    "price_3m_ago": 1459, "price_1m_ago": 1459, "price_high": 6540, "price_low": 1820,
    "margin_buy": 1000, "margin_sell": 200, "margin_sell_peak_6m": 400,
    "company_rev_growth": 0.25,
    "company_op_growth": 0.27,
    "narrative": {
        "date": "2026-06-17",
        "axis_1_label_status":        2,
        "axis_2a_catalyst_potential": 2,
        "axis_2b_catalyst_realization":2,
        "axis_3_diffusion_stage":     2,
        "axis_4_earnings_materiality":2,
        "axis_5_narrative_durability":2,
        "axis_1_note":  "スポットワーク国内最大手、ブランド認知は確立済み",
        "axis_2a_note": "物流・介護福祉への新業界進出、スキマワークスM&A",
        "axis_2b_note": "決算期変更で短期は見えにくいが新業界展開は着手済み",
        "axis_3_note":  "既に発見済み。機関・個人とも注目、セルサイドカバレッジあり",
        "axis_4_note":  "営業利益率約20%、黒字大（寄与極大）。粗利92%",
        "axis_5_note":  "人手不足は構造的追い風。但しメルカリハロ・リクルート等競争激化",
        "tam_oku_jpy":         5000,
        "expected_share_pct":  40,
        "segment_opm_pct":     20,
        "current_operating_profit_oku": 67,
        "fundamental_verdict": "OBSERVE",
    },
}

dcf_path_215A = os.path.join(project_root, "models", "215A_DCF_Model_20260617.xlsx")
out_path_215A = os.path.join(project_root, "reports", "215A_market_analysis_20260617.xlsx")

if not os.path.exists(dcf_path_215A):
    print(f"DCF Excel が見つかりません: {dcf_path_215A}")
    print("   215A_DCF_Model_20260617.xlsx を models/ に置いてください。")
    sys.exit(1)

os.makedirs(os.path.dirname(out_path_215A), exist_ok=True)
generate_market_analysis_excel(config_215A, out_path_215A, dcf_excel_path=dcf_path_215A)
print("\n完了：逆算DCF・逆算Compsシートを生成しました。")
print(f"  出力: {out_path_215A}")
