"""
325A TENTIAL 逆算DCF・逆算Comps 実行スクリプト
market_analysis_template.py を使用。
前提：325A_DCF_Model_20260616.xlsx が models/ にあること。

配置: scripts/run_325A_market.py
テンプレ: templates/market_analysis_template.py
"""
import os, sys

# --- import パス設定 ---
# このスクリプトは scripts/ にあり、テンプレは templates/ にある。
# templates/ を import 探索先に追加する。
_here = os.path.dirname(os.path.abspath(__file__))           # .../scripts
_root = os.path.abspath(os.path.join(_here, '..'))           # リポジトリのルート
_templates = os.path.join(_root, 'templates')                # .../templates
sys.path.insert(0, _templates)

from market_analysis_template import generate_market_analysis_excel

project_root = _root

config_325A = {
    "ticker": "325A.T",
    "company_name": "株式会社TENTIAL",
    "current_price": 3675,            # ⚠️ 板で最新値に要更新（DCF Excelと同じ値推奨）
    "forward_revenue": 30200,         # 会社予想売上 FY2026/8（フォワード倍率の分母）
    "price_3m_ago": 3675, "price_1m_ago": 3675, "price_high": 3675, "price_low": 3675,
    "margin_buy": 1000, "margin_sell": 200, "margin_sell_peak_6m": 400,
    "company_rev_growth": 0.58,
    "company_op_growth": 0.65,
    "narrative": {
        "date": "2026-06-16",
        "axis_1_label_status":        1,
        "axis_2a_catalyst_potential": 1,
        "axis_2b_catalyst_realization":1,
        "axis_3_diffusion_stage":     1,
        "axis_4_earnings_materiality":2,
        "axis_5_narrative_durability":1,
        "axis_1_note":  "リカバリーウェアBAKUNEブランド、個人に浸透途上",
        "axis_2a_note": "寝具・サンダルへの横展開、店舗拡大（5都市で確認）",
        "axis_2b_note": "実需は堅調だが単発の大型カタリストは限定的",
        "axis_3_note":  "個人投資家中心、セルサイドカバレッジ薄い",
        "axis_4_note":  "営業利益率10-13%、黒字が主力（寄与大）",
        "axis_5_note":  "健康・睡眠トレンドは構造的、但しMTG等競合・単一ブランド依存",
        "tam_oku_jpy":         2000,
        "expected_share_pct":  15,
        "segment_opm_pct":     12,
        "current_operating_profit_oku": 20,
        "fundamental_verdict": "HOLD",
    },
}

dcf_path_325A = os.path.join(project_root, "models", "325A_DCF_Model_20260616.xlsx")
out_path_325A = os.path.join(project_root, "reports", "325A_market_analysis_20260616.xlsx")

# DCF Excel が models/ に無い場合のヒント
if not os.path.exists(dcf_path_325A):
    print(f"⚠️ DCF Excel が見つかりません: {dcf_path_325A}")
    print("   325A_DCF_Model_20260616.xlsx を models/ に置いてください。")
    print("   （前回 generate_dcf.py で生成したファイル。reports/ 等にあるなら移動）")
    sys.exit(1)

generate_market_analysis_excel(config_325A, out_path_325A, dcf_excel_path=dcf_path_325A)
print("\n完了：逆算DCF・逆算Compsシートを生成しました。")
print(f"  出力: {out_path_325A}")
