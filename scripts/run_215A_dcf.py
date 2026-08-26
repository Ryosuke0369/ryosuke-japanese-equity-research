"""
215A タイミー DCF本体 生成スクリプト
dcf_comps_template.generate_dcf_workbook を使用。
出力: 215A_DCF_Model_20260617.xlsx（→ models/ に移動して逆算分析の入力にする）

配置: scripts/run_215A_dcf.py
テンプレ: templates/dcf_comps_template.py
Comps:  scripts/comps_input.csv（UTF-8）

実行: python scripts/run_215A_dcf.py
"""
import os, sys

# --- import パス設定（325Aと同じ構造）---
_here = os.path.dirname(os.path.abspath(__file__))           # .../scripts
_root = os.path.abspath(os.path.join(_here, '..'))           # リポジトリのルート
_templates = os.path.join(_root, 'templates')                # .../templates
sys.path.insert(0, _templates)
sys.path.insert(0, _here)  # comps_fetcher 等が scripts/ にある場合

from dcf_comps_template import generate_dcf_workbook, get_live_market_data

# Comps CSV を読む（テンプレ standalone と同じ get_comps_data を使用）
try:
    from comps_fetcher import get_comps_data
    _comps = get_comps_data(os.path.join(_here, "comps_input.csv"))
except Exception as e:
    print(f"⚠️ comps_fetcher / CSV 読み込み失敗: {e}")
    print("   comps_input.csv が scripts/ にあるか、UTF-8か確認してください。")
    _comps = None

config = {
    # ── Company Info ──
    "company_name": "株式会社タイミー",
    "ticker": "215A.T",
    "exchange": "TSE Growth",
    "sector": "Services",
    "current_price": 3700,            # yfinance がライブ上書き（板の最新でなくてOK）
    "shares_outstanding": 100_000_000,  # ⚠️ yfinance がライブ上書き。下の get_live_market_data 参照
    "net_debt": -2300,                # JPY mn（負=ネットキャッシュ）。FY2025/10 BS概算

    # ── Historical Financials (JPY mn) ── 決算原典・短信より（過去3年＋当期）
    # FY2022/10, FY2023/10, FY2024/10(=base_year、売上34,289)
    "hist_years": ["FY2022 (Oct-22)", "FY2023 (Oct-23)", "FY2024 (Oct-24)", "FY2025 (Oct-25)"],
    "hist_revenue":          [16144, 26880, 34289, 34289],   # ※末尾=base。4年枠に合わせ直近を複写
    "hist_operating_income": [1957, 4247, 6747, 6747],
    "hist_net_income":       [1802, 2797, 5310, 5310],
    "hist_cogs":             [672, 1274, 1912, 1912],
    "hist_sga":              [13514, 21358, 25630, 25630],
    "hist_ocf":              [-749, 1183, 2674, 2674],        # CF計算書より（FY2024/10=2,674）
    "hist_capex":            [493, 138, 256, 256],
    "hist_cash":             [7996, 12238, 16540, 16540],
    "hist_debt":             [10643, 14213, 14225, 14225],

    # ── DCF Assumptions — Future Projections ──
    # cogs/sga は5年配列。原価極小・販管費高の事業構造を反映。利益率は sga で表現。
    "scenarios": {
        "Base": {
            "revenue_growth": [0.25, 0.22, 0.20, 0.18, 0.16],
            "cogs_pct": [0.056, 0.056, 0.056, 0.056, 0.056],
            "sga_pct":  [0.740, 0.734, 0.729, 0.724, 0.719],   # 営業益率 20.4%→22.5%
            "dso_days": [41, 41, 41, 41, 41],
            "dih_days": [0, 0, 0, 0, 0],                        # 在庫なし事業
            "dpo_days": [28, 28, 28, 28, 28],
        },
        "Upside": {
            "revenue_growth": [0.30, 0.27, 0.24, 0.21, 0.18],
            "cogs_pct": [0.055, 0.055, 0.055, 0.055, 0.055],
            "sga_pct":  [0.710, 0.700, 0.695, 0.690, 0.685],   # 営業益率 23.5%→26.0%
            "dso_days": [41, 41, 41, 41, 41],
            "dih_days": [0, 0, 0, 0, 0],
            "dpo_days": [28, 28, 28, 28, 28],
        },
        "Management": {
            "revenue_growth": [0.27, 0.24, 0.21, 0.19, 0.17],
            "cogs_pct": [0.055, 0.055, 0.055, 0.055, 0.055],
            "sga_pct":  [0.700, 0.695, 0.690, 0.685, 0.680],   # 営業益率 24.5%→26.5%
            "dso_days": [41, 41, 41, 41, 41],
            "dih_days": [0, 0, 0, 0, 0],
            "dpo_days": [28, 28, 28, 28, 28],
        },
        "Downside 1": {
            "revenue_growth": [0.18, 0.15, 0.13, 0.11, 0.10],
            "cogs_pct": [0.058, 0.058, 0.058, 0.058, 0.058],
            "sga_pct":  [0.760, 0.762, 0.764, 0.767, 0.772],   # 営業益率 18.2%→17.0%
            "dso_days": [45, 45, 45, 45, 45],
            "dih_days": [0, 0, 0, 0, 0],
            "dpo_days": [28, 28, 28, 28, 28],
        },
        "Downside 2": {
            "revenue_growth": [0.15, 0.12, 0.10, 0.09, 0.08],
            "cogs_pct": [0.060, 0.060, 0.060, 0.060, 0.060],
            "sga_pct":  [0.790, 0.795, 0.800, 0.805, 0.810],   # 営業益率 15.0%→13.0%
            "dso_days": [48, 48, 48, 48, 48],
            "dih_days": [0, 0, 0, 0, 0],
            "dpo_days": [28, 28, 28, 28, 28],
        },
    },
    "capex_pct": 0.01,        # アセットライト（売上比 約1%）
    "da_pct": 0.005,          # 減価償却 売上比 約0.5%（FY2024/10: 153/34289）
    "tax_rate": 0.30,
    "risk_free": 0.012,
    "beta": 1.5,
    "erp": 0.060,
    "size_premium": 0.030,    # 小型グロースだが黒字大なので 325A(0.05) より低め
    "cost_of_debt_at": 0.015,
    "de_ratio": 0.05,         # 手動（EDINET自動値は1/10になる罠）。実質無借金
    "terminal_growth": 0.015,
    "exit_multiple": 12.0,    # EV/EBITDA。高収益高成長をやや評価
    "projection_years": 5,
    "base_year_revenue": 34289,
    "base_year_cogs": 1912,

    # ── NWC Base Year Actuals (JPY mn) ── 画像BSより
    "base_year_ar":   3859,   # 売掛金
    "base_year_inv":  0,      # 在庫なし
    "base_year_ap":   2656,   # 未払金（買掛相当）

    # ── Comparable Companies ──
    "comps": _comps,

    # ── Kudan(逆算) Comps Data ──
    "core_ebitda": 6747 + 153,        # 営業利益+減価償却（FY2024/10）
    "core_net_income": 5310,

    # ── Investment Thesis & Risks ──
    "investment_thesis": [
        "1. スポットワーク市場の国内最大手、先行者ネットワーク効果",
        "2. 粗利92%・営業益率約20%の高収益プラットフォーム、利益レバレッジ大",
        "3. 物流・介護福祉への新業界展開、スキマワークスM&Aで成長余地拡大",
    ],
    "key_risks": [
        "1. メルカリハロ・LINEスキマニ・リクルート参入による競争激化",
        "2. FY2026は『仕込みの時期』で投資先行、利益の伸びが一時鈍化",
        "3. 決算期変更（12ヶ月→6ヶ月変則）で短期の業績比較が見えにくい",
    ],

    # ── V3 Settings ──
    "primary_multiple": "EV/EBITDA",   # 黒字なので EV/EBITDA
}

# Base配列をフラットに復元（感応度分析の後方互換。テンプレ standalone と同じ処理）
_base = config["scenarios"]["Base"]
config["revenue_growth"] = _base["revenue_growth"]
config["cogs_pct"]       = _base["cogs_pct"]
config["sga_pct"]        = _base["sga_pct"]

# ライブ市場データを yfinance で取得・上書き
# ※テンプレのバージョンにより戻り値が (price, shares) の2つか
#   (price, shares, beta) の3つか異なるため、個数非依存で受ける
_live = get_live_market_data(
    config.get("ticker", ""),
    config.get("current_price", 0),
    config.get("shares_outstanding", 0),
)
config["current_price"] = _live[0]
config["shares_outstanding"] = _live[1]
if len(_live) >= 3:
    config["beta"] = _live[2]   # ライブbetaで上書き（範囲外なら内部でフォールバック済み）

# 出力先（models/ に置くと逆算分析がそのまま読める）
_out = os.path.join(_root, "models", "215A_DCF_Model_20260617.xlsx")
os.makedirs(os.path.dirname(_out), exist_ok=True)

generate_dcf_workbook(config)  # テンプレ既定の出力名で保存される
print("\n完了：DCF本体を生成しました。")
print("  ⚠️ テンプレ既定の出力名（215AT_Equity_Research_V3.xlsx 等）で")
print("     カレントに保存されます。逆算分析の前に models/215A_DCF_Model_20260617.xlsx に")
print("     リネーム＆移動してください。")
