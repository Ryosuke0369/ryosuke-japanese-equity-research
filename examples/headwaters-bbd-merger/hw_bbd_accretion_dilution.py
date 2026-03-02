# -*- coding: utf-8 -*-
"""
ヘッドウォータース（4011.T）によるBBDイニシアティブ（5259.T）吸収合併
M&A Accretion/Dilution Analysis（合併による増益/希薄化分析）
分析日: 2026年2月19日
"""

import math

OUTPUT_FILE = "hw_bbd_accretion_dilution_report.txt"

# ============================================================
# ユーティリティ
# ============================================================

def fmt(v, unit="百万円", decimals=1):
    """数値フォーマット（None対応）"""
    if v is None:
        return "N/A"
    if isinstance(v, float) and math.isnan(v):
        return "N/A"
    return f"{v:,.{decimals}f} {unit}"

def fmt_pct(v, decimals=1):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "N/A"
    return f"{v:+.{decimals}f}%"

def divz(a, b):
    """ゼロ除算回避"""
    if b == 0 or b is None:
        return None
    return a / b

# ============================================================
# ① マスターデータ
# ============================================================

# --- 取引概要 ---
MERGER_RATIO       = 0.50          # BBD 1株 → HW 0.50株
BBD_SHARES_TOTAL   = 6_120_221     # BBD発行済株式数
HW_HOLDS_BBD       = 1_599_100     # HWが既保有のBBD株
BBD_TREASURY       = 296           # BBD自己株式
BBD_ALLOC_SHARES   = 4_520_825     # 合併割当対象株式数（= 発行済 - HW保有 - 自己株）
HW_NEW_SHARES      = 2_260_412     # HW新規発行株式数

# --- ヘッドウォータース（HW）財務データ ---
hw = {
    # FY2024 実績
    "rev_24":     2_905,   # 売上高（百万円）
    "ebit_24":      307,   # 営業利益
    "da_24":         20,   # 減価償却費
    "ebitda_24":    327,   # EBITDA
    "ni_24":        272,   # 純利益
    "eps_24":      72.0,   # EPS（円）
    # FY2025 実績
    "rev_25":     3_900,
    "ebit_25":      229,
    "da_25":         27,
    "ebitda_25":    256,
    "ni_25":         58,
    "eps_25":      15.3,
    # FY2026 会社予想
    "rev_26":     5_764,
    "ebit_26":      436,
    "ni_26":        223,
    "eps_26":      60.1,
    # 株式・バランスシート
    "shares":     3_844_144,   # 発行済株式数
    "mktcap":    11_090,       # 時価総額（百万円、2,885円ベース）
    "stock_price":  2_885,     # 株価（円）
    "so":           399_200,   # ストックオプション残高
    "equity_24":  1_273,       # 純資産（百万円）
    "assets_24":  1_800,       # 総資産
    "cash_24":      625,       # 現金
}

# --- BBDイニシアティブ（BBD）財務データ ---
bbd = {
    # FY2024 実績（黒字期）
    "rev_24":     4_131,
    "ebit_24":      273,
    "da_24":        312,
    "ebitda_24":    585,
    "ni_24":        156,
    "eps_24":      31.2,
    # FY2025 実績（赤字期）
    "rev_25":     4_399,
    "ebit_25":     -351,
    "da_25":        342,
    "ebitda_25":     -9,
    "ni_25":      1_750,    # 特別利益による（営業ベースは大幅赤字）
    "eps_25_oper": -63.6,   # 営業ベースEPS
    # FY2026 会社予想
    "rev_26":     4_803,
    # 株式・バランスシート
    "shares":     6_120_221,
    "mktcap":     7_310,       # 時価総額（百万円、1,194円ベース）
    "stock_price":  1_194,     # 株価（円）
    "equity_24":  1_758,       # 純資産
    "assets_24":  3_989,       # 総資産
    "arr":        1_640,       # サブスクリプションARR（百万円）
    # 発表前日株価・合併換算価値
    "pre_announce_price": 1_591,  # 発表前日BBD株価（円）
    "implied_price": hw["stock_price"] * MERGER_RATIO,  # = 2885 × 0.50 = 1442.5円
}

# 合併後HW株式数（ベース）
HW_SHARES_POST_MERGER  = hw["shares"] + HW_NEW_SHARES   # 6,104,556株
HW_SHARES_FULLY_DILUTED = HW_SHARES_POST_MERGER + hw["so"]  # 6,503,756株

# 取引対価（時価）
DEAL_VALUE_MKTCAP = HW_NEW_SHARES * hw["stock_price"] / 1e6  # 百万円
# BBD株式全体のEquity Value（HW保有分含む）
BBD_FULL_EQ_VALUE = BBD_ALLOC_SHARES * hw["stock_price"] * MERGER_RATIO / 1e6

# のれん推定
GOODWILL_EST = DEAL_VALUE_MKTCAP - bbd["equity_24"]  # 取引対価 - BBD純資産

# 税率前提
TAX_RATE = 0.30  # 実効税率30%

# ============================================================
# レポート書き出し用バッファ
# ============================================================

lines = []

def h1(text):
    lines.append("")
    lines.append("=" * 72)
    lines.append(f"  {text}")
    lines.append("=" * 72)

def h2(text):
    lines.append("")
    lines.append("-" * 60)
    lines.append(f"  {text}")
    lines.append("-" * 60)

def h3(text):
    lines.append(f"\n【{text}】")

def row(label, value, indent=2):
    lines.append(f"{'  ' * indent}{label}: {value}")

def note(text, indent=2):
    lines.append(f"{'  ' * indent}※ {text}")

def blank():
    lines.append("")

def calc_row(label, formula, value):
    """計算式と結果を出力"""
    lines.append(f"    {label}")
    lines.append(f"      計算式: {formula}")
    lines.append(f"      結  果: {value}")

# ============================================================
# ヘッダー
# ============================================================

h1("ヘッドウォータース（4011.T）× BBDイニシアティブ（5259.T）")
h1("M&A Accretion / Dilution Analysis（合併による増益/希薄化分析）")
lines.append("  分析基準日 : 2026年2月19日")
lines.append("  合併効力日 : 2026年5月1日")
lines.append("  通貨単位  : 百万円（特記なき場合）")
lines.append("  作成者    : Pythonスクリプト自動生成")

# ============================================================
# Ⅰ. 取引概要
# ============================================================

h1("Ⅰ. 取引概要（Transaction Overview）")

h2("1-1. 基本条件")
row("取引形態", "吸収合併（株式対価のみ、現金なし）")
row("存続会社", "ヘッドウォータース株式会社（4011.T）")
row("消滅会社", "BBDイニシアティブ株式会社（5259.T）")
row("合併比率", f"BBD 1株 → HW {MERGER_RATIO}株")
row("合併効力発生日", "2026年5月1日")

blank()
h2("1-2. 株式数の確認")

# 計算検証
alloc_check = BBD_SHARES_TOTAL - HW_HOLDS_BBD - BBD_TREASURY
calc_row(
    "合併割当対象株式数 = BBD発行済 - HW保有分 - BBD自己株式",
    f"{BBD_SHARES_TOTAL:,} - {HW_HOLDS_BBD:,} - {BBD_TREASURY:,}",
    f"{alloc_check:,}株  ({'検証OK' if alloc_check == BBD_ALLOC_SHARES else '注意: 不一致'})"
)

new_hw_check = round(BBD_ALLOC_SHARES * MERGER_RATIO)
calc_row(
    "HW新規発行株式数 = 割当対象株式数 × 合併比率",
    f"{BBD_ALLOC_SHARES:,} × {MERGER_RATIO}",
    f"{new_hw_check:,}株  ({'検証OK' if abs(new_hw_check - HW_NEW_SHARES) <= 1 else f'参考値({HW_NEW_SHARES:,}株と端数差あり)'})"
)

post_shares = hw["shares"] + HW_NEW_SHARES
calc_row(
    "合併後HW発行済株式数 = HW既存 + HW新規発行",
    f"{hw['shares']:,} + {HW_NEW_SHARES:,}",
    f"{post_shares:,}株"
)

fully_diluted = post_shares + hw["so"]
calc_row(
    "合併後完全希薄化株式数 = 発行済 + SO残高",
    f"{post_shares:,} + {hw['so']:,}",
    f"{fully_diluted:,}株"
)

blank()
h2("1-3. 取引価値（Deal Value）")

deal_val = HW_NEW_SHARES * hw["stock_price"] / 1e6
calc_row(
    "株式対価（HW株式換算）= 新規発行株数 × HW株価",
    f"{HW_NEW_SHARES:,}株 × {hw['stock_price']:,}円",
    fmt(deal_val)
)

# BBD 1株当たり合併対価（HW株価換算）
per_bbd_value = hw["stock_price"] * MERGER_RATIO
row("BBD 1株当たり合併対価（円換算）",
    f"{hw['stock_price']:,}円 × {MERGER_RATIO} = {per_bbd_value:,.1f}円")

# プレミアム vs 発表前日株価
premium_vs_pre = (per_bbd_value / bbd["pre_announce_price"] - 1) * 100
row("対発表前日株価プレミアム",
    f"({per_bbd_value:,.1f}円 / {bbd['pre_announce_price']:,}円 - 1) = {premium_vs_pre:+.1f}%")

# プレミアム vs 合併発表日株価（1,194円）
premium_vs_mkt = (per_bbd_value / bbd["stock_price"] - 1) * 100
row("対直近市場株価プレミアム",
    f"({per_bbd_value:,.1f}円 / {bbd['stock_price']:,}円 - 1) = {premium_vs_mkt:+.1f}%")

note("合併対価 = 2,885円 × 0.50 = 1,442.5円/BBD株")
note(f"発表前日BBD株価 1,591円に対しディスカウント: 合併換算価値は市場株価を下回る")

# ============================================================
# Ⅱ. Pre-Deal Analysis（合併前スタンドアロン分析）
# ============================================================

h1("Ⅱ. Pre-Deal Analysis（合併前スタンドアロン分析）")

h2("2-1. スタンドアロンEPS比較")

blank()
lines.append("    ■ ヘッドウォータース（HW）")
lines.append(f"    {'指標':<20} {'FY2024実績':>15} {'FY2025実績':>15} {'FY2026予想':>15}")
lines.append("    " + "-" * 68)
lines.append(f"    {'売上高（百万円）':<20} {hw['rev_24']:>15,} {hw['rev_25']:>15,} {hw['rev_26']:>15,}")
lines.append(f"    {'営業利益（百万円）':<20} {hw['ebit_24']:>15,} {hw['ebit_25']:>15,} {hw['ebit_26']:>15,}")
lines.append(f"    {'純利益（百万円）':<20} {hw['ni_24']:>15,} {hw['ni_25']:>15,} {hw['ni_26']:>15,}")
lines.append(f"    {'EPS（円）':<20} {hw['eps_24']:>15.1f} {hw['eps_25']:>15.1f} {hw['eps_26']:>15.1f}")
lines.append(f"    {'発行済株式数':<20} {hw['shares']:>15,} {hw['shares']:>15,} {hw['shares']:>15,}")

blank()
lines.append("    ■ BBDイニシアティブ（BBD）")
lines.append(f"    {'指標':<20} {'FY2024実績':>15} {'FY2025実績':>15} {'FY2026予想':>15}")
lines.append("    " + "-" * 68)
lines.append(f"    {'売上高（百万円）':<20} {bbd['rev_24']:>15,} {bbd['rev_25']:>15,} {bbd['rev_26']:>15,}")
lines.append(f"    {'営業利益（百万円）':<20} {bbd['ebit_24']:>15,} {bbd['ebit_25']:>+15,} {'-':>15}")
lines.append(f"    {'純利益（百万円）':<20} {bbd['ni_24']:>15,} {bbd['ni_25']:>15,} {'-':>15}")
lines.append(f"    {'EPS（円）':<20} {bbd['eps_24']:>15.1f} {'*'+str(bbd['eps_25_oper']):>15} {'-':>15}")
lines.append(f"    {'発行済株式数':<20} {bbd['shares']:>15,} {bbd['shares']:>15,} {bbd['shares']:>15,}")
note("FY2025純利益1,750百万円は特別利益（負ののれん等）を含む。営業ベースEPS = -63.6円")
note("BBD決算期は9月末。FY2024=2024年9月期、FY2025=2025年9月期")

blank()
h2("2-2. 合併比率の妥当性検証")

h3("① 市場株価法（Market Price Method）")
hw_price  = hw["stock_price"]
bbd_price = bbd["stock_price"]
implied_ratio_mkt = bbd_price / hw_price
calc_row(
    "市場株価ベース合併比率 = BBD株価 / HW株価",
    f"{bbd_price:,}円 / {hw_price:,}円",
    f"{implied_ratio_mkt:.4f}"
)
row("  採用合併比率", f"{MERGER_RATIO:.4f}")
deviation_mkt = (MERGER_RATIO / implied_ratio_mkt - 1) * 100
row("  採用比率の乖離", f"{deviation_mkt:+.1f}%（市場比率に対し）")
note(f"合併比率{MERGER_RATIO}は市場株価ベース比率{implied_ratio_mkt:.4f}より低く、BBD株主にとってやや不利")

blank()
h3("② EV/EBITDA法（FY2024黒字期ベース）")

# HW EV
hw_ev_24 = hw["mktcap"] + 0 - hw["cash_24"]  # 有利子負債は簡略化でゼロ想定
# BBD EV（取引対価ベース）
bbd_eq_deal = deal_val  # 合併割当部分のみの対価
bbd_ev_deal = bbd_eq_deal + (bbd["assets_24"] - bbd["equity_24"]) - 0  # 負債含む推定
# 市場株価ベースBBD EV
bbd_ev_mkt  = bbd["mktcap"] + (bbd["assets_24"] - bbd["equity_24"])

calc_row(
    "HW EV（市場株価ベース）= 時価総額 - 現金",
    f"{hw['mktcap']:,} - {hw['cash_24']:,}",
    fmt(hw_ev_24)
)
calc_row(
    "HW EV/EBITDA（FY2024）= HW EV / HW EBITDA",
    f"{hw_ev_24:,} / {hw['ebitda_24']:,}",
    f"{hw_ev_24/hw['ebitda_24']:.1f}x"
)
blank()
calc_row(
    "BBD EV（市場株価ベース）= 時価総額 + 純負債",
    f"{bbd['mktcap']:,} + ({bbd['assets_24']:,} - {bbd['equity_24']:,})",
    fmt(bbd_ev_mkt)
)
calc_row(
    "BBD EV/EBITDA（FY2024）= BBD EV / BBD EBITDA",
    f"{bbd_ev_mkt:,} / {bbd['ebitda_24']:,}",
    f"{bbd_ev_mkt/bbd['ebitda_24']:.1f}x"
)
blank()
# 取引ベースBBD EV
bbd_net_debt = bbd["assets_24"] - bbd["equity_24"]  # 純負債（簡易推定）
bbd_ev_trans = deal_val + bbd_net_debt
calc_row(
    "BBD EV（取引対価ベース）= 合併対価 + BBD純負債推定",
    f"{deal_val:.1f} + {bbd_net_debt:,}",
    fmt(bbd_ev_trans)
)
calc_row(
    "取引ベースBBD EV/EBITDA（FY2024）",
    f"{bbd_ev_trans:.1f} / {bbd['ebitda_24']:,}",
    f"{bbd_ev_trans/bbd['ebitda_24']:.1f}x"
)

blank()
h3("③ PER法（Price/Earnings Ratio）")
hw_per_24  = hw["stock_price"] / hw["eps_24"]
bbd_per_24 = bbd["stock_price"] / bbd["eps_24"]
calc_row(
    "HW PER（FY2024実績）= HW株価 / HW EPS",
    f"{hw['stock_price']:,}円 / {hw['eps_24']:.1f}円",
    f"{hw_per_24:.1f}x"
)
calc_row(
    "BBD PER（FY2024実績）= BBD株価 / BBD EPS",
    f"{bbd['stock_price']:,}円 / {bbd['eps_24']:.1f}円",
    f"{bbd_per_24:.1f}x"
)
implied_bbd_by_hw_per = hw_per_24 * bbd["eps_24"]
row("  HW PER適用時のBBD理論株価", f"{implied_bbd_by_hw_per:.0f}円")
implied_ratio_per = implied_bbd_by_hw_per / hw["stock_price"]
row("  PER法による理論合併比率", f"{implied_ratio_per:.4f}")

# ============================================================
# Ⅲ. Pro Forma EPS Analysis
# ============================================================

h1("Ⅲ. Pro Forma EPS Analysis（合併後EPS分析）")

h2("3-1. 前提条件の確認")

row("HW既存発行済株式数", f"{hw['shares']:,}株")
row("HW新規発行株式数（合併対価）", f"{HW_NEW_SHARES:,}株")
row("合併後発行済株式数（ベース）", f"{HW_SHARES_POST_MERGER:,}株")
row("合併後完全希薄化株式数（+SO）", f"{HW_SHARES_FULLY_DILUTED:,}株")
row("HW FY2026会社予想純利益", fmt(hw["ni_26"]))
row("HW FY2026会社予想EPS（スタンドアロン）", f"{hw['eps_26']:.1f}円")
note("以降の分析では分析基準年度をFY2026（2026年12月期）とする")
note("BBDの決算期（9月末）はHWの12月期と3か月ずれあり。簡略化のためカレンダー年度調整は行わない")

blank()
h2("3-2. シナリオ別Pro Forma EPS計算")

# --- ベースとなるHW FY2026純利益 ---
hw_ni_base = hw["ni_26"]  # 223百万円

# --- シナリオ定義 ---
scenarios = {
    "A（ベースケース）": {
        "bbd_contrib": 0,
        "synergy":     0,
        "desc": "BBD損益ゼロ想定（赤字継続だが合併後立て直し中）",
    },
    "B（ダウンサイド）": {
        "bbd_contrib": bbd["ebit_25"] * (1 - TAX_RATE),  # 営業損失が継続、税後換算
        "synergy":     0,
        "desc": "BBD FY2025営業損失（△351百万円）がそのまま継続",
    },
    "C（シナジーケース）": {
        "bbd_contrib": 0,
        "synergy":     hw_ni_base * 0.20,  # 合併後純利益20%増
        "desc": "統合シナジーにより合併後利益20%増加",
    },
}

# スタンドアロンEPS（FY2026予想）
hw_eps_standalone = hw["eps_26"]  # 60.1円（会社予想ベース、旧株式数）

# 発行済ベースと完全希薄化ベースで両方計算
for scenario_name, params in scenarios.items():
    bbd_c   = params["bbd_contrib"]
    syn     = params["synergy"]
    desc    = params["desc"]

    # Pro Forma純利益
    pf_ni = hw_ni_base + bbd_c + syn

    # Pro Forma EPS（発行済ベース）
    pf_eps_basic = pf_ni * 1e6 / HW_SHARES_POST_MERGER  # 円

    # Pro Forma EPS（完全希薄化ベース）
    pf_eps_diluted = pf_ni * 1e6 / HW_SHARES_FULLY_DILUTED  # 円

    # Accretion / Dilution（スタンドアロンEPS比）
    # スタンドアロンEPS = 223百万円 ÷ 3,844,144株
    hw_eps_sa_calc = hw_ni_base * 1e6 / hw["shares"]
    accretion_basic   = (pf_eps_basic   / hw_eps_sa_calc - 1) * 100
    accretion_diluted = (pf_eps_diluted / hw_eps_sa_calc - 1) * 100

    params["pf_ni"]           = pf_ni
    params["pf_eps_basic"]    = pf_eps_basic
    params["pf_eps_diluted"]  = pf_eps_diluted
    params["accretion_basic"] = accretion_basic
    params["accretion_diluted"] = accretion_diluted
    params["hw_eps_sa"]       = hw_eps_sa_calc

    h3(f"シナリオ {scenario_name}")
    lines.append(f"    説明: {desc}")
    blank()
    calc_row(
        "Pro Forma純利益 = HW純利益 + BBD貢献 + シナジー",
        f"{hw_ni_base} + ({bbd_c:.1f}) + {syn:.1f}",
        fmt(pf_ni)
    )
    calc_row(
        "Pro Forma EPS（発行済ベース）= PF純利益 / 合併後発行済株式数",
        f"{pf_ni:.1f}百万円 × 1,000,000 / {HW_SHARES_POST_MERGER:,}株",
        f"{pf_eps_basic:.2f}円"
    )
    calc_row(
        "Pro Forma EPS（完全希薄化）= PF純利益 / 完全希薄化株式数",
        f"{pf_ni:.1f}百万円 × 1,000,000 / {HW_SHARES_FULLY_DILUTED:,}株",
        f"{pf_eps_diluted:.2f}円"
    )
    calc_row(
        "HW スタンドアロン EPS（FY2026）= HW純利益 / HW既存株式数",
        f"{hw_ni_base}百万円 × 1,000,000 / {hw['shares']:,}株",
        f"{hw_eps_sa_calc:.2f}円"
    )
    lines.append(f"    → Accretion / Dilution（発行済ベース）   : {accretion_basic:+.1f}%  "
                 f"({'Accretive（増益）' if accretion_basic >= 0 else 'Dilutive（希薄化）'})")
    lines.append(f"    → Accretion / Dilution（完全希薄化ベース）: {accretion_diluted:+.1f}%  "
                 f"({'Accretive（増益）' if accretion_diluted >= 0 else 'Dilutive（希薄化）'})")

blank()
h2("3-3. シナリオ比較サマリー")

hw_eps_sa_summary = scenarios["A（ベースケース）"]["hw_eps_sa"]
lines.append(f"    {'シナリオ':<22} {'PF純利益':>10} {'PF EPS(発)':>12} {'PF EPS(希)':>12} {'Accretion(発)':>14} {'Accretion(希)':>14}")
lines.append("    " + "-" * 88)
for name, p in scenarios.items():
    lines.append(
        f"    {name:<22} {p['pf_ni']:>8.1f}百万 {p['pf_eps_basic']:>10.2f}円 "
        f"{p['pf_eps_diluted']:>10.2f}円 {p['accretion_basic']:>+13.1f}% {p['accretion_diluted']:>+13.1f}%"
    )
lines.append(f"    {'[参考] HW スタンドアロン':<22} {hw_ni_base:>8.1f}百万 {hw_eps_sa_summary:>10.2f}円 {'-':>12} {'-':>14} {'-':>14}")

# ============================================================
# Ⅳ. Break-Even Synergy Analysis
# ============================================================

h1("Ⅳ. Break-Even Synergy Analysis（損益分岐シナジー分析）")

h2("4-1. EPS希薄化ゼロに必要な最低シナジー（発行済ベース）")

# スタンドアロンEPS = hw_eps_sa_summary
# Post-merger shares = HW_SHARES_POST_MERGER
# Break-even condition: (hw_ni_base + bbd_contrib + synergy) / POST_SHARES = hw_eps_sa_summary / 1,000,000
# → synergy = hw_eps_sa_summary * POST_SHARES / 1e6 - hw_ni_base - bbd_contrib

# シナリオAベース（BBD貢献ゼロ）
bbd_c_a = scenarios["A（ベースケース）"]["bbd_contrib"]
be_synergy_a_basic = (hw_eps_sa_summary * HW_SHARES_POST_MERGER / 1e6) - hw_ni_base - bbd_c_a
calc_row(
    "損益分岐シナジー（シナリオA、発行済ベース）",
    f"= {hw_eps_sa_summary:.4f}円 × {HW_SHARES_POST_MERGER:,}株 / 1,000,000 - {hw_ni_base} - {bbd_c_a:.1f}",
    fmt(be_synergy_a_basic)
)
note(f"合併後EPS = スタンドアロンEPSを維持するために税後ベースで {be_synergy_a_basic:.1f}百万円のシナジーが必要")

be_synergy_a_pretax = be_synergy_a_basic / (1 - TAX_RATE)
calc_row(
    "税引前ベース必要シナジー（実効税率30%前提）",
    f"{be_synergy_a_basic:.1f} / (1 - {TAX_RATE})",
    fmt(be_synergy_a_pretax)
)

blank()
h2("4-2. EPS希薄化ゼロに必要な最低シナジー（完全希薄化ベース）")

be_synergy_a_diluted = (hw_eps_sa_summary * HW_SHARES_FULLY_DILUTED / 1e6) - hw_ni_base - bbd_c_a
calc_row(
    "損益分岐シナジー（シナリオA、完全希薄化ベース）",
    f"= {hw_eps_sa_summary:.4f}円 × {HW_SHARES_FULLY_DILUTED:,}株 / 1,000,000 - {hw_ni_base}",
    fmt(be_synergy_a_diluted)
)
be_synergy_a_diluted_pretax = be_synergy_a_diluted / (1 - TAX_RATE)
calc_row(
    "税引前ベース必要シナジー（完全希薄化）",
    f"{be_synergy_a_diluted:.1f} / (1 - {TAX_RATE})",
    fmt(be_synergy_a_diluted_pretax)
)

blank()
h2("4-3. シナリオB（BBD損失継続）の損益分岐シナジー")

bbd_c_b = scenarios["B（ダウンサイド）"]["bbd_contrib"]
be_synergy_b_basic = (hw_eps_sa_summary * HW_SHARES_POST_MERGER / 1e6) - hw_ni_base - bbd_c_b
calc_row(
    "損益分岐シナジー（シナリオB、発行済ベース）",
    f"= {hw_eps_sa_summary:.4f}円 × {HW_SHARES_POST_MERGER:,}株 / 1,000,000 - {hw_ni_base} - ({bbd_c_b:.1f})",
    fmt(be_synergy_b_basic)
)
be_synergy_b_pretax = be_synergy_b_basic / (1 - TAX_RATE)
calc_row(
    "税引前ベース必要シナジー（シナリオB）",
    f"{be_synergy_b_basic:.1f} / (1 - {TAX_RATE})",
    fmt(be_synergy_b_pretax)
)

note("シナリオBではBBD損失を引き継ぐため、必要シナジーはシナリオAより大幅に拡大")

# ============================================================
# Ⅴ. Transaction Multiple Analysis
# ============================================================

h1("Ⅴ. Transaction Multiple Analysis（取引マルチプル分析）")

h2("5-1. 取引対価の計算")

row("合併割当対象株式数", f"{BBD_ALLOC_SHARES:,}株")
row("HW株価（基準日）", f"{hw['stock_price']:,}円")
row("合併比率", f"{MERGER_RATIO}")

deal_val_yen = BBD_ALLOC_SHARES * hw["stock_price"] * MERGER_RATIO  # 円
deal_val_mn  = deal_val_yen / 1e6  # 百万円

calc_row(
    "取引対価（Equity Value対価）= 割当株数 × HW株価 × 合併比率",
    f"{BBD_ALLOC_SHARES:,} × {hw['stock_price']:,}円 × {MERGER_RATIO}",
    f"{deal_val_mn:,.1f}百万円 / {deal_val_yen/1e8:,.2f}億円"
)

# BBD全株主ベース（HW保有分も含めた全体評価額）
bbd_full_eq = bbd["shares"] * hw["stock_price"] * MERGER_RATIO / 1e6
calc_row(
    "BBD Equity Value全体（HW保有含む）= BBD全発行済株数 × HW株価 × 合併比率",
    f"{bbd['shares']:,} × {hw['stock_price']:,} × {MERGER_RATIO} / 1,000,000",
    f"{bbd_full_eq:,.1f}百万円"
)

# Enterprise Value（取引ベース）
# 純負債 = 総資産 - 純資産（簡易推定）
bbd_net_debt_detail = bbd["assets_24"] - bbd["equity_24"]
bbd_ev = bbd_full_eq + bbd_net_debt_detail
calc_row(
    "BBD Enterprise Value（取引ベース）= Equity Value + 純負債（簡易推定）",
    f"{bbd_full_eq:.1f} + ({bbd['assets_24']:,} - {bbd['equity_24']:,})",
    fmt(bbd_ev)
)

blank()
h2("5-2. 暗示的マルチプル（FY2024黒字期ベース）")

ev_rev_24   = bbd_ev / bbd["rev_24"]
ev_ebitda_24 = bbd_ev / bbd["ebitda_24"]
per_deal_24  = (bbd_full_eq / bbd["shares"] * 1e6) / bbd["eps_24"]  # 取引対価÷EPS
# bbd_full_eq は百万円なので、1株あたりは百万円×1e6÷株数
price_per_bbd = bbd_full_eq * 1e6 / bbd["shares"]  # 円
per_deal_24   = price_per_bbd / bbd["eps_24"]

calc_row(
    "EV/Revenue（FY2024） = BBD EV / BBD売上高",
    f"{bbd_ev:.1f} / {bbd['rev_24']:,}",
    f"{ev_rev_24:.2f}x"
)
calc_row(
    "EV/EBITDA（FY2024）  = BBD EV / BBD EBITDA",
    f"{bbd_ev:.1f} / {bbd['ebitda_24']:,}",
    f"{ev_ebitda_24:.1f}x"
)
calc_row(
    "PER（FY2024）        = 合併換算価格 / BBD EPS",
    f"{price_per_bbd:.1f}円 / {bbd['eps_24']:.1f}円",
    f"{per_deal_24:.1f}x"
)

blank()
h2("5-3. 暗示的マルチプル（FY2025赤字期ベース）")

# FY2025 EBITDA = -9 百万円（マルチプル計算が無意味になるためコメント）
if bbd["ebitda_25"] < 0:
    lines.append("    ※ FY2025 EBITDAは△9百万円（マイナス）のため、EV/EBITDAは意味をなさない（N/M）")
ev_rev_25 = bbd_ev / bbd["rev_25"]
calc_row(
    "EV/Revenue（FY2025） = BBD EV / BBD売上高",
    f"{bbd_ev:.1f} / {bbd['rev_25']:,}",
    f"{ev_rev_25:.2f}x"
)
lines.append("    EV/EBITDA（FY2025）  : N/M（EBITDAがマイナスのため算出不可）")
lines.append("    PER（FY2025）        : N/M（純利益が特別利益に歪められているため参考外）")

blank()
h2("5-4. マルチプル評価コメント")

lines.append(f"""
    【分析コメント】
    ・取引ベースEV/Revenue（FY2024）: {ev_rev_24:.2f}x
      → IT・SaaS系M&Aの一般的な水準（1x〜5x）と比較すると{"妥当な水準" if 1 <= ev_rev_24 <= 5 else "参考水準外"}
    ・取引ベースEV/EBITDA（FY2024）: {ev_ebitda_24:.1f}x
      → ソフトウェア・IT企業M&Aの典型的レンジ（8x〜15x）と比較すると{"比較的低い水準" if ev_ebitda_24 < 8 else "標準的水準"}
    ・BBDはFY2025に大幅な業績悪化（営業損失△351百万円）を計上しており、
      FY2024黒字期ベースのマルチプルは将来収益力を過大評価している可能性がある
    ・サブスクリプションARR 1,640百万円を有することから、
      EV/ARR = {bbd_ev:.1f} / 1,640 = {bbd_ev/1640:.2f}x であり、
      SaaS企業としてのARRマルチプルはやや保守的な水準
    ・合併比率が発表前日株価に対しディスカウント（{premium_vs_pre:+.1f}%）であることは
      HW株主にとって有利、BBD株主には不利な条件となっている
""")

# ============================================================
# Ⅵ. Pro Forma Balance Sheet（簡易版）
# ============================================================

h1("Ⅵ. Pro Forma Balance Sheet（簡易合算貸借対照表）")

h2("6-1. 合算前の両社バランスシート（FY2024ベース）")

lines.append(f"    {'項目':<25} {'HW':>12} {'BBD':>12} {'単純合算':>12}")
lines.append("    " + "-" * 64)
lines.append(f"    {'総資産（百万円）':<25} {hw['assets_24']:>12,} {bbd['assets_24']:>12,} {hw['assets_24']+bbd['assets_24']:>12,}")
lines.append(f"    {'純資産（百万円）':<25} {hw['equity_24']:>12,} {bbd['equity_24']:>12,} {hw['equity_24']+bbd['equity_24']:>12,}")
liabilities_hw  = hw["assets_24"]  - hw["equity_24"]
liabilities_bbd = bbd["assets_24"] - bbd["equity_24"]
lines.append(f"    {'負債合計（百万円）':<25} {liabilities_hw:>12,} {liabilities_bbd:>12,} {liabilities_hw+liabilities_bbd:>12,}")

blank()
h2("6-2. Pro Forma Balance Sheet（合併後）")

# パーチェス法での修正
# のれん = 取引対価（合算割当部分）- BBD純資産
goodwill = deal_val_mn - bbd["equity_24"]
if goodwill < 0:
    goodwill = 0
    note("のれんゼロ（取引対価がBBD純資産を下回るバーゲンパーチェスの可能性あり）")

calc_row(
    "のれん推定額 = 合併対価（割当部分）- BBD純資産",
    f"{deal_val_mn:.1f} - {bbd['equity_24']:,}",
    fmt(goodwill)
)
note(f"合併対価は割当対象株式分（{deal_val_mn:.1f}百万円）で計算。HW既保有分は再評価が発生しないと仮定")

# Pro Forma 総資産 = HW資産 + BBD資産 + のれん
pf_assets  = hw["assets_24"] + bbd["assets_24"] + goodwill
# Pro Forma 純資産 = HW純資産 + HW新規発行株式対価（=deal_val_mn）
pf_equity  = hw["equity_24"] + deal_val_mn
pf_liab    = pf_assets - pf_equity
# 確認
pf_liab_check = liabilities_hw + liabilities_bbd  # 負債はそのまま引き継ぎ

calc_row(
    "Pro Forma 総資産 = HW総資産 + BBD総資産 + のれん",
    f"{hw['assets_24']:,} + {bbd['assets_24']:,} + {goodwill:.1f}",
    fmt(pf_assets)
)
calc_row(
    "Pro Forma 純資産 = HW純資産 + 合併対価（株式発行増加分）",
    f"{hw['equity_24']:,} + {deal_val_mn:.1f}",
    fmt(pf_equity)
)
calc_row(
    "Pro Forma 負債合計 = HW負債 + BBD負債（引継）",
    f"{liabilities_hw:,} + {liabilities_bbd:,}",
    fmt(pf_liab_check)
)

blank()
h2("6-3. D/E レシオ（負債資本比率）")

de_hw_pre  = liabilities_hw  / hw["equity_24"]
de_bbd_pre = liabilities_bbd / bbd["equity_24"]
de_pf      = pf_liab_check   / pf_equity
calc_row(
    "HW（合併前）D/E = HW負債 / HW純資産",
    f"{liabilities_hw:,} / {hw['equity_24']:,}",
    f"{de_hw_pre:.2f}x"
)
calc_row(
    "BBD（合併前）D/E = BBD負債 / BBD純資産",
    f"{liabilities_bbd:,} / {bbd['equity_24']:,}",
    f"{de_bbd_pre:.2f}x"
)
calc_row(
    "合併後 D/E = PF負債 / PF純資産",
    f"{pf_liab_check:,} / {pf_equity:.1f}",
    f"{de_pf:.2f}x"
)

lines.append(f"""
    【バランスシートコメント】
    ・のれん推定額: {goodwill:.1f}百万円（{goodwill/1e3:.1f}十億円）
      合併後総資産に占めるのれん比率: {goodwill/pf_assets*100:.1f}%
    ・のれんは毎期償却対象（日本基準では20年以内定額）。
      仮に20年償却の場合: 年間 {goodwill/20:.1f}百万円の追加償却負担
    ・合併後D/Eは{de_pf:.2f}xと合併前HW（{de_hw_pre:.2f}x）から悪化する見込み
    ・BBDの総資産は {bbd['assets_24']:,}百万円と大きく、合算後レバレッジ管理が重要
""")

# ============================================================
# Ⅶ. Sensitivity Analysis
# ============================================================

h1("Ⅶ. Sensitivity Analysis（感応度分析）")

h2("7-1. EPS感応度マトリクス")
lines.append("    ■ Pro Forma EPS（発行済ベース）感応度テーブル")
lines.append("      縦軸: BBD営業利益改善額（税後）/ 横軸: シナジー率（HW純利益ベース）")
blank()

bbd_delta_range  = [-400, -300, -200, -100, 0, 100, 200, 300, 400]  # 百万円
synergy_pct_range = [0, 5, 10, 15, 20, 25, 30]  # %

# ヘッダー行
header = f"    {'BBDΔvs.Synergy%':>18}"
for s in synergy_pct_range:
    header += f" {s:>8}%"
lines.append(header)
lines.append("    " + "-" * (18 + 4 + 9 * len(synergy_pct_range)))

for bbd_delta in bbd_delta_range:
    bbd_contrib_tax = bbd_delta * (1 - TAX_RATE)  # 税後
    row_str = f"    BBD改善 {bbd_delta:>+5}百万"
    for syn_pct in synergy_pct_range:
        syn_amt = hw_ni_base * syn_pct / 100.0
        pf_ni_s = hw_ni_base + bbd_contrib_tax + syn_amt
        pf_eps_s = pf_ni_s * 1e6 / HW_SHARES_POST_MERGER
        row_str += f"  {pf_eps_s:>7.1f}円"
    lines.append(row_str)

blank()
h2("7-2. Accretion / Dilution マトリクス")
lines.append("    ■ Accretion / Dilution（%、発行済ベース、スタンドアロンEPS比）")
blank()

hw_eps_sa_ref = hw_ni_base * 1e6 / hw["shares"]

header2 = f"    {'BBDΔvs.Synergy%':>18}"
for s in synergy_pct_range:
    header2 += f" {s:>8}%"
lines.append(header2)
lines.append("    " + "-" * (18 + 4 + 9 * len(synergy_pct_range)))

breakeven_cells = []
for bbd_delta in bbd_delta_range:
    bbd_contrib_tax = bbd_delta * (1 - TAX_RATE)
    row_str = f"    BBD改善 {bbd_delta:>+5}百万"
    for syn_pct in synergy_pct_range:
        syn_amt = hw_ni_base * syn_pct / 100.0
        pf_ni_s = hw_ni_base + bbd_contrib_tax + syn_amt
        pf_eps_s = pf_ni_s * 1e6 / HW_SHARES_POST_MERGER
        acc_s = (pf_eps_s / hw_eps_sa_ref - 1) * 100
        row_str += f"  {acc_s:>+7.1f}%"
        if abs(acc_s) < 5.0:
            breakeven_cells.append((bbd_delta, syn_pct, acc_s))
    lines.append(row_str)

blank()
lines.append("    ■ EPS Accretionのブレークイーブン近傍（Accretion ±5%以内のセル）")
if breakeven_cells:
    for bc in breakeven_cells:
        lines.append(f"      BBD改善 {bc[0]:>+5}百万円 × シナジー {bc[1]:>2}%  →  Accretion: {bc[2]:+.1f}%")
else:
    lines.append("      該当セルなし（全域がDilutiveまたはAccretive）")

blank()
h2("7-3. ブレークイーブンラインのまとめ")

lines.append(f"""
    【感応度分析 総括】
    ・スタンドアロンEPS（FY2026予想）: {hw_eps_sa_ref:.2f}円
    ・合併後発行済株式数: {HW_SHARES_POST_MERGER:,}株（スタンドアロンから+{HW_NEW_SHARES:,}株、+{HW_NEW_SHARES/hw['shares']*100:.1f}%希薄化）

    ブレークイーブン条件（EPS希薄化ゼロ）:
    ① シナリオA（BBD貢献ゼロ）必要シナジー: {be_synergy_a_basic:.1f}百万円（税後）
       ≒ HW FY2026予想純利益の {be_synergy_a_basic/hw_ni_base*100:.1f}%
    ② シナリオB（BBD損失継続）必要シナジー: {be_synergy_b_basic:.1f}百万円（税後）
       ≒ HW FY2026予想純利益の {be_synergy_b_basic/hw_ni_base*100:.1f}%

    ・BBD FY2025の営業損失（△351百万円）がそのまま継続するシナリオBでは、
      スタンドアロンEPSを維持するだけで大規模なシナジーが必要となる
    ・現実的なシナジー水準（10〜20%程度）では、BBD損益が▲100百万円以上の
      赤字水準が継続する限り、EPS希薄化は不可避な見込み
    ・HWがBBDのサブスクARR（1,640百万円）を活用して収益化を加速させ、
      中期的にBBD損益をブレークイーブン以上に改善できるかが鍵
""")

# ============================================================
# Ⅷ. 総合評価・結論
# ============================================================

h1("Ⅷ. 総合評価・結論（Executive Summary）")

lines.append(f"""
  【取引構造の特徴】
  ・本取引は100%株式対価のため現金流出なし。HWにとって財務負担は軽微。
  ・合併比率0.50はHW株価2,885円ベースでBBD1株あたり1,442.5円相当。
    発表前日BBD株価1,591円に対し△9.3%ディスカウントであり、BBD株主には
    不利な条件だが、HW株主には資産効率上有利な取引条件となっている。

  【EPS影響の見通し】
  ・スタンドアロンEPS（FY2026予想）: {hw_eps_sa_ref:.1f}円
  ・シナリオAベース（BBD損益ゼロ）Pro Forma EPS: {scenarios["A（ベースケース）"]["pf_eps_basic"]:.1f}円
    → Accretion/Dilution: {scenarios["A（ベースケース）"]["accretion_basic"]:+.1f}%（希薄化）
  ・シナリオBダウンサイド（BBD損失継続）Pro Forma EPS: {scenarios["B（ダウンサイド）"]["pf_eps_basic"]:.1f}円
    → Accretion/Dilution: {scenarios["B（ダウンサイド）"]["accretion_basic"]:+.1f}%（希薄化）
  ・シナリオCシナジー（利益20%増）Pro Forma EPS: {scenarios["C（シナジーケース）"]["pf_eps_basic"]:.1f}円
    → Accretion/Dilution: {scenarios["C（シナジーケース）"]["accretion_basic"]:+.1f}%

  【リスク要因】
  ①  BBD FY2025の大幅赤字（営業損失△351百万円）は重大なダウンサイドリスク
  ②  のれん{goodwill:.0f}百万円（約{goodwill/1e3:.1f}十億円）の計上と将来の減損リスク
  ③  合算後レバレッジ（D/E {de_pf:.2f}x）の上昇
  ④  BBD ARR 1,640百万円の解約・縮小リスク

  【ポジティブ要因】
  ①  BBDのAIインフラ・データ基盤とHWのAIソリューションの補完関係
  ②  完全子会社化による意思決定の迅速化・コスト効率化
  ③  100%株式対価のためHW財務体力を温存
  ④  HW FY2026売上予想5,764百万円（対FY2025比+47.8%）の高成長軌道

  【総合判断】
  短期的にはEPS希薄化が発生するが（全シナリオ共通）、
  BBDの経営再建・シナジー実現が進めば中長期的なEPS回復・増益が期待できる。
  ただしFY2025の赤字幅が大きく、統合コストや競争環境次第では
  希薄化が長期化するリスクも無視できない。
  本合併の成否は、①BBD赤字体質の改善スピード、②ARRの維持・成長、
  ③AI事業領域でのシナジー実現可否の3点に集約される。
""")

# ============================================================
# 免責事項
# ============================================================

h1("免責事項（Disclaimer）")
lines.append("""
  本分析は公開情報・有価証券報告書等に基づく教育・参考目的の試算であり、
  投資助言・推奨を構成するものではありません。
  実際の合併比率・財務数値については各社公表資料をご確認ください。
  将来の業績・シナジーについては保証するものではなく、実際の結果と
  異なる可能性があります。
  分析における税率・財務構造は簡略化のため一部を仮定しています。
""")

# ============================================================
# ファイル出力
# ============================================================

report_text = "\n".join(lines)

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write(report_text)

print(report_text)
print(f"\n\n{'='*60}")
print(f"レポートを '{OUTPUT_FILE}' に出力しました。")
print(f"{'='*60}")
