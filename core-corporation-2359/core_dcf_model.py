"""
Core Corporation (2359.T) DCFバリュエーションモデル

yfinanceから過去の財務データを取得し、5年間のFCFを予測、
WACC算出、ターミナルバリュー（2手法）、理論株価、感応度分析を実行する。
結果はコンソールに表示し、core_dcf_report.txt にも保存する。
"""

import sys
import io
import numpy as np
import yfinance as yf
import pandas as pd

# 表示設定
pd.set_option('display.float_format', lambda x: f'{x:,.0f}')
pd.set_option('display.width', 120)

# === 出力をコンソールとファイルの両方に書き出すクラス ===
class TeeWriter:
    """標準出力とファイルに同時書き込みするクラス"""
    def __init__(self, file_path):
        self.terminal = sys.stdout
        self.file = open(file_path, 'w', encoding='utf-8')

    def write(self, message):
        self.terminal.write(message)
        self.file.write(message)

    def flush(self):
        self.terminal.flush()
        self.file.flush()

    def close(self):
        self.file.close()


# ============================================================================
# 関数定義
# ============================================================================

def calc_wacc(risk_free_rate, beta, market_risk_premium,
              after_tax_cost_of_debt, de_ratio, size_premium=0.0):
    """
    WACCを計算する関数

    引数:
        risk_free_rate:         リスクフリーレート
        beta:                   ベータ
        market_risk_premium:    マーケットリスクプレミアム
        after_tax_cost_of_debt: 税引後負債コスト
        de_ratio:               D/E比率（負債÷株主資本）
        size_premium:           サイズプレミアム（小型株リスクプレミアム）

    戻り値:
        wacc, cost_of_equity, weight_equity, weight_debt
    """
    # 修正CAPMによる株主資本コスト: Ke = Rf + β × MRP + サイズプレミアム
    cost_of_equity = risk_free_rate + beta * market_risk_premium + size_premium

    # 負債比率と株主資本比率の算出: D/(D+E) = D/E / (1 + D/E)
    weight_debt = de_ratio / (1 + de_ratio)
    weight_equity = 1 - weight_debt

    # WACC = We × Ke + Wd × Kd(税引後)
    wacc = weight_equity * cost_of_equity + weight_debt * after_tax_cost_of_debt

    return wacc, cost_of_equity, weight_equity, weight_debt


def forecast_fcf(base_revenue, revenue_growth, operating_margin,
                 capex_ratio, tax_rate, wc_change_ratio, n_years=5):
    """
    n年間のFCFを予測する関数

    FCF = NOPAT + 減価償却費 - Capex - 運転資本の増加
    （減価償却費 = Capex と仮定）
    """
    prev_revenue = base_revenue
    forecast = []

    for i in range(n_years):
        revenue = prev_revenue * (1 + revenue_growth)
        operating_income = revenue * operating_margin
        nopat = operating_income * (1 - tax_rate)
        capex = revenue * capex_ratio
        depreciation = capex  # Capexと同額と仮定
        wc_increase = (revenue - prev_revenue) * wc_change_ratio
        fcf = nopat + depreciation - capex - wc_increase

        forecast.append({
            '売上高': revenue,
            '営業利益': operating_income,
            'NOPAT': nopat,
            '減価償却費': depreciation,
            '設備投資(Capex)': capex,
            '運転資本の増加': wc_increase,
            'FCF': fcf,
        })
        prev_revenue = revenue

    return forecast


def calc_intrinsic_price(base_revenue, revenue_growth, operating_margin,
                         capex_ratio, tax_rate, wc_change_ratio,
                         wacc, terminal_growth, net_debt, diluted_shares):
    """
    DCFモデルで理論株価を算出する関数（永久成長率法）

    エクイティバリュー ÷ 完全希薄化株式数で1株あたり理論株価を返す
    """
    forecast = forecast_fcf(base_revenue, revenue_growth, operating_margin,
                            capex_ratio, tax_rate, wc_change_ratio)

    # FCFの現在価値の合計
    fcfs = [row['FCF'] for row in forecast]
    pv_fcf = sum(f / (1 + wacc) ** (i + 1) for i, f in enumerate(fcfs))

    # ターミナルバリュー（永久成長率法）
    tv = fcfs[-1] * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_tv = tv / (1 + wacc) ** len(fcfs)

    # EV → エクイティバリュー → 理論株価
    ev = pv_fcf + pv_tv
    equity = ev - net_debt
    price = equity * 1_000_000 / diluted_shares

    return price


def get_row(df, key, all_dates):
    """データフレームから指定の行を取得し、百万円単位に変換"""
    if key in df.index:
        return df.loc[key].reindex(all_dates) / 1_000_000
    return pd.Series([None] * len(all_dates), index=all_dates)


# ============================================================================
# メイン処理
# ============================================================================

def main():
    # 出力をファイルにも保存
    report_path = "core_dcf_report.txt"
    tee = TeeWriter(report_path)
    sys.stdout = tee

    # === 銘柄データの取得 ===
    TICKER = "2359.T"
    stock = yf.Ticker(TICKER)
    info = stock.info

    print("=" * 80)
    print("Core Corporation (2359.T) DCF Valuation Model")
    print("=" * 80)
    print()

    # ====================================================================
    # セクション1: 過去の財務データ
    # ====================================================================
    income_stmt = stock.financials       # 損益計算書
    cashflow = stock.cashflow            # キャッシュフロー計算書
    balance = stock.balance_sheet        # 貸借対照表

    # 全期間の年度を収集し、古い順にソート
    all_dates = sorted(
        set(income_stmt.columns) | set(cashflow.columns) | set(balance.columns)
    )
    labels = [d.strftime('%Y-%m') for d in all_dates]

    # 各項目の抽出
    hist_revenue = get_row(income_stmt, 'Total Revenue', all_dates)
    hist_op_income = get_row(income_stmt, 'Operating Income', all_dates)
    hist_depr = get_row(cashflow, 'Depreciation And Amortization', all_dates)
    hist_capex = get_row(cashflow, 'Capital Expenditure', all_dates)
    hist_ca = get_row(balance, 'Current Assets', all_dates)
    hist_cl = get_row(balance, 'Current Liabilities', all_dates)
    hist_wc = hist_ca - hist_cl
    hist_wc_change = hist_wc.diff()
    hist_fcf = get_row(cashflow, 'Free Cash Flow', all_dates)
    hist_ocf = get_row(cashflow, 'Operating Cash Flow', all_dates)

    # サマリーテーブル
    hist_summary = pd.DataFrame({
        '売上高': hist_revenue.values,
        '営業利益': hist_op_income.values,
        '減価償却費': hist_depr.values,
        '設備投資': hist_capex.values,
        '運転資本': hist_wc.values,
        '運転資本の変動': hist_wc_change.values,
    }, index=labels).T

    print("=" * 80)
    print("1. 過去の財務データ")
    print("=" * 80)
    print("（単位: 百万円）")
    print()
    print(hist_summary.to_string())
    print()

    # 営業利益率
    print("営業利益率:")
    for i, label in enumerate(labels):
        r = hist_revenue.iloc[i]
        oi = hist_op_income.iloc[i]
        if pd.notna(r) and pd.notna(oi) and r != 0:
            print(f"  {label}: {oi / r * 100:.1f}%")
    print()

    # 売上高成長率
    print("売上高成長率（前年比）:")
    for i in range(1, len(labels)):
        prev = hist_revenue.iloc[i - 1]
        curr = hist_revenue.iloc[i]
        if pd.notna(prev) and pd.notna(curr) and prev != 0:
            print(f"  {labels[i]}: {(curr / prev - 1) * 100:+.1f}%")
    print()

    # FCF・営業CF
    if hist_fcf.notna().any():
        print("フリーキャッシュフロー（yfinance算出）:")
        for i, label in enumerate(labels):
            val = hist_fcf.iloc[i]
            if pd.notna(val):
                print(f"  {label}: {val:,.0f} 百万円")
        print()

    if hist_ocf.notna().any():
        print("営業キャッシュフロー:")
        for i, label in enumerate(labels):
            val = hist_ocf.iloc[i]
            if pd.notna(val):
                print(f"  {label}: {val:,.0f} 百万円")
        print()

    # 基礎データ
    print("現在の株価: {} 円".format(info.get('currentPrice', 'N/A')))
    shares_outstanding = info.get('sharesOutstanding', None)
    if shares_outstanding:
        print(f"発行済株式数: {shares_outstanding:,.0f}")
        print(f"時価総額: {info.get('marketCap', 0) / 1_000_000:,.0f} 百万円")

    total_debt = (info.get('totalDebt', 0) or 0) / 1_000_000
    total_cash = (info.get('totalCash', 0) or 0) / 1_000_000
    NET_DEBT = total_debt - total_cash

    print(f"\n有利子負債: {total_debt:,.0f} 百万円")
    print(f"現金及び現金同等物: {total_cash:,.0f} 百万円")
    print(f"ネットデット: {NET_DEBT:,.0f} 百万円")

    # ====================================================================
    # セクション2: WACC算出
    # ====================================================================
    RISK_FREE_RATE = 0.022
    BETA = 0.8
    MARKET_RISK_PREMIUM = 0.06
    AFTER_TAX_COST_OF_DEBT = 0.007
    DE_RATIO = 0.045
    SIZE_PREMIUM = 0.04

    WACC, COST_OF_EQUITY, WEIGHT_EQUITY, WEIGHT_DEBT = calc_wacc(
        RISK_FREE_RATE, BETA, MARKET_RISK_PREMIUM,
        AFTER_TAX_COST_OF_DEBT, DE_RATIO, SIZE_PREMIUM
    )

    print()
    print()
    print("=" * 80)
    print("2. WACC（加重平均資本コスト）の算出")
    print("=" * 80)
    print()
    print("【前提条件】")
    print(f"  リスクフリーレート (Rf):     {RISK_FREE_RATE * 100:.1f}%")
    print(f"  ベータ (beta):               {BETA:.1f}")
    print(f"  マーケットリスクプレミアム:   {MARKET_RISK_PREMIUM * 100:.1f}%")
    print(f"  サイズプレミアム:             {SIZE_PREMIUM * 100:.1f}%")
    print(f"  税引後負債コスト (Kd):       {AFTER_TAX_COST_OF_DEBT * 100:.1f}%")
    print(f"  D/E比率:                     {DE_RATIO * 100:.1f}%")
    print()
    print("【計算過程】")
    print(f"  株主資本コスト (Ke) = Rf + beta x MRP + サイズプレミアム")
    print(f"                      = {RISK_FREE_RATE*100:.1f}% + {BETA} x {MARKET_RISK_PREMIUM*100:.1f}% + {SIZE_PREMIUM*100:.1f}%")
    print(f"                      = {COST_OF_EQUITY * 100:.2f}%")
    print()
    print(f"  負債比率 (Wd)     = D/E / (1 + D/E) = {DE_RATIO*100:.1f}% / {(1+DE_RATIO)*100:.1f}% = {WEIGHT_DEBT * 100:.2f}%")
    print(f"  株主資本比率 (We) = 1 - Wd = {WEIGHT_EQUITY * 100:.2f}%")
    print()
    print(f"  WACC = We x Ke + Wd x Kd")
    print(f"       = {WEIGHT_EQUITY*100:.2f}% x {COST_OF_EQUITY*100:.2f}% + {WEIGHT_DEBT*100:.2f}% x {AFTER_TAX_COST_OF_DEBT*100:.1f}%")
    print(f"       = {WACC * 100:.2f}%")

    # ====================================================================
    # セクション3: 5年間のFCF予測
    # ====================================================================
    BASE_REVENUE = 24_599
    REVENUE_GROWTH = 0.072
    OPERATING_MARGIN = 0.1374
    CAPEX_RATIO = 0.0054
    TAX_RATE = 0.3062
    WC_CHANGE_RATIO = 0.10

    years = ['2026-03', '2027-03', '2028-03', '2029-03', '2030-03']
    forecast = forecast_fcf(BASE_REVENUE, REVENUE_GROWTH, OPERATING_MARGIN,
                            CAPEX_RATIO, TAX_RATE, WC_CHANGE_RATIO)

    # 表示用データフレーム
    display_data = []
    for row in forecast:
        display_data.append({
            '売上高': row['売上高'],
            '売上高成長率': REVENUE_GROWTH * 100,
            '営業利益': row['営業利益'],
            '営業利益率': OPERATING_MARGIN * 100,
            'NOPAT': row['NOPAT'],
            '減価償却費': row['減価償却費'],
            '設備投資(Capex)': row['設備投資(Capex)'],
            '運転資本の増加': row['運転資本の増加'],
            'FCF': row['FCF'],
        })
    fcf_df = pd.DataFrame(display_data, index=years).T

    print()
    print()
    print("=" * 80)
    print("3. 5年間のFCF予測")
    print("=" * 80)
    print()
    print("【前提条件】")
    print(f"  基準売上高（2025-03）: {BASE_REVENUE:,} 百万円")
    print(f"  売上高成長率: {REVENUE_GROWTH * 100:.1f}%")
    print(f"  営業利益率: {OPERATING_MARGIN * 100:.2f}%")
    print(f"  設備投資: 売上高の {CAPEX_RATIO * 100:.2f}%")
    print(f"  法人税実効税率: {TAX_RATE * 100:.2f}%")
    print(f"  減価償却費: Capexと同額")
    print(f"  運転資本の変動: 売上増分の {WC_CHANGE_RATIO * 100:.0f}%")
    print(f"  FCF = NOPAT + 減価償却費 - Capex - 運転資本の増加")
    print()
    print("（単位: 百万円）")
    print()

    pct_rows = ['売上高成長率', '営業利益率']
    for row_name in fcf_df.index:
        vals = []
        for col in fcf_df.columns:
            v = fcf_df.loc[row_name, col]
            if row_name in pct_rows:
                vals.append(f'{v:.2f}%')
            else:
                vals.append(f'{v:,.0f}')
        label = f'{row_name:<12s}'
        print(f"  {label}  {'  '.join(f'{v:>10s}' for v in vals)}")

    print()
    print(f"  * 減価償却費とCapexは同額のため相殺、FCF = NOPAT - 運転資本の増加")

    # ====================================================================
    # セクション4: DCFバリュエーション
    # ====================================================================
    TERMINAL_GROWTH = 0.015
    EXIT_MULTIPLE = 10.0
    DILUTED_SHARES = 14_844_000
    CURRENT_PRICE = 2_240

    # 予測最終年度の数値
    last_fcf = forecast[-1]['FCF']
    last_ebitda = forecast[-1]['営業利益'] + forecast[-1]['減価償却費']

    # FCFの現在価値を計算
    pv_fcfs = []
    for i, row in enumerate(forecast):
        year_num = i + 1
        discount_factor = 1 / (1 + WACC) ** year_num
        pv = row['FCF'] * discount_factor
        pv_fcfs.append({
            '年度': years[i],
            'FCF': row['FCF'],
            '割引係数': discount_factor,
            'FCFの現在価値': pv,
        })
    sum_pv_fcf = sum(item['FCFの現在価値'] for item in pv_fcfs)

    # 方法1: 永久成長率法（Gordon Growth Model）
    tv_gordon = last_fcf * (1 + TERMINAL_GROWTH) / (WACC - TERMINAL_GROWTH)
    pv_tv_gordon = tv_gordon / (1 + WACC) ** 5
    ev_gordon = sum_pv_fcf + pv_tv_gordon
    equity_gordon = ev_gordon - NET_DEBT
    price_gordon = equity_gordon * 1_000_000 / DILUTED_SHARES
    upside_gordon = (price_gordon / CURRENT_PRICE - 1) * 100

    # 方法2: Exit Multiple法
    tv_exit = last_ebitda * EXIT_MULTIPLE
    pv_tv_exit = tv_exit / (1 + WACC) ** 5
    ev_exit = sum_pv_fcf + pv_tv_exit
    equity_exit = ev_exit - NET_DEBT
    price_exit = equity_exit * 1_000_000 / DILUTED_SHARES
    upside_exit = (price_exit / CURRENT_PRICE - 1) * 100

    print()
    print()
    print("=" * 80)
    print("4. DCFバリュエーション")
    print("=" * 80)

    # FCF現在価値の内訳
    print()
    print("【FCFの現在価値】（単位: 百万円）")
    print()
    print(f"  {'年度':<10s}  {'FCF':>10s}  {'割引係数':>10s}  {'現在価値':>10s}")
    print(f"  {'-'*10}  {'-'*10}  {'-'*10}  {'-'*10}")
    for item in pv_fcfs:
        print(f"  {item['年度']:<10s}  {item['FCF']:>10,.0f}  {item['割引係数']:>10.4f}  {item['FCFの現在価値']:>10,.0f}")
    print(f"  {'合計':<10s}  {'':>10s}  {'':>10s}  {sum_pv_fcf:>10,.0f}")

    # ネットデット
    print()
    print("【ネットデット】")
    print(f"  有利子負債:         {total_debt:>10,.0f} 百万円")
    print(f"  現金及び現金同等物: {total_cash:>10,.0f} 百万円")
    print(f"  ネットデット:       {NET_DEBT:>10,.0f} 百万円（マイナス=ネットキャッシュ）")

    # 方法1
    print()
    print("-" * 80)
    print("【方法1】永久成長率法（Gordon Growth Model）")
    print(f"  ターミナル成長率: {TERMINAL_GROWTH * 100:.1f}%")
    print()
    print(f"  ターミナルバリュー = FCF(最終年) x (1+g) / (WACC-g)")
    print(f"                    = {last_fcf:,.0f} x (1+{TERMINAL_GROWTH*100:.1f}%) / ({WACC*100:.2f}%-{TERMINAL_GROWTH*100:.1f}%)")
    print(f"                    = {tv_gordon:,.0f} 百万円")
    print(f"  TVの現在価値       = {pv_tv_gordon:,.0f} 百万円")
    print()
    print(f"  エンタープライズバリュー (EV) = {sum_pv_fcf:,.0f} + {pv_tv_gordon:,.0f}")
    print(f"                               = {ev_gordon:,.0f} 百万円")
    print(f"  エクイティバリュー            = {ev_gordon:,.0f} - ({NET_DEBT:,.0f})")
    print(f"                               = {equity_gordon:,.0f} 百万円")
    print(f"  完全希薄化株式数              = {DILUTED_SHARES:,} 株")
    print()
    print(f"  -> 理論株価 = {price_gordon:,.0f} 円（現在株価 {CURRENT_PRICE:,} 円 / 乖離率: {upside_gordon:+.1f}%）")

    # 方法2
    print()
    print("-" * 80)
    print("【方法2】Exit Multiple法")
    print(f"  EV/EBITDA倍率: {EXIT_MULTIPLE:.1f}x")
    print(f"  最終年度EBITDA: {last_ebitda:,.0f} 百万円")
    print()
    print(f"  ターミナルバリュー = EBITDA(最終年) x EV/EBITDA倍率")
    print(f"                    = {last_ebitda:,.0f} x {EXIT_MULTIPLE:.1f}")
    print(f"                    = {tv_exit:,.0f} 百万円")
    print(f"  TVの現在価値       = {pv_tv_exit:,.0f} 百万円")
    print()
    print(f"  エンタープライズバリュー (EV) = {sum_pv_fcf:,.0f} + {pv_tv_exit:,.0f}")
    print(f"                               = {ev_exit:,.0f} 百万円")
    print(f"  エクイティバリュー            = {ev_exit:,.0f} - ({NET_DEBT:,.0f})")
    print(f"                               = {equity_exit:,.0f} 百万円")
    print()
    print(f"  -> 理論株価 = {price_exit:,.0f} 円（現在株価 {CURRENT_PRICE:,} 円 / 乖離率: {upside_exit:+.1f}%）")

    # サマリー
    print()
    print("=" * 80)
    print("【バリュエーションサマリー】")
    print("=" * 80)
    print(f"                        {'永久成長率法':>14s}  {'Exit Multiple法':>14s}")
    print(f"  ターミナルバリュー    {tv_gordon:>14,.0f}  {tv_exit:>14,.0f}")
    print(f"  TV現在価値            {pv_tv_gordon:>14,.0f}  {pv_tv_exit:>14,.0f}")
    print(f"  EV                    {ev_gordon:>14,.0f}  {ev_exit:>14,.0f}")
    print(f"  エクイティバリュー    {equity_gordon:>14,.0f}  {equity_exit:>14,.0f}")
    print(f"  理論株価              {price_gordon:>14,.0f}  {price_exit:>14,.0f}")
    print(f"  乖離率                {upside_gordon:>13.1f}%  {upside_exit:>13.1f}%")
    print(f"                                         （単位: 百万円、株価は円）")

    # ====================================================================
    # セクション5: 感応度分析
    # ====================================================================
    print()
    print()
    print("=" * 80)
    print("5. 感応度分析")
    print("=" * 80)

    # --- テーブル1: WACC × ターミナル成長率 ---
    wacc_range = np.arange(0.090, 0.125, 0.005)    # 9.0%〜12.0%
    tg_range = np.arange(0.005, 0.030, 0.005)      # 0.5%〜2.5%

    print()
    print("【テーブル1】WACC x ターミナル成長率 -> 理論株価（円）")
    print(f"  （固定: 売上成長率 {REVENUE_GROWTH*100:.1f}%, 営業利益率 {OPERATING_MARGIN*100:.2f}%）")
    print()

    header = f"  {'WACC \\ TGR':>12s}"
    for tg in tg_range:
        header += f"  {tg*100:>7.1f}%"
    print(header)
    print("  " + "-" * (12 + len(tg_range) * 9))

    for w in wacc_range:
        row = f"  {w*100:>7.1f}%    "
        for tg in tg_range:
            price = calc_intrinsic_price(
                BASE_REVENUE, REVENUE_GROWTH, OPERATING_MARGIN,
                CAPEX_RATIO, TAX_RATE, WC_CHANGE_RATIO,
                w, tg, NET_DEBT, DILUTED_SHARES
            )
            row += f"  {price:>7,.0f}"
        print(row)

    print()
    print(f"  * 現在株価: {CURRENT_PRICE:,} 円")

    # --- テーブル2: 売上成長率 × 営業利益率 ---
    rg_range = np.arange(0.05, 0.11, 0.01)         # 5%〜10%
    om_range = np.arange(0.11, 0.165, 0.01)        # 11%〜16%

    print()
    print()
    print("【テーブル2】売上成長率 x 営業利益率 -> 理論株価（円）")
    print(f"  （固定: WACC {WACC*100:.2f}%, ターミナル成長率 {TERMINAL_GROWTH*100:.1f}%）")
    print()

    header = f"  {'Growth \\ OPM':>12s}"
    for om in om_range:
        header += f"  {om*100:>7.0f}%"
    print(header)
    print("  " + "-" * (12 + len(om_range) * 9))

    for rg in rg_range:
        row = f"  {rg*100:>7.0f}%    "
        for om in om_range:
            price = calc_intrinsic_price(
                BASE_REVENUE, rg, om,
                CAPEX_RATIO, TAX_RATE, WC_CHANGE_RATIO,
                WACC, TERMINAL_GROWTH, NET_DEBT, DILUTED_SHARES
            )
            row += f"  {price:>7,.0f}"
        print(row)

    print()
    print(f"  * 現在株価: {CURRENT_PRICE:,} 円")

    # ====================================================================
    # 完了
    # ====================================================================
    print()
    print("=" * 80)
    print(f"レポートを {report_path} に保存しました。")
    print("=" * 80)

    # 標準出力を元に戻す
    sys.stdout = tee.terminal
    tee.close()


if __name__ == '__main__':
    main()
