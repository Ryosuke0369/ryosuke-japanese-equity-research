"""
Core Corporation (2359.T) Comparable Company Analysis

同業他社とのバリュエーション比較分析・Implied Valuation・
DCFとCompsの統合バリュエーションサマリーを出力する。
結果は core_comps_report.txt にも保存する。
"""

import sys
import yfinance as yf
import pandas as pd
import numpy as np


# === 出力をコンソールとファイルの両方に書き出すクラス ===
class TeeWriter:
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


def fetch_company_data(ticker, name):
    """
    1社分の財務データ・バリュエーション指標を取得する関数

    yfinanceのinfo、financials、cashflowから売上高・営業利益・純利益・
    EBITDA・各種マルチプル・収益性指標を抽出する。
    """
    stock = yf.Ticker(ticker)
    info = stock.info
    financials = stock.financials
    cashflow = stock.cashflow

    # 株価・時価総額
    price = info.get('currentPrice', None)
    market_cap = info.get('marketCap', 0) or 0

    # EV = 時価総額 + 有利子負債 - 現金
    total_debt = info.get('totalDebt', 0) or 0
    total_cash = info.get('totalCash', 0) or 0
    ev = market_cap + total_debt - total_cash

    # 損益計算書から取得
    revenue = None
    revenue_prev = None
    op_income = None
    net_income = None

    if not financials.empty:
        if 'Total Revenue' in financials.index:
            revenue = financials.loc['Total Revenue'].iloc[0]
            if len(financials.columns) >= 2:
                revenue_prev = financials.loc['Total Revenue'].iloc[1]
        if 'Operating Income' in financials.index:
            op_income = financials.loc['Operating Income'].iloc[0]
        if 'Net Income' in financials.index:
            net_income = financials.loc['Net Income'].iloc[0]

    # 減価償却費 → EBITDA
    depreciation = 0
    if not cashflow.empty and 'Depreciation And Amortization' in cashflow.index:
        dep_val = cashflow.loc['Depreciation And Amortization'].iloc[0]
        if pd.notna(dep_val):
            depreciation = dep_val
    ebitda = (op_income or 0) + depreciation if op_income is not None else None

    # マルチプル算出
    ev_ebitda = ev / ebitda if ebitda and ebitda > 0 else None
    ev_revenue = ev / revenue if revenue and revenue > 0 else None
    ev_ebit = ev / op_income if op_income and op_income > 0 else None
    per = info.get('trailingPE', None)
    pbr = info.get('priceToBook', None)

    # 収益性指標
    op_margin = (op_income / revenue * 100) if (op_income and revenue and revenue > 0) else None
    roe = info.get('returnOnEquity', None)
    if roe is not None:
        roe = roe * 100

    # 売上成長率
    rev_growth = None
    if revenue and revenue_prev and revenue_prev > 0:
        rev_growth = (revenue / revenue_prev - 1) * 100

    return {
        '銘柄': name,
        'ティッカー': ticker,
        '株価': price,
        '時価総額': market_cap,
        'EV': ev,
        '売上高': revenue,
        'EBITDA': ebitda,
        '営業利益': op_income,
        '純利益': net_income,
        'EV/EBITDA': ev_ebitda,
        'EV/Revenue': ev_revenue,
        'EV/EBIT': ev_ebit,
        'PER': per,
        'PBR': pbr,
        '営業利益率(%)': op_margin,
        'ROE(%)': roe,
        '売上成長率(%)': rev_growth,
    }


def implied_price_ev(core_metric, multiple, net_debt, shares):
    """EVベースのマルチプルから理論株価を算出"""
    return (core_metric * multiple - net_debt) / shares


def implied_price_eq(core_metric, multiple, shares):
    """エクイティベースのマルチプルから理論株価を算出"""
    return core_metric * multiple / shares


def main():
    report_path = "core_comps_report.txt"
    tee = TeeWriter(report_path)
    sys.stdout = tee

    CURRENT_PRICE = 2_240

    print("=" * 100)
    print("Core Corporation (2359.T) Comparable Company Analysis")
    print("=" * 100)
    print()

    # === 対象銘柄 ===
    companies = [
        ("2359.T", "コア"),
        ("2317.T", "システナ"),
        ("3626.T", "TIS"),
        ("9719.T", "SCSK"),
        ("4684.T", "オービック"),
        ("8056.T", "BIPROGY"),
        ("9682.T", "DTS"),
        ("4674.T", "クレスコ"),
        ("9759.T", "NSD"),
    ]

    # === データ取得 ===
    print("データ取得中...")
    results = []
    for ticker, name in companies:
        print(f"  {name} ({ticker}) ...", end=" ")
        try:
            data = fetch_company_data(ticker, name)
            results.append(data)
            print("OK")
        except Exception as e:
            print(f"FAILED ({e})")
    print()

    if not results:
        print("データを取得できませんでした。")
        sys.stdout = tee.terminal
        tee.close()
        return

    df = pd.DataFrame(results)
    peers = df[df['銘柄'] != 'コア']
    core = df[df['銘柄'] == 'コア'].iloc[0]

    # コアの基礎データ
    core_price = core['株価']
    core_mktcap = core['時価総額']
    core_ev_val = core['EV']
    net_debt = core_ev_val - core_mktcap
    shares = core_mktcap / core_price
    core_book_value = core_mktcap / core['PBR'] if core['PBR'] and core['PBR'] > 0 else None

    # ====================================================================
    # 1. 基本情報
    # ====================================================================
    print("=" * 100)
    print("1. 基本情報（単位: 百万円、株価は円）")
    print("=" * 100)
    print()

    basic_fields = ['株価', '時価総額', 'EV', '売上高', 'EBITDA', '営業利益', '純利益']
    header = f"  {'銘柄':<10s}"
    for f in basic_fields:
        header += f"  {f:>12s}"
    print(header)
    print("  " + "-" * (10 + len(basic_fields) * 14))

    for _, row in df.iterrows():
        line = f"  {row['銘柄']:<10s}"
        for f in basic_fields:
            val = row[f]
            if pd.isna(val) or val is None:
                line += f"  {'N/A':>12s}"
            elif f == '株価':
                line += f"  {val:>12,.0f}"
            else:
                line += f"  {val / 1_000_000:>12,.0f}"
        print(line)

    # ====================================================================
    # 2. バリュエーションマルチプル
    # ====================================================================
    multiples = ['EV/EBITDA', 'EV/Revenue', 'EV/EBIT', 'PER', 'PBR']

    print()
    print()
    print("=" * 100)
    print("2. バリュエーションマルチプル")
    print("=" * 100)
    print()

    header = f"  {'銘柄':<10s}"
    for m in multiples:
        header += f"  {m:>12s}"
    print(header)
    print("  " + "-" * (10 + len(multiples) * 14))

    for _, row in df.iterrows():
        line = f"  {row['銘柄']:<10s}"
        for m in multiples:
            val = row[m]
            if pd.notna(val) and val is not None:
                line += f"  {val:>11.1f}x"
            else:
                line += f"  {'N/A':>12s}"
        print(line)

    # パーセンタイル算出
    pct_data = {}
    print()
    print(f"  {'--- 同業他社統計 ---'}")

    for pct_label, pct_val in [('25th', 25), ('中央値(50th)', 50), ('75th', 75)]:
        line = f"  {pct_label:<12s}"
        for m in multiples:
            vals = peers[m].dropna()
            if len(vals) > 0:
                p = np.percentile(vals, pct_val)
                pct_data.setdefault(m, {})[pct_label] = p
                line += f"  {p:>11.1f}x"
            else:
                line += f"  {'N/A':>12s}"
        print(line)

    line = f"  {'コア(現在)':<12s}"
    for m in multiples:
        val = core[m]
        if pd.notna(val) and val is not None:
            line += f"  {val:>11.1f}x"
        else:
            line += f"  {'N/A':>12s}"
    print(line)

    # ====================================================================
    # 3. 収益性・成長性
    # ====================================================================
    perf_fields = ['営業利益率(%)', 'ROE(%)', '売上成長率(%)']

    print()
    print()
    print("=" * 100)
    print("3. 収益性・成長性")
    print("=" * 100)
    print()

    header = f"  {'銘柄':<10s}"
    for f in perf_fields:
        header += f"  {f:>14s}"
    print(header)
    print("  " + "-" * (10 + len(perf_fields) * 16))

    for _, row in df.iterrows():
        line = f"  {row['銘柄']:<10s}"
        for f in perf_fields:
            val = row[f]
            if pd.isna(val) or val is None:
                line += f"  {'N/A':>14s}"
            else:
                line += f"  {val:>13.1f}%"
        print(line)

    print()
    print(f"  {'--- 同業他社統計 ---'}")
    for label, func in [('平均', 'mean'), ('中央値', 'median')]:
        line = f"  {label:<10s}"
        for f in perf_fields:
            vals = peers[f].dropna()
            if len(vals) > 0:
                line += f"  {getattr(vals, func)():>13.1f}%"
            else:
                line += f"  {'N/A':>14s}"
        print(line)

    line = f"  {'コア':<10s}"
    for f in perf_fields:
        val = core[f]
        if pd.notna(val) and val is not None:
            line += f"  {val:>13.1f}%"
        else:
            line += f"  {'N/A':>14s}"
    print(line)

    # ====================================================================
    # 4. Implied Valuation（理論株価レンジ）
    # ====================================================================
    print()
    print()
    print("=" * 100)
    print("4. Implied Valuation - 理論株価レンジ")
    print("=" * 100)
    print()
    print(f"  現在株価: {CURRENT_PRICE:,} 円")
    print(f"  ネットデット: {net_debt / 1_000_000:,.0f} 百万円（マイナス=ネットキャッシュ）")
    print(f"  発行済株式数: {shares:,.0f}")
    print()

    # ヘッダー
    header = (f"  {'マルチプル':<14s}"
              f"  {'コア財務値':>12s}"
              f"  {'25th':>8s}  {'株価':>8s}  {'乖離率':>8s}"
              f"  {'50th':>8s}  {'株価':>8s}  {'乖離率':>8s}"
              f"  {'75th':>8s}  {'株価':>8s}  {'乖離率':>8s}")
    print(header)
    print("  " + "-" * 118)

    valuation_methods = [
        ('EV/EBITDA',  core['EBITDA'],    'ev'),
        ('EV/Revenue', core['売上高'],     'ev'),
        ('EV/EBIT',    core['営業利益'],   'ev'),
        ('PER',        core['純利益'],     'equity'),
        ('PBR',        core_book_value,   'equity'),
    ]

    # 全手法の理論株価を蓄積
    all_prices = {'25th': [], '中央値(50th)': [], '75th': []}

    for multi_name, core_metric, method_type in valuation_methods:
        if core_metric is None or pd.isna(core_metric) or core_metric <= 0:
            continue
        if multi_name not in pct_data:
            continue

        metric_mm = core_metric / 1_000_000
        line = f"  {multi_name:<14s}  {metric_mm:>10,.0f}M"

        for pct_label in ['25th', '中央値(50th)', '75th']:
            m_val = pct_data[multi_name].get(pct_label)
            if m_val is None:
                line += f"  {'N/A':>8s}  {'N/A':>8s}  {'N/A':>8s}"
                continue

            if method_type == 'ev':
                price = implied_price_ev(core_metric, m_val, net_debt, shares)
            else:
                price = implied_price_eq(core_metric, m_val, shares)

            upside = (price / CURRENT_PRICE - 1) * 100
            all_prices[pct_label].append(price)
            line += f"  {m_val:>7.1f}x  {price:>7,.0f}  {upside:>+7.1f}%"

        print(line)

    # ====================================================================
    # 5. 理論株価レンジ（サマリー）
    # ====================================================================
    print()
    print()
    print("=" * 100)
    print("5. Comps 理論株価レンジ（サマリー）")
    print("=" * 100)
    print()

    header = f"  {'':<16s}  {'25th':>10s}  {'50th':>10s}  {'75th':>10s}"
    print(header)
    print("  " + "-" * 52)

    for multi_name, core_metric, method_type in valuation_methods:
        if core_metric is None or pd.isna(core_metric) or core_metric <= 0:
            continue
        if multi_name not in pct_data:
            continue

        line = f"  {multi_name:<16s}"
        for pct_label in ['25th', '中央値(50th)', '75th']:
            m_val = pct_data[multi_name].get(pct_label)
            if m_val is None:
                line += f"  {'N/A':>10s}"
                continue
            if method_type == 'ev':
                p = implied_price_ev(core_metric, m_val, net_debt, shares)
            else:
                p = implied_price_eq(core_metric, m_val, shares)
            line += f"  {p:>10,.0f}"
        print(line)

    print("  " + "-" * 52)

    for agg_label, agg_func in [('平均', np.mean), ('中央値', np.median)]:
        line = f"  {agg_label:<16s}"
        for pct_label in ['25th', '中央値(50th)', '75th']:
            vals = all_prices[pct_label]
            line += f"  {agg_func(vals):>10,.0f}" if vals else f"  {'N/A':>10s}"
        print(line)

    for agg_label, agg_func in [('乖離率(平均)', np.mean), ('乖離率(中央値)', np.median)]:
        line = f"  {agg_label:<16s}"
        for pct_label in ['25th', '中央値(50th)', '75th']:
            vals = all_prices[pct_label]
            if vals:
                v = agg_func(vals)
                line += f"  {(v / CURRENT_PRICE - 1) * 100:>+9.1f}%"
            else:
                line += f"  {'N/A':>10s}"
        print(line)

    print()
    print(f"  現在株価: {CURRENT_PRICE:,} 円")

    # ====================================================================
    # 6. フットボールチャート
    # ====================================================================
    print()
    print()
    print("=" * 100)
    print("6. バリュエーションレンジ（フットボールチャート）")
    print("=" * 100)
    print()

    for multi_name, core_metric, method_type in valuation_methods:
        if core_metric is None or pd.isna(core_metric) or core_metric <= 0:
            continue
        if multi_name not in pct_data:
            continue

        p25v = pct_data[multi_name].get('25th')
        p50v = pct_data[multi_name].get('中央値(50th)')
        p75v = pct_data[multi_name].get('75th')

        calc = implied_price_ev if method_type == 'ev' else lambda m, v, nd, s: implied_price_eq(m, v, s)

        if method_type == 'ev':
            low = implied_price_ev(core_metric, p25v, net_debt, shares)
            mid = implied_price_ev(core_metric, p50v, net_debt, shares)
            high = implied_price_ev(core_metric, p75v, net_debt, shares)
        else:
            low = implied_price_eq(core_metric, p25v, shares)
            mid = implied_price_eq(core_metric, p50v, shares)
            high = implied_price_eq(core_metric, p75v, shares)

        bar_left = int(max(0, (mid - low) / 40))
        bar_right = int(max(0, (high - mid) / 40))
        print(f"  {multi_name:<12s}  {low:>6,.0f} |{'=' * bar_left}*{'=' * bar_right}| {high:>6,.0f}")

    print(f"  {'現在株価':<12s}         {CURRENT_PRICE:,} 円")

    # ====================================================================
    # 7. 統合バリュエーションサマリー（DCF + Comps）
    # ====================================================================
    # DCFの結果（core_dcf_model.py の算出値）
    DCF_GORDON = 2_481      # 永久成長率法
    DCF_EXIT = 3_098        # Exit Multiple法

    # Compsの中央値ベースの理論株価
    comps_25 = {m: [] for m in ['25th', '中央値(50th)', '75th']}
    for multi_name, core_metric, method_type in valuation_methods:
        if core_metric is None or pd.isna(core_metric) or core_metric <= 0:
            continue
        if multi_name not in pct_data:
            continue
        for pct_label in ['25th', '中央値(50th)', '75th']:
            m_val = pct_data[multi_name].get(pct_label)
            if m_val is None:
                continue
            if method_type == 'ev':
                p = implied_price_ev(core_metric, m_val, net_debt, shares)
            else:
                p = implied_price_eq(core_metric, m_val, shares)
            comps_25[pct_label].append(p)

    comps_med_25 = np.median(comps_25['25th']) if comps_25['25th'] else None
    comps_med_50 = np.median(comps_25['中央値(50th)']) if comps_25['中央値(50th)'] else None
    comps_med_75 = np.median(comps_25['75th']) if comps_25['75th'] else None

    print()
    print()
    print("*" * 100)
    print("7. 統合バリュエーションサマリー（DCF + Comps）")
    print("*" * 100)
    print()
    print(f"  現在株価: {CURRENT_PRICE:,} 円")
    print()

    header = f"  {'手法':<40s}  {'理論株価':>10s}  {'乖離率':>10s}"
    print(header)
    print("  " + "-" * 64)

    summary_rows = [
        ("DCF - 永久成長率法 (TGR 1.5%)", DCF_GORDON),
        ("DCF - Exit Multiple法 (10.0x)", DCF_EXIT),
    ]

    # Comps各マルチプル（50thベース）
    for multi_name, core_metric, method_type in valuation_methods:
        if core_metric is None or pd.isna(core_metric) or core_metric <= 0:
            continue
        if multi_name not in pct_data:
            continue
        m_val = pct_data[multi_name].get('中央値(50th)')
        if m_val is None:
            continue
        if method_type == 'ev':
            p = implied_price_ev(core_metric, m_val, net_debt, shares)
        else:
            p = implied_price_eq(core_metric, m_val, shares)
        summary_rows.append((f"Comps - {multi_name} (50th: {m_val:.1f}x)", p))

    for label, price in summary_rows:
        upside = (price / CURRENT_PRICE - 1) * 100
        print(f"  {label:<40s}  {price:>10,.0f}  {upside:>+9.1f}%")

    print("  " + "-" * 64)

    # 全手法の統計
    all_summary_prices = [p for _, p in summary_rows]
    avg_price = np.mean(all_summary_prices)
    med_price = np.median(all_summary_prices)
    min_price = np.min(all_summary_prices)
    max_price = np.max(all_summary_prices)

    print(f"  {'全手法 平均':<40s}  {avg_price:>10,.0f}  {(avg_price/CURRENT_PRICE-1)*100:>+9.1f}%")
    print(f"  {'全手法 中央値':<40s}  {med_price:>10,.0f}  {(med_price/CURRENT_PRICE-1)*100:>+9.1f}%")
    print(f"  {'全手法 レンジ':<40s}  {min_price:>5,.0f} - {max_price:>5,.0f}")
    print()

    # DCF vs Compsの比較サマリー
    print("  " + "-" * 64)
    print(f"  {'DCF レンジ':<40s}  {DCF_GORDON:>5,.0f} - {DCF_EXIT:>5,.0f}")
    if comps_med_25 and comps_med_75:
        print(f"  {'Comps レンジ (25th-75th中央値)':<40s}  {comps_med_25:>5,.0f} - {comps_med_75:>5,.0f}")
    print(f"  {'現在株価':<40s}  {CURRENT_PRICE:>10,}")
    print()

    # フットボールチャート風
    print("  [バリュエーションマップ]")
    print()
    print(f"  DCF永久成長率法      {DCF_GORDON:>6,} |")
    print(f"  DCF Exit Multiple    {DCF_EXIT:>6,}       |")
    if comps_med_25 and comps_med_50 and comps_med_75:
        print(f"  Comps 25th(中央値)   {comps_med_25:>6,.0f} |")
        print(f"  Comps 50th(中央値)   {comps_med_50:>6,.0f}    |")
        print(f"  Comps 75th(中央値)   {comps_med_75:>6,.0f}          |")
    print(f"  現在株価             {CURRENT_PRICE:>6,} *")
    print()

    print("*" * 100)
    print(f"レポートを {report_path} に保存しました。")
    print("*" * 100)

    sys.stdout = tee.terminal
    tee.close()


if __name__ == '__main__':
    main()
