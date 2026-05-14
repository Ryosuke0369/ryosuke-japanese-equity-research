"""Independent verification of Block 3 alpha values.

Recomputes the alpha-scan implied prices in Python (mirroring the Excel
formulas) and verifies that the new local-linear interpolation gives the
expected alphas for the key checkpoints.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates'))

from market_analysis_template import extract_dcf_data, ALPHA_VALUES


def implied_prices_for_dd(dd):
    """Reproduce the row-45 implied-share-price grid for each alpha."""
    fy_rev = dd['fy_actual_total_rev']
    fy_opm = dd['fy_actual_opm']
    base_growth = dd['base_total_growth']
    base_opm = dd['base_total_opm']
    da, capex, nwc = dd['da'], dd['capex'], dd['nwc_change']
    wacc = dd['wacc']
    tax = dd['tax_rate']
    tg = dd['terminal_growth']
    stub = dd['stub_fraction']
    shares = dd['shares']
    net_debt = dd['net_debt']

    out = []
    for alpha in ALPHA_VALUES:
        rev = fy_rev
        revs = []
        for yr in range(5):
            rev *= (1 + alpha * base_growth[yr])
            revs.append(rev)
        ops = [revs[yr] * (fy_opm + alpha * (base_opm[yr] - fy_opm)) for yr in range(5)]
        fcfs = [
            max(0.0, ops[yr] * (1 - tax)) + da[yr] - capex[yr] - nwc[yr]
            for yr in range(5)
        ]
        sum_pv = sum(fcfs[yr] / (1 + wacc) ** (stub + yr) for yr in range(5))
        if wacc > tg:
            tv = fcfs[-1] * (1 + tg) / (wacc - tg)
        else:
            tv = 0
        pv_tv = tv / (1 + wacc) ** (stub + 4)
        ev = sum_pv + pv_tv
        equity = ev - net_debt
        price = equity * 1_000_000 / shares if shares else 0.0
        out.append(price)
    return out


def local_interp_alpha(target_price, alphas, prices):
    """Mirror the new INDEX/MATCH local linear interpolation in Block 3."""
    # MATCH(price, prices, 1) -> largest index i s.t. prices[i] <= target_price
    idx = -1
    for i, p in enumerate(prices):
        if p <= target_price:
            idx = i
        else:
            break
    if idx < 0 or idx >= len(prices) - 1:
        # Out of range; clamp like Excel would (extrapolation off the high end
        # would error in MATCH(...,1) when there's no smaller price; we surface
        # that as "off-grid").
        return None
    a_lo, a_hi = alphas[idx], alphas[idx + 1]
    p_lo, p_hi = prices[idx], prices[idx + 1]
    return a_lo + (target_price - p_lo) / (p_hi - p_lo) * (a_hi - a_lo)


def global_forecast_alpha(target_price, alphas, prices):
    """Excel FORECAST = simple linear regression slope through all points."""
    n = len(alphas)
    mean_x = sum(prices) / n
    mean_y = sum(alphas) / n
    cov = sum((prices[i] - mean_x) * (alphas[i] - mean_y) for i in range(n))
    var = sum((prices[i] - mean_x) ** 2 for i in range(n))
    slope = cov / var
    intercept = mean_y - slope * mean_x
    return slope * target_price + intercept


def verify(label, dcf_path, segment_layout, checkpoints):
    print(f'\n=== {label} ===')
    dd = extract_dcf_data(dcf_path, segment_layout)
    prices = implied_prices_for_dd(dd)
    print('Implied prices on alpha grid:')
    for a, p in zip(ALPHA_VALUES, prices):
        print(f'  alpha={a:+.2f}  price={p:>10,.0f}')
    print()
    for name, target_price, expected_alpha in checkpoints:
        local = local_interp_alpha(target_price, ALPHA_VALUES, prices)
        glb = global_forecast_alpha(target_price, ALPHA_VALUES, prices)
        local_str = f'{local:+.4f}' if local is not None else 'OFF-GRID'
        exp_str = f'{expected_alpha:+.4f}' if expected_alpha is not None else 'n/a'
        print(f'  {name}: price={target_price:,}  '
              f'local_interp={local_str}  forecast(global)={glb:+.4f}  '
              f'expected~{exp_str}')


if __name__ == '__main__':
    here = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(here, '..'))

    seg_2359 = {
        'segments': [
            {'name': 'a', 'dcf_fy26_cell': 'F6',  'dcf_growth_base_row': 34, 'dcf_opm_base_row': 41},
            {'name': 'b', 'dcf_fy26_cell': 'F11', 'dcf_growth_base_row': 48, 'dcf_opm_base_row': 55},
            {'name': 'c', 'dcf_fy26_cell': 'F16', 'dcf_growth_base_row': 62, 'dcf_opm_base_row': 69},
        ]
    }
    verify(
        'コア (2359)',
        os.path.join(project_root, 'models', '2359_DCF_Model_20260509.xlsx'),
        seg_2359,
        [
            ('Current Price',       2006, -0.108),
            ('DCF Base PGM Target', 3290,  1.000),
            ('Entry Price',         2127, None),
        ],
    )

    seg_6365 = {
        'segments': [
            {'name': 'a', 'dcf_fy26_cell': 'F6',  'dcf_growth_base_row': 34, 'dcf_opm_base_row': 41},
            {'name': 'b', 'dcf_fy26_cell': 'F11', 'dcf_growth_base_row': 48, 'dcf_opm_base_row': 55},
            {'name': 'c', 'dcf_fy26_cell': 'F16', 'dcf_growth_base_row': 62, 'dcf_opm_base_row': 69},
        ]
    }
    verify(
        '電業社 (6365)',
        os.path.join(project_root, 'models', '6365_DCF_Model_20260413.xlsx'),
        seg_6365,
        [
            ('Current Price',       5490,  0.770),
            ('DCF Base PGM Target', 5882,  1.000),
        ],
    )
