"""test_dcf_crosscheck_matcher.py — SOTP cross-check label matching

Guards sotp_template.match_crosscheck_key(), which maps an Executive Summary
row label to one of the four cross-check methods.

Why this exists: the cross-check used to read fixed rows 16-19, which imports
one method's number under another method's name whenever the Valuation Summary
is restructured (bank-type models promote DDM / Residual Income to the top).
Matching by label fixed that, but label matching has its own failure mode —
a DDM row labelled "... perpetuity growth" was captured as the DCF PGM value,
displacing the real one. These cases lock in both directions.
"""

import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:
        pass

from sotp_template import match_crosscheck_key

CASES = [
    # --- plain DCF layout ---
    ('DCF - Perpetuity Growth',                          'pgm_fair_value'),
    ('DCF - Exit Multiple',                              'exit_fair_value'),
    ('Comps - EV/EBITDA Median',                         'comps_ev_ebitda'),
    ('Comps - EV/Sales Median',                          'comps_ev_ebitda'),
    ('Comps - PER Median',                               'comps_per'),

    # --- restructured bank-type summary: DDM / RI must NOT feed DCF keys ---
    ('DDM (Dividend Discount Model, perpetuity growth)', None),
    ('DDM - Gordon Growth',                              None),
    ('Residual Income',                                  None),
    ('Residual Income (perpetuity)',                     None),
    ('Exit strategy note (DDM)',                         None),

    # --- EV/EBIT vs EV/EBITDA exclusivity (substring trap, both ways) ---
    # 'ev/ebit' is a substring of 'ev/ebitda': matching the short form would
    # swallow every EBITDA row. An EBIT multiple must not fill an EBITDA key.
    ('Comps - EV/EBIT Median',                           None),

    # --- 'per' must be a word, not a substring ---
    ('Comps - Peer group median',                        None),
    ('Holding period return',                            None),

    # --- summary rows that must never be mistaken for a method ---
    ('Target Price (Mid)',                               None),
    ('Integrated Valuation Range',                       None),
    ('Current Price',                                    None),
]


def main():
    print('=' * 70)
    print('SOTP DCF cross-check label matcher')
    print('=' * 70)
    failures = 0
    for label, expected in CASES:
        got = match_crosscheck_key(label, set())
        ok = (got == expected)
        failures += (not ok)
        print(f"{'PASS' if ok else 'FAIL'}  {label:<52} -> {got} (expected {expected})")

    # Order matters too: the real DCF row must still win when a DDM row with
    # 'perpetuity' in its label sits above it.
    rows = ['DDM (Dividend Discount Model, perpetuity growth)',
            'Residual Income',
            'DCF - Perpetuity Growth',
            'DCF - Exit Multiple',
            'Comps - EV/EBIT Median',
            'Comps - PER Median']
    taken, hits = set(), {}
    for lbl in rows:
        key = match_crosscheck_key(lbl, taken)
        if key:
            hits[key] = lbl
            taken.add(key)
    expected_hits = {
        'pgm_fair_value': 'DCF - Perpetuity Growth',
        'exit_fair_value': 'DCF - Exit Multiple',
        'comps_per': 'Comps - PER Median',
    }
    ok = hits == expected_hits
    failures += (not ok)
    print(f"\n{'PASS' if ok else 'FAIL'}  restructured-summary row order")
    for k, v in sorted(hits.items()):
        print(f"        {k:18} <- {v}")
    if not ok:
        print(f"        expected {expected_hits}")

    print('\n' + '=' * 70)
    print(f"Result: {len(CASES) + 1 - failures}/{len(CASES) + 1} checks passed")
    print('=' * 70)
    if failures:
        sys.exit(1)


if __name__ == '__main__':
    main()
