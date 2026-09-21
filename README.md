# Independent Japanese Equity Research

Coverage of under-researched TSE small caps, published in English from primary Japanese sources (EDINET / TDnet).

Most listed companies on the Tokyo Stock Exchange below ~¥50B market cap have no sell-side coverage in any language. This repository publishes full research cycles on selected names — thesis, pre-registered estimates, and public reconciliation against actuals — together with the Python + Excel valuation models behind each call.

## Coverage

| Company | Ticker | Latest Report | Rating | Status |
|---|---|---|---|---|
| Core Corporation | 2359.T | [Post-Earnings Verification Report](reports/Core_2359_Verification_Report.pdf) (Jul 2026) | HOLD | Full cycle complete: thesis → pre-registered estimates → public reconciliation |
| SpiderPlus & Co. | 4192.T | [Initial Coverage](reports/SpiderPlus_4192_Equity_Research.pdf) (Jul 2026) | BUY | Active coverage |
| DMW Corporation | 6365.T | [v2 Revised](reports/DMW_6365_Equity_Research_v2.pdf) (Jul 2026) | BUY | Target ¥8,120 — v2 corrects the SOTP sensitivity matrix, unifies the DCF exit-multiple base at 8.5x, and states the target methodology explicitly (see Revision Note in the report) |

Each report links to its underlying model in [`reports/`](reports/) (e.g. [4192 reverse-DCF / implied-multiple model](reports/4192_market_analysis_20260618.xlsx)).

## Track Record — Core Corporation (2359.T), Full Cycle

The core of this repository is not the ratings — it is the verification discipline. Estimates are published before earnings, timestamped in this repository and on LinkedIn, and reconciled publicly afterward, right or wrong.

- **Mar 29, 2026** — FY2026 estimates published before earnings: OP ¥4,050M, EPS ¥201
- **Apr 28, 2026** — Actuals: OP ¥3,819M, EPS ¥200.40 — **EPS error 0.3%, net income error 0.7%**; revenue missed (**-4.6%**)
- **Stock outcome** — BUY at ¥2,305 → ¥1,980 (**-14.1%** as of Jul 17, 2026): the thesis failed on the guidance leg despite the earnings hit
- **[The verification report](reports/Core_2359_Verification_Report.pdf)** reconciles the full cycle publicly — what was right, what was lucky, what was wrong, and the resulting method changes (implied-expectations audit, dual-trigger catalysts, freshness stamps on positioning data)

Both halves of that record matter. The earnings estimates landed within 1% on the bottom line; the trade still lost 14%. Publishing the miss alongside the hit — and the process changes it forced — is the point.

## Methodology

- **DCF (5-scenario)** — Base / Upside / Management / Downside 1 / Downside 2, perpetuity-growth and exit-multiple terminal values
- **Reverse DCF (implied growth)** — inverts the model to quantify the growth expectations embedded in the current price before any catalyst
- **Implied multiple positioning** — the EV/Sales / EV/EBITDA / PER the market currently awards, positioned against the peer distribution
- **SOTP** — segment-level valuation for multi-business names
- **Pre-registration & public verification cycle** — estimates published before the print, reconciled publicly after
- **Python + EDINET XBRL automation pipeline** — financial data extracted directly from EDINET filings; models generated and validated by code in [`scripts/`](scripts/) and [`templates/`](templates/)

## Earnings Pre-emption Screener (`screener/`) — research in progress

**Purpose.** Automate the "read every earnings release" step: detect the moment a change in the slope of a
company's results shows up *as reported numbers* (not plans or promises), before the market prices it, and narrow
the TSE universe to a short weekly list of evidence. The screener outputs evidence and its sources; it does not
issue buy/sell recommendations. Valuation (reverse DCF) and the final judgement stay with the model pipeline above
and a human.

**Four layers** (plus one observation-only layer):

| Layer | Package | What it does |
|---|---|---|
| 1. Fetch | `screener/fetch/` | TDnet daily archive (earnings releases, forecast revisions), EDINET bulk (annual / half-year reports), J-Quants (universe, prices, earnings summaries) |
| 2. Extract | `screener/extract/` | XBRL → cumulative and stand-alone quarterly financials, guidance and revisions, disclosure flags |
| 3. Signals | `screener/signals/` | Point-in-time evidence scores (S1–S5 …); S5c = progress ratio against the company's own seasonality (display only) |
| 4. Report | `screener/report/` | Weekly screen, earnings-window filter, paper-trading ledger, pre-registered measurement reports |
| (Technical) | `screener/technical/` | Price-reaction records only. Isolated by tests: the scoring path never imports it |

**Current phase (Sep 2026).** Forward paper-trading (records only — this code never places live orders) running
alongside pre-registered shadow variants. Every rule change is first registered in
[`docs/backtest_acceptance_criteria.md`](docs/backtest_acceptance_criteria.md) with its parameters and pass/fail
criteria fixed *before* the measurement; results, negative ones included, are logged in
[`docs/calibration_backlog.md`](docs/calibration_backlog.md).

**Main measured results so far** (in-sample, ~5 years of disclosures; section numbers refer to the calibration log):

- **The initial reaction to an upward guidance revision is completed in the overnight gap.** For the 1,825 upward
  revisions disclosed after the close (5 years, 43.9% price coverage), the move from the next open to the close is
  +0.00% (date-clustered t −1.55); the reaction is the gap. The median gap widened from +0.32% (2022) to
  +1.25% (2026). (§44, §56)
- **S5c (progress vs. the company's own 5-year seasonal median, R ≥ 1.30) has predictive power but cannot be
  captured.** The upward-revision rate within 90 days is 30.9% vs. a 12.0% base rate (+18.9pt, t 7.51). But a fixed
  one-month hold averages +0.71% (t 0.64; −0.39% vs. TOPIX), and exiting on the revision event (5/10/20-day
  variants) fails all three pre-registered checks. Conclusion: *front-running revisions is not viable with the
  current data*. S5c stays at zero weight, display only. (§51–§57)
- The raw progress-ratio signal (S5) adds little: +0.39pt over a matched base rate, from a thin sample. (§42, §46)
- Volume-profile "acceptance zones" did not improve entries or exits in pre-registered tests; the technical layer
  remains observation-only. (§39, §43)

**Data.** Raw TDnet / EDINET / J-Quants data, databases, caches and logs are never committed (J-Quants terms
prohibit redistribution). The repository holds code, configuration, documentation and the author's own analysis
outputs.

## Repository Layout

```
reports/     Published research PDFs and their final Excel models
models/      Generated DCF / SOTP / market-analysis workbooks (drafts included)
templates/   Generic model templates (DCF, SOTP, market analysis, narrative stage)
scripts/     Per-ticker runners, EDINET fetcher/parser, validators
data/        Per-ticker assumption overrides, segments, adjustments (JSON). Comps CSVs stay local, not committed
screener/    Earnings pre-emption screener (fetch / extract / signals / report / technical)
docs/        Design docs, pre-registrations, calibration log and analysis notes
```

## Disclaimer

This repository is for educational and informational purposes only. Nothing here constitutes investment advice, a recommendation, or a solicitation to buy or sell any security. The author may hold positions in covered companies; any such position is disclosed in the corresponding report. See [DISCLAIMER.md](DISCLAIMER.md).

## Contact

- LinkedIn: [linkedin.com/in/ryosuke-sato-44148339b](https://www.linkedin.com/in/ryosuke-sato-44148339b)
