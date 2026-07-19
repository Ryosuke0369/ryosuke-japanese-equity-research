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

## Repository Layout

```
reports/     Published research PDFs and their final Excel models
templates/   Generic model templates (DCF, SOTP, market analysis, narrative stage)
scripts/     Per-ticker runners, EDINET fetcher/parser, validators
data/        Per-ticker assumption overrides (JSON) and comps inputs (CSV)
docs/        Design docs and analysis notes
```

## Disclaimer

This repository is for educational and informational purposes only. Nothing here constitutes investment advice, a recommendation, or a solicitation to buy or sell any security. The author may hold positions in covered companies; any such position is disclosed in the corresponding report. See [DISCLAIMER.md](DISCLAIMER.md).

## Contact

- LinkedIn: [linkedin.com/in/ryosuke-sato-44148339b](https://www.linkedin.com/in/ryosuke-sato-44148339b)
