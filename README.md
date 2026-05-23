# Japanese Equity Research — Modeling Toolkit

An independent research project on Japanese small/mid-cap equities. To ground investment decisions in *quantified scenarios* rather than price action, this repository accumulates a modeling toolkit combining DCF / Comps / SOTP / Implied Growth Analysis / Market Scorecard / Narrative Stage Assessment.

## Coverage (as of May 2026)

| Ticker | Sector | Price | Verdict | Total Score | Action |
|---|---|---|---|---|---|
| Core (2359) | IT / Defense | ¥2,006 | BUY | +0.55 | Hold; awaiting late-June catalyst |
| Torishima (6363) | Pumps | — | HOLD | — | Fairly valued on 5-method average |
| Denyo-sha (6365) | Pumps | ¥5,490 | CAUTION | -0.55 | Entry on hold; wait for ¥4,800–5,000 |
| IHI (7013) | Heavy Industry | ¥2,824 | (in progress) | — | Re-valuing with MTP-reflected DCF |

## Analysis Framework

The framework uses a layered valuation structure. The first three layers measure **Gap 1** (the *surprise gap*: realized results vs prior expectations). A fourth layer, **Block 5**, measures **Gap 2** (the *catalyst gap*: today's expectation vs the expectation about to form).

### 1. Absolute Value Layer
- **DCF (Perpetual Growth Method)**: perpetuity growth model
- **DCF (Exit Multiple)**: exit-multiple terminal value
- **SOTP (Sum-of-the-Parts)**: sum of segment-level valuations

### 2. Relative Value Layer
- **Comps (EV/EBITDA)**: peer comparison
- **Comps (PER)**: price-to-earnings comparison

### 3. Market Expectation Layer
- **Implied Growth Analysis**: reverse-engineers the growth rate (α) the market prices into the share price
- **Market Scorecard**: composite buy/sell judgment from a 4-factor weighted score

### 4. Narrative Stage Layer (Block 5) — *new*
- **Narrative Stage Assessment**: scores where a stock sits in the lifecycle of a narrative re-rating, to identify doubler candidates before the re-rating spreads. See [docs/narrative_stage.md](docs/narrative_stage.md).

## Market Scorecard — 4-Factor Design

| Factor | Weight | What it measures |
|---|---|---|
| ① Implied Growth | 40% | 1 − market α |
| ② Price Momentum | 25% | 3-month change |
| ③ Margin Balance | 20% | Margin ratio × short-interest depth |
| ④ Forecast Gap | 15% | My forecast vs company guidance |

**Verdict scale (5 tiers):**

| Total Score | Verdict |
|---|---|
| ≥ +1.0 | STRONG BUY |
| +0.5 to +1.0 | BUY |
| -0.5 to +0.5 | HOLD |
| -1.0 to -0.5 | CAUTION |
| ≤ -1.0 | AVOID |

## Block 5 — Narrative Stage Assessment

The Market Scorecard (Blocks 1–4) is strong at catching **Gap 1** (the surprise gap, observed via reverse-DCF). It cannot score **Gap 2** — the gap between the market's *current* expectation and the expectation *about to form*. Doublers are usually triggered by Gap 2 and then compounded by Gap 1 (e.g. a mid-term plan lifts expectations → next earnings beat the new bar → expectations rise again). Block 5 fills this hole.

### Six axes (each scored 0 / +1 / +2)

Four core axes track the narrative re-rating using a *fire* metaphor:

- **Axis 1 — Label Status** *(is there a spark?)*: how far the old label has been rewritten into a new one. Label-rewrite progress.
- **Axis 2A — Catalyst Potential** *(is there fuel?)*: existence of real, factual sources that could drive the rewrite (counts facts, not wishes).
- **Axis 2B — Catalyst Realization** *(has it ignited?)*: whether the catalyst has fired into the market.
- **Axis 3 — Diffusion Stage** *(how wide is the fire?)*: how far the new label has reached (retail → media → sell-side → institutions). **This axis gates the stage.**
- **Axis 4 — Earnings Materiality**: is the narrative backed by real earnings?
- **Axis 5 — Narrative Durability**: does the story survive the next cycle?

### Two key design choices

1. **Stage is gated by REACH (Axis 3), not the sum.** A brilliant story nobody has heard is Stage 1 — price moves with recognition, not with the truth of the thesis. Axis 1 (label rewrite) is *independent* of Axis 3 (reach) and does **not** move the stage; instead the gap between them is reported as a headroom / reversal-risk read (Axis 1 > Axis 3 = room to spread; Axis 1 < Axis 3 = shallow-rewrite risk).
2. **If Earnings Materiality (Axis 4) = 0, the stage is capped at 2.** This blocks narrative-only stocks that excite the market until the theme cools. The cap only ever *lowers* a stage; it never raises one.

### Five stages

Stage 0 (no label change) → Stage 1 (emergence) → Stage 2 (diffusion) → Stage 3 (recognition / sell-side) → Stage 4 (peer-group rewrite).

### Final Verdict matrix

The Block 2–4 score crosses with the Block 5 stage:

| Block 2–4 | Stage 0–1 | Stage 2 | Stage 3–4 |
|---|---|---|---|
| BUY | **STRONG BUY** (doubler zone) | BUY | HOLD (consider exit) |
| HOLD | OBSERVE | HOLD | TRIM |
| CAUTION | AVOID | CAUTION | EXIT |

The doubler sweet spot is **cheap on fundamentals × Stage 0–1** (top-left).

### Companion metric: absolute earnings impact

Stage measures the *phase* of the narrative, not the *size of the prize*. A separate check on TAM × share × segment OPM (vs current operating profit) estimates the absolute uplift — because an identical stage profile can hide wildly different ceilings.

### Implementation

- `templates/narrative_stage_template.py` — core scoring logic (`assess()`), stage determination, headroom signal, earnings-impact and verdict calculation.
- `templates/test_narrative_stage.py` — validation against 4 reference cases (Core, Denyo-sha, IHI, retroactive Fujikura 2023) plus 3 edge cases. Run with `PYTHONIOENCODING=utf-8 python test_narrative_stage.py`.
- `templates/market_analysis_template.py` — emits a "Narrative Stage" sheet (six-axis table, stage judgment, headroom signal, earnings impact, verdict, radar chart) alongside the existing Implied Growth and Market Scorecard sheets. Backward compatible: skipped when no `narrative` config is supplied.
- Full design: [docs/narrative_stage.md](docs/narrative_stage.md).

## May 2026 — Key Improvements

### Improvement 1: Higher numerical precision in Implied Growth Analysis

**Problem**: Excel's `FORECAST` linear regression could not handle the convexity of the α–price relationship (the exponential influence of terminal value), producing a systematic bias (0.025–0.032) in the low-α region.

**Fix**: Replaced with local linear interpolation (`MATCH` + `INDEX`). Interpolating only between adjacent points preserves the local character of the convex function.

**Validation**: Mathematically guarantees that the α of the DCF Base PGM Target is exactly 1.0000.

**Measured impact**:
- Core (2359) α at ¥2,006: -0.083 → -0.108
- Denyo-sha (6365) α at ¥5,490: +0.738 → +0.770
- Both retained Total Score / Verdict
- Core's Downside-1 scenario upgraded from Mildly Bullish → BULLISH

### Improvement 2: Automated price-data retrieval

`yfinance` integration auto-fetches: current price, 3-month-ago close, 1-month-ago close, and the 3-month high/low.

Design:
- Config values take priority when present (backward compatible)
- Explicit `ValueError` on fetch failure (avoids silent `None` propagation)
- High/low window unified to 3 months (consistent with momentum calc)

## Directory Structure

```
ryosuke-japanese-equity-research/
├── templates/        # Generic templates (DCF, SOTP, Market Analysis, Narrative Stage, etc.)
├── scripts/          # Per-ticker execution scripts (incl. recalc.py formula checker)
├── models/           # Per-ticker DCF/SOTP Excel (regenerable, so .gitignored)
├── reports/          # Generated Market Analysis reports
├── data/overrides/   # Per-ticker assumption overrides (JSON)
├── docs/             # Per-ticker analysis notes, design docs, LinkedIn posts
├── notes/            # Session handoffs, learning notes
└── comps/            # Comps data
```

## Roadmap

### Short term (within weeks)
- **IHI (7013)**: reflect FY2026/3 results in the DCF; integrate the new mid-term plan (Phases 1–3, through FY2034) into the Management scenario and run Market Analysis
- **Core (2359)**: await late-June catalyst; hold stop-loss at ¥1,750
- **Denyo-sha (6365)**: wait for a pullback to ¥4,800–5,000

### Medium term
- Re-value Torishima (6363) with the new template
- Initial analysis of Kimura Chemical Plants (6378) (fusion theme)
- Automated margin-balance data retrieval (Kabutan scraping)
- Multi-ticker batch wrapper (`auto_generate_for_ticker`)
- Backtest Block 5 retroactively on past doublers (Lasertec 2019, Hitachi 2020–21, Mitsubishi Heavy 2022) for statistical validation

## Tech Stack

- Python 3.x
- openpyxl (Excel generation/manipulation)
- yfinance (price data)
- pandas / numpy (data processing)
- EDINET API (financial data, `scripts/edinet_fetcher.py`)

## Design Philosophy

Draw a scenario and wait; when reality breaks it, draw a new one. Don't get tossed around by price action.

The purpose of quantification is not to make judgments *correct* but to make them *consistent* — an anchor for returning to one's own assumptions instead of being swept along by the market's pessimism or optimism.

A BUY/CAUTION verdict is the median of a probability distribution; the final call integrates quantitative and qualitative judgment.

## License / Disclaimer

- This repository is for personal study and record-keeping.
- It is not investment advice.
- All analyses are point-in-time and change with market conditions.
