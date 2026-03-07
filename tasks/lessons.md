# Lessons Learned

## EDINET API: Post-April 2024 Reform Doc Types
- **docTypeCode=140** (四半期報告書): ABOLISHED after April 2024. Q1/Q3 quarterly reports no longer filed to EDINET.
- **docTypeCode=130** (半期報告書, old): Still exists but rarely used for new filings.
- **docTypeCode=160** (半期報告書, NEW): This is the replacement. Semi-annual (H1/Q2) reports post-2024 use this doc type.
- **Q1/Q3 data**: Only available via TDNet 決算短信, NOT on EDINET.
- **Always include 160 in interim report searches** — it's the primary doc type for modern filings.

## EDINET Semi-Annual XBRL Context IDs (docType 160)
- Old quarterly contexts: `CurrentAccumulatedQ{n}Duration`, `CurrentQuarterInstant`
- New semi-annual contexts: `InterimDuration`, `InterimInstant`, `Prior1InterimDuration`, `Prior1InterimInstant`
- Parser must handle BOTH patterns in `identify_quarterly_contexts()`.

## Adaptive Search: Year Offset Must Include +1
- `fiscal_year_end` from the latest annual report covers the COMPLETED FY.
- The CURRENT (ongoing) FY's interim reports require `year_offset=+1` to generate correct candidate dates.
- Without +1, the search only finds old-FY interim data, missing the most recent one.

## Adaptive Search: Quarter-End Month Calculation
- FY start month = (fy_month % 12) + 1
- Quarter-end offsets from FY start: **[2, 5, 8]** (NOT [3, 6, 9])
- Example: March FY → April start → Q1=June, Q2=September, Q3=December

## API Call Efficiency
- Search one doc type at a time across all spiral offsets, not all doc types per offset.
- For post-2024: Q2 uses [160], Q1/Q3 use broad search [None] as fallback.
- `_search_single_date` with `doc_type_code=None` matches any doc type by secCode.
- Budget of ~50 API calls is sufficient with targeted searching.
