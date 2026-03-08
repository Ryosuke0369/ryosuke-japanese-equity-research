# NWC Dynamic Forecasting Implementation

## Status: COMPLETE

### Steps
- [x] Step 1: Add `nwc_pct` and `base_year_nwc` to config
- [x] Step 2: Update row constants (+3 rows: R_DRV_NWC_PCT=31, R_NWC=43, R_CHG_NWC=44)
- [x] Step 3: Add C19 assumption (Base Year NWC) to Assumptions section
- [x] Step 4: Add NWC% driver row label + input cell (yellow/blue)
- [x] Step 5: Add NWC and Change in NWC rows to waterfall
- [x] Step 6: Update UFCF formula: `NOPAT + D&A - Capex - Change in NWC`
- [x] Step 7: Update `row_labels_fcf` list with NWC rows
- [x] Step 8: Update `calc_dcf_pgm` and `calc_dcf_exit` with NWC logic

### Review
- PGM/Exit sections use relative row definitions → auto-adjusted, no code change needed
- WACC section starts at row 20, C19 is the last assumption → no collision
- All Excel f-string formulas use row constants → consistent
- Year 1 Change in NWC references C19 (Base Year NWC); Year 2+ references prior year NWC
- Sensitivity functions now include NWC impact in FCF calculation
