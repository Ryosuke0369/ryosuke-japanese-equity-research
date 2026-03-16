## Known Issues

- Small-cap companies (especially TSE Standard/Growth) may use NonConsolidatedMember context instead of ConsolidatedMember
- Manufacturing companies use CostOfProductsManufactured instead of CostOfSales for COGS
- Always run data verification (revenue/cogs/oi check) before generating DCF model for a new ticker
- NWC Base Year uses latest annual (FY-end) BS values, not LTM snapshot, to avoid seasonal distortions in AR/Inv/AP
- This is especially important for order-driven businesses (equipment makers, construction) where receivables fluctuate significantly by quarter
- Base Year Revenue uses latest FY actuals (not LTM) so that DCF projection Year 1 connects naturally to the last historical FY column
- LTM Revenue is kept in C20 as a reference value for stub period discounting
- ContractAssets (契約資産) is excluded from accounts_receivable to avoid inflating DSO for progress-billing companies (e.g., equipment makers using percentage-of-completion revenue recognition)
