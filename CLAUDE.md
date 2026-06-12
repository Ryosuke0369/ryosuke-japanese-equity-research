## Known Issues

- Small-cap companies (especially TSE Standard/Growth) may use NonConsolidatedMember context instead of ConsolidatedMember
- Manufacturing companies use CostOfProductsManufactured instead of CostOfSales for COGS
- IFRS filers (e.g., IHI 7013, Sony, etc.) use IFRS-suffixed XBRL tags (RevenueIFRS, CostOfSalesIFRS, OperatingProfitLossIFRS, etc.) — edinet_parser.py now supports these tags with priority over J-GAAP equivalents
- Always run data verification (revenue/cogs/oi check) before generating DCF model for a new ticker
- NWC Base Year uses latest annual (FY-end) BS values, not LTM snapshot, to avoid seasonal distortions in AR/Inv/AP
- This is especially important for order-driven businesses (equipment makers, construction) where receivables fluctuate significantly by quarter
- Base Year Revenue uses latest FY actuals (not LTM) so that DCF projection Year 1 connects naturally to the last historical FY column
- LTM Revenue is kept in C20 as a reference value for stub period discounting
- ContractAssets (契約資産) is excluded from accounts_receivable to avoid inflating DSO for progress-billing companies (e.g., equipment makers using percentage-of-completion revenue recognition)

## NWC Model Methods

Two methods for projecting Net Working Capital, toggled via `nwc_method` in overrides JSON:

### `"days"` (default)
- Projects AR/Inv/AP individually using DSO/DIH/DPO days
- Scenario matrix has 3 blocks: DSO, DIH, DPO
- Best when EDINET XBRL captures full trade receivables/payables

### `"revenue_pct"`
- Projects total NWC as a single percentage of revenue
- Scenario matrix has 1 block: NWC % of Revenue
- Use when EDINET only captures partial receivables (e.g., `accounts_receivable` misses 受取手形, 電子記録債権, 契約資産)
- `trade_receivables_total` / `trade_payables_total` items capture broader operating receivables/payables for base year NWC computation

### Critical Invariant
- Row 19 of NWC Schedule = Change in NWC for BOTH methods
- DCF Model always references `='NWC Schedule'!{col}19` — this must not change

## Segment / Driver Analysis
- overrides JSON に `segments` キーを追加することで、Segment AnalysisとDriver Analysisシートが自動生成される
- `segments` が未定義の場合、既存6シートのみ生成（後方互換）
- 各セグメントの `driver_type` で計算ロジックが分岐:
  - `backlog`: 受注残 → 検収 → 売上（装置メーカー向け）
  - `manmonth`: 人月 × 稼働率 × 単価 + ソリューション売上（ITサービス向け）
  - `growth_rate`: 前年比成長率ベース（汎用）
  - `manual`: 売上・OPを直接入力
  - `retail`: 店舗数ロールフォワード + SSSG + 新規店寄与（小売・外食向け）
    - historical: revenue, op, store_count（期末）
    - projections: new_stores, closures, sssg, new_store_months（デフォルト6）
    - 売上 = 既存店売上(前期Rev×(1+SSSG)) + 新規店売上(出店数×avg_rev×寄与月/12)
    - avg_store_rev = 前期Segment Revenue ÷ 前期末店舗数（自動算出）
    - KPI行: Avg Store Count = (期首+期末)/2
  - `subscription`: ARRブリッジ（チャーン/拡大/新規分離）+ NRRフォールバック（SaaS向け）
    - historical: revenue, op, arr_end
    - projections: nrr, churn_rate（任意）, new_arr
    - churn_rate未入力時: churn=0, expansion=NRR-1 で自動フォールバック
    - churn_rate入力時: expansion = NRR-(1-churn) で自動逆算
    - Revenue = (期首ARR + 期末ARR) / 2
    - KPI行: NRR%, Gross Churn%, ARR YoY Growth
- Segment Analysis は DCF Model の Revenue/EBIT の single source of truth（方式A: Segment→DCF完全リンク）
- DCF Model の Revenue/EBIT は `='Segment Analysis'!{col}{total_rev/op_row}` で直接参照
- COGS は `=Revenue - SGA - EBIT` で逆算（バックカルキュレーション）、COGS%/Revenue Growth%も逆算表示
- SGA%のみ従来通りDCF Model側のシナリオ入力マトリクスからCHOOSE参照
- Segment Analysis 内の Revenue/OP Margin は `=CHOOSE('DCF Model'!$D$27, ...)` でシナリオ切替
- NWC% consolidated inputs（Segment Analysis下部）は `nwc_method == "revenue_pct"` の場合のみ表示（days方式では非表示）
- Segment Scenario Input Matrix（シート下部）に5シナリオ×セグメント別の Revenue/OP Margin を入力
- `scenario_projections` キーで Base 以外のシナリオ値を定義（未定義時は `projections` の値を全シナリオで使用）
- `segments` が未定義の場合、従来の Revenue Growth × Base Year 方式にフォールバック（後方互換）
- Driver Analysis から Segment Analysis への数式リンクはなし（両シート独立、値はJSON入力で共有）
- projection 配列の長さは --years オプションと一致させること

### scenario_projections 構造
```json
{
  "segments": [{
    "projections": { "revenue": [...], "op_margin": [...] },
    "scenario_projections": {
      "Upside": { "revenue": [...], "op_margin": [...] },
      "Management": { "revenue": [...], "op_margin": [...] },
      "Downside 1": { "revenue": [...], "op_margin": [...] },
      "Downside 2": { "revenue": [...], "op_margin": [...] }
    }
  }]
}
```
- Base シナリオは `projections` の値を使用（scenario_projections に "Base" キーは不要）
- 未定義のシナリオは `projections` にフォールバック

### Capex/D&A方式切替
- `capex_method`: "revenue_pct"（デフォルト）または "direct"
- `da_method`: "revenue_pct"（デフォルト）または "direct"
- direct方式: overridesに `capex_direct` / `da_direct` オブジェクトで historical/projections 配列を指定
- revenue_pct方式: 従来通り `capex_pct` / `da_pct` (= `capex_pct_revenue` / `da_pct_revenue`) を使用
- キーが存在しない場合はrevenue_pctとして動作（後方互換）
- direct方式でもprojection配列が足りない年はrevenue_pct式にフォールバック
- Sensitivity Analysis (calc_dcf_pgm/calc_dcf_exit) でもdirect方式の値が使われる
- 注意: `capex_pct` / `da_pct` はdirect方式でもフォールバック用に残すこと

## リスクフリーレート
- テンプレートデフォルト: 0.022（2.2%、日本10年国債利回りベース）
- overridesの`risk_free`キーで企業別に上書き可能
- 金利環境の変化に応じてデフォルト値の定期見直しを推奨

### 受注生産型企業のNWC処理
- 契約資産が大きい装置メーカー等では、広義NWC（契約資産込み）は不安定
- 推奨: base_year_*をoverridesで狭義NWC（Pattern B: 電子記録含む、契約資産・前受金除外）に上書き
- nwc_pctを狭義ベースの比率（例: -0.05〜-0.10）に設定
- backlogドライバーとの二重カウントを回避

## Overrides バリデーション（2026-06-12導入 — 黙ってデフォルトに落ちない）

- `scripts/overrides_validator.py` が generate_dcf.py の JSON 読込直後に契約を強制する。
  **未知キー / ネスト構造（wacc_inputs等）/ 独自シナリオ名 / 配列長不一致 / `__CONFIRM__`
  残存 はすべて実行前にエラー停止**。エラーなく完走した実行は全キーが消費されている。
- 契約仕様の正本: **`docs/overrides_schema.md`**。新銘柄の overrides は推測でなくこれに従う。
- `__CONFIRM__` 残存は既定でエラー。検証ラン等で意図的に自動値を使う場合のみ
  `--allow-unconfirmed` を付ける（最終ランでは禁止）。
- コメント・メモは `_` 接頭辞キーに書く（常に許可）。
- 生成完了時に「Effective WACC inputs」(C7:C12相当) と Comps 5社名がコンソールに出る。
  **必ず目視照合する**こと。

## Comps CSVの注意事項

### パス契約（重要 — 5246事故の再発防止）
- comps の唯一の入力は **`data/comps/<ticker>_comps.csv`**（または `--comps-csv PATH`）。
  **`.txt` 拡張子は読まれない**。CSVが無いと generate_dcf.py は**エラー停止**する
  （comps なしで生成したい場合のみ `--no-comps` を明示）。
- 旧 `scripts/comps_input.csv` は廃止済み（パイプラインは一度も読んでいなかった遺物。
  2026-06-04 の 5246 モデルに 4192 の5社が混入した原因）。

### Trailing Tabs/Spaces問題（自動対応済み）
- comps_fetcher.py がCSV読み込み時に各行のtrailing whitespaceを自動除去する
- generate_dcf.py実行前の手動sedクリーンアップは**不要**になった
- ただし、CSVを新規作成・編集する際は余分なwhitespaceが入らないよう注意すること

### CSVフォーマット
- UTF-8エンコーディング、カンマ区切り
- 必須カラム: Ticker, Name, Revenue, EBITDA, Operating_Income, Net_Income, Book_Value, Net_Debt
- Tickerは `.T` サフィックス付き（例: 6245.T）
- 任意カラム `Market_Cap`（JPY mn）: 指定時は yfinance を呼ばない（上場廃止・TOB銘柄や
  レート制限対策に推奨）

## 非標準決算期・上場間もない銘柄の取得（overridesキー）

- **`fiscal_year_end_month`** (1-12): 決算期末の月。EDINETの有報探索窓は (期末月+3) を
  中心に動的生成される。デフォルト4窓(3/12/6/9月期)は (期末月+3) 式で再現されるため、
  標準期の銘柄は当キー不要。**11月期(例: ELEMENTS 5246)など2月提出の銘柄は当キー必須**
  （無いと探索窓が2月をカバーせず、Step 1で長時間スピンして取得失敗する）。
- **`fundamentals_fy<YYYY>`**: 決算短信由来の最新期実績（revenue/operating_income/
  net_income/ebitda/net_assets 等）。当キーがあると generate_dcf.py が最新FYとして
  merged_data に**EDINETより優先で上書き/挿入**する（最新有報がEDINET未掲載でも完走）。
  D&A は `EBITDA − OperatingIncome` で逆算。当キーが無い銘柄は完全に従来動作（副作用なし）。
  EDINETが完全に空でも当キーがあれば非致命化して構築する。
- **`__CONFIRM__` プレースホルダ**: 値文字列に `__CONFIRM__` を含む override が残っていると
  バリデータが**エラー停止**する（2026-06-12〜）。`--allow-unconfirmed` 指定時のみ警告に
  格下げされ、当該キーはスキップ（自動値が使われる）。最終ランでは必ず確定値を入れること。
- 探索のスピード/暴走防止: `get_document_ids` は「最初に有報が見つかった提出期で打ち切り」、
  かつ `MAX_API_CALLS=400` の上限を持つ（会社の提出期は1つ。no-result時の暴走を抑制）。
- テンプレ(`dcf_comps_template.py`)はシナリオ名 `Base/Upside/Management/
  Downside 1/Downside 2` の5固定、WACCはトップレベル平坦キーを読む。独自名・ネストは
  かつて黙って無視されていたが、現在は**バリデータが実行前にエラーで検出**する
  （docs/overrides_schema.md 参照）。
- `market_analysis_template.extract_dcf_data` は **recalc済み** DCF xlsx を要求する。
  未recalc（WACC=C26 が空）の場合はエラー停止する（旧実装は黙って WACC=0 で完走していた）。
  先に `python scripts/recalc_excel_com.py <xlsx>` を実行すること。

## File Protection Rules
- `models/` directory: Auto-generated by generate_dcf.py. Safe to overwrite.
- `reports/` directory: Manually curated final models. NEVER overwrite or delete.
- generate_dcf.py will refuse to overwrite existing files unless --force is used.
- Always move finalized/integrated models from models/ to reports/ after manual edits.
- Workflow: generate_dcf.py -> models/ -> manual edits -> copy to reports/
