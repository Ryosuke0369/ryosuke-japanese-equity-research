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

## 銘柄コード入りスクリプトを作らない（2026-08-26〜）

**scripts/ と templates/ の .py ファイル名に銘柄コードを入れてはいけない。**
銘柄固有の値は必ず `data/` 配下の設定ファイルで供給する。

| やりたいこと | 正しいやり方 |
|---|---|
| DCF の前提を銘柄別に変える | `data/overrides/<ticker>_overrides.json` + `scripts/generate_dcf.py` |
| セグメントブリッジを作る | `data/segments/<ticker>_segments.json` + `scripts/add_segment_bridge.py` |
| Adjustments Log を埋める | `data/adjustments/<ticker>_adjustments.json` + `scripts/fill_adjustments_log.py` |
| 逆算DCFシート | overrides の `reverse_dcf` ブロック（テンプレが標準8枚目として生成） |
| Comps | `data/comps/<ticker>_comps.csv` |

理由: `add_segment_bridge_3110.py` / `fill_adjustments_log_5726.py` のように動くコードを
銘柄ごとにコピーすると、ロジックが分岐して片方に入れた修正がもう片方に届かず、次の銘柄で
3つ目のコピーが生まれる。実際に2026-08-26時点で2件発生していた（両方とも汎用化して削除済み）。

- 検査: `python scripts/check_script_naming.py`（違反があれば exit 1）。
  `generate_dcf.py` は起動時に同じ検査を**警告として**実行する。
- 既存の `run_*_<ticker>.py`（market_analysis / narrative のランナー）は
  grandfathered として許可リストに入っている。**このリストは閉じている**——
  追加したくなったら、それは設定ファイルにすべきものである。

## File Protection Rules
- `models/` directory: Auto-generated by generate_dcf.py. Safe to overwrite.
- `reports/` directory: Manually curated final models. NEVER overwrite or delete.
- generate_dcf.py will refuse to overwrite existing files unless --force is used.
- Always move finalized/integrated models from models/ to reports/ after manual edits.
- Workflow: generate_dcf.py -> models/ -> manual edits -> copy to reports/

## 決算先回りスクリーナー（screener/）運用コマンド — 新しいチャットはここから復元する

2026-09-13 明文化。**即興スクリプトを書く前に、ここにあるコマンドで足りないかを確認すること。**

### 置き場所（新PC: 2026-09-09 移行済み）

| もの | パス |
|---|---|
| リポジトリ | `C:\dev\ryosuke-japanese-equity-research` |
| データ正本 DATA_ROOT | `C:\screener_data`（`.env` の `DATA_ROOT` が正本。コードに直書きしない） |
| DB | `C:\screener_data\screener.db`（本体）/ `C:\screener_data\projection.db`（生成物・作り直してよい） |
| ログ | `C:\screener_data\logs\`（`run_daily_YYYYMM.log` / `screener_YYYYMM.log`） |
| 作業ログ・既知課題 | **リポジトリの** `tasks/todo.md`（追記専用フック有）と `docs/calibration_backlog.md`。`C:\screener_data` 配下には無い |
| Python | 必ず `.venv\Scripts\python.exe`。システム python には依存が入っていない |

### 1. 日次収集

- タスク: `ScreenerTdnetArchiver`（平日 19:00 / 23:15）→ `screener\run_daily.ps1 -BackfillDays 14`
  （登録: `screener\install_task.ps1`。**venv を PATH に通してから登録**。`docs/バックアップ運用手順書.md` §2 の注意と同じ）
- 手動実行:
  ```powershell
  $repo='C:\dev\ryosuke-japanese-equity-research'; $env:PATH="$repo\.venv\Scripts;$env:PATH"
  powershell -ExecutionPolicy Bypass -File $repo\screener\run_daily.ps1 -BackfillDays 14 -PythonExe "$repo\.venv\Scripts\python.exe"
  ```
- 中身（直列）: `tdnet_archiver --backfill 14` → `xbrl_parser --all`（TDnet）→
  `tdnet_titles --days 5`（非決算の開示タイトル）→ `edinet_bulk --recent 7`（EDINET 再索引+欠損補完+取得）→
  `xbrl_parser --source edinet --all --resume`（未解析の EDINET 書類だけ）→ `tdnet_archiver --report 14`
- 所要: xbrl_parser は1回ごとに全DBの coverage/unknown 集計で約10分かかる（解析0件でも）。日付ごとに呼ばないこと
- 出力: `raw\tdnet\YYYYMMDD\`、`raw\edinet\YYYY-MM-DD\`、`filings` / `financials_cum` / `fetch_runs` /
  `disclosure_titles`、ログ `logs\run_daily_YYYYMM.log`
- 成功条件: exit 0（archiver / coverage / titles / edinet がすべて 0）、最後の coverage 表に `MISSING` が無いこと
- **fetch_runs.status の意味（2026-09-13 後条件ゲート導入）**:
  `ok` = 一覧の「全N件」= 読めた行数、対象書類がすべて保存済み、対象日の翌日0時以降に取得 /
  `empty` = 確定後に0件 / `provisional` = 対象日が終わる前に取得（翌日の backfill が取り直す。当日分は欠損に数えない） /
  `incomplete` = 総件数と読めた行数が不一致 / `partial` = 対象書類の保存漏れ / `failed` = 一覧が読めない。
  **covered は ok / empty だけ**
- 部分取得の再取得: `python -m screener.fetch.tdnet_archiver --date YYYYMMDD`（冪等）
- 件数の全日突合（記録ではなく一次ソースと比べる）:
  `python -m screener.report.tdnet_completeness --from 2026-07-23 --to YYYY-MM-DD --repair --csv C:\screener_data\tdnet_completeness_YYYYMMDD.csv`
  （TDnet 一覧が 404 の日は「照合不能(保持期間外)」。2026-09-13 時点で 8/05 まで遡れた）
- 注意: `tdnet_archiver` は writer_lock を取らない。**パーサ（`xbrl_parser`）実行中に並走させると
  `database is locked` で落ちる**（2026-09-13 に踏んだ）。パーサ終了を待ってから流す。

### 2. 週次スキャン

```powershell
$py='C:\dev\ryosuke-japanese-equity-research\.venv\Scripts\python.exe'
& $py -m screener.extract.quarterly_builder --all --lock-wait 3600
& $py -m screener.extract.disclosure_flags                    # 出力は C:\screener_data\logs_disclosure.txt に Tee する
& $py -m screener.projection.materialize                      # 投影DB + 発表日カレンダー + 事後条件ゲート
& $py -m screener.report.weekly_screen --as-of YYYY-MM-DD --days 30 --verify-doc-periods --csv C:\screener_data\weekly_YYYYMMDD.csv
```
- 出力: `C:\screener_data\weekly_YYYYMMDD*.csv`（列: score / fired / doc_period / doc_date / doc_kind / stale_flag / 信頼性フラグ …）
- 成功条件: materialize が `事後条件: すべて満たす` で exit 0、weekly_screen が exit 0、
  `根拠期の照合` の **不一致 0**、`スコア計算で例外` の行が出ていないこと
- 注意: `weekly_screen` の既定フィルタは projection.db の `earnings_calendar`（システム est_date）。
  est_date は LOW が大半で、7月期本決算を10月に置く等の既知のずれがある。**決算窓は §3 のスクリプトで絞る**
- 全銘柄を採点するときは `--universe`（発表日に関係なく universe_flag=1 の全社）
- 採点方針 `--policy`: `prefer_span`（現在の既定）/ `evidence_strict`（根拠期なし・直前四半期から2期以上古い根拠・
  売上前年比2倍超/半分以下のS1S2・DSO<1日を点にしない。calibration_backlog §31。既定の切替は報告・承認後）
- 根拠の健全性監査: `python -m screener.report.evidence_audit --as-of YYYY-MM-DD --policy <policy> --out <csv>`、
  方針の前後比較: `python -m screener.report.policy_shadow_compare --before <csv> --after <csv> --out <csv>`、
  再スコアの要約と差分: `python -m screener.report.rescore_diff --new <csv> --old <csv> ...`

### 3. 決算窓フィルタ（週次CSVを発表予定日の区間で絞る後処理）

```powershell
& $py -m screener.report.earnings_window --fetch-jpx                              # JPX 決算発表予定日一覧 → raw\jpx_schedule\
& $py -m screener.report.earnings_window --fetch-jquants --as-of 2026-09-13 --from 2026-09-13 --to 2026-09-30
& $py -m screener.report.earnings_window --weekly-csv C:\screener_data\weekly_YYYYMMDD.csv `
      --as-of 2026-09-13 --from 2026-09-13 --to 2026-09-30 --mid 2026-09-18 `
      --out C:\screener_data\earnings_window_YYYYMMDD_MMDD.csv
```
- 発表日の根拠は行ごとに `basis` 列: 1=TDnet適時開示タイトル / 1b=JPX一覧(会社届出) / 2=前年同期実績+364日 / 3=システムest_date単独
- 並べ替えの第一基準は `prev_quarter_in_db`（発表される四半期の**直前四半期**が本体DBにあるか）。stale_flag は補助
- 休場日は別枠 `common/jp_calendar.py` で判定（2026-09 は 21/22/23 が休場）
- 成功条件: exit 0、ログに「窓内(未発表) N」「根拠」「直前四半期が本体にある」の3行が出ていること
- テスト: `python -m unittest screener.tests.test_earnings_window`

### 既知の落とし穴（再発防止）

- **期ラベル `FY2026-Q2` を暦日と比べない。** 暦の期末は `earnings_window.fiscal_label_ym(period, q_no, 期末月)` で写す
- TDnet の一覧は約1ヶ月しか遡れない。2026-07-23 より前の開示（1月期Q1・7月期Q3 等）は永久欠落（calibration_backlog §24）
- 取得を「ok」と記録していても、朝に走った回は当日分の一部しか無いことがある（2026-09-13 に 5日ぶん発見）
- `disclosure_titles.code` が NULL だった（2026-09-13 修正。既存行は `tdnet_titles --rederive-codes` で再導出、
  会社名まで一致しない行は埋めない）
- 決算期末月は**期間11ヶ月以上のタイトル**からだけ取る。半期報告書のタイトルは半期の期間しか書かない会社がある
  （4396/4495 で6月期→12月期と誤判定していた。2026-09-13 修正）
- 別枠の位置ベース比較（根拠期なし）は `prefer_span` では点になる。偽陽性の主因（3475 グッドコムアセット）
- `--verify-doc-periods` は TDnet リンク（`.../inbs/<doc_id>.pdf`）も解決する（2026-09-13 修正。以前は全件「不一致」に見えた）。
  全銘柄スキャンで不一致が大量に出たら、まずリンク形式の読み取りを疑う
