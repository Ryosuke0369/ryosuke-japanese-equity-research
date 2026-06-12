# Overrides 契約仕様書（data/overrides/<ticker>_overrides.json）

**バージョン: 2026-06-12**（`scripts/overrides_validator.py` がこの契約を実行前に強制する）

## 大原則

1. **平坦キーのみ。** WACC・DCF前提はすべてトップレベルの平坦キー。ネスト（`wacc_inputs` 等）は**エラーで停止**する。
2. **未知キーはエラー。** 黙って無視されることは構造的に起きない。コメント・メモは `_` 接頭辞キーに書く（`_note`, `_wacc_note` 等は常に許可）。
3. **エラーなく完走した実行は、全 override キーが消費されたことを意味する。**
4. `__CONFIRM__` を含む値が残っていると**エラーで停止**。`--allow-unconfirmed` で明示した場合のみ警告に格下げ（自動値で続行）。最終ランでは使わないこと。

## トップレベルキー一覧

### 会社・メタ
| キー | 型 | 説明 |
|---|---|---|
| `ticker` | str/int | 4桁コード可（`"4192"`）。パイプラインが `.T` を自動付与 |
| `company_name` | str | モデル表記名（未指定時は EDINET名） |
| `exchange` / `sector` | str | 表示用 |
| `fiscal_year_end_month` | int 1-12 | **非標準決算期は必須**（11月期=ELEMENTS等）。EDINET探索窓を (期末月+3) 中心に生成 |

### 市場データ
| キー | 型 | 説明 |
|---|---|---|
| `current_price` | number | yfinance より優先。`__CONFIRM__` 残存はエラー |
| `shares_outstanding` | number | 同上（発行済株式総数 実数） |
| `shares` | dict | `{"fully_diluted_shares": N}`。DCF/SOTP 共通の single source of truth |
| `net_debt` | number | JPY mn。負 = ネットキャッシュ |
| `beta` | number | テンプレで [0.6, 1.75] にクランプ（範囲外→1.0） |
| `de_ratio` | number | 未指定時は net_debt/時価総額 から自動算出 |

### WACC（平坦キー！）
`risk_free` / `erp` / `size_premium` / `cost_of_debt_at`（**税引後**）/ `tax_rate`
※ `size_premium` を指定しない場合は時価総額から自動決定（<1000億円→3.0%等）。

### ターミナル・Exit
| キー | 説明 |
|---|---|
| `terminal_growth` | PGM 永久成長率 |
| `exit_multiple` | EV/EBITDA Exit 倍率（フォールバック用に常に残す） |
| `exit_sales_multiple` | EV/Sales Exit 倍率。`primary_multiple: "EV/Sales"` と併用で Exit が EV/Sales 経路になる |
| `primary_multiple` | `"EV/EBITDA"` or `"EV/Sales"` |

### Capex / D&A
`capex_method` / `da_method`（`"revenue_pct"` デフォルト or `"direct"`）、`capex_pct` / `da_pct`（direct でもフォールバック用に残す）、`capex_direct` / `da_direct`（`{"historical": [...], "projections": [...]}`）

### NWC
| キー | 説明 |
|---|---|
| `nwc_method` | `"days"`（デフォルト）/ `"revenue_pct"` / `"itemized"` のみ |
| `nwc_items` | itemized 用。各要素は `label` / `base_value` / `scenario_key` / `side` / `denom` 必須。`scenario_key` で宣言した名前は scenarios 内の有効キーになる |
| `base_year_ar` / `base_year_inv` / `base_year_ap` / `base_year_nwc` / `base_year_trade_receivables` / `base_year_trade_payables` | 基準年BS値の上書き |

### 予測フレーム・実績
`projection_years`（=シナリオ配列長）/ `projection_start_fy` / `stub_fraction` / `stub_months_elapsed` / `ltm_revenue` / `base_year_revenue` / `base_year_cogs` / `core_ebitda` / `core_net_income`（null 指定で自動再計算にフォールバック）

`hist_years` / `hist_revenue` / `hist_operating_income` / `hist_net_income` / `hist_cogs` / `hist_sga` / `hist_ocf` / `hist_capex` / `hist_cash` / `hist_debt` / `hist_depreciation` / `hist_nwc_pct`（すべて oldest-first 配列）
※ `hist_da` ではなく **`hist_depreciation`**。`hist_ordinary_income` は消費されない（書くなら `_` 接頭辞）。

### scenarios（固定5名のみ）
```json
"scenarios": {
  "Base":       { "revenue_growth": [..5], "cogs_pct": [..5], "sga_pct": [..5],
                  "dso_days": [..5], "dih_days": [..5], "dpo_days": [..5] },
  "Upside":     { ... },
  "Management": { ... },
  "Downside 1": { ... },
  "Downside 2": { ... }
}
```
- シナリオ名は **Base / Upside / Management / Downside 1 / Downside 2** の5固定。Bull/Bear 等の独自名は**エラー**。
- 配列キー: `revenue_growth` / `cogs_pct` / `sga_pct` / `dso_days` / `dih_days` / `dpo_days` / `nwc_pct`（+ `nwc_items` で宣言した `scenario_key`）。各配列長 = `projection_years`。
- 一部シナリオのみの定義も可（残りは自動生成値とマージ）。

### fundamentals_fy<YYYY>（決算短信由来の最新期実績）
`revenue` 必須。`operating_income` / `net_income` / `ebitda` / `net_assets` / `total_assets` / `total_liabilities` 等。EDINET 未掲載の最新期を**EDINETより優先で**注入する。D&A は `EBITDA − OI` で逆算。

### segments / sotp / cost_structure
従来どおり（CLAUDE.md の Segment / Driver Analysis 節、SOTP仕様参照）。`segments` 定義時は `cogs_pct` がセグメントEBITからの逆算に切り替わる。

## 書いてはいけないもの（バリデータが名指しでエラーにする）
- `wacc_inputs` / `wacc` / `dcf_assumptions` … ネスト構造は読まれない → 平坦キーで
- **トップレベル `comps`** … Step 6 で必ず `data/comps/<ticker>_comps.csv` から上書きされるため無意味。Comps データは CSV へ
- `risk_free_rate`→`risk_free`、`cost_of_debt`→`cost_of_debt_at`、`hist_da`→`hist_depreciation` 等の揺れ（サジェスト付きエラー）

## Comps CSV 契約（data/comps/<ticker>_comps.csv）
- **パスと拡張子が契約**: `data/comps/<ticker>_comps.csv`。`.txt` は読まれない（5246事故の原因）。CSV が無いと**エラー停止**（`--no-comps` 明示時のみ comps なし生成可）
- UTF-8 カンマ区切り推奨（UTF-16/タブ区切りも自動吸収はされる）
- 必須列: `Ticker`（`.T` 付き）, `Name`, `Revenue`, `EBITDA`, `Operating_Income`, `Net_Income`, `Book_Value`, `Net_Debt`
- 旧 `scripts/comps_input.csv` は**廃止済み**（パイプラインは一度も読んでいなかった）

### Market_Cap 列（任意・JPY mn）— 2026-06-12 仕様
- **列がある場合**: 各社の時価総額（JPY mn）として**その値をそのまま使用**し、yfinance は呼ばない。
  生成時に `[Comps] Market cap source: CSV (static)` がエコーされる。EV・PBR はこの値から導出。
- **全行ルール**: 列が存在する場合は**全行記入が必須**。1行でも空欄があると `ValueError` で
  **エラー停止**する（静的/ライブの黙った混在を許容しない）。許されるのは「全行記入」か
  「列ごと省略」のどちらかのみ。
- **列が無い場合**: yfinance ライブ取得にフォールバックするが、必ず
  `[Comps] WARNING: Market cap source: yfinance live — 出力は実行時点で変動する` が
  エコーされる。**再現性が必要な最終ランでは Market_Cap 列を記入すること**
  （ライブ取得だと implied 値が実行時点の株価で±変動し、同一CSVでも結果が再現しない）。
- **as-of 運用**: Market_Cap を記入・更新したら、下表の as-of を更新する。値は原則
  yfinance の当日終値ベース。上場廃止・TOB銘柄（yfinance 404）は最終観測値を据え置き、
  その旨を Note に書く。

| CSV | Market_Cap as-of | Note |
|---|---|---|
| `5246_comps.csv` | 2026-06-12 終値 | 全5社 yfinance 取得 |
| `4192_comps.csv` | 2026-06-12 終値 | LightWorks 4267 のみ上場廃止（yfinance 404）のため TOB 価格固定の継承値 10,730 を据え置き |

## market_analysis ランナー設定の補足

- `price_targets`（任意・Block D「利確/損切り価格の倍率翻訳」）: 指定すると Implied Multiple
  Analysis シートに価格水準ごとの倍率翻訳が出力される。キー無しなら Block D はスキップ。
  **実際の運用値（個人の売買水準）はコミット対象に書かないこと**。設定例（架空値）:

  ```python
  "price_targets": {"損切り": 100, "現在": 200, "第1利確": 300},
  ```

## 実行フロー上の注意
- 生成完了時に「Effective WACC inputs」と Comps 5社名がコンソールに出る。**必ず目視照合する**こと
- `market_analysis_template` は **recalc 済み**の DCF xlsx を要求する。未 recalc（WACC セルが空）はエラー停止 → 先に `python scripts/recalc_excel_com.py <xlsx>`
- バリデーション違反は全件まとめて報告される（1件ずつ直す必要はない）
