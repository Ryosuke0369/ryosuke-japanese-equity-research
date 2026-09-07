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
| `company_type` | str | `A`/`B`/`C`/`D`/`E`（手順書 §2 の銘柄型）。宣言は任意だが**型D は必ず明示すること** — 下記 |

### 市場データ
| キー | 型 | 説明 |
|---|---|---|
| `current_price` | number | yfinance より優先。`__CONFIRM__` 残存はエラー |
| `shares_outstanding` | number | 同上（発行済株式総数 実数） |
| `shares` | dict | `{"fully_diluted_shares": N}`。DCF/SOTP 共通の single source of truth |
| `net_debt` | number | JPY mn。負 = ネットキャッシュ |
| `beta` | number | **実測（回帰）β = raw を入れる**。調整後の値ではない。テンプレが Blume 調整 `0.67×raw + 0.33×1.00` を掛け、`[0.3, 2.0]` にクランプする（下記） |
| `de_ratio` | number | 未指定時は net_debt/時価総額 から自動算出（**市場データ override 適用後**の時価総額を使う） |

#### `company_type` と裁定の適用範囲（2026-09-06 追補12 §A-3）

- 値は手順書 §2 の銘柄型: `A` 通常の事業会社 / `B` シクリカル /
  `C` captive finance 持ち製造業 / `D` 銀行 / `E` 銀行を連結に持つ持株会社。
  それ以外の値はバリデータがエラー停止する。
- **用途**: `generate_dcf.py` は生成の最終段（Step 8.5、validate の直前）で
  追補6 §X の脚裁定を自動適用する。**型D は DCF が成立せず Target は DDM+RI が主手法**
  であるため、裁定の対象外としてスキップし、**スキップ理由をコンソールに出す**。
- **未宣言のときは DCF 型（A/B/C/E）とみなして裁定を適用する**。銀行を型判定なしで
  通すと DCF 2脚の裁定が走ってしまうため、**型D は必ず `company_type: "D"` を書くこと**。
- **型を推測しない理由**: 8410 の overrides は `net_debt=0` /
  `base_year_ar=inv=ap=0` / `nwc_method="revenue_pct"` という形をしており、
  一見「銀行の指紋」に見える。しかしこれは**銀行だからではなく ΔNWC を強制ゼロにする
  実装手段**であり、非銀行でも同じ形を取りうる。構造から型を当てるのは
  サイレントな推測にあたるので、明示キーで受け取る。

#### β の扱い（2026-09-06 フェーズ2 #6 で改訂）

- **`beta` に入れるのは実測（回帰）βの生値**。yfinance 由来でも自前回帰でも同じ扱いで、
  overrides 経由でも yfinance 経由でも**同一の処理を通る**。
- テンプレが **Blume 調整を既定**で適用する: `β_adj = 0.67 × β_raw + 0.33 × 1.00`。
  回帰βは平均回帰するため。2026-09-05 バッチでは旧ルール `[0.6, 1.75]` の
  **下限 0.60 に 85件中 57件（67%）が張り付き**、上側では太陽誘電の raw 1.561 が
  WACC 13.35% を生んでいた。
- クランプ域は **`[0.3, 2.0]`**。範囲外は**無言で置換せず WARN を出す**。
  Blume 調整後にこの域を外れるには raw が概ね `[-0.05, 2.49]` の外である必要があり、
  クランプは日常経路ではなく稀なガードになる。
- `beta` を渡さない（yfinance も取れない）場合は市場β 1.00 を採用し、WARN を出す。
- Adjustments Log（`DCF Model!C8` 行）と Pipeline Metadata
  （`beta_raw` / `beta_blume_adjusted` / `beta_adopted_c8` / `beta_basis` / `beta_clamped`）に
  **raw / adjusted / 採用値の3点**が必ず残る。
- **移行時の注意**: 旧ルール下で作られた overrides には**クランプ後の値**（0.60 等）が
  書かれていることがある。それを raw として渡すと二重に縮小される。
  既存銘柄の `beta` は raw を再導出して差し替えること。

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
※ `hist_da` ではなく **`hist_depreciation`**（キー名は既存契約のまま。バリデータが `hist_da` をサジェスト付きエラーにする）。`hist_ordinary_income` は消費されない（書くなら `_` 接頭辞）。

**配列長は契約**（2026-07-31〜）: `hist_years` を指定する場合、上記 `hist_*` 配列は
**すべて `hist_years` と同じ長さ**でなければならない。不一致は**エラー停止**する。
（長さ違いの配列が位置コピーされ、ある年度のCFが別年度の列に入る事故の再発防止。）

#### OCF / 現金 / 有利子負債の年度突合（2026-07-31〜）

`hist_ocf` / `hist_cash` / `hist_debt` / `hist_capex` / `hist_depreciation` の値は
**位置ではなく会計年度キーで** Financial Statements の各列に割り当てられる。

優先順位: **overrides 指定 > EDINET の年度キー一致 > 空欄**

EDINET 側キー（`FY2024` 等）と `hist_years`（`FY2024/3` 等）の突合ラダー:

1. ラベル完全一致
2. 年 + 決算月が一致（`FY2024/3` ↔ `FY2024/3`）
3. 年が一致し、その年が**両側で一意**（`FY2024/3` ↔ `FY2024`）

いずれにも当たらない年は**空欄**にする（詰めない・ずらさない）。生成ログ最終行に
`WARNING: OCF/Cash/Debt coverage OCF a/m, Cash b/m, Debt c/m` を出力し、
Financial Statements シートにも注記セルを置く。埋めたい場合は overrides に
`hist_ocf` / `hist_cash` / `hist_debt` を明示すること。

#### `ltm_revenue`（C20 の上書き）

未指定なら `LTM = 直近本決算FY実績 − 前年同期累計Q + 当期累計Q` で自動構築し、
**3成分と結果をコンソールに必ずログ出力**する（`[LTM]` 行）。スコープ調整が必要な銘柄
（7203=自動車事業のみ、8267=営業収益合計 等）では本キーで上書きする。上書き時は
C20 のラベルに `(override)` が付き、自動構築値との乖離が **20%超なら警告**（停止はしない）。

#### `book_value`（新規, JPY mn）

Comps Analysis の自社行 Book Value 列（P列）に使う純資産。未指定時は comps CSV の
自社行 `Book_Value` を使う。PBR / ROE はこの列を参照する数式（`=D/P` / `=I/P`）。

### investment_thesis / key_risks のトークン（2026-07-31〜）

本文中の以下のトークンは、テンプレートが `="…"&TEXT(<セル参照>,"<書式>")&"…"` の
連結数式に変換して書き込む（株価・株数を更新すると本文の数値も追随する）。
**トークンを含まない行は従来どおり素のテキスト**（後方互換）。

| トークン | 参照先 | TEXT書式 |
|---|---|---|
| `{price}` | Executive Summary!C9（現在株価） | `#,##0` |
| `{target_price}` | Executive Summary!C10 | `#,##0` |
| `{upside_pct}` | Executive Summary!C12 | `+0.0%;-0.0%` |
| `{pb}` | Comps Analysis 自社行 PBR | `0.00"x"` |
| `{per}` | Comps Analysis 自社行 PER | `0.0"x"` |
| `{wacc}` | DCF Model!C26 | `0.00%` |

- 上表以外の `{...}` は**バリデータがエラー停止**（タイプミス検出）。
- 参照先セルが存在しない場合（comps 無しで `{pb}` 等）は警告のうえ素のテキストに降格。
- 1セル 8,192 字を超える場合は文単位で複数セルに分割出力される。

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

### 実績ベース負債コスト（オプション・2026-08-26〜）

`interest_expense` を与えると、C11（税引後負債コスト）が**推定値から実績値に切り替わる**。

| キー | 型 | 意味 |
|---|---|---|
| `interest_expense` | number | 支払利息（JPY mn） |
| `loan_fees` | number | ローン関連手数料（任意、利息に加算） |
| `debt_beginning` / `debt_ending` | number | 有利子負債の期首・期末。**省略時は `hist_debt` の直近2年の平均**を使う |

`実績Kd(税引前) = (支払利息 + 手数料) ÷ 平均有利子負債`、`C11 = 実績Kd × (1 − tax_rate)`（小数4桁に丸め、C5/C18 と同じ規約）。
`interest_expense` を与えたのに債務残高が特定できない場合は**エラー停止**する（黙って既定値に落ちない）。

導出内容は Adjustments Log の自動行と Pipeline Metadata（`cost_of_debt_basis` /
`cost_of_debt_interest` / `cost_of_debt_avg_debt` / `cost_of_debt_pretax`）に記録され、
**マージナルコスト注記**（実績は既存借入の平均コストであり新規調達の限界コストではない）が併記される。
validate_output のチェック17 が C11 と記録済み導出の整合を機械検証する。

### fx_sensitivity（為替感応度 Table 3・オプション・2026-08-26〜）

輸出型銘柄のみ。`enabled: true` で Sensitivity シートに Table 3 を生成する（既定は非生成）。

| キー | 型 | 意味 |
|---|---|---|
| `enabled` | bool | 輸出型フラグ。false / 未指定なら Table 3 は作られない |
| `assumption_rate` | number | 会社の為替前提（例 150）。**必須**・0より大 |
| `usd_revenue_ratio` | number | 売上のUSD連動比率（小数 0〜1）。**必須** |
| `usd_cogs_ratio` | number | 売上原価のUSD連動比率（小数 0〜1）。**必須** |
| `currency_pair` | str | 既定 `"USD/JPY"` |
| `assumption_source` | str | 前提の出所（例 `"H2 guidance"`）。ラベルに付記される |
| `offsets` | list | レート格子（既定 `[-20,-10,-5,0,5,10,20]`、前提レートからの円差） |
| `estimated` | bool | 既定 true。比率が推定であることを表題・ラベルに明示し、Adjustments Log に「推定・要確認」で記録する |
| `note` | str | 表の下の注記 |

`感応度 = (売上 × 売上連動比率 − COGS × 原価連動比率) ÷ 前提レート`。原価側の連動が円高メリットを
一部相殺する構造を織り込む。売上・COGS・営業利益は `DCF Model` の Year 1 列への生き参照なので、
**シナリオ切替に追随**する。比率と前提レートは青字入力セル。
validate_output のチェック18 が「前提レートの列が DCF Model の Year 1 営業利益を再現するか」を検証する。

### normalized_net_income（Comps 正常化純利益の参考行・2026-08-26〜）

特損等で分母が歪んだ PER の隣に、正常化後の参考行を置く（標準メモ §1 / §2「残置＋除外＋理由記録」）。
`{"pretax": 3814, "addbacks": 2619, "label": "...", "note": "..."}` で
`=ROUND((pretax + addbacks) × (1 − 'DCF Model'!C6), 0)` を書く（税率セルへの生き参照）。
`{"value": 4465}` または素の数値でも可。**Target には入らない**（行ラベルに `[参考・Target不算入]`）。

### reverse_dcf（逆算DCFシート・2026-08-26〜）
`Reverse DCF` は**標準8シートの4枚目**として毎回生成される。全キー任意 — 省略すると
`hist_operating_income` / `hist_revenue` / `hist_years` から自動導出される。

| キー | 型 | 意味 |
|---|---|---|
| `enabled` | bool | `false` でシート生成をスキップ（既定 true） |
| `op0` / `op0_label` | number / str | ランプ開始の営業利益と年度ラベル（既定: 実績最終年） |
| `peak_op` / `peak_label` | number / str | サイクルピーク営業利益と年度ラベル（既定: 実績の最大値の年） |
| `peak_opm` | number | ピーク営業利益率（小数。既定: ピーク年の OP÷売上） |
| `opm_grid` | list | Block C の定常OPM列。`null` を混ぜるとその行が**ピークOPMセルへの生き参照**になる（既定: ピークOPM×0.5/0.75/live/1.25/1.5） |
| `n_years` | list[int] | Block B の到達年数（既定 `[3,5,7,10]`。2番目が Block C / Block F の見出し値） |
| `benchmark_ticker` | str | Block E（取引ベンチマーク）に使う **comps CSV 上のティッカー**。表に無ければ Block E を出さずに警告 |
| `deal_note` | list[str] | Block E に添える取引条件の注記 |

**自動導出できないときはシートを作らない**（実績営業利益が無い／ピークが赤字／ピークOPMが
負）。理由は `Adjustments Log` の Pipeline Metadata `reverse_dcf_sheet` に残り、
validate_output のチェック16が WARN で報告する。ゼロ埋めのシートは作らない。

既存ブック（この変更以前に生成したもの）への後付けは
`python scripts/add_reverse_dcf_sheet.py <xlsx> [--benchmark-ticker ...]` →
`recalc_excel_com.py` → `validate_output.py`。実装は同じモジュールを呼ぶ。

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
- **`EBITDA` = 営業利益 + 減価償却費**。D&A が取れない社は **EBITDA を空欄で提出**すること
  （営業利益をそのまま入れない）。空欄／営業利益と完全一致の行は自動検知され、EBITDA セルを
  空欄化 + `(D&A n/a)` 注記のうえ **EV/EBITDA 統計から除外**される。有効 peer が 3社未満に
  なると implied は `INVALID (n<3)` となり Target Price の平均から外れる（2026-07-31〜）。
- **`Book_Value` は必須**。欠損行は PBR / ROE が空欄になり警告が出る。
- **自社行**（CSV 1行目に置く運用）は統計から**自動除外**される（ティッカー突合。
  `add_bank_valuation.py` 等の生成後パッチに頼らない）。自社行の時価総額・EV は
  `'Executive Summary'!C9 × 株数` の数式でリンクされる。
- peer 行の時価総額は CSV の生成時点値のまま。シート上部に `Peer prices as of <生成日>` を自動注記。
- 生成時に各 peer の直近株価日付を yfinance で照会し、**45日超古い / 取得不能なら統計から除外**
  （行は Note 付きで残る）。オフライン等で過半が取得失敗した場合は環境要因とみなし除外しない。
  `--no-peer-check` でスキップ可。
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
| `285A_comps.csv` | 2026-07-29終値ベース（SNDK/MU）、SK Hynix FY2025実績ベース | 外貨comps（USD/KRW）のためレート制限回避と再現性を優先し静的値を採用。FX: USD/JPY 163.46、KRW/JPY 0.11339（2026-07-29/30時点、xe.com）。Samsung Electronics（005930.KS）はメモリ単体の完全財務諸表が非開示のため comps から除外（DS部門営業利益のみ判明、EBITDA/NI/BV不可）。SanDisk はTTM値（2026年4月期まで）を使用、FY2025単独では赤字のため。 |
| `3687_comps.csv` | 2026-07-30時点(stockanalysis.com) | 全社日本企業のためFX換算不要。**プロンプトのティッカー誤りを2件発見・訂正**: (1)「ベース株式会社」は`4483.T`ではなく`4481.T`が正しい(4483.Tは無関係のJMDC)。(2)`4726.T`(SBテクノロジー)は2024-09-06にソフトバンクのTOBで上場廃止済み(時価総額取得不可)のため、プロンプト指定の代替候補`2158.T`(FRONTEO)に差し替え。EBITDA=営業利益+減価償却費(CF計算書)で全社統一、D&A取得不能で除外した社はなし。統計からフィックスターズ自身を除外(`_comps_stats_helper`スクリプトなし、`overrides_validator`側でなく市場分析側での別処理が必要な場合は要確認)。 |
| `8267_comps.csv` | 2026-07-29終値ベース(3382/9843)、7532は日付未確認のIR公表値 | 銀行(イオン銀行)連結子会社があるため、自社行のNet_Debtは有利子負債(預金除く)−現金+非支配株主持分(984,094)で算出(schema上部「net_debtにMIを織り込む」設計、DCF側overridesのnet_debtと整合)。プロンプト指定peer5社のうち`3141.T`(ウエルシアHD)は2025-11-27にツルハHDへ吸収合併され上場廃止、`8905.T`(イオンモール)は2025-06-27に株式交換で完全子会社化され上場廃止と判明(いずれも時価総額取得不可)のため**両方除外**、3社(3382/7532/9843)のみで統計を構成(元々イオン自身は統計除外の設計だったため、実質的な変更は「表示専用2社」が無くなった点のみ)。`scripts/add_sotp_crosscheck.py`が生成後にComps Analysisの統計式(row5除外)をパッチ。 |
| `8410_comps.csv` | 2026-07-29終値ベース(Rakuten Bank/Japan Post Bank)、AEON Financial Serviceは日付未確認のIR公表値 | 銀行のため全行Net_Debt=0固定(schema上部「銀行のNet_Debt」注記参照)。プロンプト指定のpeer4社のうち`7163.T`(住信SBIネット銀行)は2025-09-25にNTT Docomoの TOB で東証上場廃止済みと判明(時価総額なし、FY2026/3科目も未確認)のため**除外**、3社(5838/7182/8570)で統計を構成。統計からセブン銀行自身を除外する要件があるが`dcf_comps_template.py`の統計式(`{col}5:{col}{last_row}`)には自社除外の仕組みがなく、`scripts/add_bank_valuation.py`が生成後にopenpyxlでComps Analysisの統計式レンジ(row5除外)を書き換えるポストプロセスとして対応。AEON Financial Serviceは決算期が2月期(3月期ではない)のためFY2026/2実績を使用(約1か月のズレ、僅少)。Japan Post Bankの純資産9,260,000は非支配株主持分等未調整の総額(自己資本の厳密値ではない)。 |
| `7203_comps.csv` | 2026-07-29/30終値ベース（Market_Cap）、各社直近期末実績（P/L・BS） | 連結ベースで統一（DCFは自動車事業のみ、両者の乖離は既知の設計で最終レポートに明記）。FX: USD/JPY 158.75、EUR/JPY 183.44（2026-03-31時点、valutafx.com。GM/F/VOW3のBS・PL全項目にこの期末レート1本を簡便法として統一適用、プロンプト許容範囲）。`comps_fetcher.py`に`_`接頭辞コメント行のスキップ機能が無いため、CSV内へのFXコメント埋め込みは行わずこの表に記録する運用とした。Honda/SubaruはFY2026/3に関税等一過性費用で赤字/大幅減益、Suzuki実績は情報源により営業利益に約3%の差異（602.9 vs 604.6十億円、大きい方を採用）。VW時価総額は普通株(VOW)/優先株(VOW3)二種類の合算方法が情報源で一致せず、37-38EURbnの中間値37,780EURmnを近似値として採用。 |
| `1433_comps.csv` | 2026-09-02 終値ベース（Market_Cap）、財務は各社直近本決算実績 | 全社日本企業のためFX換算不要。EBITDA = 営業利益 + D&A（yfinance Depreciation And Amortization、無形・のれん償却込み）で全社統一。自社行も同定義（741.1 + 減価償却29.4 + のれん償却45.5 = 816.0）。Net_Debt = Total Debt − 現金及び現金同等物 + 非支配株主持分で統一、Book_Value は Stockholders Equity（非支配株主持分を除く）。ピア5社（6379/1968/1945/1966/1716）は 2026-09-02 の終値を取得でき上場廃止・TOB進行中の社は無い。**1716 第一カッター興業のみ財務が FY2025/6**（yfinance が FY2026/6 実績を未反映、約14ヶ月古い）— 時価総額だけ最新のため倍率が『新しい株価 ÷ 古い利益』になっている点に注意。**604A ビーエイブル（2026-07-29上場、廃炉・原子力プラント工事）はCSVに含めていない**: プロンプトは『残置＋統計除外（東邦方式）』を指定したが、現行テンプレで peer を統計から外せる経路は（a）株価鮮度チェック（45日超・取得不能）と（b）D&A欠損（EBITDA空欄）のみで、株価も D&A も正常に取れる live 銘柄には適用できないため、統計汚染を避けてCSVから外し数値は Adjustments Log と Reverse DCF Block E 注記に記録した（同社 2026-09-02 実勢: 時価総額11,225 / PER 23.6x / EV/EBITDA 13.6x）。**改善候補**: CSV に任意列 `Exclude_From_Stats` / `Exclude_Reason` を設け comps_fetcher.py が `comp['exclude_from_stats']` にセットすれば既存の stale 経路（全統計から除外＋行は残置＋Note）を live 銘柄にも再利用できる（列が無ければ現行と同一挙動）。 |

## market_analysis ランナー設定の補足

- `price_targets`（任意・Block D「利確/損切り価格の倍率翻訳」）: 指定すると Implied Multiple
  Analysis シートに価格水準ごとの倍率翻訳が出力される。キー無しなら Block D はスキップ。
  **実際の運用値（個人の売買水準）はコミット対象に書かないこと**。設定例（架空値）:

  ```python
  "price_targets": {"損切り": 100, "現在": 200, "第1利確": 300},
  ```

## `sotp` ブロック（generate_sotp.py / sotp_template.py）

### 単位契約（最重要 — 桁事故の温床）
| キー | 単位 | 備考 |
|---|---|---|
| `sotp.consolidated.shares_outstanding` | **千株** | `dcf_comps_template` の `shares_outstanding` は【株】。**1,000倍違う**。generate_sotp.py が `overrides.shares.fully_diluted_shares ÷ 1000` を自動注入する |
| `sotp.consolidated.da_total` / `net_debt` / `minority_interest` | ¥百万 (¥M) | |
| セグメント OP | ¥百万 (¥M) | |

- Fair Value = `Equity Value(¥M) × 1,000 ÷ Shares(千株) ÷ split_ratio`。
- **ガード（2026-07-31〜）**: `shares_outstanding > 1e8` なら「株単位で渡している疑い」として
  **エラー停止**する（従来は静かに 1/1000 の株価が出ていた）。

### 必須キー（2026-07-31〜 デフォルト廃止）
- `sotp.sensitivity.primary_segment_key`: **必須**。実在するセグメント key でなければエラー停止。
  （旧デフォルト `"aero"` は IHI 専用で、他社では無言でラベル空・別セグメント基準の表になっていた）
- `sotp.sensitivity.table2` を書く場合、`row_segment_key` / `col_segment_key` も**必須**
  （旧デフォルト `"industrial"` / `"energy"` を廃止）。

### 任意キー
- `sotp.da_allocation_intro`（文字列リスト）: D&A Allocation シート冒頭の方法論説明。
  未指定時は汎用文（「セグメント別D&Aは非開示のため固定資産集約度で按分」）。
- `sotp.da_allocation_notes`（文字列リスト）: セグメント別の按分根拠。未指定時は
  「合計100%・残差は全社/消去」の共通注記のみ。
  （旧実装は IHI の Aero/Niigata 等の注記を全銘柄に固定出力していた）

### DCF クロスチェックの読み取り
Executive Summary の**列Bラベル**（`Perpetuity` / `Exit` / `EV/EBITDA` / `PER`）で行を特定する。
行16-19 の位置決めは廃止（銀行型など Valuation Summary が再構成されたモデルで別手法の値を
誤ったラベルで取り込むため）。読み取れたラベルは Cover シートにそのまま表示する。
クロスチェック値は生成時スナップショットのため、表の下に取得日時と元ファイル名を自動注記。

## market_analysis の入力契約（補足・2026-07-31〜）

- `dcf_excel_path` は**必須**（config だけのフォールバック経路は存在しない。旧実装は
  「fallback を使う」と印字してから NotImplementedError で落ちていた）。
- 逆算Comps の分布は **peer のみ**（自社は統計対象外、`（参考・統計対象外）` 行に別掲）。
  peer 数は可変なので見出しは `{n}社` と動的表示。**p25/中央値/p75 は peer 行から
  Excel PERCENTILE.INC 互換で再計算**する（古い DCF の統計セルは自社込みのため）。
  再計算値と DCF シートの統計セルが食い違う場合はコンソールに NOTE を出す
  → **DCF を現行テンプレで再生成すれば一致する**。
- `price_points` を指定する場合、**先頭要素の label は `Current Price`** にすること
  （Block 4/5 が先頭行の alpha を市場 alpha として参照する。違う場合は警告）。
- Block 4 の alpha（1.00/1.50/0.85/0.40/0.15）は DCF シナリオの実データではなく
  **Base 成長にかける倍率プロキシ**（列見出しは `alpha (proxy)`）。
- Narrative Stage シートの値は生成時の静的値。**Excel 上で編集しても Stage/Verdict は
  動かない**（スコア変更は config の `narrative` を更新して再生成）。

## 生成物の自動検証（2026-07-31〜）

- `scripts/generate_dcf.py` は生成後に **Excel COM で recalc → `scripts/validate_output.py`** を
  自動実行する。FAIL があれば **exit 1**（生成物ファイルは削除せず残す）。
  `--no-recalc` / `--no-validate` で個別にスキップ可。
- 単体実行: `python scripts/validate_output.py <xlsx>`（FAIL で exit 1、
  `<xlsx>_validation.txt` にレポート出力）。判定は FAIL / WARN / SKIP / PASS。
- **対象ファイル種別はシート名で自動判別**する:
  - DCF（`DCF Model`）: チェック 1-18
    （17: C11 が記録済みの負債コスト導出と整合、
    18: 為替感応度 Table 3 の前提レート列が Year 1 営業利益を再現）
    （14: Target Mid が C16:C17 のDCF2法のみで Comps 行を参照していない、
    15: Exit 法にも EV<ネットデットの INVALID ガードがある、
    16: `Reverse DCF` シートが `DCF Model` の直後に存在し生きた数式である）
  - market_analysis（`Implied Growth Analysis`）: 1, 14（Block3 の IFERROR）,
    15（逆算Comps に自社が混入していない）, 16（implied price が昇順）, 20（「N社」表記の整合）
  - SOTP（`SOTP Valuation`）: 1, 17（D&A 按分 Check = OK）, 18（Cover の SOTP 行が
    `'SOTP Valuation'` 参照の数式）, 19（Fair Value が 10〜10^6 の範囲＝単位事故検出）
- 検証の根拠は xlsx 内の **`Adjustments Log` シート**下部 `Pipeline Metadata` ブロック
  （生成時の LTM 3成分・C5/C18 の算出根拠・年度突合の結果・自社行/peer 行番号等）。
  **このブロックは手で編集しないこと**（手修正の記録は同シート上部の表に書く）。

## 実行フロー上の注意
- 生成完了時に「Effective WACC inputs」と Comps 5社名がコンソールに出る。**必ず目視照合する**こと
- `market_analysis_template` は **recalc 済み**の DCF xlsx を要求する。未 recalc（WACC セルが空）はエラー停止 → 先に `python scripts/recalc_excel_com.py <xlsx>`
- バリデーション違反は全件まとめて報告される（1件ずつ直す必要はない）

---

## 型E の節（2026-09-07 追加）

型E は**銀行・消費者金融を連結に持つ事業会社**である（手順書 §2）。型D（銀行そのもの）
と違い、事業の主体は非金融の側にあるが、連結 BS には預金・貸出金が独立科目で載るため、
**連結のまま DCF を組むと預金を有利子負債として割り引く**ことになる。

型判定は `python batch/fin_business_screen.py <docID>` が推奨を出す。境界は
**金融「資産」の総資産比 50%**（資産側だけで測る。預金は負債であり、貸出金と足すと
同じ事業を二度数える）。50% 以上なら型D、未満なら型E。25〜50% は
セグメント利益構成で人が確定する帯である。

### `company_type: "E"` を宣言すると実行前に必須化される4項目

| キー | なぜ必須か |
|---|---|
| `net_debt` | 連結 BS からの自動抽出は銀行の預金を有利子負債に、貸出金を資産に含める。**非金融ベース**（銀行預金・貸出金・コールローン/コールマネーを除外し、**非支配株主持分を加算**）で一次資料から作る |
| `de_ratio` | **自動計算は禁止**。自動計算は `net_debt ÷ 時価総額` なので、金融込みの連結 net_debt を使うと WACC の資本構成が金融の資金調達で膨らむ |
| 非金融 P/L | `segments`（セグメント分解が開示されている場合）**または** `hist_revenue` + `hist_operating_income` + `base_year_revenue` を非金融ベースで供給。金融が独立の報告セグメントでない銘柄が多い（9433 の金融はパーソナルセグメントの内側、4689 はストラテジーの内側） |
| `sotp` | 金融部分は DCF ではなく PBR×純資産で評価し、非金融の事業価値と SOTP で合算する。`valuation_method: "pbr"` のセグメントが最低1つ必要 |

### `sotp.segments[].valuation_method: "pbr"`

金融セグメント専用。基数が EBITDA ではなく**純資産**、倍率が EV/EBITDA ではなく
**PBR**、結果が EV ではなく**株主価値**になる。

- `net_assets_mn`（必須）: 金融セグメントの純資産（JPY mn）。推測値は不可
- `peers[].pbr` / `peers[].roe` を読む（`ev_ebitda` / `opm` ではない）。Peer Comps の
  見出しも自動で PBR / ROE に切り替わる
- `da_allocation_pct` は 0 にする（D&A は EV セグメント側に全額配賦）

### 型E の Equity Bridge

```
非金融 EV（EV セグメントの合計 + 持分法）
  − 非金融 net debt              ← sotp.consolidated.net_debt（MI を含めない）
= 非金融の株主価値
  + 金融セグメント株主価値（PBR × 純資産）
= SOTP 株主価値（割引前）
  × (1 − コングロマリット・ディスカウント)
  − 非支配株主持分                ← sotp.consolidated.minority_interest
= Equity Value（親会社株主帰属）
```

金融子会社の株主価値は**自身の預金を既にネットしている**ので、EV に足してから親の
net debt を引くと資金調達を二重に控除する。だから net debt 控除の**後**に加算する。

**注意（実際に踏んだ落とし穴）**: top-level の `net_debt`（MI を**含む**定義）と
`sotp.consolidated.net_debt`（MI を**含まない**定義）は型E では意図的に別物である。
`generate_sotp.py` は型A〜D では top-level を SOTP に注入するが、型E では注入しない
（注入すると MI を二重に引き、9433 の初回生成は 1株 2,532 円 ―― 正しくは 2,657 円 ――
になっていた）。

### 実例

`data/overrides/9433_overrides.json`（KDDI、auフィナンシャルホールディングス 100%）。
KDDI は「auフィナンシャルホールディングスを除く連結ベース」の P/L・BS を自ら開示して
おり（2026年3月期 決算詳細資料 p.3）、連結消去後の非金融系列が推計なしで得られる
数少ない例である。他社では金融子会社の純資産を子会社側の開示（銀行法ディスクロージャー
誌、単独上場子会社の決算短信）から取りに行く必要がある。

---

## 型F の節と加算脚（2026-09-07 追加 / 追補15 で更新）

### `company_type: "F"` — 持分法主導

連結利益の相当部分が持分法投資損益で、対応するキャッシュは受取配当のみ、という会社
（総合商社、日本郵船など）。連結 FCF を割り引く通常の DCF は投資そのものの価値を取りこぼす。

```
1株 Target = コア DCF の1株値（フロア適用後） ＋ 加算脚の1株値
```

宣言時に実行前チェックが要求するもの: `equity_method` / コアベースの `hist_*`（または
`segments`）/ `net_debt` の明示。持分法投資は資産側にあるが**コアの FCF を生まないので
`net_debt` で相殺してはいけない**（相殺すると加算脚と二重に効く）。

### `equity_method` ブロック（型E / 型F で使える）

| キー | 意味 |
|---|---|
| `method` | `book_value` / `listed_stakes` / `market_stakes` のいずれか（必須） |
| `balance_mn` | 持分法投資の連結BS 残高。`market_stakes` では不要 |
| `multiple` | `book_value` の倍率。既定 1.0 |
| `listed_stakes[]` | `name` / `ticker` / `market_cap_mn` と、`ownership`（listed_stakes 方式）または **`economic_interest`**（market_stakes 方式）。`voting_interest` は併記用で評価には使わない |
| `listed_book_mn` | `listed_stakes` 方式で必須。非上場分 = `balance_mn − listed_book_mn` |
| `unlisted_book_mn` | `market_stakes` 方式で使う非上場分の簿価 |
| `fair_value_based` | 残高そのものが公正価値で計上されている場合に `true`。check 26 の「簿価だから過小」WARN を抑止する（6971 の FVTOCI 政策保有株式） |
| `fx_note` | 外貨建ての時価を使うときの為替の明記 |
| `label` | 加算行のラベル。既定「持分法投資価値」 |

**方式の選び方は開示側の事情で決まる。**

- `market_stakes` … 上場子会社の**時価 × 経済的持分** ＋ 非上場分の簿価。
  連結子会社が上場している場合（4689/9434 の PayPay = Nasdaq: PAYP）。
  **経済的持分は議決権比率とは別物**で、評価に使うのは経済的持分のほう
  （4689 の PayPay は議決権 54.6% / 経済的持分 31.1%）。
  入れ子連結（親も子も同じ上場子会社を連結する）では、各モデルが自社の経済的持分のみを評価する。
- `listed_stakes` … 上場持分先の**簿価が注記から取れる**とき。非上場分は残高 − 上場分簿価。
- `book_value` … 上記が取れないときの最終手段。採ると check 26 が WARN を出す
  （上場持分の含み益ぶん過小に出るため）。ただし `fair_value_based: true` なら PASS。

### フロア規則（追補15 A-1 §3）

コア脚のエクイティ（EV − net_debt）が負または算定不能なら **コア = 0**、Target は加算脚のみ。
Excel 側で `=MAX(0, IFERROR((<コア式>)*1, 0)) + C{加算行}` として表に見える形で床を張る。
発動すると check 26 が WARN「コアDCF不成立」を出し、§AF の相対評価として読む必要がある。
型F では連結の net_debt を全額コア側で引くため、コアの株主価値が負になること自体は矛盾ではない
（負債の一部が持分法投資を賄っている）。

### 行を挿入しないこと（重要）

加算脚は **Executive Summary の既存の空き行に書く**。`openpyxl` の `insert_rows` で行をずらすと、
上にあった**結合セル**（注記行の B:C 結合など）が別の行に被さり、被った側は `MergedCell` になって
**書き込みが黙って捨てられる**。追補15 でこれが実際に起き、"Comps - EV/EBITDA" が空セルになった
まま、check 20 が空を「意図的な除外」と読んで 9 モデルすべてで PASS していた。
型D（`ddm_ri.py`）は 2 行必要なので挿入を残しているが、**挿入直後に C 列にかかる結合を解除**している。
