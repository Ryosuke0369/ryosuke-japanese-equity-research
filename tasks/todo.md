# DCFパイプライン テンプレート恒久修正 (2026-07-31)

Branch: `template-hardening-20260731` / 仕様書: ユーザー提供「ClaudeCode実行仕様書」

## A群: テンプレート本体
- [x] A1 シナリオ切替 D27 を MATCH 式に統一（segments 有無で分岐しない）
- [x] A2 Comps 統計範囲から subject company を自動除外（動的行算出）
- [x] A3 Comps 時価総額/EV/PBR/ROE 数式化 + Book Value 列 + Peer prices as-of 注記

## B群: データ整合
- [x] B1 FS の OCF/現金/有利子負債を年度キー突合で割当（位置ベース廃止）
- [x] B2 C5/C18 を実績3年平均ベースに（予測からの逆算を廃止）
- [x] B3 C20 LTM Revenue の3成分ログ出力 + `ltm_revenue` override
- [x] B4 D&A欠損 peer を EV/EBITDA 統計から自動除外（有効 n<3 で INVALID）
- [x] B5 thesis/key_risks のトークン→数式化

## C群: 自動検証層
- [x] C1 `scripts/validate_output.py` 新規（13チェック）+ generate_dcf からの自動呼出
- [x] C2 validator 追加チェック（配列長・トークン・ターミナルcapex）+ peer 鮮度チェック
- [x] C3 Adjustments Log シートの標準生成

## 回帰テスト
- [x] 3687 再生成 → validate FAIL 0 / WARN 2
- [x] 285A 再生成 → validate FAIL 0 / WARN 0、D27=MATCH、自社除外、Downside 2 実動

## Review（2026-07-31）

### コード変更の影響（main ワークツリーでの A/B、同一入力・同日）
`git worktree add ../_ctrl_main main` に同じ `data/` を置き 3687 を生成して全セル比較。
差分は **71セルのみ**で、すべて意図した修正:

| 箇所 | main | 新 | 由来 |
|---|---|---|---|
| Comps 統計 15-17行・C27/C28 | 自社込み 1,519 / 1,835 | **1,256 / 1,737** | A2 自社除外 |
| Comps D5/E5/J5-M5 | 静的 67,054 | 数式 67,053.69 | A3 |
| Comps P列 Book Value / Q列 Note / B3 as-of | なし | 追加 | A3 |
| FS 19/21/22/26/27/28行 | 3年ズレ | 年度一致 | B1 |
| DCF C5 / C18 とラベル | 2.786% / 2.547% | 1.25% / 1.83% | B2 |
| Adjustments Log シート | なし | 追加 | C3 |
| NWC Schedule / Sensitivity / WACC / PGM / Exit | — | **完全一致** | 影響なし |

### 不変6値の照合
`size_premium` を 0.03 に戻した検証ランで **WACC 11.54% / Comps EV-EBITDA 1,256 /
PER 1,737 が完全一致**。PGM 970→1,017・Exit 1,559→1,638・Target 1,381→1,412 の差は
**main コードでも同値**（1,017 / 1,638）であり、原因は `stub_fraction` 1.0→0.5
（当日の LTM が 2Q 進行）。**コード起因の valuation 変化はゼロ**。
なお現行 `data/overrides/3687_overrides.json` の `size_premium` は 0.025 のため、
そのまま生成すると WACC は 11.04% になる（overrides 側の値、要確認）。

### 285A
validate FAIL 0 / WARN 0。D27=`MATCH(C27,B69:B73,0)`、Comps 中位 EV/EBITDA **16.86x**
（自社除外）、FS の OCF は FY2024/3 に **195,111**。ドロップダウン実動確認:
Base 9,121 → Upside 15,181 → Management 30,952 → Downside 1 3,442 → **Downside 2 188**（PGM）。

### 残メモ
- 3687 FY2025/9 の有利子負債は EDINET に無く**空欄**（旧モデルは FY2024 の 2,085 が
  誤って入っていた）。埋めるなら overrides に `hist_debt` を入れる。
- 仕様書 §6-4 の WARN 期待（285A の terminal capex / Exit 乖離）は現行 overrides では
  発火しない（exit_multiple が 14.0 → 5.0 に、terminal capex/D&A も 0.98x に更新済み）。
  両チェックは 3687（0.86x / 1.94x 乖離）および 7203・8267 の overrides 事前警告で発火を確認。
- `reports/2359_market_analysis_20260509_v2.xlsx` は本作業前から未コミット変更あり（未関与）。

---

## 追補: 姉妹テンプレート修正（2026-07-31）

- [x] M1 `_read_comps_stats` のラベルベース化 + 自社除外（p25/中央値/p75 も peer から再計算）
- [x] M2 Block 3 補間式に IFERROR + alpha-scan 単調性チェック + NOPATフロア注記
- [x] M3 DCF 読取のラベル探索化（WACC/g/tax/net debt/shares/stub/PGM/Exit/行42-44）+ フォールバック警告
- [x] M4 Scorecard Factor 3 の空欄→中立0化、C9 を C36 と同じ買/売に統一
- [x] M5 Narrative Stage の擬似入力セルを OUTPUT 化 + 再生成注記（案A採用、案Bは将来課題）
- [x] M6 fallback パスのエラー明確化 / `market_alpha_cell` 必須化 / price_points 先頭検証 /
      Block 4 を `alpha (proxy)` 表記化 / デモ出力を reports→tmp へ
- [x] S1 `_read_dcf_crosscheck` ラベルベース化（scripts/generate_sotp.py の COM 版も同様に修正）
- [x] S2 IHI 固有文言を config 供給化 / segment key デフォルト廃止（未指定はエラー停止）
- [x] S3 単位契約の明文化 + `shares_outstanding > 1e8` でエラー停止
- [x] S4 Table2 の割引率をセル参照化 / 行レイアウト先決め / Cover 行番号 return 化 /
      as-of 注記 / float 等値比較を許容誤差比較へ
- [x] N stage_pre_cap + pre_cap_reason 追加、bool 弾き、Stage 0 優先順位を docstring 明記
- [x] validate_output にチェック 14-20 追加（シート名で種別自動判別）

### 検証結果
- `test_narrative_stage.py`: **5/5 PASS**（新規2アサーション追加後も既存期待値は不変）
- 3687 market_analysis（新DCF入力）: validate **FAIL 0 / PASS 5**、
  peer 5社 [6-10] を使用（旧実装は rows 5..9 = 自社込み4社で **5社目を取りこぼし**）、
  peer-only 中央値と DCF シート統計セルが**完全一致**（NOTE なし）、alpha=1.0 = PGM 1,064（-0.02%）
- 2359 / 4192 デモ（旧DCF入力）: 完走。peer-only 中央値が旧DCFの自社込み中央値と食い違う旨を
  NOTE で明示（4192: 3.3227 vs 3.5818 → justified price 502 vs C27 538）。**旧DCFを再生成すれば解消**
- 7013 SOTP 再生成: validate **FAIL 0 / PASS 4**、Fair Value **3,072.064691736476 で修正前と完全一致**、
  Sensitivity C27 も一致。単位ガード / segment key 必須化とも意図どおりエラー停止を確認
- 8410（銀行型）read-only 読取テスト: ラベル探索で 389/431/505/196 を正しく取得
  （当該ファイルの Exec Summary は結局 16-19 のままだったが、ラベル読取でも同値）

---

# Fix: 株価更新 + Comps時価総額の固定化 + 旧成果物整理 (2026-06-12 夕)

## TODO
- [x] 1. overrides 株価更新（承認値）: 4192 = 244 / 5246 = 557（2026-06-12終値, yfinance,
       承認 2026-06-12）。_price_note に値・取得日・出所・承認日を記録
- [x] 2. comps_fetcher.py: Market_Cap 列の all-or-nothing 化（一部空欄は ValueError 停止）
       + ソースエコー（CSV (static) / yfinance live 警告）。行単位サイレント混在を廃止
- [x] 3. comps CSV へ本日終値ベースの時価総額を記入: 5246 = 列新設（5社）、
       4192 = 既存列を更新（LightWorks 4267 は上場廃止/yfinance 404 → 継承値 10,730 据置）
- [x] 4. docs/overrides_schema.md: Market_Cap 列仕様（全行ルール・エコー・as-of 表）追記
- [x] 5. models/5246_DCF_Model_20260604.xlsx → models/archive/ へ移動
- [x] 6. 4192/5246 再生成 + recalc PASS + 実セル照合
- [x] 7. ドリフト排除実証: 5246 二連続生成で comps 由来値が浮動小数まで完全一致
- [x] 8. フォールバック検証: 2359（列なし）で警告エコー実機確認、DCF値不変
- [x] 9. lessons.md「ライブ取得依存は再現性を壊す」追記

## レビュー（実装結果 2026-06-12）
- 5246: C9 557 / PGM 347 / Exit 1129 / EV/S 456 / PER N/A / Target Mid 644 / BUY (+15.6%)
- 4192: C9 244 / PGM 131 / Exit 371 / EV/S **538→517** / PER N/A / Target Mid **347→340** /
  BUY (+39.3%)。517 への変化は時価総額の本日値リフレッシュによるもの:
  Toyokumo 22,846→19,918 (-12.8%) で EV/Rev 中央値が Smaregi 3.582 → Toyokumo 3.434 に
  交代 → (4,895×3.434+1,559)×1e6/35,507,400 = 517。コード変更由来の変化なし。
- 2359（ライブ経路）: 警告エコー確認。PGM 3150 / Exit 4343 不変、comps 由来 ±数円の
  ドリフト（2460→2454 等）はライブ経路の既知挙動 = 警告が必要な理由そのもの。
- 残課題: da_pct（4.16% vs 12.45%）・税金前提（NOL未考慮）は**未着手のまま保留**。
  2359 の comps CSV 静的化も将来候補（最終ラン前に Market_Cap 列を記入）。

---

# Fix: Target Mid の無意味値混入（赤字×PER）+ 残課題 (2026-06-12 フェーズB/C)

## 背景
- Executive Summary!C10 (Target Mid) = AVERAGE(C16:C19) に、赤字企業では
  「PER中央値×マイナス純利益」の無意味値が混入（5246: -534、4192: -11）。
- フェーズA確定事項: PGM 347 は overrides指定 g=1.5% + 自動算出 da_pct=4.16% による
  正しい計算結果。da_pct の変更はユーザー判断で保留。5246 の数値前提は変更しない。

## TODO
- [x] 1. dcf_comps_template.py: core_net_income<=0 で Comps!C28 を "N/A" 化
       （AVERAGE/MIN/MAX はテキストを無視 → Target Mid は有効手法のみの平均に）
- [x] 2. 同: Exec Summary に除外注記セル（B20）を明示出力（黙って除外しない）
- [x] 3. 同: EV/EBITDA primary × core_ebitda<=0 の同型リスクも同方針で処理
       （PBR法は Target Mid に存在せず → 対象外と報告）
- [x] 4. market_analysis_template の C28 参照影響確認（_num→None で安全、と確認済み）
- [x] 5. run_market_analysis_4192.py: 最新日付モデル自動選択 + --dcf 引数化
- [x] 6. run_market_analysis_5246.py 新規（単一セグメントfallback、市場データは中立placeholder）
- [x] 7. models/4192_DCF_Model_20260604.xlsx（未recalc地雷）→ models/archive/ へ移動
- [x] 8. 5246/4192 再生成 + recalc + 実セル照合（4192 も赤字（NI -17mn）と判明 →
       PER 除外で Target Mid 257→347 / HOLD→BUY。黒字回帰は 2359 で実証）
- [x] 9. market_analysis を 5246/4192 最新モデルで実行・完走確認（alpha=1.0⇔PGM 一致、
       EV/Sales median ⇔ C27 一致）
- [x] 10. lessons.md 追記（複数手法平均への無意味値混入）、一時ファイル掃除

## レビュー（実装結果 2026-06-12）
- 変更: `templates/dcf_comps_template.py`（PER_EXCLUDED / EBITDA_EXCLUDED ガード、
  Comps!C27/C28 の "N/A" 化、Exec B20 注記、D18/D19 を IF(ISNUMBER(...)) 化）、
  `scripts/run_market_analysis_4192.py`（最新dated モデル自動選択 + --dcf、reports/
  上書きガード）、`scripts/run_market_analysis_5246.py`（新規）。
- 5246 (recalc PASS): PER=N/A・B20注記あり。Target Mid 349→**644** = AVG(347,1129,456)、
  SELL→**BUY**（+15.4% vs 558）。Range "347 - 1129"。PGM 347 / Exit 1129 は不変。
  ※Comps EV/Sales 455→456 は comps CSV に Market_Cap 列が無く時価総額を都度
  yfinance ライブ取得しているための日中ドリフト（テンプレ修正と無関係）。
- 4192 (recalc PASS): 赤字（NI -17.357mn）→ PER=N/A・注記。Target Mid 257→**347** =
  AVG(131,371,538)、HOLD→**BUY**（+18.4% vs 293）。手法値 131/371/538 は完全不変。
  C9 240→293: overrides `current_price:293` の再適用契約どおり（旧モデルの 240 は
  契約外の live 値。293 自体は 06-02 終値で古い → 要リフレッシュ）。
- 2359 黒字回帰 (recalc PASS): PER 法は従来どおり数式で包含（1,972）、注記なし、
  DCF 3150/4343 不変（Comps/現在値の±数円は yfinance ドリフト）。
- market_analysis: 5246/4192 とも 20260612 モデルを自動選択し完走。
  5246: alpha=1.0 implied 347 ⇔ PGM 347（+0.02%）、median 456 ⇔ C27（+0.5円）。
  4192: 131 ⇔ 131（+0.17%）、538 ⇔ 538（-0.3円）。
- 残課題: (1) 4192 overrides の current_price=293 が stale、(2) comps CSV に
  Market_Cap 列を追加して時価総額を固定すべき（ドリフト排除）、(3)
  models/5246_DCF_Model_20260604.xlsx（comps混入事故時代の旧成果物）の archive 移動は
  ユーザー判断待ち、(4) da_pct 4.16% vs 12.45% 問題はフェーズA報告のとおり保留。

---

# Fix: overrides サイレント不反映の恒久修正（方針Y: バリデーション層） (2026-06-12)

## 背景（フェーズ1で確定した事実）

- **WACC**: 現存4192/5246モデルのWACC入力は正規化済みoverridesどおり（初期化値ではなかった）。
  ただし旧overrides形式（ネスト`wacc_inputs`/Bull・Bear名）が黙殺される構造は実在し、
  06-04正規化以前の5246生成では実際に発生していた（本ファイル下部アーカイブ参照）。
- **Comps**: 5246モデルに4192の5社が混入。原因は (a) ELEMENTSの正しい5社が
  `data/comps/5246_comps.txt`（.txt/UTF-16）で契約パス `.csv` にヒットせず黙殺、
  (b) 遺物 `scripts/comps_input.csv`（4192セット）を `--comps-csv` で明示指定していたこと。
- **market_analysis**: `float(C26 or 0)` により未recalcのDCFを渡すとWACC=0で黙って完走する地雷。

## TODO

- [x] 1. `scripts/overrides_validator.py` 新規作成（ホワイトリスト＋構造検証、違反は全件列挙でエラー）
- [x] 2. `scripts/generate_dcf.py` 統合: バリデータ呼出し / `--allow-unconfirmed` / `--no-comps` /
       comps CSV 不在・パース失敗のエラー化（.txt検出時はリネーム案内）/ ticker `.T` 正規化 /
       Effective WACC inputs + Comps 5社名のエコー出力
- [x] 3. `templates/market_analysis_template.py`: C26/C13/C6/C15/C16 の `or 0` 撤廃 →
       未recalc検出でエラー停止（recalc_excel_com.py への案内付き）。WACC<=0 もエラー
- [x] 4. `data/comps/5246_comps.csv` を `5246_comps.txt` から正規化して作成（ELEMENTS 5社）
- [x] 5. 遺物 `scripts/comps_input.csv` 削除
- [x] 6. 5246 確定値反映: FY2025/11決算短信（2026-01-13開示）より shares=27,115,114 /
       net_debt=-548mn（借入金のみ定義、ユーザー承認済み）/ price=558円（ユーザー承認済み）
- [x] 7. 再生成検証: 4192 / 5246 / 2359 → 実セル照合 + recalc PASS
- [x] 8. 違反系テスト: 全11ケースでエラー検出を確認
- [x] 9. `docs/overrides_schema.md` 契約仕様書作成、CLAUDE.md 更新
- [x] 10. lessons.md 更新、調査用スクリプト削除

## レビュー（実装結果）

### 変更ファイル
- `scripts/overrides_validator.py`（新規）: ホワイトリスト（実際に消費されるキーのみ）、
  ネスト構造・旧シナリオ名の名指しエラー、typoサジェスト、配列長検証、nwc_items 形状検証
  （itemized の scenario_key は動的に許可）、`__CONFIRM__` 検出。違反は全件まとめて報告。
- `scripts/generate_dcf.py`: 検証呼出し（json.load直後）、comps契約強制、ticker正規化、反映エコー。
- `templates/market_analysis_template.py`: 未recalc DCF をエラー化（旧: 黙ってWACC=0）。
- `data/overrides/4192_overrides.json`: 死にキーを契約準拠化（hist_da→hist_depreciation、
  fiscal_year_end→fiscal_year_end_month:12、company_name_jp→company_name、他は `_` 接頭辞化）。
- `data/overrides/5246_overrides.json`: `__CONFIRM__` 3値を短信確定値で置換。
- `data/overrides/6363/6365_overrides.json`: 死にキー `comps`（Step6でCSVから常に上書きされる）
  を `_comps_reference` にリネーム。
- `data/comps/5246_comps.csv` 新規（Showcase/Cyber Security Cloud/FFRI/Headwaters/User Local）。
- `scripts/comps_input.csv` 削除。
- `docs/overrides_schema.md` 新規、`CLAUDE.md` 更新。

### DoD 結果（2026-06-12）
1. ✅ バリデータ回帰: 既存8 overridesファイル全PASS（修正後）。違反系11ケース全てエラー検出。
2. ✅ 4192 再生成+recalc: C7-C12 = [0.023, 1.45, 0.06, 0.04, 0.01, 0.03]、WACC=14.59%、
   Comps=eWeLL/Toyokumo/Smaregi/LightWorks/SMS。Exec Summary (257/131/371/538/-11) は
   recalc済み旧0603モデルと完全一致 → リグレッションなし。
3. ✅ 5246 再生成+recalc: WACC=13.54%（指定どおり）、terminal 0.015、Exit=EV/Sales 4.0、
   shares=27,115,114 / price=558 / net_debt=-548（全て確定値）、Comps=ELEMENTS固有5社。
   Target Mid 349 vs 558 → SELL。PGM 347 / Exit 1129 / Comps EV/S 455。
   （注: Comps PER -534 は赤字企業ゆえ無意味な行 — 既知の表示仕様）
4. ✅ 2359（3月期・revenue_pct・segments・第3銘柄）再生成+recalc: override指定キー
   （risk_free/size_premium/cost_of_debt_at）反映、未指定キーは自動値。WACC=12.80%。
   Comps=Systena/SRA HD/CEC/Cresco/DTS。
5. ✅ market_analysis ガード: 未recalcの 4192_0604 モデル → エラー停止、
   recalc済み 4192_0603 → 正常動作（wacc=0.145922）。

### 残メモ
- `models/4192_DCF_Model_20260604.xlsx` は未recalcの旧成果物（地雷だったもの）。
  最新は `*_20260612.xlsx`。旧dated modelの整理はユーザー判断。
- `scripts/run_market_analysis_4192.py` は dcf_excel_path が 0603 固定。次回market_analysis
  実行時は 20260612 モデルへ更新すること。

---

# ［アーカイブ・完了］Fix: generate_dcf.py 5246 が「5年分EDINET取得」で停止する問題 (2026-06-04)

（以下は完了済みの過去タスク記録。当時の「残課題」が今回のサイレント不反映問題の伏線だった）

## 真因（コードで特定）

ELEMENTS(5246) は **11月期決算**。有報提出は2月だが探索窓が2月をカバーせず取得失敗。
`get_document_ids` にバジェット無しで約4分スピン。詳細は git 履歴参照。

## 実装結果（要点）
- `edinet_fetcher.py`: fiscal_year_end_month による動的SEASON生成、MAX_API_CALLS=400、
  found-season short-circuit。
- `generate_dcf.py`: fundamentals_fy<YYYY> 注入（Step 1.5）、`__CONFIRM__` スキップ。
- DoD: 4192リグレッションなし、5246完走（`--comps-csv scripts/comps_input.csv` 使用
  ← **これが後にComps混入事故の原因と判明**）。
- 当時の残課題（Bull/Bear名・wacc_inputs ネストが反映されない）→ 2026-06-12 の
  バリデーション層導入で構造的に解決。
