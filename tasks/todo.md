# 5726 大阪チタニウムテクノロジーズ DCFモデル生成 (2026-08-26)

> ⚠️ **復元メモ (2026-08-26 incident — クローズ済み)**
> 2026-08-26 のテンプレート修正作業の冒頭で、このファイルの未コミット版を全文上書きした。
> 以下「銘柄型判定」〜「設計判断」は上書き直前に読み取れていた範囲の**逐語復元**。
> 末尾にあった「## Review (2026-08-26)」本文は **lost (2026-08-26 incident)** ——
> 逐語では復元不能と最終判断した（復元経路の確認結果は下記）。
> 代わりに、一次情報から再構成できる内容を「### Review 再構成」に書き戻した。
>
> **確認した復元経路（すべて空振り）**
> | 経路 | 結果 |
> |---|---|
> | `git reflog` | 当該状態のコミットは存在しない（e5effd6 の次は e26b801） |
> | `git stash list` | 空 |
> | `git fsck --dangling` の blob 全走査 | todo.md の内容を含む blob なし（一度も `git add` されていない） |
> | VS Code ローカル履歴 (`AppData/Roaming/Code/User/History`) | 10件、いずれも本リポジトリ外 |
> | Windows File History | 未構成 |
> | OneDrive バージョン履歴 | 本リポジトリは OneDrive 配下ではない |
> | JetBrains / Notepad++ / Sublime のバックアップ | いずれも不在 |
> | VSS シャドウコピー (`vssadmin list shadows`) | **管理者権限が必要で本セッションからは確認不可**。復元を試みる場合は管理者権限のコンソールで `vssadmin list shadows` → 該当時点のスナップショットから `tasks/todo.md` を取り出すこと（これが唯一未確認の経路） |

出力: `models/5726_DCF_Model_20260826.xlsx` (DRAFT — ユーザー検証後にFINAL化)
手順書: docs/DCFパイプライン標準運用手順書.md (v2) / 契約正本: docs/overrides_schema.md

## 銘柄型判定
- **型B: シクリカル** (素材・スポンジチタン)。銀行なし・captive financeなし。
  FY2022/3 は営業赤字 △1,914 → FY2025/3 ピーク 10,088 → FY2026/3 5,524 と符号反転級の振れ。
  → サイクル型シナリオを設計。ピーク利益×高Exit倍率の二重計上を避ける。

## タスク
- [x] 1. 手順書v2 / overrides_schema / lessons.md / CLAUDE.md を読む
- [x] 2. EDINET から FY2021/3〜FY2026/3 の一次データを取得・プロンプト値と照合
      → FY2024-26 の 売上/COGS/SGA/OP/NI がプロンプトと完全一致。FY2022/3・FY2023/3 も取得
- [x] 3. beta 実測 (週次2年 vs 1306.T) = 1.553 / corr 0.442
- [x] 4. risk_free: 財務省 jgbcm.csv から 10年JGB = 2.897% (2026-08-25基準)
- [x] 5. Peer選定と上場状態確認
      → **5727 東邦チタニウムは2026-05-28上場廃止**(JX金属5016が株式交換0.70で完全子会社化)
      → 海外チタン/航空機素材ピア (ATI / CRS / 宝钛股份) を追加して基準を是正
- [x] 6. comps CSV 作成 (data/comps/5726_comps.csv、Market_Cap全行静的)
- [x] 7. overrides JSON 作成 (data/overrides/5726_overrides.json)
- [x] 8. generate_dcf.py 実行 → recalc → validate (FAIL 0 を確認)
- [x] 9. Reverse DCF シート追加 (scripts/add_reverse_dcf_sheet.py) — 逆算が本件の最重要成果物
- [x] 10. Adjustments Log に全推定セルを記録 (DRAFT状態)
- [x] 11. 最終 recalc + validate 再実行、3行サマリー作成

## 設計判断（要点）
- **正常化**: 為替差益/補助金(営業外)・特損(除却1,722/減損461/圧縮92/環境引当343)は
  営業利益ベースのモデルに元々含まれない → 追加調整不要。core_ebitda は FY2026/3 実績ベース。
- **capex/D&A は direct方式**: 能力増強プログラムは売上比例しないため。
  terminal capex/D&A = 3,900/3,950 = 0.99x (validator [0.90,1.15] 内)。
- **NWC は days方式**: 棚卸資産がFY2026/3で380日(COGS基準)と異常に膨張。
  歴史平均240日への正常化速度をシナリオの主要ドライバーにする。
- **Reverse DCF が最重要**: 定常OP逆算 / FY25ピーク比 / 到達年数感応度。

### Review 再構成 (2026-08-26, 逐語ではない)

逐語の Review 本文は失われたため、`models/5726_DCF_Model_20260826.xlsx` の
Adjustments Log 21件と `docs/DCFフォーマット標準メモ_20260826.md` から再構成した。
**この節は一次情報からの再構成であり、当時書かれた文章そのものではない。**

- **納品状態**: DRAFT。validate_output は FAIL 0 / WARN 0。FINAL 化は未実施（手順書v2 §6-3）。
- **結論値**: 現値 2,727 / Target(DCF Mid) 379 / PGM 200 / Exit 558 / SELL / WACC 9.06%。
  Comps は EV/EBITDA 1,582・PER 2,247 だが Target 不算入。
- **逆算DCF（本件の最重要成果物）**: 現値は定常営業利益 18,466百万円（FY2025/3 ピーク
  10,088 の 1.83倍）を5年で到達し永続することを織り込んだ価格。ランプ無し（Block A）でも
  15,919（ピーク比 1.58倍）が必要。ベースケースが説明できる EV は 53,209 にとどまり、
  1株あたり 2,527円 が中期economicsで説明できない。
- **確定した推定・設計判断**（Adjustments Log 参照）: risk_free 2.90%（財務省 jgbcm.csv）/
  beta 1.55（1306.T 回帰、相関0.442）/ size_premium 1.5%（閾値の直上0.3%）/
  de_ratio 0.4568（ネット基準を採用）/ terminal_growth 1.5% / exit_multiple 8.0x（国内素材ピア水準）/
  capex・D&A は direct 方式 / NWC は days 方式（DIH 260〜420日がFCF最大のスイング）。
- **未解決として残した項目**: 増設分の稼働時期・能力増分・売上寄与が未開示 /
  FY2027/3 Q1進捗率34%と会社計画の保守性 / core_net_income 2,576 が特損2,619で歪んだ分母
  （正常化なら約4,465、PER 約23倍）。
- **重要な外部事実**: 5727 東邦チタニウムは 2026-05-28 上場廃止（JX金属が株式交換0.70で
  完全子会社化）。国内唯一の直接比較対象を失ったため海外ピア3社を追加した。
- **当時テンプレ課題として報告した2件は、その後 e26b801 で恒久修正済み**:
  Target Mid が Comps を含む AVERAGE(C16:C19) だった件 / Exit 法に負値ガードが無く
  Downside 2 で負の株価が平均に残った件。

---

# DCFパイプライン テンプレート修正 + 構造リファクタ (2026-08-26)

前提: docs/DCFパイプライン標準運用手順書.md (v2) + docs/DCFフォーマット標準メモ_20260826.md
出典: 標準メモ §3「テンプレート(Python)修正バックログ — 恒常問題4件」+ §3-5「構造リファクタ」

2フェーズ = 2コミット。フェーズをまたいだ変更は混ぜない。

## Phase 1: テンプレート恒久修正 (標準メモ §3-1〜§3-4)

- [x] 0. 標準メモを docs/ へ格納 (Desktop/Downloads の最新版 7,874 bytes を正とする)
- [x] 1. **Exec Summary の Target 計算式**: `AVERAGE(C16:C19)` → DCF2本 `AVERAGE(C16:C17)` のみ
      - Comps 行 (C18/C19) は `[参考・Target不算入]` ラベルで残置
      - C10/C11/C12 に COUNT/ISNUMBER ガード（両手法 INVALID 時に #DIV/0! を出さない）
      - Integrated Valuation Range も全手法テキスト時に "0 - 0" を出さないようガード
- [x] 2. **負の Exit 株価ガード**: C17 に PGM と同等の `EV < net debt → INVALID` ガード
- [x] 3. **Financial Statements の負債ラベル**: `Total Interest-bearing Debt (short + long)`
- [x] 4. **Reverse DCF の標準化**: `templates/reverse_dcf_sheet.py` を新設し、
      dcf_comps_template が標準8枚目（DCF Model の直後）として生成
      - 他シート参照は全てテンプレの行定数から解決（G33/C53/D16 のハードコードを廃止）
      - B_N 閉形式の自己テストを生成のたびに実行
      - `reverse_dcf` overrides ブロック（全キー任意）+ hist 配列からの自動導出
      - add_reverse_dcf_sheet.py は既存ブックへの retrofit 用ラッパに縮退（347→190行）
- [x] 5. 契約/検証の追随: overrides_validator の `reverse_dcf` 検証 /
      docs/overrides_schema.md / validate_output のチェック14-16
- [x] 6. 検証: 合成シクリカル config で生成→recalc→validate（FAIL 0）、
      Downside 2 で両脚 INVALID → Target "N/A" を確認、
      5726 実モデルへの retrofit で 17項目 PASS

## Phase 2: 構造リファクタ (標準メモ §3-5)

- [x] 7. `add_segment_bridge.py` を設定ファイル駆動に全面書き換え
      (`data/segments/<ticker>_segments.json`)。3110 と 3687 の両方を設定で再現し、
      `add_segment_bridge_3110.py` を削除
- [x] 8. `fill_adjustments_log.py` を新設 (`data/adjustments/<ticker>_adjustments.json`)。
      5726 の21件を移設して `fill_adjustments_log_5726.py` を削除
- [x] 9. `scripts/check_script_naming.py` を追加。generate_dcf.py が起動時に警告として実行
- [x] 10. CLAUDE.md / 手順書v2 に運用ルールを明記

## Review

### Phase 1 (2026-08-26)

**変更の性質**: 5726 のモデルに**手修正として入れた4点をテンプレート本体へバックポート**した。
手修正版 `models/5726_DCF_Model_20260826.xlsx` と同じ挙動を、以後は生成時点で得る。

| メモ§3 | 対応 | 効果 |
|---|---|---|
| 1 | Exec C10 = `IF(COUNT(C16:C17)=0,"N/A",ROUND(AVERAGE(C16:C17),0))` | Comps が Target に混入しなくなった。合成テストで旧式 629 → 新式 488（Exit のみ） |
| 2 | Exec C17 に `EV < net debt → "INVALID"` | Downside 2 で Exit の負値が平均に残る事故が消えた（両脚 INVALID → Target "N/A"、"SELL" の誤表示も消える） |
| 3 | FS B27 = `Total Interest-bearing Debt (short + long)` | ラベルと実体（EDINET `total_debt` = 短期+長期）が一致 |
| 4 | `templates/reverse_dcf_sheet.py` を新設、標準8枚目に | ad-hoc スクリプトの後付けが不要。参照先アドレスはテンプレの行定数から解決するのでレイアウト変更に追随する |

**設計判断**:
- 逆算シートは *generate_dcf.py のステップ追加*ではなく **テンプレート内で生成**した。
  post-process にすると (a) openpyxl の再ロード／再保存で全シートのキャッシュ値が落ちる、
  (b) 他シートのアドレスをハードコードで持ち直すことになる、の2点が避けられないため。
  結果としてパイプラインのステップ数は 9 のまま。
- 自動導出できない銘柄（実績営業利益なし／ピークが赤字）は**シートを作らない**。
  ゼロ埋めのシートは「逆算した」という誤った既成事実を作るため。理由は
  Pipeline Metadata の `reverse_dcf_sheet` に残し、validate のチェック16が WARN で報告する。
- 比率セルは全て `IFERROR(...,"N/A")` で包んだ。Downside シナリオは分母（ピーク利益・
  終年D&A・WACC−g）をゼロや負に振らせるため、素の #DIV/0! はチェック1を FAIL させる。

**検証**:
- 合成シクリカル config（5726 に似せた5期・営業赤字→ピーク→減速）で生成 →
  Excel COM recalc → validate: **FAIL 0 / PASS 14**（WARN 2 は合成データ由来の
  capex/D&A 1.50x と PGM 逆算倍率乖離で、テンプレの問題ではない）
- 同ブックの Active Scenario を Downside 2 に切替えて recalc:
  EV_PGM △28,108 / EV_EXIT △11,048 < ネットデット 45,837 → C16/C17 とも INVALID、
  C10/C11/C12 は全て "N/A"（旧テンプレなら Target △25 と "SELL" を表示していた）
- 実モデル `models/5726_DCF_Model_20260826.xlsx` のコピーに retrofit ラッパを適用 →
  recalc → validate: **FAIL 0 / WARN 0 / PASS 17**。シート表題の重複括弧
  (`5726.T (TSE Prime)` 由来) も解消。
- `templates/reverse_dcf_sheet.py` 単体実行で B_N 閉形式の自己テスト OK。

**未対応（意図的）**:
- 標準メモ §3 の「任意」3件（実績ベース負債コスト、セグメントブリッジ、為替感応度Table3）は
  オプションモジュール扱いのため Phase 1 の範囲外。

### Phase 2 (2026-08-26)

**問題**: `scripts/` に動くコードの銘柄別コピーが生まれ始めていた
(`add_segment_bridge_3110.py` / `fill_adjustments_log_5726.py`)。コピーはロジックごと
分岐するので、片方に入れた修正がもう片方に届かない。放置すれば次の銘柄で3つ目が生まれる。

**方針**: 「汎用スクリプト + 銘柄固有は `data/` 配下の設定ファイル」に統一した。

| 対象 | 変更 |
|---|---|
| `scripts/add_segment_bridge.py` | 設定ファイル駆動に全面書き換え。セグメント売上／利益／D&A の各ブロックを共通ビルダーで生成し、連結値へのブリッジと tie-out 差異行を出す。自由形式の追加表 (`extra_tables`) は `{profit_r1}` `{c1}` 等の名前付きアンカーで上のブロックへ生き参照できる |
| `data/segments/3110_segments.json` | 旧 `add_segment_bridge_3110.py` の埋め込み定数を無改変で移設 |
| `data/segments/3687_segments.json` | 旧 `add_segment_bridge.py` (Fixstars 固有) の定数を移設。Normalization ブロックは定数の再掲ではなくセグメント行への生き参照に改善 |
| `scripts/fill_adjustments_log.py` | 新設。5/6フィールドのエントリ配列を読み、Pipeline Metadata 帯の**上**に挿入 |
| `data/adjustments/5726_adjustments.json` | 旧 `fill_adjustments_log_5726.py` の21件を無改変で移設 |
| `scripts/check_script_naming.py` | 新設。`scripts/` `templates/` の .py 名に銘柄コードがあれば exit 1。既存の `run_*` 系10本は grandfathered (リストは閉じている) |
| `scripts/generate_dcf.py` | 起動時に上記チェックを**警告として**実行（モデル生成自体は止めない） |
| CLAUDE.md / 手順書v2 | 運用ルールを明記 |

**設計判断**:
- 旧 `add_segment_bridge.py` にあった `patch_comps_exclude_self()` は移植しなかった。
  Comps 統計から自社行を除外するのはテンプレート本体の役目で（commit f61d242 以降、
  validate のチェック3が検証している）、この関数は「自社=5行目・統計=固定行」を仮定して
  数式を上書きする旧世代の後付けだった。今適用すると正しい数式を壊す。
- 開示のない期間には SUM を書かない。空セルの SUM は「0」という確定値に見えるが、
  正しくは「未開示」である（3110 のセグメント D&A は FY2026/3 のみ開示）。
- マージン表はセグメントを**ラベルで突合**する。売上非開示のセグメント
  （3687 の "Other (incl. CVC)"）が、位置合わせで別のセグメントと組まされないようにするため。

**検証**:
- `models/3110_DCF_Model_20260822.xlsx` のコピーに新スクリプトを適用 → recalc:
  エラーセル 0。tie-out 差異は Q1営業利益 **-12**、FY2026/3 D&A **+40** で、
  旧スクリプトが Adjustments Log に書き残していた値と一致。
  電子材料の FY2026/3 セグメント利益シェア 84%・セグメントマージン 31.6% も一致。
- `models/3687_DCF_Model_20260731.xlsx` のコピーに適用 → recalc: エラーセル 0。
  Solution マージン FY2025/9 = 35.3%、Normalized OI (ex-CVC) 2,810 = 29.2%、
  Fully Normalized 3,234 = 33.6%、顧客集中 Kioxia 17.1% — いずれも旧実装の記述と一致。
- `fill_adjustments_log.py` の出力を旧スクリプトの `ROWS` 定数と全21件×6列で突合: **差分0**。
  Pipeline Metadata 帯も無傷。
- `check_script_naming.py`: 削除前は2件を正しく検出して exit 1、削除後は exit 0。

**未対応（意図的）**:
- 既存の `run_market_analysis_<ticker>.py` 系10本は grandfathered。これらは
  「1銘柄の入力を汎用テンプレに渡すだけ」の薄いドライバでロジックの fork ではないため、
  今回の目的（分岐の解消）には該当しない。許可リストは閉じてあり、新規追加はできない。

---

# フォローアップ: 残タスク3件 (2026-08-26)

前回作業 (Phase 1: e26b801 / Phase 2: 3d2e453) の未達・要確認3件。タスクごとに別コミット。

## タスク1: オプションモジュール2件 — commit fe67463

- [x] 実装状態の確認 → **両件とも未着手**だった (grep で interest_expense / fx 系の
      実装が scripts/ templates/ のどこにも無いことを確認)
- [x] 実績ベース負債コスト: overrides `interest_expense` / `loan_fees`、平均有利子負債は
      `debt_beginning`/`debt_ending` か hist_debt 直近2年平均。C11 を実績値に差し替え、
      Adjustments Log にマージナルコスト注記を併記
- [x] 為替感応度 Table 3: `fx_sensitivity.enabled` でのみ生成。5726手修正版 Sensitivity
      28-40行を参照実装とした
- [x] あわせて標準メモ §1 の「Comps は正常化純利益の参考行」を `normalized_net_income` で実装
- [x] 検収: 合成config 4通り (フラグON/OFF × overrides有無) で発動条件どおり。
      validate_output にチェック17/18 を追加し、改ざん版で FAIL することも確認

## タスク2: 5726 の真の回帰テスト — commit 7460fbc

- [x] 新テンプレで EDINET からゼロ再生成 → `models/5726_DCF_Model_20260826_regen.xlsx`
      (手修正版は上書きせず温存)
- [x] `scripts/diff_models.py` を新規追加し主要46セルを突合 → **differences: 0**
- [x] 全セル走査で見つかった実差2件はテンプレ側を参照実装に合わせた
      (Valuation Range を DCF 2法に / C10 ラベル)
- [x] 期待値の食い違い (Block A 16,811 / B-3 19,643) は**入力差**と特定 ——
      負債コスト修正前 WACC 9.48% の値。閉形式で再現して確認
- [x] 標準メモ §1-2 に回帰テスト完了を追記

## タスク3: todo.md 上書きインシデントの後始末 — 本コミット

- [x] 復元可否の最終判断 → **逐語復元は不能**。確認した経路は本ファイル冒頭の表のとおり。
      唯一未確認は VSS シャドウコピー (管理者権限が必要)。
      再構成できる内容は「### Review 再構成」として書き戻し、逐語部分は
      `lost (2026-08-26 incident)` と明記してクローズ
- [x] `tasks/lessons.md` をコミット (2026-08-23 の 2962 `=` 事故 + 今回の上書き事故)
- [x] 本作業前からの未コミット変更3件を判定 (下記 Review)

## Review — フォローアップ

### 本作業前からの未コミット変更3件の判定

| ファイル | 何の変更か | 判定 |
|---|---|---|
| `templates/sotp_template.py` | `dcf_crosscheck` の `labels` / `source_file` が読まれていたのにマージされず、Cover が汎用の手法名と「from no DCF workbook」を表示していたバグの修正 (+8行) | **コミットする**。extract_dcf_data (L103-104) が生成し L248/L284 が消費するキーで、マージ漏れは明らかなバグ。`templates/test_dcf_crosscheck_matcher.py` 17/17 PASS で回帰なしを確認 |
| `tasks/lessons.md` | 2026-08-23 の 2962 事故 (`=` 始まりのラベルで Excel が開けなくなる) + 今回の上書き事故の再発防止ルール | **コミットする**。どちらも再発防止の資産 |
| `reports/2359_market_analysis_20260509_v2.xlsx` | 新しい market_analysis テンプレでの**再生成物**。`Implied Multiple Analysis` と `Narrative Stage` の2シートが増えている。共有2シートの差分は数式の float 表記のみ (`14366101.0`→`14366101`、`0.10`→`0.1`) で**値の変化なし**。ただし `B3 "Price Data Date: Manual"` が消えている | **判断不能 — 触らず報告のみ**。CLAUDE.md の File Protection Rules は `reports/` を「絶対に上書きしない」と定めており、この上書きが意図的な差し替えなのか runner の事故なのかは外形から判別できない。内容としては B3 を除き上位互換。**ユーザーの判断待ち**: コミットするなら B3 の日付ラベルを復元してから、破棄するなら `git restore reports/2359_market_analysis_20260509_v2.xlsx` |

### 意図的に残す差分

- `reports/2359_market_analysis_20260509_v2.xlsx` (上記のとおりユーザー判断待ち)
- `data/` `models/` `reports/` 配下の未追跡ファイル群 (本作業以前から未追跡。
  今回追跡対象にしたのは回帰テストに必要な `data/overrides/5726_overrides.json` と
  `models/5726_DCF_Model_20260826_regen.xlsx` のみ)

### regen が手修正版の完全な置き換えではない点 (既知・意図的)

`_regen` の Adjustments Log には自動記録2行しか無く、手記入21件は入っていない
(`scripts/fill_adjustments_log.py` を当てていない)。当てなかった理由は、21件のうち3件が
テンプレ修正で**陳腐化**しているため:

1. `DCF Model!C11 = 1.94%(推定・未解決)` → 実績 0.59% に置き換わった
2. `Reverse DCF シートを add_reverse_dcf_sheet.py で追加 / テンプレ標準7シート` → 標準8枚目になった
3. `Target Mid = AVERAGE(C16:C19) のシナリオ依存挙動 (未解決・テンプレ課題)` → e26b801 で解消

`_regen` は回帰テストのベースラインであって納品物ではない。納品物として使うなら、
上記3件を `data/adjustments/5726_adjustments.json` で更新してから
`fill_adjustments_log.py` を当てること。



---

# 3441 山王 DCFモデル生成 (2026-08-29)

出力: `models/3441_DCF_Model_20260829_FINAL.xlsx` (FINAL — 2026-08-29 ユーザー承認済み。validate FAIL 0 / WARN 0 / PASS 19)
手順書: docs/DCFパイプライン標準運用手順書.md (v2) / 標準メモ docs/DCFフォーマット標準メモ_20260826.md (v2)
設定ファイル: `data/overrides/3441_overrides.json` / `data/comps/3441_comps.csv` /
`data/segments/3441_segments.json` / `data/adjustments/3441_adjustments.json`

## 銘柄型判定
**型A: 通常の事業会社 / 傾き実現型(2962型)**。銀行なし・captive financeなし。
貴金属表面処理(金・銀・パラジウムめっき)加工、国内+フィリピン(SPMC)の2セグメント。
粗利率の階段 17.5%(FY23) → 16.4%(FY24) → 20.3%(FY25) → 24.2%(FY26 9M) は既に数字に出ており、
分析の主眼は「現値が傾きの継続をどこまで織り込んだか」の判定 → 逆算DCF(Reverse DCF)が主役。

## タスク
- [x] 手順書v2 / 標準メモ / overrides_schema の読み込み
- [x] EDINET 取得確認 (fiscal_year_end_month=7、有報3期 + 半期報告書 S100XQSL)
- [x] 入力値の相互検算(プロンプト転記値 vs EDINET実績 → 完全一致)
- [x] 株数の確定(自己株控除後 4,241,897株。プロンプト暫定値5,000,000株から変更)
- [x] overrides JSON 作成 + overrides_validator 通過
- [x] comps CSV 作成(ピア5社、ティッカー誤り1件を訂正)
- [x] Segment Bridge 設定(EDINETセグメント注記 2期)
- [x] generate_dcf.py 実行 → recalc → validate (FAIL 0 / WARN 0 / PASS 19)
- [x] Adjustments Log 36件記入
- [x] 5シナリオの単調性を Excel COM で実測確認

## 設計判断(要点。全件は Adjustments Log)
1. **株数 4,241,897株**(自己株758,103株控除)。EDINET有報のBPS×株数=純資産、
   yfinance BPS×株数=純資産、自己株簿価の増分と取得単価、の3本が一致。
   時価総額 11,725百万 / EV 12,849百万(プロンプト前提の 13,820 / 14,755 から −15%)。
2. **基準年 = FY2026/7E**(Q3累計実績 + Q4推定)、予測Y1 = FY2027/7。
   分析日2026-08-29 は FY2026/7期が既に終了(7/31)し本決算未発表という状態。
3. **stub_fraction = 0.92 を明示指定**。テンプレ自動値0.50は半期報告書ラベル由来で、
   既に終了した期の残り半年を指すため誤り(下記 lessons 参照)。
4. **Q4の粗利率は Q3単独実績21.92%**。半期報告書からQ3単独を復元すると
   H1 25.71% → Q3単独 21.92% と粗利率は既に低下している。
5. **リース債務189をEVに算入**(net_debt 1,124)。除外なら935で1株価値差は約44円。
6. **税率25%**(推定)。Q3累計の実効税率8.1%は持続性なし。30.6%ならTarget −7〜8%。
7. **fx_sensitivity は enabled にしない**。比国生産型でUSD連動比率を推定する根拠がない。
8. **Reverse DCF の peak_op = 2,109**(Q3年率換算)にしてプロンプトの問いに直接答える形にした。
   Block E(取引ベンチマーク)は該当事例なしのため生成せず、理由をログに記録。

## Review (2026-08-29)
- validate_output: **FAIL 0 / WARN 0 / SKIP 0 / PASS 19**。
- **Target (DCF Mid) = 3,211円 / BUY / 上値 +16.2%** (PGM 3,189 / Exit 3,232)。
  WACC 11.30% (RF 2.90% / β0.98 / ERP 5.5% / SP 4.0% / Kd_at 1.00% / D/E 0.0959)。
- シナリオ別 Target: Upside 4,933 (+78.5%) / Management 3,782 (+36.8%) / Base 3,211 (+16.2%) /
  Downside 1 1,982 (−28.3%) / Downside 2 1,201 (−56.5%)。Y5営業利益は 4,077 / 3,049 / 2,467 /
  1,476 / 905 で単調。
- **逆算の答え**: 現値2,764円(EV 12,849)は、Q3年率営業利益2,109百万の **84%(即時到達)〜85%
  (5年ランプ)** の定常化を織り込んだ価格。Base の DCF EV 14,653 を **1,804百万(425円/株)下回る**。
- **粗利率20%回帰なら** Y5営業利益1,476百万、株価1,982円(−28.3%)。
- Comps: 自社 EV/EBITDA 5.28x(ピア中央値5.94x) / PER 6.7x(同15.96x) / PBR 1.42x(同0.91x) /
  ROE 21%(同6%)。Target には不算入([参考])。
- **Segment Bridge の発見**: FY2025/7 の利益改善は日本セグメント(−188 → +509)で起きており、
  フィリピン(+348 → +243)ではない。粗利率階段のドライバー候補③「比国子会社の円安効果」は
  主因から外れる。
- 未解決(全て Adjustments Log に記載): 基準年Q4推定 / 税率25% / base_year_ap 500 /
  capex・D&Aの予測配列 / FY2026期のセグメント別開示 / 取引ベンチマーク事例。
- **次アクション**: 9月中旬の FY2026/7 本決算 + FY2027/7 初回ガイダンスで
  (1) Q4実績 (2) Q4単独の粗利率 (3) FY2027会予 (4) 通期セグメント別 (5) 通期実効税率
  が判明する。その時点で再生成すること。

---

# J-Quants V2 移行 (jquants_universe.py) — 2026-08-31

## 背景
J-Quants API は V2 へ移行し、**V1 は 2026-06-01 に終了**。全面 403 の原因はこれ。
`/v1/token/auth_user` `/v1/token/auth_refresh` は廃止、ダッシュボード発行の
API キーを `x-api-key` ヘッダーで送る方式になった。

公式仕様で確認できた V1→V2 の差分（認証以外にもある）:

| 項目 | V1 | V2 |
|---|---|---|
| ベース | `https://api.jquants.com/v1` | `https://api.jquants.com/v2` |
| 上場一覧 | `/listed/info` | `/equities/master` |
| 日次四本値 | `/prices/daily_quotes` | `/equities/bars/daily` |
| 認証 | `Authorization: Bearer <idToken>` | `x-api-key: <APIキー>` |
| 本体 | エンドポイント毎に別名の配列 | `data` 配列で統一 |
| 銘柄名/業種 | `CompanyName` `Sector33CodeName` `MarketCodeName` | `CoName` `S33Nm` `MktNm` + `ProdCat` |
| 四本値 | `Close` `Volume` `TurnoverValue` | `C` `Vo` `Va` |
| 時価総額 | 無し | **`MktCap` (JPY mn) が daily bars に入った** |
| レート | 実質無制限 | Free 5 / Light 60 / Standard 120 / Premium 500 req/min |

## やること
- [x] 公式ドキュメントで V2 仕様を確認（パス・ヘッダー・項目名・ステータス）
- [x] 認証を API キー方式へ: `.env` の `JQUANTS_API_KEY` を読み `x-api-key` で送る
- [x] `Authorization` ヘッダーは同時送信不可 → セッションから明示的に削除
- [x] トークン自動更新・キャッシュ (`jq_refresh_token` / `jq_id_token` / `jquants_token.json`) を撤去
- [x] API キーの空白・改行を除去してから使う
- [x] 400 / 403 / 429 をエラーメッセージで区別してログ出力
- [x] ベース URL と各パスを V2 へ
- [x] レスポンス項目名を V2 へ（これをやらないと build が空回りする）
- [x] `MktCap` を取り込む（V2 で初めて取れるようになった。除外内訳が変わる）
- [x] `ProdCat` で REIT/ETF/出資証券を除外（V2 は市場区分名に商品種別が入らなくなった）
- [x] レートリミット対応: `--rpm`（既定 5 = 無料プラン）と 429 バックオフ
- [x] `--check-auth` で疎通確認
- [x] `--source jquants --build` を実行し、社数と除外内訳を報告

## 想定外だった点（要報告）
1. **認証以外も壊れていた**: パスと項目名が総取っ替え。認証だけ直しても 0 件になる。
2. **`by_market_name` の除外が V2 で機能しなくなる**: V1 の `MarketCodeName` は
   「プライム（内国株式）」「ETF・ETN」のように市場と商品種別の合成だった。
   V2 の `MktNm` は「プライム」等の純粋な市場名で、商品種別は `ProdCat` に分離。
   放置すると REIT/ETF がユニバースに紛れ込む。
3. **`PRO Market` は元から一度も一致していなかった**: 実データは
   「TOKYO PRO MARKET」で、`in` は大文字小文字を区別する。潜在バグ。
4. **`adv20` が price_days 日平均だった**: 既定 40 日なので「20日平均売買代金」に
   なっていない。`liquidity.adv_window_days` を設けて 20 日に限定。

## Review（2026-08-31 完了）

### 結果
- `--check-auth`: **OK**。`/equities/master code=86970` → 日本取引所グループ / プライム /
  その他金融業（データ基準日 **2026-06-08**、無料プランの遅延配信）。
- `--source jquants --build --rpm 5`: **成功**（7分36秒、429 ゼロ、30 リクエスト）。
  上場一覧 4,449 行 / 日次バー 25 営業日 × 約4,449 行 = 4,467 コード。
- **ユニバース 1,078 社** / 除外 3,390 / 未判定 15。

### 除外内訳
| 理由 | 社数 |
|---|---:|
| 時価総額レンジ外 | 1,841 (小 713 / 大 1,128) |
| 流動性不足 (ADV20 < 30百万円) | 723 |
| 市場区分除外 その他(ETF) | 411 |
| 市場区分除外 TOKYO PRO MARKET | 181 |
| 業種除外 銀行業 | 79 |
| 市場区分除外 その他(REIT) | 63 |
| 市場区分除外 その他(外国ETF) | 55 |
| 業種除外 保険業 | 14 |
| 市場区分除外 ETF・ETN（過去取込の綴り） | 11 |
| 市場区分除外 PRO Market（同上） | 8 |
| 市場区分除外 スタンダード(優先出資証券) | 2 |
| コード帯除外 | 2 |
| **未判定**（規模・流動性データ未取得 14 / 時価総額未取得 1） | **15** |

未判定 15 の中身は説明が付く: 2026年の新規上場 8 社（581A〜604A。J-Quants の
基準日 2026-06-08 より後の上場なので無料プランではまだ来ない）、優先株式 2
（JAL 92015 / SoftBank 94346。MktCap の概念が無い）、TDnet 由来の上場廃止銘柄 5。
**「条件を満たさない」と「まだ判定していない」の区別は保たれている。**

### 妥当性検証
- ユニバース内の実測レンジ: 時価総額 5,003〜59,765 百万円 / ADV20 30.0〜21,280.7 百万円。
  ルール（5,000〜60,000 / ≥30）と完全に整合。
- 7203 トヨタ = 45,015,714 百万円 → 時価総額レンジ外。`MktCap` の単位が
  **百万円** であることの裏取り。
- 8306 三菱UFJ → 業種除外(銀行業)。5246 ELEMENTS / 3441 山王 / 4192 スパイダープラス
  / 2359 コア はいずれも universe_flag=1。
- ユニットテスト 16 件 OK。

### 途中で見つけて直したこと
1. **レート設定**: 13.2秒間隔(4.5 req/min)では1分窓の境界で必ず 429 になり、
   429 ごとに 65 秒失って実効 2.3 req/min まで落ちていた。マージンを +30%
   (15.6秒)にし、429 を見たら間隔を 1.3 倍する自動減速を入れた。→ 429 ゼロ。
2. **`--price-days` 40 → 25**: ADV の窓は 20 日なので、40 日ぶん取っても
   新しい 20 日しか使っておらず 15 リクエストが無駄撃ちだった。結果は同一。
3. **`PRO Market` の取りこぼし**: 大小が合わず 8 社が「除外」ではなく
   「未判定」に化けていた。市場区分の照合を大文字化して両綴りに当てる。
