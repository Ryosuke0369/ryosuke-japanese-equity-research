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

---

# 再開タスク（2026-09-01）— 昨夜のセッション断からの復旧

前提: 実測済み。integrity_check ok / 破損なし / EDINET索引は全平日被覆済みだが
140・150 のみ 2024-01-18 で途切れ / 四半期報告書DLは0件で未開始。

## E. 起動方法の恒久対策（最優先）
- [ ] E1 fetch_runs の running 放置4行 (686/1906/1909/1944) を interrupted に倒す
- [ ] E2 screener/run_prices.ps1 を新設（スリープ抑止＋ログ、run_edinet_full.ps1 に倣う）
- [ ] E3 screener/start_detached.ps1 を新設（ワンショットのタスクスケジューラ登録で
        親セッションから完全に切り離す。Start-Process はフォールバック）
- [ ] E4 切り離しの実証（親を殺してもジョブが生き残ることを確認）

## B. 株価・全項目版（単独実行、C/Dと並走させない）
- [ ] B1 初日400で死ぬ事故の恒久対策を実装
        400 の message が明示する契約範囲を読み、開始日を繰り上げて継続する。
        範囲を読めないときだけ従来どおり中断（推測で埋めない方針は維持）
- [ ] B2 単体テストを追加（範囲パース／繰り上げ継続／読めない時は中断）
- [ ] B3 --from 2021-09-01 で切り離し起動。約1,305営業日 / 実測2.6秒per日 ≒ 55-60分
- [ ] B4 完了検証: adj_close NULL が 2021-09-01 以降で 0 になること

## C. EDINET索引の補完（B完了後）
- [ ] C1 2024年2月の提出集中日を数日だけ叩き、140/150 が 2024-01-18 以降に
        実在するかを確認して報告
- [ ] C2 実在するなら --from 2024-01-19 --to 2024-12-31 で本走。
        実在しないならそこで停止して報告

## D. 四半期報告書ダウンロード（C完了後）
- [ ] D1 ユニバース内の未取得 140/150 のみ（120/130は対象外）
- [ ] D2 完了検証: 140/150 の xbrl_ok=1 件数とディスク上のzip実数の一致

## Review
（実施後に追記）

## Review (2026-09-01) — セッション断からの復旧

全項目完了。実測は以下。

### 成果
- **prices**: 279,015 行 / 235営業日 → **1,533,219 行 / 1,224営業日**。取得した1,222営業日で
  adj_close/OHLC の欠けゼロ（残NULLは窓外の2021-08-31と土曜日付の旧行 2026-06-06 のみ）。
  未取得83日はすべて市場休日。TOPIXの日数と一致。
- **EDINET索引**: 50,370 → **55,672 件**。140 は 2024-01-17 → 2024-10-10、
  150 は 2024-01-16 → **2026-08-27** まで到達し、他種別と終端が揃った。
- **四半期報告書DL**: 0 件 → **8,406 件**（本走8,336 + 追加パス70）。失敗0。
  DBの xbrl_ok=1 と raw/edinet の zip 実数がともに **15,867 で完全一致**、0バイトファイルなし。
  ユニバース内の未取得 140/150 は **0 件**。

### 入れた仕組み
- `start_detached.ps1`: ワンショットのタスク登録で親セッションから切り離す。
  実行中の同名タスクを踏み潰さないガード付き（Dを殺す事故を未然に防いだ）。
- `run_prices.ps1`: 株価用ランナー（スリープ抑止＋ログ一元化）。
- `covered_range()`: 400 が明示する契約範囲を読んで開始日を繰り上げる。
  5年ローリング窓が日付を跨いで動く事故の恒久対策。
- `index_day` を日単位で握る: 一過性の例外1件で数百日のスイープが消えないようにした。
- `--subtypes`: ダウンロードを書類種別で絞る。
- ランナー2本の stderr 握り潰し（PS5.1 の NativeCommandError）を解消。

### 判明した運用上の事実
- EDINET の書類一覧APIは、同一IPからダウンロードと並走させると **1.2秒/日 → 14秒/日** に落ちる。
  ダウンロード側は影響を受けない。昨夜の「15秒/日」の正体はこれ。
- 四半期報告書の平均サイズは約113KB で、有報(520KB)より大幅に小さい。
  8,336件で 0.9GB（当初見積り4.2GBは有報の平均を当てた誤り）。

### 残件（未着手・要判断）
- prices の 2026-06-06 は土曜日付に4,229銘柄の行がある不審データ。旧日次の誤挿入と思われる。
- ユニバース内の未取得: 120 が 2,338 件 / 130 が 465 件 / 160 が 157 件。今回の対象外。
- financials_q / scores / signals が 0 行のまま（下流未構築）。

## 2026-09-02 span-matched スコアラー移行（S5 → S1/S2/S4）

- [x] `screener/signals/span_scorers.py` — S1/S2/S4/S5 の span-matched 版
      （閾値・係数は別枠から写し。span の扱いのみ変更）
- [x] `screener/signals/span_runner.py` — 別枠と span 版の合流。policy は
      `span_only`（既定・承認どおり）/ `prefer_span`
- [x] `weekly_screen.py` に接続（`--policy`、発火行に [span-matched: ...] を表示）
- [x] `screener/tests/test_span_scorers.py` 14件（既存と合わせ 124件 OK）
- [x] switch_point の穴を修正（span=1 が1本も無い 774 銘柄が永久に不可視だった）
- [x] 投影層の item_key 別名を修正（S1/S3/S6 が全銘柄 available=False だった）
- [x] materialize が earnings_calendar を再構築するよう修正
- [ ] **要判断**: 規則A（5期連続で切替）→ 規則B（期ごとに最細粒度）
      FY2025以降の比較可能ペア 2,954 → 3,851。事前登録の変更が必要
- [ ] v4 として合格基準を事前登録（実装完了後）
- [ ] v3 シャドウに span-matched を載せて前向き検証
- [ ] `forecast()` の span-matched 化（B3 再検証はこの後）

## 2026-09-02 規則B採用 → 完成条件2件（製品定義: スクリーニング自動化）

- [x] 1. 規則B（期ごとに最細粒度・常に同span同士）を既定化
      - span_matched_design.md §9 に変更記録（変更日・理由・実測値・規則A不発）
      - span_runner の既定 policy を prefer_span に
      - evidence 先頭に [期 vs 前年期 span=N/mode] を一律付与
      - 週次レポートに「比較粒度」行を追加
      - 傾き系の同一span連鎖をテストで担保（S4 の期間重複バグを検出・修正）
      - テスト 130件 OK（span_scorers 20件を含む）
- [x] 2. materialize 事後条件ゲート（カレンダー/available率/候補数/例外/表の存在）
      違反で exit 1。現行DBは0件通過、旧DBで2件検出することを確認
- [x] 記録: calibration_backlog §12(縮退セット) §13(偽S5) §14(保留) §15(ゲート)
- [ ] 保留（バックログ §14）: v4事前登録 / v3シャドウ搭載 /
      forecast() span-matched化 / B3再検証。paper・shadow は裏で継続

## 2026-09-02 原文リンクの健全性検査

- [x] 形式検証を weekly_screen に組み込み（既定・毎回）
- [x] 生存確認 --verify-links（明示時のみ外部通信）
- [x] 照合キーを銘柄コード→提出者名に変更（3070 訂正有報の誤検知を修正）
- [x] CSV に doc_company 列を追加
- [x] 2026-09-02 実測: 全9行 形式OK・生存OK。6336 は切れていない
- [ ] 弱点: _latest_doc は最新開示に飛ぶ。根拠期の書類とは限らない（要設計）

## 2026-09-02 原文リンクの根拠期対応

- [x] 投影層 filings に doc_id 列（35,002行すべて充填）
- [x] span_scorers の結果に doc_id/doc_date/doc_source/doc_path/period_note
- [x] S4/S5（複数期）は最新根拠期にリンク＋含まれる期を period_note に明記
- [x] 週次レポートを「原文(根拠期)」表示に、シグナル別リンク行を追加
- [x] CSV に doc_kind / doc_period / doc_note / evidence_docs 列
- [x] --verify-doc-periods（本体 financials_cum で独立照合）
- [x] 検証: 週次9銘柄 一致15/不一致0、カレンダー28銘柄 一致44/不一致0
- [x] テスト132件 OK（根拠期リンクの不変条件2件を追加）
- [ ] 未着手: 前年の同一spanを集約して作る（3565 の根拠が2年古い件。要設計）

## 2026-09-02 週次スクリーン品質改善 5件(+1)

- [x] T1 鮮度フラグ（表示のみ・スコア不変）
      screener/signals/freshness.py / stale_flag・stale_lag・stale_detail 列
      fixture: 3565 basis=FY2025-Q2 latest=FY2026-Q2 → stale=1 lag=4 / 1433 → stale=0
      シャドウ計測: ペーパー22.0% / シャドウ23.1% / 今週22.2% が古い証拠に依存
- [x] T2 開示信頼性フィルタ（移行009）
      security_flags(人が管理・除外) + disclosure_flags(自動検知・表示のみ)
      4813 を主出力から除外し監査リストへ。**指定日と一次ソースURLは未確認＝要手入力**
- [x] T3 S12 マクロ定型文除外（移行010・記録先行→適用）
      段落分類 macro/industry/company。**Tier C だけでなく A/B にも適用**（症例の
      9692 cost_pressure は Tier A のため）。7,488ペア再計算の before/after は §18
      → 本適用（s12_narrative --all）は日次パースのロック解放待ちでキュー済み
- [x] T4 会計処理変更の検知（表示のみ）
      注記2系統。素朴なキーワードは42%陽性で使い物にならず17%まで絞った
      本文取得を full_body に変更（3565有報 17,181字→111,087字）
- [x] T5 S13 受注残（シャドウ・0点）
      fixture: ベステラQ1の4数値と +49.8% を再現。有報/短信/セグメント表の3形式
      → 全件パースは日次パースのロック解放待ちでキュー済み（業種別可用率は解放後）
- [x] T6 テールリスク開示フラグ（任意）
      実装済みだが **全59,979書類でヒット0件**。TDnet は決算短信・予想修正・
      配当修正・説明資料しか収集しておらず、非決算の適時開示が入っていないため
      構造的に発火しない。取り込み範囲の変更が要る（今回のスコープ外）
- [x] 事後条件ゲート拡張（disclosure_flags / s13_orders / s12 段落種別）
- [x] テスト 132 → **153件** OK
- [x] 記録: calibration_backlog §18〜§23
- [ ] 解放待ちジョブ: disclosure_flags → s13_orders → s12_narrative --all（直列）
- [ ] 要手入力: security_flags の 4813 指定日・一次ソースURL
- [ ] 取り込みの穴: TDnet アーカイブが 2026-07-23 開始。2026年6月開示のQ1短信が
      欠けており、これが stale 4期集中の主因

## 2026-09-02 フォローアップ: 取り込みの穴の修復 + §17

- [x] A1 段階別切り分け(1433/3565/4813): **チェーンは切れていなかった**。
      FY2026-Q4 は投影層まで届いており valid_flag=0（連結範囲変更）で無効化。
      FY2027-Q1 は TDnet 保持(約1か月)切れで遡及取得不能（一次ソースで404確認）
- [x] A2 決算期月別ヒストグラム: 全体は中央26日/p90 35日で健全。
      遅れは 1月(132日)/7月(173日)/8月(141日)/2月(99日)/10月(82日) に集中
- [x] A3 stale_audit 再実行 → §E の効果として下記に記載
- [x] A4 sanityゲートに鮮度条件（決算シーズンに全体で30日以上未更新→exit 1）
      期末月別の遅れは落とさずに毎回表示
- [x] B  件数見積もり: 全開示 222件/日、非決算84.6%(年約46,000件)、
      リスク語ヒット0.20%(年約108件) → **本文は落とさずタイトルのみ記録**する
      方針で screener/fetch/tdnet_titles.py を実装（disclosure_titles 表）
- [x] C  security_flags 投入（移行011）: 4813 special_alert(2025-08-27, JPX URL付)
      + listing_maintenance(2026-04-30, 除外せず表示のみ)
- [x] D  lag 定義を明文化（期インデックス差・四半期単位）。
      FY2025-Q2→FY2027-Q1 は **7**（当初fixtureの4は算術誤り）。テストで固定
- [x] E  前年H1合成（Q1+Q2→span=2）。実装中に「行はあるが値がNULL」の
      前年行が合成到達を阻んでいた欠陥も発見・修正（find_yoy_peer に item）
      効果: 週次 22.2%→**0%** / ペーパー 22.0%→12.8% / シャドウ 23.1%→13.5%
      残る遅れは lag=4→**lag=2** に短縮
- [x] テスト 153 → **158件** OK
- [x] 記録: calibration_backlog §24〜§28
- [ ] 移行010/011 と tdnet_titles 初回取得はロック解放待ち
- [ ] 未解決: 会計処理変更の注記由来検知(a)が0件。短信本文を読めるまで
      構造的に発火しない（(b)の単独運用と理解する）
- [ ] 未解決: 下期(H2)の32%が連結範囲変更で無効化される非対称。最新期が
      落ちやすく stale を生む。ガード自体は正しいので設計判断が要る

### 2026-09-02 フォローアップ 完了分（キュー解放後）
- [x] 移行010（paragraph_class）／security_flags 実データ投入（4813 の指定日・URL）
- [x] listing_maintenance は除外せず表示のみ（除外は種別で判定）
- [x] disclosure_flags 全件: 4,989走査/1,643記録、継続企業50件(3.0%)、
      会計処理変更474件(28.8%、**注記由来0件**)
- [x] s13_orders 全件: 1,622走査/472件(29.1%)。業種別は受注生産型に集中
      （自動車61%/建設56%/電機50%/機械48% ↔ 小売8%/医薬7%/金融0%）
- [x] S13 リードラグ評価 → **接続しない**判断（+30%超は有望だが n=30、
      下側が全体と区別できない）
- [x] S12 マクロ除外を本適用。**孤児の根拠787行**を発見・修正し、
      シャドウ予測と完全一致を確認
- [x] tdnet_titles 初回取得 14営業日/2,152件/リスク語5件
- [x] 事後条件ゲート 違反0件で通過。テスト **162件** OK

---

# 1433 ベステラ DCFモデル生成 (2026-09-02) — DRAFT

## 計画
- [x] 手順書v2 / フォーマット標準メモv2 / overrides_schema.md を読む
- [x] EDINET から FY2022/1〜FY2026/1 の5期を取得(有報5本)
- [x] 自己株控除後の株数を確定(3経路で検算)
- [x] 株価の一次確認(プロンプト1,433円 vs 実測)→ ユーザー確認で 1,269円 採用
- [x] 実績Kd・税率・BS明細を有報XBRLで裏取り
- [x] ピア6社の上場状態・財務・時価総額を確認
- [ ] data/overrides/1433_overrides.json 作成
- [ ] data/comps/1433_comps.csv 作成
- [ ] data/adjustments/1433_adjustments.json 作成
- [ ] generate_dcf.py 実行 → recalc → validate
- [ ] fill_adjustments_log.py → recalc → validate 再実行
- [ ] 目視確認9項目 + 最終レポート16項目

## Review
(生成後に追記)

## Review (2026-09-02 分析基準日 / 生成 2026-09-03)

納品: `models/1433_DCF_Model_20260902.xlsx`(DRAFT)。validate_output: **FAIL 0 / WARN 0 / SKIP 0 / PASS 19**。

**結論**: Target(DCF Mid)= **1,195円**(PGM 1,163 / Exit 1,226)、現値1,269円に対し **-5.8%**、判定 **SELL**。
逆算: 現値は5年で定常営業利益 **1,664百万**(会社予想1,000の1.66倍)を織り込む。
最大リスク: 労働災害の再発による工事中断・指名停止(Downside 2 = 259円)。

**プロンプトから変更した5点**(いずれも根拠つきでAdjustments Logに記録):
1. 株価 1,433 → 1,269円(実測終値。1,433は実測レンジ外・ユーザー確認済み)
2. 株数 9,297,200 → 8,860,910株(自己株436,290控除。3経路で検算一致)
3. de_ratio 0.060 → 0.0762(スキーマの de_ratio は D/E であって D/(D+E) ではない)
4. ltm_revenue を明示指定(1月期はQ1のXBRLがEDINETに無く自動構築が働かない)
5. 投資テーゼ1・2を短縮(Excel の数式内文字列リテラル255字制限。原文だとブックが開けない)

**テンプレート改善候補2件**(手順書v2 §6-6):
- thesis/risks の分割ロジックが 8,192字基準で、Excel の実制約(リテラル255字)を検出しない
- comps CSV に `Exclude_From_Stats` 列が無く、live銘柄の「残置＋統計除外(東邦方式)」ができない

**賞味期限**: 2026-09-09 のFY2027/1中間期決算以降は再計算が必要。

---

# フェーズ2 — パイプライン一括修正と全85件再生成 (2026-09-06)

上位文書: `フェーズ2_パイプライン修正と全件再生成プロンプト`（§1 修正11項目 / §2 新βルール /
§4 全件再生成 / §5 再走査・反転再計測 / §6 キュー20件 / §7 最終レポート v2）。
修正の正本記録: `docs/phase2_pipeline_fixes_20260906.md`。

## 原則
- **1修正 = 1検証 = 1commit**。各修正の後に 5726 回帰（`scripts/diff_models.py`、主要46セル）。
- 基準ワークブック: `models/5726_DCF_Model_20260826.xlsx`（recalc 済み）。
- ベースライン測定済み（2026-09-06）: 未修正コードで再生成した `5726_DCF_Model_20260906.xlsx`
  との差分 **0/46**。回帰ハーネスは決定論的に動く。
- `_FINAL` を作らない / `reports/` を触らない / 推測埋めをしない。

## §1 修正11項目（実施順 — 依存順に並べ替え済み）

- [x] #3  `--force` なしスキップの終了コード（generate_dcf.py）
- [x] #4  `--date YYYYMMDD` のネイティブ対応（generate_dcf.py）
- [x] #8  validate: `SKIP > 0` を PASS にしない（validate_output.py）
- [x] #1  `_val(..., default=0)` → 欠損は None（generate_dcf.py）
- [x] #5  validate: `core_ebitda > 0` と Comps 参考株価の sanity band
- [x] #2  市場データのサイレント・プレースホルダをハードエラー化
- [x] #6  新βルール（Blume 調整既定化・クランプ域 [0.3, 2.0]・置換は WARN）
- [x] #10 `hist_capex` 指定で C5 の根拠が変わる問題（優先順位の統一）
- [x] #9  guidance（会社予想）取得の修復
- [x] #11 非3月期の FY 末月の自動判定（EDINET 探索窓）
- [x] #7  EDINET 探索窓が最新有報を取りこぼす問題

## §2 β再導出（全85件の overrides 更新）
- [x] 旧ルールでクランプ後の値が overrides に書かれているため、raw β を全件再導出して差し替え
- [x] Adjustments Log に raw / adjusted / 採用値 の3点を記録

## §4 全件再生成（85件）
- [x] `TARGET_DATE=20260906` に統一、`--date 20260906`
- [x] state_part1〜4 を統合し phase2 として85件全件を再処理（done スキップ禁止）
- [x] 3ゲート（validate / core_ebitda / market_data）+ assert_recalculated + stale
- [x] 最初の数件で #1 の効果（補完なしで `#DIV/0!` が出ない）を確認

## §5 再走査・反転再計測
- [x] §X / §AD / §Y の機械再走査
- [x] 新旧 Target 対比表・判定反転の集計と主因特定

## §6 キュー20件の再スクリーン
- [x] ブロッカーが消えた可能性のあるものだけ再挑戦

## §7 最終レポート v2
- [x] `batch/batch_report_20260906_v2.md`

## Review (2026-09-06 完了)

**成果物**: `batch/batch_report_20260906_v2.md`(最終レポート v2) /
`docs/phase2_pipeline_fixes_20260906.md`(修正11項目の仕様書兼実施記録) /
`models/<ticker>_DCF_Model_20260906.xlsx` 85件(全件 DRAFT)。

**完了条件**: 3ゲート + freshness + recalculated の5点で **85/85 clean**、
SKIP 0 / stale 0 / 数式エラー 0 / 中点平均 Target 0 / 本バッチ105銘柄の `_FINAL` 0。
1銘柄あたりの生成時間は 393秒 → 81〜120秒。

**判定**: 反転14件(12件がβ単独)。Ke 恒等式 ΔKe = Δβ×ERP + Δsize_premium が 85/85 で
成立し、β が WACC の唯一の変化要因であることを機械確認した。

**指示から逸脱した1点(明示)**: §2-4 の raw β の出所を yfinance から TOPIX 回帰に変更した。
yfinance の beta フィールドが日本株で NTT −0.165 / 大阪ガス −0.201 / 4205 ちょうど 0.000 /
第一三共・任天堂 null と検証に耐えなかったため。TOPIX 回帰は 5726 の overrides が
文書化している当リポジトリ自身の手法で、手計算 1.553 を 1.547 で再現する。
両方の値を全85銘柄の `_beta_note` に併記し、`--source yfinance` で指示どおりの再現も可能。

**本フェーズで新たに見つけた問題(いずれも修正済み or 登録済み)**:
1. 追補6 §X の脚降格は再生成で必ず消える(後処理のため)。`batch/apply_arbitration.py` で運用。
2. §X の §U(トラフ)判定式「直近OPM ≤ p25」は**あらゆる単調減少で必ず発火**し、
   §Y が「トレンドに平均回帰を当てるのは誤り」と禁じた処理へ送っていた。単調性判定を入れて解消。
   **完了条件「中点平均ゼロ」の機械確認がこれを検出した**(当初4件が中点平均のまま残っていた)。
3. screener の FY 採番が短信タイトルに引きずられる(2871)。#9 の FY ガードが検出。
4. 先行報告の「--force なしで exit 0」は generate_dcf.py ではなく regen.sh 由来だった。

**申し送り**: §F 予防的補完は不要になった(2897/2801 で実測確認)が、
`hist_years` の置換は観測窓の設計判断でもあるため今回は残置した。次バッチで外すこと。

---

# 決算先回りスクリーナー: 偽陽性の恒久対策（2026-09-13）

発端: 9/13 決算窓スキャン（42社・13銘柄発火）の上位2銘柄 3475 / 2776 が偽陽性。
状態ファイルは **リポジトリの** tasks/todo.md・docs/calibration_backlog.md（C:\screener_data 配下には無い）。

## フェーズ0 恒久対策
- [x] CLAUDE.md に「日次収集」「週次スキャン」「決算窓フィルタ」のコマンド・パス・出力・成功条件・落とし穴を明文化
- [x] `screener/report/earnings_window.py`（根拠優先順 1/1b/2/3・休場日・直前四半期の有無で並べ替え）
      + `tests/test_earnings_window.py` 7件。9/13 即興版と 42社・推定日・直前四半期判定が完全一致
- [x] JPX 一覧を `raw/jpx_schedule/`、J-Quants 日次を `cache/jquants_fins_summary/` に固定配置

## フェーズ1 収集の完全性
- [x] tdnet_archiver 後条件ゲート: 総件数=読めた行数 / 対象書類すべて保存 / 対象日終了後の取得 を満たすときだけ ok。
      incomplete / partial / provisional を導入、covered は ok/empty のみ（missing_days 変更）
- [x] `screener/report/tdnet_completeness.py`: 7/23〜9/11 全営業日を一次ソース件数で突合 → 修復
      一致18 / 不一致10（8/05〜8/19）/ 照合不能9（7/23〜8/04 保持期間外）。追加 1,099 ファイルセット。
      8/05 は PDF/XBRL リンクの無い行1件で partial のまま（取得不能）
- [x] EDINET: `edinet_bulk --recent N` 追加、run_daily に組込（既存タスク ScreenerTdnetArchiver が venv で実行）。
      9/01〜9/13 索引 58件・取得 29件・失敗0、解析済み
- [x] disclosure_titles: code_raw を読むよう修正、既存 2,148行を再導出（NULL 0）、9/03〜9/13 1,318件取得、run_daily に組込
- [x] run_daily の EDINET 解析は1回（--all --resume）。xbrl_parser は毎回 約10分の全DB集計がある
- [x] テスト 176件 OK（ゲート5件・titles 2件を追加）

## フェーズ2 根拠健全性の監査（修正前・prefer_span）
- [x] `screener/report/evidence_audit.py`。1,327社: available 3,743 / 根拠期なし 1,295（全て external）/
      根拠期なしで発火 450（342銘柄）/ 照合 一致2,448・不一致0・照合不能1,295 /
      偽陽性ルール R1 450・R2 80・R5 39・R4 29・R3 18、R4∩R5 = 2776/3479/4222/4434

## フェーズ3 スコアリング修正（記録先行: calibration_backlog §31）
- [x] span_runner policy `evidence_strict`: 根拠期なし不採用 / lag0 満額・lag1 ×0.5・lag≥2 不採用 /
      S1S2 売上比≥2.0or≤0.5 不採用 / DSO<1日 不採用。不採用は消さず strict_flags と raw_score を残す
- [x] フィクスチャ `tests/fixtures/false_positive_20260913.json`（3475/2776/5136）+ `test_false_positive_fixtures.py` 7件
- [x] シャドウ比較（同一投影DB）: スコアあり 1,296→1,164、根拠期なし発火 450→0、閾値以上→消えた 120 / 新規 65、
      3475・2776 は無評価、5136 0.575 維持、3441 −0.305（lag1 ×0.5 と S3 除外が相殺し見かけ同値）
- [x] 副作用を記録: 単一シグナル銘柄の ±1.0 張り付き（採用1本 501銘柄、±1.0 204銘柄）
- [x] 決算期末月の誤導出を発見・修正（半期報告書タイトルの半期期間を期末月にしていた。ユニバース22社）+ テスト5件
- [ ] **既定 policy の切替（prefer_span → evidence_strict）はユーザー承認待ち。** フェーズ4/5 は --policy 明示で実行

## フェーズ4 全銘柄の再スコア化（evidence_strict 明示・修復後データ・投影DB 18:55 再生成）
- [x] 出力 `C:\screener_data\weekly_20260913_full_rescored.csv`（1,326行。4813 は security_flags で除外）
- [x] 発火率（分母1,326）: S1 280(21.1%) / S2 189(14.3%) / S4 141(10.6%) / S12 541(40.8%、うち正 389=29.3%・負 212)。S5 472(35.6%)
- [x] スコア付与 1,179 / 閾値0.10以上 501。理由: 証拠不足 678 / 証拠不採用 116 / データ欠損 31
- [x] 上位30 は S5+S12 が大半（S12 加算後に1.0超）。閾値以上501のうち S12以外の発火が1本 289銘柄
- [x] サニティ: 3475・2776 無評価（証拠不採用）/ 3441 −0.175（S1 DSO 60→68日 −0.611 を lag1 で ×0.5、S12 +0.13）/ 5136 0.575 維持
- [x] 旧出力との差分（旧の母集団内）: 9/02 v3 7銘柄→3（消えた 2776/3415/6184/6336）、9/13 決算窓 13→6（消えた 8、新規 3544）
- [x] 途中で直した不具合2件: 根拠期照合が TDnet リンクの doc_id を読めず 1,089件を誤って不一致（→ 一致2,199/不一致0）/
      no_score_reason を S12 加算前に決めていた（6銘柄）。テスト 192件 OK

## フェーズ5 決算窓の再抽出
- [x] 出力 `C:\screener_data\earnings_window_20260913_0930_rescored.csv`（42社、9/18まで31）
- [x] 並べ替え第一基準 prev_quarter_in_db: あり22 / 無し18 / 判定不能2（5903・3271 は根拠が est_date 単独）
- [x] 休場日（平日）: 9/21・9/22（国民の休日）・9/23。JPX 一覧は 9/3 版から更新なし

## 残件
- [ ] **既定 policy を evidence_strict に切り替えるか（ユーザー判断）**。切替箇所は span_runner.DEFAULT_POLICY と weekly_screen の --policy 既定
- [ ] 単一シグナル銘柄の ±1.0 張り付き、S12 加算で 1.0 超（§6-1 系・未対処）
- [ ] 8/05 の TDnet 1件（リンク無し）は取得不能のため partial のまま

## 2026-09-13 追記: 既定 policy 切替・コミット
- [x] ユーザー承認により既定を evidence_strict に切替（span_runner.DEFAULT_POLICY / weekly_screen --policy）
- [x] 測定条件を変えない呼び出しは prefer_span を明示: materialize.sanity_check・候補数集計 / stale_audit
      （paper_weekly は別枠 run_scorers 直呼びで影響なし）
- [x] コミット: ブランチ screener-false-positive-20260913（1e05c1a ほか）。7203 のステージ済みリネームは含めていない
- [ ] push 前確認: 未プッシュ60コミットのうち旧 DCF 作業59コミットに旧PCのローカルパス `<HOME>` が
      756箇所（batch/logs_*.txt・tasks/lessons.md 等）。リモート（公開）には0件 → 公開可否はユーザー判断

---

# シグナル設計の修正4件 + 決算窓の再抽出（2026-09-18）

発端: 実運用2銘柄（3441 山王・6838 多摩川HD）の決算を人手検証し、S1 が
「売上減少に伴う債権減」を改善として加点していたことが判明。ブランチ
`screener-signal-fixes-20260918`（57245ce / 8af8e0a / c467228）。記録は
docs/calibration_backlog.md §32〜§34。

## 実施した4件
- [x] 1. S1 に売上方向ガード（§32）。`screener/signals/sales_direction.py` +
      `span_runner` 規則5。判定は**最新の売上タイル**で、span で割った
      1四半期あたり run-rate。縮小は S1b（0点・表示のみ）に分離（§32-2）
- [ ] **1-b. OR/AND の既定はユーザー判断待ち（§32-3）。** 既定は指示の文言どおり
      `sales_direction.MODE="any"`。`"all"` なら加点側S1 380銘柄のうち189が不採用
      （any は71）。6838 は any では残り（警告つき）、all では落ちる
- [x] 2. S5 を本番経路に実装（§33）。経過四半期比・OP進捗・暗黙の残存四半期利益・
      guidance_dead。**スコアの式は変えていない**（検出・表示・監査列のみ）
- [x] 3. S13 の四半期推移を週次レポートに常時表示（§34）。`s13_series.py`
- [ ] **3-b. 四半期 B/B は現データでは作れない（§34-2）。** s13_orders はほぼ年次(有報)。
      決算説明資料の取り込みが要る（`presentation_materials` は全体87行）
- [x] 4. 累計/単独の監査（§34）。S1・S2 の根拠期 span: 1=2,131 / 2=835 / 3=146 銘柄
      → 31.5% が累計判定。S4・S5 の累計は定義どおりで切替対象外

## 収集・スキャン（2026-09-18）
- [x] run_daily（backfill 14）exit 0。TDnet 9/18分 34件取得、3,734件解析。
      後条件ゲート: 8/19〜9/18 で ok 23 / provisional 1（当日）/ incomplete 0 / partial 0
- [x] quarterly_builder / disclosure_flags / materialize（事後条件すべて満たす）
- [x] JPX 一覧を再取得（kessan08_0904 → 0918。追加24 / 削除0 / 予定日変更4）
- [x] 全銘柄スキャン `C:\screener_data\weekly_20260918_full.csv`（1,326行 / 閾値以上503 /
      根拠期照合 一致2,171・不一致0 / 例外0）
- [x] 決算窓 9/16〜9/30 `C:\screener_data\weekly_20260916_earnings_window_v2.csv`（12行）
- [x] 3441 の実績を `models/3441_DCF_Model_20260829_FINAL_actuals_reconciliation.csv` に記入

## 申し送り
- **単一シグナル銘柄の ±1.0 張り付き（§6-1 / §31）が S5 の本番化で顕在化した。**
  全銘柄上位15のうち11銘柄が「S5 + S12」の2本だけで 1.0 超。6838 も S1 を半減させた
  結果、S5 単独に近い形で 0.78 に残った。分母が採用本数なので、シグナルを1本減らすと
  スコアが**上がる**銘柄がある（シャドウ比較で 6466 0.742→1.0 等、32銘柄が0.2以上変動）
- S5 の `guidance_dead` は全銘柄で17件。**3441 は同じ進捗113%で通期未達に終わっており、
  陳腐化の検出は上振れ方向の予測ではない**（§33 にフィクスチャとして固定）

## 2026-09-18 追記: 7銘柄の発表日照会で見つけた J-Quants キャッシュの穴（§35・修正済み）
- [x] `cache/jquants_fins_summary/` に 2025-11-05〜2026-08-03 の穴があり、前年同期基準（根拠2）が
      成立せず est_date 単独（LOW）に落ちていた。63日ぶんを追加取得
- [x] 効果: 窓 9/18〜12/31 で 根拠3 618件 → **0件** / 候補 1,095 → 1,313 / 直前四半期の判定不能 618 → 1
- [ ] 申し送り: 決算窓を組む前に前年同期（as_of−400〜−330日）のキャッシュ充足を確認する

## 2026-09-18 追記: S5b 分離とシャドウD（§35b / §36・適用済み）
- [x] S5b: `guidance_dead` の S5 は加点しない（evidence_strict 規則6）。数値は 0点・表示のみで残す。
      シャドウ比較: 17銘柄が不採用、閾値0.10 以上→消えた 7 / 新規 0。4839 WOWOW 1.10→無評価、6838 0.78→0.53
- [x] シャドウD: `score_adj = 平均 × n/(n+1)`。paper_weekly の variant='D' に追加（v2/A/B/C には触れない）
- [x] P3 事前登録に3件（P3-1 OR/AND・P3-2 縮小 k=0/1/2・P3-3 S5方向判別）
- [ ] **申し送り: シャドウD は paper_weekly（別枠スコアラー）上で走るため、±1.0 張り付きが
      最も激しい evidence_strict 経路の問題を直接は測れない。** 別枠プールの採用本数は
      n=3 が39% / n=1 は10% にとどまる（8週・938候補）。10枠の顔ぶれが変わったのは 8週中5週。
      evidence_strict でスコアする前向き検証は「スコアリングは共通」（v3-3）を崩すので、
      やるなら新規事前登録が要る

## 2026-09-18 追記: 次の決算窓（2026-10-01〜11-30）の監視候補一覧
- [x] `weekly_screen` に `n_adopted`（採用シグナル本数）と `score_adj_k1`（信頼度縮小後）を追加。
      `n_available` は発火本数（S12込み・score>0）で合成の分母とは別物なので、両方を出す
- [x] 窓 10/01〜11/30: 候補1,313 / 窓内(未発表) 1,204 / 既発表で除外 92
      根拠: JPX一覧 137 / 前年同期実績 1,067 / est_date単独 **0**
- [x] 出力 `C:\screener_data\watchlist_20261001_1130.csv`（1,204行・group列で4分割）
      候補562 / 別表A 採用1本 460 / 別表B 直前四半期なし 116 / 別表C スコア未付与 66
- [ ] **JPX 一覧は 7月期・8月期分しか公表されていない**（kessan07/kessan08）。
      窓の主力である3月期Q2・9月期本決算は未公表。公表され次第 `--fetch-jpx` して
      根拠を 1b に格上げすること（現状は前年同期実績+364日が 1,067件）

---

# ユニバース時価総額上限 1,000億 → 3,000億（2026-09-18〜19）

ユーザー戦略判断（投資対象として）。検証銘柄を入れるための変更ではない（仕様書 §6 の原則は維持）。
下限50億・ADV20 3,000万は据え置き。データ源は J-Quants V2 `MktCap`（スクレイピングなし）。

- [x] 拡大前の保存: `C:\screener_data\snapshots\pre_mktcap3000_20260918\`（weekly_20260918_full_v2.csv ほか・companies_pre.csv・投影DB）
      + 本体DB `backups\screener_pre_mktcap3000_20260918.db`（run_daily 終了後に backup API）
- [x] `universe_rules.yaml` v2（mktcap_max_mn 300000）/ 仕様書 §2 変更履歴 / adapter.py の docstring
- [x] 再build（J-Quants V2, --rpm 60, MktCap は 2026-06-24 時点）: **1,656社**（拡大前 1,327）
      = 新データ×旧上限 1,243（データ更新 −84、主に流動性不足）+ 上限拡大 **+413（1,000〜3,000億帯）**
      帯別 50-600: 996 / 600-1,000: 247 / 1,000-3,000: 413。新規入り 456・脱落 127（流動性不足 106・レンジ外 21）
- [x] 副作用: 検証8銘柄の 3905 / 6855 がユニバース入り（理由ではなく結果。仕様書に記録）
- [x] EDINET 差分: `edinet_bulk --download --codes-file ... --download-since 2023-09-18`（両オプションを追加）
      新規456社・3,655件・失敗0・約1.5GB。既取得29件は再取得なし
- [x] 後段: EDINET解析 → quarterly_builder → S12 --all → disclosure_flags/order_backlog（新規456社）→ materialize（事後条件すべて満たす）
- [x] 出力 `C:\screener_data\weekly_20260919_full_v3.csv`（1,655行 = 1,656 − 4813、`--universe --all-calendar`、v2 と同形式）
      閾値0.10以上 609 / 根拠期照合 一致2,608・不一致0 / 例外0。score_adj_k1 上位30 のうち新規入り 11
- [x] P3-4 事前登録（時価総額帯 50-600 / 600-1000 / 1000-3000、エントリー前営業日の prices.mktcap で判定）
      + `backtest_eval.tag_mktcap_bands` / `band_report`、`backtest_trades_banded.csv`、v2 にも帯別表。テスト `test_mktcap_band.py`
- [ ] **P3-4 の前提: 新規ユニバース入り銘柄の株価履歴（2021-08-31〜）が未取得。** 現状の 1,000-3,000 帯は
      旧ユニバース銘柄の40トレードのみ（<100件）。取得するまで拡大の効果は測れない
- [ ] 25935 伊藤園（優先株式）がユニバースに入っている（拡大前から）。ProdCat が内国株券扱いの疑い。除外ルールの要否はユーザー判断
- 事故記録: 後段スクリプトの待機ループが自分のコマンドラインに一致して 11:52〜12:43 空転 /
  コードリストの CRLF で disclosure_flags・order_backlog が初回0件（再実行済み）

## 2026-09-19 追記: 株価履歴・優先株式除外・コミット
- [x] 新規ユニバース入り456銘柄の株価履歴: `jquants_prices --prices --codes-file ...`（銘柄単位モードを追加。
      日付単位の既存モードは「埋まっている日は飛ばす」ので銘柄の追加では全日スキップになる）
      事前見積り 約10〜12分・約56万行 → 実績 10分・544,719行・457 req・失敗0。Light 5年ローリングで開始は 2021-09-21
- [x] materialize（事後条件すべて満たす）→ `backtest_v2 --quick`。主セル候補 4,591 / v2 建玉 432
      P3-4 成立件数（候補トレード）: 50-600 3,431 / 600-1000 596 / **1000-3000 357**（うち新規入り 278）/ 範囲外 207 / 不明 0
      v2 建玉: 323 / 50 / 23 / 36 → 建玉単位では 600-1000・1000-3000 が100件未満。詳細は backtest_acceptance_criteria P3-4
      旧記録（09-09・1,743件）との差は拡大分だけではない（その後の信号・データ修正を含む）。旧記録は snapshots に退避
- [x] 優先株式等の除外（`exclude_share_classes`、正規化後5桁コード＋名前「優先株式」）。該当7銘柄、ユニバース 1,656 → 1,655
      `weekly_20260919_full_v3.csv` は除外前の出力なので 25935 の1行を含む（読み飛ばす）
- [ ] ±3シフト投影DB は 09-02 生成で新規銘柄なし。P3 頑健性の前に作り直す
- [ ] ユニバース build は全上場銘柄に1日分の prices 行（2026-06-24）を書く（2,441銘柄が1行だけ）。バックテストには無害だが投影DBが膨らむ
- [x] コミット: ブランチ `screener-universe-mktcap3000-20260919`（拡大分のみ。.env.example・DCF関連docs・models は含めない）

## 2026-09-19 追記(2): ±3投影DB再生成・P3-4判定不能の事前登録・build の株価書き込み修正
- [x] **不具合修正**: `jquants_universe.build_from_jquants` が prices に5列だけの INSERT OR REPLACE をしていた。
      (1) 全上場銘柄に1日だけの行（2026-06-24、2,441銘柄）を作り、(2) **既存の同日行の調整後終値・時価総額を
      NULL で上書き**していた（約1,300行）。prices.adv20 を読む箇所は無いので書き込み自体を削除。回帰テスト追加
- [x] データ修復（事前に `backups\screener_pre_prices0624_repair_20260919.db`）:
      実履歴のある1,796銘柄の 06-24 を再取得（1,777行）→ 拡大前バックアップと 1,327行照合・差分は 1447 のみ
      （2026-09-11 の株式分割で調整後終値が遡及変更されたため）→ 1447 は全履歴を銘柄単位で再取得し、
      5年窓外の14行は分割係数で換算して連続にした。捏造行 2,444（1日行 2,441 + 06-24 に取引の無い3銘柄
      1807/7565/9720 に前日終値が入った行）を削除。prices は 1,798銘柄、1行だけの銘柄は既存の2件のみ
- [x] 投影DB 3本（shift 0 / +3 / −3）を再生成。3本とも事後条件を満たす（daily_prices 2,070,052）。旧 ±3 は snapshots に退避
- [x] backtest_v2（3シフト）: 建玉 432 / 420 / 435、3本とも期待値プラス・勝率変動 3.3pt
- [x] P3-4 の成立件数を修復後の値で更新。1000-3000 帯は候補 357（2025年より前 116）/ 建玉 25、600-1000 は建玉 50
- [x] **事前登録に追記**: 判定単位は v2 建玉に固定。1000-3000・600-1000 帯は「判定不能」（記述統計のみ）。
      P3-3 と同じく複数サイクルで累積し、累積建玉100件到達で初めて帯別評価

# テクニカル層 シャドウE フェーズA —— イベント→価格の反応記録（2026-09-19）

ブランチ `screener-technical-event-response-20260919`。記録は calibration_backlog §37、事前登録は
backtest_acceptance_criteria「シャドウE」。**測定のみ。シグナル化・スコア接続はしていない。**

- [x] 事前登録（E-0 設計原則 / E-A 定義）を集計前にコミット（3d7b4c4）
- [x] 9月の株価（1,188銘柄・09-01〜09-18、16,610行・失敗0）と TOPIX（〜09-18）を取得。
      それまで 09-01 以降の株価は新規入り456銘柄にしか無く、TOPIX は 08-31 で止まっていた
- [x] `screener/technical/`（パッケージ docstring に原則）/ `event_response.py` / フィクスチャ / テスト2本
- [x] A-1 補正: 期中レビュー完了の再掲・補足資料・差替・再掲載・データ追加をその他へ（試走で発見、経緯を事前登録に追記）
- [x] 本番実行: 1,789イベント / long 63,708行。独立計算との照合 一致（12件 + 分割9件）
- [ ] **申し送り: 株価の日次更新が run_daily に入っていない。** 9月分が抜けていたのはこのため。
      フェーズBの前向き記録には日次の株価・TOPIX が要る
- [ ] 申し送り: 受注関連は3件。非決算タイトルが 08-20〜09-11 しか無く（その後も日次で増える）、判定不能
- [x] フェーズB: 出口4分岐と受容帯 N=60/B=20/X=70%/M=3 を E-B に提案・事前登録（成績は未計算）
- [ ] **フェーズBの実装・インサンプル符号測定（E0〜E4 × ±3シフト）はユーザー承認待ち**

## 2026-09-20 フェーズB 実装・測定 + フェーズA 追加測定（承認済み）

記録は calibration_backlog §38（追加測定）/ §39（出口4分岐）。事前登録は backtest_acceptance_criteria
E-A2 / B-2b（承認事項・OR採用の理由・veto化の申し送り）。**判定は符号の記録まで。採用していない。**

- [x] 承認事項を事前登録に追記（OR採用の理由3点 / entry_below_val の申し送り / 追加測定2件）
- [x] `screener/technical/acceptance_zone.py`（N=60/B=20/X=70%/M=3）+ `screener/report/exit_sim.py`（E0〜E4・±3シフト）
      テスト `test_exit_sim.py` 14件（受容帯の算術・分岐の優先順・OR/AND・帯がエントリー日以降を見ないこと）
- [x] **照合**: E0 の再計算 vs 別枠の ret 最大差 5.0e-06。最初の実装は価格系列の本数で出口日を数えて
      7.1e-02 ずれた（休場日）。別枠と同じ `jp_calendar.add_business_days` に直した
- [x] 測定: E4 − E0 = −0.0092 / −0.0071 / −0.0067（3本ともマイナス・符号一致）。悪化はほぼ分岐1
- [x] `screener/report/event_strata.py`（E-A2）。発火320 / 非発火548 / スコアなし918
- [x] **株価・TOPIX の日次更新を run_daily に追加**（`jquants_prices --prices --topix --recent 10`）
- [x] **不具合修正**: 日次モードが「行が1つでもある日」を取得済みと見なしていた（部分取得の日を永久にスキップ）。
      `covered_days` で「最大の日の50%以上の行がある日」だけを取得済みにする。回帰テスト追加。
      2026-09-01〜09-18 の欠落はこれが原因
- [ ] **申し送り（ユーザー判断）: veto 化（受容帯の下なら入らない）は、この測定では支持されない。**
      下限割れ入場の E0 期待値は 3本中2本で他の建玉より**良い**（+0.0254 / +0.0155 vs +0.0066 / +0.0038）。
      §39 の表を見て判断すること
- [ ] **申し送り: 業績予想修正の方向が取れない。** `guidance.revision_direction` が全行 `initial`。
      修正開示の XBRL を guidance に取り込むまで、方向別測定も分岐2の第3条件（予想引き下げ）も動かない
- [ ] **申し送り: T+60 は 2026年11月頃に再測定**（定義は B-1 のまま動かさない）
- [ ] 前向きのシャドウE（variant='E'）並走は、本報告への承認を得てから

## 2026-09-20 追記: 取得済み判定の横断点検と、投影DB再生成による再測定

- [x] 「行/ファイルがあれば済み」型の横断点検（calibration_backlog §40）。
      同型は `xbrl_parser --resume`（1行でも解析済み）と `presentation_text`（ファイルがあれば済み）の2件。
      いずれも**現データでは実害0**（1〜4行の filing 0件 / 空テキスト 0件）
- [x] 実害あり: **J-Quants 決算サマリーのキャッシュに平日131日の穴（2026-02〜07）。**
      影響は 2027-02〜07 の決算窓（前年同期の基準）。今の窓には影響しない
- [ ] **要実施: 穴埋め** `earnings_window --fetch-jquants --from 2026-02-01 --to 2026-07-31`（約131リクエスト）
- [ ] 提案: `xbrl_parser --resume` を項目数ベースに、`presentation_text` に最小サイズを入れる（存在ではなく完全性で持つ）
- [x] **再測定**: 初回の出口測定は投影DB（09-19 14:30 生成）が9月株価を含まず。
      投影DB3本を再生成（事後条件すべて満たす）→ exit_sim 再実行。E4−E0 の符号・結論は変わらず。
      スコア層別（§38）は再実行して完全一致、フェーズA（§37）は本体DB直読みで影響なし

## 2026-09-20 追記(2): 承認5件の実施（キャッシュ穴埋め・完全性判定・予想修正の取り込み）

- [x] 1. J-Quants 決算サマリーのキャッシュ穴埋め: 160日ぶん取得（2026-02〜07 の平日131日 + 祝日）。
      2027-02〜07 の決算窓で前年同期基準が引ける
- [x] 2. 「取得済みは存在ではなく完全性で持つ」の徹底（§40）:
      `xbrl_parser --resume` を**項目数5以上**（予想の修正開示は guidance 行があれば済み）に、
      `presentation_text` に **200バイト下限**。テスト `test_resume_completeness.py`
- [x] 3. veto 化は**不採用**（測定が支持していない）。§39 の表を根拠として残す
- [x] 4. 前向きは E4 の形では回さない。**テーゼ破綻の観測ログのみ**（売買を変えない記録専用）。
      事前登録に範囲を追記 / `screener/report/thesis_observe.py` + テスト4件。
      運用: 週次ペーパーの後に `python -m screener.report.thesis_observe` を回す
- [ ] 5. T+60 再測定は 2026年11月（定義は B-1 のまま）
- [x] **最優先: 業績予想修正の XBRL を guidance に取り込み（§41）。** 方向は up 685 / down 240 / flat 73。
      連結優先・レンジ除外・赤字は大小比較。再解析は 280件 + 連結単体併記の 201件
- [x] E-A2-2 を新出典で再測定: 上方61件の初動 +4.78%（日付クラスタ t 3.91）/ 下方11件は件数不足
- [x] 新規事前登録（別件）: **シャドウF**（受容帯を入口の拒否・サイジング F1a/F1b/F1c、
      テーゼ破綻を入口の除外 F2a/F2b）。パラメータは測定前に固定。**測定は未実施**
- [ ] 申し送り: `materialize` が `company_forecasts` に方向を運んでいない（前向きの第3条件で要る）
- [ ] 申し送り: 修正開示 618件のうち XBRL は 210件（34%）。残り 408件は PDF のみで方向不明

## 2026-09-20 追記(3): シャドウG（S5→上方修正）とシャドウF（入口側）の測定

- [x] 1382 型の連結/単体の取り違えは**他にもあった**。修正前後の差分で **48銘柄282行**が変化。
      原文 zip との照合 12/12 一致（前回の不一致 3161 も解消）
- [x] あわせて「中間期だけの修正」が通期予想として保存される問題を修正（`guidance.q_no` + 通期優先）。
      3161 は H1 のみの修正で q_no=2
- [x] materialize が `company_forecasts` に `revision_direction` / `prev_op` を運ぶ（通期のみ）。
      投影DB3本を再生成（事後条件すべて満たす）: up 108 / down 38 / flat 14
- [x] **シャドウG 事前登録 → 測定（§42）**: S5 発火の up率 1.0% 対 ベースレート 1.3%（差 −0.3pt [−0.7,+0.1]）。
      **上回らない。** スコア四分位にも単調性なし。90日窓は満了0件（実際は4〜53日）
- [x] `guidance_dead` の検出漏れを修正（規則6で `strict_flags` に落ちる）。テストに固定
- [x] **シャドウF 測定（§43）**: 3本で符号一致したのは F1c のみで**マイナス**。F1a/F2a/F2b は符号不定。
      F1b は期待値では動かない（MDD は悪化）。見送った候補は「帯の上」+0.0179 =**外すと損**
- [x] `entry_sim.verdict` の表示バグ（差0を「マイナス一致」と書いていた）を修正 + テスト
- [ ] 次の判断: 入口・出口とも受容帯は採用を支持されない。テクニカル層は**観測のみ**に留めるかどうか
- [ ] 申し送り: PDFのみの修正開示 408件は方向不明のまま（保留・ユーザー判断済み）
- [ ] 申し送り: シャドウG は窓が短い。**2026-11 以降に 90日窓が満了してから再測定**（定義は G のまま）

## 2026-09-20 追記(4): 上方修正の分解・S5再測定・株探アーカイブ・設計提案

- [x] **シャドウH（最優先・§44）**: 上方修正の初動を分解。**引け後開示（31件）はギャップ +5.07%（t 4.60）／
      寄付後 +0.00%（t −0.61）** = 初動のすべてがギャップ。値幅なしを除いても同じ（+3.29% / +0.21%）。
      場中29件は寄付後 +4.03% だが**場中開示では始値が開示より前**なので反応分ではない（日中足が要る）
- [x] **シャドウG は修正後データで走っていた**（DBの生成時刻で確認）。修正前DBで再実行すると
      **up 0件・unknown 701**（測定不能）。通期予想の変化が S5 の群に与えた影響は 3,456→3,452 と小さい（§46）
- [x] **選択バイアスの所在を特定（§46）**: S5 評価可能な銘柄は例外なく通期予想がDBにある。
      ベースレート（ユニバース全体）には予想がDBに無い観測（up率 2.2%）が混ざっていた。
      母集団を揃えると 発火 0.98% 対 全体 0.59%（+0.39pt [+0.02,+0.76]）。**ただし銘柄ユニークは 7/10 で薄い**
- [x] **常習性は判定不能（§47）**: 方向つき修正164件、2回以上ある銘柄は4社のみ
- [x] **株探ウォッチの手動アーカイブ（§45）**: `C:\screener_data\kabutan_watch\` + 取り込み + テスト8件。
      自動取得なし・記録専用（スコア側が読まないことをテストで検査）
- [x] シャドウI 提案（開示義務30%ライン・前年同期進捗で季節性を代用）。**3〜5年の季節性は作れない**
      （4年=4社 / 5年=0社）ことを実測して提案に明記。**実装・測定は承認後**
- [x] G-7 に11月再測定の検討案（窓を「次の四半期決算まで」に / 母集団を揃える G2）を記録
- [x] F2a/F2b の MDD が3本とも v2 と同等以下である点を §43 に申し送りとして明記
- [ ] **承認待ち**: シャドウI の実装・測定 / 日中足の取得可否（場中開示の反応を測るなら要る）
