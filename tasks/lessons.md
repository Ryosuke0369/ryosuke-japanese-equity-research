# Lessons

## 位置ベースの配列結合は「静かな年度ズレ」を生む (2026-07-31)
**Mistake (構造):** EDINET 由来の OCF/現金/有利子負債を `hist_years` に**位置で**
コピーしていた。overrides が `hist_years` を丸ごと差し替える（年数もラベルも違う）
運用と組み合わさり、5銘柄すべてで**最大3年ぶんズレた値**が別年度の見出しの下に
座っていた（3687: FY2025/9 欄の OCF 1,488 は実は FY2022/9 の値）。値は実在し、
桁も自然なので**目視では絶対に気づかない**。

**Rule for myself:**
- 2つの系列を結合するときは**必ずキーで突合**する。`zip()`／インデックス一致は
  「両者が同じ順序・同じ長さである」という検証されていない仮定であり、その仮定が
  崩れた瞬間に沈黙して壊れる。
- 突合できなかった要素は**空欄にして警告**する。詰める・ずらす・0で埋めるは禁止
  （「データが無い」と「値が0」は別物）。
- 突合ラダー（完全一致→年+月→年が両側一意）を作ったら、**曖昧な場合は必ず空欄**に
  倒す。1件でも取りこぼしたくない気持ちが誤配を生む。
- 配列長は契約にしてバリデータで**実行前にエラー停止**させる。

## 生成物には「機械検証できる根拠」を同梱する (2026-07-31)
**Mistake (構造):** 「D27 が定数 1 になっている」「Comps 統計に自社が入っている」
「C5 が予測からの逆算」——いずれも**生成後に人が見て気づくしかない**壊れ方だった。
実際に 5銘柄中 3銘柄で数か月間そのまま運用され、生成後パッチスクリプト
（add_bank_valuation.py 等）で個別に手当てされていた。

**Rule for myself:**
- テンプレートを直したら、**同じ不具合を毎回検出する検査**（`validate_output.py`）を
  必ずセットで作る。修正だけでは再発を止められない（次の手修正で戻る）。
- 検査が「生成時の意図」と照合できるよう、**成果物自身に機械可読なメタデータ**を
  埋める（Adjustments Log の Pipeline Metadata）。数値から意図を逆算させない。
- 生成後パッチで直している症状があれば、それは**テンプレート側のバグの証拠**。
  パッチを増やさずテンプレートを直す。
- コード変更の影響は **main のワークツリーで同一入力を流した A/B 差分**で示す。
  「旧成果物との差分」は入力差（overrides の編集、株価、四半期進行）が混ざるため
  そのままでは証拠にならない。

## ライブ取得依存は再現性を壊す (2026-06-12)
**Mistake (構造):** comps の時価総額を生成時に yfinance ライブ取得していたため、
同一CSV・同一コードでも実行時刻によって implied 値が±数円変動した（5246 EV/Sales
455→456 等）。「同じ入力 → 同じ出力」が成り立たないと、回帰確認・差分レビュー・
過去成果物の検証がすべて「ドリフトか実変化か」の切り分けから始まることになる。
さらに行単位の「列に値があれば使う、なければライブ」というサイレント混在許容は、
1つの comps セット内で静的値とライブ値が黙って混ざる地雷だった。

**Rule for myself:**
- 最終成果物の入力は**全て静的ファイルに固定**する（comps CSV の Market_Cap 列、
  overrides の current_price）。ライブ取得は「下書き用の利便」であり、使う場合は
  **必ず警告エコー**を出して黙ってライブにしない。
- 静的/ライブの**行単位混在を許さない**: 列があれば全行必須（空欄はエラー停止）、
  なければ列ごと省略。「一部だけ静的」は再現性の観点で最悪（どの行が動くか
  わからない）。
- 静的化した値には **as-of（取得日・出所）を記録**する（CSVにコメントを置けない
  形式なら docs 側に表で持つ）。値の鮮度はデータの一部。
- 再現性の修正は**二連続生成の完全一致**（浮動小数まで）で実証する。
- 上場廃止・TOB銘柄（yfinance 404）はライブ経路では永久に欠損する —
  静的列はフォールバックではなく**必須の正本**になるケースがある（LightWorks 4267）。

## 複数手法平均への無意味値混入 — 「式は正しいのに答えが嘘」型 (2026-06-12)
**Mistake (構造):** Executive Summary の Target Mid = AVERAGE(4手法) が、各手法の
「適用可能性」を確認せずに機械平均していた。赤字企業では PER 法が
「PER中央値 × マイナス純利益」となり、数式としては正しく計算された無意味値
（5246: -534、4192: -11）が平均を汚染（5246: 真値644 → 349 に歪み、推奨も
BUY → SELL に反転していた）。recalc も整合チェックも通る — サイレント
フォールバックと同型の「エラーなく間違う」パターンの集計版。

**Rule for myself:**
- 複数手法を集計（平均・中央値・レンジ）する箇所では、各手法に**前提条件
  （適用可能性ガード）**を必ず定義する: 倍率法は「中央値倍率 × 対象指標」の
  対象指標が正（PER→純利益>0、EV/EBITDA→EBITDA>0）であること。
- ガードに引っかかった手法は **テキスト "N/A" で出力**する（AVERAGE/MIN/MAX は
  テキストを無視するので、集計式は据え置きで有効手法のみの平均になる）。
  `=NA()`（エラー値）は集計ごと壊すので使わない。
- **黙って除外しない**: 除外が発生したら、その旨を成果物上に注記セルで明示する
  （Exec Summary B20）。下流の `% vs price` セルは `IF(ISNUMBER(...))` で防護。
- 修正の検証は3方向: ①対象銘柄（赤字）で除外＋注記が出る、②黒字銘柄で
  従来どおり含まれ値が不変（2359で確認）、③下流消費者（market_analysis の
  `_num('C28')` は非数値→None で安全）への影響を grep して確認する。
- 「モデルが黒字/赤字どちらか」は推測せず実セル（Comps!C22）で確定する —
  4192 は黒字と思われていたが実際は赤字（NI -17mn）で、回帰前提が崩れていた。

## サイレントフォールバック禁止 — 「完走 ≠ 反映」事故 (2026-06-12)
**Mistake (構造):** パイプライン全体に「黙ってデフォルトに落ちる」点が多数あった:
(1) overrides の未知キー・ネスト構造・独自シナリオ名は無検出で消える、
(2) comps CSV 不在は print のみで comps=[] のまま完走、
(3) market_analysis の `float(C26 or 0)` は未recalcモデルで WACC=0 のまま完走、
(4) `.txt` 拡張子の comps ファイルは契約パス `.csv` にヒットせず黙殺。
結果、「実行完走・recalcエラーゼロ」なのに 5246 モデルに 4192 の comps が混入し、
旧形式 overrides の WACC が無視される事故が起きた。**投資判断に使う数値で最悪のパターン
は「エラーで止まる」ではなく「エラーなく間違う」。**

**Rule for myself:**
- 入力契約（キー名・パス・拡張子・構造）は**実行前にバリデートしてエラー停止**する。
  デフォルトへのフォールバックは明示フラグ（--allow-unconfirmed / --no-comps）でのみ許す。
- `x or 0` / `dict.get(k, default)` を外部入力・数式キャッシュに使うときは、
  「None がここに来るのはどういう状態か」を必ず問う。未recalc・契約違反なら raise。
- 生成物の重要値（WACC入力・comps名）は完了時にエコーして目視照合可能にする。
- 「過去の数値が壊れていたか」は推測せず、**生成済み xlsx の実セルを読んで事実確定**する
  （今回: 4192 は正しかった、5246 の comps/terminal/exit は誤りだった、と確定できた）。
- 契約は docs/overrides_schema.md が正本。新銘柄の overrides を推測で書かない。

**副教訓:** 死にキーに見えても間接消費がある（scenarios の itemized NWC キーは
`nwc_items[*].scenario_key` 経由で消費）。キーを「未使用」と断ずる前に間接参照を grep する。

## EDINET document search — budget vs. multi-season fallback (2026-06-04)
**Mistake:** Added a `MAX_API_CALLS=60` budget to `get_document_ids` to stop the
"hang", but 60 was exhausted during the early filing seasons before reaching the
ticker's actual season. This broke 4192 (a December-FY filer found in season 2),
which had worked before — a regression.

**Root pattern:** A global API-call cap interacts badly with a sequential
multi-season fallback: legitimate tickers that file in a *later* season need to
"waste" calls scanning earlier seasons first. A cap sized for the *common* case
silently breaks the *uncommon-season* case.

**Rule for myself:**
- When adding a safety budget on top of an existing sequential-fallback search,
  size it to cover a FULL scan of all fallback branches (here: 4 seasons × ~7
  years ≈ 310 calls → set 400), not just the expected fast path.
- The real fix for a slow search is to make the *right* branch run first
  (dynamic season from `fiscal_year_end_month`) + short-circuit once any result
  is found — not to clamp the cap aggressively.
- Always run the regression ticker (4192) immediately after touching shared
  search code, before declaring the target ticker fixed.

## Overrides schema must match the template's hard-coded contract
**Observation (5246):** The overrides file used custom scenario names
(Bull/StrongBull/Bear/StrongBear) and a nested `wacc_inputs` block, but
`dcf_comps_template.py` only reads the 5 canonical names
(`Base/Upside/Management/Downside 1/Downside 2`) and top-level WACC keys
(`risk_free`, `beta`, `erp`, ...). Result: only `Base` was applied; the custom
scenarios and WACC inputs were silently ignored.

**Rule:** When an overrides file doesn't visibly take effect, check it against
the template's hard-coded `SCENARIO_NAMES` and the top-level config keys —
mismatched names are dropped without error. Surface these to the user rather
than renaming (which would misrepresent their intent).

---

## 2026-08-23 — 「ラベルを `=` で始めない」を自分で踏んだ(2962)

**何が起きたか**: Adjustments Log の「元の値」列に `='Comps Analysis'!C28 => -69` と書いた。
openpyxl はこれを**数式として保存**し(`<f>'Comps Analysis'!C28 =&gt; -69</f>`)、Excel が
ブックを一切開けなくなった(`Workbooks.Open` が com_error -2146827284)。openpyxl では
読めるため、検証を openpyxl だけで済ませていると**気づけない**。

**なぜ踏んだか**: 手順書 §5-5-1 と §7-1 に明記されている罠なのに、「ラベル」= 表示用の
見出しだけの話だと読んでいた。実際は**セルに入る全ての文字列**が対象。「元の値」「変更前」
のような、数式を記録するための列がいちばん危ない。

**ルール**:
1. openpyxl でセルに文字列を書く箇所は、**書き込みループ内で `assert not str(v).startswith('=')`**
   を必ず入れる。1〜2セルだけ手で確認するのでは足りない(今回まさに2セルだけ assert していて、
   ループ内の6行×6列は素通りした)。
2. 数式を記録したいときは `formula: ...` / `旧: ...` のように**必ず非 `=` の接頭辞**を付ける。
3. **手修正の直後に Excel COM で開けることを確認する**。openpyxl で読めることは無傷の証明に
   ならない。`recalc_excel_com.py` が com_error で落ちたら、まず疑うのはこの罠。
4. 切り分けは「行単位で消して Excel Open を試す」二分探索が速い。今回は6行→2行→F列と
   3ステップで特定できた。

**手修正の順序**: 生成 → (recalc なしで) openpyxl 手修正 → recalc → validate が安全。
Excel が保存し直したブックを openpyxl で書き戻す往復を減らせる。

---

## 2026-08-26 — 未コミットのファイルを読まずに上書きした (tasks/todo.md)

**何が起きたか**: テンプレート修正作業の冒頭で `tasks/todo.md` に新しい計画を書いた。
このファイルには前セッション(5726)の**未コミットの作業ログ**が入っていて、`git diff` の
統計 (329行変更) は見ていたのに中身は先頭40行しか読まずに全文を上書きした。
末尾の「Review (2026-08-26)」本文は git にも無く、復元できなかった。

**なぜ踏んだか**: 「todo.md はタスクごとに書き換えるファイル」という思い込み。
`git status` で ` M tasks/todo.md` を見た時点で、HEAD にも無い内容が working tree に
だけ存在することは確定していた。

**ルール**:
1. **上書きの前に `git status` を見る。` M` が付いていたら working tree の内容が唯一の
   コピーである** — HEAD に戻しても復元できない。読むか、退避するか、追記にする。
2. 既存ファイルへの Write は、全文を読んでいないなら**追記**にする。特に
   `tasks/todo.md` `tasks/lessons.md` のような「積み上げ型」のファイル。
3. 作業前に `git stash create` か `cp` で退避すれば1コマンドで守れる。



---

## 決算期末の直後に評価すると stub_fraction が二重にズレる (2026-08-29, 3441)

**何が起きたか**: 3441 山王は**7月期決算**で、分析日は 2026-08-29 —— FY2026/7期は
7/31 に既に終了しているが本決算は9月中旬まで出ない、という「期は終わったが数字は無い」窓。
generate_dcf.py の自動経路は EDINET の最新中間報告(半期報告書 = 2Q)ラベルから
`stub_fraction = (12 - 6)/12 = 0.50` を導く。これは「**当期の残り半年**」の意味だが、
その当期は既に終わっている。放置すると Y1 の FCF が 0.5年で割り引かれ、実際には
11ヶ月先にある FY2027/7 の FCF が半年後に入ってくることになり、**全年度の割引が
半年ぶん過小**になる(Target が数%上振れる)。

**なぜ静かに壊れるか**: 0.50 も 0.92 も「もっともらしい」数字で、Executive Summary の
どこにも矛盾が出ない。validate_output のチェックにも stub の意味論を見るものは無い
(C19 の値が記録と一致するかは見るが、その値が正しいかは見ない)。

**Rule for myself**:
1. 非3月期銘柄では `fiscal_year_end_month` を入れて終わりにしない。**分析日が会計年度の
   どこにあるかを必ず手で数える**。`stub_fraction = (基準年の次のFY期末 − 分析日) / 365`。
2. 「直近本決算FY」と「基準年」を混同しない。**期末は過ぎたが未発表**という窓では、
   基準年は「実績Qの積み上げ + 残Qの推定」で自分で作る(本件は9ヶ月実績 + Q4推定)。
   その場合 `base_year_revenue` / `base_year_cogs` / `hist_years` の最終列 / `ltm_revenue` /
   `projection_start_fy` / `stub_fraction` / `stub_months_elapsed` を**セットで**上書きする
   —— 1つでも自動値のまま残すと基準年の定義が2つ混在する。
3. 推定で作った基準年の列ラベルには **`E` を付ける**(`FY2026/7E`)。Financial Statements は
   `hist_years` をそのまま見出しにするので、実績と推定が同じ見た目で並ぶのを防げる。

## 自己株控除後の株数は「BPS × 株数 = 純資産」で検算できる (2026-08-29, 3441)

**何が起きたか**: プロンプトは「自己株式の株数が未取得。暫定で発行済500万株を使用」と
指定していた。しかし自己株控除後の株数は、開示済みの数字だけから**3本の独立した突合**で
確定できた: (1) EDINET有報の `NetAssetsPerShareSummaryOfBusinessResults`(BPS)× 株数 =
開示純資産 → BPSの分母が自己株控除後であることが確定、(2) yfinance の
`sharesOutstanding` × `bookValue` = 直近BSの純資産、(3) 自己株簿価の期間増分 ÷ 株数増分 =
その期間の株価水準として妥当か。結果 4,241,897株(自己株758,103株)。
発行済 5,000,000株のままだと時価総額が **+17.9%**、EV が +14.8% になり、
このモデルの主役である逆算DCFの答え(「価格は何%の定常化を織り込んだか」)が
84% ではなく 97% になる —— 結論が「まだ織り込まれていない」から「ほぼ織り込み済み」に変わる。

**Rule for myself**:
1. 株数が「暫定」と言われたら、まず **BPS × 株数 = 純資産** の恒等式を試す。
   有報のサマリー(主要な経営指標等の推移)には BPS も純資産も必ずある。
2. `TotalNumberOfSharesHeldTreasurySharesEtc` / `NumberOfIssuedSharesAsOfFiscalYearEnd...`
   は EDINET の XBRL に素で入っている。自己株「簿価」しか無くても、株数タグを直接引ける。
3. 自己株は増えるとは限らない(本件は 674,400 → 656,800 → 758,103 と一度減っている)。
   簿価から株数を割り戻すときは**平均取得単価ではなく期間の増分**で割ること。
4. 株数はモデルの結論に直結する**唯一の分母**。プロンプトの暫定値をそのまま使うのは
   「自己株控除未反映」と注記すれば済む話ではない —— 取れるなら取る。


---

## 数式内の文字列リテラルは255文字まで — 超えるとブックがExcelで開けなくなる (2026-08-29, 278A)

**何が起きたか**: 278A の初回生成で `recalc_excel_com.py` が com_error で落ちた。openpyxl は
問題なく読み書きでき、`validate_output.py` も PASS 14 / FAIL 0 を返す。壊れているのは
**Excelがファイルを開けない**という一点だけで、生成パイプラインのどの検査にも引っかからない。
シート単位の二分探索で `Executive Summary` を特定し、さらに行単位で B24/B25 —— つまり
`investment_thesis` を TEXT 連結数式に変換したセル —— に絞り込んだ。

**原因**: Excel は**数式の中の文字列定数を255文字までしか受け付けない**。最小再現で確認:
リテラル255文字は開ける、256文字は開けない。テンプレート
(`dcf_comps_template.py` の NARRATIVE_TOKEN 処理)は「1セル8,192文字」での分割は実装しているが、
「1リテラル255文字」での分割はしていない。トークン(`{price}` 等)を含む thesis/risks 行の、
**トークンとトークンの間のテキストが255文字を超えた瞬間**にブックが壊れる。
3441 や 5726 で起きなかったのは、たまたま各セグメントが短かったからにすぎない。

**なぜ静かなのか**: 生成は成功する。openpyxl で開ける。validate_output は数式の**文字列**を
見るので構造チェックは全部通る。唯一の症状が Excel COM の com_error で、これは
「Excelが起動していない」「別プロセスがファイルを掴んでいる」といったよくある環境要因と
区別がつかない —— 実際、最初はそちらを疑って Excel プロセスを kill してから再試行した。

**Rule for myself**:
1. thesis / key_risks を書くとき、**トークンで区切られた各セグメントを255文字以内**に収める。
   トークンを含まない行は素のテキストセルになるので上限32,767文字で問題ない。
   書く前に機械的に検算する: `re.split(r'\{[a-z_]+\}', line)[::2]` の各要素の len を見る。
2. `recalc_excel_com.py` が com_error で落ちたら、**環境要因を疑う前に既知の良品を1つ recalc して
   みる**。良品が通ってその銘柄だけ落ちるなら、原因はブックの中身である。
3. 切り分けは**シート単位 → 行単位の二分探索**が速い(openpyxl で他シートを削除して Open を試す)。
   今回は 8シート → Executive Summary → B24:B26 → B24/B25 の3ステップで特定できた。
4. テンプレ修正候補: リテラルを255文字ごとに文字列連結演算子で分割して書き出す。
   後方互換(255文字以下のリテラルは分割不要なので既存銘柄の出力は不変)。

## 生成後スクリプトの実行順序は固定 — add_segment_bridge を後から再実行しない (2026-08-29, 278A)

**何が起きたか**: `data/segments/278A_segments.json` の注記を1行直したかったので
`add_segment_bridge.py` だけを再実行した。その時点で `Adjustments Log` は
`fill_adjustments_log.py` によって38件が記入済みだった。`append_open_items()` は
「列Bが空になる行」まで下に歩いて open_items を書き込む実装なので、記入済みの表を通り過ぎ、
表と Pipeline Metadata バンドの間の**空行に書き込み、その次の行 = `Pipeline Metadata` の
見出し行を上書き**した。結果、validate_output がメタデータを読めず PASS 18 → SKIP 6 / PASS 12 に落ちた。

**Rule for myself**:
1. 生成後スクリプトの順序は **generate_dcf → add_segment_bridge → fill_adjustments_log →
   recalc_excel_com → validate_output** で固定。この順序に「戻る」ときは、
   **途中から再実行せず必ず generate_dcf からやり直す**。
2. 設定ファイル(segments / adjustments の JSON)を直したら、xlsx は**作り直す**。
   xlsx は生成物であって編集対象ではない。`models/` は上書き可という規約はそのためにある。
3. 症状の見分け方: `fill_adjustments_log.py` が
   `Pipeline Metadata band not found - refusing to guess row numbers` を出したら、これが起きている。
   スクリプト側は「推測を拒否して止まる」正しい振る舞いをしているので、素直に再生成する。

## データ保存先を外出しするときは「DBに書かれたパス」まで見る (2026-08-31, screener)

**何が起きたか**: C: の空きが約5GBになったので screener の保存先を D: へ移す作業。
コード側のパスは `common.py` に集約されていたので `DATA_ROOT` を1箇所足すだけ…と思ったが、
`filings` テーブルの `path / pdf_path / xbrl_path` が **リポジトリルート相対**
(`screener\data\raw\tdnet\...`) で保存されていた。ファイルを移した瞬間に
353行すべてが行き先を失う。`os.path.relpath(dest, C.ROOT)` で書き、
`os.path.join(C.ROOT, ...)` で読む対称な実装だったので grep しないと気付けない。

**Rule for myself**:
1. 保存先を外出しする作業では、**コードのパス定数だけでなくDBに永続化されたパスも
   grep する**。`relpath` / `abspath` / `join(.*ROOT` を必ず検索する。
2. 永続化するパスは **データルート相対**にする。リポジトリ相対でも絶対でもない。
   ドライブや母艦が変わっても DB を書き換えずに済むのはこの形だけ。
   (`C.store_path()` / `C.full_path()` を経由し、各モジュールで組み立てない)
3. 移行スクリプトは **dry-run を既定**にし、書き換え後のパスが実在するか全件確認してから
   apply する。1件でも解決できなければ apply を拒否する ——
   壊れたパスで書き換えると「動いているのに中身が空」という一番気付きにくい壊れ方をする。
4. 検証は「DB→ディスク」だけでなく **「ディスク→DB」の孤児チェック**も行う。
   今回は 437ファイル全てが参照済み・孤児0で移行完了を確認できた。

## Git Bash のヒアドキュメントはバックスラッシュを畳む (2026-08-31, 環境)

**何が起きたか**: `cat > f.py <<'PYEOF'` (クォート付き=リテラルのはず) で
`"screener\data\\"` と書いたのに、ファイルには `screener\data\` が落ちた。
Python が `\d` `\x` を unicode escape として解釈して SyntaxError。
README でも `raw\tdnet` の `\t` がタブ文字になって表が壊れた。

**Rule for myself**: **Windows パスやバックスラッシュを含むファイルを書くときは
heredoc を使わず Write/Edit ツールを使う**。heredoc は「バックスラッシュを含まない
テキスト」専用と考える。書いた後は `cat -A` か Read で実際のバイトを確認する。

## 2026-08-31 — ファイルを上書きする前に中身を読む

**やらかし**: `tasks/todo.md` に計画を書くとき、既存の内容を読まずに丸ごと上書きした。
HEAD の内容は `git show HEAD:tasks/todo.md` で復元できたが、**コミットされていなかった
作業コピーの差分は失われた**。

**ルール**: 既存ファイルへの書き込みは、まず読む。読んだうえで
「追記」か「置換」かを決める。`tasks/todo.md` `tasks/lessons.md` のような
累積ログは原則 **追記**。丸ごと上書きしてよいのは自分が同一セッションで
作ったファイルだけ。

## 2026-08-31 — 「認証が通らない」を認証だけの問題だと決めつけない

**やらかし**: J-Quants が全面 403 になったとき、リフレッシュトークンの文字数を数えて
「トークンが不正」と結論づけ、README にもそう書いた。実際は **API の V1 が
2026-06-01 に終了** していた。トークンをいくら取り直しても通らない類のもの。

**ルール**: 4xx が **全エンドポイントで一様に** 出るときは、資格情報より先に
**API のバージョン/提供状況を公式ドキュメントで確認する**。個別エンドポイントだけが
落ちるなら権限・パラメータ、全部落ちるなら世代・契約・提供終了を疑う。

**ついでの教訓**: メジャーバージョン移行では認証だけが変わるとは限らない。
今回はパス・レスポンスのキー・項目名・レートリミットの単位まで変わっていた。
「認証だけ直して build」は必ず空振りする。移行表を先に作ってから手を動かす。
