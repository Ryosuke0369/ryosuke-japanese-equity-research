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

