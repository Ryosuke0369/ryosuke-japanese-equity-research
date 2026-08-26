# 5726 大阪チタニウムテクノロジーズ DCFモデル生成 (2026-08-26)

> ⚠️ 復元メモ: 2026-08-26 のテンプレート修正作業の冒頭で、このファイルの未コミット版を
> 上書きしてしまった。以下は上書き直前に読み取れていた範囲の復元。末尾にあった
> 「## Review (2026-08-26)」本文は復元できていない（実質的な内容は xlsx の
> Adjustments Log と docs/DCFフォーマット標準メモ_20260826.md に残っている）。

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

- [ ] 7. `add_segment_bridge.py` を汎用化 (銘柄固有データは `data/segments/<ticker>_segments.json`)
- [ ] 8. `fill_adjustments_log_5726.py` を汎用 `fill_adjustments_log.py` + JSON 化して削除
- [ ] 9. 銘柄コード入りスクリプトの新規作成を機械的に禁止するガードを追加
- [ ] 10. CLAUDE.md / 手順書v2 に運用ルールを明記

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
