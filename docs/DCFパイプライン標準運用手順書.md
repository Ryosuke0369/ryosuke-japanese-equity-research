# 日本株DCFパイプライン 標準運用手順書(ランブック)

作成: 2026-07-31(v2 — テンプレート恒久修正の適用完了を反映)
対象リポジトリ: `C:\Users\ryosuke0923\ryosuke-japanese-equity-research`
用途: 新しいチャットでClaudeに新規銘柄のDCFモデル作成を依頼するとき、この文書を渡す(またはプロジェクト知識に入れる)。この文書だけで前提説明なしに作業を開始できることを目的とする。

前提バージョン: 本手順書は**テンプレート恒久修正(本編+追補、2026-07-31)適用済み**のコードを前提とする(コミット f61d242 / 28dbb9f / 1e0fecc / 1fef442 以降)。git log で確認できる。

関連文書:
- `引継ぎ_20260730_日本株DCFパイプライン.md` — 5銘柄の実績・銘柄型の設計判断・**8410の正しい結論(§4)**
- `姉妹テンプレート標準運用手順書.md` — market_analysis / SOTP / narrative の続編手順書
- `テンプレート恒久修正_ClaudeCodeプロンプト_20260731.md` / 同追補 — 修正の仕様(参照用)

---

## 0. 新チャット冒頭に貼るテンプレート

```
日本株DCF自動化パイプラインで新規銘柄のモデルを作りたい。
まず「DCFパイプライン標準運用手順書」を読んで手順に従って進めて。

銘柄: <ticker> <社名>
決算期: <例: 3月期 / 2月期 / 9月期>
現在株価: <円>(<日付>時点)
これから決算短信のスクショを渡す。
```

---

## 1. パイプラインの全体像(Claudeが最初に理解すべきこと)

役割分担は次の通り。**この分担を崩さない**こと。

```
[ユーザー] 決算短信のスクショを提供
    ↓
[Claude/チャット] 数値を抽出・検算 → ClaudeCode用プロンプトを作成
    ↓
[ユーザー] ClaudeCodeでプロンプトを実行(必要なら手動修正)
    ↓
[Claude/チャット] 生成されたxlsxを検証 → 誤りを修正 → FINAL版を出力
    ↓
[ユーザー] FINALを models/ に保存(§6-5の保存確認まで)
```

### リポジトリ構造(要点のみ)

| パス | 役割 |
|---|---|
| `scripts/generate_dcf.py` | オーケストレーター。EDINET取得→LTM構築→config組立→overrides深マージ→yfinance→Excel生成→**validate_output.py 自動実行**の8ステップ |
| `templates/dcf_comps_template.py` | 計算エンジン。openpyxlで**生きた数式**をExcelに書く(Pythonが答えを埋めるのではない) |
| `data/overrides/<ticker>_overrides.json` | アナリスト判断の注入層。**モデルの品質はここで決まる** |
| `data/comps/<ticker>_comps.csv` | 類似企業。**このパス・この拡張子のみ読まれる**(.txtは無視される) |
| `scripts/overrides_validator.py` | 実行前の契約チェック(未知キー/ネスト/独自シナリオ名/配列長/`__CONFIRM__`残存/未定義トークン/ターミナルcapex事前警告でエラー・警告) |
| `docs/overrides_schema.md` | **契約の正本。プロンプトと食い違ったらスキーマが勝つ** |
| `scripts/validate_output.py` | 生成後セルフチェック。DCF/market_analysis/SOTPをシート名で自動判別。**FAIL(exit 1)/WARN/PASS**の3段階。結果は `<xlsx名>_validation.txt` にも出る |
| `models/` | 出力先。上書き可 |
| `reports/` | **絶対に上書き・削除しない** |

### overridesの契約(破ってはいけない)

1. シナリオ名は `Base / Upside / Management / Downside 1 / Downside 2` の**5固定**。増減・改名不可
2. WACC関連はトップレベル平坦キー。ネスト不可
3. `capex_method` / `da_method` は `"revenue_pct"`(既定)か `"direct"` のみ。`"fixed"` `"absolute"` は**存在しない**
4. direct方式でも `capex_pct` / `da_pct` はフォールバック用に**必ず残す**
5. `__CONFIRM__` を残すとエラー停止。`--allow-unconfirmed` は検証ラン限定、**最終ランでは禁止**
6. `nwc_method` は `"days"`(既定)か `"revenue_pct"`
7. **非3月期銘柄は `fiscal_year_end_month` 必須**(EDINET探索窓 = 期末月+3)
8. コメントは `_` 接頭辞キーに書く
9. **hist系キー**: `hist_ocf` / `hist_cash` / `hist_debt` / `hist_capex` / `hist_da`(いずれも hist_years と同長の配列、指定があればEDINET値より優先)。EDINETデータは年度キー突合で割り当てられ、一致しない年度は空欄+警告になる(ズレて埋まることはない)
10. **`ltm_revenue`**(任意): LTM自動構築値を上書き。スコープ調整が必要な銘柄(型Cの事業のみ売上等)で使う
11. **thesis/risksの数値トークン**: `{price}` `{target_price}` `{upside_pct}` `{pb}` `{per}` `{wacc}` が使える(TEXT連結数式に変換され株価更新に追随する)。**新規銘柄では数値は必ずトークンで書く**。未定義トークンはエラー停止
12. SOTP用: `da_allocation_intro` / `da_allocation_notes`(按分根拠の説明文。姉妹手順書§5-2参照)

---

## 2. 手順1: 銘柄タイプの判定(最初にやる)

短信スクショを受け取ったら、数値抽出の**前に**、貸借対照表と事業内容から銘柄型を判定する。型によってoverridesの設計とプロンプトの構成が根本的に変わる。

| 型 | 判定基準 | 主な対処 |
|---|---|---|
| **A: 通常の事業会社** | 銀行なし、captive financeなし | 標準パイプラインそのまま。単位(千円/百万円)だけ注意 |
| **B: シクリカル** | 営業利益が数期で符号反転級に振れる(半導体・海運・素材等) | サイクル型シナリオを設計。「Baseは緩やかな減速」の流用禁止。ピーク利益×高Exit倍率の二重計上に注意。size_premiumは時価総額で見直す |
| **C: captive finance持ち製造業** | 有利子負債の大半が金融債権見合い(自動車OEM等) | 事業のみDCF。短信PLが金融分離表示なら事業/金融EBITを分解。net_debtは事業ネットキャッシュのみ。金融事業価値+持分法投資は**1株あたりで別途加算**。Compsは連結倍率×事業EBITDAの基準混在になるため**PER/PBRに切替** |
| **D: 銀行** | 預金が主要負債 | DCF不成立。`net_debt=0`、`nwc_method:"revenue_pct"`+`nwc_pct=0`+`base_year_ar/inv/ap=0`でΔNWC強制ゼロ。EV/EBITDA・EV/Revenueは使わずPER/PBRのみ。comps CSVは全社`Net_Debt=0`。別スクリプトで**DDM+Residual Income**を追加し、これを主手法にする |
| **E: 連結内に銀行を持つ持株会社** | 小売+金融子会社等 | MIをnet_debtに織込み。**`de_ratio`必須明示指定**(自動計算はMIを負債扱いしWACCを不当に下げる)。銀行業の預金・貸出金はnet_debt/NWCに入れない。SOTPを検討(上場子会社の**上場廃止を必ず最新確認**) |

型C/D/Eの詳細な実装例は引継ぎ文書§3(7203/8410/8267)を参照。

---

## 3. 手順2: 入力データの抽出と検算(Claude/チャット)

### 3-1. ユーザーから受け取るもの

- 決算短信のスクショ(PL・BS・CF・セグメント情報・業績予想のページ)
- 現在株価と時点
- (あれば)会社予想、直近の適時開示、Peer候補

### 3-2. Claudeが抽出・確定する項目

1. **単位**: 短信が千円表記か百万円表記か。千円なら百万円への換算表を作る(3687の教訓)
2. **株数**: 発行済株式数 − 自己株式。検算式を明記(3687: 33,635,000 − 1,382,142 = 32,252,858)
3. **決算期末月**: 非3月期なら `fiscal_year_end_month` を確定
4. **LTM売上**: 自動構築式は `LTM = 直近本決算FY実績 − 前年同期累計Q + 当期累計Q`(3成分がコンソールにログされる)。スコープ調整が要る場合のみ `ltm_revenue` を指定
5. **hist_* 系列**: 売上・営業利益に加え、**営業CF・現金・有利子負債・capex・D&A**を年度ラベル付きで抽出し、overridesの `hist_ocf` / `hist_cash` / `hist_debt` / `hist_capex` / `hist_da` に入れる(EDINET欠損年の空欄化を防ぎ、C5/C18の実績比率の元データにもなる)
6. **net_debt**: 型に応じた定義で(§2参照)。構成要素と計算式を明記
7. **WACC入力**: β、size premium(**時価総額で判断。既存モデルの値を流用しない**)、負債コスト、de_ratio(型Eは必須明示)
8. **税率・永久成長率・Exit倍率**: **ターミナル年のcapexはD&Aに収束させる**(g≤1.5%なら比率0.90〜1.15。範囲外はvalidator/validateが警告する。8267のPGMマイナス事故の再発防止)
9. **シナリオ5本**: 各シナリオが含意する営業利益率の検算テーブルを作る
10. **Peer選定**: 各社の**上場状態を最新確認**(生成時にもyfinanceの鮮度チェックが走り、45日超陳腐化Peerは統計から自動除外+警告されるが、プロンプト段階で確認するのが正)。**全社のD&Aを取得**し、取れない社はCSVに入れない(入れてもEBITDA=営業利益の行は統計から自動除外され、有効Peer<3ならEV/EBITDAがINVALID化される)。海外Peerは為替レートを明記し、異常値は一次ソース(SEC/短信)で検証してから採否を決める(285A Micronの教訓)
11. **thesis/risks**: 英語3点ずつ。**数値は必ずトークンで**(契約11)

### 3-3. ClaudeCode用プロンプトの構成(実証済みの型)

```
## 0. 実行方針
  - __CONFIRM__ は一切使わない。全項目に確定値かフォールバックを与える
  - 「取得を試みよ」項目は失敗したら指定フォールバックで続行し、使った値をレポートに記載
  - docs/overrides_schema.md が正本。食い違えばスキーマ優先で報告
## 1. この銘柄の構造上の特殊性(型B〜Eなら必須)
## 2. 前提(ticker、決算期、株価、株数の根拠と検算)
## 3. overrides JSON 仕様(hist_*5系列、NWC、net_debt、WACC、税率、ターミナル、
     capex/D&A、ltm_revenue要否、シナリオ5本+営業利益率検算、
     thesis/risks 英語3点ずつ・数値はトークン)
## 4. comps CSV 仕様(パス契約、必須カラム=Book_Value含む、
     EBITDA=営業利益+D&A全社統一、D&A欠損社は入れない、選定/除外理由)
## 5. 追加シート(D型: DDM+RI / E型: SOTP / セグメント重要: Segment Bridge)
## 6. 実行コマンド + 目視確認項目
## 7. EDINET取得失敗時のフォールバック
## 8. 最終レポート必須項目(15項目前後。validate_output.py の結果全文を含める)
```

設計目標は**「途中でユーザーに確認が来ずに最後まで完走する」**こと。

---

## 4. 手順3: ClaudeCode実行(ユーザー)

1. Claudeが作成したプロンプトをClaudeCodeに渡して実行
2. 実行コマンドの基本形:
   ```bash
   python scripts/generate_dcf.py --ticker <ticker>
   # validate_output.py は自動で走る。FAILなら生成物を残したままエラー終了する
   ```
3. ClaudeCodeの最終レポート(validate結果含む)を**全文チャットのClaudeに貼り戻す**

---

## 5. 手順4: 生成物の検証(Claude/チャット)

### 5-1. validate_output.py の結果レビュー

- **FAIL 0件が前提**。FAILが残っている生成物は検証に進まない(原因を潰して再生成)
- **WARNは1件ずつ判断を記録する**。設計上必ず出る警告と、対処が要る警告を区別する:

| WARN | 意味 | 対応 |
|---|---|---|
| ターミナルcapex/D&A比率が範囲外 | 予測判断がPGMと不整合の可能性 | 意図的ならその根拠を、そうでなければcapex予測を見直す |
| PGM逆算Exit倍率と仮定Exit倍率の乖離>1.8倍 | 2つの手法の世界観が食い違っている | どちらが正しいか判断し、Adjustments Logに記録(285A/7203では既知) |
| Peer鮮度(45日超) | 上場廃止・ティッカー誤りの疑い | Peerを差し替えるか除外を受け入れる |
| EBITDA=EBITのPeer行 | D&A欠損。統計からは自動除外済み | CSVから外すのが本筋 |

### 5-2. 機械検証(数式ダンプ)

validateが構造をカバーするので、ダンプは**数値の妥当性確認**(WACC inputs、net_debt構成、シナリオ値)に使う:

```bash
python3 -c "
import openpyxl, warnings; warnings.filterwarnings('ignore')
wbv = openpyxl.load_workbook('X.xlsx', data_only=True)
wbf = openpyxl.load_workbook('X.xlsx', data_only=False)
ws, wf = wbv['DCF Model'], wbf['DCF Model']
for r in range(1, ws.max_row+1):
    row=[]
    for c in range(1, ws.max_column+1):
        cv, cf = ws.cell(r,c).value, wf.cell(r,c).value
        if cv is None and cf is None: continue
        co = ws.cell(r,c).coordinate
        row.append(f'{co}[{cf}]=>{cv}' if isinstance(cf,str) and cf.startswith('=') else f'{co}={cv}')
    if row: print(' | '.join(row))
"
```

### 5-3. 分析的検証(機械チェックでは拾えないもの)

- **Effective WACC inputs をoverridesと目視照合**
- EV < net_debt(+MI)ならPGMは自動でINVALIDラベル+Target平均から除外される — その状態を**受け入れるか、ターミナルcapexを見直すか**を判断して記録
- 型C/Eでは基準混在(連結倍率×事業EBITDA等)がないか
- 「無意味」と判断した指標がTarget Price計算に残っていないか(8410の教訓)

### 5-4. 旧テンプレ生成ファイル(2026-07-30以前)を参照するときの注意

7/30のFINAL群(285A/7203/8267/3687)と8410ファイルは旧テンプレ生成。参照時は引継ぎ文書§2の旧バグ一覧(FS年度ズレ、D27定数等)を頭に置くこと。特に:

- **8410_DCF_Model_20260730.xlsx のExec Summary(Target 380/BUY/EV/EBITDA 505)は無効**。手修正の最終ラウンドが未保存の中間状態。**正しい結論は引継ぎ文書§4の 276円/SELL(DDM 258/RI 293の平均)**。8410はDDM/RIシートと引継ぎ文書を正とする
- これらの銘柄を次に更新するときは、恒久修正後テンプレで**再生成**するのが原則(手修正の再適用より速く、validateが効く)

### 5-5. openpyxlで手修正するときの罠(重要度順)

1. **ラベル文字列を `=` で始めるとファイル自体が破損し得る**(#VALUE!で済まず、Excelでブックが開けなくなった実例あり)。ラベルは必ず非 `=` 始まりに
2. セル一括クリアのループは**参照されているセルを巻き込まない**(8267で株数セルを消して#DIV/0!×8)
3. 条件付き書式の拡張(x14)は保存時に落ちる(標準ルールは保持)
4. Windowsコンソール出力はcp932で落ちることがある — `sys.stdout.reconfigure(encoding='utf-8')`
5. **修正後は必ず再計算**して数式エラー0件を確認:
   ```bash
   python3 /mnt/skills/public/xlsx/scripts/recalc.py X.xlsx 180
   ```

---

## 6. 手順5: 最終化

1. すべての変更を **`Adjustments Log` シートに記録**(テンプレートが空シートを自動生成する。日付/セル/変更内容/理由/元の値/状態)。未解決事項も「未解決」として全件記載
2. validate_output.py を**最終版に対して再実行**し、FAIL 0件を確認
3. ファイル名: `<ticker>_DCF_Model_<YYYYMMDD>_FINAL.xlsx`
4. Target Priceの構成(どの手法を平均に入れ、どれをINVALID除外したか)をExecutive Summaryの注記に残す
5. **保存確認**: チャットで修正したファイルはダウンロード→`models/`配置までがセッションの一部。終了前に必ず
   ```
   dir models\*FINAL*
   ```
   で当該銘柄のFINALが存在することを確認する(8410でFINAL未保存のまま7/30セッションを終え、修正が失われた事故の再発防止)
6. 検証で見つかったテンプレート起因の問題は、修正仕様書側への追記候補としてユーザーに報告する

---

## 7. 手順6: セッション終了時

次チャット用に、以下を含む短い引継ぎメモを作る(長大な引継ぎ文書は不要。Adjustments Logが正)。

```
<ticker>のDCF完了。<ticker>_DCF_Model_<日付>_FINAL.xlsx のAdjustments Logに全変更履歴。
Target: <値>(<判定>)。validate: FAIL 0 / WARN <n>(<要旨>)。
未解決: <箇条書き>。次回: <やること>
```

---

## 8. 禁止事項(全手順共通)

1. `reports/` の上書き・削除
2. `__CONFIRM__` を残したままの最終ラン
3. validate_output.py のFAILを残したまま `_FINAL` を付けること
4. size_premium・シナリオ設計・Exit倍率の**既存銘柄からの無検討流用**
5. 上場状態・ティッカーの未確認のままのPeer/子会社リスト記載
6. 「この指標は無意味」と判断した指標をTarget Price計算に残すこと
7. D&Aが取れないPeerのEBITDA行混入(自動除外はされるが、入れないのが本筋)
8. ターミナル年に成長期水準のcapexを置くこと(capex ≈ D&Aに収束させる)
9. Peerの異常値を一次ソース検証なしに「データ誤り」と断定すること
10. thesis/risksに生の数値を書くこと(トークンを使う)
11. FINALの `models/` 保存確認をせずにセッションを終えること
