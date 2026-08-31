# 別枠 earnings_screener 接続アダプタ 設計と実装 (2026-08-31)

対象: `screener/projection/` / `screener/extract/adjustment_recorder.py` /
`screener/config/earnings_screener_mapping.yaml` / `screener/config/adjustment_detection.yaml`
関連: 統合引継ぎ書 v1.1 §3・§4-5〜4-8・§6-1・§9-3 / 仕様書 §3-2・§4

## 0. 方針 —— アダプタではなく「投影層」

```
本体ファクト層                     投影層 (screener/projection/)      別枠モジュール
─────────────                 ─────────────────────────       ──────────────
filings                    ┐
financials_cum (見出し)     ├─→  pit.py     as_of + 訂正世代解決   ┐
financials_dim (次元付き)   │                                      ├─→ scorers
financials_q  (単独値)      ├─→  adapter.py §3-1 の4関数を実装      │   forecaster
pl_adjustments ★新設       │              item_key 写像 / span_q   │   weekly_screener
guidance                   ┘              単位変換 / 調整反映      ┘   backtest
```

**別枠のコードは1行も変更しない。** `data_access.py` と同じシグネチャを本体側で
実装し、import 先だけを差し替える。理由は2つ:

1. 別枠は `tests/test_pit.py` で不変条件が固定されている。書き換えるとその保証を失う
2. v1.1 の `generation` 機構は「モック側でのPIT検証用」と v1.1 §0-1 自身が位置づけている。
   本体は `financials_cum(filing_id, item, context_ref)` で同等以上を自然に表現できる

## A. pl_adjustments（一時収入分離層）

### A-1. 実データによる必要性の証明

3905 の FY2027 Q1（TDnet短信 2026-08-14、`span_q=1` の真の四半期）:

| | 当Q | 前年同Q |
|---|---:|---:|
| 売上 | 6,327,633千円 | 668,083千円 |
| 粗利 | 5,752,383千円 | 184,252千円 |
| 粗利率 | **90.90%** | **27.58%**（+63.3pt） |
| 営業利益 | +4,843,382千円 | −342,241千円（赤字→黒字） |

一時収入55.8億円を除くと売上7.48億（+11.9%）に対し**粗利は前年割れ（−6.4%）**。
この層が無いと **S1 が +63.3pt の買いシグナルを、S5 が黒字転換を、同時に誤発火**する。
したがって §9 の手順5ではなく、投影層と同時に実装した。

### A-2. 根拠なき調整を登録できない構造

`pl_adjustments` は `source_note` を `NOT NULL` + `CHECK(length(trim(...)) >= 20)` にしてある。
**アプリ層の検証ではなくスキーマ制約**に置いたのは、検証を通さない経路（手作業のSQL、
別スクリプト）からの登録を防ぐため。`source_filing_id` は `filings` への外部キーで、
`source_locator`（注記の所在）と `confirmed_by`（記録者）も必須。

### A-3. 半手動プロセス（恒久的な運用手順）

一過性収益の特定は**機械抽出しない**。3905 の55.8億円は注記テキスト由来で
数値パーサー（`nonFraction` のみ）の対象外であり、金額の確定には人が注記を読む必要がある。
これは一時的な技術制約ではなく、「何が一過性か」が判断を要する作業だから恒久的に半手動。

```
1. 検出  python -m screener.extract.adjustment_recorder --scan
         前年同Qに対し 売上2倍超/半分未満、粗利率±20pt、営業損益の符号反転
         → tasks/pl_adjustment_queue.md に追記（tasks/ は追記専用）
2. 確認  人が該当書類の注記を読む（PDF は raw/tdnet/ に保存済み）
3. 記録  --record で登録。引用20文字以上・所在・記録者が無いと弾かれる
4. 反映  投影層が normalized_sales に自動反映
5. 監査  --queue で登録済み一覧。放置を可視化する
```

閾値は `config/adjustment_detection.yaml`。**緩めに置いてある** —— 見逃した一過性収益は
S1/S5 の誤発火として直接損失につながるが、余計に積まれた期は人が1分で棄却できる。

### A-4. 適用規則

`Series.__init__` が `revenue` / `gross_profit` / `operating_income` から
`one_time_revenue` を控除する（v1.1 §6-1 の「OPに全額フロー」既定仮定）。
**調整した期は根拠文に「(一時収入控除後)」と出る** —— 調整したことが見えない調整は、
調整していないより危険。

## C-1. PIT 意味論

### 訂正世代の解決

`financials_cum` の主キーが `(filing_id, item, context_ref)` で書類ごとにファクトを
保持しているため、訂正短信は自動的に別行になる。`ROW_NUMBER() OVER (PARTITION BY
code, period, q_no, item, context_ref ORDER BY f.date DESC, f.id DESC)` の1行で
「as_of 時点の最新版」が決まる。**`generation` 列は作らない。**

### 引け後発表

TDnet は引け後発表が多く、発表日当日の値を当日の判定に使うとルックアヘッドになる。

| mode | 条件 | 用途 |
|---|---|---|
| `strict`（既定） | `filings.date < as_of` | バックテスト。当日発表を見ない |
| `live` | `filings.date <= as_of` | 実運用。寄りまでに発表済みを見る |

**危険な側を既定にしない**ため `strict` を既定にし、`live` は明示的に選ばせる。

### 不変条件（`screener/tests/test_projection.py`）

| ID | 内容 | 結果 |
|---|---|---|
| P1 | 訂正を跨いで as_of を1日動かすと返る値が変わる | ✅ |
| P2 | as_of 指定の結果は未指定の結果の部分集合 | ✅ |
| P3 | 発表日当日は strict で見えず live で見える | ✅ |
| P4 | スコアラーが投影層を経由せずDBを読んでいない（AST検査） | ✅ |

**P4 は実装中に本物の違反を検出した。** `signal_defs.py` が3箇所でDBを直読しており、
うち2箇所は S5 の `guidance` 読み出し —— **別枠 S5 が v1.1 で修正したのと同じ違反を、
本体で独立に作っていた**。投影層経由に修正済み。P4 は「結果の書き込み(INSERT)」は
対象外とし、SELECT を含む文字列リテラルだけを探す。

## C-2. item_key 写像表

`config/earnings_screener_mapping.yaml`。`status` は3値。

| 別枠 item_key | 本体 item | status |
|---|---|---|
| `accounts_receivable` | `trade_receivables` | verified |
| `contract_liabilities` | `contract_liabilities` | verified |
| `construction_in_progress` | `construction_in_progress` | verified |
| `machinery` | `machinery_and_equipment` | verified |
| `inventory` | `inventories_total` | verified |
| `deposits_received` | `advances_received` | **unverified**（預り金 vs 前受金。意味が一致するか未確認） |
| `shares_outstanding` | `shares_issued` | **unverified**（自己株控除の有無が未確認） |
| `advances_paid` | — | **missing** |
| `rev_over_time` | — | **missing**（S10 に必須） |
| `rev_point_in_time` | — | **missing**（S10 に必須） |
| `diluted_shares` | — | **missing**（S11 に必須） |

**verified 以外は投影層が空を返し、理由をログに出す。** 名前が近いだけで意味がずれる
可能性がある以上、推測で繋ぐと間違った数字が静かに流れる。0 や None を黙って返すと、
下流は「データが無い」と「まだ繋いでいない」を区別できなくなる。

missing 4項目は `account_mapping.yaml` への追加が前提。全件再解析後の `unknown_tags` で
該当タグの実在と頻度を確認する。

## C-3. span_q の分離

**`span_q=1` 以外は既定で渡さない。** 混入させない方法ではなく、渡さない方法を採る。

| 本体 `financials_q` | 投影 | 扱い |
|---|---|---|
| `span_q=1`, `valid_flag=1` | `quarter_type` に写像して渡す | 通常評価 |
| `span_q=2`（半期粒度） | `allow_span=(1,2)` 指定時のみ `is_half=True` 付き | 半期専用分析のみ |
| `span_q>=3` / `valid_flag=0` | 渡さない | — |

`fiscal_year` は本体の `period`（FY2027）から取る。提出日から推定してはいけない ——
8月提出の「2027年3月期第1四半期」が FY2026 になり別の会計年度と引き算される事故を
実際に起こした（2026-08-31 修正済み）。

## C-4. スコア意味論の互換層

`Signal` に `strength: float | None` を追加。**`value` を流用しない** —— `value` は
単位が指標ごとに違う（S1はpt、S2は%、S5はpt）ので、そのまま平均すると単位の違う数を足す。

```python
def _strength(value, fire_pt):
    """発火閾値の2倍で飽和する線形正規化。閾値ちょうどで ±0.5。"""
    return max(-1.0, min(value / (2.0 * abs(fire_pt)), 1.0))
```

正規化の分母は `signal_thresholds.yaml` の発火閾値から導くので、**魔法の数字を新しく増やさない**。

| 本体 | 別枠 |
|---|---|
| `strength` | `score` |
| `value is not None` | `available` |
| `evidence` | `evidence` |

**未解決**: 本体の集計方法（重み付き和）は未実装で `signal_weights.yaml` も未作成。
別枠は単純平均を前提にしているため、重み導入時に §8 の `SCORE_THRESHOLD=0.10` の再校正が要る。

## PTS 出来高のデータソース（承認事項4への回答）

**当面「取得しない」を既定とする**（承認済み）。

- J-Quants V2 は東証の公式配信で、取引所外取引である PTS の出来高は**エンドポイントに存在しない**
- PTS 出来高を配信する無料APIは把握していない。証券会社ツールは画面表示のみでAPI提供がなく、
  スクレイピングは規約上の問題がある
- 有料ベンダーは月額が発生し、他が全て無料枠で動いていることと釣り合わない

v1.1 §5-5 の fallback「翌日寄り指値固定」を既定運用とし、`pts_observations` は
スキーマだけ用意して空のままにする。`decide_exit()` は `pts_ratio=None` を受けると
自動的に「翌日寄り指値（PTSデータなし）」を返すので、**コード変更は不要**。

## 設計の穴（未解決・実装していない）

1. **`quarter_type='FY'` の意味**。`_seasonal_share()` が通期累計シェアを計算しており、
   単独値の系列に対して正しく動くか未検証
2. **`operating_cf` の四半期単独化**。短信のCF開示は半期のみが通例で、差分の相手が無い期の
   扱いが本体と別枠で食い違う可能性
3. **`cogs_quantity` / `cogs_price`**（原価の数量/価格分解）。**本体にも XBRL にも該当科目が無い**。
   別枠のモックは生成時に注入しているだけ。投影層は `None` を明示的に返す（0 にしない）。
   別枠 S6 と本体 S2 が同じ問題を抱える
