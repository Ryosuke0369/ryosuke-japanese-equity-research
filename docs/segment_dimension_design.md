# セグメント次元の取り込み設計 (2026-08-31)

対象: `screener/extract/xbrl_parser.py` / `screener/db/schema.sql` /
`screener/config/account_mapping.yaml`
関連: 仕様書 §3-1(XBRL科目マッピング) / §4(指標層 S1〜S8) / §5(DBスキーマ)

## 0. なぜ必要か

XBRL のファクトは `contextRef` に次元(ディメンション)を持つ。同じ
`jppfs_cor:NetSales` でも、

    CurrentYearDuration                                   全社の売上（見出し）
    CurrentYearDuration_ReportableSegmentsMember          報告セグメント計
    CurrentYearDuration_..._JapanReportableSegmentsMember 日本セグメントの売上

はまったく別の数字である。これを同じ `item='revenue'` で `financials_cum` に
入れると、`revenue` を引いたときに全社・セグメント計・個別セグメントが混ざって
返り、**合計が2〜3倍になる**。2026-08-31 時点の EDINET パーサは次元付きを
取り込み対象外にして件数だけ数えており(`dimensional` カウンタ)、TDnet 側は
そのまま `financials_cum` に混入させていた。

一方、仕様書 §4 の S1〜S8 は**セグメント別の傾き**を見る必要がある。捨てても
混ぜても駄目で、軸を分けて保持するのが正解。

## 1. 実測: 次元付きファクトの内訳

有報40件・短信60件を読み直して分類した(マッピング済み項目のみ)。

| 次元 | EDINET有報 | TDnet短信 | 用途 |
|---|---:|---:|---|
| セグメント | 24.9% | **68.8%** | **S1〜S8 に必須** |
| 株主資本変動(資本金・利益剰余金・自己株式ほか) | 62.3% | 3.1% | 使わない |
| 調整額 ReconcilingItems | 4.1% | 9.3% | セグメント合計の検算 |
| その他(大株主・行番号ほか) | 8.7% | 18.8% | 使わない |

TDnet 短信ではマッピング済みファクトの **27.7% が次元付き**で、その約7割が
セグメント。EDINET は株主資本等変動計算書が量を占めるだけで、価値の中心は
どちらもセグメントである。

セグメント member には2種類ある。混ぜると二重計上になる。

    ReportableSegmentsMember                          報告セグメント計（集計）
    TotalOfReportableSegmentsAndOthersMember          報告セグメント+その他 計
    jpcrp030000-asr_E02121-000JapanReportableSegmentsMember   日本（個別）

## 2. 設計

### 2-1. 別テーブルにする（`financials_cum` に列を足さない）

```sql
CREATE TABLE IF NOT EXISTS financials_dim (
    filing_id   INTEGER,
    code        TEXT,
    period      TEXT,
    q_no        INTEGER,
    item        TEXT,        -- account_mapping.yaml の内部項目名（既存と共通）
    axis        TEXT,        -- segment / segment_total / adjustment / other
    member      TEXT,        -- 正規化後: 'Japan' 'Philippines'
    member_raw  TEXT,        -- 元文字列。正規化を後から検証できるように残す
    value       REAL,
    unit        TEXT,
    context_ref TEXT,
    source_tag  TEXT,
    valid_flag  INTEGER DEFAULT 1,   -- 0 = セグメント区分変更で時系列が切れている
    invalid_reason TEXT,
    PRIMARY KEY (filing_id, item, context_ref),
    FOREIGN KEY (filing_id) REFERENCES filings (id) ON DELETE CASCADE
);
```

**`financials_cum` に `segment` 列を足す案は採らない。** その場合すべての既存
クエリが `WHERE segment IS NULL` を必要とし、**書き忘れると集計が黙って2〜3倍
になる**。テーブルを分ければその間違いは構造的に起こりえない。
`financials_cum` は「見出し数値だけ」という不変条件を保てる。

### 2-2. `axis` は集計と個別を必ず分ける

`ReportableSegmentsMember`(報告セグメント計)を `segment` として個別セグメントと
同じ軸に置くと、S1 のセグメント別集計で全社ぶんがもう一度足される。
集計は `axis='segment_total'` に隔離し、既定の集計対象から外す。用途は
「個別セグメントの和 == セグメント計」の検算。

### 2-3. member 正規化

    jpcrp030000-asr_E02121-000JapanReportableSegmentsMember  ->  Japan

EDINET企業コード接頭辞(`jpcrp\d+-\w+_E\d+-\d+`)と `ReportableSegmentsMember`
接尾辞を剥がす。日本語のセグメント名は `lab.xml`(ラベルリンクベース)にあるが、
S1〜S8 の計算には英字IDで足りるので初手では解決しない。`member_raw` を必ず
残すので、後からラベルを引き直せる。

### 2-4. 株主資本変動は保存しない（件数のみ）

EDINET の次元付きファクトの62%を占めるが S1〜S8 では使わない。全件保存すると
`financials_dim` が約40万行になる。`axis='equity_component'` と判定した分は
**保存せず件数だけ計上**し、ログに出す。必要になったら再解析で足せる
(生の zip は `D:\screener_data\raw\edinet\` に残っている)。

### 2-5. セグメント区分変更のガード

会社は報告セグメントの区分を変更する。member が変われば時系列は切れており、
そこで前期比を取ると**存在しない変化を検出する**。`financials_q.valid_flag` と
同じ思想で、

- ある `(code, item, member)` について前期に同一 member が存在しない場合、
  当期の行は `valid_flag=0` / `invalid_reason='セグメント区分変更'` にする
- 新設セグメントと区分変更を区別しない。どちらも「前期比は取れない」で同じ

「区分が変わった」と「まだ判定していない」を混同しないため、**行は消さない**。

## 3. 仕様書 §4 S1 への但し書き

セグメント別の粗利率は**多くの会社で算出できない**。XBRL のセグメント注記に
COGS が無いため。実測で取れるのは売上(`RevenuesFromExternalCustomers`)、
セグメント利益、減価償却費、資本的支出まで。

したがって **S1 をセグメント軸に適用する場合は「セグメント別営業利益率」で
代替する**。全社レベルの S1 は従来どおり粗利率で計算する(3441 の主砲はここ)。

## 4. 影響範囲と移行手順

**破壊的変更は無い。** `financials_cum` の消費者は 2026-08-31 時点で存在しない
(`quarterly_builder.py` は未実装、`financials_q` 0行、`signals` 0行)。

| 手順 | 内容 | リスク |
|---|---|---|
| 1 | `schema.sql` に `financials_dim` 追加(`IF NOT EXISTS` なので既存DBはそのまま) | なし |
| 2 | `store_filing` の `continue` を「捨てる」から「`financials_dim` へ書く」に変更。**TDnet/EDINET共通**(現在の source 分岐を廃止) | なし。TDnet は現在 `financials_cum` に混入させているので、これは**バグ修正**でもある |
| 3 | 既存の次元付き行を掃除: `DELETE FROM financials_cum WHERE context_ref LIKE '%Member%'` | 再解析前に必須 |
| 4 | TDnet(数分)・EDINET(約76分)を再解析 | なし |
| 5 | `quarterly_builder` は最初から `financials_cum`(見出しのみ)を読む | 今から書くので追随不要 |

**手順2で TDnet 側も同時に変える。** 分けて後回しにすると、TDnet のセグメント値が
見出し数値と同じ `item` 名で `financials_cum` に居座り続け、S1 が全社粗利率と
セグメント粗利率を取り違える余地が残る。消費者がいない今が最も安全なタイミング。
