-- screener/db/schema.sql
-- 傾き検出スクリーナー DBスキーマ。仕様書 §5 に準拠。
--
-- 規約:
--   * §5 が定めた列は名前・意味を変えない。実装上必要になった列は §5 の列の
--     「後ろ」に追加し、各テーブルのコメントで【§5拡張】と明示する。黙って足さない。
--   * 取得の失敗・欠測は必ず行として残す(fetch_runs)。仕様書 §2-1
--     「取得失敗日はDBに欠損記録(黙って飛ばさない)」。
--   * マッピング不能タグは捨てず unknown_tags に頻度を積む。仕様書 §3-1。

PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

-- ---------------------------------------------------------------- companies
-- §5: companies(code, name, market, sector, mktcap, adv20, universe_flag)
CREATE TABLE IF NOT EXISTS companies (
    code            TEXT PRIMARY KEY,          -- 4桁 or 3桁+英字 (7203 / 285A)
    name            TEXT,
    market          TEXT,                      -- プライム/スタンダード/グロース 等
    sector          TEXT,                      -- 33業種名
    mktcap          REAL,                      -- JPY mn
    adv20           REAL,                      -- 20日平均売買代金 JPY mn
    universe_flag   INTEGER DEFAULT 0,         -- 1 = universe_rules.yaml を満たす
    -- 【§5拡張】
    sector17        TEXT,
    scale_category  TEXT,
    exclude_reason  TEXT,                      -- universe_flag=0 の理由(銀行/流動性不足等)
    source          TEXT,                      -- jquants / jpx / tdnet
    updated_at      TEXT
);

-- ------------------------------------------------------------------ filings
-- §5: filings(id, code, date, type[短信/修正/有報/半期], source[tdnet/edinet], path, xbrl_ok)
CREATE TABLE IF NOT EXISTS filings (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    code            TEXT,
    date            TEXT,                      -- YYYY-MM-DD (開示日)
    type            TEXT,                      -- 短信 / 修正 / 有報 / 半期 / その他
    source          TEXT,                      -- tdnet / edinet
    path            TEXT,                      -- 主たる保存物(XBRLがあればそれ、無ければPDF)
    xbrl_ok         INTEGER DEFAULT 0,         -- 1 = XBRLを保存できた
    -- 【§5拡張】
    doc_id          TEXT,                      -- TDnet: 140120260828527542 / EDINET: S100XY8O
    title           TEXT,
    disclosed_at    TEXT,                      -- YYYY-MM-DD HH:MM
    subtype         TEXT,                      -- 決算短信 / 業績予想修正 / 配当予想修正 ...
    pdf_path        TEXT,
    xbrl_path       TEXT,
    market_place    TEXT,                      -- TDnet の「東」等
    company_name    TEXT,
    fetched_at      TEXT,
    UNIQUE (source, doc_id)
);
CREATE INDEX IF NOT EXISTS ix_filings_code_date ON filings (code, date);
CREATE INDEX IF NOT EXISTS ix_filings_date      ON filings (date);
CREATE INDEX IF NOT EXISTS ix_filings_type      ON filings (type, subtype);

-- ----------------------------------------------------------- financials_cum
-- §5: financials_cum(filing_id, code, period, q_no, item, value) — 累計原本
CREATE TABLE IF NOT EXISTS financials_cum (
    filing_id       INTEGER,
    code            TEXT,
    period          TEXT,                      -- 会計期間ラベル (FY2026/3 等)
    q_no            INTEGER,                   -- 1..4 (4 = 通期), NULL = 不明
    item            TEXT,                      -- account_mapping.yaml の内部項目名
    value           REAL,
    -- 【§5拡張】
    unit            TEXT,                      -- JPY / JPY mn / shares / pure
    context_ref     TEXT,
    source_tag      TEXT,                      -- 元のXBRL要素名(追跡用)
    PRIMARY KEY (filing_id, item, context_ref),
    FOREIGN KEY (filing_id) REFERENCES filings (id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS ix_cum_code_item ON financials_cum (code, item);

-- 【§5拡張・docs/segment_dimension_design.md】
-- 次元(セグメント等)付きファクト。financials_cum は「見出し数値だけ」という
-- 不変条件を保ちたいので別テーブルにする。cum に segment 列を足すと、全ての
-- 既存クエリが WHERE segment IS NULL を要求するようになり、書き忘れた瞬間に
-- 集計が黙って2〜3倍になる。分けておけばその間違いは構造的に起こらない。
CREATE TABLE IF NOT EXISTS financials_dim (
    filing_id       INTEGER,
    code            TEXT,
    period          TEXT,
    q_no            INTEGER,
    item            TEXT,                      -- account_mapping.yaml の内部項目名
    axis            TEXT,                      -- segment / segment_total / adjustment / other
    member          TEXT,                      -- 正規化後 (Japan / Philippines)
    member_raw      TEXT,                      -- 元文字列。正規化を後から検証できる
    value           REAL,
    unit            TEXT,
    context_ref     TEXT,
    source_tag      TEXT,
    valid_flag      INTEGER DEFAULT 1,         -- 0 = セグメント区分変更で時系列が切れている
    invalid_reason  TEXT,
    PRIMARY KEY (filing_id, item, context_ref),
    FOREIGN KEY (filing_id) REFERENCES filings (id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS ix_dim_code_item ON financials_dim (code, item, member);
CREATE INDEX IF NOT EXISTS ix_dim_axis ON financials_dim (axis);

-- ------------------------------------------------------------- financials_q
-- §5: financials_q(code, period, q_no, item, value, valid_flag) — 単独値(生成)
CREATE TABLE IF NOT EXISTS financials_q (
    code            TEXT,
    period          TEXT,
    q_no            INTEGER,
    item            TEXT,
    value           REAL,
    valid_flag      INTEGER DEFAULT 1,         -- 0 = 決算期変更/遡及修正/連結範囲変更で無効
    -- 【§5拡張】
    invalid_reason  TEXT,
    -- この単独値が何四半期ぶんか。1 = 真の四半期単独値。
    -- 短信が揃っていない会社は EDINET の有報(q4累計)と半期(q2累計)しか無く、
    -- 差分は「下期6ヶ月」になる。それを四半期と名乗らせないための列 ——
    -- 3ヶ月と6ヶ月を同じ土俵に載せると傾きの大きさが二重になる。
    span_q          INTEGER DEFAULT 1,
    built_at        TEXT,
    PRIMARY KEY (code, period, q_no, item)
);

-- ----------------------------------------------------------------- guidance
-- §5: guidance(code, date, fy, item, value, revision_direction)
CREATE TABLE IF NOT EXISTS guidance (
    code                TEXT,
    date                TEXT,
    fy                  TEXT,
    item                TEXT,
    value               REAL,
    revision_direction  TEXT,                  -- up / down / flat / initial
    -- 【§5拡張】
    filing_id           INTEGER,
    prev_value          REAL,
    PRIMARY KEY (code, date, fy, item),
    FOREIGN KEY (filing_id) REFERENCES filings (id) ON DELETE CASCADE
);

-- 【統合引継ぎ書 v1.1 §6-1 / docs/adapter_design.md A】
-- 一時収入の分離層。単独値ビルダーが正しくても、一過性の売上は傾きを汚染する。
-- 3905 の FY2027Q1 は粗利率 90.90%(前年 27.58%)で S1 が +63.3pt の買いシグナルと
-- して誤発火するが、一時収入を除くと粗利は前年割れ。この層が無いと S1/S5 が
-- 実銘柄で誤作動する。
--
-- source_note を NOT NULL + CHECK(20文字以上) にしてあるのは、「金額だけ書いて
-- 済ませる」ことを構造的に不可能にするため。アプリ層の検証ではなくスキーマ制約に
-- 置くのは、検証を通さない経路(手作業のSQL・別スクリプト)からの登録を防ぐため。
-- 一過性収益の特定は機械抽出しない。注記テキスト由来で判断を要するので半手動。
CREATE TABLE IF NOT EXISTS pl_adjustments (
    code             TEXT NOT NULL,
    period           TEXT NOT NULL,      -- FY2027 (本体の period 語彙)
    q_no             INTEGER NOT NULL,
    item_key         TEXT NOT NULL,      -- one_time_revenue / one_time_cost / one_time_gain
    amount           REAL NOT NULL,      -- 控除する額(正値)。単位は本体と同じ「円」
    source_note      TEXT NOT NULL,      -- 短信/有報の注記からの引用(原文ママ)
    source_filing_id INTEGER NOT NULL,   -- 引用元の書類
    source_locator   TEXT NOT NULL,      -- 注記の所在(例:「(セグメント情報等) 3.」)
    confirmed_by     TEXT NOT NULL,      -- 記録した人
    confirmed_at     TEXT NOT NULL,
    note             TEXT,               -- 判断メモ(任意)
    PRIMARY KEY (code, period, q_no, item_key),
    FOREIGN KEY (source_filing_id) REFERENCES filings (id) ON DELETE CASCADE,
    CHECK (length(trim(source_note)) >= 20),
    CHECK (length(trim(source_locator)) > 0),
    CHECK (length(trim(confirmed_by)) > 0),
    CHECK (amount > 0)
);
CREATE INDEX IF NOT EXISTS ix_adj_code ON pl_adjustments (code, period, q_no);

-- ------------------------------------------------------------------ signals
-- §5: signals(code, eval_date, signal_id, value, fired, evidence_text)
CREATE TABLE IF NOT EXISTS signals (
    code            TEXT,
    eval_date       TEXT,
    signal_id       TEXT,                      -- S1..S8
    value           REAL,
    fired           INTEGER,
    evidence_text   TEXT,
    PRIMARY KEY (code, eval_date, signal_id)
);

-- ------------------------------------------------------------------- scores
-- §5: scores(code, week, total, rank)
CREATE TABLE IF NOT EXISTS scores (
    code            TEXT,
    week            TEXT,                      -- ISO week (2026-W35)
    total           REAL,
    rank            INTEGER,
    PRIMARY KEY (code, week)
);

-- ------------------------------------------------------------------- prices
-- §5: prices(code, date, close, volume, adv20) — J-Quants
CREATE TABLE IF NOT EXISTS prices (
    code            TEXT,
    date            TEXT,
    close           REAL,
    volume          REAL,
    adv20           REAL,                      -- 20日平均売買代金 JPY mn
    -- 【§5拡張】
    turnover_value  REAL,                      -- 当日売買代金 JPY
    -- 【§5拡張・2026-08-31】始値と調整後株価。J-Quants は5年ローリングで、
    -- 窓から落ちた日付は二度と取得できない。取れるうちに全部取る。
    --   open   出口ルール「翌日寄り指値」(v1.1 §5-5)の再現に必須
    --   adj_*  株式分割の調整。未調整だと分割が偽の -50% リターンになる
    --          (86970 の 2022-01-04 は C=2518.5 / AdjC=1259.3)
    --   mktcap 日次時価総額。ユニバース条件の時点再現に使える
    open            REAL,
    high            REAL,
    low             REAL,
    adj_factor      REAL,
    adj_close       REAL,
    adj_volume      REAL,
    mktcap          REAL,                      -- JPY mn
    PRIMARY KEY (code, date)
);

-- 【§5拡張・統合引継ぎ書 v1.1 §4-2「超過リターン評価」】
-- 市場指数(TOPIX)の日次終値。バックテストの層別分析(相場環境別)と、
-- 絶対リターンではなく超過リターンで評価するために要る。
-- 「持っていただけ」の効果とイベント効果を分離できないと、エッジの有無が判定できない。
CREATE TABLE IF NOT EXISTS market_index (
    date   TEXT PRIMARY KEY,               -- YYYY-MM-DD
    close  REAL NOT NULL,
    name   TEXT DEFAULT 'TOPIX'
);

-- =====================================================================
-- 以下は §5 に無いが仕様書本文が要求する運用テーブル。追加である旨を明示する。
-- =====================================================================

-- 【§5拡張・仕様書 §2-1 の「欠損記録」要件】
-- 取得を試みた単位(ソース×対象日)を必ず1行残す。status:
--   ok       取得成功(件数 > 0)
--   empty    正常に取得できたが対象0件(休日・非営業日など)
--   partial  一部のダウンロードに失敗(n_failed > 0)
--   failed   一覧の取得自体に失敗 → 再実行対象
CREATE TABLE IF NOT EXISTS fetch_runs (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    source          TEXT NOT NULL,             -- tdnet / edinet / jquants
    target_date     TEXT NOT NULL,             -- YYYY-MM-DD
    status          TEXT NOT NULL,
    n_listed        INTEGER DEFAULT 0,         -- 一覧に載っていた件数
    n_target        INTEGER DEFAULT 0,         -- 保存対象と判定した件数
    n_saved         INTEGER DEFAULT 0,
    n_failed        INTEGER DEFAULT 0,
    started_at      TEXT,
    finished_at     TEXT,
    attempt         INTEGER DEFAULT 1,
    error           TEXT,
    note            TEXT
);
CREATE INDEX IF NOT EXISTS ix_runs_src_date ON fetch_runs (source, target_date);

-- 【§5拡張・仕様書 §3-1 の「マッピング不能タグは unknown として記録し頻度上位を定期レビュー」】
CREATE TABLE IF NOT EXISTS unknown_tags (
    source          TEXT NOT NULL,             -- tdnet_summary / tdnet_attachment / edinet
    tag             TEXT NOT NULL,             -- XBRL 要素名(名前空間接頭辞を除く)
    n               INTEGER DEFAULT 0,         -- 出現回数
    n_filings       INTEGER DEFAULT 0,         -- 出現した filing 数
    sample_value    TEXT,
    first_seen      TEXT,
    last_seen       TEXT,
    PRIMARY KEY (source, tag)
);
CREATE INDEX IF NOT EXISTS ix_unknown_n ON unknown_tags (n DESC);
