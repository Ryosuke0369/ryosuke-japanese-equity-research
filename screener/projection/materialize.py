"""screener/projection/materialize.py — 別枠 earnings_screener 形の投影DBを作る。

なぜ別ファイルなのか
--------------------
別枠は `SELECT ... FROM filings` を素のSQLで叩くが、本体の `filings` は
書類索引という**別の意味**のテーブルで、同名なのでビューで覆えない
(daily_prices のようにはいかない)。したがって別枠スキーマの形をした
DBファイルを1本生成し、別枠には `db_path` を差し替えて渡す。
**別枠のコードは1行も変更しない。**

投影DBは生成物であって正本ではない
----------------------------------
本体DBから何度でも作り直せる。壊れたら消して作り直すのが正しい対処で、
中身を手で直してはいけない。日次バッチから毎回作り直せるように、
このスクリプト1本で完結させてある。

PIT の焼き込み
--------------
別枠の data_access.visible_generations は
`filings(ticker, period_end, filing_date, generation)` の MAX(generation)
を as_of で絞って「その時点の版」を決める。投影側はこの契約に合わせて
**(銘柄×期×四半期) ごとに、その四半期を最初に開示した書類の日付**を
filing_date として入れる。したがって別枠が as_of を動かせば、
「その時点で公知だった四半期だけ」が見える。

限界（重要・黙って隠さない）:
  本体は同じ (期,四半期,項目) を後の書類が言い直した場合、行を差し替えず
  **valid_flag=0 で無効化**する設計になっている。つまり「訂正前の値」を
  保持していない。よって投影DBの generation は常に 1 で、
  「訂正前を再現する」再生は**できない**。訂正された四半期は
  is_valid=0 として渡るので、別枠からは最初から無効に見える。
  これは前方視バイアスではなく逆方向（当時は有効だった値を落とす）で、
  安全側だがサンプルは減る。

    python -m screener.projection.materialize
    python -m screener.projection.materialize --shift-days -3
    python -m screener.projection.materialize --shift-days 3
"""
from __future__ import annotations

import argparse
import os
import re
import sqlite3
import sys
import unicodedata
from collections import defaultdict

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.projection import adapter

# 別枠 common/schema.py の DDL のうち投影に必要な部分。別枠から import せず
# 写しているのは、別枠が untracked な配布物でパスが動きうるため。
# 差異が出たら test_materialize が気づく。
DDL = """
CREATE TABLE IF NOT EXISTS universe (
    ticker TEXT PRIMARY KEY, company_name TEXT, market TEXT, sector TEXT,
    fiscal_year_end INTEGER NOT NULL,
    market_cap_ok INTEGER DEFAULT 1, liquidity_ok INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS filings (
    id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT NOT NULL,
    filing_date TEXT NOT NULL, period_end TEXT NOT NULL,
    quarter_type TEXT NOT NULL, fiscal_year INTEGER NOT NULL,
    source TEXT DEFAULT 'mock', pdf_path TEXT, xbrl_path TEXT,
    -- doc_id: その期を初めて開示した書類。**根拠期の原文に飛ばすのに要る。**
    doc_id TEXT,
    generation INTEGER DEFAULT 1, UNIQUE(ticker, period_end, generation));
CREATE INDEX IF NOT EXISTS idx_filings_ticker ON filings(ticker, filing_date);
CREATE TABLE IF NOT EXISTS quarterly_standalone (
    id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT NOT NULL,
    fiscal_year INTEGER NOT NULL, quarter_type TEXT NOT NULL,
    period_end TEXT NOT NULL, sales REAL, operating_profit REAL,
    gross_profit REAL, cogs_quantity REAL, cogs_price REAL, operating_cf REAL,
    is_valid INTEGER DEFAULT 1, invalid_reason TEXT, generation INTEGER DEFAULT 1,
    span_q INTEGER DEFAULT 1, period_start TEXT,
    UNIQUE(ticker, period_end, generation));
-- span-matched 用。**全 span をここに入れる。**
--
-- 設計書は「本体テーブルに全 span を入れ、別枠には span=1 のビューを
-- 見せる」としていたが、別枠 data_access は `quarterly_standalone` を
-- **テーブル名で直に叩く**（ビュー名を差し替える余地が無い）。
-- したがって逆にする: 既存テーブルは span=1 のまま（別枠は無変更で動く）、
-- 全 span はこちらへ。別枠のコードを1行も変えない方針を優先した。
--
-- period_start は「その累計がどこから始まるか」。span_q だけでは
-- Q2累計の6ヶ月と Q3-Q4 の6ヶ月が区別できない。
CREATE TABLE IF NOT EXISTS quarterly_standalone_all (
    id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT NOT NULL,
    fiscal_year INTEGER NOT NULL, quarter_type TEXT NOT NULL,
    period_end TEXT NOT NULL, span_q INTEGER NOT NULL, period_start TEXT,
    sales REAL, operating_profit REAL, gross_profit REAL, operating_cf REAL,
    is_valid INTEGER DEFAULT 1, invalid_reason TEXT, generation INTEGER DEFAULT 1,
    UNIQUE(ticker, period_end, span_q, generation));
CREATE INDEX IF NOT EXISTS idx_qs_all ON quarterly_standalone_all
    (ticker, fiscal_year, quarter_type, span_q);
CREATE TABLE IF NOT EXISTS pl_adjustments (
    id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT NOT NULL,
    period_end TEXT NOT NULL, item_key TEXT NOT NULL, amount REAL NOT NULL,
    note TEXT, generation INTEGER DEFAULT 1,
    UNIQUE(ticker, period_end, item_key, generation));
CREATE TABLE IF NOT EXISTS company_forecasts (
    id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT NOT NULL,
    fiscal_year INTEGER NOT NULL, forecast_sales REAL NOT NULL,
    forecast_op REAL NOT NULL, source_date TEXT NOT NULL,
    -- 改訂方向と前回予想（本体 guidance 由来。2026-09-20 追加）。
    -- 出口の分岐2 の第3条件（通期営業利益予想の引き下げ）はこれを読む。
    revision_direction TEXT, prev_op REAL,
    UNIQUE(ticker, fiscal_year, source_date));
CREATE TABLE IF NOT EXISTS balance_sheet_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT NOT NULL,
    period_end TEXT NOT NULL, item_key TEXT NOT NULL, value REAL NOT NULL,
    generation INTEGER DEFAULT 1, UNIQUE(ticker, period_end, item_key, generation));
CREATE TABLE IF NOT EXISTS daily_prices (
    ticker TEXT NOT NULL, date TEXT NOT NULL, close REAL NOT NULL, volume REAL,
    PRIMARY KEY (ticker, date));
CREATE INDEX IF NOT EXISTS idx_prices_date ON daily_prices(date);
CREATE TABLE IF NOT EXISTS market_index (date TEXT PRIMARY KEY, close REAL NOT NULL);
-- 別枠 module_b/forecaster.py が参照する。本体に対応する概念が無いので
-- **空のまま置く**。テーブルごと無いと OperationalError になり、
-- forecaster を呼ぶ側が「データ不足」と誤認する（2026-09-02 に B3 で踏んだ）。
-- 空であることと存在しないことは別物、をここでも守る。
CREATE TABLE IF NOT EXISTS evidence_events (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker        TEXT NOT NULL,
    event_date    TEXT NOT NULL,
    event_type    TEXT NOT NULL,
    amount        REAL,
    evidence_flag INTEGER NOT NULL,
    source_doc    TEXT,
    note          TEXT
);
-- 別枠 Module A (earnings_calendar.py) が書き込む先。投影側は器だけ用意する。
-- 中身は build_calendar() が「同四半期の前年発表日 + 1年 → 営業日寄せ」で埋める。
CREATE TABLE IF NOT EXISTS earnings_calendar (
    ticker              TEXT PRIMARY KEY,
    next_earnings_date  TEXT NOT NULL,
    quarter_type        TEXT NOT NULL,
    fiscal_year         INTEGER NOT NULL,
    fiscal_year_end     INTEGER NOT NULL,
    confidence_level    TEXT NOT NULL,
    estimated_from      TEXT NOT NULL,
    updated_at          TEXT NOT NULL
);
"""

# 「2025/05/01-2026/04/30」形式（新しい書類）
_RE_SLASH = re.compile(r"(\d{4})/(\d{2})/(\d{2})\s*[^\d]{1,3}\s*(\d{4})/(\d{2})/(\d{2})")
# 「令和3年4月1日-令和4年3月31日」形式（2022年頃までの書類）
_RE_WAREKI = re.compile(
    r"(?:平成|令和)\s*\d+\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日"
    r"[^\d]{1,3}"
    r"(?:平成|令和)\s*\d+\s*年\s*(\d{1,2})\s*月")

_FLOW = {"revenue": "sales", "operating_income": "operating_profit",
         "gross_profit": "gross_profit", "operating_cf": "operating_cf"}
# 別枠 quarterly_standalone のフロー列。BS項目と混ぜない。
_NON_BS = ("revenue", "operating_income", "gross_profit", "operating_cf")


def fiscal_year_end_month(title):
    """書類タイトルの会計期間から決算期末月を取る。読めなければ None。

    決算期末月は別枠 universe の NOT NULL 列。推測で埋めると
    「3月期だと思って読んだ12月期の会社」が静かに混ざるので、
    読めない銘柄は universe に入れない（sector が引けず不明扱いになるだけ）。

    **期間が1年（11ヶ月以上）のタイトルだけを使う。** 半期報告書のタイトルは
    会社によって「第47期(2025/07/01－2025/12/31)」と**半期の期間**を書く。
    これの終了月を決算期末月とみなすと6月期の会社が12月期になる
    （2026-09-13、4396 / 4495 で発見。直前四半期の判定が丸ごと狂う）。
    変則決算で1年未満の期のタイトルも同じ理由で使わない。
    """
    t = unicodedata.normalize("NFKC", title or "")
    m = _RE_SLASH.search(t)
    if m:
        months = ((int(m.group(4)) * 12 + int(m.group(5)))
                  - (int(m.group(1)) * 12 + int(m.group(2))) + 1)
        return int(m.group(5)) if months >= 11 else None
    m = _RE_WAREKI.search(title or "")
    if m:
        return int(m.group(1))
    return None


def _shift_bday(date_str, n, bdays):
    """営業日で n ずらす。窓の外に出たら端で止める（イベントを消さない）。"""
    if n == 0:
        return date_str
    lo, hi = 0, len(bdays) - 1
    while lo <= hi:
        mid = (lo + hi) // 2
        if bdays[mid] < date_str:
            lo = mid + 1
        else:
            hi = mid - 1
    j = min(max(lo + n, 0), len(bdays) - 1)
    return bdays[j]


def build(out_path, shift_days=0, src_con=None):
    """src_con を渡さなければ本体DBを開く。テストは一時DBを渡す。"""
    src = src_con if src_con is not None else C.init_db()
    src.row_factory = sqlite3.Row

    def q(sql, *args):
        return src.execute(sql, args).fetchall()

    div = adapter._div()
    if os.path.exists(out_path):
        os.remove(out_path)
    dst = sqlite3.connect(out_path)
    dst.executescript(DDL)
    stats = {}

    bdays = [r["date"] for r in
             q("SELECT DISTINCT date FROM daily_prices ORDER BY date")]

    # ---- universe -------------------------------------------------------
    fy_end = {}
    for r in q("SELECT code, title FROM filings WHERE source='edinet' "
               "AND subtype IN ('120','160') ORDER BY date DESC"):
        if r["code"] not in fy_end:
            m = fiscal_year_end_month(r["title"])
            if m:
                fy_end[r["code"]] = m
    uni = [(r["code"], r["name"], r["market"], r["sector"], fy_end[r["code"]], 1, 1)
           for r in q("SELECT code, name, market, sector FROM companies")
           if r["code"] in fy_end]
    dst.executemany(
        "INSERT OR REPLACE INTO universe (ticker, company_name, market, sector,"
        " fiscal_year_end, market_cap_ok, liquidity_ok) VALUES (?,?,?,?,?,?,?)", uni)
    stats["universe"] = len(uni)
    stats["universe_skipped_no_fy_end"] = q("SELECT COUNT(*) c FROM companies")[0]["c"] - len(uni)

    # ---- filings（PIT の土台。初出開示日 + その書類の doc_id） -----------
    # doc_id まで持つのは、**スコアの根拠になった期の書類**に人間を
    # 飛ばすため。期だけ持っていると「最新の開示」に飛ばすしかなく、
    # 根拠期と別の書類を「原文」と称することになる。
    first = {}
    for r in q("SELECT code, period, q_no, d, doc_id, source, pdf_path FROM ("
               " SELECT fc.code, fc.period, fc.q_no, f.date d, f.doc_id,"
               "        f.source, f.pdf_path,"
               "        ROW_NUMBER() OVER (PARTITION BY fc.code, fc.period, fc.q_no"
               "          ORDER BY f.date, f.id) rn"
               " FROM financials_cum fc JOIN filings f ON f.id=fc.filing_id"
               " WHERE fc.q_no IS NOT NULL AND fc.period IS NOT NULL) WHERE rn=1"):
        first[(r["code"], r["period"], r["q_no"])] = r

    # ---- quarterly_standalone（span_q=1 のみ） ---------------------------
    cells, meta = defaultdict(dict), {}
    for r in q("SELECT code, period, q_no, item, value, valid_flag, invalid_reason "
               "FROM financials_q WHERE span_q = 1"):
        key = (r["code"], r["period"], r["q_no"])
        if r["item"] in _FLOW and r["value"] is not None:
            cells[key][_FLOW[r["item"]]] = r["value"] / div
        m = meta.setdefault(key, {"valid": 1, "reason": None})
        if not r["valid_flag"]:
            m["valid"] = 0
            m["reason"] = m["reason"] or r["invalid_reason"]

    fil, qs = [], []
    for key in sorted(meta):
        code, period, q_no = key
        doc = first.get(key)               # src は接続変数。名前を食わない
        d = doc["d"] if doc else None
        qt, fy = adapter._qtype(q_no), adapter._fy(period)
        if not d or qt is None or fy is None:
            continue
        pe = adapter._period_end(period, q_no)
        m, c = meta[key], cells.get(key, {})
        fil.append((code, _shift_bday(d, shift_days, bdays), pe, qt, fy,
                    doc["source"] or "edinet", doc["doc_id"], doc["pdf_path"], 1))
        qs.append((code, fy, qt, pe, c.get("sales"), c.get("operating_profit"),
                   c.get("gross_profit"), None, None, c.get("operating_cf"),
                   m["valid"], m["reason"], 1, 1, adapter._period_end(period, q_no)))
    dst.executemany(
        "INSERT OR REPLACE INTO filings (ticker, filing_date, period_end,"
        " quarter_type, fiscal_year, source, doc_id, pdf_path, generation)"
        " VALUES (?,?,?,?,?,?,?,?,?)", fil)
    dst.executemany(
        "INSERT OR REPLACE INTO quarterly_standalone (ticker, fiscal_year,"
        " quarter_type, period_end, sales, operating_profit, gross_profit,"
        " cogs_quantity, cogs_price, operating_cf, is_valid, invalid_reason,"
        " generation, span_q, period_start)"
        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", qs)
    # ---- quarterly_standalone_all（全 span。span-matched の入力）--------
    cells_all, meta_all = defaultdict(dict), {}
    for r in q("SELECT code, period, q_no, item, value, valid_flag, "
               " invalid_reason, span_q FROM financials_q"):
        key = (r["code"], r["period"], r["q_no"], r["span_q"] or 1)
        if r["item"] in _FLOW and r["value"] is not None:
            cells_all[key][_FLOW[r["item"]]] = r["value"] / div
        m = meta_all.setdefault(key, {"valid": 1, "reason": None})
        if not r["valid_flag"]:
            m["valid"] = 0
            m["reason"] = m["reason"] or r["invalid_reason"]
    qs_all = []
    for key in sorted(meta_all):
        code, period, q_no, span = key
        qt, fy = adapter._qtype(q_no), adapter._fy(period)
        if qt is None or fy is None:
            continue
        pe = adapter._period_end(period, q_no)
        # 累計の起点。span_q だけでは「どこからの6ヶ月か」が決まらない。
        start_q = max(1, q_no - span + 1)
        ps = adapter._period_end(period, start_q)
        m, c = meta_all[key], cells_all.get(key, {})
        qs_all.append((code, fy, qt, pe, span, ps, c.get("sales"),
                       c.get("operating_profit"), c.get("gross_profit"),
                       c.get("operating_cf"), m["valid"], m["reason"], 1))
    dst.executemany(
        "INSERT OR REPLACE INTO quarterly_standalone_all (ticker, fiscal_year,"
        " quarter_type, period_end, span_q, period_start, sales, operating_profit,"
        " gross_profit, operating_cf, is_valid, invalid_reason, generation)"
        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", qs_all)
    stats["quarterly_standalone_all"] = len(qs_all)
    stats["qs_all_span2plus"] = sum(1 for x in qs_all if x[4] > 1)

    stats["filings"] = len(fil)
    stats["quarterly_standalone"] = len(qs)
    stats["quarterly_invalid"] = sum(1 for x in qs if not x[10])

    # ---- balance_sheet_items -------------------------------------------
    bs = [(r["code"], adapter._period_end(r["period"], r["q_no"]), r["item"],
           r["value"] / div, 1)
          for r in q("SELECT code, period, q_no, item, value FROM financials_q "
                     "WHERE valid_flag=1 AND value IS NOT NULL AND span_q=1 "
                     "AND item NOT IN ('revenue','operating_income','gross_profit',"
                     "'operating_cf')")
          if adapter._qtype(r["q_no"]) and adapter._fy(r["period"])]
    # **別枠の item_key に翻訳した行も入れる。** 投影層の仕事は
    # 「別枠の語彙で喋る」こと。ここを本体名のまま書いていたせいで
    # 別枠の get_bs_series(conn, t, "accounts_receivable") が空を引き、
    # **S1/S3/S6 が全銘柄で available=False になっていた**（2026-09-02 検出）。
    # 名前が違うだけで値は同じなので、本体名の行も残したまま別名を足す
    # （本体側から本体名で引く経路を壊さないため）。
    alias = {}
    for bkey, e in (adapter.cfg().get("items") or {}).items():
        if e.get("status") == "verified" and e.get("item") and e["item"] != bkey:
            alias.setdefault(e["item"], []).append(bkey)
    bs += [(t, pe, bkey, v, g) for (t, pe, item, v, g) in bs
           for bkey in alias.get(item, ())]
    dst.executemany("INSERT OR REPLACE INTO balance_sheet_items (ticker, period_end,"
                    " item_key, value, generation) VALUES (?,?,?,?,?)", bs)
    stats["balance_sheet_items"] = len(bs)
    stats["bs_aliased"] = sum(1 for x in bs if x[2] in
                              {k for ks in alias.values() for k in ks})

    # ---- pl_adjustments -------------------------------------------------
    adj = [(r["code"], adapter._period_end(r["period"], r["q_no"]), r["item_key"],
            (r["amount"] or 0) / div, r["note"], 1)
           for r in q("SELECT code, period, q_no, item_key, amount, note "
                      "FROM pl_adjustments")
           if r["period"] and r["q_no"]]
    dst.executemany("INSERT OR REPLACE INTO pl_adjustments (ticker, period_end,"
                    " item_key, amount, note, generation) VALUES (?,?,?,?,?,?)", adj)
    stats["pl_adjustments"] = len(adj)

    # ---- company_forecasts（売上と営業利益が揃った版だけ） ---------------
    g = defaultdict(dict)
    # 通期予想だけを採る。修正開示には中間期だけの修正があり（3161 は H1 のみ）、
    # それを通期予想として運ぶと S5 の進捗率の分母が壊れる。q_no が NULL の行は
    # 列を足す前に書かれたもので、当時の既定（通期）として扱う。
    for r in q("SELECT code, date, fy, item, value, revision_direction, prev_value, q_no "
               "FROM guidance WHERE item IN ('revenue','operating_income') "
               "AND value IS NOT NULL AND (q_no IS NULL OR q_no = 4)"):
        slot = g[(r["code"], r["fy"], r["date"])]
        slot[r["item"]] = r["value"]
        if r["item"] == "operating_income":
            slot["dir"] = r["revision_direction"]
            slot["prev_op"] = r["prev_value"]
    fc = [(k[0], adapter._fy(k[1]), v["revenue"] / div, v["operating_income"] / div, k[2],
           v.get("dir"), (v["prev_op"] / div) if v.get("prev_op") is not None else None)
          for k, v in g.items()
          if "revenue" in v and "operating_income" in v and adapter._fy(k[1])]
    dst.executemany("INSERT OR REPLACE INTO company_forecasts (ticker, fiscal_year,"
                    " forecast_sales, forecast_op, source_date, revision_direction,"
                    " prev_op) VALUES (?,?,?,?,?,?,?)", fc)
    stats["company_forecasts"] = len(fc)

    # ---- daily_prices / market_index（投影ビュー経由＝調整後終値） --------
    dst.executemany(
        "INSERT OR REPLACE INTO daily_prices (ticker, date, close, volume)"
        " VALUES (?,?,?,?)",
        [(r["ticker"], r["date"], r["close"], r["volume"])
         for r in q("SELECT ticker, date, close, volume FROM daily_prices")])
    stats["daily_prices"] = dst.execute("SELECT COUNT(*) FROM daily_prices").fetchone()[0]
    dst.executemany(
        "INSERT OR REPLACE INTO market_index (date, close) VALUES (?,?)",
        [(r["date"], r["close"]) for r in
         q("SELECT date, close FROM market_index WHERE close IS NOT NULL")])
    stats["market_index"] = dst.execute("SELECT COUNT(*) FROM market_index").fetchone()[0]

    dst.commit()
    dst.close()
    return stats


def rebuild_calendar(out_path, as_of=None):
    """発表日カレンダーを組み直す。**再生成のたびに呼ばないと空のままになる。**

    `earnings_calendar` は投影層が器だけ作り、中身は別枠 Module A が書く。
    build() は器を作り直すので、**再生成すると前回の中身が消える** ——
    2026-09-02 に実際に踏んだ（週次スクリーンが「カレンダー該当 0 銘柄」に
    なった）。「1コマンドで再生成できること」を満たすには、ここまで含めて
    1コマンドである必要がある。別枠のコードは呼ぶだけで変更しない。
    """
    from datetime import date
    from screener.report import backtest_eval as V1
    root = str(V1._external_root())
    if root not in sys.path:
        sys.path.insert(0, root)
    from module_a.earnings_calendar import build_calendar
    return len(build_calendar(out_path, as_of or date.today()))


# 事後条件ゲート。**「行はあるが中身がない」型の故障を機械で捕まえる。**
# この形の故障はエラーを出さない —— テーブルは存在し、クエリは成功し、
# 結果が空なだけなので、下流は「今週は該当なし」と区別がつかない。
# 2026-09-02 だけで3件踏んだ:
#   1. item_key の写像未適用で S1/S3/S6 が全銘柄 available=False
#   2. 再生成で earnings_calendar が空になり週次が「該当0銘柄」
#   3. row_factory 未設定で全候補 skip_score（過去の同型事故）
# いずれも「出力がゼロでも正常に見える」。だから閾値ではなく
# **非ゼロであること自体**を検査する。
SANITY_SIGNALS = ("S1", "S2", "S3", "S4", "S5")
SANITY_SAMPLE = 300
SANITY_CALENDAR_DAYS = 60


def sanity_check(out_path, as_of=None, sample=SANITY_SAMPLE):
    """生成物が「使える状態か」を検査する。違反のリストを返す（空なら合格）。"""
    import random
    from datetime import date, timedelta
    from screener.report import backtest_eval as V1
    root = str(V1._external_root())
    if root not in sys.path:
        sys.path.insert(0, root)
    from module_b.run_scorers import SCORERS_ALL
    from screener.signals.span_runner import score_ticker

    as_of = as_of or date.today()
    con = sqlite3.connect(out_path)
    con.row_factory = sqlite3.Row
    bad, note = [], {}

    hi = (as_of + timedelta(days=SANITY_CALENDAR_DAYS)).isoformat()
    try:
        n_cal = con.execute(
            "SELECT COUNT(*) FROM earnings_calendar WHERE next_earnings_date "
            "BETWEEN ? AND ?", (as_of.isoformat(), hi)).fetchone()[0]
    except sqlite3.OperationalError as e:
        n_cal, bad = 0, bad + ["earnings_calendar が引けない: %s" % e]
    note["calendar_%dd" % SANITY_CALENDAR_DAYS] = n_cal
    if n_cal <= 0:
        bad.append("発表日カレンダーが今後%d日で0件（build_calendar が走っていない）"
                   % SANITY_CALENDAR_DAYS)

    # **テーブルが無いこと自体を違反として返す。** ここで例外を投げると
    # 呼び出し側が「検査が落ちた」と「検査に落ちた」を区別できなくなる。
    try:
        tickers = [r[0] for r in con.execute(
            "SELECT DISTINCT ticker FROM quarterly_standalone_all")]
    except sqlite3.OperationalError as e:
        con.close()
        return bad + ["quarterly_standalone_all が引けない: %s" % e], note
    random.Random(0).shuffle(tickers)
    tickers = tickers[:sample]
    avail = {k: 0 for k in SCORERS_ALL}
    n_err = 0
    for t in tickers:
        try:
            # 配管検査は採点方針に依存させない（evidence_strict では S3 が構造的に 0% になり、
            # 「写像・配線の故障」と区別できなくなる）。2026-09-13 の既定切替後も prefer_span 固定。
            res = score_ticker(con, t, SCORERS_ALL, policy="prefer_span")
        except Exception:
            n_err += 1
            continue
        for k, v in res.items():
            if isinstance(v, dict) and v.get("available"):
                avail[k] = avail.get(k, 0) + 1
    n = max(1, len(tickers))
    note["available"] = {k: round(v * 100.0 / n, 1) for k, v in sorted(avail.items())}
    note["score_errors"] = n_err
    if n_err > n * 0.05:
        bad.append("スコア計算の例外が %d/%d 件（データ不足ではない）" % (n_err, n))
    for k in SANITY_SIGNALS:
        if avail.get(k, 0) <= 0:
            bad.append("%s の available 率が 0%%（n=%d）。写像か配線を疑う" % (k, n))

    # ---- 本体側の派生テーブル（2026-09-02 追加分） ------------------
    # **「行はあるが中身がない」型の故障をここでも捕まえる。**
    # 新しいパース・フラグ系は対象0件でも例外を出さずに全滅しうる。
    try:
        mcon = C.init_db()
    except Exception as e:                              # pragma: no cover
        mcon, bad = None, bad + ["本体DBを開けない: %s" % e]
    if mcon is not None:
        for tbl, key, label in (
                ("disclosure_flags", "going_concern", "開示フラグ"),
                ("s13_orders", "available", "S13受注残"),
        ):
            try:
                n_all = mcon.execute("SELECT COUNT(*) FROM %s" % tbl).fetchone()[0]
                n_hit = mcon.execute(
                    "SELECT COUNT(*) FROM %s WHERE %s=1" % (tbl, key)).fetchone()[0]
            except sqlite3.OperationalError as e:
                bad.append("%s: %s" % (label, e))
                continue
            note["%s" % tbl] = "%d行 / %s=1 は %d (%.1f%%)" % (
                n_all, key, n_hit, n_hit * 100.0 / max(n_all, 1))
            if n_all == 0:
                bad.append("%s が0行（パースが走っていない）" % label)
            elif n_hit == 0:
                bad.append("%s の %s が全件0（静かに全滅した疑い）" % (label, key))
        try:
            n_ev = mcon.execute("SELECT COUNT(*) FROM s12_evidence").fetchone()[0]
            n_cls = mcon.execute(
                "SELECT COUNT(*) FROM s12_evidence "
                "WHERE paragraph_class IS NOT NULL").fetchone()[0]
            note["s12_paragraph_class"] = "%d/%d 行に段落種別あり" % (n_cls, n_ev)
            if n_ev and n_cls == 0:
                bad.append("s12_evidence の段落種別が全件 NULL（再計算前）")
        except sqlite3.OperationalError as e:
            bad.append("s12_evidence: %s" % e)

    # ---- 鮮度（2026-09-02 追加） --------------------------------------
    # **今回の穴はカレンダーも行数も候補数も全部素通りした。**
    # 「行はあるが古い」は「行はあるが空」と同じ型の故障で、件数の検査では
    # 絶対に捕まらない。鮮度は独立した事後条件として要る。
    #
    # 決算シーズン（3-5月/6-7月/9月/12月）に**ユニバース全体で30日以上
    # 新しい開示が1件も入っていない**なら取り込みが止まっている。
    # 銘柄単位で見ないのは、決算期によって正当に間隔が空くため
    # （1月期は TDnet 保持切れで132日空くが、それは別の既知の穴）。
    SEASON_MONTHS = (3, 4, 5, 6, 7, 9, 12)
    STALE_LIMIT_DAYS = 30
    if mcon is not None:
        try:
            newest = mcon.execute(
                "SELECT MAX(f.date) FROM filings f "
                "JOIN financials_cum fc ON fc.filing_id=f.id").fetchone()[0]
        except sqlite3.OperationalError as e:
            newest, bad = None, bad + ["最新パース済み開示日が取れない: %s" % e]
        if newest:
            gap = (as_of - date.fromisoformat(newest)).days
            note["newest_parsed"] = "%s（%d日前）" % (newest, gap)
            if as_of.month in SEASON_MONTHS and gap > STALE_LIMIT_DAYS:
                bad.append("決算シーズン(%d月)なのに最新パース済み開示が %d日前"
                           "（%s）。取り込みが止まっている" % (as_of.month, gap, newest))
        # 決算期末月ごとの遅れも出す。**落とさないが見せる。**
        try:
            rx = re.compile(r"\((\d{4})/(\d{2})/(\d{2})－(\d{4})/(\d{2})/(\d{2})\)")
            fym, lastd = {}, {}
            for r in mcon.execute(
                    "SELECT code, title FROM filings "
                    "WHERE title LIKE '%有価証券報告書%' ORDER BY date DESC"):
                if r[0] not in fym:
                    m2 = rx.search(r[1] or "")
                    if m2:
                        fym[r[0]] = int(m2.group(5))
            for r in mcon.execute(
                    "SELECT f.code, MAX(f.date) d FROM filings f "
                    "JOIN financials_cum fc ON fc.filing_id=f.id GROUP BY f.code"):
                lastd[r[0]] = r[1]
            buckets = {}
            for code, mo in fym.items():
                d = lastd.get(code)
                if d:
                    buckets.setdefault(mo, []).append(
                        (as_of - date.fromisoformat(d)).days)
            note["stale_by_fy_month"] = {
                mo: "n=%d 中央%d日" % (len(v), sorted(v)[len(v) // 2])
                for mo, v in sorted(buckets.items())}
        except Exception as e:                          # pragma: no cover
            note["stale_by_fy_month"] = "算出できず: %s" % e

    n_cand = 0
    for t in tickers:
        try:
            s = score_ticker(con, t, SCORERS_ALL, policy="prefer_span").get("evidence_score")
        except Exception:
            continue
        if s is not None and s >= 0.10:
            n_cand += 1
    note["candidates_at_0.10"] = n_cand
    if n_cand <= 0:
        bad.append("スコア0.10以上の候補が0件（n=%d）。全銘柄が無評価の疑い" % n)
    con.close()
    return bad, note


def out_path_for(base, shift_days):
    if not shift_days:
        return base
    root, ext = os.path.splitext(base)
    return "%s_%+d%s" % (root, shift_days, ext)


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="別枠形の投影DBを本体から生成する")
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "projection.db"))
    p.add_argument("--shift-days", type=int, default=0,
                   help="開示日を営業日で前後にずらす。代理発表日の誤差に対する"
                        "感度分析用 (-3 / 0 / 3)")
    p.add_argument("--no-calendar", action="store_true",
                   help="発表日カレンダーの再構築を省く（既定は再構築する）")
    p.add_argument("--no-sanity", action="store_true",
                   help="事後条件ゲートを省く（既定は検査して違反なら exit 1）")
    p.add_argument("--sanity-sample", type=int, default=SANITY_SAMPLE,
                   help="事後条件の抜き取り銘柄数")
    a = p.parse_args(argv)
    out = out_path_for(a.out, a.shift_days)
    C.log("投影DB生成: %s (shift_days=%+d)" % (out, a.shift_days))
    st = build(out, a.shift_days)
    if not a.no_calendar:
        st["earnings_calendar"] = rebuild_calendar(out)
    for k, v in st.items():
        C.log("  %-32s %10s" % (k, format(v, ",")))
    C.log("  サイズ %.0f MB" % (os.path.getsize(out) / 1024 / 1024))
    if a.no_sanity:
        C.log("  事後条件ゲート: スキップ（--no-sanity）")
        return 0
    bad, note = sanity_check(out, sample=a.sanity_sample)
    C.log("  事後条件ゲート (n=%d):" % a.sanity_sample)
    for k, v in note.items():
        C.log("    %-22s %s" % (k, v))
    if bad:
        for b in bad:
            C.log("  ! 事後条件ちがい: %s" % b)
        C.log("  ** 生成物は使える状態ではない。exit 1 **")
        return 1
    C.log("  事後条件: すべて満たす")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
