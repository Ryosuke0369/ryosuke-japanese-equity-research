"""screener/report/earnings_window.py — 決算窓フィルタ（週次CSVを発表予定日の区間で絞る後処理）。

なぜ正式なスクリプトにしたか
----------------------------
2026-09-13 に「今月中に発表がある銘柄だけを見る」をその場の即興スクリプトで
行った。即興だと毎回、発表日の根拠の優先順・休場日・期ラベルの暦変換を
組み直すことになり、実際に期ラベル（FY2026-Q2）を暦日と比較する誤りを
一度踏んだ。**同じ誤りを二度踏まないために1本に固定する。**

発表予定日の根拠（優先順・行ごとに `basis` 列へ明示）
---------------------------------------------------
  1   TDnet 適時開示（決算発表日のお知らせ等）  disclosure_titles のタイトル
  1b  JPX「決算発表予定日一覧」(会社届出)       raw/jpx_schedule/*.xlsx
  2   前年同期の実際の発表日 + 364日            J-Quants /fins/summary キャッシュ
  3   システム est_date（projection.db）         confidence LOW が多く単独では信用しない
1 は日付をタイトルから確定できないので、日付は 1b 以下から取り、1 は存在を注記する。

並べ替え
--------
第一基準: **発表される四半期の直前四半期が本体DB（financials_cum）に揃っているか**。
直前四半期が無い銘柄のスコアは、発表をまたぐ証拠として先行性を持たない。
第二基準: 推定発表日。第三基準: スコア。stale_flag は補助列として出すだけ。

休場日
------
別枠 `common/jp_calendar.py`（祝日・振替・国民の休日）で判定し、推定日が
休場日なら翌営業日に寄せて `est_business_day=0` の注記を残す。

    python -m screener.report.earnings_window --weekly-csv C:/screener_data/weekly_20260913_full_rescored.csv \
        --as-of 2026-09-13 --from 2026-09-13 --to 2026-09-30 --mid 2026-09-18 \
        --out C:/screener_data/earnings_window_20260913_0930.csv
    python -m screener.report.earnings_window --fetch-jpx          # JPX 一覧を raw/jpx_schedule に保存
    python -m screener.report.earnings_window --fetch-jquants --as-of 2026-09-13 --from 2026-09-13 --to 2026-09-30
"""
from __future__ import annotations

import argparse
import csv
import glob
import json
import os
import re
import sqlite3
import sys
import time
from collections import Counter
from datetime import date, datetime, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

JPX_DIR = os.path.join(C.RAW_DIR, "jpx_schedule")
JQ_CACHE = os.path.join(C.CACHE_DIR, "jquants_fins_summary")
JPX_PAGE = "https://www.jpx.co.jp/listing/event-schedules/financial-announcement/index.html"
JPX_BASE = "https://www.jpx.co.jp"

QMAP_JPX = {"第１四半期": "1Q", "第２四半期": "2Q", "第３四半期": "3Q", "本決算": "FY",
            "第1四半期": "1Q", "第2四半期": "2Q", "第3四半期": "3Q"}
QLABEL = {"1Q": "第1四半期", "2Q": "第2四半期", "3Q": "第3四半期", "FY": "本決算(Q4)"}
Q_NO = {"1Q": 1, "2Q": 2, "3Q": 3, "FY": 4}
TITLE_NOTICE = re.compile(r"決算発表日|発表予定日|発表日の(変更|決定)|決算短信の開示延期")

# 週次CSVから引き継ぐ列（存在する列だけ）
PASS_COLS = ("score", "score_base", "s12", "fired", "divergence", "doc_kind", "doc_period",
             "doc_date", "stale_flag", "stale_lag", "evidence_recency", "evidence_lag_q",
             "false_positive_flags", "scope_change_suspect", "reliability_flag",
             "going_concern", "accounting_change", "comparability_caution",
             "score_category", "no_score_reason", "doc_url")


# ------------------------------------------------------------------ 暦
def _jp_calendar():
    from screener.report import backtest_eval as V1
    root = str(V1._external_root())
    if root not in sys.path:
        sys.path.insert(0, root)
    from common import jp_calendar
    return jp_calendar


def ym_add(ym, months):
    """(年, 月) に月数を足す。"""
    y, m = ym
    t = y * 12 + (m - 1) + months
    return (t // 12, t % 12 + 1)


def fiscal_label_ym(period, q_no, fy_end_month):
    """本体の期ラベル FY{Y}-Q{q} → 期末の (年, 月)。

    FY{Y} は「期末が Y 年 fy_end_month 月の会計年度」
    （3441 7月期 FY2025-Q4 = 2025年7月、有報 第67期 2024/08-2025/07 で確認）。
    **期ラベルを暦日文字列と直接比べてはいけない**（2026-09-13 に踏んだ）。
    """
    m = re.match(r"FY(\d{4})", period or "")
    if not m or not fy_end_month or not q_no:
        return None
    return ym_add((int(m.group(1)), int(fy_end_month)), -3 * (4 - int(q_no)))


def roll_business_day(d, cal=None):
    """休場日なら翌営業日へ。(寄せた日, 元が営業日か)。"""
    cal = cal or _jp_calendar()
    if cal.is_business_day(d):
        return d, True
    x = d
    while not cal.is_business_day(x):
        x += timedelta(days=1)
    return x, False


# ------------------------------------------------------------------ 取得
def fetch_jpx(fetcher=None):
    """JPX の決算発表予定日一覧（xlsx）を raw/jpx_schedule/ に保存する。"""
    os.makedirs(JPX_DIR, exist_ok=True)
    fetcher = fetcher or C.Fetcher(min_interval=1.0)
    r = fetcher.get(JPX_PAGE)
    html = r.content.decode("utf-8", "replace")
    links = sorted(set(re.findall(r'href="([^"]+/kessan\d{2}_\d{4}\.xlsx)"', html)))
    saved = []
    for href in links:
        dest = os.path.join(JPX_DIR, os.path.basename(href))
        if not os.path.exists(dest):
            fetcher.download(JPX_BASE + href if href.startswith("/") else href, dest)
        saved.append(dest)
    C.log("JPX 決算発表予定日一覧: %d 件 -> %s" % (len(saved), JPX_DIR))
    return saved


def fetch_jquants(start, end, cache_dir=JQ_CACHE):
    """J-Quants /fins/summary を日付ごとに取得してキャッシュする（既取得日は飛ばす）。"""
    import requests
    from screener.fetch import jquants_universe as J
    os.makedirs(cache_dir, exist_ok=True)
    key = J.jq_api_key()
    d, n_new = start, 0
    while d <= end:
        path = os.path.join(cache_dir, "%s.json" % d.isoformat())
        if d.weekday() < 5 and not os.path.exists(path):
            rows, pk = [], None
            while True:
                params = {"date": d.isoformat()}
                if pk:
                    params["pagination_key"] = pk
                for _ in range(4):
                    r = requests.get(J.JQ + "/fins/summary", headers={"x-api-key": key},
                                     params=params, timeout=60)
                    if r.status_code != 429:
                        break
                    time.sleep(20)
                r.raise_for_status()
                j = r.json()
                rows += [{k: x.get(k) for k in ("DiscDate", "DiscTime", "Code", "DocType",
                                                 "CurPerType", "CurPerEn", "CurFYEn")}
                         for x in j.get("data", [])]
                pk = j.get("pagination_key")
                time.sleep(1.1)
                if not pk:
                    break
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(rows, fh, ensure_ascii=False)
            n_new += 1
        d += timedelta(days=1)
    C.log("J-Quants fins/summary: %d 日ぶん新規取得 -> %s" % (n_new, cache_dir))
    # **頼んだ範囲が実際に埋まったか**を、ファイルの有無で突き合わせる（§53）。
    have = {os.path.basename(f)[:-5] for f in glob.glob(os.path.join(cache_dir, "*.json"))}
    C.verify_range("決算サマリーのキャッシュ", start, end, have)


# ------------------------------------------------------------------ 読み込み
def load_jpx(as_of, jpx_dir=JPX_DIR):
    """code -> {date, qt, target_ym, file}。同じ kessanMM は新しいファイルを優先。"""
    import openpyxl
    files = {}
    for f in sorted(glob.glob(os.path.join(jpx_dir, "kessan*.xlsx"))):
        m = re.match(r"kessan(\d{2})_(\d{4})\.xlsx", os.path.basename(f))
        if m:
            files[m.group(1)] = f               # sorted → 同月は後勝ち（MMDD 昇順）
    out = {}
    for mm, f in sorted(files.items()):
        month = int(mm)
        year = as_of.year if month <= as_of.month else as_of.year - 1
        ws = openpyxl.load_workbook(f, read_only=True).worksheets[0]
        for row in ws.iter_rows(values_only=True):
            if not row or not isinstance(row[0], datetime):
                continue
            qt = QMAP_JPX.get(str(row[7] or "").strip())
            fye = row[4] if isinstance(row[4], datetime) else None
            out[str(row[1]).strip()] = {
                "date": row[0].date(), "qt": qt, "target_ym": (year, month),
                "fy_end_ym": (fye.year, fye.month) if fye else None,
                "file": os.path.basename(f)}
    return out


def load_jquants(cache_dir=JQ_CACHE):
    rows = []
    for f in sorted(glob.glob(os.path.join(cache_dir, "*.json"))):
        with open(f, encoding="utf-8") as fh:
            rows += json.load(fh)
    return rows


def _c4(code):
    s = str(code or "").strip()
    return C.normalise_code(s) if len(s) == 5 else s


# ------------------------------------------------------------------ 本体
def build(as_of, lo, hi, mid=None, weekly_rows=None, mcon=None, pcon=None,
          jpx=None, jq_rows=None, universe_only=True, cal=None):
    cal = cal or _jp_calendar()
    mcon = mcon or C.init_db()
    mcon.row_factory = sqlite3.Row
    if pcon is None:
        pdb = os.path.join(C.DATA_DIR, "projection.db").replace("\\", "/")
        pcon = sqlite3.connect("file:%s?mode=ro" % pdb, uri=True)
        pcon.row_factory = sqlite3.Row
    jpx = load_jpx(as_of) if jpx is None else jpx
    jq_rows = load_jquants() if jq_rows is None else jq_rows
    weekly = {r["code"]: r for r in (weekly_rows or [])}

    comp_sql = "SELECT code, name, market FROM companies" + (
        " WHERE universe_flag=1" if universe_only else "")
    comps = {r["code"]: dict(r) for r in mcon.execute(comp_sql)}
    fym = {r["ticker"]: r["fiscal_year_end"]
           for r in pcon.execute("SELECT ticker, fiscal_year_end FROM universe")}
    sysc = {r["ticker"]: dict(r) for r in pcon.execute("SELECT * FROM earnings_calendar")}

    notices = {}
    try:
        for r in mcon.execute("SELECT code, date, title FROM disclosure_titles "
                              "WHERE code IS NOT NULL AND date<=?", (as_of.isoformat(),)):
            if TITLE_NOTICE.search(r["title"] or ""):
                notices.setdefault(r["code"], []).append("%s %s" % (r["date"], r["title"][:50]))
    except sqlite3.OperationalError:
        pass

    # J-Quants: 前年同期（期末が target_ym - 12ヶ月）と当期既発表
    jq_by_code = {}
    for x in jq_rows:
        if "FinancialStatements" not in (x.get("DocType") or ""):
            continue
        pe = x.get("CurPerEn") or ""
        if len(pe) < 7:
            continue
        jq_by_code.setdefault(_c4(x["Code"]), []).append(
            (x["DiscDate"], (int(pe[:4]), int(pe[5:7])), x.get("CurPerType")))

    # 窓の月: lo の2ヶ月前〜 hi の月（四半期末から45日以内が法定の目安）
    lo_ym, hi_ym = ym_add((lo.year, lo.month), -2), (hi.year, hi.month)

    rows = []
    for code, comp in comps.items():
        j, s, jqs = jpx.get(code), sysc.get(code), jq_by_code.get(code, [])
        # 対象期（期末 ym）と四半期種別
        target_ym = qt = None
        if j:
            target_ym, qt = j["target_ym"], j["qt"]
        prev = None
        cand_prev = [x for x in jqs if x[1] and ym_add(x[1], 12) >= lo_ym
                     and ym_add(x[1], 12) <= hi_ym and x[0] < as_of.isoformat()]
        if target_ym:
            cand_prev = [x for x in jqs if ym_add(x[1], 12) == target_ym]
        if cand_prev:
            prev = sorted(cand_prev)[0]
            if not target_ym:
                target_ym, qt = ym_add(prev[1], 12), prev[2]
        if not target_ym and s:
            qt = s["quarter_type"]
        if not (j or prev or (s and lo.isoformat() <= s["next_earnings_date"] <= hi.isoformat())):
            continue

        f_month = fym.get(code) or (j["fy_end_ym"][1] if j and j.get("fy_end_ym") else None)
        # 既発表（当期）: J-Quants の同じ期末 or 本体 TDnet 短信
        announced = None
        if target_ym:
            hit = [x for x in jqs if x[1] == target_ym and x[0] <= as_of.isoformat()]
            if hit:
                announced = sorted(hit)[0][0]
            else:
                nxt = ym_add(target_ym, 1)
                r = mcon.execute(
                    "SELECT MIN(date) FROM filings WHERE code=? AND source='tdnet' "
                    "AND subtype='決算短信' AND title NOT LIKE '%訂正%' AND date>=? AND date<=?",
                    (code, "%04d-%02d-01" % nxt, as_of.isoformat())).fetchone()
                announced = r[0] if r and r[0] else None

        basis, est, conf, notes = None, None, None, []
        if notices.get(code):
            basis, conf = "1:TDnet適時開示(タイトル)", "開示あり(日付は下位根拠)"
            notes.append(" / ".join(notices[code][:3]))
        if j:
            notes.append("JPX=%s(%s)" % (j["date"], j["file"]))
            if est is None:
                est = j["date"]
                basis = basis or "1b:JPX決算発表予定日一覧(会社届出)"
                conf = conf or "会社届出"
        if prev:
            d0 = date.fromisoformat(prev[0])
            e2, _ok = roll_business_day(d0 + timedelta(days=364), cal)
            notes.append("前年実績=%s→%s" % (prev[0], e2))
            if est is None:
                est, basis, conf = e2, "2:前年同期実績から推定", "推定(前年1点)"
        if s:
            notes.append("est_date=%s %s %s" % (s["next_earnings_date"], s["quarter_type"],
                                               s["confidence_level"]))
            if est is None:
                est = date.fromisoformat(s["next_earnings_date"])
                basis, conf = "3:システムest_date(単独)", s["confidence_level"]
        est_bd = None
        if est:
            est2, est_bd = roll_business_day(est, cal)
            if not est_bd:
                notes.append("推定日%sは休場日→%s" % (est, est2))
            est = est2

        # 直前四半期が本体に揃っているか
        prev_ym = ym_add(target_ym, -3) if target_ym else None
        have = {}
        for r in mcon.execute("SELECT fc.period, fc.q_no, MIN(f.date) d FROM financials_cum fc "
                              "JOIN filings f ON f.id=fc.filing_id WHERE fc.code=? "
                              "AND fc.q_no IS NOT NULL GROUP BY fc.period, fc.q_no", (code,)):
            ym = fiscal_label_ym(r["period"], r["q_no"], f_month)
            if ym and (ym not in have or r["d"] < have[ym][1]):
                have[ym] = ("%s-Q%s" % (r["period"], r["q_no"]), r["d"])
        latest = max(have) if have else None
        prev_in_db = None if prev_ym is None or not f_month else int(prev_ym in have)
        gap = ""
        if prev_in_db == 0:
            gap = "直前四半期(%04d-%02d期末)が本体に無い" % prev_ym
            if f_month == 1:
                gap += " / 1月期Q1欠落(TDnet保持窓外)"
            elif f_month == 7:
                gap += " / 7月期Q3欠落(TDnet保持窓外)"

        in_window = bool(est and lo <= est <= hi and not announced)
        row = {
            "code": code, "name": comp["name"], "market": comp["market"],
            "fy_end_month": f_month or "",
            # 「YYYY-MM期末」= 発表される四半期の期末月。決算期（7月期等）とは別物なので
            # 「2026年7月期」と読める書き方をしない。
            "period": ("%04d-%02d期末 %s" % (target_ym[0], target_ym[1], QLABEL.get(qt, qt or "?"))
                       if target_ym else "?"),
            "target_period_ym": "%04d-%02d" % target_ym if target_ym else "",
            "prev_quarter_ym": "%04d-%02d" % prev_ym if prev_ym else "",
            "prev_quarter_in_db": "" if prev_in_db is None else prev_in_db,
            "est_date": est.isoformat() if est else "",
            "est_business_day": "" if est_bd is None else int(est_bd),
            "basis": basis or "", "confidence": conf or "",
            "already_announced": announced or "",
            "in_window": int(in_window),
            "in_mid": int(bool(in_window and mid and est <= mid)),
            "tdnet_gap_impact": gap,
            "latest_parsed_period": have[latest][0] if latest else "",
            "latest_parsed_ym": "%04d-%02d" % latest if latest else "",
            "latest_parsed_filing_date": have[latest][1] if latest else "",
            "sys_est_date": s["next_earnings_date"] if s else "",
            "sys_confidence": s["confidence_level"] if s else "",
            "in_weekly_csv": int(code in weekly),
            "notes": " ｜ ".join(notes),
        }
        w = weekly.get(code, {})
        for k in PASS_COLS:
            row[k] = w.get(k, "")
        rows.append(row)
    return rows


def sort_key(r):
    """第一: 直前四半期が本体にある(1)→無い(0)→判定不能('')。第二: 推定日。第三: スコア降順。"""
    pq = r.get("prev_quarter_in_db")
    rank = 0 if pq == 1 else (1 if pq == 0 else 2)
    try:
        sc = -float(r.get("score"))
    except (TypeError, ValueError):
        sc = 9.0
    return (rank, r.get("est_date") or "9999", sc, r["code"])


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--weekly-csv", help="weekly_screen の CSV（無ければカレンダーだけ出す）")
    p.add_argument("--as-of", help="基準日 YYYY-MM-DD（既定は今日）")
    p.add_argument("--from", dest="dfrom", help="窓の開始日（既定は as-of）")
    p.add_argument("--to", dest="dto", help="窓の終了日（既定は as-of+17日）")
    p.add_argument("--mid", help="中間集計日（例: 2026-09-18）")
    p.add_argument("--out", help="出力CSV")
    p.add_argument("--scored-only", action="store_true", help="週次CSVに行がある銘柄だけ出す")
    p.add_argument("--all-universe", action="store_true", help="universe_flag=0 も含める")
    p.add_argument("--fetch-jpx", action="store_true", help="JPX 一覧を取得して終了")
    p.add_argument("--fetch-from", help="決算サマリーの取得開始日（--fetch-jquants と併用）")
    p.add_argument("--fetch-to", help="決算サマリーの取得終了日（--fetch-jquants と併用）")
    p.add_argument("--fetch-jquants", action="store_true",
                   help="前年同期と当期既発表の判定に要る J-Quants 日次データを取得して終了")
    a = p.parse_args(argv)

    as_of = date.fromisoformat(a.as_of) if a.as_of else date.today()
    lo = date.fromisoformat(a.dfrom) if a.dfrom else as_of
    hi = date.fromisoformat(a.dto) if a.dto else as_of + timedelta(days=17)
    mid = date.fromisoformat(a.mid) if a.mid else None
    if a.fetch_jpx:
        fetch_jpx()
        return 0
    if a.fetch_jquants:
        # 既定は「窓の前年同期（−400〜−330日）」と「直近40日」。**--from/--to は窓の指定であって
        # 取得範囲ではない**（2026-09-20、これを取り違えて別の期間を埋めてしまった）。
        # 任意の期間を埋めたいときは --fetch-from/--fetch-to を明示する。
        if a.fetch_from or a.fetch_to:
            f0 = date.fromisoformat(a.fetch_from) if a.fetch_from else as_of - timedelta(days=40)
            f1 = date.fromisoformat(a.fetch_to) if a.fetch_to else as_of
            C.log("J-Quants 決算サマリー: 指定された期間 %s 〜 %s を取得する" % (f0, f1))
            fetch_jquants(f0, f1)
            return 0
        C.log("J-Quants 決算サマリー: 前年同期 %s 〜 %s と 直近40日を取得する"
              % (lo - timedelta(days=400), hi - timedelta(days=330)))
        fetch_jquants(lo - timedelta(days=400), hi - timedelta(days=330))
        fetch_jquants(as_of - timedelta(days=40), as_of)
        return 0

    weekly_rows = []
    if a.weekly_csv:
        with open(a.weekly_csv, encoding="utf-8-sig") as fh:
            weekly_rows = list(csv.DictReader(fh))
    rows = build(as_of, lo, hi, mid, weekly_rows, universe_only=not a.all_universe)
    sel = [r for r in rows if r["in_window"] and (r["in_weekly_csv"] or not a.scored_only)]
    sel.sort(key=sort_key)

    cal = _jp_calendar()
    hol = [lo + timedelta(days=i) for i in range((hi - lo).days + 1)]
    hol = [d for d in hol if d.weekday() < 5 and not cal.is_business_day(d)]
    C.log("決算窓 %s〜%s（as_of %s）休場日(平日): %s" % (lo, hi, as_of, ", ".join(map(str, hol)) or "なし"))
    C.log("  候補(何らかの根拠あり) %d / 窓内(未発表) %d%s / 既発表で除外 %d"
          % (len(rows), len(sel),
             (" / うち %s まで %d" % (mid, sum(r["in_mid"] for r in sel))) if mid else "",
             sum(1 for r in rows if r["already_announced"])))
    C.log("  根拠: %s" % dict(Counter(r["basis"] for r in sel)))
    C.log("  直前四半期が本体にある: %s" % dict(Counter(str(r["prev_quarter_in_db"]) for r in sel)))
    C.log("  週次CSVに行がある: %d" % sum(r["in_weekly_csv"] for r in sel))
    if a.out:
        with open(a.out, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(sel[0].keys()) if sel else ["code"])
            w.writeheader()
            w.writerows(sel)
        C.log("  CSV: %s (%d 行)" % (a.out, len(sel)))
    # 成功条件: 窓内が1件以上、または明示的に 0 件であることが表示されている
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
