"""screener/fetch/tdnet_archiver.py - daily TDnet 適時開示 archiver (仕様書 §2-1).

Why this runs every weekday: the TDnet public listing keeps only about one
month. Anything not pulled today is gone for good, so the archive is the one
part of this system where a missed day is unrecoverable. That is the whole
reason fetch_runs exists — a day is either saved or visibly recorded as missing.

What it saves
    決算短信(四半期含む) / 業績予想の修正 / 配当予想の修正
    PDF (always) + XBRL zip (when TDnet publishes one), under
    <DATA_ROOT>/raw/tdnet/YYYYMMDD/, with metadata in `filings`.

Usage
    python -m screener.fetch.tdnet_archiver --today
    python -m screener.fetch.tdnet_archiver --date 20260828
    python -m screener.fetch.tdnet_archiver --days 3          # last 3 weekdays
    python -m screener.fetch.tdnet_archiver --backfill 30     # fill gaps, 30d window
    python -m screener.fetch.tdnet_archiver --report 30       # coverage only, no fetch
    python -m screener.fetch.tdnet_archiver --all-types ...   # archive every disclosure

Exit codes: 0 = every requested day is ok/empty, 1 = at least one day failed or
is still missing (so the Task Scheduler entry shows red and the retry matters).
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import date, datetime, timedelta

from bs4 import BeautifulSoup

try:
    from screener import common as C
except ImportError:                                     # run as a plain script
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

SOURCE = "tdnet"
LIST_URL = "https://www.release.tdnet.info/inbs/I_list_{page:03d}_{ymd}.html"
DOC_URL = "https://www.release.tdnet.info/inbs/{name}"
MAX_PAGES = 40                                          # ~4,000 rows/day ceiling

# Disclosure classification. Order matters: the first match wins, and 訂正 is
# checked inside each branch rather than as a separate type so that a corrected
# 短信 still lands in the 短信 bucket (it carries the same XBRL).
SUBTYPE_RULES = (
    ("決算短信",     re.compile(r"決算短信")),
    ("業績予想修正", re.compile(r"(業績予想|業績見通し).*(修正|変更)|通期予想の修正")),
    ("配当予想修正", re.compile(r"配当予想.*(修正|変更)|(修正|変更).*配当予想|配当予想の修正")),
    # 決算説明会資料（2026-09-02 追加）。順序は最後 —— 「決算説明会資料」は
    # 「決算短信」を含まないが、訂正版のタイトルが両方を含むことがあるので、
    # 既存3種別の判定を先に通してから拾う。既存の分類は1件も変わらない。
    #
    # 用途はテキスト差分(S12ファミリー)の材料の**収集のみ**。スコアには
    # 使わない。TDnet は約6週間しか保持しないので前年ペアが揃うのは
    # 2027年秋以降。それまでは収集と抽出率のテレメトリだけを回す。
    # 実測(2026-09-02, 5営業日): 全開示975件中60件(6.2%)、
    # ユニバース1,423銘柄に対するカバレッジは日次1.6%、年換算 約2,940件。
    # 「決算」「業績」の文脈を必須にする。素の「補足説明資料」まで拾うと
    # 「第三者割当による新株予約権の発行に関する補足説明資料」のような
    # 資金調達の資料が混ざり、テキスト差分の母集団が汚れる。
    ("決算説明資料", re.compile(
        r"(決算|業績|four|中間|通期).{0,8}(説明|補足)"
        r"|説明会資料|ファクトブック|決算プレゼンテーション")),
)
# §5 filings.type vocabulary
TYPE_OF_SUBTYPE = {"決算短信": "短信", "業績予想修正": "修正",
                   "配当予想修正": "修正", "決算説明資料": "説明"}


def classify(title: str):
    """(subtype, type) or (None, None) when the row is outside the §2-1 scope."""
    t = (title or "").replace("　", " ")
    for subtype, rx in SUBTYPE_RULES:
        if rx.search(t):
            return subtype, TYPE_OF_SUBTYPE[subtype]
    return None, None


# ------------------------------------------------------------------ listing
def parse_list_page(html: str) -> tuple[list[dict], int]:
    """Rows on one TDnet listing page + the total row count from the pager.

    The pager text is '1～100件 / 全227件'; total is what tells us how many
    pages to walk. Returning 0 rows with total 0 is a legitimate answer (a
    Saturday), and is why the caller distinguishes 'empty' from 'failed'.
    """
    soup = BeautifulSoup(html, "lxml")
    total = 0
    pager = soup.find(id="pager-box-top")
    if pager:
        m = re.search(r"全\s*([\d,]+)\s*件", pager.get_text(" ", strip=True))
        if m:
            total = int(m.group(1).replace(",", ""))

    rows = []
    for tr in soup.select("table#main-list-table tr"):
        tds = tr.find_all("td")
        if len(tds) < 6:
            continue
        a_pdf = tds[3].find("a")
        a_zip = tds[4].find("a")
        rows.append({
            "time": tds[0].get_text(strip=True),
            "code_raw": tds[1].get_text(strip=True),
            "name": tds[2].get_text(strip=True),
            "title": tds[3].get_text(strip=True),
            "pdf": a_pdf.get("href") if a_pdf else None,
            "zip": a_zip.get("href") if a_zip else None,
            "place": tds[5].get_text(strip=True),
        })
    return rows, total


def fetch_day_index(fetcher: "C.Fetcher", d: date) -> tuple[list[dict], int]:
    ymd = d.strftime("%Y%m%d")
    all_rows: list[dict] = []
    total = None
    for page in range(1, MAX_PAGES + 1):
        r = fetcher.get(LIST_URL.format(page=page, ymd=ymd))
        if r.status_code == 404:
            break
        if r.status_code != 200:
            raise RuntimeError(f"listing page {page} for {ymd}: HTTP {r.status_code}")
        rows, tot = parse_list_page(r.content.decode("utf-8", "replace"))
        if total is None:
            total = tot
        all_rows.extend(rows)
        if not rows or len(all_rows) >= (total or 0):
            break
    return all_rows, (total if total is not None else len(all_rows))


# ------------------------------------------------------------------ archive
def day_is_final(d: date, now: datetime | None = None) -> bool:
    """対象日が終わってから取得したか。

    TDnet の開示は 22:30 過ぎまで出る。当日中に取った一覧は「その時点まで」の
    一覧でしかなく、完全性を主張できない。**翌日 0:00 以降の取得だけを確定**
    とし、それ以前の回は 'provisional'（covered ではない）として次の実行に
    取り直させる。
    """
    now = now or datetime.now()
    return now >= datetime(d.year, d.month, d.day) + timedelta(days=1)


def _doc_id_of(row: dict) -> str:
    return os.path.splitext(row.get("pdf") or row.get("zip") or "")[0]


def postcondition(con, d: date, rows: list[dict], total: int,
                  targets: list[dict], now: datetime | None = None) -> dict:
    """1日ぶんの取得が「揃った」と言えるかの後条件ゲート。

    ok を名乗れるのは次の3つをすべて満たすときだけ:
      1. 一覧ページの総件数（「全N件」）= 実際に読めた行数      → 違えば incomplete
      2. 対象（短信等）の書類がすべて filings にあり、ファイルも実在 → 欠ければ partial
      3. 対象日が終わってから取得した                          → でなければ provisional
    **件数が合っていることを確かめずに ok と書かない。** 2026-09-13 に、朝の
    実行が当日分の一部（9/11: 16件 / 実際333件）で ok と記録され、backfill が
    欠損を見逃し続けていたのを発見した。
    """
    n_rows = len(rows)
    present = 0
    missing_ids = []
    for t in targets:
        doc_id = _doc_id_of(t)
        r = con.execute("SELECT path FROM filings WHERE source=? AND doc_id=?",
                        (SOURCE, doc_id)).fetchone()
        if r and r["path"] and os.path.exists(C.full_path(r["path"])):
            present += 1
        else:
            missing_ids.append(doc_id)
    note = (f"postcondition: listed_total={total} rows_read={n_rows} "
            f"in_scope={len(targets)} present={present}")
    if total != n_rows:
        status = "incomplete"
    elif missing_ids:
        status = "partial"
        note += " missing=" + ",".join(missing_ids[:10])
    elif not day_is_final(d, now):
        status = "provisional"
        note += " (対象日の終了前に取得。翌日以降の実行で確定する)"
    else:
        status = "ok"
    return {"status": status, "present": present, "note": note}


def archive_day(con, fetcher, d: date, *, all_types: bool = False,
                force: bool = False, now: datetime | None = None) -> dict:
    ymd = d.strftime("%Y%m%d")
    iso = d.isoformat()
    run_id = C.start_run(con, SOURCE, iso)
    outdir = os.path.join(C.RAW_DIR, SOURCE, ymd)

    try:
        rows, total = fetch_day_index(fetcher, d)
    except Exception as e:
        C.finish_run(con, run_id, "failed", error=f"{type(e).__name__}: {e}")
        C.log(f"  {iso}: FAILED to read the listing - {type(e).__name__}: {e}")
        return {"date": iso, "status": "failed", "listed": 0, "target": 0,
                "saved": 0, "failed": 0}

    if not rows:
        if total:
            # 「全N件」と出ているのに1行も読めない —— 空ではなく読み損ね。
            st = "incomplete"
            note = f"postcondition: listed_total={total} rows_read=0"
        elif day_is_final(d, now):
            st, note = "empty", "listing returned 0 rows (weekend/holiday or no disclosures)"
        else:
            # 当日の朝に0件なのは「開示が無い日」とは限らない（9/03 00:31 の実例）。
            st, note = "provisional", "0 rows before the target day ended; re-check later"
        C.finish_run(con, run_id, st, n_listed=total or 0, n_target=0, note=note)
        C.log(f"  {iso}: {st} (0 rows read) - {note}")
        return {"date": iso, "status": st, "listed": total or 0, "target": 0,
                "saved": 0, "failed": 0}

    targets = []
    for row in rows:
        subtype, ftype = classify(row["title"])
        if subtype is None:
            if not all_types:
                continue
            subtype, ftype = "その他", "その他"
        row["subtype"], row["type"] = subtype, ftype
        targets.append(row)

    n_saved = n_failed = 0
    for row in targets:
        try:
            n_saved += save_one(con, fetcher, row, d, outdir, force=force)
        except Exception as e:
            n_failed += 1
            C.log(f"    ! {row['code_raw']} {row['title'][:40]}: "
                  f"{type(e).__name__}: {e}")

    gate = postcondition(con, d, rows, total, targets, now)
    status = gate["status"]
    C.finish_run(con, run_id, status, n_listed=total, n_target=len(targets),
                 n_saved=n_saved, n_failed=n_failed, note=gate["note"])
    C.log(f"  {iso}: {status} listed {total} / rows {len(rows)} / in-scope {len(targets)}"
          f" / present {gate['present']} / saved(new) {n_saved}"
          + (f" / FAILED {n_failed}" if n_failed else ""))
    return {"date": iso, "status": status, "listed": total, "rows": len(rows),
            "target": len(targets), "present": gate["present"],
            "saved": n_saved, "failed": n_failed}


def save_one(con, fetcher, row: dict, d: date, outdir: str, force: bool) -> int:
    """Download one disclosure and upsert its filings row. Returns 1 if a new
    filings row was written, 0 if it was already archived."""
    pdf_name = row.get("pdf")
    zip_name = row.get("zip")
    doc_id = os.path.splitext(pdf_name or zip_name or "")[0]
    if not doc_id:
        raise RuntimeError("row has neither a PDF nor an XBRL link")

    code = C.normalise_code(row["code_raw"])
    existing = con.execute(
        "SELECT id, path FROM filings WHERE source=? AND doc_id=?", (SOURCE, doc_id)
    ).fetchone()
    if existing and not force and existing["path"] and os.path.exists(
            C.full_path(existing["path"])):
        return 0

    pdf_path = xbrl_path = None
    if pdf_name:
        dest = os.path.join(outdir, pdf_name)
        if force or not os.path.exists(dest):
            fetcher.download(DOC_URL.format(name=pdf_name), dest)
        pdf_path = C.store_path(dest)
    if zip_name:
        dest = os.path.join(outdir, zip_name)
        if force or not os.path.exists(dest):
            fetcher.download(DOC_URL.format(name=zip_name), dest)
        xbrl_path = C.store_path(dest)

    con.execute(
        "INSERT INTO companies (code, name, market, source, updated_at) "
        "VALUES (?,?,?,?,?) ON CONFLICT(code) DO UPDATE SET "
        "name=COALESCE(companies.name, excluded.name), updated_at=excluded.updated_at",
        (code, row["name"], None, SOURCE, C.utcnow()),
    )
    con.execute(
        "INSERT INTO filings (code, date, type, source, path, xbrl_ok, doc_id, "
        " title, disclosed_at, subtype, pdf_path, xbrl_path, market_place, "
        " company_name, fetched_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) "
        "ON CONFLICT(source, doc_id) DO UPDATE SET "
        " path=excluded.path, xbrl_ok=excluded.xbrl_ok, pdf_path=excluded.pdf_path,"
        " xbrl_path=excluded.xbrl_path, fetched_at=excluded.fetched_at",
        (code, d.isoformat(), row["type"], SOURCE, xbrl_path or pdf_path,
         1 if xbrl_path else 0, doc_id, row["title"],
         f"{d.isoformat()} {row['time']}", row["subtype"], pdf_path, xbrl_path,
         row["place"], row["name"], C.utcnow()),
    )
    con.commit()
    return 1


# ------------------------------------------------------------------ reports
def coverage_report(con, days: int, end: date | None = None) -> dict:
    end = end or date.today()
    start = end - timedelta(days=days - 1)
    missing = C.missing_days(con, SOURCE, start, end)
    runs = con.execute(
        "SELECT target_date, status, n_listed, n_target, n_saved, n_failed, attempt "
        "FROM fetch_runs WHERE source=? AND target_date BETWEEN ? AND ? "
        "AND id IN (SELECT MAX(id) FROM fetch_runs WHERE source=? GROUP BY target_date) "
        "ORDER BY target_date", (SOURCE, start.isoformat(), end.isoformat(), SOURCE),
    ).fetchall()
    return {"start": start.isoformat(), "end": end.isoformat(),
            "missing_weekdays": missing, "runs": [dict(r) for r in runs]}


def print_coverage(con, days: int) -> int:
    rep = coverage_report(con, days)
    C.log(f"TDnet coverage {rep['start']} .. {rep['end']}")
    for r in rep["runs"]:
        C.log(f"  {r['target_date']}  {r['status']:<11} listed={r['n_listed']:<5}"
              f" in-scope={r['n_target']:<4} saved={r['n_saved']:<4}"
              f" failed={r['n_failed']} attempt={r['attempt']}")
    # 当日分は定義上まだ確定できない（day_is_final）。当日を欠損として赤にすると
    # 毎晩の実行が必ず exit 1 になり、本物の欠損と見分けがつかなくなる。
    today = date.today().isoformat()
    if today in rep["missing_weekdays"]:
        rep["missing_weekdays"].remove(today)
        C.log(f"  {today}: 当日分は翌日以降の実行で確定（provisional は欠損に数えない）")
    if rep["missing_weekdays"]:
        C.log(f"  MISSING weekdays ({len(rep['missing_weekdays'])}): "
              + ", ".join(rep["missing_weekdays"]))
    else:
        C.log("  no missing weekdays in the window")
    return 1 if rep["missing_weekdays"] else 0


# --------------------------------------------------------------------- main
def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="TDnet daily archiver (仕様書 §2-1)")
    g = p.add_mutually_exclusive_group()
    g.add_argument("--today", action="store_true", help="archive today")
    g.add_argument("--date", help="archive one day (YYYYMMDD)")
    g.add_argument("--days", type=int, help="archive the last N weekdays")
    g.add_argument("--backfill", type=int, metavar="N",
                   help="archive every weekday in the last N days that is not "
                        "already recorded as ok/empty/partial")
    g.add_argument("--report", type=int, metavar="N",
                   help="print coverage for the last N days and exit")
    p.add_argument("--all-types", action="store_true",
                   help="archive every disclosure, not just 短信/予想修正")
    p.add_argument("--force", action="store_true", help="re-download existing files")
    p.add_argument("--min-interval", type=float, default=0.7,
                   help="seconds between HTTP requests (default 0.7)")
    a = p.parse_args(argv)

    con = C.init_db()
    if a.report:
        return print_coverage(con, a.report)

    if a.date:
        days = [C.parse_date_arg(a.date)]
    elif a.days:
        days = C.business_days_back(a.days)
    elif a.backfill:
        end = date.today()
        start = end - timedelta(days=a.backfill - 1)
        days = [C.parse_date_arg(s) for s in C.missing_days(con, SOURCE, start, end)]
        if not days:
            C.log("backfill: nothing missing in the window")
            return 0
    else:
        days = [date.today()]

    fetcher = C.Fetcher(min_interval=a.min_interval)
    C.log(f"TDnet archiver: {len(days)} day(s) -> {os.path.join(C.RAW_DIR, SOURCE)}")
    results = [archive_day(con, fetcher, d, all_types=a.all_types, force=a.force)
               for d in days]

    # provisional は「まだ確定していない」で失敗ではない（翌日の backfill が取り直す）。
    bad = [r for r in results if r["status"] in ("failed", "partial", "incomplete")]
    C.log(f"done: {len(results)} day(s), {sum(r['saved'] for r in results)} file set(s) "
          f"saved, {fetcher.n_requests} HTTP requests, {len(bad)} day(s) needing retry")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
