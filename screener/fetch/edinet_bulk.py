"""screener/fetch/edinet_bulk.py - EDINET一括取得 (仕様書 §2-2).

用途は仕様書のとおり「年次・半期の深掘り用」。四半期報告書制度は廃止済みなので
四半期のBS/PLはTDnet短信が主源泉であり、EDINETは実績時系列の土台・セグメント注記・
自己株・詳細BS科目のために使う。

なぜ日付スイープなのか
    EDINET API v2 の documents.json は **1日1回の呼び出しで全社ぶんの提出書類**を
    返す。銘柄ごとに探索する既存の scripts/edinet_fetcher.py は1銘柄あたり最大400回
    APIを叩く設計で、1,900社×3年には桁違いに向かない。日付スイープなら
    3年 = 約750営業日 = 約750回の索引呼び出しで全社を覆える。
    索引のコストは対象社数と無関係なので、試走(58社)で測った所要時間が
    そのまま本番(全ユニバース)の見積りになる。

2パス構成
    pass 1 (index)     documents.json をスイープし、対象書類を filings に
                       path=NULL / xbrl_ok=0 で登録する。ここまでは軽い。
    pass 2 (download)  filings のうち未取得のものだけZIPを落とす。
                       中断しても pass 1 の結果は残り、再開は差分だけになる。

Usage
    python -m screener.fetch.edinet_bulk --trial
    python -m screener.fetch.edinet_bulk --index --from 2025-09-01 --to 2026-08-28
    python -m screener.fetch.edinet_bulk --download --limit 200
    python -m screener.fetch.edinet_bulk --full --years 3        # 夜間実行
    python -m screener.fetch.edinet_bulk --report
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

SOURCE = "edinet"
API = "https://api.edinet-fsa.go.jp/api/v2"

# 仕様書 §2-2: 有価証券報告書・半期報告書。訂正版も取る(訂正は遡及修正の証拠であり、
# §3-2 が単独値を無効化する判断材料になる)。
# 四半期報告書(140/150)は 2024-04-01 以後に開始する四半期から廃止された。
# 「制度廃止済みだから対象外」として最初から除外していたが、**収集対象期間の
# 大半では現役の制度**であり、取りこぼしだった(2026-08-31 修正)。
# 実地確認: 2022-11-14 に996件、2023-11-14 に937件、2024-02-14 に907件が存在し、
# 2024-05-15(286件)を最後に消滅する。
#
# これが効くのは過去データの遡及。2024年4月以降の四半期粒度は EDINET に存在せず
# 短信でしか取れないが、TDnet の保持は約40日(2026-08-31 実測: 07-23 が最古)で
# 遡及できない。**FY2023〜FY2024 の連続8四半期は四半期報告書からしか作れない。**
DOC_TYPES = {
    "120": "有報",
    "130": "有報",      # 訂正有価証券報告書
    "140": "四半期",    # 2024-04-01 以後開始の四半期から廃止。過去分の遡及に必須
    "150": "四半期",    # 訂正四半期報告書
    "160": "半期",
    "170": "半期",      # 訂正半期報告書
}

# 仕様書 §6 の検証8銘柄。「電子材料」は先生リストの表記で、仕様書は
# 日本電子材料(6855)と推定・要確認としていた。JPX上場一覧4,444社のうち
# 社名に「電子材料」を含むのは6855の1社のみ(2026-08-29確認)なので確定とみなす。
VALIDATION_CODES = ["285A", "2962", "3110", "278A", "3905", "6217", "4192", "6855"]


def _headers() -> dict:
    return {"Ocp-Apim-Subscription-Key": C.require_env("EDINET_API_KEY")}


def weekdays(start: date, end: date):
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


# ------------------------------------------------------------------ pass 1
def index_day(con, fetcher, d: date, codes: set[str] | None) -> dict:
    """documents.json for one date → filings rows (path=NULL)."""
    run_id = C.start_run(con, SOURCE, d.isoformat())
    try:
        r = fetcher.get(f"{API}/documents.json",
                        params={"date": d.isoformat(), "type": "2"},
                        allow_status=(200, 400, 404))
        if r.status_code != 200:
            C.finish_run(con, run_id, "failed", error=f"HTTP {r.status_code}")
            return {"status": "failed", "listed": 0, "target": 0}
        payload = r.json()
        return _index_docs(con, run_id, d, payload, codes)
    except Exception as e:
        # 1日ぶんの失敗でスイープ全体を落とさない。失敗日は fetch_runs に
        # 残り、index_range が "failed day(s)" として数えるので、後から
        # その日だけ拾い直せる —— 434日の走査が一過性のエラー1件で
        # 消えるほうが高くつく（2026-09-01 13:33、2025-05-08 で実際に消えた）。
        try:
            con.rollback()
        except Exception:                                # pragma: no cover
            pass
        C.finish_run(con, run_id, "failed",
                     error=f"{type(e).__name__}: {e}"[:400])
        return {"status": "failed", "listed": 0, "target": 0}


def _index_docs(con, run_id: int, d: date, payload: dict,
                codes: set[str] | None) -> dict:
    """取得済みの documents.json を filings に落とす。index_day の後半。"""
    docs = payload.get("results") or []
    n_indexed = n_target = 0
    for doc in docs:
        if doc.get("docTypeCode") not in DOC_TYPES:
            continue
        if str(doc.get("xbrlFlag")) != "1":
            continue
        code = C.normalise_code(doc.get("secCode"))
        if not code:
            continue                                     # 非上場の提出者
        # 索引はユニバースに依存させない。ここで codes で絞ると、あとで
        # ユニバース条件を広げたときに「走査済みの日」に載っている新規銘柄の
        # 書類が永久に入らず、日付単位の全再走査でしか回復できなくなる
        # （2026-08-31 の上限拡大 600億→1,000億 で +249社 が出た）。
        # 索引は全上場銘柄の有報/半期を持ち、絞るのは download_pending 側。
        con.execute(
            "INSERT INTO filings (code, date, type, source, path, xbrl_ok, doc_id, "
            " title, disclosed_at, subtype, company_name, fetched_at) "
            "VALUES (?,?,?,?,NULL,0,?,?,?,?,?,NULL) "
            "ON CONFLICT(source, doc_id) DO NOTHING",
            (code, d.isoformat(), DOC_TYPES[doc["docTypeCode"]], SOURCE,
             doc.get("docID"), doc.get("docDescription"), d.isoformat(),
             doc.get("docTypeCode"), doc.get("filerName")))
        n_indexed += 1
        if codes is None or code in codes:
            n_target += 1                                # 現ユニバースでの内数
    con.commit()
    C.finish_run(con, run_id, "ok" if docs else "empty",
                 n_listed=len(docs), n_target=n_target,
                 note=f"index pass (no download); indexed={n_indexed}")
    return {"status": "ok", "listed": len(docs),
            "indexed": n_indexed, "target": n_target}


def index_range(con, fetcher, start: date, end: date,
                codes: set[str] | None) -> dict:
    days = list(weekdays(start, end))
    C.log(f"EDINET index sweep {start} .. {end} ({len(days)} weekdays), "
          f"target codes: {'all listed' if codes is None else len(codes)}")
    tot = {"days": 0, "listed": 0, "indexed": 0, "target": 0, "failed": 0}
    for i, d in enumerate(days, 1):
        res = index_day(con, fetcher, d, codes)
        tot["days"] += 1
        tot["listed"] += res["listed"]
        tot["indexed"] += res.get("indexed", 0)
        tot["target"] += res["target"]
        tot["failed"] += 1 if res["status"] == "failed" else 0
        if i % 25 == 0 or i == len(days):
            C.log(f"  [{i}/{len(days)}] {d}  cumulative: {tot['listed']} docs seen, "
                  f"{tot['indexed']} indexed, {tot['target']} in current universe, "
                  f"{tot['failed']} failed day(s)")
    return tot


# ------------------------------------------------------------------ pass 2
def download_pending(con, fetcher, limit: int | None = None,
                     codes: set[str] | None = None,
                     subtypes: set[str] | None = None) -> dict:
    """未取得の書類を落とす。取得対象の絞り込みは**ここ**で行う。

    索引は全上場銘柄を持っているので、codes を渡さないと対象外の会社まで
    落としにいく。逆に、ユニバースを広げたときは codes が広がるだけで、
    既に xbrl_ok=1 の行は WHERE から外れるため再取得は起きない ——
    「既取得分は無効にせず差分のみ追加取得」がこの1か所で成立する。

    subtypes は書類種別(docTypeCode)の絞り込み。四半期報告書(140/150)だけを
    先に埋める、のように「何を今欲しいか」で取得順を決めるために要る。
    絞っても既取得判定は変わらないので、後から広げれば残りが差分で入る。
    """
    rows = con.execute(
        "SELECT id, code, date, doc_id, type, subtype FROM filings "
        "WHERE source='edinet' AND (xbrl_ok=0 OR path IS NULL) "
        "ORDER BY date DESC").fetchall()
    n_all = len(rows)
    if codes is not None:
        rows = [r for r in rows if r["code"] in codes]
    if subtypes is not None:
        rows = [r for r in rows if r["subtype"] in subtypes]
    if limit:
        rows = rows[:int(limit)]
    filt = "" if subtypes is None else f" / 種別 {','.join(sorted(subtypes))} に限定"
    C.log(f"EDINET download: {len(rows)} pending document(s) "
          f"(索引済みの未取得 {n_all} 件のうち、取得対象は {len(rows)} 件{filt})")
    ok = failed = 0
    for i, r in enumerate(rows, 1):
        dest = os.path.join(C.RAW_DIR, SOURCE, r["date"], f"{r['doc_id']}.zip")
        try:
            if not os.path.exists(dest):
                # type=1 = XBRL一式のZIP。type を省略すると 400 になる。
                fetcher.download(f"{API}/documents/{r['doc_id']}?type=1", dest)
            con.execute("UPDATE filings SET path=?, xbrl_path=?, xbrl_ok=1, "
                        "fetched_at=? WHERE id=?",
                        (C.store_path(dest),
                         C.store_path(dest), C.utcnow(), r["id"]))
            ok += 1
        except Exception as e:
            failed += 1
            C.log(f"  ! {r['code']} {r['doc_id']}: {type(e).__name__}: {e}")
        if i % 20 == 0:
            con.commit()
            C.log(f"  [{i}/{len(rows)}] ok={ok} failed={failed}")
    con.commit()
    return {"pending": len(rows), "ok": ok, "failed": failed}


# ------------------------------------------------------------------ report
def report(con) -> None:
    tot = con.execute("SELECT COUNT(*) c FROM filings WHERE source='edinet'"
                      ).fetchone()["c"]
    got = con.execute("SELECT COUNT(*) c FROM filings WHERE source='edinet' "
                      "AND xbrl_ok=1").fetchone()["c"]
    codes = con.execute("SELECT COUNT(DISTINCT code) c FROM filings "
                        "WHERE source='edinet'").fetchone()["c"]
    C.log(f"EDINET filings indexed: {tot} / downloaded: {got} / "
          f"distinct codes: {codes}")
    for r in con.execute("SELECT type, COUNT(*) c FROM filings WHERE source='edinet' "
                         "GROUP BY type ORDER BY c DESC"):
        C.log(f"  {r['type']:<8} {r['c']:>6}")
    C.log("validation-8 coverage (仕様書 §6):")
    for code in VALIDATION_CODES:
        rows = con.execute(
            "SELECT COUNT(*) c, SUM(xbrl_ok) d, MIN(date) f, MAX(date) l "
            "FROM filings WHERE source='edinet' AND code=?", (code,)).fetchone()
        name = con.execute("SELECT name FROM companies WHERE code=?",
                           (code,)).fetchone()
        C.log(f"  {code:<5} {str(name['name'] if name else '?'):<16} "
              f"indexed={rows['c'] or 0:<3} downloaded={rows['d'] or 0:<3} "
              f"{rows['f'] or '-'} .. {rows['l'] or '-'}")


# ユニバース候補 = 「条件を満たす」+「まだ判定していない」。除外済みは含めない。
# 内国株かどうかは市場区分名では判定しない —— J-Quants V2 で MarketCodeName
# (「プライム（内国株式）」)が MktNm(「プライム」)へ変わり、'内国株式' を
# 部分一致で探す条件は 1,093 社を 8 社まで取りこぼしていた(2026-08-31)。
# 商品種別による除外は apply_universe_rules が exclude_reason に落とし済みなので、
# ここで重ねて商品種別を見る必要はない。
_CANDIDATE_SQL = ("SELECT code FROM companies "
                  "WHERE (exclude_reason IS NULL OR exclude_reason LIKE '%未取得%')")


def universe_codes(con) -> set[str]:
    """取得対象 = ユニバース候補 ∪ 検証8銘柄 (仕様書 §2-3 / §6)。

    検証8銘柄はユニバースの部分集合ではない。5銘柄は時価総額上限を超えていて
    除外されるので、和集合を取らないと §6 の検証データが欠ける。
    """
    codes = {r["code"] for r in con.execute(_CANDIDATE_SQL)}
    codes.update(VALIDATION_CODES)
    return codes


def trial_codes(con, extra: int = 50) -> set[str]:
    """検証8銘柄 + ユニバース候補から先頭 `extra` 社。

    50社は「除外条件に当たらない銘柄」からコード順に取る。恣意的に選ぶと
    試走が本番の見積りにならないので、順序は決定的にする。
    """
    codes = set(VALIDATION_CODES)
    rows = con.execute(_CANDIDATE_SQL + " ORDER BY code LIMIT ?", (extra,)).fetchall()
    codes.update(r["code"] for r in rows)
    return codes


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="EDINET bulk loader (仕様書 §2-2)")
    p.add_argument("--trial", action="store_true",
                   help="検証8銘柄+50社・直近12ヶ月で試走")
    p.add_argument("--full", action="store_true", help="全ユニバース(夜間実行)")
    p.add_argument("--index", action="store_true", help="索引パスのみ")
    p.add_argument("--download", action="store_true", help="取得パスのみ")
    p.add_argument("--report", action="store_true")
    p.add_argument("--from", dest="dfrom", help="YYYY-MM-DD")
    p.add_argument("--to", dest="dto", help="YYYY-MM-DD")
    p.add_argument("--years", type=float, default=1.0)
    p.add_argument("--trial-extra", type=int, default=50)
    p.add_argument("--limit", type=int, help="download pass の上限")
    p.add_argument("--subtypes", help="download pass を書類種別(docTypeCode)で絞る。"
                                      "カンマ区切り。例: 140,150 = 四半期報告書のみ")
    p.add_argument("--all-codes", action="store_true",
                   help="索引済みの全銘柄を取得対象にする（既定は現ユニバース）。"
                        "索引は全上場銘柄を持つので、指定すると数万件になる")
    p.add_argument("--recent", type=int, metavar="N",
                   help="日次用: 直近N日を再索引（当日中に増えた提出も拾う）+ 直近30日の"
                        "欠損日を補完 → 現ユニバースの未取得を取得。2026-09-13 に"
                        "8/31 で停止していたのを受けて追加")
    p.add_argument("--min-interval", type=float, default=1.2,
                   help="EDINETへの最短リクエスト間隔(秒)。既定1.2は "
                        "仕様書 §2-2『レート制限に注意して間隔を空ける』の実装")
    a = p.parse_args(argv)

    con = C.init_db()
    if a.report:
        report(con)
        return 0

    end = C.parse_date_arg(a.dto) if a.dto else date.today()
    start = (C.parse_date_arg(a.dfrom) if a.dfrom
             else end - timedelta(days=int(365 * a.years)))

    if a.trial:
        codes = trial_codes(con, a.trial_extra)
        C.log(f"trial: {len(codes)} code(s) = 検証8銘柄 + {a.trial_extra}社")
    elif a.all_codes:
        codes = None
        C.log("all-codes: 索引済みの全銘柄を取得対象にする")
    else:
        # --full でも素の --index/--download でも既定は現ユニバース。
        # 既定を None(全銘柄) にすると、索引が全上場銘柄を持つように
        # なった以上、--download 単独実行が数万件を落としにいく。
        codes = universe_codes(con)
        C.log(f"target: {len(codes)} code(s) = ユニバース候補 ∪ 検証8銘柄")

    fetcher = C.Fetcher(min_interval=a.min_interval, headers=_headers())
    if a.recent:
        end = date.today()
        start = end - timedelta(days=int(a.recent))
        # 再索引窓より前の30日で、成功記録の無い平日を個別に拾う（PC停止明けの自己回復）
        gaps = C.missing_days(con, SOURCE, end - timedelta(days=30), start - timedelta(days=1))
        C.log(f"EDINET recent: re-index {start}..{end} + gap days {len(gaps)}")
        for s in gaps:
            index_day(con, fetcher, C.parse_date_arg(s), codes)
        tot = index_range(con, fetcher, start, end, codes)
        dl = download_pending(con, fetcher, a.limit, codes, None)
        report(con)
        C.log(f"HTTP requests this run: {fetcher.n_requests}")
        # 成功条件: 索引の失敗日 0 かつ取得失敗 0
        return 1 if (tot["failed"] or dl["failed"]) else 0
    if a.trial or a.full or a.index:
        index_range(con, fetcher, start, end, codes)
    if a.trial or a.full or a.download:
        # PowerShell は引数モードでも `140,150` を配列と解釈し、[string] に
        # 詰め直すと "140 150" になる。区切りはカンマでも空白でも受ける。
        subtypes = ({t for t in a.subtypes.replace(",", " ").split() if t}
                    if a.subtypes else None)
        download_pending(con, fetcher, a.limit, codes, subtypes)
    report(con)
    C.log(f"HTTP requests this run: {fetcher.n_requests}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
