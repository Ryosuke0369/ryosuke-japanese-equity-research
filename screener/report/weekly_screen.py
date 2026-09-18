"""screener/report/weekly_screen.py — 週次スクリーン（人間が読む最終成果物）。

製品定義: **スクリーニングの自動化**。最終判断と投資は人間が行う。
したがってこのレポートは「買え」とは言わない。**どこを見るべきかと、
その根拠がどこにあるか**を示す。

出力仕様（2026-09-02 確定）
--------------------------
- フィルタは「**推定発表日が30日以内**」の一点のみ
- 各行: コード / 銘柄名 / 合成スコア / 発火シグナル / 乖離タイプ /
        根拠数値 / 原文リンク / 推定発表日と confidence
- **出口・サイジング・枠制約はここに出さない。** それらは
  paper / shadow が品質テレメトリとして裏で回すもので、
  スクリーンの出力ではない（ゲートでもない）

乖離タイプ
----------
「何と何がズレているか」を1行で言えるようにする。人間が最初に知りたいのは
スコアの数値ではなく**ズレの種類**だから。

  数字×ガイダンス   実績の進捗と会社予想がズレている        (S5)
  BS証拠×織り込み   BSに出ている変化が株価に入っていない    (S1/S2/S3/S4)
  文言変化          定性記述が前年から変わった              (S12・方式B加算)
  修正初回転換      予想修正の向きが初めて変わった          (guidance)

    python -m screener.report.weekly_screen
    python -m screener.report.weekly_screen --days 30 --csv out.csv
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sqlite3
import sys
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1
from screener.report import s13_series as S13S
from screener.signals import freshness as FR

WINDOW_DAYS = 30
SCORE_FLOOR = 0.10          # 表示の足切り。採否ではなく「読む価値がある」の線

# 乖離タイプの割り当て。1銘柄が複数該当することはある。
DIVERGENCE = (
    ("数字×ガイダンス", ("S5",)),
    ("BS証拠×織り込み", ("S1", "S2", "S3", "S4")),
    ("文言変化", ("S12",)),
)

EDINET_DOC = "https://disclosure2.edinet-fsa.go.jp/WZEK0040.aspx?{}"
TDNET_PDF = "https://www.release.tdnet.info/inbs/{}"


# 原文リンクの形式。**「文字列がある」は「リンクが生きている」ではない。**
# 原文リンクはこの製品の要（人間が最終判断する以上、一次資料に飛べないと
# スコアは検証できない）なので、出す前に形を検査する。
_RE_EDINET = re.compile(
    r"^https://disclosure2\.edinet-fsa\.go\.jp/WZEK0040\.aspx\?S[0-9A-Z]{7}$")
_RE_TDNET = re.compile(
    r"^https://www\.release\.tdnet\.info/inbs/[0-9A-Za-z_.-]+\.pdf$")


def check_links(rows):
    """全行の原文リンクを形式検証する。(問題のリスト, 内訳) を返す。

    検査するのは**形式だけ**。生存確認は外部への通信になるので
    `--verify-links` を明示したときだけ行う。
    """
    bad, kind = [], {"edinet": 0, "tdnet": 0, "なし": 0, "不正": 0}
    for r in rows:
        u = (r.get("doc_url") or "").strip()
        if not u:
            kind["なし"] += 1
            bad.append("%s %s: 原文リンクが空" % (r["code"], r["name"][:12]))
        elif _RE_EDINET.match(u):
            kind["edinet"] += 1
        elif _RE_TDNET.match(u):
            kind["tdnet"] += 1
        else:
            kind["不正"] += 1
            bad.append("%s %s: 形式ちがい %r" % (r["code"], r["name"][:12], u))
        if u != (r.get("doc_url") or ""):
            bad.append("%s: 前後に空白がある（CSV で壊れる）" % r["code"])
    return bad, kind


def _page_is_about(body, row):
    """そのページが当該銘柄の書類か。**照合は提出者名で行う。**

    証券コードは当てにならない —— EDINET のページは全角（６３３６）や
    5桁（63360）で出すことがあり、訂正報告書では載らないこともある。
    実際 3070 の訂正有報は本文にコードを一切含まない（2026-09-02 に
    素朴なコード照合で誤検知した）。提出者名なら書類の先頭に必ず出る。
    """
    who = (row.get("doc_company") or "").strip()
    cands = [w for w in (who, who.replace("株式会社", ""), row.get("name") or "")
             if len(w) >= 2]
    if any(w in body for w in cands):
        return True
    code = row["code"]
    wide = code.translate(str.maketrans("0123456789", "０１２３４５６７８９"))
    return any(x in body for x in (code, wide, code + "0"))


def verify_links(rows, timeout=20):
    """生存確認。**明示したときだけ通信する。** (問題のリスト) を返す。"""
    import urllib.error
    import urllib.request
    bad = []
    for r in rows:
        u = (r.get("doc_url") or "").strip()
        if not u:
            continue
        try:
            req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                body = resp.read(300000).decode("utf-8", "ignore")
            if resp.status != 200:
                bad.append("%s: HTTP %s" % (r["code"], resp.status))
            elif not _page_is_about(body, r):
                # 200 が返っても中身が別の会社なら、リンクとしては壊れている。
                bad.append("%s %s: 200 だが本文に提出者名が無い（別書類の疑い）"
                           % (r["code"], r["name"][:12]))
        except urllib.error.HTTPError as e:
            bad.append("%s: HTTPError %s" % (r["code"], e.code))
        except Exception as e:
            bad.append("%s: %s: %s" % (r["code"], type(e).__name__, e))
    return bad


# ---------------------------------------------------------------- 信頼性・注意
# 取引所の指定は人が管理する（security_flags）。自動検知は表示のみ。
EXCLUDE_UNRELIABLE = True          # 既定 ON。--include-unreliable で解除

# **除外するのは「数字の信頼性そのもの」が否定された種別だけ。**
# 上場維持基準の未適合(listing_maintenance)は上場継続性の話であって、
# 計上済み数字が疑わしいという話ではない。証拠としての数字は生きている
# ので表示フラグに留める。除外の範囲を広げすぎると、本来見るべき候補が
# 静かに消える。
EXCLUDING_FLAG_TYPES = ("special_alert", "supervision")

# TDnet の非決算適時開示から拾うテールリスク（タスク6）。
# **スコアには乗らない。** 発表またぎで人間が知っておくべき情報。
RISK_TITLE = re.compile(
    r"事故|死亡|死傷|人身|捜索|書類送検|逮捕|行政処分|業務停止|営業停止"
    r"|特別注意|不適切|不正|課徴金|改善報告書|上場契約違反")
RISK_DAYS = 180


def security_flagged(mcon):
    """人が入れた信頼性フラグ。code -> [(種類, 指定日, note)]。"""
    out = {}
    try:
        rows = mcon.execute(
            "SELECT code, flag_type, since_date, until_date, note "
            "FROM security_flags WHERE until_date IS NULL").fetchall()
    except sqlite3.OperationalError:
        return out                      # 表が無い＝まだ移行していない
    for r in rows:
        out.setdefault(r["code"], []).append(
            (r["flag_type"], r["since_date"] or "指定日未確認", r["note"] or ""))
    return out


def excludes(flags_for_code):
    """その銘柄を主出力から外すか。**種別で決める。**"""
    return any(t in EXCLUDING_FLAG_TYPES for t, _d, _n in flags_for_code or [])


def disclosure_flag(mcon, code, as_of):
    """as_of 時点で最新の自動検知フラグ。無ければ空 dict。"""
    try:
        r = mcon.execute(
            "SELECT going_concern, gc_matched, accounting_change, ac_from_notes,"
            " ac_from_narrative, ac_matched, doc_date, doc_id FROM disclosure_flags "
            "WHERE code=? AND doc_date<=? ORDER BY doc_date DESC LIMIT 1",
            (code, as_of)).fetchone()
    except sqlite3.OperationalError:
        return {}
    return dict(r) if r else {}


def s13_of(mcon, code, as_of):
    """S13 受注残（シャドウ・0点）。無ければ空 dict。"""
    try:
        r = mcon.execute(
            "SELECT backlog_yoy_pct, orders_yoy_pct, closing_backlog, doc_id,"
            " doc_date, period_note FROM s13_orders WHERE code=? AND available=1 "
            "AND doc_date<=? ORDER BY doc_date DESC LIMIT 1",
            (code, as_of)).fetchone()
    except sqlite3.OperationalError:
        return {}
    return dict(r) if r else {}


def risk_disclosures(mcon, code, as_of, days=RISK_DAYS):
    """非決算の適時開示からテールリスクを拾う。**スコアには乗らない。**

    2026-04-07 のベステラ扇島の死亡事故（翌営業日 -20%）のような、
    数字には出ないが発表またぎで効く情報。誤検出の許容度は実測してから
    相談する前提で、まずは拾って見せる。
    """
    lo = (date.fromisoformat(as_of) - timedelta(days=days)).isoformat()
    try:
        rows = mcon.execute(
            "SELECT date, title, subtype FROM filings WHERE code=? "
            "AND date BETWEEN ? AND ? ORDER BY date DESC LIMIT 40",
            (code, lo, as_of)).fetchall()
    except sqlite3.OperationalError:
        return []
    out = []
    for r in rows:
        t = r["title"] or ""
        if "決算短信" in t or "業績予想" in t:
            continue                    # 決算は別経路で見ている
        if RISK_TITLE.search(t):
            out.append("%s %s" % (r["date"], t[:50]))
    return out[:3]


def doc_id_from_url(url):
    """原文リンク → filings.doc_id。

    EDINET は `...WZEK0040.aspx?S100XXXX` のクエリ部、TDnet は
    `.../inbs/140120260813519912.pdf` のファイル名から拡張子を除いたもの
    （tdnet_archiver が doc_id をそう作る）。2026-09-13 まで EDINET 形式しか
    読んでおらず、全銘柄スキャンで TDnet 根拠の 1,089 件が「doc_id が本体に無い」
    と誤って不一致になっていた（42社の決算窓スキャンは全件 EDINET だったので出なかった）。
    """
    u = (url or "").strip()
    if "?" in u:
        return u.rsplit("?", 1)[-1]
    return os.path.splitext(os.path.basename(u))[0]


def verify_doc_periods(rows, mcon):
    """**リンク先書類の期 == スコア根拠期** を、投影層を経由せず検証する。

    投影層の filings を引いてリンクを作っているので、同じ表で照合しても
    「作った通りに作れている」しか言えない。本体の事実表 `financials_cum`
    に (period, q_no) の行が実在するかを見る —— 期の対応を独立に確かめる
    唯一の経路。戻り値: (問題のリスト, 一致数, 根拠期を持たない行数)。
    """
    bad, ok, nolink = [], 0, 0
    for r in rows:
        if not r.get("evidence_docs"):
            nolink += 1
            continue
        for part in r["evidence_docs"].split(" ｜ "):
            sig, rest = part.split(":", 1)
            per, url = rest.split(" ", 1)[0], rest.split(" ", 1)[1].split("（")[0].strip()
            doc = doc_id_from_url(url)
            fy, q = per.split("-Q")
            f = mcon.execute("SELECT id, date, title FROM filings "
                             "WHERE code=? AND doc_id=?", (r["code"], doc)).fetchone()
            if not f:
                bad.append("%s %s: doc_id %s が本体に無い" % (r["code"], sig, doc))
                continue
            hit = mcon.execute(
                "SELECT COUNT(*) FROM financials_cum WHERE filing_id=? "
                "AND period=? AND q_no=?", (f["id"], fy, int(q))).fetchone()[0]
            if hit:
                ok += 1
            else:
                bad.append("%s %s: 根拠期 %s だがリンク先 %s は %s（%s）"
                           % (r["code"], sig, per, doc, f["date"],
                              (f["title"] or "")[:30]))
    return bad, ok, nolink


def s12_of(mcon, code, as_of):
    """S12（定性文言diff）を取る。**方式B: 合成後に加算する。**

    平均に入れる（方式A）のではなく後から足すのは、S12 の available 率が
    高いので平均に入れると分母が変わり、既存シグナルの順位付けが
    S12 とは無関係に動いてしまうため。分母は変えない。

    返り値: (score, tiers, evidence[list]) / 無ければ (None, None, [])
    """
    r = mcon.execute(
        "SELECT s.filing_id, s.score, s.tier_a, s.tier_b, s.tier_c, "
        " s.segment_changed, s.new_product_mention, s.period_label "
        "FROM s12_scores s JOIN filings f ON f.id=s.filing_id "
        "WHERE s.code=? AND s.available=1 AND f.date<=? "
        "ORDER BY f.date DESC LIMIT 1", (code, as_of)).fetchone()
    if not r:
        return None, None, []
    ev = mcon.execute(
        "SELECT tier, rule_key, score, matched_text, dedup_applied "
        "FROM s12_evidence WHERE filing_id=? AND score<>0 "
        "ORDER BY ABS(score) DESC LIMIT 6", (r["filing_id"],)).fetchall()
    tiers = {"A": r["tier_a"], "B": r["tier_b"], "C": r["tier_c"],
             "segment_changed": r["segment_changed"],
             "new_product_mention": r["new_product_mention"],
             "period": r["period_label"]}
    return r["score"], tiers, [dict(x) for x in ev]


def _fired(scores):
    """発火したシグナル名（available かつ score > 0）。"""
    return [k for k, v in sorted(scores.items())
            if isinstance(v, dict) and v.get("available") and (v.get("score") or 0) > 0]


def _doc_url(source, doc_id, pdf_path):
    """書類 → URL。EDINET は doc_id、TDnet は PDF ファイル名。"""
    if source == "edinet" and doc_id:
        return EDINET_DOC.format(doc_id)
    if pdf_path:
        return TDNET_PDF.format(os.path.basename(pdf_path))
    return None


def _evidence_docs(scores):
    """シグナルごとの根拠書類。**最も新しい根拠期のものを代表にする。**

    S12 が filing_id を直に持つのと同じ形（シグナル→根拠書類）を
    S1/S2/S4/S5 にも与える。S4/S5 は複数期を積むので、リンクは最新期に
    向け、どこまでを含むかは period_note に書く —— **1本のリンクで
    複数期を表せない以上、含まれる期を明示しないと嘘になる。**

    戻り値: (代表の {url, period, date, note}, シグナル別リスト)
    """
    per = []
    for k, v in sorted(scores.items()):
        if not (isinstance(v, dict) and v.get("available") and v.get("period")):
            continue
        u = _doc_url(v.get("doc_source"), v.get("doc_id"), v.get("doc_path"))
        if not u:
            continue
        per.append({"signal": k, "period": v["period"], "url": u,
                    "date": v.get("doc_date"), "note": v.get("period_note") or ""})
    if not per:
        return None, []
    top = max(per, key=lambda x: (x["period"], x["date"] or ""))
    return top, per


def _granularity(scores):
    """各シグナルがどの粒度で比較したか。「S1:2Q同士(span=2)」の形。

    規則B は同じ銘柄の時系列に粒度が混ざることを許す（設計書 §9-4）。
    **混ざってよい代わりに、何と何を比べたかは必ず見えていなければならない。**
    """
    out = []
    for k, v in sorted(scores.items()):
        if not (isinstance(v, dict) and v.get("available")):
            continue
        sp, md = v.get("span_q"), v.get("mode")
        if not sp:
            # **空欄にしない。** 空欄は「該当なし」に見えるが、実際は
            # 「別枠の位置ベース比較で、何と比べたか分からない」。
            # span 検証済みかどうかは人間が真っ先に知りたいことなので、
            # 分からないことを分からないと書く。
            out.append("%s:粒度不明(別枠の位置ベース比較)" % k)
            continue
        label = "四半期" if md == "quarter" else ("半期" if sp == 2 else "%d期" % sp)
        peer = v.get("peer_period")
        out.append("%s:%s(span=%d%s)"
                   % (k, label, sp, "→前年 " + peer if peer else ""))
    return " ".join(out)


def _divergence(fired):
    out = [label for label, keys in DIVERGENCE if any(k in fired for k in keys)]
    return out


def _evidence_lines(scores):
    """根拠数値。スコアラーが返す evidence 文をそのまま使う。

    **根拠文の無い点数は出さない。** 数字だけ見せて出所を示さないのは、
    このシステムが一貫して避けてきたことなので、ここでも守る。
    """
    out = []
    for k, v in sorted(scores.items()):
        if isinstance(v, dict) and v.get("available") and (v.get("score") or 0) > 0:
            ev = (v.get("evidence") or "").strip()
            if ev:
                out.append("%s: %s" % (k, ev))
    return out


def _latest_doc(mcon, code):
    """原文リンク。最新の開示に飛ばす（EDINET優先、無ければTDnet）。

    提出者名も返す。**リンクの生存確認で「別書類に飛んでいないか」を
    照合するのに要る** —— 証券コードはページ上で全角や5桁表記になったり
    訂正報告書では載らなかったりして、識別子として当てにならない。
    """
    r = mcon.execute(
        "SELECT source, doc_id, pdf_path, date, title, company_name FROM filings "
        "WHERE code=? ORDER BY date DESC LIMIT 1", (code,)).fetchone()
    if not r:
        return None, None, None
    who = r["company_name"]
    if r["source"] == "edinet" and r["doc_id"]:
        return EDINET_DOC.format(r["doc_id"]), r["date"], who
    if r["pdf_path"]:
        return TDNET_PDF.format(os.path.basename(r["pdf_path"])), r["date"], who
    return None, r["date"], who


def _revision_turn(mcon, code, as_of):
    """修正初回転換: 予想修正の向きが直前と変わった最初の1回か。"""
    rows = mcon.execute(
        "SELECT date, revision_direction FROM guidance WHERE code=? AND date<=? "
        "AND revision_direction IS NOT NULL ORDER BY date DESC LIMIT 2",
        (code, as_of)).fetchall()
    if len(rows) < 2:
        return None
    now, prev = rows[0]["revision_direction"], rows[1]["revision_direction"]
    if now and prev and now != prev:
        return "%s→%s (%s)" % (prev, now, rows[0]["date"])
    return None


def _no_score_reason(scores, s, floor):
    """スコアが付かなかった／閾値未満の理由を1語で。"""
    excluded = sorted({f.split("(")[0] for v in scores.values() if isinstance(v, dict)
                       for f in (v.get("strict_flags") or [])})
    if s is None:
        if excluded:
            return "証拠不採用(%s)" % ",".join(excluded)
        return "データ欠損(全シグナル available=False)"
    if s < floor:
        return "証拠不足(閾値%.2f未満)" % floor
    return ""


def collect(pcon, mcon, as_of, days=WINDOW_DAYS, floor=SCORE_FLOOR,
            policy="evidence_strict", include_all=False, sort="score",
            exclude_unreliable=EXCLUDE_UNRELIABLE, universe=False):
    root = V1._external_root()
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from module_b.run_scorers import SCORERS_ALL
    from screener.signals.span_runner import score_ticker as _score

    def score_ticker(con, code, scorers, as_of=None):
        # span-matched 合流版。四半期報告書が廃止された期は別枠の
        # 位置ベース比較が前年同期を取り違えるので、そこだけ差し替える。
        return _score(con, code, scorers, as_of=as_of, policy=policy)

    hi = (as_of + timedelta(days=days)).isoformat()
    if universe:
        # 全ユニバース（universe_flag=1）を発表日に関係なく採点する。
        # 発表日は分かる銘柄だけ参考として付ける（決算窓は earnings_window.py で絞る）。
        est = {r["ticker"]: r for r in pcon.execute(
            "SELECT ticker, next_earnings_date, quarter_type, confidence_level "
            "FROM earnings_calendar")}
        cal = [{"ticker": r[0],
                "next_earnings_date": (est[r[0]]["next_earnings_date"] if r[0] in est else ""),
                "quarter_type": (est[r[0]]["quarter_type"] if r[0] in est else ""),
                "confidence_level": (est[r[0]]["confidence_level"] if r[0] in est else "")}
               for r in mcon.execute("SELECT code FROM companies WHERE universe_flag=1 ORDER BY code")]
    else:
        cal = pcon.execute(
            "SELECT ticker, next_earnings_date, quarter_type, confidence_level "
            "FROM earnings_calendar WHERE next_earnings_date BETWEEN ? AND ? "
            "ORDER BY next_earnings_date", (as_of.isoformat(), hi)).fetchall()
    names = {r[0]: r[1] for r in
             mcon.execute("SELECT code, name FROM companies")}

    flagged = security_flagged(mcon)
    rows, excluded, n_err = [], [], 0
    for c in cal:
        code = c["ticker"]
        try:
            res = score_ticker(pcon, code, SCORERS_ALL, as_of=as_of)
        except Exception:
            n_err += 1
            continue
        scores = {k: v for k, v in res.items()
                  if k not in ("evidence_score", "_source", "_strict")}
        s = res.get("evidence_score")
        no_score = _no_score_reason(scores, s, floor)
        strict_flags = " | ".join(
            "%s:%s" % (k, ",".join(v.get("strict_flags")))
            for k, v in sorted(scores.items())
            if isinstance(v, dict) and v.get("strict_flags"))
        strict_notes = " | ".join(
            "%s:%s" % (k, ",".join(v.get("strict_notes")))
            for k, v in sorted(scores.items())
            if isinstance(v, dict) and v.get("strict_notes"))
        # include_all は**参考出力用**。製品のフィルタ(floor)はいじらない。
        # 無評価(None)を 0 と書かないのは、「評価して0点」と
        # 「評価できていない」を混ぜないため。
        if not include_all and (s is None or s < floor):
            continue
        if s is None:
            s = None
        fired = _fired(scores)
        # 方式B: 合成後に S12 を加算する。閾値 0.10 は不変で、
        # **負の S12 で閾値を割ったら除外**（veto）。
        s12, s12_tiers, s12_ev = s12_of(mcon, code, as_of.isoformat())
        base = s
        if s12 is not None and s is not None:
            s = s + s12
            if s < floor and not include_all:
                continue                    # 文言変化が足を引いて閾値割れ
            # **0点のときは発火扱いにしない。** 他のシグナルは score>0 を
            # 要求しているのに S12 だけ「評価できた」で乖離タイプに載せると、
            # 起きていない「文言変化」を報告することになる。
            if abs(s12) > 1e-9:
                fired = fired + ["S12"]
        # 理由は S12 加算**後**のスコアで決める（加算前で決めると、S12 で閾値を
        # 越えた銘柄に「証拠不足」と書いてしまう。2026-09-13 に6銘柄で発生）。
        no_score = _no_score_reason(scores, s, floor)
        div = _divergence(fired)
        turn = _revision_turn(mcon, code, as_of.isoformat())
        if turn:
            div.append("修正初回転換")
        # 開示の信頼性・注意フラグ。**すべて表示のみ**（除外は security_flags のみ）。
        df = disclosure_flag(mcon, code, as_of.isoformat())
        s13 = s13_of(mcon, code, as_of.isoformat())
        s13ser = S13S.quarterly_series(mcon, pcon, code, as_of.isoformat())
        _s1 = scores.get("S1") if isinstance(scores.get("S1"), dict) else {}
        s1d = _s1.get("sales_direction") or {}
        s1shrink = _s1.get("shrink_signal") or {}
        _s5 = scores.get("S5") if isinstance(scores.get("S5"), dict) else {}
        s5d = (_s5.get("details") or {}) if _s5.get("period") else {}
        risks = risk_disclosures(mcon, code, as_of.isoformat())
        # 会計処理変更のあった期を使う YoY 系シグナルには比較可能性の注意を出す。
        comp_caution = 1 if (df.get("accounting_change") and any(
            k in fired for k in ("S1", "S2", "S4", "S5"))) else 0
        # 鮮度。**スコアには触らない**（表示のみ。§8 の値の変更にあたるため）。
        FR.annotate(scores, pcon, code, as_of)
        stale_flag, stale_lag, stale_detail = FR.summarize(scores, set(fired))
        # **根拠期の書類に飛ばす。** 取れないときだけ最新開示に落とす。
        top_doc, per_docs = _evidence_docs(scores)
        link, doc_date, doc_who = _latest_doc(mcon, code)
        doc_kind = "最新開示"
        if top_doc:
            link, doc_date, doc_kind = top_doc["url"], top_doc["date"], "根拠期"
            w = mcon.execute(
                "SELECT company_name FROM filings WHERE code=? AND doc_id=?",
                (code, top_doc["url"].rsplit("?", 1)[-1])).fetchone()
            if w and w[0]:
                doc_who = w[0]
        if exclude_unreliable and excludes(flagged.get(code)):
            # **主出力から外す。** 「計上済み数字＝証拠」の前提が取引所に
            # 否定されている会社を、スコアが高いという理由で人に勧めない。
            # 消さずに監査用リストへ回す（消すと外したことが見えなくなる）。
            excluded.append({"code": code, "name": names.get(code, ""),
                             "score": None if s is None else round(s, 3),
                             "est_date": c["next_earnings_date"],
                             "reason": ";".join(
                                 "%s(%s) %s" % (t, d, n[:60])
                                 for t, d, n in flagged[code])})
            continue
        rows.append({
            "code": code, "name": names.get(code, ""),
            "score": None if s is None else round(s, 3),
            "score_base": None if base is None else round(base, 3),
            "s12": None if s12 is None else round(s12, 3),
            "s12_tier_a": None if not s12_tiers else s12_tiers["A"],
            "s12_tier_b": None if not s12_tiers else s12_tiers["B"],
            "s12_tier_c": None if not s12_tiers else s12_tiers["C"],
            "s12_evidence": " ｜ ".join(
                "%s/%s %+.2f: %s" % (e["tier"], e["rule_key"], e["score"],
                                     (e["matched_text"] or "")[:60])
                for e in s12_ev),
            "n_available": len(fired),
            "fired": ",".join(fired),
            "divergence": " / ".join(div) or "-",
            "evidence": " ｜ ".join(_evidence_lines(scores)),
            "span_matched": ",".join(
                sorted(k for k, v in (res.get("_source") or {}).items()
                       if v == "span_matched")),
            "granularity": _granularity(scores),
            "revision_turn": turn or "",
            "doc_url": link or "", "doc_date": doc_date or "",
            "doc_company": doc_who or "",
            "stale_flag": stale_flag, "stale_lag": stale_lag,
            "stale_detail": stale_detail,
            "reliability_flag": ";".join(
                "%s(%s)" % (t, d) for t, d, _n in flagged.get(code, [])),
            "going_concern": df.get("going_concern") or 0,
            "gc_matched": (df.get("gc_matched") or "")[:200],
            "accounting_change": df.get("accounting_change") or 0,
            "ac_source": ("注記" if df.get("ac_from_notes") else "")
                         + ("本文" if df.get("ac_from_narrative") else ""),
            "ac_matched": (df.get("ac_matched") or "")[:200],
            "comparability_caution": comp_caution,
            "s13_order_backlog_yoy": s13.get("backlog_yoy_pct"),
            "s13_orders_yoy": s13.get("orders_yoy_pct"),
            "s13_doc_id": s13.get("doc_id") or "",
            "s13_doc_date": s13.get("doc_date") or "",
            # S13 四半期単独推移（2026-09-18・§34）。**合成スコアには接続しないが
            # 取得できた銘柄では必ず表示する。**0点と非表示は別の話。
            "s13_series": s13ser.get("text", ""),
            "s13_series_note": s13ser.get("note", ""),
            "s13_bb_latest": (s13ser["rows"][-1]["bb"]
                              if s13ser.get("rows") and s13ser["rows"][-1]["bb"] is not None
                              else ""),
            # S1 売上方向ガード（§32）
            "s1_sales_direction": (s1d or {}).get("status", ""),
            "s1_sales_quarter_level": ("" if not s1d else int(bool(s1d.get("quarter_level")))),
            "s1_sales_trend": (s1d or {}).get("trend_text", ""),
            "s1b_shrink": (s1shrink or {}).get("evidence", ""),
            # S5 進捗率（§33）
            "s5_progress_op": (s5d or {}).get("progress_op", ""),
            "s5_elapsed_q": (s5d or {}).get("elapsed_q", ""),
            "s5_pace_excess_pt": (s5d or {}).get("pace_excess_pt", ""),
            "s5_implied_rest_op": (s5d or {}).get("implied_rest_op", ""),
            "s5_guidance_dead": ("" if not s5d else int(bool(s5d.get("guidance_dead")))),
            "risk_flag": " ｜ ".join(risks),
            "doc_kind": doc_kind,
            "doc_period": (top_doc or {}).get("period", ""),
            "doc_note": (top_doc or {}).get("note", ""),
            "evidence_docs": " ｜ ".join(
                "%s:%s %s%s" % (d["signal"], d["period"], d["url"],
                                "（%s）" % d["note"] if d["note"] else "")
                for d in per_docs),
            "est_date": c["next_earnings_date"] or "",
            "confidence": c["confidence_level"] or "",
            # evidence_strict の監査列（prefer_span では空）
            "policy": policy,
            "evidence_lag_q": ",".join(
                "%s:%s" % (k, scores[k].get("evidence_lag_q")) for k in fired
                if isinstance(scores.get(k), dict) and scores[k].get("evidence_lag_q") is not None),
            "evidence_recent": (
                "" if not fired else int(all(
                    (scores[k].get("evidence_lag_q") == 0) for k in fired
                    if isinstance(scores.get(k), dict) and k != "S12"))),
            "false_positive_flags": strict_flags,
            "strict_notes": strict_notes,
            "no_score_reason": no_score,
        })
    if sort == "date":
        rows.sort(key=lambda r: (r["est_date"] or "9999", -(r["score"] or -9), r["code"]))
    else:
        rows.sort(key=lambda r: (-(r["score"] if r["score"] is not None else -9),
                                 r["est_date"] or "9999", r["code"]))
    return rows, len(cal), n_err, excluded


def render(rows, as_of, days, n_cal, n_err, excluded=None):
    C.log("=" * 96)
    C.log("週次スクリーン  %s 時点 / 推定発表日が %d 日以内" % (as_of, days))
    n_hit = sum(1 for r in rows if (r["score"] or -9) >= SCORE_FLOOR)
    C.log("  カレンダー該当 %d 銘柄 → 出力 %d 銘柄 / スコア %.2f 以上 %d 銘柄"
          % (n_cal, len(rows), SCORE_FLOOR, n_hit))
    if n_err:
        C.log("  ! スコア計算で例外 %d 件（データ不足ではない。要調査）" % n_err)
    for e in (excluded or []):
        C.log("  [除外] %s %s  スコア %s  発表 %s  理由 %s"
              % (e["code"], e["name"][:14],
                 "未評価" if e["score"] is None else "%.3f" % e["score"],
                 e["est_date"], e["reason"][:110]))
    C.log("=" * 96)
    if not rows:
        C.log("  該当なし")
        return
    for i, r in enumerate(rows, 1):
        C.log("")
        C.log("[%2d] %s %s   スコア %s (発火 %d本)   発表予定 %s (%s)"
              % (i, r["code"], r["name"][:18],
                 "未評価" if r["score"] is None else "%.3f" % r["score"],
                 r["n_available"], r["est_date"], r["confidence"]))
        C.log("     乖離タイプ : %s" % r["divergence"])
        C.log("     発火シグナル: %s%s"
              % (r["fired"] or "-",
                 ("   [span-matched: %s]" % r["span_matched"])
                 if r["span_matched"] else ""))
        if r["evidence"]:
            C.log("     根拠数値   : %s" % r["evidence"])
        if r["granularity"]:
            C.log("     比較粒度   : %s" % r["granularity"])
        if r["stale_detail"]:
            C.log("     証拠の鮮度 : %s%s"
                  % ("** 古い（最大%d期遅れ） ** " % r["stale_lag"]
                     if r["stale_flag"] else "",
                     r["stale_detail"]))
        if r["s12"] is None:
            C.log("     文言変化   : 未評価（前期ペアが無い）")
        elif abs(r["s12"]) < 1e-9:
            C.log("     文言変化   : なし（評価済み・前年から差分なし）")
        else:
            C.log("     文言変化   : S12 %+.3f (合成前 %s → %s)  "
                  "Tier A %+.2f / B %+.2f / C %+.2f"
                  % (r["s12"],
                     "未評価" if r["score_base"] is None else "%.3f" % r["score_base"],
                     "未評価" if r["score"] is None else "%.3f" % r["score"],
                     r["s12_tier_a"], r["s12_tier_b"], r["s12_tier_c"]))
            if r["s12_evidence"]:
                C.log("       文言根拠 : %s" % r["s12_evidence"])
        notes = []
        if r["reliability_flag"]:
            notes.append("上場・信頼性フラグ: %s" % r["reliability_flag"])
        if r["going_concern"]:
            notes.append("継続企業の前提に注記あり")
        if r["accounting_change"]:
            notes.append("会計処理変更(%s)" % (r["ac_source"] or "?"))
        if r["comparability_caution"]:
            notes.append("**前年比の比較可能性に注意**")
        if notes:
            C.log("     開示の注意 : %s" % " / ".join(notes))
            if r["ac_matched"]:
                C.log("       会計根拠 : %s" % r["ac_matched"][:150])
        # --- S1 売上方向ガード（§32）。**S1 が発火しているときは必ず出す。**
        if r["s1_sales_trend"]:
            C.log("     売上の方向 : %s  [Q単独確認 %s]"
                  % (r["s1_sales_direction"],
                     "済" if r["s1_sales_quarter_level"] == 1 else "**未**（累計 span での判定）"))
            C.log("       売上推移 : %s" % r["s1_sales_trend"])
        if r["s1b_shrink"]:
            C.log("     S1b(0点)   : %s" % r["s1b_shrink"])
        # --- S5 進捗率（§33）
        if r["s5_elapsed_q"] != "" and r["s5_progress_op"] not in ("", None):
            C.log("     進捗率(S5) : OP進捗 %.0f%% / 経過 %d四半期(期待 %.0f%%) / 超過 %s pt"
                  " / 暗黙の残存 OP %s%s"
                  % (float(r["s5_progress_op"]) * 100, int(r["s5_elapsed_q"]),
                     int(r["s5_elapsed_q"]) / 4.0 * 100, r["s5_pace_excess_pt"],
                     r["s5_implied_rest_op"],
                     "  ← **死んだガイダンス**" if r["s5_guidance_dead"] == 1 else ""))
        # --- S13 受注。**取得できた銘柄では必ず出す**（§34）。
        # 0点だから表示しない、では情報の損失になる。
        if r["s13_series"]:
            C.log("     受注(S13・0点): %s" % r["s13_series"])
            if r["s13_series_note"]:
                C.log("       注      : %s" % r["s13_series_note"])
        if r["s13_order_backlog_yoy"] is not None or r["s13_orders_yoy"] is not None:
            C.log("     受注残(参考): 残高YoY %s%% / 受注YoY %s%%  [S13・0点・%s]"
                  % (r["s13_order_backlog_yoy"], r["s13_orders_yoy"],
                     r["s13_doc_date"] or "-"))
        if r["risk_flag"]:
            C.log("     リスク開示 : %s" % r["risk_flag"])
        if r["revision_turn"]:
            C.log("     修正の転換 : %s" % r["revision_turn"])
        if r["doc_url"]:
            C.log("     原文(%s): %s  (%s%s)"
                  % (r["doc_kind"], r["doc_url"], r["doc_date"],
                     " / " + r["doc_period"] if r["doc_period"] else ""))
            if r["doc_note"]:
                C.log("       %s" % r["doc_note"])
            if r["evidence_docs"] and " ｜ " in r["evidence_docs"]:
                C.log("       シグナル別: %s" % r["evidence_docs"])


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--as-of", help="基準日 YYYY-MM-DD（既定は今日）")
    p.add_argument("--days", type=int, default=WINDOW_DAYS)
    p.add_argument("--floor", type=float, default=SCORE_FLOOR)
    p.add_argument("--csv", help="CSV にも書き出す")
    p.add_argument("--all-calendar", action="store_true",
                   help="カレンダー該当を全件出す（参考出力。製品のフィルタは"
                        "変えない。スコア未評価の銘柄も含む）")
    p.add_argument("--sort", default="score", choices=("score", "date"),
                   help="並び順。date は推定発表日の昇順")
    p.add_argument("--include-unreliable", action="store_true",
                   help="security_flags の銘柄も主出力に含める（既定は除外）")
    p.add_argument("--verify-doc-periods", action="store_true",
                   help="リンク先書類の期がスコア根拠期と一致するか検証する"
                        "（本体の financials_cum で独立に照合。通信なし）")
    p.add_argument("--verify-links", action="store_true",
                   help="原文リンクを実際に取得して生存確認する（外部通信）")
    p.add_argument("--policy", default="evidence_strict",
                   choices=("evidence_strict", "prefer_span", "span_only"),
                   help="evidence_strict が既定（2026-09-13 切替、calibration_backlog §31）: 根拠期なし・"
                        "直前四半期から2期以上古い根拠・前年比破壊を点にしない。"
                        "prefer_span は旧既定（根拠期なしの別枠結果も通す）、span_only は旧規則A")
    p.add_argument("--universe", action="store_true",
                   help="発表日カレンダーに関係なく universe_flag=1 の全銘柄を採点する")
    a = p.parse_args(argv)

    as_of = date.fromisoformat(a.as_of) if a.as_of else date.today()
    pdb = os.path.join(C.DATA_DIR, "projection.db")
    pcon = sqlite3.connect("file:%s?mode=ro" % pdb.replace("\\", "/"), uri=True)
    pcon.row_factory = sqlite3.Row          # 別枠は r["col"] で読む
    mcon = C.init_db()

    rows, n_cal, n_err, excluded = collect(
        pcon, mcon, as_of, a.days, a.floor, a.policy,
        include_all=a.all_calendar, sort=a.sort,
        exclude_unreliable=not a.include_unreliable, universe=a.universe)
    render(rows, as_of, a.days, n_cal, n_err, excluded)

    bad, kind = check_links(rows)
    C.log("")
    C.log("原文リンクの形式検証: %s"
          % " / ".join("%s %d" % (k, v) for k, v in kind.items() if v))
    if a.verify_doc_periods:
        vbad, vok, vnol = verify_doc_periods(rows, mcon)
        C.log("  根拠期の照合: 一致 %d / 不一致 %d / 根拠期を持たない行 %d"
              % (vok, len(vbad), vnol))
        bad += vbad
    if a.verify_links:
        bad += verify_links(rows)
        C.log("  生存確認: %d 件を実際に取得した" % len(rows))
    for b in bad:
        C.log("  ! %s" % b)
    if not bad:
        C.log("  問題なし（%d 行すべて）" % len(rows))
    if a.csv and rows:
        with open(a.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
        C.log("")
        C.log("CSV: %s" % a.csv)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
