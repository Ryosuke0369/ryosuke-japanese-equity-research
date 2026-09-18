"""screener/signals/span_scorers.py — span-matched 版のスコアラー（S5/S1/S2/S4）。

設計の正本: `docs/span_matched_design.md` §6。

なぜ本体側に作るのか
--------------------
別枠 `module_b/scorers_*.py` は**変更禁止**。あちらは `pl[-5]`（位置ベース）で
前年同期を取るので、四半期報告書の廃止で 1Q/3Q が消えた銘柄では
「5つ前の行」が前年同期でなくなる。本体側に**明示的な前年ペア探索**
(`span_matched.find_yoy_peer`) を使う版を置き、モードが span_matched に
切り替わった期だけこちらを使う（quarter モードの期は別枠のまま。
既存の挙動を1ミリも変えないため）。

**閾値・係数は別枠の実装をそのまま写している。** span の扱いを変えただけで、
感度の再調整はしていない（するなら新規事前登録）。

PIT
---
別枠 `visible_generations` と同じ規則。**filing_date <= as_of の書類で
語られた期だけを見る。** 投影 DB の generation は常に1（本体は
filings.date で訂正を解決済み）なので、実質は「その期を語る書類が
as_of までに出ていたか」の集合フィルタになる。

返り値
------
別枠 `_result` と同じ形に `mode` / `span_q` / `peer_period` を足したもの。
どちらのモードで算出したかを持たないと、後から成績を層別できない。
"""
from __future__ import annotations

import os
import sys

try:
    from screener import common as C            # noqa: F401
except ImportError:                             # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C            # noqa: F401

from screener.projection import span_matched as SM
from screener.signals import sales_direction as SD

DAYS_PER_Q = 91          # 別枠 data_access と同じ
MIN_BASE = 30            # 契約負債の少額無効ライン（百万円）。別枠と同じ

SPAN_SIGNALS = ("S1", "S2", "S4", "S5")

# 比較ペアが「直近の開示」からどれだけ遡ってよいか（四半期数）。
# 別枠は位置で 5 行前を前年同期とみなすので、系列に穴があると
# **3年前や4年前と比べて平然とスコアを返す**（7231 は FY2026-Q4 を
# FY2022-Q4 と比較していた）。こちらは前年ペアを明示的に探すぶん
# 誤った相手とは組まないが、代わりに「前年ペアが取れる最後の期」まで
# 遡って黙って古い期を評価してしまう。**古い期の証拠を今の証拠として
# 出すのは、間違った相手と比べるのと同じくらい悪い。** 直近開示から
# 4四半期を超えて古い組は available=False にする。
MAX_STALE_Q = 4


def _result(score, available, evidence, details=None, **extra):
    """別枠 `_result` と同じ形＋ `mode` / `span_q` / `period` / `peer_period`。

    **available なときは evidence の先頭に「何と何を比べたか」を必ず出す。**
    規則B は同じ銘柄の時系列に粒度が混ざることを許す（設計書 §9-4）ので、
    混ざってよい代わりに比較の根拠が人間から見えていなければならない。
    ここ1か所で付けるので、個々のスコアラーが書き忘れることがない。
    """
    r = {"score": round(float(score), 3), "available": bool(available),
         "evidence": evidence, "details": details or {}}
    r.setdefault("mode", None)
    r.setdefault("span_q", None)
    r.setdefault("period", None)
    r.setdefault("peer_period", None)
    r.setdefault("peer_synthesized", 0)
    r.setdefault("peer_synth_parts", None)
    r.setdefault("doc_id", None)
    r.setdefault("doc_date", None)
    r.setdefault("doc_source", None)
    r.setdefault("doc_path", None)
    r.setdefault("period_note", None)
    r.update(extra)
    if r["available"] and r["period"]:
        tag = r["period"]
        if r["peer_period"]:
            tag += " vs " + r["peer_period"]
            if r["peer_synthesized"]:
                # **合成であることを黙らない。** 開示された値そのものでは
                # ないので、証拠にそう書けないなら出してはいけない。
                tag += "(合成:%s)" % (r["peer_synth_parts"] or "?")
        tag += " span=%s/%s" % (r["span_q"], r["mode"])
        r["evidence"] = "[%s] %s" % (tag, r["evidence"])
    return r


def clamp(x, lo=-1.0, hi=1.0):
    return max(lo, min(hi, x))


def _iso(as_of):
    return as_of.isoformat() if hasattr(as_of, "isoformat") else as_of


def visible_periods(con, ticker, as_of=None):
    """as_of 時点で公知だった period_end の集合。None なら全期。

    別枠 `visible_generations` と同じ基準（filings.filing_date）。
    """
    if as_of is None:
        return None
    return {r[0] for r in con.execute(
        "SELECT DISTINCT period_end FROM filings WHERE ticker=? AND filing_date<=?",
        (ticker, _iso(as_of)))}


def _adjustments(con, ticker, vis):
    """一時収入の控除額。period_end → 合計。別枠 get_adjustments と同義。"""
    out = {}
    for pe, amt in con.execute(
            "SELECT period_end, amount FROM pl_adjustments WHERE ticker=?", (ticker,)):
        if vis is None or pe in vis:
            out[pe] = out.get(pe, 0) + (amt or 0)
    return out


def _bs_series(con, ticker, item_key, vis):
    """(period_end, value) の昇順。別枠 get_bs_series と同義。"""
    return [(pe, v) for pe, v in con.execute(
        "SELECT period_end, value FROM balance_sheet_items WHERE ticker=? "
        "AND item_key=? ORDER BY period_end", (ticker, item_key))
        if vis is None or pe in vis]


def _series(con, ticker, item, vis, need=None):
    """span-matched の比較系列を PIT で絞ったもの。

    need を渡すと **当期と前年の両方でその BS 項目が取れる期だけ**に絞る。
    別枠 S1 が `pl = [r for r in pl if r["period_end"] in ar]` としてから
    末尾を取るのと同じ挙動 —— BS が無い最新期で打ち切ると、その1期のせいで
    シグナルが丸ごと消える（2026-09-02 の初版で実際にそうなっていた）。
    """
    ser = SM.yoy_series(con, ticker, item)
    if vis is not None:
        ser = [r for r in ser
               if r["period_end"] in vis and r["peer_period"] in vis]
    if need is not None:
        ser = [r for r in ser
               if r["period_end"] in need and r["peer_period"] in need]
    return ser


def doc_for(con, ticker, period_end):
    """その期を初めて開示した書類。**「根拠期の原文」はこれ。**

    最新の開示に飛ばすと、根拠になった期と別の書類を「原文」と称する
    ことになる（2026-09-02 に指摘された弱点）。期→書類の対応は
    投影層の filings が持っているので、それを引くだけでよい。
    """
    r = con.execute(
        "SELECT doc_id, filing_date, source, pdf_path FROM filings "
        "WHERE ticker=? AND period_end=? ORDER BY generation DESC LIMIT 1",
        (ticker, period_end)).fetchone()
    if not r:
        return {}
    return {"doc_id": r[0], "doc_date": r[1], "doc_source": r[2], "doc_path": r[3]}


def _staleness(con, ticker, period_end, vis):
    """その期が「直近の開示期」より何四半期古いか。分からなければ None。"""
    latest = None
    for (pe,) in con.execute(
            "SELECT DISTINCT period_end FROM quarterly_standalone_all "
            "WHERE ticker=? AND is_valid=1 AND sales IS NOT NULL", (ticker,)):
        if vis is not None and pe not in vis:
            continue
        if latest is None or pe > latest:
            latest = pe
    a, b = SM.parse_pe(latest or ""), SM.parse_pe(period_end)
    if not a or not b:
        return None
    return SM._seq(*a) - SM._seq(*b)


# ------------------------------------------------------------------ S1
def s1_dso(con, ticker, as_of=None):
    """S1: DSO の改善。**同じ span の前年と比べる。**

    別枠は pl[-5] を前年同期とみなすが、チェーンが切れると位置が意味を
    失う。ここは find_yoy_peer で明示的に前年を取る。
    日数を span に合わせるので、DSO の絶対値も「その期間の日数」で正しく出る
    （比率変化は span が同じなら日数が約分されるため、スコアは別枠と一致する）。
    """
    vis = visible_periods(con, ticker, as_of)
    ar = dict(_bs_series(con, ticker, "accounts_receivable", vis))
    ser = _series(con, ticker, "sales", vis, need=ar)
    if not ser:
        return _result(0.0, False,
                       "売掛金と売上が当期・前年で揃う同じ span の組が無い")
    cur = ser[-1]
    stale = _staleness(con, ticker, cur["period_end"], vis)
    if stale is not None and stale > MAX_STALE_Q:
        return _result(0.0, False,
                       "前年同期と組める最後の期(%s)が直近開示より%d四半期古い"
                       % (cur["period_end"], stale))
    adj = _adjustments(con, ticker, vis)
    s_now = cur["value"] - adj.get(cur["period_end"], 0)
    s_prev = cur["peer_value"] - adj.get(cur["peer_period"], 0)
    if s_now <= 0 or s_prev <= 0:
        return _result(0.0, False, "調整後売上がゼロ以下（一時収入のみの期）")

    days = DAYS_PER_Q * cur["span_q"]
    dso_now = ar[cur["period_end"]] / s_now * days
    dso_prev = ar[cur["peer_period"]] / s_prev * days
    change = dso_now / dso_prev - 1
    score = clamp(-change / 0.20)
    confirm = ""
    if s_now < s_prev:
        score *= 0.5
        confirm = "（売上減少下のため半減）"
    # 売上方向ガード（2026-09-18、calibration_backlog §32）。
    # **ここでは score を変えない**—— 採否は policy（evidence_strict）の
    # 仕事にする。事前登録済みの測定（paper / stale_audit / sanity_check）は
    # prefer_span を明示しているので、黙って測定条件が変わらない。
    sd = SD.evaluate(con, ticker, cur["period_end"], vis)
    ev = ("DSO %.0f日→%.0f日（%+.1f%%）／調整後売上 %.0f→%.0f百万円%s"
          % (dso_prev, dso_now, change * 100, s_prev, s_now, confirm))
    ev += "／売上方向 %s（%s）" % (sd["status"], sd["note"])
    shrink = None
    if sd["status"] == "down" and score > 0:
        # **分離した別シグナル（S1b）として残す。**減点にしない理由は
        # calibration_backlog §32-2。合成スコアには乗せない（0点・表示のみ）。
        shrink = {"signal": "S1b", "score": 0.0,
                  "evidence": ("縮小に伴う債権減：DSO %.0f日→%.0f日（%+.1f%%）だが "
                               "売上は減少（%s）。S1 の加点は取り消した"
                               % (dso_prev, dso_now, change * 100, sd["trend_text"])),
                  "dso_change_pct": round(change, 4),
                  "sales_trend": sd["trend"]}
    return _result(score, True, ev,
                   {"dso_now": round(dso_now, 1), "dso_prev_year": round(dso_prev, 1),
                    "change_pct": round(change, 4),
                    "sales_direction": sd["status"],
                    "sales_direction_quarter_level": sd["quarter_level"],
                    "sales_trend_text": sd["trend_text"]},
                   sales_direction=sd, shrink_signal=shrink,
                   mode=cur["mode"], span_q=cur["span_q"],
                   period=cur["period_end"], peer_period=cur["peer_period"],
                   peer_synthesized=cur.get("peer_synthesized", 0),
                   peer_synth_parts=cur.get("peer_synth_parts"),
                   **doc_for(con, ticker, cur["period_end"]))


# ------------------------------------------------------------------ S2
def s2_contract_liabilities(con, ticker, as_of=None):
    """S2: 契約負債の増加。

    BS はストック（期末時点の残高）なので span に依存しない。span を意識
    するのは**売上比の確認**だけ —— 6ヶ月の売上と3ヶ月の売上で割った比率を
    並べたら意味が壊れる。
    """
    vis = visible_periods(con, ticker, as_of)
    cl = dict(_bs_series(con, ticker, "contract_liabilities", vis))
    ser = _series(con, ticker, "sales", vis, need=cl)
    if not ser:
        return _result(0.0, False,
                       "契約負債と売上が当期・前年で揃う同じ span の組が無い")
    cur = ser[-1]
    stale = _staleness(con, ticker, cur["period_end"], vis)
    if stale is not None and stale > MAX_STALE_Q:
        return _result(0.0, False,
                       "前年同期と組める最後の期(%s)が直近開示より%d四半期古い"
                       % (cur["period_end"], stale))
    v_now, v_y = cl[cur["period_end"]], cl[cur["peer_period"]]
    if v_now < MIN_BASE and v_y < MIN_BASE:
        return _result(0.0, False, "契約負債が少額（%.0f百万円）で有意でない" % v_now)

    yoy = (v_now / v_y - 1) if v_y >= MIN_BASE else (1.0 if v_now >= MIN_BASE * 3 else 0.0)
    # 前期比: 別枠 cl[-2] と同じく「BS系列の1つ前」。span に依存しない量なので
    # 比較可能ペアの有無に関係なく取れる。
    order = _bs_series(con, ticker, "contract_liabilities", vis)
    prev = [x for x in order if x[0] < cur["period_end"]]
    qoq = (v_now / prev[-1][1] - 1) if prev and prev[-1][1] > 0 else 0.0

    adj = _adjustments(con, ticker, vis)
    s_now = cur["value"] - adj.get(cur["period_end"], 0)
    s_y = cur["peer_value"] - adj.get(cur["peer_period"], 0)
    ratio_term = 0.0
    if s_now > 0 and s_y > 0:
        ratio_term = clamp(((v_now / s_now) / (v_y / s_y) - 1) / 0.30)
    score = clamp(0.5 * clamp(yoy / 0.30) + 0.2 * clamp(qoq / 0.20) + 0.3 * ratio_term)
    ev = ("契約負債 %.0f→%.0f百万円（前年比%+.1f%%・前期比%+.1f%%）"
          % (v_y, v_now, yoy * 100, qoq * 100))
    return _result(score, True, ev,
                   {"v_now": v_now, "yoy": round(yoy, 4), "qoq": round(qoq, 4)},
                   mode=cur["mode"], span_q=cur["span_q"],
                   period=cur["period_end"], peer_period=cur["peer_period"],
                   peer_synthesized=cur.get("peer_synthesized", 0),
                   peer_synth_parts=cur.get("peer_synth_parts"),
                   **doc_for(con, ticker, cur["period_end"]))


# ------------------------------------------------------------------ S4
def s4_cash_quality(con, ticker, as_of=None):
    """S4: 営業CF vs 営業利益（キャッシュ品質）。直近1年ぶんで評価。

    別枠は「CF開示期の直前Qを含めて半期単位で集計」という位置ベースの
    組み立てをする。span-matched では **CF と OP が同じ行に同じ span で
    載っている**ので、行を span 合計が4（＝1年）になるまで後ろから取れば
    期間対応が定義から保証される。組み立てが要らないぶん壊れようがない。
    """
    vis = visible_periods(con, ticker, as_of)
    rows = [r for r in con.execute(
        "SELECT period_end, span_q, operating_cf, operating_profit "
        "FROM quarterly_standalone_all WHERE ticker=? AND is_valid=1 "
        "AND operating_cf IS NOT NULL ORDER BY period_end", (ticker,))
        if vis is None or r[0] in vis]
    if not rows:
        return _result(0.0, False, "営業CFの開示期が不足（1Q/3Qは短信にCF記載なし）")

    # **span の合計が4になればよい、ではない。期間が重なってはいけない。**
    # 同じ期に span=1 と span=2 の行が両方あることは普通にあるので
    # （FY2026-Q2 の3ヶ月単独と上期6ヶ月）、合計だけ見て拾うと
    # 同じ四半期を二重に数える。**直前の区間の開始点に、次の区間の
    # 終了点がぴったり接する**ことを要求して後ろから辿る。
    by_end = {}
    for r in rows:
        by_end.setdefault(r[0], []).append(r)
    latest = max(rows, key=lambda r: r[0])[0]
    picked, total, end = [], 0, SM.parse_pe(latest)
    if not end:
        return _result(0.0, False, "期の表記を解釈できない: %s" % latest)
    cursor = SM._seq(*end)
    while total < 4:
        pe = "FY%d-Q%d" % (cursor // 4, cursor % 4 + 1)
        cand = [r for r in by_end.get(pe, ()) if total + (r[1] or 1) <= 4]
        if not cand:
            break
        r = max(cand, key=lambda x: x[1] or 1)   # 隙間を作らない最長を採る
        picked.append(r)
        total += r[1] or 1
        cursor -= (r[1] or 1)                    # 直前の区間の終端へ移る
    if total != 4:
        return _result(0.0, False,
                       "営業CFの開示が連続した1年ぶんに満たない（span合計%d）" % total)
    if any(r[3] is None for r in picked):
        return _result(0.0, False, "対応期間の営業利益が欠けている")

    sum_cf = sum(r[2] for r in picked)
    sum_op = sum(r[3] for r in picked)
    if sum_op <= 0:
        return _result(0.0, False,
                       "対象期間の営業利益が%.0f百万円（赤字期は品質評価しない）" % sum_op)
    ratio = sum_cf / sum_op
    score = clamp((ratio - 1.0) / 0.5)
    latest = picked[0]
    ev = ("直近1年 営業CF %.0f百万円 / 営業利益 %.0f百万円 = %.2f倍"
          "（span %s の積み上げ）"
          % (sum_cf, sum_op, ratio,
             "+".join(str(r[1] or 1) for r in reversed(picked))))
    covered = " + ".join("%s(span=%d)" % (r[0], r[1] or 1) for r in reversed(picked))
    return _result(score, True, ev, {"cf_op_ratio": round(ratio, 3)},
                   mode=SM.granularity_mode(latest[1]),
                   span_q=latest[1], period=latest[0], peer_period=None,
                   # **複数期を積むスコアはリンク1本に落とせない。**
                   # 最新の根拠期に飛ばし、どこまでを含むかを注記で示す。
                   period_note="根拠期間 %s（リンクは最新期の書類）" % covered,
                   **doc_for(con, ticker, latest[0]))


# ------------------------------------------------------------------ S5
def _ytd(con, ticker, fy, vis):
    """期首から連続して埋まる最長区間（YTD累計）を組み立てる。

    別枠は「その年度の単独値を全部足す」で累計を作る。四半期が揃って
    いる年度ではそれで正しいが、1Q/3Q が消えた年度では**足せる行が
    半期1本しかない**。逆に半期1本しか無くても、それ自体が期首からの
    6ヶ月累計なので進捗率は出せる。

    そこで「期首(Q1)から隙間なく覆える最長の区間」を、**span の大小に
    関係なく**組み立てる。四半期が揃っていれば 別枠 と同じ足し算になり、
    半期しか無ければ半期1本になる。どちらも同じ規則の帰結。

    戻り値: (period_end, n_quarters, sales, operating_profit) または None。
    n=4（通期が揃った）は「通期発表済み」として対象外にする ——
    別枠が latest_q == "FY" を外すのと同じ判断。
    """
    by_start = {}
    for r in con.execute(
            "SELECT period_end, period_start, span_q, sales, operating_profit "
            "FROM quarterly_standalone_all WHERE ticker=? AND fiscal_year=? "
            "AND is_valid=1 AND sales IS NOT NULL", (ticker, fy)):
        if vis is not None and r[0] not in vis:
            continue
        st = SM.parse_pe(r[1] or r[0])
        if not st or st[0] != fy:
            continue
        by_start.setdefault(st[1], []).append(r)
    # 各開始四半期から到達できる最遠点を DP で求める（長い区間を優先しない。
    # 「最も遠くまで覆える」組み合わせを選ぶ）。
    best = {1: (0, [])}                 # start_q -> (到達Q, 使った行)
    for q in range(1, 5):
        if q not in best:
            continue
        reach, used = best[q]
        for r in by_start.get(q, []):
            nq = q + (r[2] or 1) - 1
            if nq > 4:
                continue
            cand = (nq, used + [r])
            if nq + 1 not in best or len(cand[1]) < len(best[nq + 1][1]):
                best[nq + 1] = cand
    end = max((k - 1 for k in best if k > 1), default=0)
    if end < 1 or end >= 4:
        return None                     # 覆えない、または通期が揃っている
    rows = best[end + 1][1]
    ops = [r[4] for r in rows]
    return (max(r[0] for r in rows), end, sum(r[3] for r in rows),
            None if any(o is None for o in ops) else sum(ops))


def s5_progress(con, ticker, as_of=None):
    """S5: 通期予想に対する累計進捗率の異常。**span-matched と最も相性が良い。**

    進捗率はもともと「期首からの累計 ÷ 通期予想」で、四半期単独値である
    必要がない。別枠は単独値を足し上げるので四半期チェーンを要求するが、
    ここは `_ytd` が半期1本でも累計を組み立てる。
    """
    a = _iso(as_of)
    fc = con.execute(
        "SELECT fiscal_year, forecast_sales, forecast_op FROM company_forecasts "
        "WHERE ticker=?" + (" AND source_date<=?" if a else "")
        + " ORDER BY source_date DESC LIMIT 1",
        (ticker, a) if a else (ticker,)).fetchone()
    if not fc:
        return _result(0.0, False, "会社通期予想なし")
    fy, fc_sales, fc_op = fc[0], fc[1], fc[2]

    vis = visible_periods(con, ticker, as_of)
    ytd = _ytd(con, ticker, fy, vis)
    if not ytd:
        return _result(0.0, False, "当該年度の期首からの累計を組み立てられない"
                                   "（未開示、または通期発表済み）")
    pe, span, cum_sales, cum_op = ytd

    adj = _adjustments(con, ticker, vis)
    ded = sum(v for k, v in adj.items()
              if k.startswith("FY%d-" % fy) and k <= pe)
    cum_sales_rep = cum_sales
    cum_sales = cum_sales - ded
    cum_op = None if cum_op is None else cum_op - ded

    expected = _seasonal_share(con, ticker, fy, span, vis)
    if expected is None:
        expected = {1: 0.25, 2: 0.50, 3: 0.75}.get(span)
    if not expected:
        return _result(0.0, False, "季節性の期待値を決められない")

    prog_s = cum_sales / fc_sales if fc_sales else 0
    ratio_s = prog_s / expected if expected > 0 else 1.0
    if fc_op and fc_op > 0 and cum_op is not None:
        prog_o = cum_op / fc_op
        ratio_o = prog_o / expected if expected > 0 else 1.0
        ratio = 0.4 * ratio_s + 0.6 * ratio_o
        op_note = "、調整後OP進捗 %.1f%%" % (prog_o * 100)
    else:
        prog_o, ratio, op_note = None, ratio_s, ""
    score = clamp((ratio - 1.0) / 0.25)

    # ---- 経過四半期数で正規化した OP 進捗と、残存四半期の暗黙利益
    # （2026-09-18・calibration_backlog §33）。
    # 季節性シェアは「過去年度の平均的な出方」でしかなく、過去年度が
    # 取れない銘柄では 0.25/0.50/0.75 に落ちる。**会社予想が死んでいる
    # かどうかは、残りの四半期でいくら稼ぐことになっているかを引き算
    # すれば直接分かる**。累計 OP が通期予想を超えていれば、暗黙の残存
    # 四半期は赤字という不合理になる。
    pace = span / 4.0
    prog_o_pace = imp_rest = imp_rest_per_q = None
    dead = False
    if fc_op and fc_op > 0 and cum_op is not None:
        prog_o_pace = cum_op / fc_op
        imp_rest = fc_op - cum_op
        if span < 4:
            imp_rest_per_q = imp_rest / (4 - span)
        dead = imp_rest < 0
    pace_excess_pt = (None if prog_o_pace is None
                      else (prog_o_pace - pace) * 100.0)

    rep = ("／報告値ベース売上進捗 %.1f%%" % (cum_sales_rep / fc_sales * 100)
           if fc_sales and cum_sales_rep != cum_sales else "")
    ev = ("Q1〜Q%d累計：調整後売上進捗 %.1f%%%s（季節性期待 %.0f%%）%s "
          "→ 期待比 %.2f倍"
          % (span, prog_s * 100, op_note, expected * 100, rep, ratio))
    if prog_o_pace is not None:
        ev += ("／OP進捗 %.0f%%（経過四半期比 %.0f%%、超過 %+.0fpt）"
               "／暗黙の残存 %d四半期 OP %.0f百万円"
               % (prog_o_pace * 100, pace * 100, pace_excess_pt,
                  4 - span, imp_rest))
        if imp_rest_per_q is not None:
            ev += "（1Qあたり %.0f）" % imp_rest_per_q
        if dead:
            ev += " ← **死んだガイダンス（残存四半期が暗黙の赤字）**"
    return _result(score, True, ev,
                   {"progress_sales": round(prog_s, 4),
                    "expected": round(expected, 4), "ratio": round(ratio, 3),
                    "progress_op": (None if prog_o_pace is None
                                    else round(prog_o_pace, 4)),
                    "elapsed_q": span, "pace": round(pace, 4),
                    "pace_excess_pt": (None if pace_excess_pt is None
                                       else round(pace_excess_pt, 1)),
                    "implied_rest_op": (None if imp_rest is None
                                        else round(imp_rest, 1)),
                    "implied_rest_op_per_q": (None if imp_rest_per_q is None
                                              else round(imp_rest_per_q, 1)),
                    "guidance_dead": bool(dead),
                    "forecast_op": fc_op, "cum_op": (None if cum_op is None
                                                     else round(cum_op, 1))},
                   guidance_dead=bool(dead),
                   mode=SM.granularity_mode(span), span_q=span, period=pe,
                   peer_period=None,
                   period_note=("根拠期間 FY%d-Q1〜%s の累計（リンクは最新期の書類）"
                                % (fy, pe)),
                   **doc_for(con, ticker, pe))


def _seasonal_share(con, ticker, fy, span, vis):
    """過去年度の「Q1〜Q{span} の累計が通期に占める割合」の平均。

    別枠 `_seasonal_share` と同じ定義（過去の全年度を使う。直近数年に
    絞らない）。分母の通期は**単独値4本の合計**を優先し、無ければ
    span=4 の行を使う —— 別枠は前者しか持たないが、四半期が消えた年度は
    後者しか無い。
    """
    years = [r[0] for r in con.execute(
        "SELECT DISTINCT fiscal_year FROM quarterly_standalone_all "
        "WHERE ticker=? AND fiscal_year<? AND is_valid=1", (ticker, fy))]
    shares = []
    for y in sorted(years):
        q1 = [r for r in con.execute(
            "SELECT period_end, sales FROM quarterly_standalone_all WHERE ticker=? "
            "AND fiscal_year=? AND span_q=1 AND is_valid=1 AND sales IS NOT NULL "
            "ORDER BY period_end", (ticker, y))
            if vis is None or r[0] in vis]
        if len(q1) == 4:
            full = sum(r[1] for r in q1)
        else:
            f = con.execute(
                "SELECT period_end, sales FROM quarterly_standalone_all WHERE ticker=? "
                "AND fiscal_year=? AND span_q=4 AND is_valid=1 AND sales IS NOT NULL "
                "LIMIT 1", (ticker, y)).fetchone()
            if not f or (vis is not None and f[0] not in vis):
                continue
            full = f[1]
        if not full or full <= 0:
            continue
        cum = _ytd_for_span(con, ticker, y, span, vis, q1)
        if cum is not None:
            shares.append(cum / full)
    return sum(shares) / len(shares) if shares else None


def _ytd_for_span(con, ticker, fy, span, vis, q1_rows):
    """その年度の Q1〜Q{span} 累計売上。組み立てられなければ None。"""
    if len(q1_rows) >= span:
        return sum(r[1] for r in q1_rows[:span])
    r = con.execute(
        "SELECT period_end, sales FROM quarterly_standalone_all WHERE ticker=? "
        "AND fiscal_year=? AND span_q=? AND period_start=? AND is_valid=1 "
        "AND sales IS NOT NULL LIMIT 1",
        (ticker, fy, span, "FY%d-Q1" % fy)).fetchone()
    if r and (vis is None or r[0] in vis):
        return r[1]
    return None


SPAN_SCORERS = {
    "S1": s1_dso,
    "S2": s2_contract_liabilities,
    "S4": s4_cash_quality,
    "S5": s5_progress,
}
