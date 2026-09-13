"""screener/projection/adapter.py — 別枠 earnings_screener への投影層。

統合引継ぎ書 v1.1 §3-1 の4関数と同じシグネチャを本体側で実装する。
**別枠のコードは1行も変更しない。** import 先を data_access からここへ
差し替えるだけで繋がる形にしてある。理由:

  - 別枠は tests/test_pit.py で不変条件が固定されており、書き換えると
    その保証を失う
  - v1.1 の generation 機構は「モック側でのPIT検証用」と v1.1 §0-1 自身が
    位置づけている。本体は financials_cum(filing_id, item, context_ref) で
    同等以上を自然に表現できるので、本体側で実装し直す方が正しい

## 返さないものは黙って返さない

写像が unverified / missing の item_key を要求されたら、空を返したうえで
**なぜ返せないかをログに出す**。0 や None を黙って返すと、下流は
「データが無い」と「まだ繋いでいない」を区別できなくなる —— 本体が
一貫して避けてきた失敗そのもの。
"""
from __future__ import annotations

import os
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.projection import pit

_CFG = None
_WARNED: set = set()


def cfg() -> dict:
    global _CFG
    if _CFG is None:
        _CFG = C.load_yaml("earnings_screener_mapping.yaml")
    return _CFG


def resolve_item(item_key: str) -> str | None:
    """別枠 item_key → 本体 item。繋げないものは None を返しログに残す。"""
    e = (cfg().get("items") or {}).get(item_key)
    if e is None:
        _warn(item_key, f"写像表に item_key '{item_key}' が無い")
        return None
    if e.get("status") != "verified":
        _warn(item_key,
              f"item_key '{item_key}' は status={e.get('status')} のため未接続 "
              f"({str(e.get('note') or '').strip()[:80]})")
        return None
    return e.get("item")


def _warn(key: str, msg: str) -> None:
    if key not in _WARNED:            # 同じ警告で画面を埋めない
        _WARNED.add(key)
        C.log(f"[adapter] {msg}")


def _div() -> float:
    return float(cfg().get("unit_divisor") or 1.0)


def _qtype(q_no: int) -> str | None:
    return (cfg().get("quarter_type") or {}).get(q_no)


def _fy(period: str) -> int | None:
    if period and period.startswith("FY") and period[2:].isdigit():
        return int(period[2:])
    return None


def _period_end(period: str, q_no: int) -> str:
    """別枠は period_end を系列のキーに使う。本体は (period, q_no) を持つので
    ソート可能な合成キーを作る。実日付ではないので日付として解釈しないこと。"""
    return f"{period}-Q{q_no}"


# ------------------------------------------------------- §3-1 の4関数
# 単独値のフロー項目。BS(ストック)項目と違い、span_q が「何ヶ月ぶんか」を
# 直接意味するので、粒度の判定はこれらだけを見る。
_FLOW_ITEMS = ("revenue", "operating_income", "gross_profit", "operating_cf")
_ALL_SPANS = (1, 2, 3, 4)


def get_pl_series(con, ticker, as_of=None, *, mode: str = "strict",
                  allow_span=None) -> list[dict]:
    """単独値PL時系列。有効行のみ・古い順。

    既定は span_q=1（真の四半期単独値）のみ。span_q=2（半期粒度）は別枠に
    対応概念が無く、渡すと6ヶ月の値を四半期として扱う。戻り値には span_q と
    is_half を必ず含め、受け手が無視できないようにする。
    """
    allow = tuple(allow_span or cfg().get("allowed_span_q") or (1,))
    # 許容 span の行だけでなく全 span を引く。理由: BS項目(span=1)だけが
    # 残った行が「フローは半期粒度なので落とした」ことを名乗れず、
    # span_q=1 / is_half=False のまま sales=None を返していた。
    # 「開示が無い」と「粒度が合わないので渡さない」は別物で、
    # 受け手がそこを区別できないのが投影層として最も高くつく嘘になる。
    all_rows = pit.visible_q(con, ticker, as_of, mode, allow_span=_ALL_SPANS)
    by_pq: dict = {}
    by_pq_all: dict = {}
    for r in all_rows:
        by_pq_all.setdefault((r["period"], r["q_no"]), {})[r["item"]] = r
        if (r["span_q"] or 1) in allow:
            by_pq.setdefault((r["period"], r["q_no"]), {})[r["item"]] = r
    div = _div()

    out = []
    for (period, q_no), items in sorted(by_pq.items()):
        qt = _qtype(q_no)
        if qt is None:
            continue
        # span はフロー項目の実体から決める。BS項目の span を借りると、
        # フローを落とした行が「四半期粒度」を名乗ってしまう。
        flow = [items[k] for k in _FLOW_ITEMS if k in items]
        span = next((it["span_q"] or 1 for it in flow), None)
        excluded = None
        if span is None:
            dropped = [by_pq_all[(period, q_no)][k]
                       for k in _FLOW_ITEMS if k in by_pq_all.get((period, q_no), {})]
            if dropped:
                excluded = next((it["span_q"] or 1 for it in dropped), None)
            span = next((it["span_q"] or 1 for it in items.values()), 1)
        def v(name):
            r = items.get(name)
            return None if r is None or r["value"] is None else r["value"] / div
        out.append({
            "period_end": _period_end(period, q_no),
            "quarter_type": qt,
            "fiscal_year": _fy(period),
            "sales": v("revenue"),
            "operating_profit": v("operating_income"),
            "gross_profit": v("gross_profit"),
            # 原価の数量/価格分解は本体にも XBRL にも該当科目が無い。
            # 別枠 S6 の入力だが作れないので None を明示する（0 にしない）。
            "cogs_quantity": None,
            "cogs_price": None,
            "operating_cf": v("operating_cf"),
            "span_q": span,
            "is_half": (excluded or span) == 2,
            # フローが「開示されていない」のか「粒度が合わず落とした」のかを
            # 受け手が区別できるようにする。None なら前者。
            "flow_span_excluded": excluded,
            "period": period,
            "q_no": q_no,
        })
    return out


def get_bs_series(con, ticker, item_key, as_of=None, *,
                  mode: str = "strict") -> list[tuple]:
    """BS科目のスナップショット。[(period_end, value)] を古い順で返す。"""
    item = resolve_item(item_key)
    if item is None:
        return []
    rows = pit.visible_q(con, ticker, as_of, mode, allow_span=(1, 2, 3, 4))
    div = _div()
    out = [( _period_end(r["period"], r["q_no"]), r["value"] / div)
           for r in rows if r["item"] == item and r["value"] is not None]
    return sorted(out)


def get_adjustments(con, ticker, as_of=None, *, mode: str = "strict") -> dict:
    """一時収入等の調整。period_end → 売上から控除する合計額。

    one_time_revenue のみを売上控除に使う。one_time_cost / one_time_gain は
    利益側の調整なので別関数（get_adjustment_detail）で返す。
    """
    raw = pit.visible_adjustments(con, ticker, as_of, mode)
    div = _div()
    return {_period_end(p, q): sum(v for k, v in d.items()
                                   if k == "one_time_revenue") / div
            for (p, q), d in raw.items()
            if any(k == "one_time_revenue" for k in d)}


def get_adjustment_detail(con, ticker, as_of=None, *,
                          mode: str = "strict") -> dict:
    """period_end → {item_key: 金額}。根拠つきで何を引いたかを見せるため。"""
    raw = pit.visible_adjustments(con, ticker, as_of, mode)
    div = _div()
    return {_period_end(p, q): {k: v / div for k, v in d.items()}
            for (p, q), d in raw.items()}


def get_latest_forecast(con, ticker, as_of=None, *, mode: str = "strict"):
    """会社通期予想の最新版（as_of 指定でその時点の版）。v1.1 §3-1。"""
    row = pit.visible_guidance(con, ticker, "operating_income", as_of, mode)
    if row is None:
        return None
    sales = pit.visible_guidance(con, ticker, "revenue", as_of, mode)
    div = _div()
    return {"fiscal_year": _fy(row["fy"]),
            "forecast_op": (row["value"] or 0) / div,
            "forecast_sales": ((sales["value"] or 0) / div) if sales else None,
            "source_date": row["date"]}


# ------------------------------------------------------------- 補助
def normalized_sales(pl_row, adjustments) -> float:
    """調整後売上。別枠 data_access.normalized_sales と同じ意味論。"""
    return (pl_row["sales"] or 0) - adjustments.get(pl_row["period_end"], 0)


def unmapped_keys() -> dict:
    """繋がっていない item_key の一覧。統合状況の可視化用。"""
    out = {"unverified": [], "missing": []}
    for k, e in (cfg().get("items") or {}).items():
        st = e.get("status")
        if st in out:
            out[st].append(k)
    return out


# --------------------------------------------------------------- 価格の投影
# 別枠は daily_prices ビュー(schema.sql)を素のSQLで読む。ここに置く関数は
# 同じ意味論を Python から使うためのもので、**唯一の正本はビューのほう**。
# 二重定義にならないよう、この層は必ずビュー経由で読む。

def get_price_series(con, ticker: str, as_of=None, *, start=None) -> dict:
    """調整後終値の時系列 {date: close} を返す。

    close は **調整後終値**。未調整終値へ落ちる経路は用意しない
    （adj_close が NULL の日はビューが行ごと落とすので、ここにも来ない）。

    as_of: その日までに市場が知り得た価格だけを返す。株価は当日の引けで
    公知になるので境界は `date <= as_of` —— 開示情報の strict/lax とは
    別の話で、ここに strict は無い。
    """
    sql = "SELECT date, close FROM daily_prices WHERE ticker = ?"
    args = [ticker]
    if start:
        sql += " AND date >= ?"
        args.append(pit._as_of_str(start))
    if as_of:
        sql += " AND date <= ?"
        args.append(pit._as_of_str(as_of))
    return {r["date"]: r["close"] for r in con.execute(sql + " ORDER BY date", args)}


def get_universe_metrics(con, ticker: str, as_of, *, adv_days: int = 20) -> dict:
    """as_of 時点のユニバース判定に要る指標を返す。

    {"date":…, "mktcap":…, "adv_turnover":…, "n_days":…} または空 dict。

    なぜ要るか: 「2023年時点でユニバースに入っていたか」を今の companies 表で
    代用すると、当時は条件を満たしていたのに今は外れている銘柄が消え、
    サンプルが生き残りだけに偏る（サバイバーシップバイアス）。時価総額と
    売買代金を**その時点の値**で持っておかないと過去再現ができない。

    adv_turnover は as_of 以前の直近 adv_days 営業日の売買代金の平均。
    日数が足りない場合は n_days に実数を返す（黙って薄い平均を返さない）。

    単位に注意: **mktcap は百万円、adv_turnover は円**（J-Quants の返す
    ままで、投影層では変換しない）。仕様書のユニバース条件に当てるなら
      時価総額 50〜1,000億円     -> 5000 <= mktcap <= 100000
      20日平均売買代金 3,000万円 -> adv_turnover >= 30_000_000
    桁を取り違えても例外にはならず「該当0件」として静かに出るので、
    ここを読まずに閾値を書かないこと。
    """
    a = pit._as_of_str(as_of)
    rows = list(con.execute(
        "SELECT date, mktcap, turnover_value FROM daily_prices "
        "WHERE ticker = ? AND date <= ? ORDER BY date DESC LIMIT ?",
        (ticker, a, int(adv_days))))
    if not rows:
        _warn(f"prices:{ticker}", f"{a} 以前に調整後終値のある日が無い")
        return {}
    vals = [r["turnover_value"] for r in rows if r["turnover_value"] is not None]
    return {
        "date": rows[0]["date"],
        "mktcap": rows[0]["mktcap"],
        "adv_turnover": (sum(vals) / len(vals)) if vals else None,
        "n_days": len(vals),
    }
