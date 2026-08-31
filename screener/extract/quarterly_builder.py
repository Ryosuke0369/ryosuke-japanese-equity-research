"""screener/extract/quarterly_builder.py — 四半期単独値ビルダー (仕様書 §3-2)。

本システムの心臓。短信は累計表示なので、

    Q単独値 = 当期累計 − 前四半期累計

を全項目で機械生成する。3441 で実証済みのとおり、9ヶ月累計の粗利率24.2%の裏で
Q3単独は21.9%まで落ちていた。**累計は平均で嘘をつく。単独値だけが傾きを語る**。

無効化(valid_flag=0)する場合 —— 仕様書が名指しした3つに、実装上どうしても
必要な2つを足した。**行は消さない**。「無効」と「まだデータが無い」を
取り違えないため。

  決算期変更        決算期末が変わると四半期の長さが変わり、差分が意味を失う
  遡及修正          同じ(期,四半期,項目)を後の書類が違う値で言い直した
  連結範囲変更      連結子会社数が変わった。前四半期と同じ会社を見ていない
  前四半期累計なし  引き算の相手が無い（取得漏れか、まだ開示されていない）
  四半期不明        書類から第何四半期か決められなかった（推測しない）

Q1 は累計そのものが単独値なので引き算しない。

Usage
    python -m screener.extract.quarterly_builder --all
    python -m screener.extract.quarterly_builder --code 3441 --verbose
    python -m screener.extract.quarterly_builder --report
"""
from __future__ import annotations

import argparse
import os
import sys
from collections import defaultdict

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

# 累計から単独値を作ってよいのはフロー項目だけ。ストック項目(BS)は期末残高
# なので引き算してはいけない —— 総資産の「四半期単独値」に意味は無い。
FLOW_ITEMS = {
    "revenue", "cogs", "gross_profit", "sga", "operating_income",
    "ordinary_income", "profit_before_tax", "net_income_parent",
    "profit_loss_total", "operating_cf", "investing_cf", "financing_cf",
    "segment_revenue_external", "segment_depreciation", "segment_capex",
    "segment_intersegment",
}
# ストック項目。単独値は作らず、期末残高をそのまま q_no 付きで持つ。
STOCK_ITEMS = {
    "total_assets", "net_assets", "owners_equity", "inventories_total",
    "merchandise_and_finished_goods", "work_in_process",
    "raw_materials_and_supplies", "trade_receivables", "trade_payables",
    "contract_assets", "contract_liabilities", "advances_received",
    "construction_in_progress", "machinery_and_equipment",
    "property_plant_and_equipment", "cash_and_deposits", "short_term_loans",
    "long_term_loans", "electronically_recorded_claims",
    "electronically_recorded_obligations", "treasury_stock",
}
REL_TOL = 0.001          # 遡及修正判定の相対許容誤差


def _cum_rows(con, code: str | None):
    """financials_cum から「見出しの実績値」を取る。

    予想は guidance テーブルに分かれているのでここには来ない。次元付きは
    financials_dim に分かれているのでここには来ない（設計の効き目）。
    """
    sql = ("SELECT fc.code, fc.period, fc.q_no, fc.item, fc.value, fc.context_ref, "
           "       fc.source_tag, fc.filing_id, f.date AS filing_date, f.source "
           "FROM financials_cum fc JOIN filings f ON f.id = fc.filing_id "
           "WHERE fc.q_no IS NOT NULL AND fc.value IS NOT NULL ")
    args: tuple = ()
    if code:
        sql += "AND fc.code = ? "
        args = (code,)
    sql += "ORDER BY fc.code, fc.period, fc.item, fc.q_no, f.date"
    return con.execute(sql, args).fetchall()


def _is_nonconsolidated(context_ref: str | None) -> bool:
    return "NonConsolidatedMember" in (context_ref or "")


def _pick_latest(rows) -> tuple[dict, dict]:
    """(code, period, item, q_no) → 最新書類の累計値。あわせて遡及修正を検出。

    **連結と単体を必ず分ける。** 有報は同じ期・同じ項目を連結と単体の両方で
    載せる(3441 FY2023 の売上は連結95.6億/単体75.8億)。分けずに1つのキーへ
    詰めると、連結と単体の差を「後の書類が言い直した」= 遡及修正と誤検出する。
    実際に 3441 で65件の偽陽性が出た。

    連結を正とし、連結が無い期だけ単体を使う(単体しか出さない会社があるため)。
    遡及修正の判定も同じ基準どうしで比べる。
    """
    picked: dict = {}
    by_tag: dict = {}
    restated: set = set()
    for r in rows:
        basis = "nc" if _is_nonconsolidated(r["context_ref"]) else "c"
        key = (r["code"], r["period"], r["item"], r["q_no"], basis)

        # 遡及修正の判定は**同じXBRLタグどうし**でしか成り立たない。
        # 一つの内部項目に複数のタグが入る(3441 の trade_payables は有報が
        # NotesAndAccountsPayableTrade=4.2億、半期が AccountsPayableTrade=3.0億)。
        # タグを見ずに比べると、別勘定の差を「言い直した」と読む。
        tkey = key + (r["source_tag"],)
        tprev = by_tag.get(tkey)
        if tprev is not None:
            a, b = tprev, r["value"]
            if abs(a - b) > REL_TOL * max(abs(a), abs(b), 1.0):
                # 項目単位で立てる。一つの騒がしい項目が期まるごとを無効に
                # しては、生きている売上・粗利まで捨てることになる。
                restated.add((r["code"], r["period"], r["q_no"], r["item"]))
        by_tag[tkey] = r["value"]

        prev = picked.get(key)
        if prev is None or (r["filing_date"] or "") >= (prev["filing_date"] or ""):
            picked[key] = {"value": r["value"], "filing_date": r["filing_date"],
                           "filing_id": r["filing_id"], "source": r["source"]}

    latest: dict = {}
    for (code, period, item, q, basis), v in picked.items():
        k = (code, period, item, q)
        if basis == "c":
            latest[k] = v                   # 連結が最優先。常に上書きする
        elif k not in latest:
            latest[k] = v                   # 単体は連結が無いときだけ
    return latest, restated


def _scope_changes(latest) -> dict:
    """連結子会社数が変わった (code, period, q_no) を返す。

    数が変われば前四半期と同じ会社を見ていない。累計の差分で作った単独値は
    前四半期と比較可能ではないので無効化する(仕様書 §3-2)。

    **精度の限界を明記する**: 子会社数は有報の期末値でしか取れないことが多く、
    期中のいつ変わったかは分からない。3441 は FY2024 の1社から FY2025 は2社に
    増えているが、それが上期か下期かは開示から読めない。したがって変化を検出
    した期は保守的に丸ごと無効化する —— 実際には下期だけ影響しているかも
    しれないが、「影響していない」と決めつけるよりは安全側に倒す。
    値は消さないので、無効理由を見たうえで人が採用することはできる。
    """
    by_code = defaultdict(dict)
    for (code, period, item, q), v in latest.items():
        if item == "consolidated_subsidiaries":
            by_code[code][(period, q)] = v["value"]
    out = set()
    for code, series in by_code.items():
        keys = sorted(series)
        for i in range(1, len(keys)):
            if series[keys[i]] != series[keys[i - 1]]:
                out.add((code, keys[i][0], keys[i][1]))
    return out


def _fiscal_changes(con) -> set:
    """決算期末の月が変わった (code, period)。

    EDINET の DEI から期末日を取れる書類だけで判定する。短信からは取れない
    ので、EDINET を持たない会社では検出できない —— 検出できないことを
    「変更が無い」と言い換えないよう、判定不能は無効化しない(理由が別だから)。
    """
    rows = con.execute(
        "SELECT fc.code, fc.period, MIN(f.date) AS d FROM financials_cum fc "
        "JOIN filings f ON f.id = fc.filing_id WHERE f.source='edinet' "
        "GROUP BY fc.code, fc.period").fetchall()
    # period は FY<year> なので月は持っていない。ここでは期の連番の欠落を見る。
    by_code = defaultdict(list)
    for r in rows:
        by_code[r["code"]].append(r["period"])
    out = set()
    for code, periods in by_code.items():
        ys = sorted(int(p[2:]) for p in periods if p.startswith("FY") and p[2:].isdigit())
        for i in range(1, len(ys)):
            if ys[i] - ys[i - 1] != 1:
                out.add((code, f"FY{ys[i]}"))       # 期が飛んでいる = 期変更の疑い
    return out


def build(con, code: str | None = None, verbose: bool = False) -> dict:
    rows = _cum_rows(con, code)
    latest, restated = _pick_latest(rows)
    scope = _scope_changes(latest)
    fiscal = _fiscal_changes(con)
    C.log(f"累計 {len(rows)} 行 / ユニークな (銘柄,期,項目,四半期) {len(latest)} 件")

    if code:
        con.execute("DELETE FROM financials_q WHERE code=?", (code,))
    else:
        con.execute("DELETE FROM financials_q")

    n = {"written": 0, "flow": 0, "stock": 0, "invalid": 0, "skipped_item": 0}
    reasons: dict[str, int] = defaultdict(int)
    now = C.utcnow()

    for (c, period, item, q), v in sorted(latest.items()):
        if item in STOCK_ITEMS:
            # ストックは引き算しない。期末残高をそのまま置く。
            con.execute(
                "INSERT OR REPLACE INTO financials_q (code, period, q_no, item, "
                " value, valid_flag, invalid_reason, span_q, built_at) "
                "VALUES (?,?,?,?,?,?,?,?,?)",
                (c, period, q, item, v["value"], 1, None, 1, now))
            n["written"] += 1
            n["stock"] += 1
            continue
        if item not in FLOW_ITEMS:
            n["skipped_item"] += 1
            continue

        reason = None
        if q is None:
            reason = "四半期不明"
        elif (c, period) in fiscal:
            reason = "決算期変更の疑い(期が連続していない)"
        elif (c, period, q, item) in restated:
            reason = "遡及修正(後の書類が違う値で言い直した)"
        elif (c, period, q) in scope:
            reason = "連結範囲変更(連結子会社数が変わった)"

        span = 1
        if q == 1:
            value = v["value"]              # Q1 は累計そのものが単独値
        else:
            # 直前の四半期が無ければ、同じ期の中でそれより手前の累計を探す。
            # 短信が揃っていない会社は有報(q4)と半期(q2)しか無く、その差は
            # 「下期6ヶ月」になる。値は作るが span_q に何四半期ぶんかを必ず
            # 残す —— 3ヶ月と6ヶ月を同じ土俵に載せると傾きが二重になる。
            prev = prev_q = None
            for back in range(q - 1, 0, -1):
                cand = latest.get((c, period, item, back))
                if cand is not None:
                    prev, prev_q = cand, back
                    break
            if prev is None:
                # 手前の累計が1つも無い期は、累計そのものが「期首からの単独値」。
                # 半期報告書しか無い会社の q2 がこれで、値は H1 そのもの。
                # Q1 と同じ形なので無効にはしない —— span_q に何四半期ぶんかを
                # 残せば、四半期と半期を取り違える危険は無い。
                value = v["value"]
                span = q
            else:
                value = v["value"] - prev["value"]
                span = q - prev_q

        flag = 0 if (reason or value is None) else 1
        if flag == 0:
            n["invalid"] += 1
            reasons[reason or "値を作れない"] += 1
        con.execute(
            "INSERT OR REPLACE INTO financials_q (code, period, q_no, item, value, "
            " valid_flag, invalid_reason, span_q, built_at) VALUES (?,?,?,?,?,?,?,?,?)",
            (c, period, q, item, value, flag, reason, span, now))
        n["written"] += 1
        n["flow"] += 1
        if verbose and value is not None:
            C.log(f"  {c} {period} Q{q} {item:<18} 単独={value:>16,.0f} "
                  f"{'' if flag else '[無効: ' + str(reason) + ']'}")

    con.commit()
    C.log(f"financials_q: {n['written']} 行 (フロー {n['flow']} / ストック {n['stock']}) "
          f"/ 無効 {n['invalid']}")
    for k, v in sorted(reasons.items(), key=lambda x: -x[1]):
        C.log(f"    {k:<44} {v:>6}")
    return n


def report(con) -> None:
    tot = con.execute("SELECT COUNT(*) c FROM financials_q").fetchone()["c"]
    ok = con.execute("SELECT COUNT(*) c FROM financials_q WHERE valid_flag=1").fetchone()["c"]
    codes = con.execute("SELECT COUNT(DISTINCT code) c FROM financials_q").fetchone()["c"]
    C.log(f"financials_q: {tot} 行 / 有効 {ok} / {codes} 銘柄")
    C.log("無効理由:")
    for r in con.execute("SELECT invalid_reason, COUNT(*) c FROM financials_q "
                         "WHERE valid_flag=0 GROUP BY invalid_reason ORDER BY c DESC"):
        C.log(f"  {str(r['invalid_reason']):<44} {r['c']:>6}")


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="四半期単独値ビルダー (仕様書 §3-2)")
    p.add_argument("--all", action="store_true")
    p.add_argument("--code")
    p.add_argument("--report", action="store_true")
    p.add_argument("--verbose", action="store_true")
    a = p.parse_args(argv)

    con = C.init_db()
    if a.report:
        report(con)
        return 0
    if not (a.all or a.code):
        p.error("--all か --code を指定する")
    build(con, a.code, a.verbose)
    report(con)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
