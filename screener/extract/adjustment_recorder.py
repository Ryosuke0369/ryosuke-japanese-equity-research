"""screener/extract/adjustment_recorder.py — 一時収入の検出キューと記録。

docs/adapter_design.md A-3。**検出は機械、確定は人、記録は構造化** の三分割。

なぜ機械抽出しないか
--------------------
3905 の手数料収入 55.8億円は短信の注記テキスト由来で、数値パーサー
(nonFraction のみを読む)の対象外。金額の確定には人が注記を読む必要がある。
これは一時的な技術制約ではなく、「何が一過性か」が判断を要する作業だから
恒久的に半手動にする。

  1. 検出   異常な期を機械が拾い、キューに積む（--scan）
  2. 確認   人が該当書類の注記を読む
  3. 記録   --record で登録。引用・所在・記録者が無いとDBが受け付けない
  4. 反映   投影層が normalized_sales に自動反映
  5. 監査   --queue で未処理の残件を出す（放置を可視化する）

Usage
    python -m screener.extract.adjustment_recorder --scan
    python -m screener.extract.adjustment_recorder --queue
    python -m screener.extract.adjustment_recorder --record \\
        --code 3905 --period FY2027 --q 1 --item one_time_revenue \\
        --amount 5580000000 --filing-id 1234 \\
        --locator "(セグメント情報等) 3. 報告セグメントごとの売上高" \\
        --note "当第1四半期連結累計期間において、...手数料収入5,580百万円を計上している" \\
        --by ryosuke
"""
from __future__ import annotations

import argparse
import os
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

ITEM_KEYS = ("one_time_revenue", "one_time_cost", "one_time_gain")


def _rules() -> dict:
    return C.load_yaml("adjustment_detection.yaml")["detect"]


def scan(con) -> list[dict]:
    """一過性の疑いがある期を拾う。**判定はしない。人に見せるだけ。**

    拾う条件は「前年同四半期と比べて説明のつかない動き」。閾値は
    config/adjustment_detection.yaml。ここを厳しくすると見逃し、緩めると
    キューが溢れる —— 溢れる方がまだ安全なので初期値は緩めにしてある。
    """
    r = _rules()
    q = con.execute(
        "SELECT code, period, q_no, item, value, span_q FROM financials_q "
        "WHERE valid_flag=1 AND value IS NOT NULL AND span_q=1 "
        "AND item IN ('revenue','gross_profit','operating_income')").fetchall()
    by = {}
    for x in q:
        by.setdefault((x["code"], x["period"], x["q_no"]), {})[x["item"]] = x["value"]

    def prev_year(code, period, q_no):
        if not (period.startswith("FY") and period[2:].isdigit()):
            return None
        return by.get((code, f"FY{int(period[2:]) - 1}", q_no))

    done = {(a["code"], a["period"], a["q_no"]) for a in con.execute(
        "SELECT DISTINCT code, period, q_no FROM pl_adjustments")}

    out = []
    for (code, period, q_no), cur in sorted(by.items()):
        if (code, period, q_no) in done:
            continue                       # 既に調整済み
        prev = prev_year(code, period, q_no)
        if not prev:
            continue
        reasons = []
        cs, ps = cur.get("revenue"), prev.get("revenue")
        if cs and ps and ps > 0:
            ratio = cs / ps
            if ratio >= r["revenue_ratio_high"]:
                reasons.append(f"売上が前年同Qの{ratio:.1f}倍")
            elif ratio <= r["revenue_ratio_low"]:
                reasons.append(f"売上が前年同Qの{ratio:.2f}倍")
        cg, pg = cur.get("gross_profit"), prev.get("gross_profit")
        if cs and ps and cg is not None and pg is not None and cs > 0 and ps > 0:
            d = (cg / cs - pg / ps) * 100
            if abs(d) >= r["gross_margin_shift_pt"]:
                reasons.append(f"粗利率が前年同Qから{d:+.1f}pt")
        co, po = cur.get("operating_income"), prev.get("operating_income")
        if co is not None and po is not None and (po < 0 <= co or co < 0 <= po):
            reasons.append(f"営業損益が{po:,.0f}→{co:,.0f}で符号反転")
        if reasons:
            nm = con.execute("SELECT name FROM companies WHERE code=?",
                             (code,)).fetchone()
            out.append({"code": code, "name": (nm["name"] if nm else "?"),
                        "period": period, "q_no": q_no,
                        "reasons": " / ".join(reasons)})
    return out


def write_queue(con, rows: list[dict]) -> str:
    """キューを tasks/ に**追記**する。tasks/ は追記専用（フックが上書きを拒む）。"""
    path = os.path.join("tasks", "pl_adjustment_queue.md")
    body = [f"\n## 検出 {C.utcnow()} — {len(rows)} 件\n",
            "| コード | 社名 | 期 | Q | 検出理由 | 確認 |",
            "|---|---|---|---:|---|---|"]
    for x in rows:
        body.append(f"| {x['code']} | {x['name'][:14]} | {x['period']} | "
                    f"{x['q_no']} | {x['reasons']} | ☐ |")
    body.append("")
    with open(path, "a", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(body))
    return path


def record(con, *, code, period, q_no, item_key, amount, filing_id,
           locator, note, by) -> None:
    """調整を登録する。根拠が欠けていればDBの CHECK 制約が弾く。

    アプリ層でも先に確かめるのは、SQLite のエラーメッセージが
    「どのCHECKに引っかかったか」を教えてくれないため。
    """
    if item_key not in ITEM_KEYS:
        raise SystemExit(f"ERROR: item_key は {ITEM_KEYS} のいずれか: {item_key}")
    if amount <= 0:
        raise SystemExit("ERROR: amount は正値（控除する額）で入れる")
    if len((note or "").strip()) < 20:
        raise SystemExit(
            "ERROR: --note は注記からの引用を20文字以上。\n"
            "  金額だけを根拠なく登録できない構造にしてある（設計 A-2）。\n"
            "  該当書類の注記を開いて、一過性である旨が書かれた箇所を引用すること。")
    if not (locator or "").strip():
        raise SystemExit("ERROR: --locator に注記の所在を入れる（例:「(セグメント情報等) 3.」）")
    if not (by or "").strip():
        raise SystemExit("ERROR: --by に記録者を入れる")
    f = con.execute("SELECT id, code, date, title FROM filings WHERE id=?",
                    (filing_id,)).fetchone()
    if f is None:
        raise SystemExit(f"ERROR: filing_id={filing_id} が filings に無い")
    if f["code"] != code:
        raise SystemExit(f"ERROR: filing_id={filing_id} は {f['code']} の書類。"
                         f"{code} の調整の根拠にはできない")
    con.execute(
        "INSERT OR REPLACE INTO pl_adjustments (code, period, q_no, item_key, "
        " amount, source_note, source_filing_id, source_locator, confirmed_by, "
        " confirmed_at, note) VALUES (?,?,?,?,?,?,?,?,?,?,NULL)",
        (code, period, q_no, item_key, float(amount), note.strip(), filing_id,
         locator.strip(), by.strip(), C.utcnow()))
    con.commit()
    C.log(f"記録: {code} {period} Q{q_no} {item_key} {amount:,.0f} 円")
    C.log(f"  根拠: {f['date']} {str(f['title'])[:50]} / {locator}")


def report(con) -> None:
    rows = con.execute(
        "SELECT code, period, q_no, item_key, amount, source_locator, confirmed_by "
        "FROM pl_adjustments ORDER BY code, period, q_no").fetchall()
    C.log(f"登録済みの調整: {len(rows)} 件")
    for r in rows:
        C.log(f"  {r['code']} {r['period']} Q{r['q_no']} {r['item_key']:<17} "
              f"{r['amount']:>16,.0f} 円  [{r['source_locator'][:30]}] "
              f"by {r['confirmed_by']}")


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="一時収入の検出キューと記録")
    p.add_argument("--scan", action="store_true", help="異常な期を検出しキューへ追記")
    p.add_argument("--queue", action="store_true", help="登録済みの調整を一覧")
    p.add_argument("--record", action="store_true", help="調整を登録")
    for k in ("code", "period", "item", "locator", "note", "by"):
        p.add_argument(f"--{k}")
    p.add_argument("--q", type=int)
    p.add_argument("--amount", type=float)
    p.add_argument("--filing-id", type=int)
    a = p.parse_args(argv)

    con = C.init_db()
    if a.scan:
        rows = scan(con)
        C.log(f"検出: {len(rows)} 件")
        for x in rows[:30]:
            C.log(f"  {x['code']} {x['name'][:14]:<14} {x['period']} Q{x['q_no']}  {x['reasons']}")
        if rows:
            C.log(f"キューへ追記: {write_queue(con, rows)}")
        return 0
    if a.queue:
        report(con)
        return 0
    if a.record:
        record(con, code=a.code, period=a.period, q_no=a.q, item_key=a.item,
               amount=a.amount, filing_id=a.filing_id, locator=a.locator,
               note=a.note, by=a.by)
        return 0
    p.error("--scan / --queue / --record のいずれかを指定する")


if __name__ == "__main__":
    raise SystemExit(main())
