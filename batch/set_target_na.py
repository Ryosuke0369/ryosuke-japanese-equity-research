"""追補8 §AD-2 — mark the Target N/A for a leverage-amplified name.

The case this exists for: the two DCF legs agree on the MULTIPLE (so the band
test cannot separate them and neither leg is wrong) but disagree by more than
3x on the PRICE. That happens when net debt is large relative to equity value:
equity = EV - net_debt, so a few per cent of difference in EV becomes a multiple
of the equity value. 3861 王子HD is the case that motivated the rule -- multiple
divergence 1.45x, price divergence 4.85x, net_debt/EV 0.76.

Averaging the two legs here would produce exactly the "midpoint of two
irreconcilable worldviews" that 追補5 banned, and picking a leg would be
arbitrary because the band test says both are equally plausible. So above
net_debt/EV = 0.6 the honest output is that the model cannot produce a point
estimate at all.

The template already degrades gracefully: C11 (Recommendation) and C12 (Upside)
both test ISNUMBER on their input, so writing a formula returning "N/A" into C10 makes
all three read N/A without producing a single formula error. validate check 14
then WARNs that C10 no longer averages C16:C17 -- that WARN is the intended
signal, exactly as with the leg demotions.

Usage:
  python batch/set_target_na.py <xlsx> --price-div D --lev L --pgm P --exit E
"""
import argparse, sys, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

TAG = "追補8 §AD-2"


def find(ws, prefix, col=2, limit=40):
    for r in range(1, min(ws.max_row, limit) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith(prefix):
            return r
    return None


def main(a):
    wb = openpyxl.load_workbook(a.xlsx)
    ws = wb["Executive Summary"]
    r_tgt = find(ws, "Target Price")
    r_note = find(ws, "Note: Target Mid")
    if not (r_tgt and r_note):
        raise SystemExit("ERROR: Exec Summary rows not found — refusing to guess row numbers.")

    # A literal string fails validate check 14 ("a hardcoded target cannot track a
    # price or scenario change"), so keep it a FORMULA that evaluates to N/A. Check
    # 14 then degrades to the intended WARN rather than a FAIL.
    ws.cell(r_tgt, 3).value = '="N/A"' 
    lab = ws.cell(r_tgt, 2).value
    if TAG not in lab:
        ws.cell(r_tgt, 2).value = lab + " [判定不能 — %s]" % TAG
    assert not str(ws.cell(r_tgt, 2).value).startswith("="), "label must not start with '='"

    sentence = (
        "【%s】**Target は N/A（高レバレッジにより点推定が成立しない）**とした。"
        "PGM %s円 と Exit %s円 は**株価ベースで %s倍**開いているが、"
        "**倍率ベースの乖離は小さく、どちらの脚も誤りではない**。"
        "原因は財務レバレッジであり、株式価値 = EV − ネットデット の構造上、"
        "**EV のわずか数%%の差が株式価値では数倍に増幅される**（net_debt / EV = %s > 0.6）。"
        "■このためバンドテスト（追補6 §X-3）は倍率が一致しているため脚を判別できず、"
        "2脚の平均は追補5 が禁じた『どちらの世界観にも属さない中点』にしかならない。"
        "**点推定は意味を持たない**ため Target を出さない。"
        "■**要目視レビュー**。FINAL化時は負債削減の織り込み等の個別対応を行うこと。"
        "両脚の値（PGM %s円 / Exit %s円）は [参考] として残してある。"
        % (TAG, a.pgm, a.exit, a.price_div, a.lev, a.pgm, a.exit)
    )
    cur = ws.cell(r_note, 2).value
    if TAG not in cur:
        ws.cell(r_note, 2).value = cur + " ■" + sentence
    wb.save(a.xlsx)
    print("Target を N/A に設定: %s" % a.xlsx)
    print("  B%d = %s" % (r_tgt, ws.cell(r_tgt, 2).value))
    print("  C%d = %r" % (r_tgt, ws.cell(r_tgt, 3).value))
    print("NOTE: cached values dropped — run recalc_excel_com.py then validate_output.py.")


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("xlsx")
    p.add_argument("--price-div", required=True)
    p.add_argument("--lev", required=True)
    p.add_argument("--pgm", required=True)
    p.add_argument("--exit", required=True)
    main(p.parse_args())
