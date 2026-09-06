"""Demote the Exit-multiple leg to a reference row so Target = PGM alone (追補4 §Q).

Applies to the growth-premium names where the two DCF legs describe incompatible
worlds: the perpetuity leg implies a 5-7x exit multiple (because g is fixed at
1.0%) while the peer median sits above 25x. Averaging them produces a Target that
is only the midpoint of two irreconcilable views, so the batch spec demotes Exit
to [参考] and keeps PGM as a conservative floor.

Trigger (all three must hold; the caller checks them and passes the numbers in):
  * WARN 11 divergence   > 3.0x
  * PGM-implied exit     < 10x
  * assumed exit (peer)  > 25x

What this does, following 手順書§5-5 (never write a label starting with '='):
  1. B17 label  -> marked [参考・Target不算入 — 追補4 §Q]
  2. C10 target -> =IF(ISNUMBER(C16),ROUND(C16,0),"N/A")   (PGM alone; still a
     formula, still free of the comps rows, so validate check 14 stays out of FAIL
     and reports WARN "does not average C16:C17" — which is the intended signal)
  3. B20 note   -> the ticker-specific reason is appended

Generic: every number comes from the command line.
Usage:
  python batch/demote_exit_leg.py <xlsx> <divergence> <pgm_implied> <assumed>
"""
import sys, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

TAG = "[参考・Target不算入 — 追補4 §Q]"


def find(ws, prefix, col=2, limit=40):
    for r in range(1, min(ws.max_row, limit) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith(prefix):
            return r
    return None


def main(path, div, pgm_implied, assumed):
    wb = openpyxl.load_workbook(path)
    ws = wb["Executive Summary"]
    r_exit = find(ws, "DCF - Exit Multiple")
    r_note = find(ws, "Note: Target Mid")
    r_tgt = find(ws, "Target Price")
    if not (r_exit and r_note and r_tgt):
        raise SystemExit("ERROR: Exec Summary rows not found — refusing to guess row numbers.")

    label = ws.cell(r_exit, 2).value
    if TAG not in label:
        ws.cell(r_exit, 2).value = label + " " + TAG
    assert not str(ws.cell(r_exit, 2).value).startswith("="), "label must not start with '='"

    ws.cell(r_tgt, 3).value = '=IF(ISNUMBER(C%d),ROUND(C%d,0),"N/A")' % (r_exit - 1, r_exit - 1)

    sentence = (
        "【追補4 §Q】Exit法を[参考]に降格し Target = PGM 単独とした。"
        "PGM逆算Exit倍率 %s倍 に対し仮定Exit倍率(Peer中央値) %s倍 で乖離 %s倍。"
        "永久成長率1.0%%固定の下ではPGM逆算Exitは5〜7倍にしかならず、"
        "Peer中央値が織り込む長期成長と構造的に両立しないため、両者の平均は"
        "『相容れない2つの世界観の中点』にすぎない。"
        "PGM単独Targetは保守フロアであり、現値との差は市場が織り込む成長プレミアムとして読むこと。"
        % (pgm_implied, assumed, div)
    )
    cur = ws.cell(r_note, 2).value
    if "追補4 §Q" not in cur:
        ws.cell(r_note, 2).value = cur + " ■" + sentence
    wb.save(path)
    print("demoted Exit leg in %s" % path)
    print("  B%d = %s" % (r_exit, ws.cell(r_exit, 2).value))
    print("  C%d = %s" % (r_tgt, ws.cell(r_tgt, 3).value))
    print("NOTE: cached values dropped — run recalc_excel_com.py then validate_output.py.")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
