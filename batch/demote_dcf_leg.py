"""Demote ONE of the two DCF legs to a reference row so Target = the other leg alone.

Two distinct triggers use this, and they demote opposite legs:

  追補4 §Q (leg=exit)  growth-premium names. The perpetuity leg implies a 5-7x exit
      multiple (g fixed at 1.0%) while the peer median sits above 25x. Demote Exit,
      keep PGM as a conservative floor.
      Trigger (all three): divergence > 3.0x AND PGM-implied < 10x AND assumed > 25x.

  追補5 §U (leg=pgm)   denominator-trough names, AFTER the 型B mid-cycle
      re-generation has been applied and the divergence still exceeds 3.0x.
      Here the PGM leg is the distorted one: for a capital-intensive company
      (capex/D&A persistently > 1.5x) the perpetuity leg implies a terminal
      EV/EBITDA far BELOW anything observable in the market, so it fails an
      empirical sanity test as a terminal value. Demote PGM, keep Exit.
      §U requires the choice of authoritative leg to be recorded in the
      Adjustments Log; this script only writes the workbook-side evidence.

In both cases the point is that 追補5 forbids the naive midpoint: a Target that is
the average of two irreconcilable worldviews belongs to neither.

What this does, following 手順書§5-5 (never write a label starting with '='):
  1. the demoted leg's label -> marked [参考・Target不算入 — <rule>]
  2. C10 target -> =IF(ISNUMBER(Cn),ROUND(Cn,0),"N/A") pointing at the SURVIVING
     leg. Still a formula, still free of the comps rows, so validate check 14
     stays out of FAIL and reports WARN "does not average C16:C17" — the intended
     signal that a leg was deliberately dropped.
  3. B20 note -> the ticker-specific reason is appended.

Generic: every number and the reason text come from the command line.
Usage:
  python batch/demote_dcf_leg.py <xlsx> --leg exit|pgm --div D --pgm-implied P \
      --assumed A [--reason "extra sentence"]
"""
import argparse, sys, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

RULE = {"exit": "追補4 §Q", "pgm": "追補5 §U", "x": "追補6 §X"}
PREFIX = {"exit": "DCF - Exit Multiple", "pgm": "DCF - Perpetuity Growth"}


def find(ws, prefix, col=2, limit=40):
    for r in range(1, min(ws.max_row, limit) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith(prefix):
            return r
    return None


def main(a):
    wb = openpyxl.load_workbook(a.xlsx)
    ws = wb["Executive Summary"]
    r_dem = find(ws, PREFIX[a.leg])
    r_keep = find(ws, PREFIX["pgm" if a.leg == "exit" else "exit"])
    r_note = find(ws, "Note: Target Mid")
    r_tgt = find(ws, "Target Price")
    if not (r_dem and r_keep and r_note and r_tgt):
        raise SystemExit("ERROR: Exec Summary rows not found — refusing to guess row numbers.")

    rule = RULE["x"] if a.rule == "x" else RULE[a.leg]
    tag = "[参考・Target不算入 — %s]" % rule
    label = ws.cell(r_dem, 2).value
    if tag not in label:
        ws.cell(r_dem, 2).value = label + " " + tag
    assert not str(ws.cell(r_dem, 2).value).startswith("="), "label must not start with '='"

    # each methodology row carries its implied value in column C of the SAME row
    # (B16/C16 = PGM, B17/C17 = Exit), so the surviving leg's value cell is C<r_keep>.
    c_keep = r_keep
    ws.cell(r_tgt, 3).value = '=IF(ISNUMBER(C%d),ROUND(C%d,0),"N/A")' % (c_keep, c_keep)

    if a.rule == "x":
        kept = "Exit" if a.leg == "pgm" else "PGM"
        sentence = (
            "【追補6 §X】PGMとExitの乖離が **%s倍** で 3.0倍を超えたため、"
            "**乖離そのものを裁定トリガーとする追補6 §X により中点平均を禁止**し、どちらの脚を落とすかを判定した。"
            "■**観測倍率バンドテスト**: Peer の実測 EV/EBITDA の第1〜第3四分位バンドは **%s倍**。"
            "これに対し **PGM逆算 %s倍 / 仮定Exit %s倍**。"
            "**%s法のみがバンドの外にあるため %s法を[参考]に降格し、Target = %s 単独**とした。"
            "■バンドテストは市場倍率を無条件のアンカーにするものではない（それではスクリーナーが市場の複写に堕する）。"
            "**実際に取引されているレンジの外に出た脚だけを取り除く**という位置づけである。"
            % (a.div, a.band or "—", a.pgm_implied, a.assumed,
               "PGM" if a.leg == "pgm" else "Exit",
               "PGM" if a.leg == "pgm" else "Exit", kept)
        )
    elif a.leg == "exit":
        sentence = (
            "【追補4 §Q】Exit法を[参考]に降格し Target = PGM 単独とした。"
            "PGM逆算Exit倍率 %s倍 に対し仮定Exit倍率(Peer中央値) %s倍 で乖離 %s倍。"
            "永久成長率1.0%%固定の下ではPGM逆算Exitは5〜7倍にしかならず、"
            "Peer中央値が織り込む長期成長と構造的に両立しないため、両者の平均は"
            "『相容れない2つの世界観の中点』にすぎない。"
            "PGM単独Targetは保守フロアであり、現値との差は市場が織り込む成長プレミアムとして読むこと。"
            % (a.pgm_implied, a.assumed, a.div)
        )
    else:
        sentence = (
            "【追補5 §U】型Bミッドサイクル正常化を適用して再生成した後もPGMとExitの乖離が %s倍 残ったため、"
            "§U の手順に従いどちらを正とするかを個別判断し、**PGM法を[参考]に降格して Target = Exit 単独**とした。"
            "PGM逆算のターミナルEV/EBITDAは %s倍 で、上場する事業会社が実際に取引される倍率の下限を大きく下回り、"
            "ターミナル価値として経験的に成立しない。一方の仮定Exit倍率 %s倍(Peer中央値)は市場で観測される水準の内側にある。"
            "片方の脚が観測可能な範囲の外に出た場合は、内側にある脚を正とする。"
            "乖離の原因は本銘柄固有の洞察ではなく、永久成長率1.0%%固定の単段階モデルが"
            "資本集約型(capex/D&A が持続的に1.5倍超)を評価しきれないという既知のモデル上の限界である"
            "(バッチ後の2段階成長モデル導入課題として登録済み)。"
            % (a.div, a.pgm_implied, a.assumed)
        )
    if a.reason:
        sentence += a.reason
    cur = ws.cell(r_note, 2).value
    if rule not in cur:
        ws.cell(r_note, 2).value = cur + " ■" + sentence
    wb.save(a.xlsx)
    print("demoted %s leg in %s" % (a.leg.upper(), a.xlsx))
    print("  B%d = %s" % (r_dem, ws.cell(r_dem, 2).value))
    print("  C%d = %s" % (r_tgt, ws.cell(r_tgt, 3).value))
    print("NOTE: cached values dropped — run recalc_excel_com.py then validate_output.py.")


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("xlsx")
    p.add_argument("--leg", required=True, choices=["exit", "pgm"])
    p.add_argument("--div", required=True)
    p.add_argument("--pgm-implied", required=True)
    p.add_argument("--assumed", required=True)
    p.add_argument("--reason", default="")
    p.add_argument("--rule", choices=["auto", "x"], default="auto",
                   help="auto = the legacy §Q/§U tag; x = 追補6 §X (the unified arbitration)")
    p.add_argument("--band", default="", help="peer EV/EBITDA band 'p25-p75' for the §X sentence")
    main(p.parse_args())
