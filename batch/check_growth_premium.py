"""Apply the 追補4 §Q trigger mechanically to generated models.

§Q demotes the Exit leg to [参考] and makes Target = PGM alone when ALL of:
    WARN 11 divergence      > 3.0x
    PGM-implied exit        < 10x
    assumed exit (peer med) > 25x

The three numbers all appear in validate_output's check 11 line, so this reads
the *_validation.txt written next to each workbook. Reports which tickers qualify;
it does not modify anything (batch/demote_exit_leg.py does that).

Usage: python batch/check_growth_premium.py [<validation.txt> ...]
"""
import sys, os, re, glob
sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PAT = re.compile(r"PGM implies ([\d.]+)x vs assumed ([\d.]+)x \(([\d.]+)x")


def check(path):
    code = os.path.basename(path)[:4]
    txt = open(path, encoding="utf-8", errors="replace").read()
    m = PAT.search(txt)
    if not m:
        return code, None
    pgm, assumed, div = (float(x) for x in m.groups())
    hit = div > 3.0 and pgm < 10.0 and assumed > 25.0
    return code, (pgm, assumed, div, hit)


if __name__ == "__main__":
    paths = sys.argv[1:] or sorted(glob.glob(os.path.join(ROOT, "models", "*_DCF_Model_20260905_validation.txt")))
    qual = []
    for p in paths:
        code, r = check(p)
        if r is None:
            continue
        pgm, assumed, div, hit = r
        mark = "*** §Q 該当 ***" if hit else "                "
        print(f"{mark} {code}: PGM逆算 {pgm:6.2f}x (<10) / 仮定 {assumed:6.2f}x (>25) / 乖離 {div:5.2f}x (>3.0)")
        if hit:
            qual.append(code)
    print("\n§Q 該当: %s" % (", ".join(qual) if qual else "なし"))
