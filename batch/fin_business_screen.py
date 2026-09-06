"""Screen a cached EDINET XBRL for consolidated financial-business line items.

Catches the 9433 KDDI case: a group that consolidates a bank / credit business
shows dedicated balance-sheet elements the ordinary type-A net_debt definition
must not absorb. Run after the pipeline has downloaded a ticker's XBRL.

Also routes the finding to a 銘柄型 (手順書 §2): see classify() for the boundary
and how it was calibrated.

Usage: python batch/fin_business_screen.py <docID> [...]
       python batch/fin_business_screen.py --latest    (scan every cached doc)
"""
import sys, os, glob, re
sys.stdout.reconfigure(encoding="utf-8")
from bs4 import BeautifulSoup

ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "tmp", "edinet_data")
PAT = re.compile(r"(ForFinancialBusiness|BankingBusiness|CallLoan|CallMoney|"
                 r"DepositsFromCustomers|LoansAndBillsDiscounted|InsuranceContract|"
                 r"PolicyReserve|InstallmentReceivable|LeaseReceivable|"
                 r"CreditCard|AccountsReceivableInstallment|OperatingLoans|"
                 # captive finance の債権。6326 クボタは FinancialReceivables で
                 # 総資産の 35.8%(2,221,256mn)を占めるのに、この語が無かったため
                 # スクリーンが CLEAN を返していた。
                 r"FinancialReceivables|FinanceReceivables|SalesFinanceReceivables)")

# IFRS タグは側面を名前に持つ: ...AssetsIFRS / ...CAIFRS(流動資産) /
# ...NCAIFRS(非流動資産) が資産、...LiabilitiesIFRS / ...CLIFRS(流動負債) /
# ...NCLIFRS(非流動負債) が負債。日本基準タグには標識が無いので語彙で補う。
ASSET_TAG = re.compile(r"(Assets(IFRS)?$|NCAIFRS$|CAIFRS$|"
                       r"LoansAndBillsDiscounted|CallLoan|InstallmentReceivable|"
                       r"LeaseReceivable)")
LIAB_TAG = re.compile(r"(Liabilities(IFRS)?$|NCLIFRS$|CLIFRS$|"
                      r"Deposits(From|InBanking)|CallMoney|PolicyReserve)")

def scan(doc, return_hits=False):
    files = glob.glob(os.path.join(ROOT, doc, "XBRL", "PublicDoc", "*.xbrl"))
    if not files:
        print(f"{doc}: no PublicDoc xbrl cached")
        return {} if return_hits else None
    soup = BeautifulSoup(open(files[0], encoding="utf-8").read(), "xml")
    hits, total = {}, None
    for el in soup.find_all():
        n = el.name.split(":")[-1]
        ctx = el.get("contextRef", "")
        if "CurrentYearInstant" not in ctx or "Member" in ctx:
            continue
        try:
            v = float(el.text)
        except (TypeError, ValueError):
            continue
        if n in ("AssetsIFRS", "Assets"):
            total = v
        if PAT.search(n):
            hits[n] = v
    if not hits:
        print(f"{doc}: CLEAN - no financial-business balance-sheet elements")
        return {} if return_hits else None
    print(f"{doc}: *** FINANCIAL-BUSINESS ELEMENTS FOUND ***"
          f"{'' if not total else f'  (total assets {total/1e6:,.0f} mn)'}")
    for n, v in sorted(hits.items(), key=lambda kv: -kv[1]):
        share = f"  = {v/total:.1%} of assets" if total else ""
        print(f"   {n:<58}{v/1e6:>15,.0f} mn{share}")
    mn = {n: round(v / 1e6) for n, v in hits.items()}
    route = classify(hits, total)
    if route is None:
        print("   [型ルーティング] 総資産が読めず比率判定できません")
    else:
        if route.get("unknown"):
            print(f"   [型ルーティング] 資産/負債を判定できなかったタグ: "
                  f"{', '.join(route['unknown'])} — 比率から除外した")
        if route["type"] is None:
            print(f"   [型ルーティング] {route['reason']}"
                  f"(負債側 {route['liab_share']:.1%} of assets)")
        else:
            print(f"   [型ルーティング] 推奨: 型{route['type']}  "
                  f"(金融資産 {route['share']:.1%} / 金融負債 {route['liab_share']:.1%} "
                  f"of assets) — {route['reason']}")
            print(f"   overrides に company_type: \"{route['type']}\" を宣言すること"
                  f"(宣言が無いと型A として扱われる)")
    if return_hits:
        return mn

def _side(name):
    """タグ名から資産側 / 負債側を判定する。判定できないものは None。"""
    if LIAB_TAG.search(name):
        return "L"
    if ASSET_TAG.search(name):
        return "A"
    return None


def classify(hits, total):
    """金融資産の総資産比から 型D / 型E を推奨する。

    使うのは【資産側】の比率だけである。預金は金融事業の規模を示す指標ではあるが
    負債であり、資産比の分子に足すと同じ事業を二度数えることになる(最初の実装は
    9433 の預金を資産として数え、67.2% という実態の倍近い比率を出していた)。

    境界は資産側 50%:

        8410 セブン銀行  ほぼ100%  → 型D (銀行そのもの)
        9433 KDDI        37.5%     → 型E (貸出金+有価証券。非金融が主体)
        4689 LINEヤフー   28.7%     → 型E (同上。PayPay銀行の貸出金・有価証券)

    50% を超えると連結の資産の過半が金融になり、非金融を切り出しても残りが会社を
    代表しない —— DDM/RI(型D)へ送る。下回る場合は非金融が主体なので、金融を
    分離して SOTP(型E)で解ける。

    返り値は【推奨】であって決定ではない。25%〜50% の帯は「金融は無視できないが
    主体でもない」領域で、セグメント利益の構成を見て人が確定する。事業多角化や
    政策保有株が理由で分離が要る銘柄(6971 京セラ)は、そもそもこの検知に載らない。
    """
    if not hits or not total:
        return None
    assets = {n: v for n, v in hits.items() if _side(n) == "A"}
    liabs = {n: v for n, v in hits.items() if _side(n) == "L"}
    unknown = [n for n in hits if _side(n) is None]
    if not assets:
        return {"type": None, "share": None, "liab_share": sum(liabs.values()) / total,
                "unknown": unknown,
                "reason": "負債側の金融科目のみ検出。資産側が無いため比率判定は保留"}
    share = sum(assets.values()) / total
    if share >= 0.50:
        t, why = "D", "金融資産が総資産の過半。連結DCF は成立せず DDM + Residual Income"
    else:
        t, why = "E", "非金融が主体。金融を net_debt / NWC から分離し 非金融DCF + 金融PBR の SOTP"
    if 0.25 <= share < 0.50:
        why += "(25〜50% の帯 — セグメント利益構成で確認すること)"
    return {"type": t, "share": share, "liab_share": sum(liabs.values()) / total,
            "unknown": unknown, "reason": why}


if __name__ == "__main__":
    args = sys.argv[1:]
    docs = ([os.path.basename(p) for p in glob.glob(os.path.join(ROOT, "S*"))]
            if args == ["--latest"] else args)
    for d in docs:
        scan(d)
