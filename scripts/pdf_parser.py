"""
Kudan (4425) 決算短信PDF 財務データ抽出スクリプト (V3)

pdfplumberの空間座標ベースで抽出:
- Page 1: 連結経営成績テーブル (売上高・営業利益・当期純利益) in 百万円
- Inner pages: PL詳細 (COGS, SGA), BS (現金, 借入金), CF (営業CF, 投資CF) in 千円

3つのPDFからcurrent + previousの2年分ずつ抽出し、
期間ラベルベースで重複排除、FY2022〜FY2025の4年分を時系列配列に整理する。

V3 changes:
  - extract_inner_page_items(): PL/BS/CF全項目を内部ページからスキャン
  - extract_all_financials(): dict形式で全フィールドを返却（旧V2はtuple）
  - CF抽出: narrative page誤検出を防ぐため、tabular CF pageのみ対象
"""

import pdfplumber
import re
import os
import sys
import json


def parse_number(text):
    """Parse a number string, handling △ (negative) and commas."""
    s = text.replace(",", "").replace("，", "")
    if "△" in s:
        s = "-" + s.replace("△", "")
    s = re.sub(r"[^\d\-]", "", s)
    if s and s != "-":
        return int(s)
    return None


def extract_row_values(row_y, col_headers, words, y_tolerance=5):
    """Extract values from a data row by matching x-coordinates to column headers.

    Uses x-center distance matching (30px threshold) to align numbers with headers.
    """
    row_values = {}

    # Collect all number-like words on this row (within ±y_tolerance of row_y)
    row_words = []
    for w in words:
        if abs(w["top"] - row_y) < y_tolerance:
            text = w["text"]
            # Skip period labels and percentage values
            if re.search(r"\d{4}年", text):
                continue
            if "％" in text or "%" in text:
                continue
            if text in ("―", "─", "-", "–"):
                continue
            # Check if it looks like a number (with optional △, commas)
            if re.search(r"[△\d]", text):
                x_center = (w["x0"] + w["x1"]) / 2
                row_words.append((x_center, text))

    # Match each number to the nearest column header
    for key, col_x in col_headers.items():
        best_match = None
        best_dist = float("inf")
        for x_center, text in row_words:
            dist = abs(x_center - col_x)
            if dist < best_dist and dist < 30:  # within 30px
                best_dist = dist
                best_match = text
        if best_match:
            row_values[key] = parse_number(best_match)
        else:
            row_values[key] = None

    return row_values


def extract_row_numbers_by_x(label_y, words, y_tolerance=5):
    """Extract all numbers on a row, sorted by x-coordinate.

    Returns list of (x0, parsed_int) tuples, sorted left-to-right.
    Used for inner page extraction where there are no column headers —
    just [previous_year, current_year] sorted by x position.
    """
    numbers = []
    for w in words:
        if abs(w["top"] - label_y) < y_tolerance:
            text = w["text"]
            if re.search(r"[△\d]", text) and not re.search(r"\d{4}年", text):
                val = parse_number(text)
                if val is not None:
                    numbers.append((w["x0"], val))
    numbers.sort(key=lambda x: x[0])
    return numbers


def extract_inner_page_items(pdf_path):
    """Extract PL, BS, CF items from inner pages of a 決算短信 PDF.

    Inner pages report in 千円. Values are converted to JPY mn (÷1000, rounded).

    Returns:
        dict with 'current' and 'previous' dicts, each containing:
            - cogs: int (JPY mn) or None
            - sga: int (JPY mn) or None
            - cash: int (JPY mn) or None
            - debt_short: int (JPY mn) or None
            - ocf: int (JPY mn) or None
            - investing_cf: int (JPY mn) or None
        or None on failure
    """
    # Target labels and their dict keys
    # (search_substring, dict_key, page_type)
    targets = [
        ("売上原価", "cogs", "pl"),
        ("販売費", "sga", "pl"),
        ("現金及び預金", "cash", "bs"),
        ("短期借入金", "debt_short", "bs"),
    ]

    # CF targets need special handling: subtotal lines are indented (x0 > 70)
    cf_targets = [
        ("営業活動によるキャッシュ・フロー", "ocf"),
        ("投資活動によるキャッシュ・フロー", "investing_cf"),
    ]

    current = {}
    previous = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            if page_idx == 0:
                continue  # Skip page 1, handled by extract_financials()

            words = page.extract_words()
            page_text = " ".join(w["text"] for w in words)

            # --- PL items: find on pages with 売上原価 ---
            if "売上原価" in page_text:
                for search_str, key, ptype in targets:
                    if ptype != "pl":
                        continue
                    for w in words:
                        if search_str in w["text"]:
                            nums = extract_row_numbers_by_x(w["top"], words)
                            if len(nums) >= 2:
                                previous[key] = round(nums[0][1] / 1000)
                                current[key] = round(nums[1][1] / 1000)
                            elif len(nums) == 1:
                                current[key] = round(nums[0][1] / 1000)
                            break

            # --- BS items: find on pages with 貸借対照表 + 現金及び預金 ---
            if "貸借対照表" in page_text and "現金及び預金" in page_text:
                for search_str, key, ptype in targets:
                    if ptype != "bs":
                        continue
                    for w in words:
                        if search_str in w["text"]:
                            nums = extract_row_numbers_by_x(w["top"], words)
                            if len(nums) >= 2:
                                previous[key] = round(nums[0][1] / 1000)
                                current[key] = round(nums[1][1] / 1000)
                            elif len(nums) == 1:
                                current[key] = round(nums[0][1] / 1000)
                            break

            # --- CF items: find tabular CF (subtotal lines with numbers) ---
            # Only process pages that have all 3 CF section headers (tabular CF page)
            has_ocf_header = any("営業活動によるキャッシュ・フロー" in w["text"] for w in words)
            has_icf_header = any("投資活動によるキャッシュ・フロー" in w["text"] for w in words)
            has_fcf_header = any("財務活動によるキャッシュ・フロー" in w["text"] for w in words)
            if has_ocf_header and has_icf_header and has_fcf_header:
                for cf_search, cf_key in cf_targets:
                    if cf_key in current:
                        continue  # Already found
                    for w in words:
                        # Match subtotal line: indented (x0 > 70), label-only (no digits in text)
                        if (cf_search in w["text"] and w["x0"] > 70
                                and not re.search(r"\d", w["text"])):
                            nums = extract_row_numbers_by_x(w["top"], words)
                            if len(nums) >= 2:
                                previous[cf_key] = round(nums[0][1] / 1000)
                                current[cf_key] = round(nums[1][1] / 1000)
                            elif len(nums) == 1:
                                current[cf_key] = round(nums[0][1] / 1000)

    if not current and not previous:
        return None

    print(f"  [Inner Pages] current: {current}")
    print(f"  [Inner Pages] previous: {previous}")

    return {"current": current, "previous": previous}


def extract_financials(pdf_path):
    """
    決算短信PDFの1ページ目から連結経営成績テーブルを抽出し、
    内部ページからPL/BS/CF詳細を追加取得する。

    Returns:
        dict with 'current' and 'previous' rows, each containing:
            - period: str (e.g. "2023年３月期")
            - values: {revenue, operating_income, net_income, cogs, sga, cash, ...}
        or None on failure
    """
    basename = os.path.basename(pdf_path)
    print(f"\n--- Parsing: {basename} ---")

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()

    # --- Step 1: Locate column headers ---
    col_headers = {}  # key -> x-center coordinate
    header_y = None

    for w in words:
        if w["text"] == "売上高" and "revenue" not in col_headers:
            col_headers["revenue"] = (w["x0"] + w["x1"]) / 2
            header_y = w["top"]
        elif w["text"] == "営業利益" and "operating_income" not in col_headers:
            col_headers["operating_income"] = (w["x0"] + w["x1"]) / 2
            if header_y is None:
                header_y = w["top"]

    # Net income header: look for 当期純利益 or 四半期純利益
    for w in words:
        if ("当期純利益" in w["text"] or "四半期純利益" in w["text"]) and "net_income" not in col_headers:
            col_headers["net_income"] = (w["x0"] + w["x1"]) / 2

    if not col_headers or header_y is None:
        print(f"  ERROR: Could not find column headers")
        return None

    print(f"  Column headers found: {list(col_headers.keys())} at y~{header_y:.1f}")

    # --- Step 2: Find data rows via unit marker and period labels ---
    unit_y = None
    for w in words:
        if w["text"] == "百万円" and w["top"] > header_y:
            unit_y = w["top"]
            break

    if unit_y is None:
        unit_y = header_y + 15  # fallback

    # Find period labels (e.g. "2023年３月期", "2024年３月期第２四半期")
    data_rows = []
    for w in words:
        if w["top"] > unit_y and re.search(r"\d{4}年", w["text"]):
            data_rows.append((w["top"], w["text"]))

    data_rows.sort(key=lambda x: x[0])
    data_rows = data_rows[:2]

    if not data_rows:
        print(f"  ERROR: No data rows found")
        return None

    print(f"  Data rows: {[r[1] for r in data_rows]}")

    # --- Step 3: Extract numbers for each row ---
    rows_data = []
    for row_y, period_label in data_rows:
        values = extract_row_values(row_y, col_headers, words)
        rows_data.append({
            "period": period_label,
            "values": values,
        })
        print(f"  {period_label}: {values}")

    # --- Step 4: Extract inner page items and merge ---
    inner = extract_inner_page_items(pdf_path)
    if inner:
        for row_key, row_data in zip(["current", "previous"], rows_data):
            inner_vals = inner.get(row_key, {})
            if inner_vals:
                row_data["values"].update(inner_vals)

    return {
        "current": rows_data[0] if len(rows_data) > 0 else None,
        "previous": rows_data[1] if len(rows_data) > 1 else None,
    }


def get_fiscal_year(period_label):
    """Extract fiscal year integer from period label like '2023年３月期'."""
    m = re.search(r"(\d{4})", period_label)
    return int(m.group(1)) if m else 0


def is_annual(period_label):
    """Check if a period label is annual (not quarterly)."""
    return "四半期" not in period_label


def extract_all_financials(folder=None):
    """
    Extract historical financials from all 3 Kudan PDFs.

    Returns:
        dict: {
            "hist_revenue": [...],
            "hist_operating_income": [...],
            "hist_net_income": [...],
            "hist_cogs": [...],
            "hist_sga": [...],
            "hist_ocf": [...],
            "hist_capex": [...],       # abs(investing_cf)
            "hist_cash": [...],
            "hist_debt": [...],
            "latest_net_debt": int,    # debt[-1] - cash[-1]
        }
        Each list contains 4 ints for FY2022-FY2025 in JPY mn.
        or None on failure
    """
    if folder is None:
        folder = os.path.dirname(os.path.abspath(__file__))

    files = [
        ("FY22_23", "Kudan_2022&2023.pdf"),
        ("FY24_25", "Kudan_2024&2025.pdf"),
        ("Newest", "Kudan_Newest.pdf"),
    ]

    # Collect all fiscal year data points
    fiscal_data = {}  # period_label -> {revenue, operating_income, net_income, cogs, sga, ...}

    for label, filename in files:
        path = os.path.join(folder, filename)
        if not os.path.exists(path):
            print(f"File not found: {filename}")
            continue

        result = extract_financials(path)
        if result is None:
            continue

        for row_key in ["current", "previous"]:
            row = result.get(row_key)
            if row and row["values"]:
                period = row["period"]
                # Deduplicate: later PDFs overwrite earlier ones for same period
                if period in fiscal_data:
                    fiscal_data[period].update(row["values"])
                else:
                    fiscal_data[period] = dict(row["values"])

    # Sort periods chronologically
    sorted_periods = sorted(fiscal_data.keys(), key=get_fiscal_year)

    # Filter to annual periods only
    annual_periods = [p for p in sorted_periods if is_annual(p)]
    quarterly_periods = [p for p in sorted_periods if not is_annual(p)]

    print("\n" + "=" * 60)
    print("Extracted Financial Data (JPY mn)")
    print("=" * 60)

    if not annual_periods:
        print("  ERROR: No annual periods extracted")
        return None

    # All fields to extract
    fields = [
        ("revenue", "hist_revenue"),
        ("operating_income", "hist_operating_income"),
        ("net_income", "hist_net_income"),
        ("cogs", "hist_cogs"),
        ("sga", "hist_sga"),
        ("ocf", "hist_ocf"),
        ("investing_cf", "hist_investing_cf"),
        ("cash", "hist_cash"),
        ("debt_short", "hist_debt"),
    ]

    result_dict = {}
    for field_key, hist_key in fields:
        arr = []
        for p in annual_periods:
            arr.append(fiscal_data[p].get(field_key))
        result_dict[hist_key] = arr

    # Capex = abs(investing_cf)
    result_dict["hist_capex"] = [
        abs(v) if v is not None else None
        for v in result_dict["hist_investing_cf"]
    ]

    # Net debt = debt - cash (positive = net debt, negative = net cash)
    cash_vals = result_dict["hist_cash"]
    debt_vals = result_dict["hist_debt"]
    if cash_vals and debt_vals and cash_vals[-1] is not None:
        latest_debt = debt_vals[-1] if debt_vals[-1] is not None else 0
        result_dict["latest_net_debt"] = latest_debt - cash_vals[-1]
    else:
        result_dict["latest_net_debt"] = None

    # Print summary
    for p in annual_periods:
        d = fiscal_data[p]
        parts = [f"rev={d.get('revenue')}", f"cogs={d.get('cogs')}", f"sga={d.get('sga')}",
                 f"op={d.get('operating_income')}", f"ni={d.get('net_income')}",
                 f"ocf={d.get('ocf')}", f"icf={d.get('investing_cf')}",
                 f"cash={d.get('cash')}", f"debt={d.get('debt_short')}"]
        print(f"  {p}: {', '.join(parts)}")

    print()
    print("# --- Extracted Historical Financials (JPY mn) ---")
    for field_key, hist_key in fields:
        print(f'"{hist_key}": {result_dict[hist_key]},')
    print(f'"hist_capex": {result_dict["hist_capex"]},')
    print(f'"latest_net_debt": {result_dict["latest_net_debt"]},')

    if quarterly_periods:
        print("\n# --- Quarterly Data (JPY mn) ---")
        for p in quarterly_periods:
            d = fiscal_data[p]
            print(f"# {p}: rev={d.get('revenue')}, op={d.get('operating_income')}, ni={d.get('net_income')}")

    # --- Verification vs expected values ---
    expected = {
        "hist_revenue": [272, 333, 491, 518],
        "hist_operating_income": [-433, -599, -527, -801],
        "hist_net_income": [-2237, -414, -70, -802],
    }

    if len(annual_periods) >= 4:
        print("\n# --- Verification vs Expected Config ---")
        print("# (±1 differences expected: PDF truncates, config rounds from 千円)")
        for key, label in [("revenue", "hist_revenue"),
                           ("operating_income", "hist_operating_income"),
                           ("net_income", "hist_net_income")]:
            extracted = [fiscal_data[p].get(key) for p in annual_periods[:4]]
            exp = expected[label]
            diffs = [abs(a - b) for a, b in zip(extracted, exp)
                     if a is not None and b is not None]
            max_diff = max(diffs) if diffs else 0
            status = "OK" if max_diff <= 1 else f"DIFF (max={max_diff})"
            print(f"# {label}: extracted={extracted} expected={exp} [{status}]")

    # Save to JSON
    output_path = os.path.join(folder, "financials_spatial.json")
    json_data = {}
    for p in annual_periods:
        fy = get_fiscal_year(p)
        json_data[f"FY{fy}"] = fiscal_data[p]
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, indent=4, ensure_ascii=False)
    print(f"\nSaved extracted data to {output_path}")

    return result_dict


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    folder = os.path.dirname(os.path.abspath(__file__))
    result = extract_all_financials(folder)
    if result is None:
        print("\nERROR: Failed to extract financials")
        sys.exit(1)
    print(f"\nFinal result:")
    for key, val in result.items():
        print(f"  {key}: {val}")


if __name__ == "__main__":
    main()
