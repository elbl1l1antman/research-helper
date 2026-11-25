"""Compare every value cell: source Excel -> package -> generated HWPX.

1) package value cell (display_text) vs the original Excel cell (source_cell on source_sheet):
   the number shown must equal the Excel value at the shown precision (percent formats x100).
2) the HWPX tables (in writing order) must contain each table's formatted value texts.
Usage: numcheck.py <source.xlsx> <report_package.json> <writer_report.json> <output.hwpx> <style_config.json>
"""
import json, re, sys, zipfile
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from xml.etree import ElementTree as ET

sys.path.insert(0, str(__import__("pathlib").Path(__file__).resolve().parents[2]))
import openpyxl
from report_automation_engine.hwp_style_config import TABLE_OUTPUT_DEFAULTS
from report_automation_engine.hwp_table_output import format_table_display

src, package_path, writer_path, hwpx_path, style_path = sys.argv[1:6]
wb = openpyxl.load_workbook(src, data_only=True)
package = json.load(open(package_path, encoding="utf-8"))
writer = json.load(open(writer_path, encoding="utf-8"))
style = json.load(open(style_path, encoding="utf-8")) if style_path != "-" else {}
options = {**TABLE_OUTPUT_DEFAULTS, **style.get("table_output", {})}
NUM = re.compile(r"[+-]?[0-9][0-9,]*(?:\.[0-9]+)?")

value_issues, checked = [], 0
for table in package["tables"]:
    sheet = table.get("source_sheet")
    ws = wb[sheet] if sheet in wb.sheetnames else None
    for row in table.get("matrix") or []:
        for cell in row:
            if cell.get("role") != "value" or cell.get("covered_by") or not cell.get("source_cell") or ws is None:
                continue
            text = str(cell.get("display_text") or "").strip()
            match = NUM.search(text)
            raw = ws[cell["source_cell"]].value
            if not match or not isinstance(raw, (int, float)):
                continue
            checked += 1
            shown = Decimal(match.group(0).replace(",", ""))
            fmt = ws[cell["source_cell"]].number_format or ""
            value = Decimal(str(raw)) * (100 if "%" in fmt else 1)
            places = -shown.as_tuple().exponent if shown.as_tuple().exponent < 0 else 0
            if abs(value.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP) - shown) > Decimal(1).scaleb(-places) / 2:
                value_issues.append(f"{table['table_key']} {cell['source_cell']}: Excel {raw} ({fmt}) vs package '{text}'")

# HWPX tables in document order
z = zipfile.ZipFile(hwpx_path)
names = sorted([n for n in z.namelist() if re.match(r"Contents/section\d+\.xml", n)], key=lambda n: int(re.findall(r"\d+", n)[0]))
hwpx_tables = []
for name in names:
    root = ET.fromstring(z.read(name))
    for tbl in root.iter():
        if tbl.tag.endswith("}tbl"):
            texts = []
            for tc in tbl.iter():
                if tc.tag.endswith("}tc"):
                    texts.append("".join(t.text or "" for t in tc.iter() if t.tag.endswith("}t")).strip())
            if tbl.get('rowCnt') == '1' and tbl.get('colCnt') == '1' and not any(texts):
                continue
            hwpx_tables.append(texts)

layouts = {item["table_key"]: len(item["parts"]) for item in writer.get("table_layouts", [])}
order = [section["table_key"] for section in package["sections"] if section["table_key"] in layouts]
by_key = {t["table_key"]: t for t in package["tables"]}
cursor, missing, compared = 0, [], 0
for key in order:
    parts = layouts[key]
    got = Counter(text for part in hwpx_tables[cursor:cursor + parts] for text in part)
    cursor += parts
    shown = format_table_display(by_key[key], options) if by_key[key].get("matrix") else by_key[key]
    expected = Counter(str(c.get("display_text") or "").strip() for row in shown.get("matrix") or [] for c in row
                       if c.get("role") == "value" and not c.get("covered_by") and str(c.get("display_text") or "").strip())
    for text, count in expected.items():
        compared += count
        if got[text] < count:
            missing.append(f"{key}: '{text}' x{count - got[text]} not found in HWPX table")

print(json.dumps({"excel_vs_package_checked": checked, "excel_vs_package_issues": value_issues[:40],
                  "excel_vs_package_issue_count": len(value_issues), "hwpx_tables": len(hwpx_tables),
                  "tables_expected": sum(layouts.values()), "hwpx_values_compared": compared,
                  "hwpx_missing_count": len(missing), "hwpx_missing": missing[:40]}, ensure_ascii=False, indent=1))
