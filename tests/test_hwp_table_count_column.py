"""The sample-count column is sized to its text with no side margins; other columns share the rest."""

from report_automation_engine.hwp_style_config import DEFAULT_HWP_STYLE_CONFIG
from report_automation_engine.hwp_table_layout import analyze_table_columns, plan_table_parts


def cell(row, col, text, role="value", **extra):
    return {"row": row, "col": col, "display_text": text, "role": role, "rowspan": 1, "colspan": 1, "covered_by": "", **extra}


def table():
    header = [cell(1, 1, "구분", "header", colspan=2), cell(1, 2, "", "blank", covered_by="R1C1"),
              cell(1, 3, "사례수", "banner_horizontal"), cell(1, 4, "지면 내 기사 비율 평균(%)", "banner_horizontal"),
              cell(1, 5, "지면 내 광고 비율 평균(%)", "banner_horizontal")]
    rows = [header, [cell(2, 1, "전체", "stub", colspan=2), cell(2, 2, "", "blank", covered_by="R2C1"),
                     cell(2, 3, "(1,449)"), cell(2, 4, "70.90"), cell(2, 5, "29.10")]]
    for r, name in enumerate(["일간신문", "주간신문"], 3):
        rows.append([cell(r, 1, "매체 유형", "banner_vertical"), cell(r, 2, name, "stub"),
                     cell(r, 3, "(225)"), cell(r, 4, "69.14"), cell(r, 5, "30.86")])
    return {"table_key": "T1", "matrix": rows, "row_count": len(rows), "col_count": 5}


def test_count_column_has_no_side_margin_and_keeps_its_measured_width():
    analysis = analyze_table_columns(table(), 8, DEFAULT_HWP_STYLE_CONFIG)
    assert analysis["tight_columns"] == {2}
    assert analysis["horizontal_margin"](table()["matrix"][1][2], 1, 2) == 0
    part = plan_table_parts(table(), 48000, 8, DEFAULT_HWP_STYLE_CONFIG)[0]
    widths = part["column_widths_hwpunit"]
    assert part["tight_columns"] == [2]
    assert widths[2] == analysis["widths"][2]          # not stretched to fill the body
    assert widths[2] < widths[3] and sum(widths) == 48000 - 500
    # "(1,449)" is the longest entry: 7 narrow characters, wider than the 3-syllable header.
    assert widths[2] >= 7 * 0.55 * 8 * 100
