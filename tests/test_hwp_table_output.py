from report_automation_engine.hwp_style_config import DEFAULT_HWP_STYLE_CONFIG, validate_hwp_style_config
from report_automation_engine.hwp_table_layout import prepare_report_table
from report_automation_engine.hwp_table_output import format_table_display


def cell(row, col, text, role="value", **extra):
    return {"row": row, "col": col, "display_text": text, "role": role, "rowspan": 1, "colspan": 1, **extra}


def test_numeric_options_leave_percent_raw_values_and_input_unchanged():
    headers = [cell(1, i, name, "banner_horizontal") for i, name in
               enumerate(["사례수", "평균(백만원)", "%", "평균(명)"], 1)]
    values = [cell(2, i, value, raw_value=3333333.333) for i, value in
              enumerate(["(3232)", "1234.55", "63.3", "-"], 1)]
    table = {"matrix": [headers, values], "col_count": 4}
    output = format_table_display(table, {"sample_grouping": True, "value_decimals": 1, "value_grouping": True})
    assert [c["display_text"] for c in output["matrix"][1]] == ["(3,232)", "1,234.6", "63.3", "-"]
    assert output["matrix"][1][1]["raw_value"] == 3333333.333
    assert values[1]["display_text"] == "1234.55"


def test_guard_column_preserves_data_merge_and_uses_existing_width_budget():
    anchor = {**cell(1, 1, "머리행", "banner_horizontal"), "rowspan": 2, "colspan": 2}
    table = {"matrix": [[anchor, cell(1, 2, "", covered_by="A1")],
                        [cell(2, 1, "", covered_by="A1"), cell(2, 2, "", covered_by="A1")],
                        [cell(3, 1, "전체", "stub"), cell(3, 2, "1")]],
             "column_widths_hwpunit": [10000, 10000], "body_width_hwpunit": 20500, "col_count": 2,
             "merged_ranges": [{"row": 1, "col": 1, "rowspan": 2, "colspan": 2}]}
    output = prepare_report_table(table)
    assert output["protected_merge_rows"] == [3]
    assert output["matrix"][2][2]["role"] == "merge_guard"
    assert output["merged_ranges"][1] == {"row": 2, "col": 1, "rowspan": 2, "colspan": 2}
    assert output["row_count"] == 4 and output["col_count"] == 3
    assert sum(output["column_widths_hwpunit"]) == 20000
    assert len(table["matrix"][0]) == 2


def test_caption_and_display_config_accepts_existing_presets():
    old = {key: value for key, value in DEFAULT_HWP_STYLE_CONFIG.items() if key != "table_output"}
    assert validate_hwp_style_config(old)["table_output"]["number_pattern"] == "<표 {number}>"
