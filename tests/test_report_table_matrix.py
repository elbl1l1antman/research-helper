from report_automation_engine.report_table_matrix import build_table_matrix, build_table_matrix_from_cells, format_display_value


def test_format_display_value_uses_one_decimal_for_float():
    assert format_display_value(3.333535353, decimal_places=1) == "3.3"


def test_format_display_value_preserves_percent_unit():
    assert format_display_value(63.25, unit="%", decimal_places=1) == "63.3%"


def test_build_table_matrix_adds_display_cells_without_losing_rows():
    table = {
        "table_key": "T001",
        "title": "만족도",
        "rows": [
            {"category": "전체", "percent": 63.25, "weighted_n": 1200, "raw_n": 1198, "unit": "%", "source_cell": "D5"},
            {"category": "매우 만족", "percent": 29.44, "weighted_n": 353, "raw_n": 351, "unit": "%", "source_cell": "D6"},
        ],
    }
    matrix_table = build_table_matrix(table, decimal_places=1)
    assert matrix_table["row_count"] == 3
    assert matrix_table["col_count"] == 4
    assert matrix_table["matrix"][0][0]["role"] == "header"
    assert matrix_table["matrix"][1][1]["display_text"] == "63.3%"
    assert matrix_table["matrix"][1][1]["raw_value"] == 63.25
    assert matrix_table["matrix"][1][1]["source_cell"] == "D5"
    assert matrix_table["roles"]["value"] >= 1


def test_cell_contract_reports_overlapping_merges_and_hash_display():
    rows = [
        {"table_key": "T001", "row": 1, "col": 1, "rowspan": 1, "colspan": 2, "role": "banner_horizontal", "display_text": "구분", "source_range": "A1:B2"},
        {"table_key": "T001", "row": 1, "col": 2, "rowspan": 2, "colspan": 1, "role": "banner_horizontal", "display_text": "비율", "source_range": "A1:B2"},
        {"table_key": "T001", "row": 2, "col": 1, "role": "value", "display_text": "###", "raw_value": 123.4, "source_range": "A1:B2"},
    ]

    table = build_table_matrix_from_cells(rows)
    errors = [item["message"] for item in table["qa"] if item["severity"] == "error"]

    assert any("병합" in message for message in errors)
    assert any("###" in message for message in errors)
