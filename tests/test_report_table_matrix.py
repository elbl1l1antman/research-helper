from report_automation_engine.report_table_matrix import build_table_matrix, format_display_value


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
