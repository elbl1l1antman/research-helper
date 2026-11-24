import copy
from pathlib import Path
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from report_automation_engine.hwp_style_config import margin_units
from report_automation_engine.hwp_table_layout import (
    apply_vertical_banner_policy,
    plan_table_parts,
    prepare_report_table,
)
from report_automation_engine.hwp_table_layout_options import (
    LAYOUT_DEFAULTS,
    dump_table_layout,
    layout_for_table,
    normalize_table_layout,
)


def cell(row, col, text, role, **extra):
    base = {"row": row, "col": col, "rowspan": 1, "colspan": 1, "covered_by": "",
            "source_cell": f"R{row}C{col}", "display_text": text, "role": role}
    base.update(extra)
    return base


def wide_table(cols=10, body_rows=3):
    """Header row + body rows; column 1 is the stub, the rest are values."""
    matrix = [[cell(1, 1, "구분", "header")] + [cell(1, c, f"보기{c}", "banner_horizontal") for c in range(2, cols + 1)]]
    for r in range(2, body_rows + 2):
        matrix.append([cell(r, 1, f"항목{r}", "stub")] + [cell(r, c, f"{c}.0", "value") for c in range(2, cols + 1)])
    return {"table_key": "T", "matrix": matrix, "row_count": len(matrix), "col_count": cols, "merged_ranges": []}


def merge(table, row, col, rowspan=1, colspan=1):
    matrix = table["matrix"]
    anchor = matrix[row - 1][col - 1]
    anchor.update(rowspan=rowspan, colspan=colspan)
    for r in range(row - 1, row - 1 + rowspan):
        for c in range(col - 1, col - 1 + colspan):
            if (r, c) != (row - 1, col - 1):
                matrix[r][c].update(covered_by=anchor["source_cell"], display_text="", role="blank")
    table["merged_ranges"].append({"row": row, "col": col, "rowspan": rowspan, "colspan": colspan})
    return table


def text_by_source(parts):
    seen = {}
    for part in parts:
        for row in part["matrix"]:
            for item in row:
                if not item.get("covered_by"):
                    seen.setdefault(item.get("source_cell"), set()).add(item["display_text"])
    return seen


# --- options file ---------------------------------------------------------

def test_layout_file_round_trip_and_table_override():
    config = normalize_table_layout({
        "schema_version": "1.0",
        "defaults": {"wide_table_policy": "landscape_if_needed"},
        "tables": {"Q1": {"cuts": [7, 4], "column_widths_mm": {"1": 30}, "body_merge_policy": "split_repeat"}},
    })
    assert config["tables"]["Q1"]["cuts"] == [4, 7]
    assert config["tables"]["Q1"]["column_widths_mm"] == {1: 30.0}
    merged = layout_for_table(config, "Q1")
    assert merged["wide_table_policy"] == "landscape_if_needed"
    assert merged["body_merge_policy"] == "split_repeat"
    assert layout_for_table(config, "other") == {**LAYOUT_DEFAULTS, "wide_table_policy": "landscape_if_needed"}
    dumped = dump_table_layout(config)
    assert dumped["tables"]["Q1"]["column_widths_mm"] == {"1": 30.0}
    assert normalize_table_layout(dumped) == config


@pytest.mark.parametrize("bad, field", [
    ({"schema_version": "2"}, "schema_version"),
    ({"schema_version": "1.0", "defaults": {"wide_table_policy": "rotate"}}, "wide_table_policy"),
    ({"schema_version": "1.0", "tables": {"Q": {"cuts": [1]}}}, "cuts"),
    ({"schema_version": "1.0", "tables": {"Q": {"cuts": [3, 3]}}}, "cuts"),
    ({"schema_version": "1.0", "tables": {"Q": {"column_widths_mm": {"x": 3}}}}, "column_widths_mm"),
    ({"schema_version": "1.0", "tables": {"Q": {"column_widths_mm": {"2": 0.5}}}}, "column_widths_mm"),
    ({"schema_version": "1.0", "tables": {"Q": {"colour": "red"}}}, "알 수 없는"),
])
def test_invalid_layout_file_names_the_field(bad, field):
    with pytest.raises(ValueError, match=field):
        normalize_table_layout(bad)


# --- explicit split boundaries and widths --------------------------------

def test_explicit_cuts_override_automatic_split():
    parts = plan_table_parts(wide_table(), 60000, 8.5, layout={"cuts": [5, 8]})
    assert [p["source_columns"] for p in parts] == [[1, 2, 3, 4], [1, 5, 6, 7], [1, 8, 9, 10]]
    assert all(sum(p["column_widths_hwpunit"]) == 60000 - 500 for p in parts)


def test_empty_cuts_keep_a_table_whole_or_report_that_it_does_not_fit():
    assert len(plan_table_parts(wide_table(4), 60000, 8.5, layout={"cuts": []})) == 1
    with pytest.raises(ValueError, match="본문 폭을 넘습니다"):
        plan_table_parts(wide_table(30), 20000, 8.5, layout={"cuts": []})


def test_cut_inside_repeated_columns_is_rejected():
    table = wide_table()
    with pytest.raises(ValueError, match="반복 보기/BASE"):
        plan_table_parts(table, 60000, 8.5, layout={"cuts": [11]})


def test_fixed_column_width_is_kept_exactly():
    parts = plan_table_parts(wide_table(4), 40000, 8.5, layout={"column_widths_mm": {"1": 40}})
    widths = parts[0]["column_widths_hwpunit"]
    assert widths[0] == margin_units(40)
    assert sum(widths) == 40000 - 500


def test_fixed_widths_wider_than_the_page_are_rejected():
    with pytest.raises(ValueError, match="본문 폭"):
        plan_table_parts(wide_table(4), 30000, 8.5, layout={"column_widths_mm": {"1": 100, "2": 100}})


def test_width_override_outside_table_is_rejected():
    with pytest.raises(ValueError, match="표 범위"):
        plan_table_parts(wide_table(4), 40000, 8.5, layout={"column_widths_mm": {"9": 20}})


# --- body merges -----------------------------------------------------------

def test_body_merge_blocks_a_cut_by_default_and_names_the_column():
    table = merge(wide_table(), 2, 4, colspan=3)  # body cell C4:C6
    with pytest.raises(ValueError, match="5열 앞에서 나누면"):
        plan_table_parts(table, 60000, 8.5, layout={"cuts": [5]})


def test_split_repeat_cuts_body_merge_and_repeats_its_text():
    table = merge(wide_table(), 2, 4, colspan=3)
    table["matrix"][1][3]["display_text"] = "공통값"
    before = copy.deepcopy(table)
    parts = plan_table_parts(table, 60000, 8.5, layout={"cuts": [5], "body_merge_policy": "split_repeat"})
    assert table == before
    assert len(parts) == 2
    assert parts[0]["matrix"][1][3]["display_text"] == "공통값"
    assert parts[0]["matrix"][1][3]["colspan"] == 1
    assert parts[1]["matrix"][1][1]["display_text"] == "공통값"
    assert parts[1]["matrix"][1][1]["colspan"] == 2
    assert "공통값" in parts[0]["layout_warnings"][0]
    # Every other value is still written exactly once.
    texts = text_by_source(parts)
    for c in range(2, 11):
        if c not in (4, 5, 6):
            assert texts[f"R3C{c}"] == {f"{c}.0"}


def test_split_repeat_lets_automatic_split_pass_a_wide_body_merge():
    table = merge(wide_table(12), 2, 2, colspan=11)
    with pytest.raises(ValueError):
        plan_table_parts(table, 18000, 8.5)
    parts = plan_table_parts(table, 18000, 8.5, layout={"body_merge_policy": "split_repeat"})
    assert len(parts) > 1
    assert [c for p in parts for c in p["source_columns"][1:]] == list(range(2, 13))


# --- header/body crossing merges -----------------------------------------

def header_crossing_table():
    table = wide_table(4, body_rows=3)
    # Two header rows; the corner label runs from the header into the first body row.
    second = [cell(2, 1, "", "header")] + [cell(2, c, f"세부{c}", "header") for c in range(2, 5)]
    table["matrix"].insert(1, second)
    for r, row in enumerate(table["matrix"], 1):
        for c, item in enumerate(row, 1):
            item.update(row=r, source_cell=f"R{r}C{c}")
    table["row_count"] = len(table["matrix"])
    merge(table, 1, 1, rowspan=3)  # 구분 spans header rows 1-2 and body row 3
    table["matrix"][0][0]["role"] = "banner_vertical"
    return table


def test_header_crossing_merge_repeats_only_base_by_default():
    prepared = prepare_report_table(header_crossing_table())
    assert prepared["header_row_count"] == 1
    assert "BASE 행만" in prepared["header_repeat_warning"]


def test_split_at_header_cuts_the_label_merge_and_repeats_header_rows():
    table = header_crossing_table()
    prepared = prepare_report_table(table, "split_at_header")
    matrix = prepared["matrix"]  # row 0 is the BASE note
    assert prepared["header_row_count"] == 3  # BASE + two header rows
    assert matrix[1][0]["rowspan"] == 2 and matrix[1][0]["display_text"] == "구분"
    remainder = matrix[3][0]
    assert remainder["covered_by"] == "" and remainder["display_text"] == "" and remainder["rowspan"] == 1
    assert {"row": 2, "col": 1, "rowspan": 2, "colspan": 1} in prepared["merged_ranges"]
    assert "헤더 경계" in prepared["header_repeat_warning"]


def test_split_at_header_never_cuts_a_value_merge():
    table = header_crossing_table()
    table["matrix"][0][0]["role"] = "base"  # a BASE count is data, not a label
    prepared = prepare_report_table(table, "split_at_header")
    assert prepared["header_row_count"] == 1
    assert "값 셀" in prepared["header_repeat_warning"]


# --- vertical banners -----------------------------------------------------

def banner_table():
    matrix = [[cell(1, 1, "구분", "header"), cell(1, 2, "", "header"), cell(1, 3, "값", "banner_horizontal")]]
    for r, (label, item) in enumerate([("성별", "남"), ("", "여"), ("연령", "20대"), ("", "30대")], 2):
        matrix.append([cell(r, 1, label, "banner_vertical"), cell(r, 2, item, "stub"), cell(r, 3, f"{r}.0", "value")])
    table = {"table_key": "B", "matrix": matrix, "row_count": 5, "col_count": 3, "merged_ranges": []}
    merge(table, 2, 1, rowspan=2)
    merge(table, 4, 1, rowspan=2)
    table["matrix"][2][0]["role"] = table["matrix"][4][0]["role"] = "blank"
    return table


def test_vertical_banner_merge_is_unchanged_by_default():
    table = banner_table()
    assert apply_vertical_banner_policy(table, "merge") is table


def test_vertical_banner_repeat_per_row_writes_the_label_in_each_row():
    table = banner_table()
    before = copy.deepcopy(table)
    result = apply_vertical_banner_policy(table, "repeat_per_row")
    assert table == before
    assert [row[0]["display_text"] for row in result["matrix"][1:]] == ["성별", "성별", "연령", "연령"]
    assert all(row[0]["rowspan"] == 1 and not row[0]["covered_by"] for row in result["matrix"][1:])
    assert result["merged_ranges"] == []
    assert result["vertical_banner_repeated"] == 2
    # Values and stubs stay where they were.
    assert [row[2]["display_text"] for row in result["matrix"][1:]] == ["2.0", "3.0", "4.0", "5.0"]
    parts = plan_table_parts(result, 40000, 8.5)
    assert parts[0]["matrix"][2][0]["display_text"] == "성별"
