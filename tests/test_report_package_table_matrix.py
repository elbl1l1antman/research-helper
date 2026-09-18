from pathlib import Path

import openpyxl

from report_automation_engine.report_package import build_preflight, build_report_package


def make_workbook(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "보고서_분석문"
    ws.append(["table_key", "문항/표 제목", "분석문_기본", "최종 사용문", "사용자 수정문", "주요 수치 요약", "검토 상태"])
    ws.append(["T001", "만족도", "만족도는 63.3%로 나타남.", "", "", "63.3%", ""])
    table_ws = wb.create_sheet("보고서_삽입표")
    table_ws.append(["table_key", "title", "category", "weighted_n", "raw_n", "percent", "unit", "source_cell"])
    table_ws.append(["T001", "만족도", "전체", 1200, 1198, 63.25, "%", "D5"])
    table_ws.append(["T001", "만족도", "매우 만족", 353, 351, 29.44, "%", "D6"])
    wb.save(path)


def make_cell_contract_workbook(path: Path) -> None:
    make_workbook(path)
    wb = openpyxl.load_workbook(path)
    source = wb.create_sheet("결과표")
    source["A5"] = "구분"
    source.merge_cells("A5:B5")
    source["A6"] = "전체"
    source["B6"] = 0.633335353
    source["B6"].number_format = "0.0%"
    ws = wb.create_sheet("보고서_삽입표셀")
    ws.append(
        [
            "table_key",
            "title",
            "row",
            "col",
            "rowspan",
            "colspan",
            "role",
            "display_text",
            "raw_value",
            "number_format",
            "source_sheet",
            "source_cell",
            "source_range",
            "horizontal_align",
            "vertical_align",
            "covered_by",
        ]
    )
    ws.append(["T001", "만족도", 1, 1, 1, 2, "banner_horizontal", "구분", "구분", "General", "결과표", "A5", "A5:B6", "center", "center", ""])
    ws.append(["T001", "만족도", 1, 2, 1, 1, "blank", "", "", "General", "결과표", "B5", "A5:B6", "center", "center", "A5"])
    ws.append(["T001", "만족도", 2, 1, 1, 1, "stub", "전체", "전체", "General", "결과표", "A6", "A5:B6", "left", "center", ""])
    ws.append(["T001", "만족도", 2, 2, 1, 1, "value", "63.3%", 0.633335353, "0.0%", "결과표", "B6", "A5:B6", "right", "center", ""])
    wb.save(path)


def test_package_tables_include_matrix(tmp_path):
    excel = tmp_path / "sample.xlsx"
    make_workbook(excel)
    package = build_report_package(excel, {"decimal_places": "1"})
    table = package["tables"][0]
    assert table["matrix"][1][1]["display_text"] == "63.3%"
    assert table["matrix"][1][1]["raw_value"] == 63.25
    assert table["roles"]["header"] == 4


def test_preflight_counts_table_matrix_warnings(tmp_path):
    excel = tmp_path / "sample.xlsx"
    make_workbook(excel)
    package = build_report_package(excel, {"decimal_places": "1"})
    preflight = build_preflight(package, [])
    assert "table_matrix_warning_count" in preflight["summary"]
    assert preflight["summary"]["table_matrix_error_count"] == 0


def test_preflight_blocks_section_missing_insert_table_rows(tmp_path):
    excel = tmp_path / "missing_table.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "보고서_분석문"
    ws.append(["table_key", "문항/표 제목", "분석문_기본", "최종 사용문", "사용자 수정문", "주요 수치 요약", "검토 상태"])
    ws.append(["T001", "만족도", "만족도는 63.3%로 나타남.", "", "", "63.3%", ""])
    table_ws = wb.create_sheet("보고서_삽입표")
    table_ws.append(["table_key", "title", "category", "weighted_n", "raw_n", "percent", "unit", "source_cell"])
    table_ws.append(["T999", "만족도", "전체", 1200, 1198, 63.25, "%", "D5"])
    wb.save(excel)

    package = build_report_package(excel, {"decimal_places": "1"})
    preflight = build_preflight(package, [])

    assert any("삽입표 데이터가 없습니다" in item["message"] for item in package["qa"])
    assert preflight["status"] == "blocked"


def test_package_prefers_cell_contract_and_preserves_display_text(tmp_path):
    excel = tmp_path / "cell_contract.xlsx"
    make_cell_contract_workbook(excel)

    package = build_report_package(excel, {"decimal_places": "1"})
    table = package["tables"][0]

    assert table["source_sheet"] == "결과표"
    assert table["source_range"] == "A5:B6"
    assert table["matrix"][0][0]["colspan"] == 2
    assert table["matrix"][0][1]["covered_by"] == "A5"
    assert table["matrix"][1][1]["display_text"] == "63.3%"
    assert table["matrix"][1][1]["raw_value"] == 0.633335353
    assert build_preflight(package, [])["summary"]["table_row_count"] == 2


def test_preflight_blocks_cell_contract_with_missing_source_sheet(tmp_path):
    excel = tmp_path / "missing_source.xlsx"
    make_cell_contract_workbook(excel)
    wb = openpyxl.load_workbook(excel)
    ws = wb["보고서_삽입표셀"]
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 11).value = "없는시트"
    wb.save(excel)

    package = build_report_package(excel, {"decimal_places": "1"})
    preflight = build_preflight(package, [])

    assert preflight["status"] == "blocked"
    assert any("source sheet" in item["message"] for item in preflight["errors"])
