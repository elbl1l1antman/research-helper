import json
from pathlib import Path
import sys

import pytest
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from report_automation_engine.document_writer import (
    ACCENT,
    DocumentWriterError,
    run_cli,
    write_document,
)

REPORT_TEMPLATE = ROOT / "templates" / "report_template_basic.pptx"
CHART_TEMPLATE = ROOT / "templates" / "chart_review_template_basic.pptx"


def cell(row, col, text, role, **extra):
    base = {"row": row, "col": col, "rowspan": 1, "colspan": 1, "covered_by": "", "display_text": text, "role": role}
    base.update(extra)
    return base


def matrix_table(key):
    matrix = [
        [cell(1, 1, "구분", "header"), cell(1, 2, "응답", "banner_horizontal", colspan=2), cell(1, 3, "", "blank", covered_by="R1C2")],
        [cell(2, 1, "", "header"), cell(2, 2, "그렇다", "banner_horizontal"), cell(2, 3, "아니다", "banner_horizontal")],
        [cell(3, 1, "전체", "stub"), cell(3, 2, "60.0", "value"), cell(3, 3, "40.0", "value")],
    ]
    return {"table_key": key, "title": "인지 여부", "matrix": matrix, "row_count": 3, "col_count": 3,
            "source_sheet": "Sheet1", "source_range": "A1:C3", "base_label": "전체", "unit": "%"}


def package(sections=2):
    keys = [f"Q{i}" for i in range(1, sections + 1)]
    return {
        "schema_version": "1.0",
        "meta": {"report_title": "2026 인식도 조사", "project_name": "테스트 사업", "created_at": "2026-11-24 10:00:00"},
        "sections": [{"table_key": key, "title": f"{key} 문항", "narrative_final": f"{key} 첫 문장.\n{key} 둘째 문장."} for key in keys],
        "tables": [matrix_table("Q1")] + [{"table_key": key, "rows": [
            {"category": "예", "percent": 55.5, "weighted_n": 100}, {"category": "아니오", "percent": 44.5, "weighted_n": 80}]}
            for key in keys[1:]],
        "charts": [
            {"table_key": "Q1", "series_group": "", "category": "그렇다", "value": 60.0, "sort_order": 1, "include_chart": True},
            {"table_key": "Q1", "series_group": "", "category": "아니다", "value": 40.0, "sort_order": 2, "include_chart": True},
        ] + [
            {"table_key": key, "series_group": group, "category": category, "value": value, "sort_order": order, "include_chart": True}
            for key in keys[1:] for group in ("남", "여")
            for order, (category, value) in enumerate([("A", 10), ("B", 30), ("C", 20)], 1)
        ],
        "qa": [{"severity": "warning", "table_key": "Q1", "message": "BASE 확인"}],
    }


@pytest.fixture
def files(tmp_path):
    def make(value=None, status="ready"):
        source = tmp_path / "report_package.json"
        source.write_text(json.dumps(value or package(), ensure_ascii=False), encoding="utf-8")
        preflight = tmp_path / "preflight_report.json"
        preflight.write_text(json.dumps({"status": status}), encoding="utf-8")
        return source, preflight
    return make


def texts(slide):
    return [shape.text_frame.text for shape in slide.shapes if shape.has_text_frame]


def all_text(prs):
    return "\n".join(text for slide in prs.slides for text in texts(slide))


def test_report_uses_template_slides_and_fills_every_placeholder(files, tmp_path):
    source, preflight = files()
    before = REPORT_TEMPLATE.read_bytes()
    out = write_document(source, tmp_path / "draft.pptx", "pptx_report", preflight, REPORT_TEMPLATE)
    assert REPORT_TEMPLATE.read_bytes() == before
    prs = Presentation(out)
    # title, two sections, QA (moved to the end)
    assert len(prs.slides) == 4
    assert prs.slide_width == Presentation(REPORT_TEMPLATE).slide_width
    assert "{{" not in all_text(prs)
    title, first, second, qa = prs.slides
    assert texts(title)[:3] == ["2026 인식도 조사", "테스트 사업", "2026-11-24 10:00:00"]
    assert "Q1 문항" in texts(first)
    assert "Q1 첫 문장.\nQ1 둘째 문장." in texts(first)
    assert any("Sheet1!A1:C3" in text and "단위: %" in text for text in texts(first))
    assert "BASE 확인" in all_text(Presentation(out)) and "BASE 확인" in "\n".join(texts(qa))
    report = json.loads((tmp_path / "draft_pptx_writer_report.json").read_text(encoding="utf-8"))
    assert report["status"] in {"ready", "ready_with_warnings"} and report["sections_written"] == 2


def test_matrix_table_is_native_with_merges_and_chart_is_editable(files, tmp_path):
    source, preflight = files()
    prs = Presentation(write_document(source, tmp_path / "draft.pptx", "pptx_report", preflight, REPORT_TEMPLATE))
    first = prs.slides[1]
    table = next(shape for shape in first.shapes if shape.has_table).table
    assert table.cell(0, 1).text == "응답" and table.cell(0, 1).is_merge_origin
    assert table.cell(2, 1).text == "60.0"
    chart_shape = next(shape for shape in first.shapes if shape.has_chart)
    assert chart_shape.chart.chart_type == XL_CHART_TYPE.PIE  # one series, two categories
    # Skeleton template boxes are thin; blocks get a usable area.
    assert chart_shape.height >= Inches(2)
    second = prs.slides[2]
    chart = next(shape for shape in second.shapes if shape.has_chart).chart
    assert chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED
    assert [series.name for series in chart.series] == ["남", "여"]
    assert chart.series[0].values == (10.0, 30.0, 20.0)
    rows_table = next(shape for shape in second.shapes if shape.has_table).table
    assert rows_table.cell(1, 0).text == "예" and rows_table.cell(1, 1).text == "55.5%"


def test_single_series_highlights_the_largest_value(files, tmp_path):
    value = package(1)
    value["charts"] = [{"table_key": "Q1", "category": c, "value": v, "sort_order": i, "include_chart": True}
                       for i, (c, v) in enumerate([("A", 5), ("B", 50), ("C", 20)], 1)]
    source, preflight = files(value)
    prs = Presentation(write_document(source, tmp_path / "d.pptx", "pptx_report", preflight, REPORT_TEMPLATE))
    chart = next(shape for shape in prs.slides[1].shapes if shape.has_chart).chart
    assert chart.series[0].points[1].format.fill.fore_color.rgb == ACCENT


def test_chart_review_has_one_slide_per_chart_and_a_note(files, tmp_path):
    source, preflight = files()
    prs = Presentation(write_document(source, tmp_path / "review.pptx", "chart_review", preflight, CHART_TEMPLATE))
    assert len(prs.slides) == 4  # title, Q1, Q2, QA
    review = prs.slides[1]
    assert "Q1 문항" in texts(review)
    note = next(text for text in texts(review) if text.startswith("차트 유형"))
    assert "강조 항목: 그렇다 (60)" in note
    chart = next(shape for shape in review.shapes if shape.has_chart)
    note_shape = next(shape for shape in review.shapes if shape.has_text_frame and shape.text_frame.text.startswith("차트 유형"))
    assert note_shape.left >= chart.left + chart.width  # note beside the chart, not on top of it


def test_chart_review_without_candidates_warns(files, tmp_path):
    value = package(1)
    value["charts"] = []
    source, preflight = files(value)
    out = write_document(source, tmp_path / "review.pptx", "chart_review", preflight, CHART_TEMPLATE)
    report = json.loads((tmp_path / "review_pptx_writer_report.json").read_text(encoding="utf-8"))
    assert any("생성할 문항/차트가 없습니다" in w for w in report["warnings"])
    assert len(Presentation(out).slides) == 2  # title and QA


def test_blocked_preflight_and_bad_schema_are_refused(files, tmp_path):
    source, preflight = files(status="blocked")
    with pytest.raises(DocumentWriterError, match="blocked"):
        write_document(source, tmp_path / "x.pptx", "pptx_report", preflight, REPORT_TEMPLATE)
    source, preflight = files({"sections": "nope"})
    with pytest.raises(DocumentWriterError, match="sections"):
        write_document(source, tmp_path / "x.pptx", "pptx_report", preflight, REPORT_TEMPLATE)
    assert not (tmp_path / "x.pptx").exists()


def test_template_cannot_be_overwritten(files, tmp_path):
    source, preflight = files()
    copy = tmp_path / "template.pptx"
    copy.write_bytes(REPORT_TEMPLATE.read_bytes())
    with pytest.raises(DocumentWriterError, match="덮어쓸"):
        write_document(source, copy, "pptx_report", preflight, copy)


def test_failing_section_becomes_a_review_slide(files, tmp_path, monkeypatch):
    import report_automation_engine.document_writer as writer

    original = writer.add_table

    def flaky(slide, table, box, report, key):
        if key == "Q2":
            raise RuntimeError("표 생성 실패")
        return original(slide, table, box, report, key)

    monkeypatch.setattr(writer, "add_table", flaky)
    source, preflight = files()
    out = write_document(source, tmp_path / "d.pptx", "pptx_report", preflight, REPORT_TEMPLATE)
    prs = Presentation(out)
    text = all_text(prs)
    assert "검토 필요: Q2 문항" in text and "Q2: 표 생성 실패" in text
    report = json.loads((tmp_path / "d_pptx_writer_report.json").read_text(encoding="utf-8"))
    assert report["status"] == "ready_with_warnings"
    assert report["failed_sections"] == [{"table_key": "Q2", "error": "표 생성 실패"}]


def test_save_failure_leaves_no_partial_file_and_writes_error_text(files, tmp_path, monkeypatch):
    import report_automation_engine.document_writer as writer

    def broken_save(self, path):
        Path(path).write_bytes(b"partial")
        raise OSError("disk full")

    monkeypatch.setattr(writer.Presentation().__class__, "save", broken_save)
    source, preflight = files()
    out = tmp_path / "d.pptx"
    with pytest.raises(OSError):
        write_document(source, out, "pptx_report", preflight, REPORT_TEMPLATE)
    assert not out.exists()
    assert not list(tmp_path.glob(".~*"))
    assert "disk full" in (tmp_path / "d.pptx.error.txt").read_text(encoding="utf-8")


def test_named_shape_template_and_repeat_marker(files, tmp_path):
    prs = Presentation()
    layout = prs.slide_layouts[6]
    title = prs.slides.add_slide(layout)
    box = title.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    box.name = "RA_REPORT_TITLE"
    box.text_frame.text = "제목 자리"
    repeat = prs.slides.add_slide(layout)
    marker = repeat.shapes.add_textbox(0, 0, Inches(1), Inches(0.3))
    marker.name = "RA_REPORT_SLIDE"
    heading = repeat.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    heading.text_frame.text = "문항: {{SECTION_TITLE}}"
    chart_box = repeat.shapes.add_shape(1, Inches(0.5), Inches(1.5), Inches(6), Inches(4))
    chart_box.name = "RA_CHART"
    template = tmp_path / "named.pptx"
    prs.save(template)

    source, preflight = files()
    out = Presentation(write_document(source, tmp_path / "named_out.pptx", "pptx_report", preflight, template))
    assert len(out.slides) == 3
    assert "문항: Q1 문항" in texts(out.slides[1])
    chart = next(shape for shape in out.slides[1].shapes if shape.has_chart)
    assert (chart.left, chart.top, chart.width, chart.height) == (Inches(0.5), Inches(1.5), Inches(6), Inches(4))


def test_without_template_falls_back_to_generated_layout(files, tmp_path, monkeypatch):
    import report_automation_engine.document_writer as writer

    monkeypatch.setattr(writer, "TEMPLATE_DIR", tmp_path / "missing")
    source, preflight = files()
    prs = Presentation(write_document(source, tmp_path / "plain.pptx", "pptx_report", preflight))
    assert len(prs.slides) == 4
    assert any(shape.has_chart for shape in prs.slides[1].shapes)


def test_cli_reports_errors_with_exit_code(files, tmp_path, capsys):
    source, preflight = files(status="blocked")
    assert run_cli(["--package", str(source), "--preflight", str(preflight), "--type", "pptx_report",
                    "--output", str(tmp_path / "o.pptx")]) == 2
    assert "blocked" in capsys.readouterr().err
