"""Report structure for PPTX drafts: parts, dividers, TOC, frame tokens, table paging, hidden QA."""

import json
from pathlib import Path
import sys

from pptx import Presentation
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from report_automation_engine.document_writer import write_document  # noqa: E402
from tests.test_document_writer_pptx import cell, files, package  # noqa: E402,F401


def texts(slide):
    return [shape.text_frame.text for shape in slide.shapes if shape.has_text_frame]


def joined(slide):
    return "\n".join(texts(slide))


def text_box(slide, text, top, name=None, left=0.5, width=9, height=0.5):
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    shape.text_frame.text = text
    if name:
        shape.name = name
    return shape


def structured_template(tmp_path, *, divider=True, toc=True, table_height=4.0):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    cover = prs.slides.add_slide(blank)
    text_box(cover, "{{REPORT_TITLE}}", 2)
    if toc:
        text_box(prs.slides.add_slide(blank), "{{TOC}}", 1, height=4)
    if divider:
        slide = prs.slides.add_slide(blank)
        text_box(slide, "PART {{PART_NUMBER}}", 1)
        text_box(slide, "{{PART_TITLE}}", 2)
        text_box(slide, "{{PART_ITEMS}}", 3, height=3)
    repeat = prs.slides.add_slide(blank)
    text_box(repeat, "{{SECTION_NUMBER}} {{SECTION_TITLE}}", 0.2)
    text_box(repeat, "{{PART_TITLE}}", 0.2, left=10, width=3)
    text_box(repeat, "{{KEY_MESSAGE}}", 0.9, name="RA_KEY_MESSAGE")
    repeat.shapes.add_shape(1, Inches(0.3), Inches(1.0), Inches(0.1), Inches(0.1)).name = "RA_KEY_MARK"
    text_box(repeat, "{{UNIT}}", 1.4, left=10, width=3)
    text_box(repeat, "{{NARRATIVE}}", 1.4, name="RA_NARRATIVE", width=9)
    table = repeat.shapes.add_shape(1, Inches(0.5), Inches(2.2), Inches(6), Inches(table_height))
    table.name = "RA_TABLE"
    chart = repeat.shapes.add_shape(1, Inches(6.8), Inches(2.2), Inches(6), Inches(table_height))
    chart.name = "RA_CHART"
    text_box(repeat, "{{SOURCE}}", 6.9, name="RA_SOURCE", width=12)
    text_box(prs.slides.add_slide(blank), "{{QA_SUMMARY}}", 1, height=5)
    path = tmp_path / "structured.pptx"
    prs.save(path)
    return path


def long_table(key, groups=3, members=12):
    """Header + total + `groups` vertical banners of `members` rows each (merged label column)."""
    header = [cell(1, 1, "구분", "header", colspan=2), cell(1, 2, "", "blank", covered_by="R1C1"),
              cell(1, 3, "사례수", "header"), cell(1, 4, "만족", "banner_horizontal"), cell(1, 5, "불만", "banner_horizontal")]
    matrix = [header, [cell(2, 1, "전체", "stub", colspan=2), cell(2, 2, "", "blank", covered_by="R2C1"),
                       cell(2, 3, "1000", "value"), cell(2, 4, "60.0", "value"), cell(2, 5, "40.0", "value")]]
    row = 3
    for g in range(groups):
        first = row
        for m in range(members):
            label = cell(row, 1, f"배너{g + 1}", "banner_vertical", rowspan=members) if m == 0 else \
                cell(row, 1, "", "blank", covered_by=f"R{first}C1")
            matrix.append([label, cell(row, 2, f"그룹{g + 1}-{m + 1}", "stub"), cell(row, 3, "50", "value"),
                           cell(row, 4, "55.0", "value"), cell(row, 5, "45.0", "value")])
            row += 1
    return {"table_key": key, "title": "긴 표", "matrix": matrix, "row_count": len(matrix), "col_count": 5,
            "source_sheet": "Sheet1", "source_range": f"A1:E{len(matrix)}", "base_label": "전체", "unit": "%"}


def write(files, tmp_path, value=None, template=None, plan=None, name="out.pptx"):
    source, preflight = files(value)
    plan_path = None
    if plan is not None:
        plan_path = tmp_path / "plan.json"
        plan_path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
    out = write_document(source, tmp_path / name, "pptx_report", preflight, template or structured_template(tmp_path),
                         plan_path=plan_path)
    report = json.loads((tmp_path / (Path(name).stem + "_pptx_writer_report.json")).read_text(encoding="utf-8"))
    return Presentation(out), report


def test_plan_parts_add_toc_dividers_and_restart_section_numbers(files, tmp_path):
    plan = {"parts": [{"title": "조사 결과 종합", "table_keys": ["Q1"]},
                      {"title": "문항별 결과", "table_keys": ["Q2", "Q3"]}]}
    prs, report = write(files, tmp_path, package(sections=3), plan=plan)
    slides = list(prs.slides)
    # cover, toc, divider 1, Q1, divider 2, Q2, Q3, QA
    assert len(slides) == 8
    assert "01 조사 결과 종합" in joined(slides[1]) and "02 문항별 결과" in joined(slides[1])
    assert "PART 01" in joined(slides[2]) and "조사 결과 종합" in joined(slides[2]) and "Q1 문항" in joined(slides[2])
    assert "01 Q1 문항" in joined(slides[3]) and "조사 결과 종합" in joined(slides[3])
    assert "PART 02" in joined(slides[4]) and "Q2 문항" in joined(slides[4]) and "Q3 문항" in joined(slides[4])
    assert "01 Q2 문항" in joined(slides[5]) and "02 Q3 문항" in joined(slides[6])
    assert "{{" not in "\n".join(joined(slide) for slide in slides)
    assert [part["title"] for part in report["structure"]["parts"]] == ["조사 결과 종합", "문항별 결과"]


def test_default_plan_is_one_part_with_every_section(files, tmp_path):
    prs, report = write(files, tmp_path, package(sections=2))
    slides = list(prs.slides)
    assert len(slides) == 6  # cover, toc, divider, Q1, Q2, QA
    assert "조사 결과" in joined(slides[2])
    assert report["structure"]["parts"][0]["table_keys"] == ["Q1", "Q2"]


def test_plan_with_unknown_or_missing_keys_warns_and_keeps_every_section(files, tmp_path):
    plan = {"parts": [{"title": "A", "table_keys": ["Q2", "NOPE"]}]}
    prs, report = write(files, tmp_path, package(sections=2), plan=plan)
    body = "\n".join(joined(slide) for slide in prs.slides)
    assert "Q1 문항" in body and "Q2 문항" in body
    assert any("NOPE" in warning for warning in report["warnings"])
    assert any("Q1" in warning and "계획" in warning for warning in report["warnings"])


def test_key_message_takes_the_first_line_and_unit_is_shown(files, tmp_path):
    prs, _ = write(files, tmp_path, package(sections=1))
    section = joined(prs.slides[3])
    assert "Q1 첫 문장." in section and "Q1 둘째 문장." in section
    narrative = next(shape for shape in prs.slides[3].shapes if shape.name == "RA_NARRATIVE")
    assert narrative.text_frame.text == "Q1 둘째 문장."
    assert "단위: %" in section


def test_templates_without_structure_slides_keep_old_behaviour(files, tmp_path):
    prs, _ = write(files, tmp_path, package(sections=2), template=structured_template(tmp_path, divider=False, toc=False))
    assert len(prs.slides) == 4  # cover, Q1, Q2, QA


def test_long_table_continues_on_extra_slides_with_repeated_header(files, tmp_path):
    value = package(sections=1)
    value["tables"] = [long_table("Q1")]
    prs, report = write(files, tmp_path, value, template=structured_template(tmp_path, divider=False, toc=False))
    pages = [slide for slide in prs.slides if any(shape.has_table for shape in slide.shapes)]
    assert len(pages) >= 2
    assert any("(계속)" in text for text in texts(pages[1]))
    assert not any(shape.has_chart for shape in pages[1].shapes)
    seen = []
    for page in pages:
        table_shape = next(shape for shape in page.shapes if shape.has_table)
        assert table_shape.top + table_shape.height <= prs.slide_height
        grid = table_shape.table
        assert grid.cell(0, 0).text == "구분"  # header repeated
        seen += [grid.cell(r, 1).text or grid.cell(r, 0).text for r in range(1, len(grid.rows))]  # 전체 spans cols 0-1
    expected = ["전체"] + [f"그룹{g}-{m}" for g in range(1, 4) for m in range(1, 13)]
    assert [text for text in seen if text] == expected
    assert report["sections_written"] == 1
    assert any("나눠" in warning for warning in report["warnings"])
    # continuation pages carry no key message, so its marker shape goes too
    assert any(shape.name == "RA_KEY_MARK" for shape in pages[0].shapes)
    assert not any(shape.name == "RA_KEY_MARK" for shape in pages[1].shapes)


def test_split_prefers_group_boundaries_and_repeats_a_cut_banner_label(files, tmp_path):
    value = package(sections=1)
    value["tables"] = [long_table("Q1", groups=2, members=40)]  # a group taller than one page must be cut
    prs, _ = write(files, tmp_path, value, template=structured_template(tmp_path, divider=False, toc=False))
    pages = [slide for slide in prs.slides if any(shape.has_table for shape in slide.shapes)]
    labels = []
    for page in pages:
        grid = next(shape for shape in page.shapes if shape.has_table).table
        first_body = next(r for r in range(1, len(grid.rows)) if grid.cell(r, 1).text)
        labels.append(grid.cell(first_body, 0).text or grid.cell(first_body, 1).text)
    # every continuation page that starts inside a banner group shows that group's label again
    assert all(label for label in labels[1:])


def wide_table(key, choices):
    header = [cell(1, 1, "구분", "header", colspan=2), cell(1, 2, "", "blank", covered_by="R1C1"), cell(1, 3, "사례수", "header")]
    header += [cell(1, 4 + i, f"보기{i + 1:02d}", "banner_horizontal") for i in range(choices)]
    total = [cell(2, 1, "전체", "stub", colspan=2), cell(2, 2, "", "blank", covered_by="R2C1"), cell(2, 3, "1000", "value")]
    total += [cell(2, 4 + i, f"{i + 1}.5", "value") for i in range(choices)]
    rows = [header, total]
    for r, name in enumerate(["남성 응답자 집단", "여성 응답자 집단"], 3):
        label = cell(r, 1, "성별", "banner_vertical", rowspan=2) if r == 3 else cell(r, 1, "", "blank", covered_by="R3C1")
        rows.append([label, cell(r, 2, name, "stub"), cell(r, 3, "500", "value")]
                    + [cell(r, 4 + i, f"{i + 2}.0", "value") for i in range(choices)])
    return {"table_key": key, "title": "넓은 표", "matrix": rows, "row_count": 4, "col_count": 3 + choices,
            "source_sheet": "Sheet1", "source_range": "A1:Z4", "base_label": "전체", "unit": "%"}


def table_slides(prs):
    return [slide for slide in prs.slides if any(shape.has_table for shape in slide.shapes)]


def grid_of(slide):
    return next(shape for shape in slide.shapes if shape.has_table)


def test_wide_table_moves_to_full_width_pages_and_chart_takes_the_first_page(files, tmp_path):
    value = package(sections=1)
    value["tables"] = [wide_table("Q1", 14)]
    prs, report = write(files, tmp_path, value, template=structured_template(tmp_path, divider=False, toc=False))
    first, following = prs.slides[1], prs.slides[2]
    assert not any(shape.has_table for shape in first.shapes)
    chart = next(shape for shape in first.shapes if shape.has_chart)
    assert chart.width > Inches(12)  # chart uses the TABLE + CHART area
    assert any("(계속)" in text for text in texts(following))
    table = grid_of(following)
    assert len(table.table.columns) == 17 and table.width <= Inches(12.31)
    widths = [column.width for column in table.table.columns]
    assert widths[1] > widths[3]  # measured widths, not equal columns
    assert any("넓어" in warning for warning in report["warnings"])


def test_table_wider_than_the_slide_is_split_by_columns_with_labels_repeated(files, tmp_path):
    value = package(sections=1)
    value["tables"] = [wide_table("Q1", 40)]
    prs, report = write(files, tmp_path, value, template=structured_template(tmp_path, divider=False, toc=False))
    pages = table_slides(prs)
    assert len(pages) >= 2
    headers = []
    for page in pages:
        grid = grid_of(page).table
        assert grid.cell(0, 0).text == "구분"  # label columns repeat in every part
        headers += [grid.cell(0, c).text for c in range(len(grid.columns)) if grid.cell(0, c).text.startswith("보기")]
    assert headers == [f"보기{i:02d}" for i in range(1, 41)]
    assert any("열" in warning and "나눠" in warning for warning in report["warnings"])


def test_wide_table_without_chart_data_starts_on_the_first_page(files, tmp_path):
    value = package(sections=1)
    value["tables"] = [wide_table("Q1", 14)]
    value["charts"] = []
    prs, _ = write(files, tmp_path, value, template=structured_template(tmp_path, divider=False, toc=False))
    first = prs.slides[1]
    assert not any(shape.has_chart for shape in first.shapes)
    assert len(grid_of(first).table.columns) == 17


def test_qa_slides_are_hidden_and_paginated(files, tmp_path):
    value = package(sections=1)
    value["qa"] = [{"severity": "warning", "table_key": f"Q{i}", "message": f"확인 {i}"} for i in range(25)]
    prs, report = write(files, tmp_path, value)
    qa = [slide for slide in prs.slides if "확인 0" in joined(slide) or "확인 24" in joined(slide)
          or any("확인 1" in t for t in texts(slide))]
    assert len(qa) >= 2
    assert all(slide._element.get("show") == "0" for slide in qa)
    assert "확인 24" in "\n".join(joined(slide) for slide in qa)
    assert report["qa_slides_hidden"] is True
