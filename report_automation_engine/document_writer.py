"""Create editable PPTX drafts from report_package.json using the user's template.

Template rules (see ``templates/report_template_basic.pptx``):

* A placeholder is ``{{TOKEN}}`` text inside a shape, or a shape named ``RA_TOKEN``.
* The **title slide** holds ``{{REPORT_TITLE}}`` / ``{{PROJECT_NAME}}`` / ``{{GENERATED_AT}}``.
* The **repeat slide** is copied once per section (``pptx_report``) or chart
  (``chart_review``). It is the slide with a shape named ``RA_REPORT_SLIDE`` /
  ``RA_CHART_REVIEW_SLIDE`` or with ``{{SECTION_TITLE}}`` / ``{{CHART_TITLE}}``.
* ``{{TABLE}}`` and ``{{CHART}}`` shapes are replaced by native, editable
  PowerPoint tables and charts in the same box. Text tokens keep the
  formatting of the run that held the token.
* ``{{QA_SUMMARY}}`` lists package QA items; that slide is moved to the end.

The template is never modified. Output is written to a temporary file and
moved into place only after a successful save; on failure the partial file is
removed and ``<output>.error.txt`` explains why. A companion
``<output stem>_pptx_writer_report.json`` records sections, warnings and
per-section failures (a failing section becomes a "검토 필요" slide instead of
stopping the whole document).
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import re
import tempfile
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.opc.packuri import PackURI
from pptx.util import Emu, Inches, Pt

try:
    from .report_table_matrix import resolve_table_metadata
except ImportError:
    from report_table_matrix import resolve_table_metadata


ACCENT = RGBColor(0xD9, 0x42, 0x2E)
BLUE = RGBColor(0x2F, 0x66, 0xA7)
GRAY = RGBColor(0x66, 0x66, 0x66)
HEADER_FILL = RGBColor(0xE8, 0xEE, 0xF6)
TOKEN = re.compile(r"\{\{([A-Z_]+)\}\}")
BLOCK_TOKENS = ("TABLE", "CHART")
REPORT_REPEAT_TOKENS = {"SECTION_TITLE", "NARRATIVE", "TABLE"}
CHART_REPEAT_TOKENS = {"CHART_TITLE", "CHART_NOTE"}
TEMPLATE_DIR = Path(__file__).resolve().parents[1] / "templates"
DEFAULT_TEMPLATES = {"pptx_report": "report_template_basic.pptx", "chart_review": "chart_review_template_basic.pptx"}
MIN_BLOCK_HEIGHT = Inches(1.2)
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CHART_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
SKIP_RELTYPES = ("/slideLayout", "/notesSlide", "/chart")


class DocumentWriterError(ValueError):
    """User-facing failure: invalid package, blocked preflight, unusable template."""


# --------------------------------------------------------------------------- entry


def write_document(
    package_path: str | Path,
    output_path: str | Path,
    document_type: str,
    preflight_path: str | Path | None = None,
    template_path: str | Path | None = None,
    *,
    chart_style_path: str | Path | None = None,
) -> Path:
    output = Path(output_path).resolve()
    package_file = Path(package_path).resolve()
    if output == package_file or (template_path and output == Path(template_path).resolve()):
        raise DocumentWriterError("입력 package나 원본 템플릿을 덮어쓸 수 없습니다.")
    package = load_json(package_path)
    preflight = load_json(preflight_path) if preflight_path else {}
    if preflight.get("status") == "blocked":
        raise DocumentWriterError("preflight status is blocked; fix errors before creating PPTX output")

    if chart_style_path is not None:
        if document_type != "chart_review" or template_path is not None:
            raise DocumentWriterError("--chart-style applies only to chart_review without --template")
        if __package__:
            from .survey_chart_style import write_chart_review
        else:
            from survey_chart_style import write_chart_review
        return write_chart_review(package, output_path, chart_style_path)

    if document_type not in DEFAULT_TEMPLATES:
        raise DocumentWriterError(f"unsupported document type: {document_type}")
    validate_package(package)
    template = resolve_template(document_type, template_path)
    report = new_report(package_file, template, output, document_type, preflight)
    report_file = output.with_name(output.stem + "_pptx_writer_report.json")
    error_file = output.with_name(output.name + ".error.txt")
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        prs = Presentation(str(template)) if template else Presentation()
        if template:
            build_from_template(prs, package, document_type, report)
        else:
            build_without_template(prs, package, document_type, report)
        renumber_slide_parts(prs)
        handle, name = tempfile.mkstemp(prefix=".~" + output.stem + "_", suffix=".pptx", dir=output.parent)
        os.close(handle)
        temporary = Path(name)
        prs.save(str(temporary))
        os.replace(temporary, output)
        temporary = None
        report.update(status="ready_with_warnings" if report["warnings"] or report["failed_sections"] else "ready",
                      finished_at=now(), slide_count=len(prs.slides))
        write_json(report_file, report)
        if error_file.exists():
            error_file.unlink()
        return output
    except Exception as exc:
        report.update(status="failed", finished_at=now(), errors=[str(exc)])
        if temporary is not None and temporary.exists():
            temporary.unlink()
        error_file.write_text(f"PPTX 초본 생성 실패: {exc}\n\n{traceback.format_exc()}", encoding="utf-8")
        write_json(report_file, report)
        raise


def resolve_template(document_type: str, template_path: str | Path | None) -> Path | None:
    if template_path:
        path = Path(template_path)
        if not path.is_file():
            raise DocumentWriterError(f"PPTX 템플릿을 찾지 못했습니다: {path}")
        return path.resolve()
    default = TEMPLATE_DIR / DEFAULT_TEMPLATES[document_type]
    return default if default.is_file() else None


def validate_package(package: Dict[str, Any]) -> None:
    if not isinstance(package, dict):
        raise DocumentWriterError("report_package.json 형식이 올바르지 않습니다: 객체가 필요합니다.")
    for key in ("sections", "tables", "charts"):
        if not isinstance(package.get(key, []), list):
            raise DocumentWriterError(f"report_package.json 형식이 올바르지 않습니다: {key}는 목록이어야 합니다.")
    if not isinstance(package.get("meta", {}), dict):
        raise DocumentWriterError("report_package.json 형식이 올바르지 않습니다: meta는 객체여야 합니다.")


def new_report(package_file, template, output, document_type, preflight) -> Dict[str, Any]:
    return {"schema_version": "1.0", "status": "running", "document_type": document_type,
            "package_path": str(package_file), "template_path": str(template or ""), "output_path": str(output),
            "preflight_status": preflight.get("status", ""), "started_at": now(), "finished_at": "",
            "sections_written": 0, "failed_sections": [], "warnings": [], "errors": []}


# --------------------------------------------------------------------- template model


def is_group(shape) -> bool:
    # shape.shape_type raises for ink (contentPart) and geometry-less shapes; read the tag instead.
    return shape._element.tag.endswith("}grpSp")


def has_text(shape) -> bool:
    try:
        return bool(shape.has_text_frame)
    except Exception:
        return False


def shape_tokens(shape) -> set[str]:
    tokens = set()
    name = str(getattr(shape, "name", "") or "")
    if name.startswith("RA_"):
        tokens.add(name[3:])
    if has_text(shape):
        tokens.update(TOKEN.findall(shape.text_frame.text))
    if is_group(shape):
        for child in shape.shapes:
            tokens |= shape_tokens(child)
    return tokens


def slide_tokens(slide) -> set[str]:
    tokens = set()
    for shape in slide.shapes:
        tokens |= shape_tokens(shape)
    return tokens


def classify_slides(prs, document_type: str) -> Dict[str, Any]:
    repeat_tokens = REPORT_REPEAT_TOKENS if document_type == "pptx_report" else CHART_REPEAT_TOKENS
    marker = "REPORT_SLIDE" if document_type == "pptx_report" else "CHART_REVIEW_SLIDE"
    roles = {"title": [], "repeat": None, "qa": [], "other": []}
    for slide in prs.slides:
        tokens = slide_tokens(slide)
        if roles["repeat"] is None and (marker in tokens or tokens & repeat_tokens
                                        or (document_type == "chart_review" and "CHART" in tokens and "TABLE" not in tokens)):
            roles["repeat"] = slide
        elif "QA_SUMMARY" in tokens:
            roles["qa"].append(slide)
        elif tokens & {"REPORT_TITLE", "PROJECT_NAME", "GENERATED_AT", "TITLE_SLIDE"}:
            roles["title"].append(slide)
        else:
            roles["other"].append(slide)
    return roles


def build_from_template(prs, package: Dict[str, Any], document_type: str, report: Dict[str, Any]) -> None:
    roles = classify_slides(prs, document_type)
    if roles["repeat"] is None:
        report["warnings"].append("템플릿에서 반복 슬라이드(RA_REPORT_SLIDE 또는 {{SECTION_TITLE}}/{{CHART_TITLE}})를 찾지 못해 기본 배치로 만들었습니다.")
        for slide in list(prs.slides):
            delete_slide(prs, slide)
        build_without_template(prs, package, document_type, report, keep_size=True)
        return
    values = report_values(package)
    for slide in roles["title"]:
        fill_slide(slide, values, report, context="제목")
    for slide in roles["qa"]:
        fill_slide(slide, values, report, context="QA")

    template_slide = roles["repeat"]
    position = slide_index(prs, template_slide)
    items = list(iter_items(package, document_type))
    if not items:
        report["warnings"].append("생성할 문항/차트가 없습니다. preflight_report.json을 확인하세요.")
    failures = []
    for offset, item in enumerate(items):
        slide = duplicate_slide(prs, template_slide, report)
        move_slide(prs, slide, position + offset + 1)
        try:
            fill_item_slide(slide, package, item, document_type, report)
            report["sections_written"] += 1
        except Exception as exc:  # one bad section must not stop the draft
            failures.append(f"{item['table_key'] or '(table_key 없음)'}: {exc}")
            report["failed_sections"].append({"table_key": item["table_key"], "error": str(exc)})
            clear_slide(slide)
            add_text(slide, f"검토 필요: {item['title'] or item['table_key']}", Inches(0.6), Inches(0.4), Inches(11), Inches(0.6), 22, BLUE, bold=True)
            add_text(slide, str(exc), Inches(0.6), Inches(1.3), Inches(11), Inches(4), 14, GRAY)
    delete_slide(prs, template_slide)
    for slide in roles["qa"]:
        move_slide(prs, slide, len(prs.slides._sldIdLst) - 1)
    if failures:
        slide = prs.slides.add_slide(blank_layout(prs))
        add_text(slide, "검토 필요", Inches(0.6), Inches(0.4), Inches(11), Inches(0.6), 22, BLUE, bold=True)
        add_text(slide, "\n".join(failures), Inches(0.6), Inches(1.3), Inches(11.5), Inches(5.4), 13, RGBColor(0, 0, 0))


def iter_items(package: Dict[str, Any], document_type: str) -> Iterable[Dict[str, Any]]:
    if document_type == "pptx_report":
        for section in package.get("sections", []):
            key = str(section.get("table_key", ""))
            yield {"table_key": key, "title": str(section.get("title") or key or "제목 없음"), "section": section}
        return
    for key, rows in group_by([row for row in package.get("charts", []) if row.get("include_chart")], "table_key").items():
        yield {"table_key": key, "title": section_title(package, key) or key or "차트", "rows": rows}


def report_values(package: Dict[str, Any]) -> Dict[str, str]:
    meta = package.get("meta", {})
    qa = package.get("qa", [])
    lines = [f"{q.get('severity', '')}: [{q.get('table_key', '')}] {q.get('message', '')}".replace("[] ", "") for q in qa[:20]]
    if len(qa) > 20:
        lines.append(f"... 외 {len(qa) - 20}건")
    return {
        "REPORT_TITLE": str(meta.get("report_title") or meta.get("source_file_name") or "보고서 초안"),
        "PROJECT_NAME": str(meta.get("project_name") or ""),
        "GENERATED_AT": str(meta.get("created_at") or now()),
        "QA_SUMMARY": "\n".join(lines) if lines else "QA 항목이 없습니다.",
    }


def fill_item_slide(slide, package, item, document_type, report) -> None:
    key = item["table_key"]
    table = find_table(package, key)
    charts = item.get("rows") if document_type == "chart_review" else chart_rows(package, key)
    values = report_values(package)
    values.update(SECTION_TITLE=item["title"], CHART_TITLE=item["title"], SOURCE=source_text(table, key),
                  NARRATIVE=narrative_text(item.get("section") or {}), CHART_NOTE=chart_note(charts) if charts else "차트 데이터 없음")
    boxes = block_boxes(slide, report, key)
    fill_slide(slide, values, report, context=key, placed=set(boxes))
    layout_blocks(slide, boxes)
    if "TABLE" in boxes:
        add_table(slide, table, boxes["TABLE"], report, key)
    if "CHART" in boxes:
        add_chart(slide, charts, boxes["CHART"], report, key)


# ------------------------------------------------------------------- placeholders


def block_boxes(slide, report: Dict[str, Any] | None = None, context: str = "") -> Dict[str, Dict[str, Any]]:
    """Find ``{{TABLE}}``/``{{CHART}}`` holders and return their boxes.

    A holder that contains only the token (or is named ``RA_TABLE``/``RA_CHART``,
    or is an existing table/chart) is removed. A token inside other text keeps
    the shape (the token is blanked later) and the object is placed over its box.
    """
    boxes = {}

    def visit(shapes, in_group):
        for shape in list(shapes):
            if is_group(shape):
                visit(shape.shapes, True)
                continue
            tokens = shape_tokens(shape)
            text = shape.text_frame.text.strip() if has_text(shape) else ""
            for token in BLOCK_TOKENS:
                if token not in tokens or token in boxes:
                    continue
                boxes[token] = {"left": shape.left, "top": shape.top, "width": shape.width, "height": shape.height}
                whole = (str(shape.name or "") == "RA_" + token or text == "{{" + token + "}}"
                         or getattr(shape, "has_chart", False) or getattr(shape, "has_table", False))
                if in_group and report is not None:
                    report["warnings"].append(f"{context}: {{{{{token}}}}}가 그룹 도형 안에 있어 위치가 어긋날 수 있습니다.")
                if whole:
                    shape._element.getparent().remove(shape._element)
                    break
                if report is not None:
                    report["warnings"].append(f"{context}: {{{{{token}}}}}가 다른 글자와 같은 도형에 있어 그 도형 위치에 겹쳐 넣었습니다.")

    visit(slide.shapes, False)
    return boxes


def layout_blocks(slide, boxes) -> None:
    """Skeleton templates stack thin text boxes; give tables/charts a usable area."""
    if not boxes or all(box["height"] >= MIN_BLOCK_HEIGHT for box in boxes.values()):
        return
    part = slide.part.package.presentation_part.presentation
    width, height = part.slide_width, part.slide_height
    margin = min(box["left"] for box in boxes.values())
    top = min(box["top"] for box in boxes.values())
    bottom = height - Inches(0.75)
    # Keep the narrative readable above the blocks and push the source line to the bottom.
    for shape in slide.shapes:
        name = str(shape.name or "")
        if name == "RA_NARRATIVE" and shape.top < top:
            shape.height = max(shape.height, Inches(1.2))
            top = max(top, shape.top + shape.height + Inches(0.15))
        elif name == "RA_SOURCE":
            shape.top = height - Inches(0.6)
            shape.height = Inches(0.4)
    top = min(top, bottom - Inches(2))
    usable = width - 2 * margin
    note = next((shape for shape in slide.shapes if str(shape.name or "") == "RA_CHART_NOTE"), None)
    if note is not None and note.height < MIN_BLOCK_HEIGHT:
        # Chart review: data note on the right, blocks on the left.
        side = int(usable * 0.36)
        usable -= side + Inches(0.3)
        note.left, note.top, note.width, note.height = margin + usable + Inches(0.3), top, side, bottom - top
    keys = [key for key in BLOCK_TOKENS if key in boxes]
    gap = Inches(0.3)
    each = (usable - gap * (len(keys) - 1)) // len(keys)
    for index, key in enumerate(keys):
        boxes[key].update(left=margin + index * (each + gap), top=top, width=each, height=bottom - top)


def fill_slide(slide, values: Dict[str, str], report: Dict[str, Any], context: str = "", placed=frozenset()) -> None:
    """Replace every text token; block tokens not in ``placed`` are blanked with a warning."""
    seen = set()
    for shape in iter_text_shapes(slide.shapes):
        frame = shape.text_frame
        for paragraph in list(frame.paragraphs):
            text = "".join(run.text for run in paragraph.runs)
            tokens = TOKEN.findall(text)
            if not tokens:
                continue
            for token in tokens:
                if token in BLOCK_TOKENS and token not in placed:
                    report["warnings"].append(f"{context}: {{{{{token}}}}} 자리를 찾지 못해 비워 두었습니다.")
                seen.add(token)
            replaced = TOKEN.sub(lambda m: values.get(m.group(1), m.group(0)) if m.group(1) not in BLOCK_TOKENS else "", text)
            set_paragraph_text(frame, paragraph, replaced)
        if shape.name and str(shape.name).startswith("RA_"):
            frame.word_wrap = True


def iter_text_shapes(shapes):
    for shape in shapes:
        if is_group(shape):
            yield from iter_text_shapes(shape.shapes)
        elif has_text(shape):
            yield shape


def set_paragraph_text(frame, paragraph, text: str) -> None:
    """Replace a paragraph's text, keeping the first run's formatting; split lines into paragraphs."""
    lines = text.split("\n")
    runs = paragraph.runs
    if runs:
        runs[0].text = lines[0]
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        paragraph.add_run().text = lines[0]
    anchor = paragraph._p
    for line in lines[1:]:
        clone = copy.deepcopy(paragraph._p)
        anchor.addnext(clone)
        anchor = clone
        new_runs = [r for r in clone.iter() if r.tag.endswith("}r")]
        for extra in new_runs[1:]:
            extra.getparent().remove(extra)
        texts = [t for t in clone.iter() if t.tag.endswith("}t")]
        if texts:
            texts[0].text = line


def narrative_text(section: Dict[str, Any]) -> str:
    blocks = section.get("narrative_blocks")
    if blocks:
        return "\n".join(str(block.get("text") or "") for block in blocks if str(block.get("text") or "").strip())
    return str(section.get("narrative_final") or "")


def source_text(table: Dict[str, Any] | None, key: str) -> str:
    if not table:
        return f"출처: {key}" if key else ""
    parts = []
    if table.get("source_sheet") and table.get("source_range"):
        parts.append(f"출처: 원본 집계표 ({table['source_sheet']}!{table['source_range']})")
    else:
        parts.append(f"출처: {key}")
    if table.get("matrix"):
        metadata = resolve_table_metadata(table)
        parts.append(f"BASE: {metadata['base_label']}")
        parts.append(f"단위: {metadata['unit']}")
    return " | ".join(parts)


# ------------------------------------------------------------------ tables/charts


def table_grid(table: Dict[str, Any] | None):
    """Return (rows of text, merges, header row count) for a package table."""
    if not table:
        return [], [], 0
    matrix = table.get("matrix")
    if matrix:
        rows = [[str(cell.get("display_text") or "") if not cell.get("covered_by") else "" for cell in row] for row in matrix]
        merges = [(r, c, int(cell.get("rowspan") or 1), int(cell.get("colspan") or 1))
                  for r, row in enumerate(matrix) for c, cell in enumerate(row)
                  if not cell.get("covered_by") and (int(cell.get("rowspan") or 1) > 1 or int(cell.get("colspan") or 1) > 1)]
        header = 0
        for row in matrix:
            roles = {cell.get("role") for cell in row if not cell.get("covered_by")}
            if "value" in roles:
                break
            header += 1
        return rows, merges, min(header, len(rows))
    body = [[str(row.get("category", "")), display_percent(row), display_number(row.get("weighted_n"))] for row in table.get("rows", [])]
    return ([["항목", "%", "N"]] + body) if body else [], [], 1


def add_table(slide, table, box, report, key) -> None:
    rows, merges, header = table_grid(table)
    if not rows:
        add_text(slide, "삽입표 데이터 없음", box["left"], box["top"], box["width"], Inches(0.5), 11, GRAY)
        report["warnings"].append(f"{key}: 삽입표 데이터가 없습니다.")
        return
    count_rows, count_cols = len(rows), max(len(row) for row in rows)
    if count_cols > 12 or count_rows > 30:
        report["warnings"].append(f"{key}: 표가 {count_rows}행 x {count_cols}열이라 슬라이드를 넘을 수 있습니다. 표를 나누거나 글자 크기를 확인하세요.")
    size = 11 if count_rows <= 10 and count_cols <= 6 else 9 if count_rows <= 18 and count_cols <= 9 else 7
    row_height = Emu(int(Pt(size) * 2.1))
    height = min(box["height"], row_height * count_rows)
    shape = slide.shapes.add_table(count_rows, count_cols, box["left"], box["top"], box["width"], height)
    shape.name = "RA_TABLE"
    grid = shape.table
    grid.first_row = header > 0
    for r, values in enumerate(rows):
        grid.rows[r].height = row_height
        for c in range(count_cols):
            cell = grid.cell(r, c)
            cell.text = values[c] if c < len(values) else ""
            cell.margin_left = cell.margin_right = Pt(3)
            cell.margin_top = cell.margin_bottom = Pt(1)
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(size)
                    run.font.bold = r < header
            if r < header:
                cell.fill.solid()
                cell.fill.fore_color.rgb = HEADER_FILL
                for run in (run for p in cell.text_frame.paragraphs for run in p.runs):
                    run.font.color.rgb = RGBColor(0, 0, 0)
    for r, c, rowspan, colspan in merges:
        if r + rowspan <= count_rows and c + colspan <= count_cols:
            grid.cell(r, c).merge(grid.cell(r + rowspan - 1, c + colspan - 1))


def chart_series(rows: List[Dict[str, Any]]):
    """Group chart rows into (categories, [(series name, values)])."""
    usable = [row for row in rows if number(row.get("value")) is not None]
    groups = group_by(usable, "series_group")
    ordered = sorted(usable, key=lambda row: (number(row.get("sort_order")) if number(row.get("sort_order")) is not None else 1e9))
    categories = list(dict.fromkeys(str(row.get("category", "")) for row in ordered))
    series = []
    for name, items in groups.items():
        by_category = {str(row.get("category", "")): number(row.get("value")) for row in items}
        series.append((name or "전체", [by_category.get(category) for category in categories]))
    return categories, series


def add_chart(slide, rows, box, report, key) -> None:
    categories, series = chart_series(rows or [])
    if not categories:
        add_text(slide, "차트 데이터 없음", box["left"], box["top"], box["width"], Inches(0.5), 12, GRAY)
        return
    data = CategoryChartData()
    data.categories = categories
    for name, values in series:
        data.add_series(name, values)
    chart_type = XL_CHART_TYPE.PIE if choose_chart_type(categories, series) == "pie" else XL_CHART_TYPE.COLUMN_CLUSTERED
    frame = slide.shapes.add_chart(chart_type, box["left"], box["top"], box["width"], box["height"], data)
    frame.name = "RA_CHART"
    chart = frame.chart
    chart.has_title = False
    chart.has_legend = chart_type == XL_CHART_TYPE.PIE or len(series) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM if len(series) > 1 else XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.number_format = '0.0'
    plot.data_labels.number_format_is_linked = False
    plot.data_labels.font.size = Pt(9)
    if len(series) == 1:
        highlight_max(chart, series[0][1])
    if len(categories) > 12:
        report["warnings"].append(f"{key}: 차트 항목이 {len(categories)}개라 글자가 겹칠 수 있습니다.")


def highlight_max(chart, values: List[float | None]) -> None:
    numbers = [value if value is not None else float("-inf") for value in values]
    max_idx = max(range(len(numbers)), key=numbers.__getitem__)
    for idx, point in enumerate(chart.series[0].points):
        fill = point.format.fill
        fill.solid()
        fill.fore_color.rgb = ACCENT if idx == max_idx else BLUE


def choose_chart_type(categories, series) -> str:
    return "pie" if len(series) == 1 and len(categories) == 2 else "column"


def chart_note(rows: List[Dict[str, Any]]) -> str:
    categories, series = chart_series(rows)
    if not categories:
        return "차트 데이터 없음"
    lines = [f"차트 유형: {'원그래프' if choose_chart_type(categories, series) == 'pie' else '세로 막대'}"]
    if len(series) == 1:
        values = series[0][1]
        best = max(range(len(values)), key=lambda i: values[i] if values[i] is not None else float("-inf"))
        lines.append(f"강조 항목: {categories[best]} ({format_value(values[best])})")
    lines += ["", "데이터"]
    for name, values in series:
        prefix = f"[{name}] " if len(series) > 1 else ""
        lines += [f"- {prefix}{category}: {format_value(value)}" for category, value in zip(categories, values)]
    return "\n".join(lines)


# ---------------------------------------------------------------- slide plumbing


def duplicate_slide(prs, source, report):
    """Copy a slide's shapes and media relationships onto a new slide with the same layout."""
    slide = prs.slides.add_slide(source.slide_layout)
    for shape in list(slide.shapes):
        shape._element.getparent().remove(shape._element)
    mapping = {}
    for rel_id, rel in source.part.rels.items():
        if any(rel.reltype.endswith(skip) for skip in SKIP_RELTYPES):
            continue
        if rel.is_external:
            mapping[rel_id] = slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            mapping[rel_id] = slide.part.relate_to(rel.target_part, rel.reltype)
    chart_rels = {rel_id for rel_id, rel in source.part.rels.items() if rel.reltype == CHART_RELTYPE}
    tree = slide.shapes._spTree
    for element in source.shapes._spTree.iterchildren():
        tag = element.tag.rsplit("}", 1)[-1]
        if tag in ("nvGrpSpPr", "grpSpPr", "extLst"):
            continue
        clone = copy.deepcopy(element)
        if any(value in chart_rels for node in clone.iter() for key, value in node.attrib.items() if key.startswith("{" + REL_NS + "}")):
            name = next((n.get("name") for n in clone.iter() if n.tag.endswith("}cNvPr")), "")
            if name not in ("RA_CHART",):
                report["warnings"].append(f"반복 슬라이드의 차트 개체({name})는 복사하지 않았습니다.")
                continue
        for node in clone.iter():
            for attr, value in list(node.attrib.items()):
                if attr.startswith("{" + REL_NS + "}") and value in mapping:
                    node.set(attr, mapping[value])
        tree.append(clone)
    background = source._element.cSld.bg
    if background is not None:
        slide._element.cSld.insert(0, copy.deepcopy(background))
    return slide


def renumber_slide_parts(prs) -> None:
    """python-pptx names new slides by count; after deletions names can collide."""
    slides = list(prs.slides)
    for index, slide in enumerate(slides, 1):
        slide.part.partname = PackURI(f"/ppt/slides/~renumber{index}.xml")
    for index, slide in enumerate(slides, 1):
        slide.part.partname = PackURI(f"/ppt/slides/slide{index}.xml")


def clear_slide(slide) -> None:
    for shape in list(slide.shapes):
        shape._element.getparent().remove(shape._element)


def slide_index(prs, slide) -> int:
    return list(prs.slides).index(slide)


def move_slide(prs, slide, index: int) -> None:
    ids = prs.slides._sldIdLst
    element = next(item for item in ids if prs.slides.get(int(item.get("id"))) is slide)
    ids.remove(element)
    ids.insert(index, element)


def delete_slide(prs, slide) -> None:
    ids = prs.slides._sldIdLst
    for item in list(ids):
        if prs.slides.get(int(item.get("id"))) is slide:
            prs.part.drop_rel(item.rId)
            ids.remove(item)
            return


def blank_layout(prs):
    for layout in prs.slide_layouts:
        if not len(layout.placeholders):
            return layout
    return prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]


def add_text(slide, text: str, left, top, width, height, size: int, color=RGBColor(0, 0, 0), bold=False) -> None:
    shape = slide.shapes.add_textbox(left, top, width, height)
    frame = shape.text_frame
    frame.word_wrap = True
    for idx, line in enumerate(str(text).splitlines() or [""]):
        paragraph = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        run = paragraph.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color


# ---------------------------------------------------------------- no template


def build_without_template(prs, package, document_type, report, keep_size=False) -> None:
    """Fallback layout used when no template (or no repeat slide) is available."""
    if not keep_size:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    width = prs.slide_width
    values = report_values(package)
    slide = prs.slides.add_slide(blank_layout(prs))
    add_text(slide, values["REPORT_TITLE"], Inches(0.8), Inches(2.2), width - Inches(1.6), Inches(0.8), 30, BLUE, bold=True)
    add_text(slide, values["GENERATED_AT"], Inches(0.8), Inches(3.1), width - Inches(1.6), Inches(0.5), 14, GRAY)
    half = (width - Inches(1.6)) // 2
    for item in iter_items(package, document_type):
        key = item["table_key"]
        slide = prs.slides.add_slide(blank_layout(prs))
        try:
            add_text(slide, item["title"], Inches(0.6), Inches(0.35), width - Inches(1.2), Inches(0.6), 22, BLUE, bold=True)
            if document_type == "pptx_report":
                add_text(slide, narrative_text(item["section"]), Inches(0.6), Inches(1.05), width - Inches(1.2), Inches(1.3), 13)
                top = Inches(2.5)
                add_table(slide, find_table(package, key), {"left": Inches(0.6), "top": top, "width": half - Inches(0.15), "height": Inches(4.2)}, report, key)
                add_chart(slide, chart_rows(package, key), {"left": Inches(0.6) + half + Inches(0.15), "top": top, "width": half - Inches(0.15), "height": Inches(4.2)}, report, key)
            else:
                add_chart(slide, item["rows"], {"left": Inches(0.6), "top": Inches(1.3), "width": half + Inches(1.2), "height": Inches(5.0)}, report, key)
                add_text(slide, chart_note(item["rows"]), Inches(0.6) + half + Inches(1.5), Inches(1.3), half - Inches(1.5), Inches(5.0), 12, GRAY)
            add_text(slide, source_text(find_table(package, key), key), Inches(0.6), prs.slide_height - Inches(0.6), width - Inches(1.2), Inches(0.3), 9, GRAY)
            report["sections_written"] += 1
        except Exception as exc:
            report["failed_sections"].append({"table_key": key, "error": str(exc)})
            add_text(slide, f"검토 필요: {exc}", Inches(0.6), Inches(1.3), width - Inches(1.2), Inches(2), 14, GRAY)
    if package.get("qa"):
        slide = prs.slides.add_slide(blank_layout(prs))
        add_text(slide, "QA 요약", Inches(0.6), Inches(0.35), width - Inches(1.2), Inches(0.6), 22, BLUE, bold=True)
        add_text(slide, values["QA_SUMMARY"], Inches(0.6), Inches(1.2), width - Inches(1.2), Inches(5.6), 13)


# ---------------------------------------------------------------------- helpers


def find_table(package: Dict[str, Any], table_key: str) -> Dict[str, Any] | None:
    return next((table for table in package.get("tables", []) if str(table.get("table_key", "")) == table_key), None)


def table_rows(package: Dict[str, Any], table_key: str) -> List[Dict[str, Any]]:
    table = find_table(package, table_key)
    return list(table.get("rows", [])) if table else []


def chart_rows(package: Dict[str, Any], table_key: str) -> List[Dict[str, Any]]:
    return [row for row in package.get("charts", []) if str(row.get("table_key", "")) == table_key and row.get("include_chart")]


def section_title(package: Dict[str, Any], table_key: str) -> str:
    for section in package.get("sections", []):
        if section.get("table_key") == table_key:
            return str(section.get("title") or "")
    return ""


def display_percent(row: Dict[str, Any]) -> str:
    value = row.get("percent")
    unit = row.get("unit") or "%"
    return "" if value in (None, "") else f"{format_value(number(value))}{unit}"


def display_number(value: Any) -> str:
    parsed = number(value)
    if parsed is None:
        return ""
    return f"{parsed:,.0f}" if parsed == int(parsed) else f"{parsed:,.1f}"


def format_value(value: float | None) -> str:
    if value is None:
        return "-"
    return f"{value:.1f}".rstrip("0").rstrip(".") if value != int(value) else str(int(value))


def group_by(rows: List[Dict[str, Any]], key: str) -> Dict[str, List[Dict[str, Any]]]:
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for row in rows:
        grouped.setdefault(str(row.get(key, "") or ""), []).append(row)
    return grouped


def number(value: Any) -> float | None:
    try:
        parsed = float(value)
    except (TypeError, ValueError):
        return None
    return parsed if parsed == parsed and parsed not in (float("inf"), float("-inf")) else None


def load_json(path: str | Path | None) -> Dict[str, Any]:
    if not path:
        return {}
    try:
        return json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise DocumentWriterError(f"JSON을 읽지 못했습니다: {path}: {exc}") from exc


def write_json(path: Path, value: Dict[str, Any]) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create editable PPTX drafts from report_package.json.")
    parser.add_argument("--package", required=True)
    parser.add_argument("--preflight")
    parser.add_argument("--template", help="PPTX template with {{TOKEN}} text or RA_ shape names (default: bundled template).")
    parser.add_argument("--chart-style")
    parser.add_argument("--type", required=True, choices=["pptx_report", "chart_review"])
    parser.add_argument("--output", required=True)
    args = parser.parse_args(argv)
    try:
        output = write_document(args.package, args.output, args.type, args.preflight, args.template, chart_style_path=args.chart_style)
    except (DocumentWriterError, OSError) as exc:
        print(f"PPTX 초본 생성 실패: {exc}", file=__import__("sys").stderr)
        return 2
    print(str(output))
    return 0


if __name__ == "__main__":
    raise SystemExit(run_cli())
