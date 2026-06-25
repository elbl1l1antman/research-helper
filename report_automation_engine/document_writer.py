"""Create first-pass PPTX drafts from report_package.json.

This writer intentionally uses only the standard library.  It creates editable
PowerPoint XML slides with text placeholders and chart data blocks; full chart
objects can replace these blocks after the package contract stabilizes.
"""

from __future__ import annotations

import argparse
import html
import json
import zipfile
from pathlib import Path
from typing import Any, Dict, Iterable, List


EMU_WIDE = (12192000, 6858000)


def write_document(
    package_path: str | Path,
    output_path: str | Path,
    document_type: str,
    preflight_path: str | Path | None = None,
    template_path: str | Path | None = None,
) -> Path:
    package = load_json(package_path)
    preflight = load_json(preflight_path) if preflight_path else {}
    if preflight.get("status") == "blocked":
        raise ValueError("preflight status is blocked; fix errors before creating PPTX output")
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if document_type == "chart_review":
        slides = chart_review_slides(package)
    elif document_type == "pptx_report":
        slides = report_slides(package)
    else:
        raise ValueError(f"unsupported document type for this writer: {document_type}")

    # ponytail: template_path is accepted for the stable CLI now; v1 writes a
    # minimal deck until template cloning is worth the extra OpenXML surface.
    _ = template_path
    write_pptx(output, slides)
    return output


def chart_review_slides(package: Dict[str, Any]) -> List[List[Dict[str, str]]]:
    slides = [title_slide(package, "Chart Review Draft")]
    charts = [row for row in package.get("charts", []) if row.get("include_chart")]
    if not charts:
        slides.append(text_slide("검토 필요", ["차트 후보가 없습니다.", "preflight_report.json을 확인하세요."]))
        return slides

    for table_key, rows in group_by(charts, "table_key").items():
        title = section_title(package, table_key) or table_key
        chart_type = choose_chart_type(rows)
        highlight = max(rows, key=lambda row: number(row.get("value")) or 0)
        body = [
            f"차트 유형: {chart_type}",
            f"강조 항목: {highlight.get('category', '')} ({highlight.get('display_value', highlight.get('value', ''))})",
            "",
            *chart_lines(rows),
            "",
            f"source: {table_key}",
        ]
        slides.append(text_slide(title, body))
    return slides


def report_slides(package: Dict[str, Any]) -> List[List[Dict[str, str]]]:
    slides = [title_slide(package, "Report Draft")]
    for section in package.get("sections", []):
        key = section.get("table_key", "")
        body = [
            section.get("narrative_final", ""),
            "",
            "표",
            *table_lines(package, key),
            "",
            "차트",
            *chart_lines([row for row in package.get("charts", []) if row.get("table_key") == key and row.get("include_chart")]),
            "",
            f"source: {key}",
        ]
        slides.append(text_slide(section.get("title") or key or "Untitled", body))
    if package.get("qa"):
        slides.append(text_slide("QA Summary", [f"{q.get('severity', '')}: {q.get('message', '')}" for q in package["qa"][:12]]))
    return slides


def title_slide(package: Dict[str, Any], fallback_title: str) -> List[Dict[str, str]]:
    meta = package.get("meta", {})
    title = meta.get("report_title") or meta.get("source_file_name") or fallback_title
    return [
        box("title", str(title)),
        box("subtitle", str(meta.get("created_at", ""))),
    ]


def text_slide(title: str, lines: Iterable[str]) -> List[Dict[str, str]]:
    return [
        box("title", title),
        box("body", "\n".join(str(line) for line in lines if line is not None)),
    ]


def box(kind: str, text: str) -> Dict[str, str]:
    return {"kind": kind, "text": text}


def section_title(package: Dict[str, Any], table_key: str) -> str:
    for section in package.get("sections", []):
        if section.get("table_key") == table_key:
            return str(section.get("title") or "")
    return ""


def table_lines(package: Dict[str, Any], table_key: str) -> List[str]:
    for table in package.get("tables", []):
        if table.get("table_key") == table_key:
            return [
                f"{row.get('category', '')}\t{row.get('percent', '')}{row.get('unit', '')}\tN={row.get('weighted_n', '')}"
                for row in table.get("rows", [])[:8]
            ]
    return ["삽입표 데이터 없음"]


def chart_lines(rows: List[Dict[str, Any]]) -> List[str]:
    if not rows:
        return ["차트 데이터 없음"]
    return [
        f"{row.get('category', '')}\t{row.get('display_value') or row.get('value', '')}"
        for row in sorted(rows, key=lambda row: number(row.get("sort_order")) or 9999)[:10]
    ]


def choose_chart_type(rows: List[Dict[str, Any]]) -> str:
    categories = {str(row.get("category", "")) for row in rows if row.get("category")}
    return "pie" if len(categories) == 2 else "column"


def group_by(rows: List[Dict[str, Any]], key: str) -> Dict[str, List[Dict[str, Any]]]:
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for row in rows:
        grouped.setdefault(str(row.get(key, "")), []).append(row)
    return grouped


def number(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def load_json(path: str | Path | None) -> Dict[str, Any]:
    if not path:
        return {}
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_pptx(path: Path, slides: List[List[Dict[str, str]]]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(slides)))
        zf.writestr("_rels/.rels", package_rels_xml())
        zf.writestr("docProps/app.xml", app_props_xml(len(slides)))
        zf.writestr("docProps/core.xml", core_props_xml())
        zf.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
        zf.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(len(slides)))
        for index, slide in enumerate(slides, start=1):
            zf.writestr(f"ppt/slides/slide{index}.xml", slide_xml(slide))


def slide_xml(slide: List[Dict[str, str]]) -> str:
    shapes = []
    for idx, item in enumerate(slide, start=1):
        if item["kind"] == "title":
            shapes.append(text_shape(idx, item["text"], 650000, 400000, 10800000, 700000, 3000))
        elif item["kind"] == "subtitle":
            shapes.append(text_shape(idx, item["text"], 650000, 1250000, 10800000, 420000, 1700))
        else:
            shapes.append(text_shape(idx, item["text"], 650000, 1350000, 10800000, 4800000, 1600))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    {''.join(shapes)}
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''


def text_shape(idx: int, text: str, x: int, y: int, cx: int, cy: int, size: int) -> str:
    escaped = html.escape(text)
    lines = escaped.split("\n") or [""]
    paragraphs = "".join(f'<a:p><a:r><a:rPr lang="ko-KR" sz="{size}"/><a:t>{line}</a:t></a:r></a:p>' for line in lines)
    return f'''
<p:sp>
  <p:nvSpPr><p:cNvPr id="{idx + 1}" name="RA_WRITER_{idx}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>{paragraphs}</p:txBody>
</p:sp>'''


def content_types_xml(slide_count: int) -> str:
    overrides = "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  {overrides}
</Types>'''


def package_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''


def presentation_rels_xml(slide_count: int) -> str:
    rels = "\n".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {rels}
</Relationships>'''


def presentation_xml(slide_count: int) -> str:
    slide_ids = "\n".join(f'<p:sldId id="{255 + i}" r:id="rId{i}"/>' for i in range(1, slide_count + 1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
  <p:sldSz cx="{EMU_WIDE[0]}" cy="{EMU_WIDE[1]}" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>'''


def app_props_xml(slide_count: int) -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>ReportAutomation</Application>
  <Slides>{slide_count}</Slides>
</Properties>'''


def core_props_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>ReportAutomation PPTX Draft</dc:title>
  <dc:creator>ReportAutomation</dc:creator>
</cp:coreProperties>'''


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create PPTX drafts from report_package.json.")
    parser.add_argument("--package", required=True)
    parser.add_argument("--preflight")
    parser.add_argument("--template")
    parser.add_argument("--type", required=True, choices=["pptx_report", "chart_review"])
    parser.add_argument("--output", required=True)
    args = parser.parse_args(argv)
    output = write_document(args.package, args.output, args.type, args.preflight, args.template)
    print(str(output))
    return 0


if __name__ == "__main__":
    raise SystemExit(run_cli())
