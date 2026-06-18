"""Create starter templates for report automation users.

PPTX templates are generated as minimal editable OpenXML presentations.  HWPX
generation without Hancom Automation is limited, so v1 creates a placeholder
package that the inspector can read and a companion guide file.
"""

from __future__ import annotations

import argparse
import html
import zipfile
from pathlib import Path
from typing import List


def create_template(template_type: str, output: str | Path) -> Path:
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if template_type == "pptx_report":
        write_pptx_template(
            output_path,
            [
                ["{{REPORT_TITLE}}", "{{PROJECT_NAME}}", "{{GENERATED_AT}}"],
                ["{{SECTION_TITLE}}", "{{NARRATIVE}}", "{{TABLE}}", "{{CHART}}", "{{SOURCE}}"],
                ["{{QA_SUMMARY}}"],
            ],
        )
    elif template_type == "chart_review":
        write_pptx_template(
            output_path,
            [
                ["{{REPORT_TITLE}}", "차트 검토본", "{{GENERATED_AT}}"],
                ["{{CHART_TITLE}}", "{{CHART}}", "{{CHART_NOTE}}", "{{SOURCE}}"],
                ["{{QA_SUMMARY}}"],
            ],
        )
    elif template_type == "hwpx_report":
        write_hwpx_placeholder_package(output_path)
    else:
        raise ValueError(f"지원하지 않는 템플릿 유형입니다: {template_type}")
    return output_path


def write_pptx_template(path: Path, slides: List[List[str]]) -> None:
    slide_ids = []
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(slides)))
        zf.writestr("_rels/.rels", package_rels_xml())
        zf.writestr("docProps/app.xml", app_props_xml(len(slides)))
        zf.writestr("docProps/core.xml", core_props_xml())
        zf.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
        zf.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(len(slides)))
        for index, texts in enumerate(slides, start=1):
            slide_ids.append(index)
            zf.writestr(f"ppt/slides/slide{index}.xml", slide_xml(texts))


def content_types_xml(slide_count: int) -> str:
    slide_overrides = "\n".join(
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
  {slide_overrides}
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
    slide_ids = "\n".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i}"/>' for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>'''


def slide_xml(texts: List[str]) -> str:
    shapes = []
    for idx, text in enumerate(texts, start=1):
        y = 450000 + (idx - 1) * 850000
        height = 620000 if idx == 1 else 520000
        width = 10800000
        shapes.append(text_shape_xml(idx, text, 700000, y, width, height))
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


def text_shape_xml(idx: int, text: str, x: int, y: int, cx: int, cy: int) -> str:
    escaped = html.escape(text)
    name = html.escape(placeholder_to_shape_name(text, idx), quote=True)
    return f'''
<p:sp>
  <p:nvSpPr><p:cNvPr id="{idx + 1}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="2200"/><a:t>{escaped}</a:t></a:r></a:p></p:txBody>
</p:sp>'''


def placeholder_to_shape_name(text: str, idx: int) -> str:
    stripped = text.strip()
    if stripped.startswith("{{") and stripped.endswith("}}"):
        token = stripped.strip("{}").replace(":", "_")
        return "RA_" + token
    return f"RA_PLACEHOLDER_{idx}"


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
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Report Automation Template</dc:title>
  <dc:creator>ReportAutomation</dc:creator>
</cp:coreProperties>'''


def write_hwpx_placeholder_package(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("mimetype", "application/hwp+zip")
        zf.writestr(
            "Contents/report_automation_template.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<reportAutomationTemplate>
  <cover>{{REPORT_TITLE}}</cover>
  <project>{{PROJECT_NAME}}</project>
  <generatedAt>{{GENERATED_AT}}</generatedAt>
  <toc>{{TOC}}</toc>
  <body>{{BODY}}</body>
  <section>{{SECTION_TITLE}}</section>
  <narrative>{{NARRATIVE}}</narrative>
  <table>{{TABLE}}</table>
  <chart>{{CHART}}</chart>
  <source>{{SOURCE}}</source>
  <qa>{{QA_SUMMARY}}</qa>
</reportAutomationTemplate>
""",
        )
    guide_path = path.with_suffix(path.suffix + ".guide.txt")
    guide_path.write_text(
        "HWPX 기본 템플릿 placeholder 패키지를 생성했습니다.\n"
        "한글에서 직접 열리는 완성 HWPX 생성을 위해서는 추후 한글 COM 기반 factory가 필요합니다.\n"
        "최소 사용자 템플릿은 한글 문서 본문 위치에 {{BODY}}를 넣고 .hwpx로 저장하면 됩니다.\n",
        encoding="utf-8",
    )


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="보고서 자동화 기본 템플릿을 생성합니다.")
    parser.add_argument("--type", required=True, choices=["hwpx_report", "pptx_report", "chart_review"])
    parser.add_argument("--output", required=True)
    args = parser.parse_args(argv)
    output = create_template(args.type, args.output)
    print(str(output))
    return 0


if __name__ == "__main__":
    raise SystemExit(run_cli())
