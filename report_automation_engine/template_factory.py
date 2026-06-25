"""Create starter templates for report automation users.

PPTX templates are generated as minimal editable OpenXML presentations.  HWPX
generation without Hancom Automation is limited, so v1 creates a placeholder
package that the inspector can read and a companion guide file.
"""

from __future__ import annotations

import argparse
import html
import json
import subprocess
import tempfile
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
    if try_write_pptx_with_powerpoint(path, slides):
        return
    write_pptx_template_openxml(path, slides)


def try_write_pptx_with_powerpoint(path: Path, slides: List[List[str]]) -> bool:
    script = r'''
param([string]$OutputPath, [string]$SlidesJson)
$ppt = $null
$presentation = $null
function ShapeName([string]$Text, [int]$Index) {
    $trimmed = $Text.Trim()
    if ($trimmed.StartsWith("{{") -and $trimmed.EndsWith("}}")) {
        return "RA_" + $trimmed.Trim("{}").Replace(":", "_")
    }
    return "RA_PLACEHOLDER_$Index"
}
try {
    $slides = Get-Content -LiteralPath $SlidesJson -Raw -Encoding UTF8 | ConvertFrom-Json
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Add()
    foreach ($slideTexts in $slides) {
        $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
        $idx = 1
        foreach ($text in $slideTexts) {
            $top = 36 + (($idx - 1) * 72)
            $height = if ($idx -eq 1) { 54 } else { 44 }
            $shape = $slide.Shapes.AddTextbox(1, 42, $top, 640, $height)
            $shape.Name = ShapeName ([string]$text) $idx
            $shape.TextFrame.TextRange.Text = [string]$text
            $shape.TextFrame.TextRange.Font.Name = "Arial"
            $shape.TextFrame.TextRange.Font.Size = if ($idx -eq 1) { 24 } else { 16 }
            $idx += 1
        }
    }
    $presentation.SaveAs($OutputPath)
    $presentation.Close()
    $ppt.Quit()
    exit 0
} catch {
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    if ($ppt -ne $null) { try { $ppt.Quit() } catch {} }
    Write-Error $_
    exit 1
}
'''
    if not path.suffix.lower() == ".pptx":
        return False
    with tempfile.TemporaryDirectory() as tmp:
        script_path = Path(tmp) / "create_pptx.ps1"
        slides_path = Path(tmp) / "slides.json"
        script_path.write_text(script, encoding="utf-8")
        slides_path.write_text(json.dumps(slides, ensure_ascii=False), encoding="utf-8")
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                str(path),
                str(slides_path),
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )
    return result.returncode == 0 and path.exists()


def write_pptx_template_openxml(path: Path, slides: List[List[str]]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(slides)))
        zf.writestr("_rels/.rels", package_rels_xml())
        zf.writestr("docProps/app.xml", app_props_xml(len(slides)))
        zf.writestr("docProps/core.xml", core_props_xml())
        zf.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
        zf.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(len(slides)))
        zf.writestr("ppt/slideMasters/slideMaster1.xml", slide_master_xml())
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels_xml())
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml())
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels_xml())
        zf.writestr("ppt/theme/theme1.xml", theme_xml())
        for index, texts in enumerate(slides, start=1):
            zf.writestr(f"ppt/slides/slide{index}.xml", slide_xml(texts))
            zf.writestr(f"ppt/slides/_rels/slide{index}.xml.rels", slide_rels_xml())


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
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
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
        f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  {rels}
</Relationships>'''


def presentation_xml(slide_count: int) -> str:
    slide_ids = "\n".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>'''


def slide_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>'''


def slide_master_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>'''


def slide_layout_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>'''


def slide_master_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>'''


def slide_layout_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
</p:sldLayout>'''


def theme_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ReportAutomation">
  <a:themeElements>
    <a:clrScheme name="Office"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="Office"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>'''


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
