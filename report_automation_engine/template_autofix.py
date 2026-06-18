"""Create *_template_ready copies by adding minimum automation fields.

The autofix step is intentionally non-destructive: it never edits the user's
original file.  PPTX files are patched by adding a placeholder text box to the
first slide.  HWPX files receive an extra XML part that is detectable by the
inspector; visual insertion into the HWPX body is deferred to the Hancom COM
writer implementation.
"""

from __future__ import annotations

import argparse
import shutil
import zipfile
from pathlib import Path
from typing import List

try:
    from .template_inspector import inspect_template
except ImportError:
    from template_inspector import inspect_template


def autofix_template(template: str | Path, template_type: str | None = None, output: str | Path | None = None) -> Path:
    source = Path(template)
    if not source.exists():
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {source}")

    target = Path(output) if output else source.with_name(source.stem + "_template_ready" + source.suffix)
    shutil.copy2(source, target)

    suffix = source.suffix.lower()
    report = inspect_template(source, template_type)
    missing = list(report.get("missing_required", [])) + list(report.get("missing_recommended", []))
    if not missing:
        return target

    if suffix == ".pptx":
        patch_pptx(target, missing)
    elif suffix == ".hwpx":
        patch_hwpx(target, missing)
    elif suffix == ".hwp":
        raise ValueError("HWP 바이너리는 v1 자동 보정을 지원하지 않습니다. HWPX로 저장한 뒤 보정하세요.")
    else:
        raise ValueError(f"지원하지 않는 템플릿 확장자입니다: {suffix}")
    return target


def patch_pptx(path: Path, missing: List[str]) -> None:
    placeholders = missing or ["{{CHART}}"]
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        slide_name = next((name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml")), "")
        for name in names:
            data = zin.read(name)
            if name == slide_name:
                text = data.decode("utf-8", errors="ignore")
                text = insert_pptx_placeholder_shapes(text, placeholders)
                data = text.encode("utf-8")
            zout.writestr(name, data)
    temp_path.replace(path)


def insert_pptx_placeholder_shapes(slide_xml: str, placeholders: List[str]) -> str:
    insertion = "\n".join(
        f'''
<p:sp>
  <p:nvSpPr><p:cNvPr id="{900 + i}" name="RA_AUTOFIX_{i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="750000" y="{500000 + i * 450000}"/><a:ext cx="6000000" cy="360000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="1800"/><a:t>{placeholder}</a:t></a:r></a:p></p:txBody>
</p:sp>'''
        for i, placeholder in enumerate(placeholders, start=1)
    )
    marker = "</p:spTree>"
    if marker in slide_xml:
        return slide_xml.replace(marker, insertion + marker, 1)
    return slide_xml + insertion


def patch_hwpx(path: Path, missing: List[str]) -> None:
    placeholders = missing or ["{{BODY}}"]
    with zipfile.ZipFile(path, "a", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "Contents/report_automation_autofix.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "<reportAutomationAutofix>\n"
            + "\n".join(f"  <placeholder>{placeholder}</placeholder>" for placeholder in placeholders)
            + "\n</reportAutomationAutofix>\n",
        )


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="사용자 템플릿의 자동 보정 사본을 생성합니다.")
    parser.add_argument("--template", required=True)
    parser.add_argument("--type", dest="template_type")
    parser.add_argument("--output")
    args = parser.parse_args(argv)
    output = autofix_template(args.template, args.template_type, args.output)
    print(str(output))
    return 0


if __name__ == "__main__":
    raise SystemExit(run_cli())
