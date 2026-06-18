"""Template inspection utilities for report automation.

The inspector is intentionally conservative.  It does not judge visual quality;
it checks whether a user-provided HWPX/PPTX file contains the automation fields
needed by the launcher and returns a JSON report that the WinForms app can show.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Sequence
import xml.etree.ElementTree as ET


PLACEHOLDER_RE = re.compile(r"\{\{[A-Z0-9_:]+\}\}")

HWPX_REQUIRED = ["{{BODY}}"]
HWPX_RECOMMENDED = [
    "{{REPORT_TITLE}}",
    "{{PROJECT_NAME}}",
    "{{GENERATED_AT}}",
    "{{TOC}}",
    "{{QA_SUMMARY}}",
]
PPTX_REPORT_REQUIRED = ["{{CHART}}"]
PPTX_REPORT_RECOMMENDED = [
    "{{REPORT_TITLE}}",
    "{{SECTION_TITLE}}",
    "{{NARRATIVE}}",
    "{{TABLE}}",
    "{{CHART}}",
    "{{SOURCE}}",
]
PPTX_CHART_REQUIRED = ["{{CHART_TITLE}}", "{{CHART}}"]
PPTX_CHART_RECOMMENDED = ["{{CHART_NOTE}}", "{{SOURCE}}"]


def inspect_template(template_path: str | Path, template_type: str | None = None) -> Dict[str, object]:
    path = Path(template_path)
    report: Dict[str, object] = {
        "template_path": str(path),
        "template_type": template_type or infer_template_type(path),
        "status": "unsupported",
        "found_placeholders": [],
        "found_shape_names": [],
        "missing_required": [],
        "missing_recommended": [],
        "auto_fix_available": False,
        "detected_layouts": [],
        "warnings": [],
    }

    if not path.exists():
        report["warnings"] = ["템플릿 파일을 찾을 수 없습니다."]
        return report

    suffix = path.suffix.lower()
    if suffix == ".pptx":
        content = read_pptx_text_content(path)
        shape_names = read_pptx_shape_names(path)
        report["found_shape_names"] = sorted(shape_names)
    elif suffix == ".hwpx":
        content = read_zip_text_content(path)
        shape_names = set()
    elif suffix == ".hwp":
        content = ""
        shape_names = set()
        report["warnings"] = ["HWP 바이너리 파일은 v1에서 직접 구조 분석을 지원하지 않습니다. HWPX로 저장한 뒤 검사하세요."]
        report["auto_fix_available"] = False
        return report
    else:
        report["warnings"] = ["지원하지 않는 템플릿 확장자입니다."]
        return report

    placeholders = sorted(set(PLACEHOLDER_RE.findall(content)))
    report["found_placeholders"] = placeholders
    report["detected_layouts"] = detect_layouts(placeholders, shape_names)

    required, recommended = required_fields(str(report["template_type"]))
    missing_required = [field for field in required if field not in placeholders and field.strip("{}") not in shape_names]
    missing_recommended = [field for field in recommended if field not in placeholders and field.strip("{}") not in shape_names]
    report["missing_required"] = missing_required
    report["missing_recommended"] = missing_recommended
    report["auto_fix_available"] = bool(missing_required)

    warnings: List[str] = list(report["warnings"])  # type: ignore[arg-type]
    if missing_required:
        warnings.append("필수 자동화 필드가 부족합니다.")
        report["status"] = "needs_autofix"
    elif missing_recommended:
        warnings.append("권장 자동화 필드가 일부 없습니다.")
        report["status"] = "usable_with_warnings"
    else:
        report["status"] = "ready"
    report["warnings"] = warnings
    return report


def infer_template_type(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".hwpx", ".hwp"}:
        return "hwpx_report"
    if suffix == ".pptx":
        return "pptx_report"
    return "unknown"


def required_fields(template_type: str) -> tuple[Sequence[str], Sequence[str]]:
    if template_type == "chart_review":
        return PPTX_CHART_REQUIRED, PPTX_CHART_RECOMMENDED
    if template_type == "pptx_report":
        return PPTX_REPORT_REQUIRED, PPTX_REPORT_RECOMMENDED
    if template_type == "hwpx_report":
        return HWPX_REQUIRED, HWPX_RECOMMENDED
    return [], []


def read_zip_text_content(path: Path) -> str:
    texts: List[str] = []
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if name.lower().endswith((".xml", ".rels", ".txt")):
                    try:
                        texts.append(zf.read(name).decode("utf-8", errors="ignore"))
                    except Exception:
                        continue
    except zipfile.BadZipFile:
        return path.read_text(encoding="utf-8", errors="ignore")
    return "\n".join(texts)


def read_pptx_text_content(path: Path) -> str:
    raw = read_zip_text_content(path)
    texts = [raw]
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if not name.startswith("ppt/") or not name.endswith(".xml"):
                    continue
                try:
                    root = ET.fromstring(zf.read(name))
                except Exception:
                    continue
                for node in root.iter():
                    if node.tag.endswith("}t") and node.text:
                        texts.append(node.text)
    except Exception:
        pass
    return "\n".join(texts)


def read_pptx_shape_names(path: Path) -> set[str]:
    names: set[str] = set()
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if not name.startswith("ppt/slides/") or not name.endswith(".xml"):
                    continue
                try:
                    root = ET.fromstring(zf.read(name))
                except Exception:
                    continue
                for node in root.iter():
                    if node.tag.endswith("}cNvPr"):
                        value = node.attrib.get("name", "").strip()
                        if value:
                            names.add(value)
    except Exception:
        pass
    return names


def detect_layouts(placeholders: Iterable[str], shape_names: Iterable[str]) -> List[str]:
    tokens = set(placeholders) | set(shape_names)
    layouts: List[str] = []
    if "{{REPORT_TITLE}}" in tokens or "RA_TITLE_SLIDE" in tokens:
        layouts.append("title")
    if "{{BODY}}" in tokens:
        layouts.append("body")
    if "{{NARRATIVE}}" in tokens or "{{TABLE}}" in tokens or "RA_REPORT_SLIDE" in tokens:
        layouts.append("report")
    if "{{CHART}}" in tokens or "RA_CHART" in tokens:
        layouts.append("chart")
    if "{{CHART_NOTE}}" in tokens or "RA_CHART_REVIEW_SLIDE" in tokens:
        layouts.append("chart_review")
    if "{{QA_SUMMARY}}" in tokens:
        layouts.append("qa")
    return layouts


def summarize(report: Dict[str, object]) -> str:
    found = ", ".join(report.get("found_placeholders", []) or ["없음"])  # type: ignore[arg-type]
    missing_required = ", ".join(report.get("missing_required", []) or ["없음"])  # type: ignore[arg-type]
    missing_recommended = ", ".join(report.get("missing_recommended", []) or ["없음"])  # type: ignore[arg-type]
    warnings = "; ".join(report.get("warnings", []) or ["없음"])  # type: ignore[arg-type]
    return (
        f"상태: {report.get('status')}\n"
        f"템플릿 유형: {report.get('template_type')}\n"
        f"발견 필드: {found}\n"
        f"필수 누락: {missing_required}\n"
        f"권장 누락: {missing_recommended}\n"
        f"경고: {warnings}"
    )


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="HWPX/PPTX 보고서 템플릿을 검사합니다.")
    parser.add_argument("--template", required=True, help="검사할 템플릿 파일")
    parser.add_argument("--type", dest="template_type", help="hwpx_report, pptx_report, chart_review")
    parser.add_argument("--output", help="JSON 리포트 저장 경로")
    args = parser.parse_args(argv)

    report = inspect_template(args.template, args.template_type)
    if args.output:
        Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(summarize(report))
    return 0 if report["status"] != "unsupported" else 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
