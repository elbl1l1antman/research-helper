"""Probe HWP/HWPX report templates for table-placement planning.

The probe does not modify source templates. HWP files are converted to HWPX
copies under the chosen output directory through the local Hancom COM object,
then every HWPX is inspected as OpenXML.
"""

from __future__ import annotations

import argparse
import collections
import gc
import json
import shutil
import time
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Tuple
from xml.etree import ElementTree as ET

try:
    from .hwp_com_writer import close_hwp, create_hwp_object, now, save_as_hwpx, set_visible, write_json
except ImportError:  # pragma: no cover
    from hwp_com_writer import close_hwp, create_hwp_object, now, save_as_hwpx, set_visible, write_json  # type: ignore


def probe_templates(paths: List[str | Path], output_dir: str | Path, visible: bool = False, table_detail_limit: int = 30) -> Dict[str, Any]:
    root = Path(output_dir).resolve()
    root.mkdir(parents=True, exist_ok=True)
    converted_dir = root / "converted_hwpx"
    converted_dir.mkdir(parents=True, exist_ok=True)

    report: Dict[str, Any] = {
        "schema_version": "1.0",
        "status": "started",
        "started_at": now(),
        "finished_at": "",
        "output_dir": str(root),
        "templates": [],
        "warnings": [],
        "errors": [],
    }

    for source in paths:
        source_path = Path(source)
        item: Dict[str, Any] = {
            "source_path": str(source_path),
            "source_exists": source_path.exists(),
            "source_type": source_path.suffix.lower(),
            "analyzed_path": "",
            "conversion": "",
            "table_count": 0,
            "tables": [],
            "headings": [],
            "warnings": [],
            "errors": [],
        }
        try:
            if not source_path.exists():
                raise FileNotFoundError(str(source_path))
            hwpx_path = ensure_hwpx_copy(source_path, converted_dir, visible, item)
            item["analyzed_path"] = str(hwpx_path)
            analyze_hwpx(hwpx_path, item, table_detail_limit)
        except Exception as exc:
            item["errors"].append(str(exc))
            report["errors"].append(f"{source_path}: {exc}")
        report["templates"].append(item)

    report["status"] = "ready_with_warnings" if report["errors"] else "ready"
    report["finished_at"] = now()
    report_path = root / "hwp_template_probe_report.json"
    summary_path = root / "hwp_template_probe_summary.json"
    report["report_path"] = str(report_path)
    report["summary_path"] = str(summary_path)
    write_json(report_path, report)
    write_json(summary_path, build_summary(report))
    return report


def build_summary(report: Dict[str, Any]) -> List[Dict[str, Any]]:
    summary: List[Dict[str, Any]] = []
    for index, item in enumerate(report.get("templates", []), start=1):
        class_counts = dict(collections.Counter(table.get("classification") for table in item.get("tables", [])))
        examples = []
        for table in item.get("tables", []):
            row_text = " ".join(" ".join(row) for row in table.get("sample_rows", [])[:3])
            if (
                table.get("classification") in {"likely_report_table", "survey_result_table"}
                or "표" in row_text
                or "구분" in row_text
                or "N" in row_text
            ):
                examples.append(
                    {
                        "row_count": table.get("row_count"),
                        "col_count": table.get("col_count"),
                        "classification": table.get("classification"),
                        "before_text": table.get("before_text", [])[:3],
                        "sample_rows": table.get("sample_rows", [])[:5],
                    }
                )
            if len(examples) >= 8:
                break
        summary.append(
            {
                "index": index,
                "source_path": item.get("source_path"),
                "conversion": item.get("conversion"),
                "table_count": item.get("table_count"),
                "class_counts": class_counts,
                "headings": item.get("headings", [])[:12],
                "table_examples": examples,
                "warnings": item.get("warnings", []),
                "errors": item.get("errors", []),
            }
        )
    return summary


def ensure_hwpx_copy(source_path: Path, converted_dir: Path, visible: bool, item: Dict[str, Any]) -> Path:
    safe_name = sanitize_filename(source_path.stem) + ".hwpx"
    target_path = converted_dir / safe_name
    if source_path.suffix.lower() == ".hwpx":
        shutil.copy2(source_path, target_path)
        item["conversion"] = "copied_hwpx"
        return target_path
    if source_path.suffix.lower() == ".hwp":
        convert_hwp_to_hwpx(source_path, target_path, visible)
        item["conversion"] = "converted_hwp_to_hwpx"
        return target_path
    raise ValueError("지원하지 않는 템플릿 확장자입니다: " + source_path.suffix)


def convert_hwp_to_hwpx(source_path: Path, target_path: Path, visible: bool) -> None:
    hwp = None
    report: Dict[str, Any] = {
        "stage": "convert",
        "action": "",
        "warnings": [],
        "errors": [],
        "com": {"prog_id": "", "file_path_checker": "", "visible_applied": None, "closed": False},
    }
    try:
        hwp = create_hwp_object(report)
        set_visible(hwp, visible, report)
        report["action"] = "open"
        if hwp.Open(str(source_path)) is False:
            raise RuntimeError("HWP Open returned False")
        save_as_hwpx(hwp, target_path, report)
    finally:
        if hwp is not None:
            close_hwp(hwp, report)
            del hwp
            gc.collect()
            time.sleep(1.0)


def analyze_hwpx(path: Path, item: Dict[str, Any], table_detail_limit: int = 30) -> None:
    if not zipfile.is_zipfile(path):
        raise ValueError("HWPX zip 패키지가 아닙니다.")

    with zipfile.ZipFile(path) as archive:
        section_names = [name for name in archive.namelist() if name.startswith("Contents/section") and name.endswith(".xml")]
        section_names.sort()
        paragraphs: List[str] = []
        tables: List[Dict[str, Any]] = []
        for section_name in section_names:
            xml = archive.read(section_name)
            root = ET.fromstring(xml)
            walk_section(root, section_name, paragraphs, tables)

    item["headings"] = guess_headings(paragraphs)
    item["table_count"] = len(tables)
    if table_detail_limit <= 0:
        item["tables"] = tables
    else:
        item["tables"] = tables[:table_detail_limit]
        if len(tables) > table_detail_limit:
            item["warnings"].append(f"표 {len(tables)}개 중 앞 {table_detail_limit}개만 상세 기록했습니다.")


def walk_section(node: ET.Element, section_name: str, paragraphs: List[str], tables: List[Dict[str, Any]]) -> None:
    for child in list(node):
        local = local_name(child.tag)
        if local == "p":
            text = normalize_text(text_of(child))
            if text:
                paragraphs.append(text)
        elif local == "tbl":
            table = extract_table(child)
            table["table_index"] = len(tables) + 1
            table["section"] = section_name
            table["before_text"] = paragraphs[-3:]
            tables.append(table)
        walk_section(child, section_name, paragraphs, tables)


def extract_table(tbl: ET.Element) -> Dict[str, Any]:
    rows: List[List[str]] = []
    for tr in descendants_by_local_name(tbl, "tr"):
        row: List[str] = []
        for tc in children_by_local_name(tr, "tc"):
            row.append(normalize_text(text_of(tc)))
        if row:
            rows.append(row)
    max_cols = max((len(row) for row in rows), default=0)
    return {
        "row_count": len(rows),
        "col_count": max_cols,
        "sample_rows": rows[:8],
        "classification": classify_table(rows),
    }


def classify_table(rows: List[List[str]]) -> str:
    flat = " ".join(" ".join(row) for row in rows[:8])
    if any(token in flat for token in ("Base", "BASE", "사례수", "N=", "n=")):
        return "survey_result_table"
    if any(token in flat for token in ("구분", "전체", "계", "%", "점")) and len(rows) >= 3:
        return "likely_report_table"
    if len(rows) <= 3 and max((len(row) for row in rows), default=0) <= 3:
        return "layout_or_summary_box"
    return "unknown"


def guess_headings(paragraphs: List[str]) -> List[str]:
    candidates = []
    for text in paragraphs:
        if len(text) > 80:
            continue
        if any(token in text for token in ("조사 결과", "분석 결과", "응답자", "특성", "만족도", "인식", "실태", "결과")):
            candidates.append(text)
    return candidates[:30]


def descendants_by_local_name(node: ET.Element, name: str) -> List[ET.Element]:
    return [child for child in node.iter() if local_name(child.tag) == name]


def children_by_local_name(node: ET.Element, name: str) -> List[ET.Element]:
    return [child for child in list(node) if local_name(child.tag) == name]


def text_of(node: ET.Element) -> str:
    values: List[str] = []
    for child in node.iter():
        if child.text and local_name(child.tag) in {"t", "p", "run", "sec", "tc"}:
            values.append(child.text)
    return " ".join(values)


def normalize_text(value: str) -> str:
    return " ".join(value.replace("\xa0", " ").split())


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def sanitize_filename(name: str) -> str:
    blocked = '<>:"/\\|?*'
    value = "".join("_" if ch in blocked else ch for ch in name).strip()
    return value[:120] or "template"


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Probe HWP/HWPX report templates.")
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--visible", action="store_true")
    parser.add_argument("--table-detail-limit", type=int, default=30, help="상세 기록할 표 개수입니다. 0이면 전체 표를 기록합니다.")
    parser.add_argument("templates", nargs="+")
    args = parser.parse_args(argv)

    report = probe_templates(args.templates, args.output_dir, args.visible, args.table_detail_limit)
    print(
        json.dumps(
            {
                "status": report.get("status"),
                "report_path": report.get("report_path"),
                "summary_path": report.get("summary_path"),
                "template_count": len(report.get("templates", [])),
                "error_count": len(report.get("errors", [])),
            },
            ensure_ascii=True,
            indent=2,
        )
    )
    return 0 if report.get("status") in {"ready", "ready_with_warnings"} else 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
