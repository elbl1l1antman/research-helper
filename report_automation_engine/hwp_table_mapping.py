"""Build a mapping between report_package sections and HWP/HWPX table blocks.

The template recognizer tells us which HWP/HWPX table block looks reusable.
This module turns that recognition report into a writer-facing contract:
which report section should use which template block, which package table
feeds it, and whether the handoff is ready, warning-only, or blocked.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List


READY_STATES = {"ready", "needs_mapping", "needs_review"}


def build_hwp_table_mapping(
    package_path: str | Path,
    recognition_path: str | Path,
    output_path: str | Path | None = None,
    block_id: str = "",
    style_table_index: int | None = None,
) -> Dict[str, Any]:
    """Create a HWP writer mapping from package + template recognition JSON."""

    package_file = Path(package_path).resolve()
    recognition_file = Path(recognition_path).resolve()
    package = load_json(package_file)
    recognition = load_json(recognition_file)

    warnings: List[Dict[str, Any]] = []
    errors: List[Dict[str, Any]] = []
    sections = list(package.get("sections", []))
    tables = {str(table.get("table_key", "")): table for table in package.get("tables", [])}

    if not sections:
        errors.append(issue("package", "report_package.json에 sections가 없습니다."))

    recognition_status = str(recognition.get("status") or "")
    if recognition_status not in READY_STATES:
        errors.append(issue("template", f"HWP 템플릿 표 인식 상태가 {recognition_status}입니다."))

    explicit_selection_error = False
    selected_block = select_result_block(recognition, block_id)
    if block_id and not selected_block:
        errors.append(issue("template", f"지정한 block_id를 찾지 못했습니다: {block_id}"))
        explicit_selection_error = True

    selected_style = None if selected_block else select_style_table(recognition, style_table_index)
    if style_table_index and not selected_style:
        errors.append(issue("template", f"지정한 style_table_index를 찾지 못했습니다: {style_table_index}"))
        explicit_selection_error = True
    if not selected_block and not selected_style and not explicit_selection_error:
        errors.append(issue("template", "사용할 결과표 블록 또는 표 서식 후보를 찾지 못했습니다."))
    elif not selected_block and selected_style:
        warnings.append(issue("template", "실제 결과표 블록이 없어 표 서식 후보를 반복 원본으로 사용합니다."))

    section_mappings = []
    for index, section in enumerate(sections, start=1):
        table_key = str(section.get("table_key") or "")
        if not table_key:
            errors.append(issue("package", f"{index}번째 section의 table_key가 비어 있습니다."))
        table = tables.get(table_key)
        if table is None:
            errors.append(issue(table_key or "package", "section에 대응되는 삽입표 데이터가 없습니다."))

        source_rows = list((table or {}).get("rows", []))
        if table is not None and not source_rows:
            errors.append(issue(table_key, "삽입표 rows가 비어 있습니다."))

        template_source = selected_block or selected_style or {}
        section_mappings.append(
            {
                "order": index,
                "table_key": table_key,
                "section_title": str(section.get("title") or table_key or f"문항 {index}"),
                "narrative_source": str(section.get("narrative_source") or ""),
                "has_narrative": bool(str(section.get("narrative_final") or "").strip()),
                "package_table": {
                    "exists": table is not None,
                    "row_count": len(source_rows),
                    "source_columns": infer_package_columns(source_rows),
                },
                "template_block": {
                    "block_id": str(template_source.get("block_id") or ""),
                    "table_index": template_source.get("table_index", 0),
                    "table_kind": str(template_source.get("table_kind") or ""),
                    "row_count": template_source.get("row_count", 0),
                    "col_count": template_source.get("col_count", 0),
                    "render_sequence": list(template_source.get("sequence", [])),
                    "recommended_use": str(template_source.get("recommended_use") or ""),
                },
                "writer_contract": {
                    "mode": "repeat_result_block" if selected_block else "repeat_style_table",
                    "requires": ["section_title", "narrative", "result_table"],
                    "defer_chart": True,
                    "preserve_template_style": True,
                },
            }
        )

    if len(section_mappings) > 1 and (selected_block or selected_style):
        warnings.append(issue("template", "하나의 템플릿 표 블록을 여러 section에 반복 적용합니다."))

    for section in sections:
        if not str(section.get("narrative_final") or "").strip():
            errors.append(issue(str(section.get("table_key") or "package"), "최종 분석문이 없습니다."))

    status = "blocked" if errors else "ready_with_warnings" if warnings else "ready"
    mapping = {
        "schema_version": "1.0",
        "status": status,
        "created_at": now(),
        "package_path": str(package_file),
        "recognition_path": str(recognition_file),
        "template_path": recognition.get("template_path", ""),
        "analyzed_path": recognition.get("analyzed_path", ""),
        "selected": {
            "block_id": str((selected_block or {}).get("block_id") or ""),
            "style_table_index": (selected_style or {}).get("table_index", 0),
            "selection_source": selection_source(block_id, style_table_index, selected_block, selected_style),
        },
        "summary": {
            "section_count": len(sections),
            "mapped_section_count": len(section_mappings),
            "package_table_count": len(tables),
            "template_result_candidate_count": int(recognition.get("summary", {}).get("result_candidate_count") or 0),
            "template_style_candidate_count": int(recognition.get("summary", {}).get("style_candidate_count") or 0),
            "warning_count": len(warnings),
            "error_count": len(errors),
        },
        "section_mappings": section_mappings,
        "warnings": warnings,
        "errors": errors,
        "writer_next_steps": [
            "HWPX writer는 section_mappings 순서대로 템플릿 블록을 복제하거나 동일 서식의 표를 생성합니다.",
            "차트는 v1에서 직접 삽입하지 않고 차트 삽입 필요 표시 또는 후속 EMF 삽입 단계로 넘깁니다.",
            "blocked 상태면 문서 생성을 중단하고 errors를 먼저 해결합니다.",
        ],
    }

    if output_path:
        write_json(output_path, mapping)
    return mapping


def select_result_block(recognition: Dict[str, Any], block_id: str) -> Dict[str, Any] | None:
    candidates = list(recognition.get("result_table_candidates", []))
    if block_id:
        for candidate in candidates:
            if str(candidate.get("block_id") or "") == block_id:
                return candidate
        return None
    recommended = str(recognition.get("summary", {}).get("recommended_block_id") or "")
    for candidate in candidates:
        if str(candidate.get("block_id") or "") == recommended:
            return candidate
    return candidates[0] if candidates else None


def select_style_table(recognition: Dict[str, Any], style_table_index: int | None) -> Dict[str, Any] | None:
    candidates = list(recognition.get("style_table_candidates", []))
    if style_table_index:
        for candidate in candidates:
            if int(candidate.get("table_index") or 0) == style_table_index:
                return candidate
        return None
    return candidates[0] if candidates else None


def selection_source(
    block_id: str,
    style_table_index: int | None,
    selected_block: Dict[str, Any] | None,
    selected_style: Dict[str, Any] | None,
) -> str:
    if block_id and selected_block:
        return "user_block_id"
    if style_table_index and selected_style:
        return "user_style_table_index"
    if selected_block:
        return "recommended_block"
    if selected_style:
        return "first_style_candidate"
    return "none"


def infer_package_columns(rows: List[Dict[str, Any]]) -> List[str]:
    columns: List[str] = []
    for row in rows:
        for key in row.keys():
            if key not in columns:
                columns.append(str(key))
    return columns


def issue(issue_type: str, message: str) -> Dict[str, str]:
    return {"type": issue_type, "message": message}


def load_json(path: str | Path) -> Dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_json(path: str | Path, value: Dict[str, Any]) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build HWP table mapping from report_package and template recognition JSON.")
    parser.add_argument("--package", required=True)
    parser.add_argument("--recognition", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--block-id", default="")
    parser.add_argument("--style-table-index", type=int)
    args = parser.parse_args(argv)

    mapping = build_hwp_table_mapping(
        args.package,
        args.recognition,
        args.output,
        args.block_id,
        args.style_table_index,
    )
    print(
        json.dumps(
            {
                "status": mapping.get("status"),
                "output": str(Path(args.output).resolve()),
                "mapped_section_count": mapping.get("summary", {}).get("mapped_section_count"),
                "selected": mapping.get("selected"),
                "warning_count": mapping.get("summary", {}).get("warning_count"),
                "error_count": mapping.get("summary", {}).get("error_count"),
            },
            ensure_ascii=True,
            indent=2,
        )
    )
    return 0 if mapping.get("status") in {"ready", "ready_with_warnings"} else 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
