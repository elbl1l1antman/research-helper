"""Recognize automation-ready tables in a user-provided HWP/HWPX template.

This is the user-facing wrapper around hwp_template_probe + template_blueprint.
It returns a compact report the launcher can show after a user selects a report
template: which tables look like survey result tables, which tables are only
style placeholders, and which tables should be treated as layout furniture.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

try:
    from .hwp_template_probe import probe_templates, write_json
    from .template_blueprint import build_blueprint, infer_table_kind, looks_like_layout, looks_like_placeholder
except ImportError:  # pragma: no cover
    from hwp_template_probe import probe_templates, write_json  # type: ignore
    from template_blueprint import build_blueprint, infer_table_kind, looks_like_layout, looks_like_placeholder  # type: ignore


def recognize_template_tables(
    template_path: str | Path,
    output_dir: str | Path,
    visible: bool = False,
    output_path: str | Path | None = None,
) -> Dict[str, Any]:
    """Probe one HWP/HWPX template and create a compact table recognition report."""

    template = Path(template_path).resolve()
    root = Path(output_dir).resolve()
    root.mkdir(parents=True, exist_ok=True)

    recognition_path = Path(output_path).resolve() if output_path else root / "hwp_template_table_recognition.json"
    probe_report = probe_templates([template], root, visible=visible, table_detail_limit=0)
    probe_path = Path(probe_report.get("report_path") or root / "hwp_template_probe_report.json")
    blueprint_path = root / "template_blueprint.json"
    style_report_path = root / "hwp_table_style_report.json"
    style_profile_path = root / "hwp_table_style_profile.json"
    blueprint = build_blueprint(probe_path, blueprint_path)

    template_probe = (probe_report.get("templates") or [{}])[0]
    template_blueprint = (blueprint.get("templates") or [{}])[0]
    result_candidates = normalize_result_candidates(template_blueprint)
    style_candidates = normalize_style_candidates(template_probe, result_candidates)
    layout_tables = normalize_layout_tables(template_probe, result_candidates, style_candidates)

    errors = list(template_probe.get("errors", [])) + list(blueprint.get("errors", []))
    warnings = list(template_probe.get("warnings", [])) + list(blueprint.get("warnings", []))
    if not result_candidates:
        warnings.append("자동화 삽입용 결과표 후보를 찾지 못했습니다. 디자인 전용 템플릿이면 표 스타일 후보를 선택해야 합니다.")
    if template.suffix.lower() == ".hwp" and template_probe.get("conversion"):
        warnings.append("HWP 원본은 수정하지 않고 HWPX 사본으로 변환해 분석했습니다.")
    status = determine_status(errors, result_candidates, style_candidates)
    style_report = build_style_report(template_probe, result_candidates, style_candidates)
    style_profile = build_style_profile(style_report)

    report: Dict[str, Any] = {
        "schema_version": "1.0",
        "status": status,
        "created_at": now(),
        "template_path": str(template),
        "analyzed_path": template_probe.get("analyzed_path", ""),
        "source_type": template.suffix.lower(),
        "conversion": template_probe.get("conversion", ""),
        "probe_report_path": str(probe_path),
        "blueprint_path": str(blueprint_path),
        "style_report_path": str(style_report_path),
        "style_profile_path": str(style_profile_path),
        "summary": {
            "table_count": template_probe.get("table_count", 0),
            "result_candidate_count": len(result_candidates),
            "style_candidate_count": len(style_candidates),
            "layout_table_count": len(layout_tables),
            "recommended_block_id": template_blueprint.get("recommended_block_id", ""),
            "style_profile_status": style_profile.get("status", ""),
        },
        "result_table_candidates": result_candidates,
        "style_table_candidates": style_candidates,
        "layout_tables": layout_tables,
        "warnings": warnings,
        "errors": errors,
        "launcher_next_steps": [
            "사용자가 result_table_candidates 중 문항별 반복 블록으로 쓸 후보를 선택합니다.",
            "실제 결과표가 없고 style_table_candidates만 있으면 해당 표는 서식 원본으로만 사용합니다.",
            "선택된 block_id와 table_index를 HWPX writer의 style_source로 넘깁니다.",
        ],
    }
    report["recognition_report_path"] = str(recognition_path)
    write_json(style_report_path, style_report)
    write_json(style_profile_path, style_profile)
    write_json(recognition_path, report)
    return report


def determine_status(
    errors: List[str],
    result_candidates: List[Dict[str, Any]],
    style_candidates: List[Dict[str, Any]],
) -> str:
    """Return a launcher-friendly state for template table recognition."""

    if errors:
        return "unsupported"
    if result_candidates:
        return "ready"
    if style_candidates:
        return "needs_mapping"
    return "needs_review"


def normalize_result_candidates(template_blueprint: Dict[str, Any]) -> List[Dict[str, Any]]:
    candidates = []
    for block in template_blueprint.get("blocks", []):
        candidates.append(
            {
                "block_id": block.get("block_id", ""),
                "table_index": block.get("table_index", 0),
                "score": block.get("score", 0),
                "title_candidate": block.get("title_candidate", ""),
                "table_kind": block.get("table_kind", ""),
                "section": block.get("section", ""),
                "row_count": block.get("style_source", {}).get("row_count", 0),
                "col_count": block.get("style_source", {}).get("col_count", 0),
                "sequence": block.get("sequence", []),
                "detected_parts": block.get("detected_parts", {}),
                "preview": block.get("preview", {}),
                "recommended_use": "repeat_result_block",
            }
        )
    return candidates


def normalize_style_candidates(template_probe: Dict[str, Any], result_candidates: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    result_indexes = {candidate.get("table_index") for candidate in result_candidates}
    candidates = []
    for table in template_probe.get("tables", []):
        table_index = table.get("table_index")
        if table_index in result_indexes:
            continue
        flat = flatten_table(table)
        table_kind = infer_table_kind(table.get("sample_rows", []), flat)
        if is_style_source_table(table, flat, table_kind):
            candidates.append(
                {
                    "table_index": table_index,
                    "section": table.get("section", ""),
                    "row_count": table.get("row_count", 0),
                    "col_count": table.get("col_count", 0),
                    "classification": table.get("classification", ""),
                    "table_kind": table_kind,
                    "style_summary": table.get("style_summary", {}),
                    "preview": {
                        "before_text": table.get("before_text", [])[-3:],
                        "sample_rows": table.get("sample_rows", [])[:5],
                    },
                    "recommended_use": "style_source_only",
                }
            )
    return candidates[:20]


def is_style_source_table(table: Dict[str, Any], flat: str, table_kind: str) -> bool:
    row_count = int(table.get("row_count") or 0)
    col_count = int(table.get("col_count") or 0)
    if any(token in flat for token in ("목 차", "목차", "표목차", "그림목차", "Contents", "보고서 제목을 입력합니다")):
        return False
    if "표제목 입력" in flat:
        return True
    return table_kind == "placeholder_style_source" and looks_like_placeholder(flat) and row_count >= 4 and col_count >= 2


def normalize_layout_tables(
    template_probe: Dict[str, Any],
    result_candidates: List[Dict[str, Any]],
    style_candidates: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    excluded = {candidate.get("table_index") for candidate in result_candidates}
    excluded.update(candidate.get("table_index") for candidate in style_candidates)
    tables = []
    for table in template_probe.get("tables", []):
        table_index = table.get("table_index")
        if table_index in excluded:
            continue
        flat = flatten_table(table)
        if table.get("classification") == "layout_or_summary_box" or looks_like_layout(flat):
            tables.append(
                {
                    "table_index": table_index,
                    "section": table.get("section", ""),
                    "row_count": table.get("row_count", 0),
                    "col_count": table.get("col_count", 0),
                    "classification": table.get("classification", ""),
                    "preview_text": compact_preview(table),
                    "recommended_use": "ignore_or_keep_layout",
                }
            )
    return tables[:40]


def build_style_report(
    template_probe: Dict[str, Any],
    result_candidates: List[Dict[str, Any]],
    style_candidates: List[Dict[str, Any]],
) -> Dict[str, Any]:
    candidate_indexes = {candidate.get("table_index") for candidate in result_candidates}
    candidate_indexes.update(candidate.get("table_index") for candidate in style_candidates)
    tables = []
    for table in template_probe.get("tables", []):
        if table.get("table_index") not in candidate_indexes:
            continue
        tables.append(
            {
                "table_index": table.get("table_index"),
                "section": table.get("section", ""),
                "row_count": table.get("row_count", 0),
                "col_count": table.get("col_count", 0),
                "classification": table.get("classification", ""),
                "style_summary": table.get("style_summary", {}),
                "preview": {
                    "before_text": table.get("before_text", [])[-3:],
                    "sample_rows": table.get("sample_rows", [])[:5],
                },
            }
        )
    return {
        "schema_version": "1.0",
        "created_at": now(),
        "template_path": template_probe.get("source_path", ""),
        "analyzed_path": template_probe.get("analyzed_path", ""),
        "table_style_count": len(tables),
        "tables": tables,
    }


def build_style_profile(style_report: Dict[str, Any]) -> Dict[str, Any]:
    tables = style_report.get("tables", [])
    source = first_style_source(tables)
    if not source:
        return {
            "schema_version": "1.0",
            "status": "needs_manual_style_selection",
            "created_at": now(),
            "message": "재사용할 표 스타일 후보가 없습니다.",
            "style_source": {},
        }
    style = source.get("style_summary", {})
    return {
        "schema_version": "1.0",
        "status": "ready",
        "created_at": now(),
        "style_source": {
            "table_index": source.get("table_index"),
            "section": source.get("section", ""),
            "row_count": source.get("row_count", 0),
            "col_count": source.get("col_count", 0),
            "classification": source.get("classification", ""),
        },
        "table_style": {
            "table_border_fill_id": style.get("table_border_fill_id", ""),
            "table_border_fill": style.get("table_border_fill", {}),
            "common_cell_border_fill_ids": style.get("common_cell_border_fill_ids", {}),
            "common_cell_border_fills": style.get("common_cell_border_fills", {}),
            "common_char_pr_ids": style.get("common_char_pr_ids", {}),
            "common_char_prs": style.get("common_char_prs", {}),
            "font_height_counts": style.get("font_height_counts", {}),
            "fill_color_counts": style.get("fill_color_counts", {}),
            "sample_cell_margins": style.get("sample_cell_margins", []),
            "sample_cell_sizes": style.get("sample_cell_sizes", []),
            "repeat_header": style.get("repeat_header", ""),
            "cell_spacing": style.get("cell_spacing", ""),
        },
    }


def first_style_source(tables: List[Dict[str, Any]]) -> Dict[str, Any] | None:
    for table in tables:
        if table.get("classification") in {"survey_result_table", "likely_report_table"}:
            return table
    return tables[0] if tables else None


def flatten_table(table: Dict[str, Any]) -> str:
    before = " ".join(str(value) for value in table.get("before_text", []))
    rows = " ".join(" ".join(str(cell) for cell in row) for row in table.get("sample_rows", []))
    return f"{before} {rows}"


def compact_preview(table: Dict[str, Any]) -> str:
    flat = flatten_table(table)
    return " ".join(flat.split())[:180]


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Recognize table candidates in a user-provided HWP/HWPX report template.")
    parser.add_argument("--template", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--output", help="table recognition JSON path")
    parser.add_argument("--visible", action="store_true")
    args = parser.parse_args(argv)

    report = recognize_template_tables(args.template, args.output_dir, args.visible, args.output)
    print(
        json.dumps(
            {
                "status": report.get("status"),
                "recognition_report_path": report.get("recognition_report_path"),
                "result_candidate_count": report.get("summary", {}).get("result_candidate_count"),
                "style_candidate_count": report.get("summary", {}).get("style_candidate_count"),
                "recommended_block_id": report.get("summary", {}).get("recommended_block_id"),
            },
            ensure_ascii=True,
            indent=2,
        )
    )
    return 0 if report.get("status") in {"ready", "needs_mapping", "needs_review"} else 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
