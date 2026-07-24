"""Generate a visible one-section HWPX demo report.

The smoke test proves the writer contract. This demo creates a small, readable
HWPX draft that a user can open in Hancom HWP and inspect as an actual automated
report page.
"""

from __future__ import annotations

import argparse
import gc
import json
import time
from pathlib import Path
from typing import Any, Dict, List

try:
    from .hwp_com_smoke import read_hwpx_text, wait_for_file_ready
    from .hwp_com_writer import (
        close_hwp,
        create_hwp_object,
        insert_text,
        now,
        parse_bool,
        run_action,
        save_as_hwpx,
        set_visible,
        write_hwp_document,
        write_json,
    )
except ImportError:  # pragma: no cover - allows direct script execution.
    from hwp_com_smoke import read_hwpx_text, wait_for_file_ready  # type: ignore
    from hwp_com_writer import (  # type: ignore
        close_hwp,
        create_hwp_object,
        insert_text,
        now,
        parse_bool,
        run_action,
        save_as_hwpx,
        set_visible,
        write_hwp_document,
        write_json,
    )


EXPECTED_TEXT = [
    "2026년 디지털 서비스 이용 만족도 조사",
    "종합 만족도",
    "만족 응답이 68.4%로 가장 높게 나타났으며",
    "만족",
    "불만족",
]


def run_demo(output_dir: str | Path, visible: bool = False, keep_open_on_error: bool = False) -> Dict[str, Any]:
    """Create a template, package, HWPX output, text preview, and run report."""

    root = Path(output_dir).resolve()
    root.mkdir(parents=True, exist_ok=True)

    template_path = root / "demo_template.hwpx"
    package_path = root / "demo_report_package.json"
    preflight_path = root / "demo_preflight_report.json"
    output_path = root / "demo_hwp_report.hwpx"
    writer_report_path = root / "demo_hwp_writer_report.json"
    preview_path = root / "demo_preview.txt"
    demo_report_path = root / "demo_run_report.json"

    report: Dict[str, Any] = {
        "schema_version": "1.0",
        "status": "started",
        "started_at": now(),
        "finished_at": "",
        "description": "One-section HWPX automation demo for visual inspection.",
        "paths": {
            "template": str(template_path),
            "package": str(package_path),
            "preflight": str(preflight_path),
            "output": str(output_path),
            "writer_report": str(writer_report_path),
            "preview": str(preview_path),
            "demo_report": str(demo_report_path),
        },
        "checks": [],
        "warnings": [],
        "errors": [],
    }

    try:
        create_demo_template(template_path, visible)
        wait_for_file_ready(template_path)
        write_json(package_path, demo_package())
        write_json(preflight_path, demo_preflight())

        write_hwp_document(
            package_path,
            preflight_path,
            template_path,
            output_path,
            visible=visible,
            report_path=writer_report_path,
            keep_open_on_error=keep_open_on_error,
        )

        writer_report = json.loads(writer_report_path.read_text(encoding="utf-8"))
        report["writer_status"] = writer_report.get("status")
        report["writer"] = {
            "sections_written": writer_report.get("sections_written"),
            "tables_written": writer_report.get("tables_written"),
            "text_table_fallbacks": writer_report.get("text_table_fallbacks"),
            "charts_deferred": writer_report.get("charts_deferred"),
            "warnings": writer_report.get("warnings", []),
            "errors": writer_report.get("errors", []),
        }

        output_xml_text = read_hwpx_text(output_path)
        preview_text = build_text_preview(output_xml_text)
        preview_path.write_text(preview_text, encoding="utf-8")

        add_check(report, "output_exists", output_path.exists() and output_path.stat().st_size > 0, str(output_path))
        for text in EXPECTED_TEXT:
            add_check(report, f"contains:{text}", text in output_xml_text, text)
        add_check(report, "body_placeholder_removed", "{{BODY}}" not in output_xml_text, "")
        add_check(report, "writer_status_ready", writer_report.get("status") == "ready", str(writer_report.get("status")))
        add_check(report, "one_section_written", int(writer_report.get("sections_written") or 0) == 1, "")
        add_check(report, "table_object_written", int(writer_report.get("tables_written") or 0) >= 1, "")

        report["status"] = "ready" if not report["errors"] else "failed"
    except Exception as exc:
        report["status"] = "failed"
        report["errors"].append(str(exc))
    finally:
        report["finished_at"] = now()
        write_json(demo_report_path, report)

    return report


def create_demo_template(path: Path, visible: bool) -> None:
    """Create a plain HWPX template that still reads like a report shell."""

    hwp = None
    report: Dict[str, Any] = {
        "stage": "template",
        "action": "create",
        "warnings": [],
        "errors": [],
        "com": {"prog_id": "", "file_path_checker": "", "visible_applied": None, "closed": False},
    }
    try:
        hwp = create_hwp_object(report)
        set_visible(hwp, visible, report)
        run_action(hwp, "FileNew", report, "template")
        insert_text(hwp, "{{REPORT_TITLE}}", report)
        run_action(hwp, "BreakPara", report, "template")
        insert_text(hwp, "프로젝트: {{PROJECT_NAME}}", report)
        run_action(hwp, "BreakPara", report, "template")
        insert_text(hwp, "생성일: {{GENERATED_AT}}", report)
        run_action(hwp, "BreakPara", report, "template")
        run_action(hwp, "BreakPara", report, "template")
        insert_text(hwp, "1. 조사 결과 요약", report)
        run_action(hwp, "BreakPara", report, "template")
        insert_text(hwp, "{{BODY}}", report)
        run_action(hwp, "BreakPara", report, "template")
        insert_text(hwp, "2. QA 요약", report)
        run_action(hwp, "BreakPara", report, "template")
        insert_text(hwp, "{{QA_SUMMARY}}", report)
        save_as_hwpx(hwp, path, report)
    finally:
        if hwp is not None:
            close_hwp(hwp, report)
            del hwp
            gc.collect()
            time.sleep(1.5)


def demo_package() -> Dict[str, Any]:
    return {
        "schema_version": "1.0",
        "meta": {
            "report_title": "2026년 디지털 서비스 이용 만족도 조사",
            "project_name": "보고서 자동화 HWPX 데모",
            "created_at": now(),
            "source_file_name": "demo_survey_result.xlsx",
            "report_profile": "인식도/만족도 조사형",
            "style_profile": "공식 보고서체",
            "decimal_places": 1,
        },
        "sections": [
            {
                "table_key": "DEMO_T001",
                "title": "종합 만족도",
                "narrative_final": (
                    "디지털 서비스의 종합 만족도는 만족 응답이 68.4%로 가장 높게 나타났으며, "
                    "보통은 21.7%, 불만족은 9.9%로 조사되었다. 전반적으로 긍정 응답이 우세하나, "
                    "불만족 응답층에 대해서는 서비스 접근성과 처리 속도 개선 필요성을 함께 검토할 필요가 있다."
                ),
            }
        ],
        "tables": [
            {
                "table_key": "DEMO_T001",
                "title": "종합 만족도",
                "rows": [
                    {"category": "만족", "percent": 68.4, "weighted_n": 342, "raw_n": 338},
                    {"category": "보통", "percent": 21.7, "weighted_n": 109, "raw_n": 112},
                    {"category": "불만족", "percent": 9.9, "weighted_n": 49, "raw_n": 50},
                ],
            }
        ],
        "charts": [
            {"table_key": "DEMO_T001", "category": "만족", "measure": "percent", "value": 68.4, "include_chart": True},
            {"table_key": "DEMO_T001", "category": "보통", "measure": "percent", "value": 21.7, "include_chart": True},
            {"table_key": "DEMO_T001", "category": "불만족", "measure": "percent", "value": 9.9, "include_chart": True},
        ],
        "qa": [
            {"severity": "info", "message": "차트는 HWPX v1 writer에서 직접 삽입하지 않고 삽입 필요 문구로 표시합니다."}
        ],
    }


def demo_preflight() -> Dict[str, Any]:
    return {
        "schema_version": "1.0",
        "status": "ready_with_warnings",
        "section_count": 1,
        "chart_candidate_count": 3,
        "table_count": 1,
        "qa_warning_count": 1,
        "qa_error_count": 0,
        "errors": [],
        "warnings": ["HWPX v1 writer는 차트 직접 삽입을 보류합니다."],
    }


def build_text_preview(raw_text: str) -> str:
    """Extract a compact text preview from the generated HWPX XML."""

    fragments: List[str] = []
    for text in EXPECTED_TEXT + ["[차트 삽입 필요]", "source: DEMO_T001", "차트는 HWPX v1 writer"]:
        if text in raw_text and text not in fragments:
            fragments.append(text)
    return "\n".join(fragments) + "\n"


def add_check(report: Dict[str, Any], name: str, ok: bool, detail: str) -> None:
    report["checks"].append({"name": name, "ok": bool(ok), "detail": detail})
    if not ok:
        report["errors"].append(f"check failed: {name}")


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generate a one-section HWPX demo report.")
    parser.add_argument("--output-dir", default=str(Path("outputs") / "hwp_com_demo"))
    parser.add_argument("--visible", default="false")
    parser.add_argument("--keep-open-on-error", action="store_true")
    args = parser.parse_args(argv)

    report = run_demo(args.output_dir, parse_bool(args.visible), args.keep_open_on_error)
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if report.get("status") == "ready" else 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
