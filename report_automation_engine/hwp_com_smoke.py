"""Smoke test for the Hancom HWP COM writer.

This module creates a minimal HWPX template through the local HWP COM object,
builds a small report_package/preflight pair, runs hwp_com_writer, and verifies
that the generated HWPX contains expected Korean text while removing {{BODY}}.
It is intentionally a smoke test, not a full visual regression test.
"""

from __future__ import annotations

import argparse
import gc
import json
import time
import zipfile
from pathlib import Path
from typing import Any, Dict, List

try:
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


def run_smoke(output_dir: str | Path, visible: bool = False, keep_open_on_error: bool = False) -> Dict[str, Any]:
    """Run a local end-to-end smoke test and return a JSON-serializable report."""

    root = Path(output_dir).resolve()
    root.mkdir(parents=True, exist_ok=True)

    template_path = root / "minimal_template.hwpx"
    package_path = root / "report_package.json"
    preflight_path = root / "preflight_report.json"
    output_path = root / "draft_output.hwpx"
    writer_report_path = root / "hwp_writer_report.json"
    smoke_report_path = root / "hwp_com_smoke_report.json"

    report: Dict[str, Any] = {
        "schema_version": "1.0",
        "status": "started",
        "started_at": now(),
        "finished_at": "",
        "paths": {
            "template": str(template_path),
            "package": str(package_path),
            "preflight": str(preflight_path),
            "output": str(output_path),
            "writer_report": str(writer_report_path),
            "smoke_report": str(smoke_report_path),
        },
        "checks": [],
        "warnings": [],
        "errors": [],
    }

    try:
        create_minimal_template(template_path, visible)
        wait_for_file_ready(template_path)
        write_json(package_path, sample_package())
        write_json(preflight_path, sample_preflight())

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

        assert_file_exists(output_path, report, "output_exists")
        assert_hwpx_contains(output_path, ["HWPX COM smoke 문항", "그렇다", "테스트 분석문"], report)
        assert_hwpx_not_contains(output_path, ["{{BODY}}"], report)
        assert_writer_ready(writer_report, report)

        report["status"] = "ready" if not report["errors"] else "failed"
    except Exception as exc:
        report["status"] = "failed"
        report["errors"].append(str(exc))
    finally:
        report["finished_at"] = now()
        write_json(smoke_report_path, report)

    return report


def create_minimal_template(path: Path, visible: bool) -> None:
    """Create a simple HWPX file with the placeholders required by the writer."""

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
        insert_text(hwp, "{{PROJECT_NAME}}", report)
        run_action(hwp, "BreakPara", report, "template")
        insert_text(hwp, "{{GENERATED_AT}}", report)
        run_action(hwp, "BreakPara", report, "template")
        run_action(hwp, "BreakPara", report, "template")
        insert_text(hwp, "{{BODY}}", report)
        run_action(hwp, "BreakPara", report, "template")
        insert_text(hwp, "{{QA_SUMMARY}}", report)
        save_as_hwpx(hwp, path, report)
    finally:
        if hwp is not None:
            close_hwp(hwp, report)
            del hwp
            gc.collect()
            time.sleep(1.5)


def sample_package() -> Dict[str, Any]:
    return {
        "schema_version": "1.0",
        "meta": {
            "report_title": "HWPX COM smoke 테스트 보고서",
            "project_name": "보고서 자동화 알파",
            "created_at": now(),
            "source_file_name": "smoke.xlsx",
            "report_profile": "인식도/만족도 조사형",
            "style_profile": "공식 보고서체",
            "decimal_places": 1,
        },
        "sections": [
            {
                "table_key": "T001",
                "title": "HWPX COM smoke 문항",
                "narrative_final": "테스트 분석문입니다. 긍정 응답은 62.5%로 가장 높게 나타났습니다.",
            }
        ],
        "tables": [
            {
                "table_key": "T001",
                "title": "HWPX COM smoke 문항",
                "rows": [
                    {"category": "그렇다", "percent": 62.5, "weighted_n": 125, "raw_n": 120},
                    {"category": "보통이다", "percent": 25.0, "weighted_n": 50, "raw_n": 48},
                    {"category": "그렇지 않다", "percent": 12.5, "weighted_n": 25, "raw_n": 24},
                ],
            }
        ],
        "charts": [
            {"table_key": "T001", "category": "그렇다", "measure": "percent", "value": 62.5, "include_chart": True}
        ],
        "qa": [],
    }


def sample_preflight() -> Dict[str, Any]:
    return {
        "schema_version": "1.0",
        "status": "ready",
        "section_count": 1,
        "chart_candidate_count": 1,
        "table_count": 1,
        "qa_warning_count": 0,
        "qa_error_count": 0,
        "errors": [],
        "warnings": [],
    }


def assert_file_exists(path: Path, report: Dict[str, Any], name: str) -> None:
    ok = path.exists() and path.stat().st_size > 0
    add_check(report, name, ok, str(path))


def wait_for_file_ready(path: Path, timeout_seconds: float = 10.0) -> None:
    """Wait until the HWPX template is fully written and can be opened as zip."""

    deadline = time.time() + timeout_seconds
    last_error = ""
    while time.time() < deadline:
        try:
            if path.exists() and path.stat().st_size > 0 and zipfile.is_zipfile(path):
                with zipfile.ZipFile(path) as archive:
                    archive.namelist()
                return
        except Exception as exc:
            last_error = str(exc)
        time.sleep(0.3)
    raise RuntimeError(f"HWPX 템플릿 파일 준비를 확인하지 못했습니다: {path} {last_error}")


def assert_writer_ready(writer_report: Dict[str, Any], report: Dict[str, Any]) -> None:
    add_check(report, "writer_status_ready", writer_report.get("status") == "ready", str(writer_report.get("status")))
    add_check(report, "body_placeholder_found", bool(writer_report.get("placeholders", {}).get("body_found")), "")
    add_check(report, "section_written", int(writer_report.get("sections_written") or 0) >= 1, "")
    add_check(report, "table_written", int(writer_report.get("tables_written") or 0) >= 1, "")


def assert_hwpx_contains(path: Path, needles: List[str], report: Dict[str, Any]) -> None:
    text = read_hwpx_text(path)
    for needle in needles:
        add_check(report, f"contains:{needle}", needle in text, needle)


def assert_hwpx_not_contains(path: Path, needles: List[str], report: Dict[str, Any]) -> None:
    text = read_hwpx_text(path)
    for needle in needles:
        add_check(report, f"not_contains:{needle}", needle not in text, needle)


def read_hwpx_text(path: Path) -> str:
    if not zipfile.is_zipfile(path):
        raise ValueError(f"HWPX 파일이 zip 패키지 형식이 아닙니다: {path}")
    chunks: List[str] = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.lower().endswith((".xml", ".rels")):
                chunks.append(archive.read(name).decode("utf-8", errors="ignore"))
    return "\n".join(chunks)


def add_check(report: Dict[str, Any], name: str, ok: bool, detail: str) -> None:
    report["checks"].append({"name": name, "ok": bool(ok), "detail": detail})
    if not ok:
        report["errors"].append(f"check failed: {name}")


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the HWPX COM writer smoke test.")
    parser.add_argument("--output-dir", default=str(Path("outputs") / "hwp_com_smoke"))
    parser.add_argument("--visible", default="false")
    parser.add_argument("--keep-open-on-error", action="store_true")
    args = parser.parse_args(argv)

    report = run_smoke(args.output_dir, parse_bool(args.visible), args.keep_open_on_error)
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if report.get("status") == "ready" else 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
