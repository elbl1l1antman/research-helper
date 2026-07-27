"""Create HWPX drafts through the local Hancom HWP COM automation API.

이 모듈은 HWPX를 직접 XML로 조립하지 않고, 사용자의 Windows PC에 설치된
아래한글을 실행해 템플릿 사본에 본문과 표를 입력한다. rhwp 기반 writer는
장기 후보로 두고, 알파 단계의 실사용 writer는 아래한글 COM을 우선한다.
"""

from __future__ import annotations

import argparse
import json
import platform
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List


BODY_PLACEHOLDER = "{{BODY}}"
REPORT_PLACEHOLDERS = {
    "{{REPORT_TITLE}}": "report_title",
    "{{PROJECT_NAME}}": "project_name",
    "{{GENERATED_AT}}": "created_at",
    "{{QA_SUMMARY}}": "qa_summary",
}
TABLE_COLUMNS = ["항목", "비율", "가중 N", "원 N"]


class HwpWriterError(RuntimeError):
    """Writer failure with a stage/action pair for the JSON report."""

    def __init__(self, stage: str, action: str, message: str) -> None:
        super().__init__(message)
        self.stage = stage
        self.action = action


def write_hwp_document(
    package_path: str | Path,
    preflight_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
    visible: bool = False,
    report_path: str | Path | None = None,
    keep_open_on_error: bool = False,
    max_sections: int | None = None,
    render_plan_path: str | Path | None = None,
    dry_run: bool = False,
    table_style_profile_path: str | Path | None = None,
    keep_open_after_save: bool = False,
) -> Path:
    """Write a report draft and always write a companion JSON report."""

    package_file = Path(package_path).resolve()
    preflight_file = Path(preflight_path).resolve()
    template_file = Path(template_path).resolve()
    output_file = Path(output_path).resolve()
    writer_report = new_report(package_file, preflight_file, template_file, output_file, visible)
    writer_report["keep_open_after_save"] = keep_open_after_save
    report_file = Path(report_path).resolve() if report_path else output_file.with_name(output_file.stem + "_hwp_writer_report.json")
    render_plan_file = Path(render_plan_path).resolve() if render_plan_path else output_file.with_name(output_file.stem + "_hwp_render_plan.json")
    table_style_profile_file = Path(table_style_profile_path).resolve() if table_style_profile_path else None
    hwp = None

    try:
        package = load_json(package_file)
        preflight = load_json(preflight_file)
        validate_preflight(preflight)
        table_style_profile = load_table_style_profile(table_style_profile_file, writer_report)
        render_plan = build_render_plan(package, max_sections, table_style_profile)
        writer_report["render_plan_path"] = str(render_plan_file)
        writer_report["section_count_total"] = render_plan["section_count_total"]
        writer_report["section_count_selected"] = render_plan["section_count_selected"]
        writer_report["table_style_profile"] = render_plan["table_style_profile"]
        writer_report["table_style_apply_plan"] = render_plan["table_style_apply_plan"]
        write_json(render_plan_file, render_plan)
        write_json(report_file, writer_report)

        if dry_run:
            writer_report["status"] = "ready"
            writer_report["dry_run"] = True
            writer_report["finished_at"] = now()
            write_json(report_file, writer_report)
            return render_plan_file

        validate_files(template_file, output_file)

        output_file.parent.mkdir(parents=True, exist_ok=True)
        if template_file.resolve() == output_file.resolve():
            raise HwpWriterError("validate", "output", "원본 템플릿과 출력 경로가 같습니다. 원본 보호를 위해 중단합니다.")
        shutil.copy2(template_file, output_file)
        writer_report["template_copied"] = True
        write_json(report_file, writer_report)

        writer_report["stage"] = "com"
        writer_report["action"] = "create_hwp_object"
        write_json(report_file, writer_report)
        hwp = create_hwp_object(writer_report, report_file)
        set_visible(hwp, visible, writer_report, report_file)
        write_json(report_file, writer_report)
        open_document(hwp, output_file, writer_report, report_file)
        write_json(report_file, writer_report)

        replace_header_placeholders(hwp, package, writer_report)
        if not find_placeholder(hwp, BODY_PLACEHOLDER):
            raise HwpWriterError("template", "find_body", "{{BODY}} placeholder를 문서 본문에서 찾지 못했습니다.")
        writer_report["placeholders"]["body_found"] = True
        run_action(hwp, "Delete", writer_report, "template")
        write_json(report_file, writer_report)

        write_body(hwp, package, writer_report, max_sections, table_style_profile)
        write_json(report_file, writer_report)
        save_as_hwpx(hwp, output_file, writer_report, report_file)
        writer_report["status"] = "ready"
        writer_report["finished_at"] = now()
        write_json(report_file, writer_report)
        return output_file
    except HwpWriterError as exc:
        writer_report["status"] = "failed"
        writer_report["stage"] = exc.stage
        writer_report["action"] = exc.action
        writer_report["errors"].append(str(exc))
        writer_report["finished_at"] = now()
        raise
    except Exception as exc:
        writer_report["status"] = "failed"
        writer_report["stage"] = writer_report.get("stage") or "unknown"
        writer_report["action"] = writer_report.get("action") or "unknown"
        writer_report["errors"].append(str(exc))
        writer_report["finished_at"] = now()
        raise
    finally:
        write_json(report_file, writer_report)
        if hwp is not None and should_close_hwp(writer_report, keep_open_on_error, keep_open_after_save):
            close_hwp(hwp, writer_report)
            write_json(report_file, writer_report)
        elif hwp is not None:
            writer_report["com"]["closed"] = False
            writer_report["warnings"].append("HWP COM document was intentionally left open.")
            write_json(report_file, writer_report)


def check_environment(report_path: str | Path | None = None, visible: bool = False) -> Dict[str, Any]:
    """Check whether pywin32 and the local HWP COM object are available."""

    report = {
        "schema_version": "1.0",
        "status": "started",
        "stage": "environment",
        "action": "check_environment",
        "started_at": now(),
        "finished_at": "",
        "platform": platform.platform(),
        "python": sys.version,
        "visible": visible,
        "com": {
            "prog_id": "",
            "current_prog_id": "",
            "file_path_checker": "",
            "visible_applied": None,
            "closed": False,
            "steps": [],
        },
        "warnings": [],
        "errors": [],
    }
    hwp = None
    try:
        if platform.system().lower() != "windows":
            raise HwpWriterError("environment", "platform", "아래한글 COM writer는 Windows에서만 실행할 수 있습니다.")
        checkpoint_file = Path(report_path).resolve() if report_path else None
        write_checkpoint(report, checkpoint_file)
        hwp = create_hwp_object(report, checkpoint_file)
        set_visible(hwp, visible, report, checkpoint_file)
        write_checkpoint(report, checkpoint_file)
        report["status"] = "ready"
        report["finished_at"] = now()
        write_checkpoint(report, checkpoint_file)
        return report
    except HwpWriterError as exc:
        report["status"] = "failed"
        report["stage"] = exc.stage
        report["action"] = exc.action
        report["errors"].append(str(exc))
        report["finished_at"] = now()
        return report
    except Exception as exc:
        report["status"] = "failed"
        report["errors"].append(str(exc))
        report["finished_at"] = now()
        return report
    finally:
        if hwp is not None:
            close_hwp(hwp, report)
        if report_path:
            write_json(report_path, report)


def validate_preflight(preflight: Dict[str, Any]) -> None:
    if preflight.get("status") == "blocked":
        raise HwpWriterError("validate", "preflight", "preflight status가 blocked입니다. 오류를 먼저 해결하세요.")


def validate_files(template_file: Path, output_file: Path) -> None:
    if platform.system().lower() != "windows":
        raise HwpWriterError("validate", "platform", "아래한글 COM writer는 Windows에서만 실행할 수 있습니다.")
    if not template_file.exists():
        raise HwpWriterError("validate", "template", f"템플릿 파일을 찾을 수 없습니다: {template_file}")
    if template_file.suffix.lower() not in {".hwpx", ".hwp"}:
        raise HwpWriterError("validate", "template", "HWPX/HWP 템플릿만 지원합니다.")
    if output_file.suffix.lower() not in {".hwpx", ".hwp"}:
        raise HwpWriterError("validate", "output", "출력 파일 확장자는 .hwpx 또는 .hwp여야 합니다.")


def create_hwp_object(report: Dict[str, Any], checkpoint_path: Path | None = None):
    report["stage"] = "com"
    report["action"] = "create_object"
    record_com_step(report, "import_win32com", "started")
    write_checkpoint(report, checkpoint_path)
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        record_com_step(report, "import_win32com", "failed", str(exc))
        write_checkpoint(report, checkpoint_path)
        raise HwpWriterError("com", "import_win32com", "pywin32(win32com)를 불러오지 못했습니다. pywin32 설치가 필요합니다.") from exc
    record_com_step(report, "import_win32com", "ready")
    write_checkpoint(report, checkpoint_path)

    last_error = None
    for prog_id in ("HWPFrame.HwpObject", "HwpFrame.HwpObject.2"):
        try:
            report["stage"] = "com"
            report["action"] = "dispatch"
            report["com"]["current_prog_id"] = prog_id
            record_com_step(report, "dispatch", "started", prog_id=prog_id)
            write_checkpoint(report, checkpoint_path)
            hwp = win32com.client.gencache.EnsureDispatch(prog_id)
            report["com"]["prog_id"] = prog_id
            record_com_step(report, "dispatch", "ready", prog_id=prog_id)
            write_checkpoint(report, checkpoint_path)
            register_file_path_checker(hwp, report, checkpoint_path)
            return hwp
        except Exception as exc:
            last_error = exc
            record_com_step(report, "dispatch", "failed", str(exc), prog_id=prog_id)
            write_checkpoint(report, checkpoint_path)
    raise HwpWriterError("com", "create_object", f"아래한글 COM 객체를 생성하지 못했습니다: {last_error}")


def register_file_path_checker(hwp, report: Dict[str, Any], checkpoint_path: Path | None = None) -> None:
    # 보안 모듈 등록은 설치 환경별로 실패할 수 있다. 실패해도 Open 단계에서 다시 명확한 오류가 난다.
    for module in ("FilePathCheckerModule", "FilePathCheckDLL"):
        try:
            report["stage"] = "com"
            report["action"] = "register_file_path_checker"
            record_com_step(report, "register_file_path_checker", "started", module=module)
            write_checkpoint(report, checkpoint_path)
            hwp.RegisterModule("FilePathCheckDLL", module)
            report["com"]["file_path_checker"] = module
            record_com_step(report, "register_file_path_checker", "ready", module=module)
            write_checkpoint(report, checkpoint_path)
            return
        except Exception as exc:
            record_com_step(report, "register_file_path_checker", "failed", str(exc), module=module)
            write_checkpoint(report, checkpoint_path)
            continue
    report["warnings"].append("아래한글 FilePathCheck 보안 모듈 등록을 확인하지 못했습니다.")
    write_checkpoint(report, checkpoint_path)


def record_com_step(report: Dict[str, Any], name: str, status: str, detail: str = "", **extra: Any) -> None:
    step = {"at": now(), "name": name, "status": status}
    if detail:
        step["detail"] = detail
    step.update(extra)
    report.setdefault("com", {}).setdefault("steps", []).append(step)


def write_checkpoint(report: Dict[str, Any], path: Path | None) -> None:
    if path:
        write_json(path, report)


def set_visible(hwp, visible: bool, report: Dict[str, Any], checkpoint_path: Path | None = None) -> None:
    report["stage"] = "com"
    report["action"] = "set_visible"
    record_com_step(report, "set_visible", "started", visible=visible)
    write_checkpoint(report, checkpoint_path)
    try:
        hwp.XHwpWindows.Item(0).Visible = bool(visible)
        report["com"]["visible_applied"] = bool(visible)
        record_com_step(report, "set_visible", "ready", "XHwpWindows", visible=visible)
        write_checkpoint(report, checkpoint_path)
    except Exception:
        try:
            hwp.Visible = bool(visible)
            report["com"]["visible_applied"] = bool(visible)
            record_com_step(report, "set_visible", "ready", "Visible", visible=visible)
            write_checkpoint(report, checkpoint_path)
        except Exception:
            report["warnings"].append("아래한글 창 표시 옵션을 적용하지 못했습니다.")
            record_com_step(report, "set_visible", "failed", visible=visible)
            write_checkpoint(report, checkpoint_path)


def open_document(hwp, path: Path, report: Dict[str, Any], checkpoint_path: Path | None = None) -> None:
    report["stage"] = "document"
    report["action"] = "open"
    record_com_step(report, "open_document", "started", path=str(path))
    write_checkpoint(report, checkpoint_path)
    attempts = [
        lambda: hwp.Open(str(path), "HWPX", "forceopen:true") if path.suffix.lower() == ".hwpx" else hwp.Open(str(path)),
        lambda: hwp.Open(str(path)),
    ]
    last_error = None
    for attempt in attempts:
        try:
            result = attempt()
            if result is False:
                last_error = "Open returned False"
                record_com_step(report, "open_document", "failed", str(last_error), path=str(path))
                write_checkpoint(report, checkpoint_path)
                continue
            report["document_opened"] = True
            record_com_step(report, "open_document", "ready", path=str(path))
            write_checkpoint(report, checkpoint_path)
            return
        except Exception as exc:
            last_error = exc
            record_com_step(report, "open_document", "failed", str(exc), path=str(path))
            write_checkpoint(report, checkpoint_path)
    raise HwpWriterError("document", "open", f"템플릿 사본을 아래한글로 열지 못했습니다: {last_error}")


def replace_header_placeholders(hwp, package: Dict[str, Any], report: Dict[str, Any]) -> None:
    meta = package.get("meta", {})
    values = {
        "report_title": str(meta.get("report_title") or meta.get("source_file_name") or "보고서 초안"),
        "project_name": str(meta.get("project_name") or meta.get("report_profile") or ""),
        "created_at": str(meta.get("created_at") or now()),
        "qa_summary": qa_summary(package),
    }
    for placeholder, value_key in REPORT_PLACEHOLDERS.items():
        if find_placeholder(hwp, placeholder):
            run_action(hwp, "Delete", report, "template")
            insert_text(hwp, values[value_key], report)
            report["placeholders"]["replaced"].append(placeholder)


def find_placeholder(hwp, text: str) -> bool:
    try:
        hwp.HAction.Run("MoveDocBegin")
    except Exception:
        pass
    try:
        params = hwp.HParameterSet.HFindReplace
        hwp.HAction.GetDefault("RepeatFind", params.HSet)
        params.FindString = text
        params.IgnoreMessage = 1
        try:
            params.Direction = hwp.FindDir("Forward")
        except Exception:
            pass
        return bool(hwp.HAction.Execute("RepeatFind", params.HSet))
    except Exception:
        return False


def write_body(
    hwp,
    package: Dict[str, Any],
    report: Dict[str, Any],
    max_sections: int | None = None,
    table_style_profile: Dict[str, Any] | None = None,
) -> None:
    tables_by_key = {str(table.get("table_key", "")): table for table in package.get("tables", [])}
    charts_by_key = group_charts(package.get("charts", []))
    sections = select_sections(package, max_sections)
    for index, section in enumerate(sections, start=1):
        key = str(section.get("table_key", ""))
        title = str(section.get("title") or key or f"문항 {index}")
        narrative = str(section.get("narrative_final") or "")

        insert_text(hwp, title, report)
        run_action(hwp, "BreakPara", report, "body")
        insert_text(hwp, narrative, report)
        run_action(hwp, "BreakPara", report, "body")
        run_action(hwp, "BreakPara", report, "body")

        table = tables_by_key.get(key)
        if table:
            insert_text(hwp, str(table.get("title") or title), report)
            run_action(hwp, "BreakPara", report, "body")
            if not insert_hwp_table(hwp, table_rows_for_hwp(table), report, table_style_profile):
                insert_text_table(hwp, table_rows_for_hwp(table), report)
            run_action(hwp, "BreakPara", report, "body")
        else:
            report["warnings"].append(f"삽입표 데이터 없음: {key}")

        if charts_by_key.get(key):
            insert_text(hwp, f"[차트 삽입 필요] {title}", report)
            run_action(hwp, "BreakPara", report, "body")
            report["charts_deferred"] += 1

        insert_text(hwp, f"source: {key}", report)
        run_action(hwp, "BreakPara", report, "body")
        run_action(hwp, "BreakPara", report, "body")
        report["sections_written"] += 1


def select_sections(package: Dict[str, Any], max_sections: int | None = None) -> List[Dict[str, Any]]:
    sections = list(package.get("sections", []))
    if max_sections and max_sections > 0:
        return sections[:max_sections]
    return sections


def build_render_plan(
    package: Dict[str, Any],
    max_sections: int | None = None,
    table_style_profile: Dict[str, Any] | None = None,
) -> Dict[str, Any]:
    """Create a COM-independent preview of what the HWP writer will insert."""

    tables_by_key = {str(table.get("table_key", "")): table for table in package.get("tables", [])}
    charts_by_key = group_charts(package.get("charts", []))
    all_sections = list(package.get("sections", []))
    selected_sections = select_sections(package, max_sections)
    plan_sections: List[Dict[str, Any]] = []

    for index, section in enumerate(selected_sections, start=1):
        key = str(section.get("table_key", ""))
        title = str(section.get("title") or key or f"문항 {index}")
        narrative = str(section.get("narrative_final") or "")
        table = tables_by_key.get(key)
        table_rows = table_rows_for_hwp(table) if table else []
        chart_rows = charts_by_key.get(key, [])
        plan_sections.append(
            {
                "index": index,
                "table_key": key,
                "title": title,
                "narrative": narrative,
                "narrative_length": len(narrative),
                "table_title": str(table.get("title") or title) if table else "",
                "table_row_count": max(len(table_rows) - 1, 0),
                "table_preview_rows": table_rows[:6],
                "chart_deferred": bool(chart_rows),
                "chart_candidate_count": len(chart_rows),
                "source": key,
            }
        )

    return {
        "schema_version": "1.0",
        "created_at": now(),
        "writer": "hwp_com_writer",
        "max_sections": max_sections if max_sections and max_sections > 0 else None,
        "section_count_total": len(all_sections),
        "section_count_selected": len(selected_sections),
        "table_count_total": len(package.get("tables", [])),
        "chart_count_total": len(package.get("charts", [])),
        "qa_count_total": len(package.get("qa", [])),
        "table_style_profile": summarize_table_style_profile(table_style_profile),
        "table_style_apply_plan": build_table_style_apply_plan(table_style_profile),
        "sections": plan_sections,
    }


def insert_hwp_table(
    hwp,
    rows: List[List[str]],
    report: Dict[str, Any],
    table_style_profile: Dict[str, Any] | None = None,
) -> bool:
    """Try to create a real HWP table. Fall back to text table when COM differs."""

    if not rows:
        return False
    report["stage"] = "table"
    report["action"] = "TableCreate"
    try:
        apply_table_style_before_create(hwp, table_style_profile, report)
        params = hwp.HParameterSet.HTableCreation
        hwp.HAction.GetDefault("TableCreate", params.HSet)
        params.Rows = len(rows)
        params.Cols = len(rows[0])
        try:
            params.WidthType = 2
            params.HeightType = 1
        except Exception:
            pass
        if not hwp.HAction.Execute("TableCreate", params.HSet):
            return False
        for row_idx, row in enumerate(rows):
            for col_idx, value in enumerate(row):
                insert_text(hwp, value, report)
                if not (row_idx == len(rows) - 1 and col_idx == len(row) - 1):
                    run_action(hwp, "TableRightCell", report, "table")
        report["tables_written"] += 1
        try:
            hwp.HAction.Run("MoveRight")
        except Exception:
            pass
        return True
    except Exception as exc:
        report["warnings"].append(f"HWP 표 객체 생성 실패, 텍스트 표로 대체합니다: {exc}")
        return False


def load_table_style_profile(path: Path | None, report: Dict[str, Any]) -> Dict[str, Any] | None:
    if path is None:
        return None
    if not path.exists():
        raise HwpWriterError("validate", "table_style_profile", f"표 스타일 profile 파일을 찾을 수 없습니다: {path}")
    profile = load_json(path)
    status = str(profile.get("status", ""))
    if status and status != "ready":
        report["warnings"].append(f"표 스타일 profile 상태가 ready가 아닙니다: {status}")
    report["table_style_profile_path"] = str(path)
    return profile


def summarize_table_style_profile(profile: Dict[str, Any] | None) -> Dict[str, Any]:
    if not profile:
        return {"loaded": False}
    table_style = profile.get("table_style", {})
    font_height = dominant_font_height(profile)
    apply_plan = build_table_style_apply_plan(profile)
    return {
        "loaded": True,
        "status": profile.get("status", ""),
        "source": profile.get("style_source", {}),
        "dominant_font_height": font_height,
        "dominant_font_pt": height_to_points(font_height),
        "header_fill_color": apply_plan.get("header_fill_color", ""),
        "dominant_border": apply_plan.get("dominant_border", {}),
        "cell_margin_summary": apply_plan.get("cell_margin_summary", {}),
        "fill_color_counts": table_style.get("fill_color_counts", {}),
        "common_cell_border_fill_ids": table_style.get("common_cell_border_fill_ids", {}),
        "cell_spacing": table_style.get("cell_spacing", ""),
        "repeat_header": table_style.get("repeat_header", ""),
        "supported_apply": ["dominant_font_height"],
        "deferred_apply": ["cell_border", "cell_fill", "cell_margin", "repeat_header"],
    }


def build_table_style_apply_plan(profile: Dict[str, Any] | None) -> Dict[str, Any]:
    """Normalize a recognized HWPX table style into writer-sized steps."""

    if not profile:
        return {"loaded": False, "steps": []}

    font_height = dominant_font_height(profile)
    font_pt = height_to_points(font_height)
    header_fill = dominant_fill_color(profile)
    border = dominant_border_summary(profile)
    margin = cell_margin_summary(profile)
    steps: List[Dict[str, Any]] = []

    if font_height:
        steps.append(
            {
                "name": "dominant_font_height",
                "status": "supported",
                "action": "CharShape",
                "value": font_height,
                "point": font_pt,
            }
        )
    if header_fill:
        steps.append(
            {
                "name": "header_fill",
                "status": "planned",
                "action": "CellBorderFill",
                "value": header_fill,
            }
        )
    if border:
        steps.append(
            {
                "name": "cell_border",
                "status": "planned",
                "action": "CellBorderFill",
                "value": border,
            }
        )
    if margin:
        steps.append(
            {
                "name": "cell_margin",
                "status": "planned",
                "action": "TablePropertyDialog",
                "value": margin,
            }
        )

    return {
        "loaded": True,
        "source": profile.get("style_source", {}),
        "dominant_font_height": font_height,
        "dominant_font_pt": font_pt,
        "header_fill_color": header_fill,
        "dominant_border": border,
        "cell_margin_summary": margin,
        "supported_apply": ["dominant_font_height"],
        "planned_apply": [step["name"] for step in steps if step.get("status") == "planned"],
        "steps": steps,
    }


def dominant_fill_color(profile: Dict[str, Any]) -> str:
    table_style = profile.get("table_style", {})
    counts = table_style.get("fill_color_counts", {})
    fill = highest_count_key(counts, skip_values={"", "none", "NONE"})
    if fill:
        return fill
    scanned: Dict[str, int] = {}
    for border_fill in table_style.get("common_cell_border_fills", {}).values():
        color = str(border_fill.get("fill_color") or "")
        if color and color.lower() != "none":
            scanned[color] = scanned.get(color, 0) + 1
    return highest_count_key(scanned, skip_values={"", "none", "NONE"})


def dominant_border_summary(profile: Dict[str, Any]) -> Dict[str, Any]:
    table_style = profile.get("table_style", {})
    border_fills = table_style.get("common_cell_border_fills", {})
    id_counts = table_style.get("common_cell_border_fill_ids", {})
    side_counts: Dict[str, Dict[str, int]] = {}
    side_values: Dict[str, Dict[str, Dict[str, str]]] = {}

    for border_fill_id, border_fill in border_fills.items():
        weight = safe_int(id_counts.get(str(border_fill_id)), 1)
        borders = border_fill.get("borders", {})
        for side, border in borders.items():
            border_type = str(border.get("type") or "")
            width = str(border.get("width") or "")
            color = str(border.get("color") or "")
            if not border_type or border_type.upper() == "NONE":
                continue
            key = "|".join([border_type, width, color])
            side_counts.setdefault(side, {})
            side_values.setdefault(side, {})
            side_counts[side][key] = side_counts[side].get(key, 0) + weight
            side_values[side][key] = {"type": border_type, "width": width, "color": color}

    summary: Dict[str, Any] = {}
    for side, counts in side_counts.items():
        key = highest_count_key(counts)
        if key:
            summary[side] = {**side_values[side][key], "weight": counts[key]}
    return summary


def cell_margin_summary(profile: Dict[str, Any]) -> Dict[str, Any]:
    table_style = profile.get("table_style", {})
    margins = table_style.get("sample_cell_margins", [])
    if not isinstance(margins, list) or not margins:
        return {}
    keys = sorted({key for margin in margins if isinstance(margin, dict) for key in margin.keys()})
    summary: Dict[str, Any] = {"sample_count": len(margins), "keys": keys}
    common: Dict[str, Any] = {}
    for key in keys:
        counts: Dict[str, int] = {}
        for margin in margins:
            if not isinstance(margin, dict):
                continue
            value = str(margin.get(key, ""))
            if value:
                counts[value] = counts.get(value, 0) + 1
        selected = highest_count_key(counts)
        if selected:
            common[key] = selected
    if common:
        summary["common"] = common
    return summary


def highest_count_key(counts: Dict[str, Any], skip_values: set[str] | None = None) -> str:
    skip = {value.lower() for value in (skip_values or set())}
    best_key = ""
    best_count = -1
    for raw_key, raw_count in counts.items():
        key = str(raw_key)
        if key.lower() in skip:
            continue
        count = safe_int(raw_count, 0)
        if count > best_count:
            best_key = key
            best_count = count
    return best_key


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def dominant_font_height(profile: Dict[str, Any] | None) -> int | None:
    if not profile:
        return None
    table_style = profile.get("table_style", {})
    counts = table_style.get("font_height_counts", {})
    best_height = None
    best_count = -1
    for raw_height, raw_count in counts.items():
        try:
            height = int(raw_height)
            count = int(raw_count)
        except (TypeError, ValueError):
            continue
        if count > best_count:
            best_height = height
            best_count = count
    if best_height:
        return best_height
    for char_pr in table_style.get("common_char_prs", {}).values():
        try:
            return int(char_pr.get("height"))
        except (AttributeError, TypeError, ValueError):
            continue
    return None


def height_to_points(height: int | None) -> float | None:
    if not height:
        return None
    return round(float(height) / 100.0, 1)


def apply_table_style_before_create(hwp, profile: Dict[str, Any] | None, report: Dict[str, Any]) -> None:
    apply_plan = build_table_style_apply_plan(profile)
    if apply_plan.get("loaded"):
        report["table_style_applied"]["apply_plan"] = apply_plan
    font_height = apply_plan.get("dominant_font_height")
    if not font_height:
        return
    report["stage"] = "style"
    report["action"] = "CharShape"
    try:
        params = hwp.HParameterSet.HCharShape
        hwp.HAction.GetDefault("CharShape", params.HSet)
        params.Height = int(font_height)
        hwp.HAction.Execute("CharShape", params.HSet)
        report["table_style_applied"]["dominant_font_height"] = int(font_height)
        report["table_style_applied"]["dominant_font_pt"] = height_to_points(font_height)
    except Exception as exc:
        report["warnings"].append(f"표 글자 크기 profile 적용 실패: {exc}")


def should_close_hwp(report: Dict[str, Any], keep_open_on_error: bool, keep_open_after_save: bool) -> bool:
    if keep_open_after_save and report.get("status") == "ready":
        return False
    if keep_open_on_error and report.get("status") == "failed":
        return False
    return True


def insert_text_table(hwp, rows: List[List[str]], report: Dict[str, Any]) -> None:
    lines = ["\t".join(row) for row in rows]
    insert_text(hwp, "\n".join(lines), report)
    run_action(hwp, "BreakPara", report, "body")
    report["text_table_fallbacks"] += 1


def table_rows_for_hwp(table: Dict[str, Any]) -> List[List[str]]:
    rows = [TABLE_COLUMNS]
    for row in table.get("rows", [])[:20]:
        rows.append(
            [
                str(row.get("category") or ""),
                display_percent(row),
                display_number(row.get("weighted_n")),
                display_number(row.get("raw_n")),
            ]
        )
    return rows


def insert_text(hwp, text: str, report: Dict[str, Any]) -> None:
    report["stage"] = "insert"
    report["action"] = "InsertText"
    try:
        params = hwp.HParameterSet.HInsertText
        hwp.HAction.GetDefault("InsertText", params.HSet)
        params.Text = text
        hwp.HAction.Execute("InsertText", params.HSet)
    except Exception as exc:
        raise HwpWriterError("insert", "InsertText", f"텍스트 입력 실패: {exc}") from exc


def run_action(hwp, action: str, report: Dict[str, Any], stage: str) -> bool:
    report["stage"] = stage
    report["action"] = action
    try:
        return bool(hwp.HAction.Run(action))
    except Exception:
        try:
            return bool(hwp.Run(action))
        except Exception:
            report["warnings"].append(f"아래한글 Action 실행 실패: {action}")
            return False


def save_as_hwpx(hwp, output_file: Path, report: Dict[str, Any], checkpoint_path: Path | None = None) -> None:
    report["stage"] = "document"
    report["action"] = "save_as"
    record_com_step(report, "save_as", "started", path=str(output_file))
    write_checkpoint(report, checkpoint_path)
    format_name = "HWPX" if output_file.suffix.lower() == ".hwpx" else "HWP"
    attempts = [
        lambda: hwp.SaveAs(str(output_file), format_name, ""),
        lambda: hwp.SaveAs(str(output_file)),
        lambda: hwp.Save(),
    ]
    last_error = None
    for attempt in attempts:
        try:
            result = attempt()
            if result is False:
                last_error = "Save returned False"
                record_com_step(report, "save_as", "failed", str(last_error), path=str(output_file))
                write_checkpoint(report, checkpoint_path)
                continue
            if output_file.exists():
                report["document_saved"] = True
                record_com_step(report, "save_as", "ready", path=str(output_file))
                write_checkpoint(report, checkpoint_path)
                return
        except Exception as exc:
            last_error = exc
            record_com_step(report, "save_as", "failed", str(exc), path=str(output_file))
            write_checkpoint(report, checkpoint_path)
    raise HwpWriterError("document", "save_as", f"HWPX 저장 실패: {last_error}")


def close_hwp(hwp, report: Dict[str, Any]) -> None:
    try:
        hwp.Clear(1)
    except Exception:
        pass
    try:
        hwp.Quit()
        report["com"]["closed"] = True
    except Exception:
        pass


def group_charts(rows: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for row in rows:
        if row.get("include_chart"):
            grouped.setdefault(str(row.get("table_key", "")), []).append(row)
    return grouped


def qa_summary(package: Dict[str, Any]) -> str:
    qa = package.get("qa", [])
    if not qa:
        return "QA 이슈 없음"
    return "\n".join(f"{item.get('severity', '')}: {item.get('message', '')}" for item in qa[:20])


def display_percent(row: Dict[str, Any]) -> str:
    value = row.get("percent")
    if value in (None, ""):
        return ""
    unit = row.get("unit") or "%"
    return f"{display_number(value)}{unit}"


def display_number(value: Any) -> str:
    if value in (None, ""):
        return ""
    try:
        parsed = float(value)
        return f"{parsed:,.1f}" if parsed % 1 else f"{parsed:,.0f}"
    except (TypeError, ValueError):
        return str(value)


def new_report(package_file: Path, preflight_file: Path, template_file: Path, output_file: Path, visible: bool) -> Dict[str, Any]:
    return {
        "schema_version": "1.0",
        "status": "started",
        "dry_run": False,
        "keep_open_after_save": False,
        "stage": "",
        "action": "",
        "started_at": now(),
        "finished_at": "",
        "platform": platform.platform(),
        "python": sys.version,
        "package_path": str(package_file),
        "preflight_path": str(preflight_file),
        "template_path": str(template_file),
        "output_path": str(output_file),
        "render_plan_path": "",
        "table_style_profile_path": "",
        "table_style_profile": {"loaded": False},
        "table_style_apply_plan": {"loaded": False, "steps": []},
        "table_style_applied": {},
        "visible": visible,
        "template_copied": False,
        "document_opened": False,
        "document_saved": False,
        "section_count_total": 0,
        "section_count_selected": 0,
        "sections_written": 0,
        "tables_written": 0,
        "text_table_fallbacks": 0,
        "charts_deferred": 0,
        "placeholders": {"body_found": False, "replaced": []},
        "com": {
            "prog_id": "",
            "current_prog_id": "",
            "file_path_checker": "",
            "visible_applied": None,
            "closed": False,
            "steps": [],
        },
        "warnings": [],
        "errors": [],
    }


def load_json(path: str | Path) -> Dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_json(path: str | Path, value: Dict[str, Any]) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_bool(value: str) -> bool:
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on", "표시", "보임"}


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create HWPX drafts through Hancom HWP COM automation.")
    parser.add_argument("--package")
    parser.add_argument("--preflight")
    parser.add_argument("--template")
    parser.add_argument("--output")
    parser.add_argument("--visible", default="false")
    parser.add_argument("--report-output")
    parser.add_argument("--keep-open-on-error", action="store_true")
    parser.add_argument("--max-sections", type=int, default=0)
    parser.add_argument("--render-plan-output")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--table-style-profile")
    parser.add_argument("--keep-open-after-save", action="store_true")
    parser.add_argument("--check-environment", action="store_true")
    args = parser.parse_args(argv)

    try:
        if args.check_environment:
            report = check_environment(args.report_output, parse_bool(args.visible))
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return 0 if report.get("status") == "ready" else 2

        required = {
            "--package": args.package,
            "--preflight": args.preflight,
            "--template": args.template,
            "--output": args.output,
        }
        missing = [name for name, value in required.items() if not value]
        if missing:
            parser.error("문서 생성 모드에는 다음 인자가 필요합니다: " + ", ".join(missing))

        output = write_hwp_document(
            str(args.package),
            str(args.preflight),
            str(args.template),
            str(args.output),
            parse_bool(args.visible),
            args.report_output,
            args.keep_open_on_error,
            args.max_sections if args.max_sections > 0 else None,
            args.render_plan_output,
            args.dry_run,
            args.table_style_profile,
            args.keep_open_after_save,
        )
        print(str(output))
        return 0
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
