"""Create HWPX drafts through the local Hancom HWP COM automation API.

이 모듈은 HWPX를 직접 XML로 조립하지 않고, 사용자의 Windows PC에 설치된
아래한글을 실행해 템플릿 사본에 본문과 표를 입력한다. rhwp 기반 writer는
장기 후보로 두고, 알파 단계의 실사용 writer는 아래한글 COM을 우선한다.
"""

from __future__ import annotations

import argparse
import json
import math
import platform
import shutil
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

try:
    from .hwp_style_config import REQUIRED_STYLE_NAMES, load_hwp_style_config, style_name_for_cell_role
    from .hwpx_style_registry import register_named_styles
    from .report_package import split_narrative_blocks
    from .hwp_table_layout import plan_table_parts, _text_units
except ImportError:
    from hwp_style_config import REQUIRED_STYLE_NAMES, load_hwp_style_config, style_name_for_cell_role
    from hwpx_style_registry import register_named_styles
    from report_package import split_narrative_blocks
    from hwp_table_layout import plan_table_parts, _text_units


BODY_PLACEHOLDER = "{{BODY}}"
REPORT_PLACEHOLDERS = {
    "{{REPORT_TITLE}}": "report_title",
    "{{PROJECT_NAME}}": "project_name",
    "{{GENERATED_AT}}": "created_at",
    "{{QA_SUMMARY}}": "qa_summary",
}
TABLE_COLUMNS = ["항목", "비율", "가중 N", "원 N"]
DISPATCH_MODES = ("ensure_dispatch", "dispatch", "dispatch_ex")


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
    dispatch_mode: str = "ensure_dispatch",
    style_config_path: str | Path | None = None,
) -> Path:
    """Write a report draft and always write a companion JSON report."""

    package_file = Path(package_path).resolve()
    preflight_file = Path(preflight_path).resolve()
    template_file = Path(template_path).resolve()
    output_file = Path(output_path).resolve()
    writer_report = new_report(package_file, preflight_file, template_file, output_file, visible)
    writer_report["keep_open_after_save"] = keep_open_after_save
    dispatch_mode = normalize_dispatch_mode(dispatch_mode)
    writer_report["com"]["dispatch_mode"] = dispatch_mode
    report_file = Path(report_path).resolve() if report_path else output_file.with_name(output_file.stem + "_hwp_writer_report.json")
    render_plan_file = Path(render_plan_path).resolve() if render_plan_path else output_file.with_name(output_file.stem + "_hwp_render_plan.json")
    table_style_profile_file = Path(table_style_profile_path).resolve() if table_style_profile_path else None
    style_config_file = Path(style_config_path).resolve() if style_config_path else None
    hwp = None
    excel = None
    excel_workbook = None
    temporary_root = None

    try:
        package = load_json(package_file)
        preflight = load_json(preflight_file)
        validate_preflight(preflight)
        style_config = load_writer_style_config(style_config_file, writer_report)
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

        temporary_root = Path(tempfile.mkdtemp(prefix="research-helper-hwp-"))
        working_template = temporary_root / "styled_template.hwpx"
        style_indexes = None
        if template_file.suffix.lower() == ".hwpx":
            style_indexes = register_styles_or_raise(template_file, working_template, style_config, writer_report)

        writer_report["stage"] = "com"
        writer_report["action"] = "create_hwp_object"
        write_json(report_file, writer_report)
        hwp = create_hwp_object(writer_report, report_file, dispatch_mode)
        set_visible(hwp, visible, writer_report, report_file)
        write_json(report_file, writer_report)
        if template_file.suffix.lower() == ".hwp":
            converted_template = temporary_root / "converted_template.hwpx"
            open_document(hwp, template_file, writer_report, report_file)
            save_as_hwpx(hwp, converted_template, writer_report, report_file)
            close_current_document(hwp)
            writer_report["document_saved"] = False
            style_indexes = register_styles_or_raise(converted_template, working_template, style_config, writer_report)
        open_document(hwp, working_template, writer_report, report_file)
        writer_report["template_copied"] = True
        write_json(report_file, writer_report)

        excel, excel_workbook = open_excel_clipboard_source(package, package_file, writer_report)

        replace_header_placeholders(hwp, package, writer_report)
        if not find_placeholder(hwp, BODY_PLACEHOLDER):
            raise HwpWriterError("template", "find_body", "{{BODY}} placeholder를 문서 본문에서 찾지 못했습니다.")
        writer_report["placeholders"]["body_found"] = True
        run_action(hwp, "Delete", writer_report, "template")
        write_json(report_file, writer_report)

        write_body(
            hwp,
            package,
            writer_report,
            max_sections,
            table_style_profile,
            excel_workbook,
            style_indexes,
            style_config,
        )
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
        close_excel_clipboard_source(excel, excel_workbook, writer_report)
        finalize_hwp_resources(
            temporary_root,
            hwp,
            writer_report,
            keep_open_on_error,
            keep_open_after_save,
        )
        write_json(report_file, writer_report)


def check_environment(
    report_path: str | Path | None = None,
    visible: bool = False,
    dispatch_mode: str = "ensure_dispatch",
) -> Dict[str, Any]:
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
            "dispatch_mode": normalize_dispatch_mode(dispatch_mode),
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
        dispatch_mode = normalize_dispatch_mode(dispatch_mode)
        report["com"]["dispatch_mode"] = dispatch_mode
        if platform.system().lower() != "windows":
            raise HwpWriterError("environment", "platform", "아래한글 COM writer는 Windows에서만 실행할 수 있습니다.")
        checkpoint_file = Path(report_path).resolve() if report_path else None
        write_checkpoint(report, checkpoint_file)
        hwp = create_hwp_object(report, checkpoint_file, dispatch_mode)
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


def load_writer_style_config(path: Path | None, report: Dict[str, Any]) -> Dict[str, Any]:
    try:
        config = load_hwp_style_config(path)
    except Exception as exc:
        raise HwpWriterError("style", "load_style_config", f"HWP 스타일 설정을 읽지 못했습니다: {exc}") from exc
    report["style_config_path"] = str(path) if path else ""
    report["style_preset_name"] = str(config.get("preset_name") or "")
    return config


def register_styles_or_raise(
    template: Path,
    working_template: Path,
    config: Dict[str, Any],
    report: Dict[str, Any],
) -> Dict[str, int]:
    try:
        indexes = register_named_styles(template, working_template, config)
    except Exception as exc:
        raise HwpWriterError("style", "register_named_styles", f"HWPX 스타일 등록에 실패했습니다: {exc}") from exc
    missing = [name for name in REQUIRED_STYLE_NAMES if name not in indexes]
    if missing:
        raise HwpWriterError("style", "register_named_styles", f"등록된 HWP 스타일 인덱스가 없습니다: {missing}")
    report["style_index_map"] = {name: int(indexes[name]) for name in REQUIRED_STYLE_NAMES}
    report["style_application_counts"] = {name: 0 for name in REQUIRED_STYLE_NAMES}
    report["working_template_path"] = str(working_template)
    return report["style_index_map"]


def close_current_document(hwp) -> None:
    try:
        hwp.Clear(1)
    except Exception as exc:
        raise HwpWriterError("document", "close_template", f"변환한 HWP 템플릿을 닫지 못했습니다: {exc}") from exc


def create_hwp_object(
    report: Dict[str, Any],
    checkpoint_path: Path | None = None,
    dispatch_mode: str = "ensure_dispatch",
):
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
            report["com"]["dispatch_mode"] = dispatch_mode
            report["com"]["current_prog_id"] = prog_id
            record_com_step(report, "dispatch", "started", prog_id=prog_id, dispatch_mode=dispatch_mode)
            write_checkpoint(report, checkpoint_path)
            hwp = dispatch_hwp_object(win32com.client, prog_id, dispatch_mode)
            report["com"]["prog_id"] = prog_id
            record_com_step(report, "dispatch", "ready", prog_id=prog_id, dispatch_mode=dispatch_mode)
            write_checkpoint(report, checkpoint_path)
            register_file_path_checker(hwp, report, checkpoint_path)
            return hwp
        except Exception as exc:
            last_error = exc
            record_com_step(report, "dispatch", "failed", str(exc), prog_id=prog_id, dispatch_mode=dispatch_mode)
            write_checkpoint(report, checkpoint_path)
    raise HwpWriterError("com", "create_object", f"아래한글 COM 객체를 생성하지 못했습니다: {last_error}")


def dispatch_hwp_object(win32_client, prog_id: str, dispatch_mode: str):
    if dispatch_mode == "dispatch":
        return win32_client.Dispatch(prog_id)
    if dispatch_mode == "dispatch_ex":
        return win32_client.DispatchEx(prog_id)
    return win32_client.gencache.EnsureDispatch(prog_id)


def normalize_dispatch_mode(value: str | None) -> str:
    mode = str(value or "ensure_dispatch").strip().lower().replace("-", "_")
    if mode not in DISPATCH_MODES:
        allowed = ", ".join(DISPATCH_MODES)
        raise HwpWriterError("validate", "dispatch_mode", f"지원하지 않는 HWP COM dispatch mode입니다: {value}. 허용값: {allowed}")
    return mode


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


def require_style_index(style_indexes: Dict[str, int], style_name: str) -> int:
    try:
        return int(style_indexes[style_name])
    except (KeyError, TypeError, ValueError) as exc:
        raise HwpWriterError("style", "Style", f"HWP 스타일 인덱스가 없습니다: {style_name}") from exc


def apply_named_style(hwp, style_index: int, report: Dict[str, Any], style_name: str) -> None:
    report["stage"] = "style"
    report["action"] = "Style"
    try:
        params = hwp.HParameterSet.HStyle
        hwp.HAction.GetDefault("Style", params.HSet)
        params.Apply = int(style_index)
        if not hwp.HAction.Execute("Style", params.HSet):
            raise RuntimeError("Style action returned False")
    except Exception as exc:
        raise HwpWriterError("style", "Style", f"HWP 스타일 적용 실패: {style_name}: {exc}") from exc
    counts = report.setdefault("style_application_counts", {})
    counts[style_name] = int(counts.get(style_name, 0)) + 1


def write_body(
    hwp,
    package: Dict[str, Any],
    report: Dict[str, Any],
    max_sections: int | None = None,
    table_style_profile: Dict[str, Any] | None = None,
    excel_workbook=None,
    style_indexes: Dict[str, int] | None = None,
    style_config: Dict[str, Any] | None = None,
) -> None:
    tables_by_key = {str(table.get("table_key", "")): table for table in package.get("tables", [])}
    charts_by_key = group_charts(package.get("charts", []))
    sections = select_sections(package, max_sections)
    for index, section in enumerate(sections, start=1):
        if index > 1:
            run_action(hwp, "BreakPage", report, "body")
        key = str(section.get("table_key", ""))
        title = str(section.get("title") or key or f"문항 {index}")
        if style_indexes is not None:
            apply_named_style(hwp, require_style_index(style_indexes, "보고서 본문1"), report, "보고서 본문1")
        insert_text(hwp, title, report)
        run_action(hwp, "BreakPara", report, "body")
        insert_narrative_blocks(hwp, section, report, style_indexes, style_config)
        run_action(hwp, "BreakPara", report, "body")

        table = tables_by_key.get(key)
        if table:
            try:
                width = current_body_width(hwp, report)
                font_size = max(style_config["paragraph_styles"][name]["font_size_pt"] for name in ("표보기", "표배너", "표숫자")) if style_config else 9
                parts = plan_table_parts(table, width, font_size) if table.get("matrix") else [table]
            except ValueError as exc:
                raise HwpWriterError("layout", "table_split", f"{key}: {exc}") from exc
            report.setdefault("table_layouts", []).append({"table_key": key, "body_width_hwpunit": width,
                "parts": [{name: part.get(name) for name in ("part_index", "part_count", "source_columns", "repeated_columns", "column_widths_hwpunit")} for part in parts]})
            if len(parts) > 1:
                report["warnings"].append(f"{key}: 본문 폭에 맞춰 {len(parts)}개 표로 가로 분할했습니다. 보기/BASE 열은 반복됩니다.")
            for part in parts:
                write_table_part(hwp, part, title, report, table_style_profile, excel_workbook, style_indexes, style_config)
            report["source_tables_written"] = report.get("source_tables_written", 0) + 1
        else:
            report["warnings"].append(f"삽입표 데이터 없음: {key}")

        if charts_by_key.get(key):
            if style_indexes is not None:
                apply_named_style(hwp, require_style_index(style_indexes, "보고서 본문1"), report, "보고서 본문1")
            insert_text(hwp, f"[차트 삽입 필요] {title}", report)
            run_action(hwp, "BreakPara", report, "body")
            report["charts_deferred"] += 1

        if style_indexes is not None:
            apply_named_style(hwp, require_style_index(style_indexes, "보고서 본문1"), report, "보고서 본문1")
        insert_text(hwp, f"source: {key}", report)
        run_action(hwp, "BreakPara", report, "body")
        run_action(hwp, "BreakPara", report, "body")
        report["sections_written"] += 1


def write_table_part(hwp, table, title, report, profile, workbook, style_indexes, style_config):
    if int(table.get("part_index") or 1) > 1:
        run_action(hwp, "BreakPage", report, "body")
    if style_indexes is not None:
        apply_named_style(hwp, require_style_index(style_indexes, "보고서 본문1"), report, "보고서 본문1")
    caption = str(table.get("title") or title)
    if int(table.get("part_count") or 1) > 1:
        caption += f" [{table['part_index']}/{table['part_count']}]"
    insert_text(hwp, caption, report)
    run_action(hwp, "BreakPara", report, "body")
    inserted = insert_clipboard_table(hwp, workbook, table, report, profile, style_indexes, style_config)
    if not inserted:
        inserted = insert_hwp_table(hwp, table_rows_for_hwp(table), report, profile,
                                    table.get("merged_ranges", []), table, style_indexes, style_config)
        report["table_results"].append(table_result(table, "contract_fallback", "tbl" if inserted else "",
            "applied" if inserted else "failed", report.pop("clipboard_failure", "")))
    if not inserted:
        raise HwpWriterError("table", "TableCreate", f"HWP 표 생성에 실패했습니다: {table.get('table_key')}")
    run_action(hwp, "BreakPara", report, "body")


def current_body_width(hwp, report):
    """Read the active section, not an assumed A4 page or a template's first page."""
    try:
        params = hwp.HParameterSet.HSecDef
        hwp.HAction.GetDefault("PageSetup", params.HSet)
        page = params.PageDef
        width = page.PaperHeight if page.Landscape else page.PaperWidth
        width -= page.LeftMargin + page.RightMargin
        if page.GutterType != 2:
            width -= page.GutterLen
        action = hwp.CreateAction("MultiColumn")
        columns = action.CreateSet()
        action.GetDefault(columns)
        if int(columns.Item("Count")) != 1:
            raise ValueError("다단 본문은 지원하지 않습니다. 한 단 템플릿으로 변경해 주세요.")
        if width <= 0:
            raise ValueError("본문 가용 폭이 없습니다.")
        return int(width)
    except Exception as exc:
        raise HwpWriterError("layout", "PageSetup", f"본문 폭 검사 실패: {exc}") from exc


def insert_narrative_blocks(
    hwp,
    section: Dict[str, Any],
    report: Dict[str, Any],
    style_indexes=None,
    style_config: Dict[str, Any] | None = None,
) -> None:
    blocks = section.get("narrative_blocks") or split_narrative_blocks(str(section.get("narrative_final") or ""))
    bullet = style_config["paragraph_styles"]["보고서 본문2"]["bullet"] if style_config is not None else "-"
    for block in blocks:
        text = str(block.get("text") or "")
        style_name = str(block.get("style") or "보고서 본문1")
        if style_name == "보고서 본문2" and bullet and text.startswith((bullet + " ", bullet + "\t")):
            text = text[len(bullet):].lstrip(" \t")
        if style_indexes is not None:
            apply_named_style(hwp, require_style_index(style_indexes, style_name), report, style_name)
        insert_text(hwp, text, report)
        run_action(hwp, "BreakPara", report, "body")


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
    merged_ranges: List[Dict[str, Any]] | None = None,
    table: Dict[str, Any] | None = None,
    style_indexes: Dict[str, int] | None = None,
    style_config: Dict[str, Any] | None = None,
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
        if table and table.get("column_widths_hwpunit"):
            widths = table["column_widths_hwpunit"]
            properties = params.TableProperties
            margin = int(properties.CellMarginLeft) + int(properties.CellMarginRight)
            properties.CellSpacing = 0
            params.WidthValue = sum(widths)
            params.CreateItemArray("ColWidth", len(widths))
            for index, width in enumerate(widths):
                # TableCreate adds the default left/right padding to each ColWidth.
                if width <= margin:
                    raise HwpWriterError("layout", "TableCreate", "열 폭이 셀 안쪽 여백보다 작습니다.")
                params.ColWidth.SetItem(index, width - margin)
        if not hwp.HAction.Execute("TableCreate", params.HSet):
            return False
        control = hwp.ParentCtrl
        for row_idx, row in enumerate(rows):
            for col_idx, value in enumerate(row):
                style_name = style_name_for_cell_role(table_cell_role(table or {}, row_idx, col_idx))
                if style_indexes is not None:
                    apply_named_style(hwp, require_style_index(style_indexes, style_name), report, style_name)
                insert_text(hwp, value, report)
                if style_config is not None:
                    apply_cell_appearance(
                        hwp,
                        style_config,
                        style_name,
                        row_idx,
                        col_idx,
                        len(rows),
                        len(row),
                        report,
                    )
                if not (row_idx == len(rows) - 1 and col_idx == len(row) - 1):
                    run_action(hwp, "TableRightCell", report, "table")
        apply_table_merges(hwp, merged_ranges or [], len(rows), report)
        verify_table_width(control, table or {})
        leave_hwp_table(hwp, control, report)
        report["tables_written"] += 1
        return True
    except HwpWriterError:
        raise
    except Exception as exc:
        report["warnings"].append(f"HWP 표 객체 생성 실패, 텍스트 표로 대체합니다: {exc}")
        return False


def apply_table_merges(
    hwp,
    merged_ranges: List[Dict[str, Any]],
    row_count: int,
    report: Dict[str, Any],
) -> None:
    """Merge contract cells from bottom-right so earlier coordinates stay stable."""

    ranges = [item for item in merged_ranges if int(item.get("rowspan") or 1) > 1 or int(item.get("colspan") or 1) > 1]
    for merged in sorted(ranges, key=lambda item: (int(item.get("row") or 1), int(item.get("col") or 1)), reverse=True):
        run_action(hwp, "TableColBegin", report, "table")
        for _ in range(row_count):
            run_action(hwp, "TableUpperCell", report, "table")
        for _ in range(int(merged.get("row") or 1) - 1):
            run_action(hwp, "TableLowerCell", report, "table")
        for _ in range(int(merged.get("col") or 1) - 1):
            run_action(hwp, "TableRightCell", report, "table")
        run_action(hwp, "TableCellBlock", report, "table")
        run_action(hwp, "TableCellBlockExtend", report, "table")
        for _ in range(int(merged.get("rowspan") or 1) - 1):
            run_action(hwp, "TableLowerCell", report, "table")
        for _ in range(int(merged.get("colspan") or 1) - 1):
            run_action(hwp, "TableRightCell", report, "table")
        run_action(hwp, "TableMergeCell", report, "table")


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


def finalize_hwp_resources(
    temporary_root: Path | None,
    hwp,
    report: Dict[str, Any],
    keep_open_on_error: bool,
    keep_open_after_save: bool,
) -> None:
    if hwp is not None and should_close_hwp(report, keep_open_on_error, keep_open_after_save):
        close_hwp(hwp, report)
    elif hwp is not None:
        report["com"]["closed"] = False
        report["warnings"].append("HWP COM document was intentionally left open.")

    retained = hwp is not None and not bool(report.get("com", {}).get("closed"))
    report["temporary_files_retained"] = retained
    report["temporary_root_path"] = str(temporary_root) if temporary_root else ""
    if temporary_root is None or retained:
        return
    try:
        shutil.rmtree(temporary_root)
    except Exception as exc:
        report["temporary_files_retained"] = True
        report["warnings"].append(f"HWP 임시 작업 폴더를 삭제하지 못했습니다: {exc}")


def insert_text_table(hwp, rows: List[List[str]], report: Dict[str, Any]) -> None:
    lines = ["\t".join(row) for row in rows]
    insert_text(hwp, "\n".join(lines), report)
    run_action(hwp, "BreakPara", report, "body")
    report["text_table_fallbacks"] += 1


def table_rows_for_hwp(table: Dict[str, Any]) -> List[List[str]]:
    if table.get("matrix"):
        return [
            [str(cell.get("display_text") or "") for cell in row]
            for row in table.get("matrix", [])
        ]
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


def table_cell_role(table: Dict[str, Any], row_index: int, col_index: int) -> str:
    matrix = table.get("matrix") or []
    try:
        return str(matrix[row_index][col_index].get("role") or "unknown")
    except (IndexError, TypeError, AttributeError):
        if row_index == 0:
            return "header"
        return "stub" if col_index == 0 else "value"


_BORDER_WIDTHS_MM = (0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0)


def border_width_code(width_mm: float) -> int:
    return min(range(len(_BORDER_WIDTHS_MM)), key=lambda index: abs(_BORDER_WIDTHS_MM[index] - float(width_mm)))


def hwp_color(hwp, value: str) -> int:
    text = str(value).lstrip("#")
    red, green, blue = (int(text[index : index + 2], 16) for index in (0, 2, 4))
    try:
        return int(hwp.RGBColor(red, green, blue))
    except Exception:
        return red | (green << 8) | (blue << 16)


def set_parameter_item(target, name: str, value: Any) -> None:
    fallback_name = {"Type": "type", "BorderColorLeft": "BorderCorlorLeft"}.get(name, name)
    if hasattr(target, fallback_name):
        setattr(target, fallback_name, value)
        return
    try:
        target.HSet.SetItem(name, value)
        return
    except AttributeError:
        pass
    try:
        target.SetItem(name, value)
        return
    except AttributeError:
        pass
    setattr(target, fallback_name, value)


def apply_cell_appearance(
    hwp,
    config: Dict[str, Any],
    style_name: str,
    row_index: int,
    col_index: int,
    row_count: int,
    col_count: int,
    report: Dict[str, Any],
    rowspan: int = 1,
    colspan: int = 1,
) -> None:
    cell_style = config["table_cell_styles"][style_name]
    border = config["table_border"]
    inner = border_width_code(float(border["inner_width_mm"]))
    outer = border_width_code(float(border["outer_width_mm"]))
    report["stage"] = "style"
    report["action"] = "CellBorderFill"
    try:
        run_action(hwp, "TableCellBlock", report, "style")
        params = hwp.HParameterSet.HCellBorderFill
        # CellFill can restore the imported borders. Apply borders after the fill.
        hwp.HAction.GetDefault("CellFill", params.HSet)
        fill = params.FillAttr
        set_parameter_item(fill, "Type", 1)
        set_parameter_item(fill, "WinBrushFaceColor", hwp_color(hwp, cell_style["fill_color"]))
        set_parameter_item(fill, "WinBrushHatchColor", 0)
        set_parameter_item(fill, "WinBrushFaceStyle", -1)
        set_parameter_item(fill, "WindowsBrush", 1)
        if not hwp.HAction.Execute("CellFill", params.HSet):
            raise RuntimeError("CellFill action returned False")
        hwp.HAction.GetDefault("CellBorderFill", params.HSet)
        set_parameter_item(params, "ApplyTo", 0)
        selected = params.SelCellsBorderFill
        color = hwp_color(hwp, border["color"])
        solid = hwp.HwpLineType("Solid")
        for side, width in {
            "Left": outer if col_index == 0 else inner,
            "Right": outer if col_index + colspan == col_count else inner,
            "Top": outer if row_index == 0 else inner,
            "Bottom": outer if row_index + rowspan == row_count else inner,
        }.items():
            for target in (params, selected):
                set_parameter_item(target, f"BorderType{side}", solid)
                set_parameter_item(target, f"BorderWidth{side}", width)
                set_parameter_item(target, f"BorderColor{side}", color)
        if not hwp.HAction.Execute("CellBorderFill", params.HSet):
            raise RuntimeError("CellBorderFill action returned False")
        run_action(hwp, "Cancel", report, "style")
        horizontal = str(config["paragraph_styles"][style_name]["alignment"]).title()
        vertical = str(cell_style["vertical_alignment"]).title()
        if not run_action(hwp, f"TableCellAlign{horizontal}{vertical}", report, "style"):
            raise RuntimeError(f"cell alignment action returned False: {horizontal}/{vertical}")
    except HwpWriterError:
        raise
    except Exception as exc:
        raise HwpWriterError("style", "CellBorderFill", f"HWP 표 셀 서식 적용 실패: {style_name}: {exc}") from exc


def is_hwp_table_control(control: Any) -> bool:
    if control is None:
        return False
    try:
        return str(control.CtrlID or "").strip().lower() == "tbl"
    except Exception:
        return False


def control_instance_id(control: Any) -> str:
    if control is None:
        return ""
    try:
        value = control.GetCtrlInstID()
        return str(value) if value is not None else ""
    except Exception:
        return ""


def same_hwp_control(left: Any, right: Any) -> bool:
    if left is None or right is None:
        return False
    left_id = control_instance_id(left)
    right_id = control_instance_id(right)
    return left_id == right_id if left_id and right_id else left is right


def snapshot_hwp_controls(hwp):
    """Capture scalar IDs: LastCtrl may belong to an unchanged template footer."""
    controls = {}
    try:
        control = hwp.HeadCtrl
        while control is not None:
            key = control_instance_id(control)
            if key in controls or (not key and str(control.CtrlID).strip().lower() in {"tbl", "gso"}):
                raise ValueError("문서 개체 ID를 안전하게 구분할 수 없습니다.")
            # Section/column controls (secd/cold) do not expose an instance ID.
            if key:
                controls[key] = control
            control = control.Next
        return controls
    except Exception as exc:
        raise HwpWriterError("table", "control_snapshot", f"문서 개체 목록 검증 실패: {exc}") from exc


def open_excel_clipboard_source(package: Dict[str, Any], package_file: Path, report: Dict[str, Any]):
    if not any(table.get("cell_contract") for table in package.get("tables", [])):
        return None, None
    source_text = str(package.get("meta", {}).get("source_workbook") or "").strip()
    if not source_text:
        report["warnings"].append("clipboard 삽입용 source workbook 경로가 없어 contract fallback을 사용합니다.")
        return None, None
    source_path = Path(source_text)
    if not source_path.is_absolute():
        source_path = (package_file.parent / source_path).resolve()
    if not source_path.exists():
        report["warnings"].append(f"clipboard 삽입용 workbook을 찾지 못했습니다: {source_path}")
        return None, None
    try:
        import win32com.client  # type: ignore

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(source_path), 0, True)
        report["excel"] = {"opened": True, "source_workbook": str(source_path), "closed": False}
        return excel, workbook
    except Exception as exc:
        report["warnings"].append(f"Excel clipboard source를 열지 못해 contract fallback을 사용합니다: {exc}")
        return None, None


def close_excel_clipboard_source(excel, workbook, report: Dict[str, Any]) -> None:
    if workbook is not None:
        try:
            workbook.Close(False)
        except Exception:
            pass
    if excel is not None:
        try:
            excel.CutCopyMode = False
            excel.Quit()
            report.setdefault("excel", {})["closed"] = True
        except Exception:
            pass


def insert_clipboard_table(
    hwp,
    workbook,
    table: Dict[str, Any],
    report: Dict[str, Any],
    profile: Dict[str, Any] | None,
    style_indexes: Dict[str, int] | None = None,
    style_config: Dict[str, Any] | None = None,
) -> bool:
    if workbook is None or not table.get("cell_contract"):
        report["clipboard_failure"] = "Excel source unavailable" if table.get("cell_contract") else "legacy table contract"
        return False
    editing_started = False
    scratch = None
    try:
        if table.get("layout_prepared"):
            scratch = workbook.Application.Workbooks.Add()
            source_range = prepare_layout_range(scratch, table, style_config)
        else:
            worksheet = workbook.Worksheets(str(table.get("source_sheet") or ""))
            source_range = worksheet.Range(str(table.get("source_range") or ""))
        last_error = None
        for _ in range(3):
            try:
                source_range.Copy()
                last_error = None
                break
            except Exception as exc:
                last_error = exc
                time.sleep(0.25)
        if last_error is not None:
            raise last_error

        before = snapshot_hwp_controls(hwp)
        disable_picture_paste(hwp, report)
        paste_result = run_action(hwp, "Paste", report, "table")
        time.sleep(0.5)
        after = snapshot_hwp_controls(hwp)
        added = set(after) - set(before)
        if set(before) - set(after):
            raise HwpWriterError("table", "Paste", "붙여넣기 중 기존 개체가 사라져 대체 표 생성을 차단합니다.")
        if not added:
            if paste_result:
                raise HwpWriterError("table", "Paste", "붙여넣기 변경 상태가 불명확하여 Undo 및 대체 표 생성을 차단합니다.")
            raise RuntimeError("Paste가 실행되지 않았습니다. 기존 문서는 되돌리지 않습니다.")
        if not paste_result or len(added) != 1 or not is_hwp_table_control(after[next(iter(added))]):
            rollback_clipboard_paste(hwp, report)
            detail = "Paste returned False; " if not paste_result else ""
            raise RuntimeError(detail + "붙여넣기 결과가 단일 신규 HWP 표 객체가 아닙니다.")
        control = after[next(iter(added))]

        rows, cols = hwp_table_dimensions(control)
        expected_rows = int(table.get("row_count") or 0)
        expected_cols = int(table.get("col_count") or 0)
        if rows and cols and (rows != expected_rows or cols != expected_cols):
            rollback_clipboard_paste(hwp, report)
            raise RuntimeError(f"붙여넣기 표 크기가 다릅니다: {rows}x{cols}, expected {expected_rows}x{expected_cols}")
        verify_table_width(control, table)

        # Paste leaves the cursor outside the object; select its anchor before editing cells.
        try:
            entered = (
                hwp.SetPosBySet(control.GetAnchorPos(0))
                and hwp.FindCtrl()
                and run_action(hwp, "ShapeObjTableSelCell", report, "table")
            )
            entered = entered and same_hwp_control(getattr(hwp, "ParentCtrl", None), control)
        except Exception as exc:
            rollback_clipboard_paste(hwp, report)
            raise RuntimeError(f"붙여넣은 표 셀 진입 API 실패: {exc}") from exc
        if not entered:
            rollback_clipboard_paste(hwp, report)
            raise RuntimeError("붙여넣은 표 셀에 진입하지 못했습니다.")
        run_action(hwp, "Cancel", report, "table")

        editing_started = True
        if style_indexes is not None and style_config is not None:
            apply_table_matrix_styles(hwp, table, style_indexes, style_config, report)
        else:
            apply_table_style_after_paste(hwp, profile, report)
        leave_hwp_table(hwp, control, report)
        report["tables_written"] += 1
        report["table_results"].append(table_result(table, "clipboard", "tbl", "applied", ""))
        report.pop("clipboard_failure", None)
        return True
    except HwpWriterError:
        raise
    except Exception as exc:
        if editing_started:
            raise HwpWriterError("style", str(report.get("action") or "clipboard"), f"붙여넣은 표 편집 중 실패하여 중복 표 생성을 차단합니다: {exc}") from exc
        report["clipboard_failure"] = str(exc)
        report["warnings"].append(f"Excel clipboard 표 삽입 실패, contract fallback을 사용합니다: {table.get('table_key')}: {exc}")
        return False
    finally:
        if scratch is not None:
            try:
                scratch.Close(False)
            except Exception as exc:
                report["warnings"].append(f"표 폭 조정 임시 통합문서 닫기 실패: {exc}")


def prepare_layout_range(workbook, table, style_config):
    """Use a disposable Excel grid; original workbook values and widths stay untouched."""
    sheet = workbook.Worksheets.Item(1)
    rows, cols = table["row_count"], table["col_count"]
    area = sheet.Range(sheet.Cells(1, 1), sheet.Cells(rows, cols))
    area.NumberFormat = "@"  # Parenthesized BASE counts must remain strings, not negative numbers.
    area.Value2 = tuple(tuple(row) for row in table_rows_for_hwp(table))
    font_size = max(style_config["paragraph_styles"][name]["font_size_pt"] for name in ("표보기", "표배너", "표숫자")) if style_config else 9
    area.Font.Name = style_config["paragraph_styles"]["표보기"]["font_family"] if style_config else "맑은 고딕"
    area.Font.Size = font_size
    area.WrapText = True
    widths = table["column_widths_hwpunit"]
    for index, width in enumerate(widths, 1):
        column = sheet.Columns.Item(index)
        low, high = 0.1, 255
        # Excel column units depend on the Normal font. Measure actual points instead.
        for _ in range(12):
            middle = (low + high) / 2
            column.ColumnWidth = middle
            if column.Width <= width / 100:
                low = middle
            else:
                high = middle
        column.ColumnWidth = low
    for merged in table.get("merged_ranges", []):
        row, col = merged["row"], merged["col"]
        sheet.Range(sheet.Cells(row, col), sheet.Cells(row + merged["rowspan"] - 1, col + merged["colspan"] - 1)).Merge()
    heights = [15.0] * rows
    for row in table["matrix"]:
        for cell in row:
            if cell.get("covered_by"):
                continue
            first = cell["col"] - 1
            available = sum(widths[first:first + cell["colspan"]]) / 100 - 6
            if available <= 0:
                raise ValueError("셀 너비가 여백보다 작습니다.")
            lines = max(1, math.ceil(_text_units(cell.get("display_text", "")) * font_size / available))
            height = (lines * font_size * 1.6 + 4) / cell["rowspan"]
            if height > 409:
                raise ValueError("셀 내용이 Excel 임시 표의 최대 행 높이를 초과합니다.")
            for index in range(cell["row"] - 1, cell["row"] - 1 + cell["rowspan"]):
                heights[index] = max(heights[index], height)
    for index, height in enumerate(heights, 1):
        sheet.Rows.Item(index).RowHeight = height
    return area


def verify_table_width(control, table):
    if not table.get("layout_prepared"):
        return
    try:
        actual = int(control.Properties.Item("Width"))
        if actual <= 0 or actual > int(table["body_width_hwpunit"]):
            raise ValueError(f"표 폭 {actual}이 본문 폭 {table['body_width_hwpunit']}을 초과합니다.")
        # Anchor to the column, not an indented narrative paragraph.
        properties = control.Properties
        for key, value in (("TreatAsChar", 0), ("HorzRelTo", 2), ("HorzAlign", 0), ("HorzOffset", 0)):
            properties.SetItem(key, value)
        control.Properties = properties
        table["actual_width_hwpunit"] = actual
    except Exception as exc:
        raise HwpWriterError("layout", "table_width", f"표 너비 검증 실패: {exc}") from exc


def rollback_clipboard_paste(hwp, report: Dict[str, Any]) -> None:
    run_action(hwp, "Cancel", report, "table")
    if not run_action(hwp, "Undo", report, "table"):
        raise HwpWriterError("table", "Undo", "붙여넣은 개체를 되돌리지 못해 계약 기반 표 생성을 중단합니다.")


def leave_hwp_table(hwp, control, report: Dict[str, Any]) -> None:
    # Return after this table, not after template footers/QA at the document end.
    try:
        run_action(hwp, "Cancel", report, "table")
        if not is_hwp_table_control(control) or not hwp.SetPosBySet(control.GetAnchorPos(0)):
            raise RuntimeError("표 앵커를 찾지 못했습니다.")
        list_id, paragraph, offset = hwp.GetPos()
        if not hwp.SetPos(list_id, paragraph, offset + 1) or is_hwp_table_control(hwp.ParentCtrl):
            raise RuntimeError("표 뒤 본문으로 이동하지 못했습니다.")
    except Exception as exc:
        raise HwpWriterError("table", "SetPos", f"표 뒤 삽입 위치 복귀 실패: {exc}") from exc


def disable_picture_paste(hwp, report: Dict[str, Any]) -> None:
    try:
        properties = hwp.EngineProperties
        properties.SetItem("PasteObjectAsPicture", 0)
        hwp.EngineProperties = properties
    except Exception as exc:
        report["warnings"].append(f"PasteObjectAsPicture 설정을 적용하지 못했습니다: {exc}")


def find_hwp_table_control(hwp):
    for name in ("CurSelectedCtrl", "ParentCtrl", "LastCtrl"):
        try:
            control = getattr(hwp, name)
        except Exception:
            continue
        if is_hwp_table_control(control):
            return control
    return None


def last_hwp_control(hwp):
    try:
        return hwp.LastCtrl
    except Exception:
        return None


def hwp_table_dimensions(control) -> tuple[int, int]:
    try:
        properties = control.Properties
        return int(properties.Rows), int(properties.Cols)
    except Exception:
        return 0, 0


def apply_table_style_after_paste(hwp, profile: Dict[str, Any] | None, report: Dict[str, Any]) -> None:
    run_action(hwp, "TableCellBlock", report, "style")
    run_action(hwp, "TableCellBlockExtend", report, "style")
    apply_table_style_before_create(hwp, profile, report)
    run_action(hwp, "Cancel", report, "style")


def apply_table_matrix_styles(
    hwp,
    table: Dict[str, Any],
    style_indexes: Dict[str, int],
    style_config: Dict[str, Any],
    report: Dict[str, Any],
) -> None:
    rows = table_rows_for_hwp(table)
    if not rows:
        return
    run_action(hwp, "TableColBegin", report, "style")
    for _ in range(len(rows)):
        run_action(hwp, "TableUpperCell", report, "style")
    matrix = table.get("matrix") or []
    cells = []
    for row_index, row in enumerate(rows):
        for col_index, value in enumerate(row):
            cell = (
                matrix[row_index][col_index]
                if row_index < len(matrix) and col_index < len(matrix[row_index])
                else {}
            )
            if cell.get("covered_by"):
                continue
            cells.append((row_index, col_index, value, len(row), int(cell.get("rowspan") or 1), int(cell.get("colspan") or 1)))
    visited_lists = set()
    for position, (row_index, col_index, value, col_count, rowspan, colspan) in enumerate(cells):
        visited_lists.add(int(hwp.GetPos()[0]))
        style_name = style_name_for_cell_role(table_cell_role(table, row_index, col_index))
        # Select the cell's text list, not the cell block (Delete on a block is not text deletion).
        if not run_action(hwp, "MoveListBegin", report, "style"):
            raise HwpWriterError("style", "MoveListBegin", "붙여넣은 표 셀 텍스트 시작 위치를 찾지 못했습니다.")
        text_start = tuple(hwp.GetPos())
        if not run_action(hwp, "MoveSelListEnd", report, "style"):
            raise HwpWriterError("style", "MoveSelListEnd", "붙여넣은 표 셀 텍스트를 선택하지 못했습니다.")
        if tuple(hwp.GetPos()) != text_start and not run_action(hwp, "Delete", report, "style"):
            raise HwpWriterError("style", "Delete", "붙여넣은 표 셀 내용을 지우지 못했습니다.")
        apply_named_style(hwp, require_style_index(style_indexes, style_name), report, style_name)
        insert_text(hwp, value, report)
        apply_cell_appearance(
            hwp,
            style_config,
            style_name,
            row_index,
            col_index,
            len(rows),
            col_count,
            report,
            rowspan=rowspan,
            colspan=colspan,
        )
        if position < len(cells) - 1:
            # Vertical merges can revisit the same cell list on the next logical row.
            for _ in range(len(rows) * max(len(row) for row in rows)):
                if not run_action(hwp, "TableRightCell", report, "style"):
                    raise HwpWriterError("style", "TableRightCell", "붙여넣은 표의 다음 셀로 이동하지 못했습니다.")
                if not is_hwp_table_control(getattr(hwp, "ParentCtrl", None)):
                    raise HwpWriterError("style", "TableRightCell", "셀 이동 중 표 밖으로 나가 서식 적용을 중단합니다.")
                if int(hwp.GetPos()[0]) not in visited_lists:
                    break
            else:
                raise HwpWriterError("style", "TableRightCell", "병합 셀 이동이 반복되어 표 서식 적용을 중단합니다.")


def table_result(table: Dict[str, Any], insert_mode: str, ctrl_id: str, style_status: str, fallback_reason: str) -> Dict[str, Any]:
    return {
        "table_key": str(table.get("table_key") or ""),
        "insert_mode": insert_mode,
        "pasted_ctrl_id": ctrl_id,
        "rows": int(table.get("row_count") or len(table.get("matrix", []))),
        "cols": int(table.get("col_count") or max((len(row) for row in table.get("matrix", [])), default=0)),
        "style_status": style_status,
        "fallback_reason": fallback_reason,
        "part_index": table.get("part_index", 1),
        "part_count": table.get("part_count", 1),
        "source_columns": table.get("source_columns", []),
        "body_width_hwpunit": table.get("body_width_hwpunit"),
        "actual_width_hwpunit": table.get("actual_width_hwpunit"),
    }


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
    temp_output = Path(tempfile.gettempdir()) / f"report_automation_{time.time_ns()}{output_file.suffix}"
    attempts = [
        lambda: hwp.XHwpDocuments.Active_XHwpDocument.SaveAs(
            str(temp_output), format_name, "lock:false;backup:false;fullsave:true"
        ),
        lambda: hwp.SaveAs(str(temp_output), format_name, "lock:false;backup:false;fullsave:true"),
        lambda: hwp.SaveAs(str(temp_output)),
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
            if temp_output.exists():
                output_file.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(temp_output, output_file)
                report["document_saved"] = True
                record_com_step(report, "save_as", "ready", path=str(output_file))
                write_checkpoint(report, checkpoint_path)
                try:
                    temp_output.unlink()
                except OSError:
                    pass
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
        "style_config_path": "",
        "style_preset_name": "",
        "style_index_map": {},
        "style_application_counts": {},
        "working_template_path": "",
        "temporary_files_retained": False,
        "temporary_root_path": "",
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
        "table_results": [],
        "excel": {"opened": False, "source_workbook": "", "closed": False},
        "placeholders": {"body_found": False, "replaced": []},
        "com": {
            "dispatch_mode": "ensure_dispatch",
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
    parser.add_argument("--style-config")
    parser.add_argument("--keep-open-after-save", action="store_true")
    parser.add_argument("--dispatch-mode", choices=DISPATCH_MODES, default="ensure_dispatch")
    parser.add_argument("--check-environment", action="store_true")
    args = parser.parse_args(argv)

    try:
        if args.check_environment:
            report = check_environment(args.report_output, parse_bool(args.visible), args.dispatch_mode)
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
            args.dispatch_mode,
            args.style_config,
        )
        print(str(output))
        return 0
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
