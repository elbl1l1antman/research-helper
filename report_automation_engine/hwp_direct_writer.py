from __future__ import annotations

import argparse
import html
import json
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List


BODY_PLACEHOLDER = "{{BODY}}"


class HwpDirectWriterError(RuntimeError):
    def __init__(self, stage: str, action: str, message: str) -> None:
        super().__init__(message)
        self.stage = stage
        self.action = action


def write_hwpx_direct(
    package_path: str | Path,
    preflight_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
    report_path: str | Path | None = None,
) -> Path:
    package_file = Path(package_path).resolve()
    preflight_file = Path(preflight_path).resolve()
    template_file = Path(template_path).resolve()
    output_file = Path(output_path).resolve()
    report_file = (
        Path(report_path).resolve()
        if report_path
        else output_file.with_name(output_file.stem + "_hwp_direct_writer_report.json")
    )
    report = new_report(package_file, preflight_file, template_file, output_file)
    try:
        package = load_json(package_file)
        preflight = load_json(preflight_file)
        if preflight.get("status") == "blocked":
            raise HwpDirectWriterError("validate", "preflight", "preflight status가 blocked입니다.")
        if not template_file.exists():
            raise HwpDirectWriterError("validate", "template", f"템플릿 파일을 찾을 수 없습니다: {template_file}")
        if template_file.resolve() == output_file.resolve():
            raise HwpDirectWriterError("validate", "output", "원본 템플릿과 출력 경로가 같습니다.")
        output_file.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_file, output_file)
        replace_body_payload(output_file, build_body_payload(package), report)
        report["status"] = "ready"
        report["finished_at"] = now()
        write_json(report_file, report)
        return output_file
    except HwpDirectWriterError as exc:
        report["status"] = "failed"
        report["stage"] = exc.stage
        report["action"] = exc.action
        report["errors"].append(str(exc))
        report["finished_at"] = now()
        write_json(report_file, report)
        raise
    except Exception as exc:
        report["status"] = "failed"
        report["stage"] = report.get("stage") or "unknown"
        report["action"] = report.get("action") or "unknown"
        report["errors"].append(str(exc))
        report["finished_at"] = now()
        write_json(report_file, report)
        raise


def replace_body_payload(hwpx_path: Path, payload: str, report: Dict[str, Any]) -> None:
    if not zipfile.is_zipfile(hwpx_path):
        raise HwpDirectWriterError("validate", "template", "HWPX zip 패키지가 아닙니다.")
    temp_path = hwpx_path.with_suffix(hwpx_path.suffix + ".tmp")
    replaced = False
    with zipfile.ZipFile(hwpx_path, "r") as src, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename.lower().endswith(".xml"):
                text = data.decode("utf-8", errors="ignore")
                if BODY_PLACEHOLDER in text:
                    text = text.replace(BODY_PLACEHOLDER, payload, 1)
                    data = text.encode("utf-8")
                    replaced = True
                    report["body_xml_part"] = info.filename
            dst.writestr(info, data)
    temp_path.replace(hwpx_path)
    if not replaced:
        raise HwpDirectWriterError("template", "find_body", "{{BODY}} placeholder를 찾지 못했습니다.")


def build_body_payload(package: Dict[str, Any]) -> str:
    tables = {str(table.get("table_key") or ""): table for table in package.get("tables", [])}
    parts: List[str] = []
    for index, section in enumerate(package.get("sections", []), start=1):
        key = str(section.get("table_key") or "")
        title = str(section.get("title") or key or f"문항 {index}")
        narrative = str(section.get("narrative_final") or "")
        parts.append(f"<ra-section-title>{escape(title)}</ra-section-title>")
        parts.append(f"<ra-narrative>{escape(narrative)}</ra-narrative>")
        table = tables.get(key)
        if table:
            parts.append(matrix_to_xml(table))
        parts.append(f"<ra-source>{escape(key)}</ra-source>")
    return "".join(parts)


def matrix_to_xml(table: Dict[str, Any]) -> str:
    rows = []
    for row in table.get("matrix", []):
        cells = "".join(
            f"<ra-cell role=\"{escape_attr(cell.get('role'))}\">{escape(cell.get('display_text'))}</ra-cell>"
            for cell in row
        )
        rows.append(f"<ra-row>{cells}</ra-row>")
    return f"<ra-table key=\"{escape_attr(table.get('table_key'))}\" title=\"{escape_attr(table.get('title'))}\">{''.join(rows)}</ra-table>"


def escape(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=False)


def escape_attr(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def load_json(path: str | Path) -> Dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_json(path: str | Path, value: Dict[str, Any]) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def new_report(package_file: Path, preflight_file: Path, template_file: Path, output_file: Path) -> Dict[str, Any]:
    return {
        "schema_version": "1.0",
        "writer": "hwp_direct_writer",
        "status": "started",
        "stage": "",
        "action": "",
        "started_at": now(),
        "finished_at": "",
        "package_path": str(package_file),
        "preflight_path": str(preflight_file),
        "template_path": str(template_file),
        "output_path": str(output_file),
        "body_xml_part": "",
        "warnings": [],
        "errors": [],
    }


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create a minimal HWPX draft by replacing {{BODY}} inside the HWPX package.")
    parser.add_argument("--package", required=True)
    parser.add_argument("--preflight", required=True)
    parser.add_argument("--template", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--report-output")
    args = parser.parse_args(argv)
    try:
        output = write_hwpx_direct(args.package, args.preflight, args.template, args.output, args.report_output)
        print(str(output))
        return 0
    except Exception as exc:
        print(str(exc))
        return 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
