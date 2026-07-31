# HWPX Table Stabilization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Excel 산출표를 HWPX writer가 안전하게 사용할 수 있는 table matrix 계약으로 정규화하고, preflight와 최소 직접 HWPX writer를 붙인다.

**Architecture:** 기존 `report_package.json` schema v1은 유지하고 `tables[].matrix`와 table QA만 추가한다. HWPX 직접 writer는 COM writer를 대체하지 않고, HWP 설치 없이 검증 가능한 최소 직접 출력 경로로 추가한다.

**Tech Stack:** Python 3, openpyxl, zipfile, xml.etree.ElementTree, existing `report_automation_engine` package, Windows launcher integration later.

## Global Constraints

- HWPX가 최우선 산출물이다.
- PPTX와 대시보드 신규 개발은 이번 계획 범위에서 제외한다.
- 기존 `tables[].rows` 계약은 깨지 않는다.
- HWPX에 삽입되는 값은 `display_text`를 우선한다.
- raw numeric value는 QA/분석용으로 보존하되 문서 삽입 문자열로 직접 쓰지 않는다.
- 원본 Excel과 원본 HWPX 템플릿은 수정하지 않는다.
- COM writer는 fallback/진단용으로 유지한다.
- 새 Python 파일은 UTF-8, ASCII 중심 식별자, 한국어 사용자 메시지는 문자열로만 둔다.
- 각 task는 테스트 또는 `compileall` 검증을 포함한다.

---

## File Structure

- Create `report_automation_engine/report_table_matrix.py`
  - Existing simple table rows를 matrix/cell 계약으로 변환한다.
  - display value formatting, role inference, table QA를 담당한다.
- Modify `report_automation_engine/report_package.py`
  - `group_table_rows()` 결과에 matrix contract를 추가한다.
  - `build_preflight()` summary와 warning/error에 table matrix QA를 반영한다.
- Create `report_automation_engine/hwp_direct_writer.py`
  - package/preflight/template/output을 받아 최소 HWPX zip 출력과 writer report를 만든다.
  - `{{BODY}}` placeholder를 section/table payload로 치환한다.
- Create `tests/test_report_table_matrix.py`
  - matrix conversion과 display value 우선 정책을 검증한다.
- Create `tests/test_report_package_table_matrix.py`
  - fake Excel 산출 시트에서 package/preflight table matrix 생성 여부를 검증한다.
- Create `tests/test_hwp_direct_writer.py`
  - 최소 HWPX zip 템플릿에서 placeholder 제거와 display value 보존을 검증한다.

### Task 1: Table Matrix Contract

**Files:**
- Create: `report_automation_engine/report_table_matrix.py`
- Create: `tests/test_report_table_matrix.py`

**Interfaces:**
- Produces: `build_table_matrix(table: dict, decimal_places: int = 1) -> dict`
- Produces: `format_display_value(value: object, unit: str = "", decimal_places: int = 1) -> str`
- Produces: `table_matrix_qa(table_matrix: dict) -> list[dict]`

- [ ] **Step 1: Write failing tests**

Create `tests/test_report_table_matrix.py` with these tests:

```python
from report_automation_engine.report_table_matrix import build_table_matrix, format_display_value


def test_format_display_value_uses_one_decimal_for_float():
    assert format_display_value(3.333535353, decimal_places=1) == "3.3"


def test_format_display_value_preserves_percent_unit():
    assert format_display_value(63.25, unit="%", decimal_places=1) == "63.3%"


def test_build_table_matrix_adds_display_cells_without_losing_rows():
    table = {
        "table_key": "T001",
        "title": "만족도",
        "rows": [
            {"category": "전체", "percent": 63.25, "weighted_n": 1200, "raw_n": 1198, "unit": "%", "source_cell": "D5"},
            {"category": "매우 만족", "percent": 29.44, "weighted_n": 353, "raw_n": 351, "unit": "%", "source_cell": "D6"},
        ],
    }
    matrix_table = build_table_matrix(table, decimal_places=1)
    assert matrix_table["row_count"] == 3
    assert matrix_table["col_count"] == 4
    assert matrix_table["matrix"][0][0]["role"] == "header"
    assert matrix_table["matrix"][1][1]["display_text"] == "63.3%"
    assert matrix_table["matrix"][1][1]["raw_value"] == 63.25
    assert matrix_table["matrix"][1][1]["source_cell"] == "D5"
    assert matrix_table["roles"]["value"] >= 1
```

- [ ] **Step 2: Run tests to verify failure**

Run: `python -m pytest tests/test_report_table_matrix.py -q`

Expected: fail because `report_table_matrix.py` does not exist.

- [ ] **Step 3: Implement module**

Create `report_automation_engine/report_table_matrix.py` with:

```python
from __future__ import annotations

from typing import Any, Dict, List


HEADER_LABELS = ["항목", "비율", "가중 N", "원 N"]
LONG_TEXT_LIMIT = 40
MANY_COLUMNS_LIMIT = 8


def build_table_matrix(table: Dict[str, Any], decimal_places: int = 1) -> Dict[str, Any]:
    rows = list(table.get("rows", []))
    matrix: List[List[Dict[str, Any]]] = [header_row()]
    for row_index, source in enumerate(rows, start=2):
        unit = str(source.get("unit") or "")
        matrix.append(
            [
                make_cell(row_index, 1, "stub", source.get("category"), source.get("category"), "", source.get("source_cell"), "left"),
                make_cell(row_index, 2, "value", format_display_value(source.get("percent"), unit, decimal_places), source.get("percent"), "0.0", source.get("source_cell"), "right"),
                make_cell(row_index, 3, "value", format_display_value(source.get("weighted_n"), "", decimal_places), source.get("weighted_n"), "#,##0", source.get("source_cell"), "right"),
                make_cell(row_index, 4, "value", format_display_value(source.get("raw_n"), "", decimal_places), source.get("raw_n"), "#,##0", source.get("source_cell"), "right"),
            ]
        )

    cells = [cell for row in matrix for cell in row]
    result = {
        **table,
        "source_sheet": table.get("source_sheet", "보고서_삽입표"),
        "source_range": table.get("source_range", ""),
        "row_count": len(matrix),
        "col_count": max((len(row) for row in matrix), default=0),
        "matrix": matrix,
        "cells": cells,
        "merged_ranges": list(table.get("merged_ranges", [])),
        "roles": role_counts(cells),
        "style_hints": {
            "table_width": "body",
            "wrap_text": True,
            "header_fill": "#E7E7E7",
            "font_size_pt": 9,
        },
    }
    result["qa"] = table_matrix_qa(result)
    return result


def header_row() -> List[Dict[str, Any]]:
    return [
        make_cell(1, index, "header", label, label, "", "", "center")
        for index, label in enumerate(HEADER_LABELS, start=1)
    ]


def make_cell(row: int, col: int, role: str, display_text: Any, raw_value: Any, number_format: str, source_cell: Any, align: str) -> Dict[str, Any]:
    text = "" if display_text is None else str(display_text)
    return {
        "row": row,
        "col": col,
        "rowspan": 1,
        "colspan": 1,
        "role": role,
        "display_text": text,
        "raw_value": raw_value,
        "number_format": number_format,
        "source_cell": "" if source_cell is None else str(source_cell),
        "align": align,
        "is_numeric": is_number(raw_value),
    }


def format_display_value(value: Any, unit: str = "", decimal_places: int = 1) -> str:
    if value in (None, ""):
        return ""
    try:
        parsed = float(str(value).replace(",", "").replace("%", "").strip())
    except (TypeError, ValueError):
        return str(value).strip()
    if parsed.is_integer():
        text = f"{parsed:,.0f}"
    else:
        text = f"{parsed:,.{max(decimal_places, 0)}f}"
    return text + (unit or "")


def table_matrix_qa(table_matrix: Dict[str, Any]) -> List[Dict[str, str]]:
    qa: List[Dict[str, str]] = []
    if int(table_matrix.get("row_count") or 0) == 0 or int(table_matrix.get("col_count") or 0) == 0:
        qa.append(issue("error", "표 matrix가 비어 있습니다."))
    if int(table_matrix.get("col_count") or 0) > MANY_COLUMNS_LIMIT:
        qa.append(issue("warning", "열 수가 많아 HWPX 본문 폭을 초과할 수 있습니다."))
    if not str(table_matrix.get("title") or "").strip():
        qa.append(issue("warning", "표 제목이 없습니다."))
    if not str(table_matrix.get("source_range") or "").strip():
        qa.append(issue("warning", "표 source_range가 없습니다."))
    value_cells = [cell for cell in table_matrix.get("cells", []) if cell.get("role") == "value"]
    if value_cells and all(not str(cell.get("display_text") or "").strip() for cell in value_cells):
        qa.append(issue("error", "값 영역이 모두 비어 있습니다."))
    for cell in table_matrix.get("cells", []):
        if not str(cell.get("display_text") or "").strip() and cell.get("raw_value") not in (None, ""):
            qa.append(issue("error", f"{cell.get('row')}행 {cell.get('col')}열 display_text가 없습니다."))
        if len(str(cell.get("display_text") or "")) > LONG_TEXT_LIMIT:
            qa.append(issue("warning", f"{cell.get('row')}행 {cell.get('col')}열 텍스트가 길어 줄바꿈됩니다."))
    return qa


def role_counts(cells: List[Dict[str, Any]]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for cell in cells:
        role = str(cell.get("role") or "unknown")
        counts[role] = counts.get(role, 0) + 1
    return counts


def is_number(value: Any) -> bool:
    if value in (None, ""):
        return False
    try:
        float(str(value).replace(",", "").replace("%", "").strip())
        return True
    except (TypeError, ValueError):
        return False


def issue(severity: str, message: str) -> Dict[str, str]:
    return {"type": "table_matrix", "severity": severity, "message": message}
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_report_table_matrix.py -q`

Expected: all tests pass.

- [ ] **Step 5: Commit**

Run:

```bash
git add report_automation_engine/report_table_matrix.py tests/test_report_table_matrix.py
git commit -m "feat: add HWPX table matrix contract"
```

### Task 2: Package and Preflight Integration

**Files:**
- Modify: `report_automation_engine/report_package.py`
- Create: `tests/test_report_package_table_matrix.py`

**Interfaces:**
- Consumes: `build_table_matrix(table: dict, decimal_places: int = 1) -> dict`
- Produces: `report_package.json` tables containing `matrix`, `cells`, `roles`, `qa`
- Produces: `preflight_report.json` summary fields `table_matrix_warning_count`, `table_matrix_error_count`

- [ ] **Step 1: Write failing integration test**

Create `tests/test_report_package_table_matrix.py`:

```python
from pathlib import Path

import openpyxl

from report_automation_engine.report_package import build_preflight, build_report_package


def make_workbook(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "보고서_분석문"
    ws.append(["table_key", "문항/표 제목", "분석문_기본", "최종 사용문", "사용자 수정문", "주요 수치 요약", "검토 상태"])
    ws.append(["T001", "만족도", "만족도는 63.3%로 나타남.", "", "", "63.3%", ""])
    table_ws = wb.create_sheet("보고서_삽입표")
    table_ws.append(["table_key", "title", "category", "weighted_n", "raw_n", "percent", "unit", "source_cell"])
    table_ws.append(["T001", "만족도", "전체", 1200, 1198, 63.25, "%", "D5"])
    table_ws.append(["T001", "만족도", "매우 만족", 353, 351, 29.44, "%", "D6"])
    wb.save(path)


def test_package_tables_include_matrix(tmp_path):
    excel = tmp_path / "sample.xlsx"
    make_workbook(excel)
    package = build_report_package(excel, {"decimal_places": "1"})
    table = package["tables"][0]
    assert table["matrix"][1][1]["display_text"] == "63.3%"
    assert table["matrix"][1][1]["raw_value"] == 63.25
    assert table["roles"]["header"] == 4


def test_preflight_counts_table_matrix_warnings(tmp_path):
    excel = tmp_path / "sample.xlsx"
    make_workbook(excel)
    package = build_report_package(excel, {"decimal_places": "1"})
    preflight = build_preflight(package, [])
    assert "table_matrix_warning_count" in preflight["summary"]
    assert preflight["summary"]["table_matrix_error_count"] == 0
```

- [ ] **Step 2: Run test to verify failure**

Run: `python -m pytest tests/test_report_package_table_matrix.py -q`

Expected: fail because package tables do not include matrix fields.

- [ ] **Step 3: Modify `report_package.py` imports**

Add fallback import near existing template inspector import:

```python
try:
    from .report_table_matrix import build_table_matrix
except ImportError:
    from report_table_matrix import build_table_matrix
```

- [ ] **Step 4: Pass decimal places into table grouping**

In `build_report_package()`, parse decimal places from meta:

```python
decimal_places = int(str((meta or {}).get("decimal_places", "1") or "1"))
tables = group_table_rows(read_rows(wb, "보고서_삽입표"), decimal_places)
```

Change signature:

```python
def group_table_rows(rows: List[Dict[str, Any]], decimal_places: int = 1) -> List[Dict[str, Any]]:
```

At the end of grouping, return:

```python
return [build_table_matrix(table, decimal_places) for table in grouped.values()]
```

- [ ] **Step 5: Add table matrix QA to contract QA**

In `add_contract_qa()`, for each table, append its `qa` items to package-level QA with `table_key`:

```python
for qa_item in table.get("qa", []):
    package["qa"].append(
        issue(table["table_key"], "table_matrix", qa_item.get("severity", "warning"), qa_item.get("message", ""))
    )
```

- [ ] **Step 6: Add preflight summary counts**

In `build_preflight()`, calculate:

```python
table_matrix_warnings = [q for q in warnings if clean(q.get("type")) == "table_matrix"]
table_matrix_errors = [q for q in errors if clean(q.get("type")) == "table_matrix"]
```

Add summary keys:

```python
"table_matrix_warning_count": len(table_matrix_warnings),
"table_matrix_error_count": len(table_matrix_errors),
```

- [ ] **Step 7: Run tests**

Run:

```bash
python -m pytest tests/test_report_table_matrix.py tests/test_report_package_table_matrix.py -q
python -m compileall -q report_automation_engine
```

Expected: all pass.

- [ ] **Step 8: Commit**

Run:

```bash
git add report_automation_engine/report_package.py tests/test_report_package_table_matrix.py
git commit -m "feat: include table matrix in report package"
```

### Task 3: Minimal HWPX Direct Writer

**Files:**
- Create: `report_automation_engine/hwp_direct_writer.py`
- Create: `tests/test_hwp_direct_writer.py`

**Interfaces:**
- Consumes: package JSON with `sections[]` and `tables[].matrix`
- Consumes: preflight JSON with `status`
- Produces: `write_hwpx_direct(package_path, preflight_path, template_path, output_path, report_path=None) -> Path`
- Produces CLI: `python -m report_automation_engine.hwp_direct_writer --package ... --preflight ... --template ... --output ...`

- [ ] **Step 1: Write failing writer tests**

Create `tests/test_hwp_direct_writer.py`:

```python
import json
import zipfile
from pathlib import Path

from report_automation_engine.hwp_direct_writer import write_hwpx_direct


def make_template(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("mimetype", "application/hwp+zip")
        zf.writestr("Contents/section0.xml", "<?xml version='1.0' encoding='UTF-8'?><doc><p>{{BODY}}</p></doc>")


def make_package(path: Path) -> None:
    package = {
        "schema_version": "1.0",
        "meta": {"source_file_name": "sample.xlsx"},
        "sections": [{"table_key": "T001", "title": "만족도", "narrative_final": "만족도는 63.3%로 나타남."}],
        "tables": [
            {
                "table_key": "T001",
                "title": "만족도",
                "matrix": [
                    [{"display_text": "항목", "role": "header"}, {"display_text": "비율", "role": "header"}],
                    [{"display_text": "전체", "role": "stub"}, {"display_text": "63.3%", "role": "value", "raw_value": 63.25}],
                ],
            }
        ],
        "qa": [],
    }
    path.write_text(json.dumps(package, ensure_ascii=False), encoding="utf-8")


def test_write_hwpx_direct_replaces_body_and_uses_display_text(tmp_path):
    template = tmp_path / "template.hwpx"
    package = tmp_path / "package.json"
    preflight = tmp_path / "preflight.json"
    output = tmp_path / "output.hwpx"
    make_template(template)
    make_package(package)
    preflight.write_text(json.dumps({"status": "ready"}, ensure_ascii=False), encoding="utf-8")
    write_hwpx_direct(package, preflight, template, output)
    with zipfile.ZipFile(output) as zf:
        content = "\n".join(zf.read(name).decode("utf-8") for name in zf.namelist() if name.endswith(".xml"))
    assert "{{BODY}}" not in content
    assert "만족도는 63.3%로 나타남." in content
    assert "63.3%" in content
    assert "63.25" not in content
```

- [ ] **Step 2: Run test to verify failure**

Run: `python -m pytest tests/test_hwp_direct_writer.py -q`

Expected: fail because `hwp_direct_writer.py` does not exist.

- [ ] **Step 3: Implement direct writer**

Create `report_automation_engine/hwp_direct_writer.py` with these functions:

```python
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


def write_hwpx_direct(package_path: str | Path, preflight_path: str | Path, template_path: str | Path, output_path: str | Path, report_path: str | Path | None = None) -> Path:
    package_file = Path(package_path).resolve()
    preflight_file = Path(preflight_path).resolve()
    template_file = Path(template_path).resolve()
    output_file = Path(output_path).resolve()
    report_file = Path(report_path).resolve() if report_path else output_file.with_name(output_file.stem + "_hwp_direct_writer_report.json")
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
        cells = "".join(f"<ra-cell role=\"{escape_attr(cell.get('role'))}\">{escape(cell.get('display_text'))}</ra-cell>" for cell in row)
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
```

- [ ] **Step 4: Run tests**

Run:

```bash
python -m pytest tests/test_hwp_direct_writer.py -q
python -m compileall -q report_automation_engine
```

Expected: all pass.

- [ ] **Step 5: Commit**

Run:

```bash
git add report_automation_engine/hwp_direct_writer.py tests/test_hwp_direct_writer.py
git commit -m "feat: add minimal direct HWPX writer"
```

### Task 4: Documentation and Priority Cleanup

**Files:**
- Modify: `README.md`
- Modify: `report_automation_engine/README.md`
- Modify: `docs/next_development_plan.md`

**Interfaces:**
- Consumes: `report_table_matrix.py`
- Consumes: `hwp_direct_writer.py`
- Produces: documentation stating HWPX table stabilization is current priority and PPT/dashboard are maintained but not active priority.

- [ ] **Step 1: Update root README**

Add a short section near “개발 방향”:

```markdown
현재 신규 개발 우선순위는 HWPX 표 안정화입니다. PPTX 보고서와 대시보드 PPTX는 기존 기능을 유지하지만, 새 기능 개발은 HWPX 표 계약과 writer 안정화 이후로 둡니다.
```

Add CLI example:

```powershell
python -m report_automation_engine.hwp_direct_writer `
  --package "C:\path\report_package.json" `
  --preflight "C:\path\preflight_report.json" `
  --template "C:\path\template.hwpx" `
  --output "C:\path\report_draft.hwpx"
```

- [ ] **Step 2: Update engine README**

Add bullets for:

- `report_table_matrix.py`
- `hwp_direct_writer.py`

State that direct writer v1 writes a minimal HWPX package payload for automated validation and that visual fidelity will be hardened after table XML validation.

- [ ] **Step 3: Update next development plan**

Replace the latest “다음 개발 우선순위” with:

```markdown
1. Report Package table matrix v2 계약 안정화
2. HWPX table preflight 차단/경고 정밀화
3. HWPX 직접 writer 최소 경로 검증
4. 실제 사용자 HWPX 템플릿 1장 생성 검수
5. COM writer는 fallback/진단용으로 정리
```

- [ ] **Step 4: Run verification**

Run:

```bash
python -m compileall -q report_automation_engine
git diff --check
```

Expected: both pass.

- [ ] **Step 5: Commit**

Run:

```bash
git add README.md report_automation_engine/README.md docs/next_development_plan.md
git commit -m "docs: refocus alpha plan on HWPX table stability"
```

## Final Verification

After all tasks:

```bash
python -m pytest tests -q
python -m compileall -q report_automation_engine
git status --short
```

Expected:

- tests pass
- compileall passes
- only intentional committed changes remain

## Self-Review

- Spec coverage: all design requirements are covered by Tasks 1-4.
- Placeholder scan: this plan contains no TBD/TODO placeholders.
- Type consistency: `build_table_matrix`, `format_display_value`, `write_hwpx_direct` signatures are consistent across tasks.
