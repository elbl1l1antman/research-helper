from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
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
        parsed = Decimal(str(value).replace(",", "").replace("%", "").strip())
    except (InvalidOperation, TypeError, ValueError):
        return str(value).strip()
    if parsed == parsed.to_integral_value():
        text = f"{parsed:,.0f}"
    else:
        places = max(decimal_places, 0)
        rounded = parsed.quantize(Decimal("1").scaleb(-places), rounding=ROUND_HALF_UP)
        text = f"{rounded:,.{places}f}"
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
