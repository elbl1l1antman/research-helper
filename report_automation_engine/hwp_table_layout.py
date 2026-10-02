"""Plan printable HWP table fragments without changing source cells or font sizes."""

import copy
import math
import unicodedata

try:
    from .report_table_matrix import resolve_table_metadata
except ImportError:
    from report_table_matrix import resolve_table_metadata


def _text_units(text):
    return sum(1 if unicodedata.east_asian_width(char) in {"W", "F"} else 0.55 for char in str(text))


def _header_row_count(matrix):
    for index, row in enumerate(matrix):
        roles = {cell.get("role") for cell in row if not cell.get("covered_by")}
        if "value" in roles or (roles & {"stub", "banner_vertical"} and "banner_horizontal" not in roles):
            return index
    return len(matrix)


def prepare_report_table(table):
    """Add a render-only merged metadata row; Excel/source coordinates stay intact."""
    result = copy.deepcopy(table)
    matrix = result["matrix"]
    headers = _header_row_count(matrix)
    crosses_body = any(not cell.get("covered_by") and row_index < headers and row_index + int(cell.get("rowspan") or 1) > headers
                       for row_index, row in enumerate(matrix) for cell in row)
    if crosses_body:
        headers = 0
        result["header_repeat_warning"] = "헤더 병합이 본문까지 이어져 BASE 행만 반복합니다."
    metadata = resolve_table_metadata(table)
    result.update(metadata)
    cols = len(matrix[0])
    note = []
    for col in range(1, cols + 1):
        note.append({"row": 1, "col": col, "rowspan": 1, "colspan": cols if col == 1 else 1,
                     "role": "metadata" if col == 1 else "blank", "display_text": f"[BASE : {metadata['base_label']} | 단위 : {metadata['unit']}]" if col == 1 else "",
                     "raw_value": None, "covered_by": "" if col == 1 else "RA_METADATA", "source_cell": "", "is_numeric": False})
    for row in matrix:
        for cell in row:
            cell["row"] = int(cell["row"]) + 1
    for merged in result.get("merged_ranges", []):
        merged["row"] += 1
    if cols > 1:
        result.setdefault("merged_ranges", []).insert(0, {"row": 1, "col": 1, "rowspan": 1, "colspan": cols})
    matrix.insert(0, note)
    result.update(row_count=len(matrix), cells=[cell for row in matrix for cell in row], header_row_count=headers + 1, data_start_row=1)
    return result


def plan_table_parts(table, body_width, font_size_pt):
    """Repeat left labels/BASE, split data columns, and project header merges.

    Header banners may be repeated across fragments. Body merges are atomic:
    a value must never be silently divided or dropped to make a table fit.
    """
    if not math.isfinite(body_width) or body_width <= 0 or not math.isfinite(font_size_pt) or font_size_pt <= 0:
        raise ValueError("본문 폭과 표 글자 크기는 양수여야 합니다.")
    matrix = table.get("matrix") or []
    if not matrix or not matrix[0] or any(len(row) != len(matrix[0]) for row in matrix):
        raise ValueError("표 matrix가 비어 있거나 행별 열 수가 다릅니다.")
    count = len(matrix[0])
    header_rows = _header_row_count(matrix)
    prefix = max((col + 1 for row in matrix for col, cell in enumerate(row) if cell.get("role") in {"stub", "banner_vertical"} and not cell.get("covered_by")), default=0)
    while prefix < count and any(row[prefix].get("role") == "base" for row in matrix):
        prefix += 1

    widths = []
    for col in range(count):
        body = [cell for row in matrix[header_rows:] if not (cell := row[col]).get("covered_by")]
        labels = [cell for row in matrix[:header_rows] if not (cell := row[col]).get("covered_by") and int(cell.get("colspan") or 1) == 1]
        text_width = max((_text_units(cell.get("display_text", "")) for cell in body), default=0)
        if any(cell.get("role") in {"stub", "banner_vertical"} for cell in body):
            units = max(4, min(text_width, 4 if col == 0 and prefix > 1 else 9))
        else:
            units = max(4, text_width, min(5, max((_text_units(cell.get("display_text", "")) for cell in labels), default=0)))
        widths.append(math.ceil(units * font_size_pt * 100 + 600))

    # A small clearance avoids clipboard pixel rounding and outer border clipping.
    budget = int(body_width) - 500
    fixed = sum(widths[:prefix])
    forbidden, preferred = set(), set()
    for row_index, row in enumerate(matrix):
        for col, cell in enumerate(row):
            if cell.get("covered_by"):
                continue
            span = int(cell.get("colspan") or 1)
            if col + span > count or row_index + int(cell.get("rowspan") or 1) > len(matrix):
                raise ValueError("병합 범위가 표 영역을 벗어납니다.")
            if span > 1:
                if row_index >= header_rows:
                    forbidden.update(range(col + 1, col + span))
                elif col >= prefix:
                    preferred.add(col + span)

    selections = []
    if sum(widths) <= budget:
        selections.append(list(range(count)))
    else:
        if prefix in forbidden:
            raise ValueError("본문 병합 셀이 반복 보기/BASE 열 경계를 가로질러 분할할 수 없습니다.")
        if prefix == count or fixed >= budget:
            raise ValueError("반복 보기/BASE 열이 본문 폭에 들어가지 않습니다.")
        start = prefix
        while start < count:
            end, total = start, fixed
            fitting = []
            while end < count and total + widths[end] <= budget:
                total += widths[end]
                end += 1
                if end not in forbidden:
                    fitting.append(end)
            if not fitting:
                raise ValueError("값 열 또는 본문 병합 묶음이 본문 폭에 들어가지 않습니다.")
            cut = max(fitting)
            if cut < count:
                cut = max((edge for edge in fitting if edge in preferred), default=cut)
            selections.append(list(range(prefix)) + list(range(start, cut)))
            start = cut

    parts = []
    for index, columns in enumerate(selections, 1):
        part = _project_columns(table, columns)
        minimum = [widths[col] for col in columns]
        allocated = [int(value * budget / sum(minimum)) for value in minimum]
        allocated[-1] += budget - sum(allocated)
        part.update(source_columns=[col + 1 for col in columns], repeated_columns=list(range(1, prefix + 1)),
                    column_widths_hwpunit=allocated, body_width_hwpunit=int(body_width),
                    part_index=index, part_count=len(selections), layout_prepared=True)
        parts.append(part)
    return parts


def _project_columns(table, columns):
    original = table["matrix"]
    matrix = [[copy.deepcopy(row[col]) for col in columns] for row in original]
    for row_index, row in enumerate(matrix):
        for col, cell in enumerate(row):
            cell.update(row=row_index + 1, col=col + 1, rowspan=1, colspan=1, covered_by="")
    merged = []
    for row_index, row in enumerate(original):
        for col, cell in enumerate(row):
            if cell.get("covered_by"):
                continue
            intersection = [index for index, source in enumerate(columns) if col <= source < col + int(cell.get("colspan") or 1)]
            if not intersection:
                continue
            target_col = intersection[0]
            rowspan, colspan = int(cell.get("rowspan") or 1), len(intersection)
            anchor = copy.deepcopy(cell)
            anchor.update(row=row_index + 1, col=target_col + 1, rowspan=rowspan, colspan=colspan, covered_by="",
                          source_row=row_index + 1, source_col=col + 1)
            matrix[row_index][target_col] = anchor
            for r in range(row_index, row_index + rowspan):
                for c in intersection:
                    if (r, c) != (row_index, target_col):
                        matrix[r][c].update(display_text="", role="blank", covered_by=anchor.get("source_cell") or f"R{row_index + 1}C{col + 1}")
            if rowspan > 1 or colspan > 1:
                merged.append({"row": row_index + 1, "col": target_col + 1, "rowspan": rowspan, "colspan": colspan})
    return {**table, "matrix": matrix, "cells": [cell for row in matrix for cell in row],
            "row_count": len(matrix), "col_count": len(columns), "merged_ranges": merged}
