"""Plan printable HWP table fragments without changing source cells or font sizes."""

import copy
import math
import unicodedata

try:
    from .report_table_matrix import resolve_table_metadata
    from .hwp_style_config import margin_units, style_name_for_cell_role
except ImportError:
    from report_table_matrix import resolve_table_metadata
    from hwp_style_config import margin_units, style_name_for_cell_role


def _text_units(text):
    return sum(1 if unicodedata.east_asian_width(char) in {"W", "F"} else 0.55 for char in str(text))


def _header_row_count(matrix):
    for index, row in enumerate(matrix):
        roles = {cell.get("role") for cell in row if not cell.get("covered_by")}
        if "value" in roles or (roles & {"stub", "banner_vertical"} and "banner_horizontal" not in roles):
            return index
    return len(matrix)


def _merged_ranges_from_matrix(matrix):
    return [{"row": r + 1, "col": c + 1, "rowspan": int(cell.get("rowspan") or 1), "colspan": int(cell.get("colspan") or 1)}
            for r, row in enumerate(matrix) for c, cell in enumerate(row)
            if not cell.get("covered_by") and (int(cell.get("rowspan") or 1) > 1 or int(cell.get("colspan") or 1) > 1)]


def _cell_id(cell, row_index, col_index):
    return cell.get("source_cell") or f"R{row_index + 1}C{col_index + 1}"


def apply_vertical_banner_policy(table, policy="merge"):
    """Optionally unmerge vertical banners so every page of a long table shows the label.

    Hancom prints a merged cell's text once, on the page where the merge starts.
    ``repeat_per_row`` writes the banner label into every row it covered instead.
    Values are never moved; only label cells (``banner_vertical``) change.
    """
    if policy == "merge":
        return table
    if policy != "repeat_per_row":
        raise ValueError(f"알 수 없는 세로 배너 정책입니다: {policy}")
    result = copy.deepcopy(table)
    matrix = result.get("matrix") or []
    headers = _header_row_count(matrix)
    changed = 0
    for row_index in range(headers, len(matrix)):
        for col_index, cell in enumerate(matrix[row_index]):
            rowspan = int(cell.get("rowspan") or 1)
            if cell.get("covered_by") or cell.get("role") != "banner_vertical" or rowspan <= 1:
                continue
            colspan = int(cell.get("colspan") or 1)
            cell["rowspan"] = 1
            for r in range(row_index + 1, min(row_index + rowspan, len(matrix))):
                target = matrix[r][col_index]
                anchor = copy.deepcopy(cell)
                anchor.update(row=target.get("row", r + 1), col=target.get("col", col_index + 1), rowspan=1, colspan=colspan,
                              covered_by="", source_cell=target.get("source_cell") or "", repeated_label=True)
                matrix[r][col_index] = anchor
                for c in range(col_index + 1, col_index + colspan):
                    matrix[r][c].update(covered_by=_cell_id(anchor, r, col_index), display_text="", role="blank")
            changed += 1
    if not changed:
        return table
    result["merged_ranges"] = _merged_ranges_from_matrix(matrix)
    result["cells"] = [cell for row in matrix for cell in row]
    result["vertical_banner_repeated"] = changed
    return result


def _split_merges_at_header(matrix, headers):
    """Cut label merges that run from the header into the body at the header boundary."""
    crossing = [(r, c, cell) for r, row in enumerate(matrix[:headers]) for c, cell in enumerate(row)
                if not cell.get("covered_by") and r + int(cell.get("rowspan") or 1) > headers]
    if any(cell.get("role") in {"value", "base"} for _, _, cell in crossing):
        return False
    for r, c, cell in crossing:
        end = r + int(cell["rowspan"])
        colspan = int(cell.get("colspan") or 1)
        cell["rowspan"] = headers - r
        body = matrix[headers][c]
        remainder = copy.deepcopy(body)
        remainder.update(rowspan=end - headers, colspan=colspan, covered_by="", display_text="",
                         role="blank", split_from=_cell_id(cell, r, c))
        matrix[headers][c] = remainder
        anchor_id = _cell_id(remainder, headers, c)
        for rr in range(headers, end):
            for cc in range(c, c + colspan):
                if (rr, cc) != (headers, c):
                    matrix[rr][cc].update(covered_by=anchor_id, display_text="", role="blank")
    return True


def prepare_report_table(table, header_merge_policy="base_only"):
    """Add a render-only merged metadata row; Excel/source coordinates stay intact."""
    if header_merge_policy not in ("base_only", "split_at_header"):
        raise ValueError(f"알 수 없는 헤더 병합 정책입니다: {header_merge_policy}")
    result = copy.deepcopy(table)
    matrix = result["matrix"]
    # Hancom collapses a row with no cell anchor, even on a native TableMergeCell.
    # A narrow, empty edge column keeps the original data coordinates unchanged.
    protected = [i + 2 for i, row in enumerate(matrix) if all(cell.get("covered_by") for cell in row)]
    if protected:
        cols = len(matrix[0])
        widths = result.get("column_widths_hwpunit")
        if not widths or min(widths) <= 1300:
            raise ValueError("병합행 보존용 여백을 확보할 수 없습니다.")
        guard_width = 1200
        budget = sum(widths) - guard_width
        adjusted = [width * budget // sum(widths) for width in widths]
        adjusted[-1] += budget - sum(adjusted)
        result["column_widths_hwpunit"] = [*adjusted, guard_width]
        result["col_count"] = cols + 1
        result["protected_merge_rows"] = protected
        for i, row in enumerate(matrix, 1):
            row.append({"row": i, "col": cols + 1, "rowspan": 1, "colspan": 1,
                        "role": "merge_guard", "display_text": "", "covered_by": "", "source_cell": ""})
    headers = _header_row_count(matrix)
    # Repeating a merged header can be unsafe even when its visual role is known.
    result["visual_header_row"] = headers if headers else None
    crosses_body = any(not cell.get("covered_by") and row_index < headers and row_index + int(cell.get("rowspan") or 1) > headers
                       for row_index, row in enumerate(matrix) for cell in row)
    if crosses_body and header_merge_policy == "split_at_header" and _split_merges_at_header(matrix, headers):
        result["merged_ranges"] = _merged_ranges_from_matrix(matrix)
        result["header_repeat_warning"] = "헤더와 본문에 걸친 병합 셀을 헤더 경계에서 나누어 헤더 행을 반복합니다."
    elif crosses_body:
        headers = 0
        result["header_repeat_warning"] = (
            "헤더 병합이 본문까지 이어져 BASE 행만 반복합니다." if header_merge_policy == "base_only"
            else "헤더와 본문에 걸친 병합에 값 셀이 있어 나누지 않고 BASE 행만 반복합니다.")
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


def analyze_table_columns(table, font_size_pt, style_config=None, layout=None):
    """Measure minimum column widths and where a table may be split.

    Returns ``count``, ``header_rows``, ``prefix`` (repeated stub/BASE columns),
    ``widths`` (HWPUNIT, user-fixed widths applied), ``fixed_columns``,
    ``forbidden`` and ``preferred`` cut edges (0-based index of the first column
    of a new fragment), ``body_merges`` and ``horizontal_margin(cell, row)``.
    """
    layout = layout or {}
    body_policy = layout.get("body_merge_policy", "block")
    matrix = table.get("matrix") or []
    if not matrix or not matrix[0] or any(len(row) != len(matrix[0]) for row in matrix):
        raise ValueError("표 matrix가 비어 있거나 행별 열 수가 다릅니다.")
    count = len(matrix[0])
    header_rows = _header_row_count(matrix)
    cell_styles = (style_config or {}).get("table_cell_styles", {})

    def horizontal_margin(cell, row):
        role = cell.get("role")
        if role == "banner_horizontal" and row == header_rows - 1:
            role = "header"
        margins = cell_styles.get(style_name_for_cell_role(role, style_config), {}).get("margins_mm")
        return sum(margin_units(margins[side]) for side in ("left", "right")) if margins else 0
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
        padding = max((horizontal_margin(row[col], index) for index, row in enumerate(matrix)
                       if not row[col].get("covered_by") and int(row[col].get("colspan") or 1) == 1), default=0)
        widths.append(math.ceil(units * font_size_pt * 100 + max(600, padding)))
    fixed_columns = set()
    for col, mm in (layout.get("column_widths_mm") or {}).items():
        col = int(col)
        if not 1 <= col <= count:
            raise ValueError(f"열 폭 보정의 {col}열이 표 범위(1~{count}열)를 벗어납니다.")
        widths[col - 1] = margin_units(mm)
        fixed_columns.add(col - 1)

    forbidden, preferred = set(), set()
    body_merges = []
    for row_index, row in enumerate(matrix):
        for col, cell in enumerate(row):
            if cell.get("covered_by"):
                continue
            span = int(cell.get("colspan") or 1)
            if col + span > count or row_index + int(cell.get("rowspan") or 1) > len(matrix):
                raise ValueError("병합 범위가 표 영역을 벗어납니다.")
            if span > 1:
                if row_index >= header_rows:
                    body_merges.append((row_index, col, span, cell))
                    if body_policy == "block":
                        forbidden.update(range(col + 1, col + span))
                elif col >= prefix:
                    preferred.add(col + span)

    return {"count": count, "header_rows": header_rows, "prefix": prefix, "widths": widths,
            "fixed_columns": fixed_columns, "forbidden": forbidden, "preferred": preferred,
            "body_merges": body_merges, "horizontal_margin": horizontal_margin, "matrix": matrix}


def plan_table_parts(table, body_width, font_size_pt, style_config=None, layout=None):
    """Repeat left labels/BASE, split data columns, and project header merges.

    Header banners may be repeated across fragments. Body merges are atomic:
    a value must never be silently divided or dropped to make a table fit.

    ``layout`` (see ``hwp_table_layout_options``) may give explicit ``cuts``
    (1-based source columns that start a new fragment), fixed
    ``column_widths_mm``, and ``body_merge_policy`` ``split_repeat`` which allows
    cutting a merged body cell; its text is then repeated in every fragment.
    """
    layout = layout or {}
    body_policy = layout.get("body_merge_policy", "block")
    if body_policy not in ("block", "split_repeat"):
        raise ValueError(f"알 수 없는 본문 병합 정책입니다: {body_policy}")
    if not math.isfinite(body_width) or body_width <= 0 or not math.isfinite(font_size_pt) or font_size_pt <= 0:
        raise ValueError("본문 폭과 표 글자 크기는 양수여야 합니다.")
    analysis = analyze_table_columns(table, font_size_pt, style_config, layout)
    matrix, count, header_rows, prefix = analysis["matrix"], analysis["count"], analysis["header_rows"], analysis["prefix"]
    widths, fixed_columns = analysis["widths"], analysis["fixed_columns"]
    forbidden, preferred, body_merges = analysis["forbidden"], analysis["preferred"], analysis["body_merges"]
    horizontal_margin = analysis["horizontal_margin"]
    # A small clearance avoids clipboard pixel rounding and outer border clipping.
    budget = int(body_width) - 500
    fixed = sum(widths[:prefix])
    selections = []
    explicit = layout.get("cuts")
    if explicit is not None:
        edges = sorted({int(c) - 1 for c in explicit})
        for edge in edges:
            if not prefix < edge < count:
                raise ValueError(f"분할 경계 {edge + 1}열은 반복 보기/BASE 열({prefix}열) 다음부터 {count}열 사이여야 합니다.")
            if edge in forbidden:
                raise ValueError(f"{edge + 1}열 앞에서 나누면 본문 병합 셀이 잘립니다. 경계를 옮기거나 본문 병합 분할을 허용하세요.")
        if not edges:
            selections.append(list(range(count)))
        else:
            bounds = [prefix, *edges, count]
            for start, end in zip(bounds, bounds[1:]):
                selections.append(list(range(prefix)) + list(range(start, end)))
        for index, columns in enumerate(selections, 1):
            if sum(widths[col] for col in columns) > budget:
                first = next((col for col in columns if col >= prefix), columns[0])
                raise ValueError(f"{index}번째 표({first + 1}~{columns[-1] + 1}열)가 본문 폭을 넘습니다. 분할 경계나 열 폭을 조정하세요.")
    elif sum(widths) <= budget:
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
                raise ValueError("값 열 또는 본문 병합 묶음이 본문 폭에 들어가지 않습니다. "
                                 "표 분할 미리보기에서 본문 병합 분할, 가로 용지 배치 또는 열 폭을 조정하세요.")
            cut = max(fitting)
            if cut < count:
                cut = max((edge for edge in fitting if edge in preferred), default=cut)
            selections.append(list(range(prefix)) + list(range(start, cut)))
            start = cut

    split_warnings = []
    edges = {columns[len(range(prefix))] for columns in selections[1:] if len(columns) > prefix}
    for row_index, col, span, cell in body_merges:
        if any(col < edge < col + span for edge in edges) or (len(selections) > 1 and col < prefix < col + span):
            text = str(cell.get("display_text") or "").strip() or "(빈 셀)"
            split_warnings.append(f"{row_index + 1}행 본문 병합 셀 '{text}'을 분할 경계에서 나누어 각 표에 반복 표시했습니다.")

    parts = []
    for index, columns in enumerate(selections, 1):
        part = _project_columns(table, columns)
        minimum = [widths[col] for col in columns]
        flexible = [i for i, col in enumerate(columns) if col not in fixed_columns]
        fixed_total = sum(minimum[i] for i in range(len(columns)) if i not in flexible)
        if fixed_total > budget:
            raise ValueError("고정한 열 폭의 합이 본문 폭을 넘습니다.")
        if flexible and len(flexible) < len(columns):
            # User-fixed columns keep their width; the rest share what remains.
            allocated = list(minimum)
            remaining = budget - fixed_total
            flexible_total = sum(minimum[i] for i in flexible)
            for i in flexible:
                allocated[i] = int(minimum[i] * remaining / flexible_total)
            allocated[flexible[-1]] += budget - sum(allocated)
        else:
            allocated = [int(value * budget / sum(minimum)) for value in minimum]
            allocated[-1] += budget - sum(allocated)
        for row_index, row in enumerate(part["matrix"]):
            for col_index, cell in enumerate(row):
                if not cell.get("covered_by") and horizontal_margin(cell, row_index) >= sum(allocated[col_index:col_index + cell["colspan"]]):
                    raise ValueError(f"{row_index + 1}행 {col_index + 1}열: 셀 여백을 빼면 글자가 들어갈 폭이 없습니다.")
        part.update(source_columns=[col + 1 for col in columns], repeated_columns=list(range(1, prefix + 1)),
                    column_widths_hwpunit=allocated, body_width_hwpunit=int(body_width),
                    part_index=index, part_count=len(selections), layout_prepared=True,
                    layout_warnings=split_warnings if index == 1 else [])
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
