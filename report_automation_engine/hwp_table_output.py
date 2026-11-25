"""Display-only options applied before HWP table width and height are planned."""

import re
from decimal import Decimal, ROUND_HALF_UP

try:
    from .report_table_evidence import aggregate_measure, table_head_and_total
except ImportError:
    from report_table_evidence import aggregate_measure, table_head_and_total


_COUNT = re.compile(r"(?:wt)?(?:사례수|표본수|응답수|빈도수|빈도|n)(?:\(명\))?", re.I)
_ABSOLUTE_UNITS = r"명|인|개|건|회|원|천원|만원|백만원|천만원|억원|십억원"
_NUMBER = re.compile(r"([+-]?(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\.[0-9]+)?)")


def is_count_header(labels):
    """True for a sample-count column (사례수, 표본수, n …) that is not a percentage."""
    labels = [re.sub(r"\s+", "", label) for label in labels if label]
    return bool(labels) and not any("%" in label or "비율" in label for label in labels) and bool(_COUNT.fullmatch(labels[-1]))


def group_digits(text):
    """1234567.25 -> 1,234,567.25 (decimals kept as written)."""
    whole, dot, fraction = text.partition(".")
    sign = whole[:1] if whole[:1] in "+-" else ""
    return f"{sign}{int(whole.lstrip('+-').replace(',', '') or 0):,}{dot}{fraction}"


def format_table_display(table, options):
    decimals = options.get("value_decimals", -1)
    group_cases = options.get("sample_grouping", True)
    group_values = options.get("value_grouping", True)
    if decimals < 0 and not group_cases and not group_values:
        return table
    headers, _ = table_head_and_total(table)
    cols = table["col_count"]
    measures = {}
    for col in range(1, cols + 1):
        labels = [re.sub(r"\s+", "", c.get("display_text", "")) for row in headers for c in row
                  if c.get("role") in {"header", "banner_horizontal"} and c["col"] <= col < c["col"] + c.get("colspan", 1)]
        if not labels:
            continue
        if is_count_header(labels):
            measures[col] = "count"
        elif not any("%" in label or "비율" in label or "점" in label for label in labels):
            leaf = labels[-1]
            if re.fullmatch(_ABSOLUTE_UNITS, leaf) or re.search(rf"\(({_ABSOLUTE_UNITS})\)$", leaf) or (
                    aggregate_measure(leaf) and re.fullmatch(_ABSOLUTE_UNITS, aggregate_measure(leaf)[1])):
                measures[col] = "value"
    matrix = []
    for row in table["matrix"]:
        converted = []
        for cell in row:
            mode = measures.get(cell["col"])
            display = cell.get("display_text") or ""
            wrapped = display.startswith("(") and display.endswith(")")
            match = _NUMBER.fullmatch(display[1:-1] if wrapped else display)
            if cell.get("role") == "value" and not cell.get("covered_by") and match and "%" not in (cell.get("number_format") or ""):
                number = Decimal(match[1].replace(",", ""))
                if mode == "count" and group_cases and number == number.to_integral_value():
                    text = format(number, ",.0f")
                elif mode == "value" and decimals >= 0:
                    text = format(number.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP),
                                  f"{',' if group_values else ''}.{decimals}f")
                elif mode == "value" and group_values:
                    text = group_digits(match[1])
                else:
                    text = None
                if text is not None:
                    cell = {**cell, "display_text": f"({text})" if wrapped else text}
            converted.append(cell)
        matrix.append(converted)
    return {**table, "matrix": matrix, "cells": [c for row in matrix for c in row]}
