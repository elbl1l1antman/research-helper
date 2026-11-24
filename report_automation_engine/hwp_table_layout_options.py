"""User-reviewed table layout overrides for HWP table output.

``table_layout.json`` keeps two kinds of settings:

* ``defaults``: page/merge policies applied to every table.
* ``tables``: per ``table_key`` overrides, usually saved from the split preview
  (explicit split boundaries and fixed column widths).

Schema (``schema_version`` ``"1.0"``)::

    {
      "schema_version": "1.0",
      "defaults": {
        "wide_table_policy": "split",
        "body_merge_policy": "block",
        "header_merge_policy": "base_only",
        "vertical_banner_policy": "merge"
      },
      "tables": {
        "Q1": {"cuts": [6, 11], "column_widths_mm": {"1": 28.0}, "wide_table_policy": "landscape_if_needed"}
      }
    }

``cuts`` are 1-based source column numbers where a new table fragment starts.
An empty list means "do not split"; a missing key means "split automatically".
"""

from __future__ import annotations

import copy
import json
import math
from pathlib import Path
from typing import Any, Dict

SCHEMA_VERSION = "1.0"
POLICY_VALUES = {
    # split: split columns to fit (current behaviour), block when impossible.
    # landscape_if_needed: rotate this table's section when portrait cannot fit it.
    # landscape_if_split: rotate when portrait would need a split or cannot fit it.
    "wide_table_policy": ("split", "landscape_if_needed", "landscape_if_split"),
    # block: never cut a merged body cell; split_repeat: cut it and repeat its text in each fragment.
    "body_merge_policy": ("block", "split_repeat"),
    # base_only: when a merge crosses the header boundary repeat only BASE (current behaviour);
    # split_at_header: cut non-value merges at the boundary so header rows repeat on every page.
    "header_merge_policy": ("base_only", "split_at_header"),
    # merge: keep vertical banners merged (Hancom shows the label once, on the first page);
    # repeat_per_row: unmerge and write the label in every row so each page shows it.
    "vertical_banner_policy": ("merge", "repeat_per_row"),
}
LAYOUT_DEFAULTS = {name: values[0] for name, values in POLICY_VALUES.items()}
TABLE_FIELDS = set(POLICY_VALUES) | {"cuts", "column_widths_mm", "col_count"}
MAX_COLUMN_WIDTH_MM = 400.0


def default_table_layout() -> Dict[str, Any]:
    return {"schema_version": SCHEMA_VERSION, "defaults": dict(LAYOUT_DEFAULTS), "tables": {}}


def _policies(value: Dict[str, Any], field: str, allowed_extra: set[str] = frozenset()) -> Dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError(f"{field}: 객체가 필요합니다.")
    unknown = set(value) - set(POLICY_VALUES) - set(allowed_extra)
    if unknown:
        raise ValueError(f"{field}: 알 수 없는 항목 {sorted(unknown)}")
    result = {}
    for name, allowed in POLICY_VALUES.items():
        if name in value:
            if value[name] not in allowed:
                raise ValueError(f"{field}.{name}: {', '.join(allowed)} 중 하나여야 합니다.")
            result[name] = value[name]
    return result


def normalize_table_layout(config: Dict[str, Any] | None) -> Dict[str, Any]:
    """Return a validated deep copy; raise ValueError naming the bad field."""
    if config is None:
        return default_table_layout()
    if not isinstance(config, dict):
        raise ValueError("table_layout: 객체가 필요합니다.")
    if config.get("schema_version") != SCHEMA_VERSION:
        raise ValueError(f"table_layout.schema_version: '{SCHEMA_VERSION}'이어야 합니다.")
    unknown = set(config) - {"schema_version", "defaults", "tables"}
    if unknown:
        raise ValueError(f"table_layout: 알 수 없는 항목 {sorted(unknown)}")
    defaults = {**LAYOUT_DEFAULTS, **_policies(config.get("defaults", {}), "defaults")}
    tables_in = config.get("tables", {})
    if not isinstance(tables_in, dict):
        raise ValueError("tables: table_key별 객체가 필요합니다.")
    tables = {}
    for key, raw in tables_in.items():
        field = f"tables.{key}"
        if not isinstance(key, str) or not key.strip():
            raise ValueError("tables: 빈 table_key는 사용할 수 없습니다.")
        entry = _policies(raw, field, {"cuts", "column_widths_mm", "col_count"})
        if "col_count" in raw:
            if type(raw["col_count"]) is not int or raw["col_count"] < 1:
                raise ValueError(f"{field}.col_count: 1 이상의 정수가 필요합니다.")
            entry["col_count"] = raw["col_count"]
        if "cuts" in raw:
            cuts = raw["cuts"]
            if not isinstance(cuts, list) or any(type(c) is not int or c < 2 for c in cuts):
                raise ValueError(f"{field}.cuts: 2 이상의 열 번호 목록이 필요합니다.")
            if len(set(cuts)) != len(cuts):
                raise ValueError(f"{field}.cuts: 같은 분할 경계가 중복되었습니다.")
            entry["cuts"] = sorted(cuts)
        if "column_widths_mm" in raw:
            widths = raw["column_widths_mm"]
            if not isinstance(widths, dict):
                raise ValueError(f"{field}.column_widths_mm: 열 번호별 mm 객체가 필요합니다.")
            normalized = {}
            for col, mm in widths.items():
                try:
                    number = int(col)
                except (TypeError, ValueError):
                    raise ValueError(f"{field}.column_widths_mm: 열 번호 '{col}'가 올바르지 않습니다.") from None
                if number < 1 or str(number) != str(col).strip():
                    raise ValueError(f"{field}.column_widths_mm: 열 번호 '{col}'가 올바르지 않습니다.")
                if isinstance(mm, bool) or not isinstance(mm, (int, float)) or not math.isfinite(mm) or not 2 <= mm <= MAX_COLUMN_WIDTH_MM:
                    raise ValueError(f"{field}.column_widths_mm.{col}: 2..{MAX_COLUMN_WIDTH_MM:g} mm 값이 필요합니다.")
                normalized[number] = float(mm)
            entry["column_widths_mm"] = normalized
        tables[key] = entry
    return {"schema_version": SCHEMA_VERSION, "defaults": defaults, "tables": tables}


def layout_for_table(config: Dict[str, Any] | None, table_key: str) -> Dict[str, Any]:
    """Merge defaults with one table's overrides. ``config`` must be normalized."""
    config = config or default_table_layout()
    result = dict(config.get("defaults") or LAYOUT_DEFAULTS)
    result.update(copy.deepcopy(config.get("tables", {}).get(str(table_key), {})))
    return result


def load_table_layout(path: str | Path | None) -> Dict[str, Any]:
    if not path:
        return default_table_layout()
    with Path(path).open("r", encoding="utf-8-sig") as stream:
        return normalize_table_layout(json.load(stream))


def dump_table_layout(config: Dict[str, Any]) -> Dict[str, Any]:
    """JSON-ready copy (column numbers become strings)."""
    normalized = normalize_table_layout(config)
    tables = {}
    for key, entry in normalized["tables"].items():
        entry = dict(entry)
        if "column_widths_mm" in entry:
            entry["column_widths_mm"] = {str(col): mm for col, mm in sorted(entry["column_widths_mm"].items())}
        tables[key] = entry
    return {**normalized, "tables": tables}
