"""One table-planning entry point shared by the HWP writer and the split preview.

The writer measures the real body width in Hancom; the preview estimates it
from the template or page setup. Both then call :func:`plan_hwp_table`, so the
preview shows exactly the fragments, widths, and warnings the writer will use.
"""

from __future__ import annotations

from typing import Any, Dict

try:
    from .hwp_style_config import TABLE_OUTPUT_DEFAULTS
    from .hwp_table_layout import apply_vertical_banner_policy, plan_table_parts, prepare_report_table
    from .hwp_table_layout_options import LAYOUT_DEFAULTS
    from .hwp_table_output import format_table_display
    from .report_table_matrix import resolve_table_metadata
except ImportError:
    from hwp_style_config import TABLE_OUTPUT_DEFAULTS
    from hwp_table_layout import apply_vertical_banner_policy, plan_table_parts, prepare_report_table
    from hwp_table_layout_options import LAYOUT_DEFAULTS
    from hwp_table_output import format_table_display
    from report_table_matrix import resolve_table_metadata


def table_font_size(style_config: Dict[str, Any] | None) -> float:
    if not style_config:
        return 9
    return max(style_config["paragraph_styles"][name]["font_size_pt"] for name in style_config["table_cell_styles"])


def choose_orientation(source, body_width, font_size, style_config, layout, landscape_width):
    """Return (orientation, width, parts) following ``wide_table_policy``."""
    policy = layout["wide_table_policy"]
    try:
        parts, error = plan_table_parts(source, body_width, font_size, style_config, layout), None
    except ValueError as exc:
        parts, error = None, exc
    rotate = landscape_width and (
        (policy == "landscape_if_needed" and parts is None)
        or (policy == "landscape_if_split" and (parts is None or len(parts) > 1))
    )
    if rotate:
        try:
            rotated = plan_table_parts(source, landscape_width, font_size, style_config, layout)
        except ValueError as exc:
            if parts is None:
                raise ValueError(f"세로 용지: {error} / 가로 용지: {exc}") from exc
        else:
            if parts is None or len(rotated) < len(parts):
                return "landscape", int(landscape_width), rotated
    if parts is None:
        raise error
    return "portrait", int(body_width), parts


def plan_hwp_table(
    table: Dict[str, Any],
    body_width: int,
    style_config: Dict[str, Any] | None = None,
    layout: Dict[str, Any] | None = None,
    *,
    landscape_width: int | None = None,
) -> Dict[str, Any]:
    """Plan render-ready fragments for one package table with a ``matrix``.

    ``body_width`` is the current section's width. ``landscape_width`` is the
    width after rotating that section; pass ``None`` to forbid rotation.
    Raises ValueError with a user-facing reason when no layout is possible.
    """
    layout = {**LAYOUT_DEFAULTS, **(layout or {})}
    key = str(table.get("table_key", ""))
    font_size = table_font_size(style_config)
    # Resolve metadata before splitting so every fragment carries the same unit.
    metadata = resolve_table_metadata(table)
    source = {**table, "base_label": metadata["base_label"], "unit": metadata["unit"]}
    source = format_table_display(source, {**TABLE_OUTPUT_DEFAULTS, **(style_config or {}).get("table_output", {})})
    source = apply_vertical_banner_policy(source, layout["vertical_banner_policy"])
    orientation, width, planned = choose_orientation(source, body_width, font_size, style_config, layout, landscape_width)
    parts = [prepare_report_table(part, layout["header_merge_policy"]) for part in planned]

    warnings = []
    if "확인 필요" in (metadata["base_label"], metadata["unit"]):
        warnings.append(f"{key}: BASE 조건 또는 단위를 확인해 주세요.")
    if source.get("vertical_banner_repeated"):
        warnings.append(f"{key}: 세로 배너 {source['vertical_banner_repeated']}개를 행마다 반복 표시했습니다.")
    for message in dict.fromkeys(m for part in planned for m in part.get("layout_warnings", [])):
        warnings.append(f"{key}: {message}")
    for message in dict.fromkeys(part["header_repeat_warning"] for part in parts if part.get("header_repeat_warning")):
        warnings.append(f"{key}: {message}")
    if orientation == "landscape":
        warnings.append(f"{key}: 세로 용지 본문 폭에 맞지 않아 이 표를 가로 용지 구역에 배치합니다.")
    if len(parts) > 1:
        warnings.append(f"{key}: 본문 폭에 맞춰 {len(parts)}개 표로 가로 분할했습니다. 보기/BASE 열은 반복됩니다.")
    return {"table_key": key, "orientation": orientation, "body_width_hwpunit": width,
            "parts": parts, "warnings": warnings, "layout": layout, "metadata": metadata}
