from __future__ import annotations

import copy
import json
import re
from pathlib import Path
from typing import Any, Dict


REQUIRED_STYLE_NAMES = ("보고서 본문1", "보고서 본문2", "표보기", "표배너", "표숫자")
TABLE_STYLE_NAMES = ("표보기", "표배너", "표숫자")
OPTIONAL_ROLE_STYLES = {
    "header": "표머리행", "banner_horizontal": "표가로배너",
    "banner_vertical": "표세로배너", "base": "표BASE",
}
MARGIN_SIDES = ("top", "bottom", "left", "right")
LINE_GROUPS = ("outer", "inner_horizontal", "inner_vertical", "header_bottom")
# Hancom 13.0.0.711: map the saved HWPX meaning, not the similar API name.
HWP_LINE_NAMES = {"none": "None", "solid": "Solid", "dot": "Dash", "dash": "Dot", "double_slim": "DoubleSlim"}
_DEFAULT_PATH = Path(__file__).with_name("config") / "default_hwp_style_config.json"
_HEX_COLOR = re.compile(r"^[0-9A-Fa-f]{6}$")


def _read_json(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise ValueError(f"config: expected object in {path}")
    return value


DEFAULT_HWP_STYLE_CONFIG = _read_json(_DEFAULT_PATH)


def _object(value: Any, field: str) -> Dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError(f"{field}: expected object")
    return value


def _number(value: Any, field: str, minimum: float, maximum: float) -> float | int:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not minimum <= value <= maximum:
        raise ValueError(f"{field}: expected {minimum}..{maximum}")
    return value


def _text(value: Any, field: str, *, allow_empty: bool = False) -> str:
    if not isinstance(value, str) or (not allow_empty and not value.strip()):
        raise ValueError(f"{field}: expected {'string' if allow_empty else 'non-empty string'}")
    return value


def _color(value: Any, field: str) -> str:
    text = _text(value, field)
    if not _HEX_COLOR.fullmatch(text):
        raise ValueError(f"{field}: expected 6-digit hex color")
    return text.upper()


def validate_hwp_style_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """Return a normalized deep copy or raise ValueError naming the invalid field."""
    normalized = copy.deepcopy(_object(config, "config"))
    if normalized.get("schema_version") != "1.0":
        raise ValueError("schema_version: expected '1.0'")
    _text(normalized.get("preset_name"), "preset_name")

    paragraph_styles = _object(normalized.get("paragraph_styles"), "paragraph_styles")
    optional = set(paragraph_styles) & set(OPTIONAL_ROLE_STYLES.values())
    if set(paragraph_styles) != set(REQUIRED_STYLE_NAMES) | optional:
        missing = [name for name in REQUIRED_STYLE_NAMES if name not in paragraph_styles]
        extra = [name for name in paragraph_styles if name not in REQUIRED_STYLE_NAMES and name not in optional]
        raise ValueError(f"paragraph_styles: unknown or missing style; missing={missing}, extra={extra}")
    for name in paragraph_styles:
        style = _object(paragraph_styles[name], f"paragraph_styles.{name}")
        prefix = f"paragraph_styles.{name}"
        _text(style.get("font_family"), f"{prefix}.font_family")
        style["font_size_pt"] = _number(style.get("font_size_pt"), f"{prefix}.font_size_pt", 6, 30)
        if not isinstance(style.get("bold"), bool):
            raise ValueError(f"{prefix}.bold: expected boolean")
        _text(style.get("bullet"), f"{prefix}.bullet", allow_empty=True)
        style["line_spacing_percent"] = _number(
            style.get("line_spacing_percent"), f"{prefix}.line_spacing_percent", 80, 300
        )
        style["left_indent_mm"] = _number(style.get("left_indent_mm"), f"{prefix}.left_indent_mm", -30, 100)
        style["first_line_indent_mm"] = _number(
            style.get("first_line_indent_mm"), f"{prefix}.first_line_indent_mm", -30, 100
        )
        if style.get("alignment") not in {"left", "center", "right", "justify"}:
            raise ValueError(f"{prefix}.alignment: expected left, center, right, or justify")

    cell_styles = _object(normalized.get("table_cell_styles"), "table_cell_styles")
    if set(cell_styles) != set(TABLE_STYLE_NAMES) | optional:
        raise ValueError(f"table_cell_styles: matching paragraph and cell styles required: {sorted(set(TABLE_STYLE_NAMES) | optional)}")
    for name in cell_styles:
        style = _object(cell_styles[name], f"table_cell_styles.{name}")
        prefix = f"table_cell_styles.{name}"
        style["fill_color"] = _color(style.get("fill_color"), f"{prefix}.fill_color")
        if style.get("vertical_alignment") not in {"top", "center", "bottom"}:
            raise ValueError(f"{prefix}.vertical_alignment: expected top, center, or bottom")
        if "margins_mm" in style:
            margins = _object(style["margins_mm"], f"{prefix}.margins_mm")
            if set(margins) != set(MARGIN_SIDES):
                raise ValueError(f"{prefix}.margins_mm: expected top, bottom, left, right")
            for side in MARGIN_SIDES:
                margins[side] = _number(margins[side], f"{prefix}.margins_mm.{side}", 0, 20)

    border = _object(normalized.get("table_border"), "table_border")
    border["inner_width_mm"] = _number(border.get("inner_width_mm"), "table_border.inner_width_mm", 0.1, 5.0)
    border["outer_width_mm"] = _number(border.get("outer_width_mm"), "table_border.outer_width_mm", 0.1, 5.0)
    border["color"] = _color(border.get("color"), "table_border.color")
    border.setdefault("show_side_borders", False)
    if not isinstance(border["show_side_borders"], bool):
        raise ValueError("table_border.show_side_borders: expected boolean")
    if "edges" in border:
        for group, raw in _object(border["edges"], "table_border.edges").items():
            field = f"table_border.edges.{group}"
            if group not in LINE_GROUPS:
                raise ValueError(f"{field}: unknown edge group")
            line = _object(raw, field)
            if not isinstance(line.get("type"), str) or line["type"] not in HWP_LINE_NAMES:
                raise ValueError(f"{field}.type: unsupported line type")
            line["width_mm"] = _number(line.get("width_mm"), f"{field}.width_mm", 0.1, 5)
            line["color"] = _color(line.get("color"), f"{field}.color")
    return normalized


def load_hwp_style_config(path: Path | None) -> Dict[str, Any]:
    return validate_hwp_style_config(DEFAULT_HWP_STYLE_CONFIG if path is None else _read_json(Path(path)))


def style_name_for_cell_role(role: str, config: Dict[str, Any] | None = None) -> str:
    separate = OPTIONAL_ROLE_STYLES.get(role)
    if separate and separate in (config or {}).get("table_cell_styles", {}):
        return separate
    if role in {"banner_horizontal", "base", "header", "title"}:
        return "표배너"
    if role == "value":
        return "표숫자"
    return "표보기"


def margin_units(mm: float) -> int:
    return int(mm * 7200 / 25.4)
