from __future__ import annotations

import copy
import json
import re
from pathlib import Path
from typing import Any, Dict


REQUIRED_STYLE_NAMES = ("보고서 본문1", "보고서 본문2", "표보기", "표배너", "표숫자")
TABLE_STYLE_NAMES = ("표보기", "표배너", "표숫자")
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
    if set(paragraph_styles) != set(REQUIRED_STYLE_NAMES):
        missing = [name for name in REQUIRED_STYLE_NAMES if name not in paragraph_styles]
        extra = [name for name in paragraph_styles if name not in REQUIRED_STYLE_NAMES]
        raise ValueError(f"paragraph_styles: fixed style names required; missing={missing}, extra={extra}")
    for name in REQUIRED_STYLE_NAMES:
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
    if set(cell_styles) != set(TABLE_STYLE_NAMES):
        raise ValueError(f"table_cell_styles: fixed style names required: {TABLE_STYLE_NAMES}")
    for name in TABLE_STYLE_NAMES:
        style = _object(cell_styles[name], f"table_cell_styles.{name}")
        prefix = f"table_cell_styles.{name}"
        style["fill_color"] = _color(style.get("fill_color"), f"{prefix}.fill_color")
        if style.get("vertical_alignment") not in {"top", "center", "bottom"}:
            raise ValueError(f"{prefix}.vertical_alignment: expected top, center, or bottom")

    border = _object(normalized.get("table_border"), "table_border")
    border["inner_width_mm"] = _number(border.get("inner_width_mm"), "table_border.inner_width_mm", 0.1, 5.0)
    border["outer_width_mm"] = _number(border.get("outer_width_mm"), "table_border.outer_width_mm", 0.1, 5.0)
    border["color"] = _color(border.get("color"), "table_border.color")
    return normalized


def load_hwp_style_config(path: Path | None) -> Dict[str, Any]:
    return validate_hwp_style_config(DEFAULT_HWP_STYLE_CONFIG if path is None else _read_json(Path(path)))


def style_name_for_cell_role(role: str) -> str:
    if role in {"banner_horizontal", "base", "header", "title"}:
        return "표배너"
    if role == "value":
        return "표숫자"
    return "표보기"
