"""Preview how report tables will be split before Hancom is started.

The launcher runs this module from the table list. It estimates the BODY
section width from the HWPX template (or the page setup chosen in the
launcher), then calls the same :func:`plan_hwp_table` the writer uses.
The JSON result lists every column's minimum and planned width, the split
fragments, which boundaries are allowed, and why a table is blocked.

Example::

    python -m report_automation_engine.hwp_table_split_preview \
        --package report_package.json --template report_template_basic.hwpx \
        --table-layout table_layout.json --output table_split_preview.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

try:
    from .hwp_page_setup import PAGE_FIELDS, normalize_page_setup
    from .hwp_style_config import load_hwp_style_config
    from .hwp_table_layout import analyze_table_columns, apply_vertical_banner_policy
    from .hwp_table_layout_options import layout_for_table, load_table_layout
    from .hwp_table_planning import plan_hwp_table, table_font_size
except ImportError:
    from hwp_page_setup import PAGE_FIELDS, normalize_page_setup
    from hwp_style_config import load_hwp_style_config
    from hwp_table_layout import analyze_table_columns, apply_vertical_banner_policy
    from hwp_table_layout_options import layout_for_table, load_table_layout
    from hwp_table_planning import plan_hwp_table, table_font_size

HWPUNIT_PER_MM = 7200 / 25.4
BODY_PLACEHOLDER = "{{BODY}}"
# Hancom's default A4 portrait page (used only when nothing better is known).
DEFAULT_PAGE = {"width": 59528, "height": 84186, "left": 8504, "right": 8504, "gutter": 0,
                "gutter_type": "LEFT_ONLY", "landscape": False}
_PAGE_PR = re.compile(r"<(?:\w+:)?pagePr\b([^>]*)>(.*?)</(?:\w+:)?pagePr>", re.S)
_MARGIN = re.compile(r"<(?:\w+:)?margin\b([^>]*)/?>")
_ATTR = re.compile(r'(\w+)\s*=\s*"([^"]*)"')


def to_mm(units: float) -> float:
    return round(units / HWPUNIT_PER_MM, 1)


def _body_widths(page: Dict[str, Any]) -> Dict[str, int]:
    gutter = 0 if page["gutter_type"] == "TOP_BOTTOM" else page["gutter"]
    portrait = page["width"] - page["left"] - page["right"] - gutter
    landscape = page["height"] - page["left"] - page["right"] - gutter
    current, rotated = (landscape, portrait) if page["landscape"] else (portrait, landscape)
    return {"current": int(current), "rotated": int(rotated), "landscape": bool(page["landscape"])}


def page_from_hwpx(path: str | Path) -> Dict[str, Any]:
    """Read the page definition that applies to ``{{BODY}}`` in an HWPX template."""
    with zipfile.ZipFile(path) as archive:
        sections = sorted(name for name in archive.namelist() if re.fullmatch(r"Contents/section\d+\.xml", name))
        chosen = None
        for name in sections:
            xml = archive.read(name).decode("utf-8", errors="replace")
            position = xml.find(BODY_PLACEHOLDER)
            if position < 0:
                continue
            # A section property applies from its paragraph onward: use the last one before BODY.
            matches = [m for m in _PAGE_PR.finditer(xml) if m.start() < position] or list(_PAGE_PR.finditer(xml))
            if matches:
                chosen = matches[-1]
            break
        if chosen is None:
            raise ValueError("HWPX 템플릿에서 {{BODY}}가 있는 구역의 편집용지 설정을 찾지 못했습니다.")
    attrs = dict(_ATTR.findall(chosen.group(1)))
    margin_tag = _MARGIN.search(chosen.group(2))
    margins = dict(_ATTR.findall(margin_tag.group(1))) if margin_tag else {}
    return {
        "width": int(attrs.get("width", DEFAULT_PAGE["width"])),
        "height": int(attrs.get("height", DEFAULT_PAGE["height"])),
        "left": int(margins.get("left", DEFAULT_PAGE["left"])),
        "right": int(margins.get("right", DEFAULT_PAGE["right"])),
        "gutter": int(margins.get("gutter", 0)),
        "gutter_type": attrs.get("gutterType", "LEFT_ONLY"),
        # OWPML: WIDELY is Hancom's normal (portrait) setting, NARROWLY is landscape.
        "landscape": attrs.get("landscape", "WIDELY").upper() == "NARROWLY",
    }


def page_from_setup(config: Dict[str, Any]) -> Dict[str, Any]:
    config = normalize_page_setup(config)
    units = {key: round(config[key] * HWPUNIT_PER_MM) for key in PAGE_FIELDS}
    return {"width": units["paper_width_mm"], "height": units["paper_height_mm"], "left": units["left_mm"],
            "right": units["right_mm"], "gutter": units["gutter_mm"], "gutter_type": "LEFT_ONLY",
            "landscape": config["orientation"] == "landscape"}


def resolve_page(template: str | Path | None, page_setup: Dict[str, Any] | None) -> Dict[str, Any]:
    if page_setup is not None:
        return {**page_from_setup(page_setup), "source": "page_setup"}
    if template and Path(template).suffix.lower() == ".hwpx":
        return {**page_from_hwpx(template), "source": "template"}
    return {**DEFAULT_PAGE, "source": "default_a4"}


def column_labels(matrix: List[List[Dict[str, Any]]], header_rows: int) -> List[str]:
    labels = []
    for col in range(len(matrix[0]) if matrix else 0):
        names = []
        for row in matrix[:header_rows]:
            item = row[col]
            if item.get("covered_by"):
                continue
            text = str(item.get("display_text") or "").strip()
            if text:
                names.append(text)
        if not names:
            body = [str(row[col].get("display_text") or "").strip() for row in matrix[header_rows:] if not row[col].get("covered_by")]
            names = [next((text for text in body if text), "")]
        labels.append(" / ".join(dict.fromkeys(names)))
    return labels


def preview_table(table: Dict[str, Any], page: Dict[str, Any], style_config: Dict[str, Any], layout_config: Dict[str, Any]) -> Dict[str, Any]:
    key = str(table.get("table_key", ""))
    layout = layout_for_table(layout_config, key)
    widths = _body_widths(page)
    result: Dict[str, Any] = {
        "table_key": key, "title": table.get("title") or "", "layout": {k: v for k, v in layout.items()},
        "body_width_mm": to_mm(widths["current"]), "rotated_body_width_mm": to_mm(widths["rotated"]),
        "status": "ok", "error": "", "warnings": [], "parts": [], "columns": [],
    }
    if not table.get("matrix"):
        result.update(status="skipped", error="셀 구조(matrix)가 없는 이전 형식 표는 미리보기를 지원하지 않습니다.")
        return result
    font_size = table_font_size(style_config)
    source = apply_vertical_banner_policy(table, layout["vertical_banner_policy"])
    analysis = analyze_table_columns(source, font_size, style_config, layout)
    labels = column_labels(analysis["matrix"], analysis["header_rows"])
    prefix = analysis["prefix"]
    result["repeated_columns"] = list(range(1, prefix + 1))
    result["allowed_cuts"] = [edge + 1 for edge in range(prefix + 1, analysis["count"]) if edge not in analysis["forbidden"]]
    result["blocked_cuts"] = [{"col": edge + 1, "reason": "본문 병합 셀이 이 경계를 가로지릅니다."}
                              for edge in sorted(analysis["forbidden"]) if edge > prefix]
    result["preferred_cuts"] = sorted(edge + 1 for edge in analysis["preferred"] if prefix < edge < analysis["count"])
    columns = [{
        "col": col + 1, "label": labels[col],
        "role": "repeated" if col < prefix else "value",
        "min_width_mm": to_mm(analysis["widths"][col]),
        "fixed": col in analysis["fixed_columns"],
        "planned_width_mm": None, "part": None,
    } for col in range(analysis["count"])]
    result["columns"] = columns
    try:
        plan = plan_hwp_table(table, widths["current"], style_config, layout,
                              landscape_width=widths["rotated"] if layout["wide_table_policy"] != "split" else None)
    except ValueError as exc:
        result.update(status="blocked", error=str(exc))
        return result
    result.update(orientation=plan["orientation"], planned_body_width_mm=to_mm(plan["body_width_hwpunit"]),
                  warnings=plan["warnings"])
    for part in plan["parts"]:
        sources = part.get("source_columns") or list(range(1, analysis["count"] + 1))
        part_widths = part.get("column_widths_hwpunit") or []
        guard = len(part_widths) - len(sources)  # prepare_report_table may add a narrow guard column
        result["parts"].append({
            "index": part.get("part_index", 1), "source_columns": sources,
            "widths_mm": [to_mm(w) for w in part_widths[:len(sources)]],
            "guard_column": guard > 0,
        })
        for source_col, width in zip(sources, part_widths):
            column = columns[source_col - 1]
            if source_col > prefix or column["planned_width_mm"] is None:
                column["planned_width_mm"] = to_mm(width)
            if source_col > prefix:
                column["part"] = part.get("part_index", 1)
    result["current_cuts"] = [part["source_columns"][prefix] for part in result["parts"][1:] if len(part["source_columns"]) > prefix]
    return result


def build_preview(package: Dict[str, Any], *, template=None, page_setup=None, style_config=None,
                  table_layout=None, table_keys: List[str] | None = None) -> Dict[str, Any]:
    page = resolve_page(template, page_setup)
    style_config = style_config if style_config is not None else load_hwp_style_config(None)
    layout_config = table_layout if table_layout is not None else load_table_layout(None)
    titles = {str(s.get("table_key", "")): str(s.get("title") or "") for s in package.get("sections", [])}
    wanted = set(table_keys or [])
    tables = []
    for table in package.get("tables", []):
        key = str(table.get("table_key", ""))
        if wanted and key not in wanted:
            continue
        item = preview_table({**table, "title": table.get("title") or titles.get(key, "")}, page, style_config, layout_config)
        tables.append(item)
    widths = _body_widths(page)
    notes = []
    if page["source"] == "default_a4":
        notes.append("템플릿 편집용지를 읽지 못해 A4 기본 여백으로 추정했습니다. 실제 생성 시 한글에서 다시 측정합니다.")
    return {
        "schema_version": "1.0", "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "page": {"source": page["source"], "landscape": page["landscape"],
                 "body_width_mm": to_mm(widths["current"]), "rotated_body_width_mm": to_mm(widths["rotated"])},
        "notes": notes,
        "summary": {"tables": len(tables), "split": sum(1 for t in tables if len(t["parts"]) > 1),
                    "landscape": sum(1 for t in tables if t.get("orientation") == "landscape"),
                    "blocked": sum(1 for t in tables if t["status"] == "blocked")},
        "tables": tables,
    }


def render_text(preview: Dict[str, Any]) -> str:
    lines = [f"본문 폭 {preview['page']['body_width_mm']}mm (회전 시 {preview['page']['rotated_body_width_mm']}mm, {preview['page']['source']})"]
    lines += preview["notes"]
    for table in preview["tables"]:
        head = f"[{table['table_key']}] {table['title']}".rstrip()
        if table["status"] != "ok":
            lines.append(f"{head}: {table['status']} - {table['error']}")
            continue
        lines.append(f"{head}: {len(table['parts'])}개 표, {table['orientation']}")
        for part in table["parts"]:
            cells = ", ".join(f"{c}열 {w}mm" for c, w in zip(part["source_columns"], part["widths_mm"]))
            lines.append(f"  {part['index']}: {cells}")
        lines += [f"  ! {w}" for w in table["warnings"]]
    return "\n".join(lines)


def _load_json(path):
    with Path(path).open("r", encoding="utf-8-sig") as stream:
        return json.load(stream)


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Preview HWP table split boundaries and column widths.")
    parser.add_argument("--package", required=True)
    parser.add_argument("--template", help="HWPX template whose {{BODY}} section sets the body width.")
    parser.add_argument("--page-setup", help="Page setup JSON chosen in the launcher (overrides the template).")
    parser.add_argument("--style-config")
    parser.add_argument("--table-layout")
    parser.add_argument("--table-key", action="append", help="Limit to one table; repeatable.")
    parser.add_argument("--output", help="Write JSON here (default: stdout).")
    parser.add_argument("--format", choices=("json", "text"), default="json")
    args = parser.parse_args(argv)
    try:
        preview = build_preview(
            _load_json(args.package), template=args.template,
            page_setup=_load_json(args.page_setup) if args.page_setup else None,
            style_config=load_hwp_style_config(Path(args.style_config) if args.style_config else None),
            table_layout=load_table_layout(args.table_layout), table_keys=args.table_key)
    except (OSError, ValueError, zipfile.BadZipFile, json.JSONDecodeError) as exc:
        print(f"표 분할 미리보기 실패: {exc}", file=sys.stderr)
        return 2
    text = render_text(preview) if args.format == "text" else json.dumps(preview, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(run_cli())
