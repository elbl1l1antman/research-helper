"""Opt-in native Hancom calibration. Run with no other Hancom documents open."""

import argparse
import hashlib
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from report_automation_engine.hwp_com_writer import (
    close_hwp, create_hwp_object, find_hwp_table_control, insert_hwp_table,
    open_document, same_hwp_control, save_as_hwpx, set_parameter_item,
    set_visible, write_json,
)
from report_automation_engine.hwp_template_probe import extract_border_fills, local_name


ROWS = [["\uad6c\ubd84", "BASE", "\uacb0\uacfc"], ["A", "(3,232)", "3.3"], ["B", "0", ""]]
FILL = "DDEBF7"


def action(hwp, name):
    if not hwp.HAction.Run(name):
        raise RuntimeError(f"{name} returned False")


def change(hwp, name, params, update):
    hwp.HAction.GetDefault(name, params.HSet)
    update(params)
    if not hwp.HAction.Execute(name, params.HSet):
        raise RuntimeError(f"{name} returned False")


def select_cells(hwp, scope):
    control = find_hwp_table_control(hwp)
    if not control or not hwp.SetPosBySet(control.GetAnchorPos(0)) or not hwp.FindCtrl():
        raise RuntimeError("Cannot enter the synthetic table")
    action(hwp, "ShapeObjTableSelCell")
    action(hwp, "Cancel")
    if scope == "center":
        for _ in range(4):
            action(hwp, "TableRightCell")
    action(hwp, "TableCellBlock")
    if scope == "all":
        action(hwp, "TableCellBlockExtend")
        action(hwp, "TableCellBlockExtend")
    elif scope == "header":
        action(hwp, "TableCellBlockExtend")
        action(hwp, "TableRightCell")
        action(hwp, "TableRightCell")
    if not same_hwp_control(control, hwp.ParentCtrl):
        raise RuntimeError("Selection left the synthetic table")


def border(hwp, axis, kind, width, rgb):
    fields = (
        (f"Type{axis}", f"Width{axis}", f"Color{axis}") if axis in {"Horz", "Vert"}
        else tuple(f"Border{field}{axis}" for field in ("Type", "Width", "Color"))
    )
    values = (kind, hwp.HwpLineWidth(width), hwp.RGBColor(*rgb))

    def update(p):
        set_parameter_item(p, "ApplyTo", 0)
        for field, value in zip(fields, values):
            set_parameter_item(p, field, value)

    change(hwp, "CellBorderFill", hwp.HParameterSet.HCellBorderFill, update)


def fill(hwp):
    def update(p):
        for key, value in {
            "Type": 1, "WinBrushFaceColor": hwp.RGBColor(221, 235, 247),
            "WinBrushHatchColor": 0, "WinBrushFaceStyle": -1, "WindowsBrush": 1,
        }.items():
            set_parameter_item(p.FillAttr, key, value)

    change(hwp, "CellFill", hwp.HParameterSet.HCellBorderFill, update)


def margins(hwp, values):
    def update(p):
        cell = p.ShapeTableCell
        set_parameter_item(cell, "HasMargin", 1)
        for side, mm in zip(("Left", "Right", "Top", "Bottom"), values):
            set_parameter_item(cell, f"Margin{side}", hwp.MiliToHwpUnit(mm))

    change(hwp, "TablePropertyDialog", hwp.HParameterSet.HShapeObject, update)


def fill_then_header_border(hwp, solid):
    fill(hwp)
    action(hwp, "Cancel")
    select_cells(hwp, "header")
    border(hwp, "Bottom", solid, "0.4mm", (31, 78, 121))


def read_table(path):
    with ZipFile(path) as archive:
        fills = extract_border_fills(ET.fromstring(archive.read("Contents/header.xml")))
        tables = [node for node in ET.fromstring(archive.read("Contents/section0.xml")).iter()
                  if local_name(node.tag) == "tbl"]
        if "Preview/PrvImage.png" in archive.namelist():
            path.with_suffix(".png").write_bytes(archive.read("Preview/PrvImage.png"))
    assert len(tables) == 1
    table = tables[0]
    assert (table.get("rowCnt"), table.get("colCnt")) == ("3", "3")
    result = {}
    for node in table.iter():
        if local_name(node.tag) != "tc":
            continue
        parts = {local_name(child.tag): child for child in node}
        row, col = (int(parts["cellAddr"].get(key)) for key in ("rowAddr", "colAddr"))
        assert (parts["cellSpan"].get("rowSpan"), parts["cellSpan"].get("colSpan")) == ("1", "1")
        key = f"{row + 1},{col + 1}"
        assert key not in result
        assert "".join(t.text or "" for t in node.iter() if local_name(t.tag) == "t") == ROWS[row][col]
        style = fills[node.get("borderFillIDRef")]
        result[key] = {"fill": style["fill_color"], "borders": style["borders"],
                       "has_margin": node.get("hasMargin"), "margin": parts["cellMargin"].attrib,
                       "size": parts["cellSz"].attrib}
    assert len(result) == 9
    return result


def main(root):
    root.mkdir(parents=True, exist_ok=False)
    report = {"status": "started", "warnings": [], "errors": [], "com": {}, "cases": {}, "line_types": {}}
    log = root / "calibration.json"
    hwp = None
    try:
        hwp = create_hwp_object(report, log, "dispatch_ex")
        set_visible(hwp, False, report, log)
        report["version"] = hwp.Version
        for label in ("None", "Solid", "Dot", "Dash", "Double", "DoubleSlim", "DoubleLine"):
            try:
                report["line_types"][label] = hwp.HwpLineType(label)
            except Exception:
                report["line_types"][label] = "rejected"
        solid = report["line_types"]["Solid"]
        assert isinstance(solid, int) and solid > 0
        blue = (31, 78, 121)
        cases = [
            ("outer_top", "all", lambda: border(hwp, "Top", solid, "0.4mm", blue)),
            ("inner_horizontal", "all", lambda: border(hwp, "Horz", solid, "0.4mm", blue)),
            ("inner_vertical_off", "all", lambda: border(hwp, "Vert", 0, "0.12mm", blue)),
            ("header_bottom", "header", lambda: border(hwp, "Bottom", solid, "0.4mm", blue)),
            ("center_fill", "center", lambda: fill(hwp)),
            ("center_margins", "center", lambda: margins(hwp, (1.0, 1.0, 0.5, 0.5))),
            ("center_margins_asymmetric", "center", lambda: margins(hwp, (1.0, 2.0, 0.8, 1.2))),
            ("fill_then_header_bottom", "center", lambda: fill_then_header_border(hwp, solid)),
        ]
        for label in ("Dot", "Dash", "DoubleSlim", "DoubleLine", "Double"):
            code = report["line_types"][label]
            if isinstance(code, int) and code > 0 and code != solid:
                cases.append((f"line_{label.lower()}", "header", lambda code=code: border(hwp, "Top", code, "0.4mm", blue)))
        action(hwp, "FileNew")
        assert insert_hwp_table(hwp, ROWS, {"warnings": [], "tables_written": 0})
        baseline = root / "baseline.hwpx"
        save_as_hwpx(hwp, baseline, report, log)
        baseline_hash = hashlib.sha256(baseline.read_bytes()).hexdigest()
        before = read_table(baseline)
        report["baseline"] = before
        for name, scope, mutate in cases:
            report["case"] = name
            write_json(log, report)
            try:
                hwp.Clear(1)
                open_document(hwp, baseline, report, log)
                select_cells(hwp, scope)
                mutate()
                action(hwp, "Cancel")
                output = root / f"{name}.hwpx"
                save_as_hwpx(hwp, output, report, log)
                after = read_table(output)
                hwp.Clear(1)
                open_document(hwp, output, report, log)
                report["cases"][name] = {"status": "saved_and_reopened", "scope": scope,
                                         "changed": {key: val for key, val in after.items() if val != before[key]}}
            except Exception as exc:
                report["cases"][name] = {"status": "failed", "error": str(exc)}
            write_json(log, report)
        assert hashlib.sha256(baseline.read_bytes()).hexdigest() == baseline_hash
        report["baseline_sha256"] = baseline_hash
        report["status"] = "measured" if all(c["status"] == "saved_and_reopened" for c in report["cases"].values()) else "partial"
    except Exception as exc:
        report["status"] = "failed"
        report["errors"].append(str(exc))
        raise
    finally:
        if hwp is not None:
            close_hwp(hwp, report)
        write_json(log, report)
    print(f"{report['status']}: {log}")
    return report["status"] == "measured"


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    raise SystemExit(0 if main(parser.parse_args().output_dir.resolve()) else 1)
