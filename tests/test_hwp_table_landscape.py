from pathlib import Path
from types import SimpleNamespace
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import report_automation_engine.hwp_com_writer as writer
from report_automation_engine.hwp_com_writer import HwpWriterError, write_body
from report_automation_engine.hwp_table_layout_options import normalize_table_layout


class FakeHwp:
    """Enough of HSecDef/PageSetup and MultiColumn for body-width logic."""

    def __init__(self):
        self.page = SimpleNamespace(PaperWidth=59528, PaperHeight=84186, LeftMargin=7087, RightMargin=7087,
                                    GutterLen=0, GutterType=0, Landscape=0)
        self.items = {}
        hset = SimpleNamespace(SetItem=lambda name, value: self.items.__setitem__(name, value))
        self.HParameterSet = SimpleNamespace(HSecDef=SimpleNamespace(PageDef=self.page, HSet=hset))
        self.HAction = SimpleNamespace(GetDefault=lambda *_: True, Execute=self.execute)
        self.executed = []

    def execute(self, name, _hset):
        self.executed.append((name, self.page.Landscape, self.items.get("ApplyTo")))
        return True

    def CreateAction(self, _name):
        columns = SimpleNamespace(Item=lambda _key: 1)
        return SimpleNamespace(CreateSet=lambda: columns, GetDefault=lambda _set: True)


def cell(row, col, text, role):
    return {"row": row, "col": col, "rowspan": 1, "colspan": 1, "covered_by": "",
            "source_cell": f"R{row}C{col}", "display_text": text, "role": role}


def merged_wide_table(key, cols=16):
    header = [cell(1, 1, "구분", "header")] + [cell(1, c, f"보기{c}", "banner_horizontal") for c in range(2, cols + 1)]
    body = [cell(2, 1, "전체", "stub")] + [cell(2, c, "" if c > 2 else "공통", "value") for c in range(2, cols + 1)]
    body[1]["colspan"] = cols - 1
    for item in body[2:]:
        item["covered_by"] = "R2C2"
    second = [cell(3, 1, "남", "stub")] + [cell(3, c, f"{c}.0", "value") for c in range(2, cols + 1)]
    return {"table_key": key, "matrix": [header, body, second], "row_count": 3, "col_count": cols,
            "merged_ranges": [{"row": 2, "col": 2, "rowspan": 1, "colspan": cols - 1}]}


def run(monkeypatch, layout, keys=("W",)):
    events, widths = [], []
    hwp = FakeHwp()
    monkeypatch.setattr(writer, "insert_text", lambda *_: None)
    monkeypatch.setattr(writer, "insert_narrative_blocks", lambda *_: None)
    monkeypatch.setattr(writer, "run_action", lambda _h, action, *_: events.append(action) or True)

    def write_part(_hwp, part, *_args):
        events.append("table")
        widths.append((part["body_width_hwpunit"], hwp.page.Landscape))

    monkeypatch.setattr(writer, "write_table_part", write_part)
    package = {"sections": [{"table_key": key} for key in keys],
               "tables": [merged_wide_table(key) for key in keys]}
    report = {"sections_written": 0, "warnings": [], "charts_deferred": 0}
    write_body(hwp, package, report, table_layout=normalize_table_layout(layout) if layout else None)
    return hwp, events, widths, report


def test_wide_merged_table_is_blocked_by_default(monkeypatch):
    with pytest.raises(HwpWriterError, match="본문"):
        run(monkeypatch, None)


def test_landscape_if_needed_rotates_only_that_section(monkeypatch):
    layout = {"schema_version": "1.0", "defaults": {"wide_table_policy": "landscape_if_needed"}}
    hwp, events, widths, report = run(monkeypatch, layout, keys=("W", "N"))
    # Title and narrative stay on the portrait page; only the table moves to a rotated section.
    assert events == ["BreakPara", "BreakPara", "BreakSection", "table",
                      "BreakSection", "BreakPara", "BreakPara", "BreakSection", "table"]
    assert widths == [(84186 - 14174, 1), (84186 - 14174, 1)]
    # rotate W, restore portrait for N's text (no extra BreakPage), rotate N.
    assert [entry[1] for entry in hwp.executed] == [1, 0, 1]
    assert all(entry[2] == 2 for entry in hwp.executed)
    assert report["table_layouts"][0]["orientation"] == "landscape"
    assert any("가로 용지" in warning for warning in report["warnings"])


def test_rotated_state_survives_to_the_next_section_through_the_report(monkeypatch):
    layout = {"schema_version": "1.0", "defaults": {"wide_table_policy": "landscape_if_needed"}}
    _hwp, _events, _widths, report = run(monkeypatch, layout)
    # The last rotated section stays open; a recovery resume restores it at the next section.
    assert report["rotated_section"] == {"table_key": "W", "original_landscape": False}


def test_table_override_can_split_the_merge_instead_of_rotating(monkeypatch):
    layout = {"schema_version": "1.0", "tables": {"W": {"body_merge_policy": "split_repeat"}}}
    hwp, events, widths, report = run(monkeypatch, layout)
    assert "BreakSection" not in events
    assert events.count("table") >= 2
    assert all(landscape == 0 for _width, landscape in widths)
    assert any("반복 표시" in warning for warning in report["warnings"])
