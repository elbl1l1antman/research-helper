import copy
import subprocess
import tempfile
import unittest
from pathlib import Path
import sys
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from report_automation_engine.hwp_com_writer import (
    HwpWriterError,
    apply_cell_appearance,
    apply_named_style,
    finalize_hwp_resources,
    insert_clipboard_table,
    insert_hwp_table,
    insert_narrative_blocks,
    leave_hwp_table,
    write_body,
    write_hwp_document,
)
from report_automation_engine.hwp_style_config import DEFAULT_HWP_STYLE_CONFIG


STYLE_INDEXES = {
    "보고서 본문1": 10,
    "보고서 본문2": 11,
    "표보기": 12,
    "표배너": 13,
    "표숫자": 14,
}


class FakeAction:
    def __init__(self, owner):
        self.owner = owner

    def GetDefault(self, action, _set):
        self.owner.events.append(("default", action))

    def Execute(self, action, params):
        if action == "Style":
            self.owner.events.append(("style", params.Apply))
        elif action == "InsertText":
            self.owner.events.append(("text", params.Text))
            self.owner.contents[self.owner.list_id] = params.Text
        elif action == "CellBorderFill":
            border = params
            self.owner.events.append(("border_type", border.BorderTypeLeft, border.BorderTypeRight, border.BorderTypeTop, border.BorderTypeBottom))
            self.owner.events.append(
                (
                    "border",
                    border.BorderWidthLeft,
                    border.BorderWidthRight,
                    border.BorderWidthTop,
                    border.BorderWidthBottom,
                )
            )
        elif action == "CellFill":
            self.owner.events.append(("fill", params.FillAttr.WinBrushFaceColor))
        elif action == "TableCreate":
            self.owner.ParentCtrl = FakeTableControl()
        return True

    def Run(self, action):
        self.owner.events.append(("run", action))
        if action == "MoveListBegin":
            self.owner.position = (self.owner.list_id, 0, 0)
        elif action == "MoveSelListEnd":
            self.owner.position = (self.owner.list_id, 0, 1)
        elif action == "TableRightCell":
            self.owner.list_id += 1
            self.owner.position = (self.owner.list_id, 0, 0)
        return True


class FakeHwp:
    def __init__(self):
        self.events = []
        self.list_id = 2
        self.contents = {}
        self.position = (2, 0, 0)
        self.ParentCtrl = None
        style = SimpleNamespace(Apply=None)
        insert_text = SimpleNamespace(Text="")
        table_creation = SimpleNamespace(Rows=0, Cols=0)
        cell_border_fill = SimpleNamespace(
            ApplyTo=0,
            FillAttr=SimpleNamespace(
                type=0,
                WinBrushFaceColor=0,
                WinBrushHatchColor=0,
                WinBrushFaceStyle=0,
                WindowsBrush=0,
            ),
            BorderTypeLeft=0,
            BorderTypeRight=0,
            BorderTypeTop=0,
            BorderTypeBottom=0,
            BorderWidthLeft=0,
            BorderWidthRight=0,
            BorderWidthTop=0,
            BorderWidthBottom=0,
            BorderColorLeft=0,
            BorderColorRight=0,
            BorderColorTop=0,
            BorderColorBottom=0,
        )
        for params in (style, insert_text, table_creation, cell_border_fill):
            params.HSet = params
        self.HParameterSet = SimpleNamespace(
            HStyle=style,
            HInsertText=insert_text,
            HTableCreation=table_creation,
            HCellBorderFill=cell_border_fill,
        )
        self.HAction = FakeAction(self)

    @staticmethod
    def RGBColor(red, green, blue):
        return red | (green << 8) | (blue << 16)

    @staticmethod
    def HwpLineType(name):
        assert name == "Solid"
        return 1

    def SetPosBySet(self, position):
        self.events.append(("anchor", position))
        self.position = (0, 7, 3)
        return True

    def SetPos(self, list_id, paragraph, offset):
        self.events.append(("set_pos", list_id, paragraph, offset))
        self.position = (list_id, paragraph, offset)
        self.ParentCtrl = None
        return True

    def FindCtrl(self):
        self.events.append(("select_table",))
        return True

    def GetPos(self):
        return self.position


class FakeTableControl:
    CtrlID = "tbl"
    Properties = SimpleNamespace(Rows=1, Cols=3)

    @staticmethod
    def GetCtrlInstID():
        return 1

    @staticmethod
    def GetAnchorPos(_mode):
        return "table_anchor"


class FakeClipboardRange:
    def Copy(self):
        return None


class FakeWorkbook:
    class Sheet:
        @staticmethod
        def Range(_address):
            return FakeClipboardRange()

    @staticmethod
    def Worksheets(_name):
        return FakeWorkbook.Sheet()


def insert_fake_clipboard_table(table):
    hwp = FakeHwp()
    hwp.LastCtrl = None
    hwp.EngineProperties = SimpleNamespace(SetItem=lambda *_args: None)
    original_run = hwp.HAction.Run

    def run(action):
        result = original_run(action)
        if action == "Paste":
            hwp.LastCtrl = FakeTableControl()
        elif action == "ShapeObjTableSelCell":
            hwp.ParentCtrl = hwp.LastCtrl
            hwp.position = (hwp.list_id, 0, 0)
        elif action == "MoveDocEnd":
            hwp.ParentCtrl = None
        return result

    hwp.HAction.Run = run
    report = {
        "warnings": [],
        "style_application_counts": {},
        "table_results": [],
        "tables_written": 0,
    }
    with patch("report_automation_engine.hwp_com_writer.time.sleep"):
        inserted = insert_clipboard_table(
            hwp,
            FakeWorkbook(),
            table,
            report,
            None,
            STYLE_INDEXES,
            DEFAULT_HWP_STYLE_CONFIG,
        )
    return hwp, report, inserted


class HwpComWriterStyleTests(unittest.TestCase):
    def test_direct_script_help_does_not_require_package_import(self):
        script = Path(__file__).resolve().parents[1] / "report_automation_engine" / "hwp_com_writer.py"
        with tempfile.TemporaryDirectory() as directory:
            result = subprocess.run(
                [sys.executable, str(script), "--help"], cwd=directory,
                capture_output=True, text=True, timeout=30,
            )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("--style-config", result.stdout)

    def test_write_body_removes_only_configured_nonempty_body2_prefix(self):
        cases = (
            ("●", ["● first", "second", "third", "- fourth", "-3232", "●fifth"]),
            ("", ["● first", "● second", "●\tthird", "- fourth", "-3232", "●fifth"]),
            ("--", ["● first", "second", "third", "- fourth", "-3232", "--fifth"]),
        )
        for bullet, expected in cases:
            with self.subTest(bullet=bullet):
                config = copy.deepcopy(DEFAULT_HWP_STYLE_CONFIG)
                config["paragraph_styles"]["보고서 본문2"]["bullet"] = bullet
                prefix = bullet or "●"
                texts = ["● first", prefix + " second", prefix + "\tthird", "- fourth", "-3232", prefix + "fifth"]
                section = {"table_key": "T001", "title": "Title", "narrative_blocks": [
                    {"style": "보고서 본문1" if index == 0 else "보고서 본문2", "text": text}
                    for index, text in enumerate(texts)
                ]}
                hwp = FakeHwp()
                report = {"warnings": [], "style_application_counts": {}, "sections_written": 0}

                write_body(hwp, {"sections": [section]}, report, style_indexes=STYLE_INDEXES, style_config=config)

                self.assertEqual([event[1] for event in hwp.events if event[0] == "text"], ["Title", *expected, "source: T001"])
                self.assertEqual([block["text"] for block in section["narrative_blocks"]], texts)

    def test_apply_named_style_uses_style_action_and_counts(self):
        hwp = FakeHwp()
        report = {"style_application_counts": {}}

        apply_named_style(hwp, 11, report, "보고서 본문2")

        self.assertEqual(hwp.events, [("default", "Style"), ("style", 11)])
        self.assertEqual(report["style_application_counts"], {"보고서 본문2": 1})

    def test_narrative_applies_body_styles_and_removes_literal_bullet(self):
        hwp = FakeHwp()
        report = {"style_application_counts": {}}
        section = {
            "narrative_blocks": [
                {"style": "보고서 본문1", "text": "첫 문장"},
                {"style": "보고서 본문2", "text": "- 다음 문장"},
                {"style": "보고서 본문2", "text": "-3232"},
                {"style": "보고서 본문2", "text": "-\tlast sentence"},
            ]
        }

        insert_narrative_blocks(hwp, section, report, STYLE_INDEXES)

        self.assertEqual(
            [event for event in hwp.events if event[0] in {"style", "text"}],
            [
                ("style", 10),
                ("text", "첫 문장"),
                ("style", 11),
                ("text", "다음 문장"),
                ("style", 11),
                ("text", "-3232"),
                ("style", 11),
                ("text", "last sentence"),
            ],
        )

    def test_table_cells_use_role_styles_and_companion_appearance(self):
        hwp = FakeHwp()
        report = {"warnings": [], "style_application_counts": {}, "table_style_applied": {}, "tables_written": 0}
        table = {
            "matrix": [
                [
                    {"display_text": "배너", "role": "banner_horizontal"},
                    {"display_text": "42.0", "role": "value"},
                    {"display_text": "항목", "role": "stub"},
                ]
            ],
            "merged_ranges": [],
        }

        self.assertTrue(
            insert_hwp_table(
                hwp,
                [["배너", "42.0", "항목"]],
                report,
                table=table,
                style_indexes=STYLE_INDEXES,
                style_config=DEFAULT_HWP_STYLE_CONFIG,
            )
        )

        self.assertEqual([event[1] for event in hwp.events if event[0] == "style"], [13, 14, 12])
        fills = [event for event in hwp.events if event[0] == "fill"]
        self.assertEqual(len(fills), 3)
        self.assertNotEqual(fills[0][1], fills[1][1])
        self.assertIn(("run", "TableCellAlignCenterCenter"), hwp.events)
        self.assertIn(("run", "TableCellAlignLeftCenter"), hwp.events)

    def test_merged_cell_uses_solid_borders_and_span_outer_edges(self):
        hwp = FakeHwp()
        apply_cell_appearance(hwp, DEFAULT_HWP_STYLE_CONFIG, "표배너", 0, 0, 2, 3, {"warnings": []}, rowspan=2, colspan=3)
        self.assertIn(("border_type", 1, 1, 1, 1), hwp.events)
        borders = [event for event in hwp.events if event[0] == "border"]
        self.assertEqual(borders, [("border", 6, 6, 6, 6)])  # 0.4 mm outer border enum.
        table = {
            "table_key": "T001", "cell_contract": True,
            "source_sheet": "Sheet1", "source_range": "A1:C1", "row_count": 1, "col_count": 3,
            "matrix": [[
                {"display_text": "병합", "role": "banner_horizontal", "colspan": 3},
                {"covered_by": "A1"}, {"covered_by": "A1"},
            ]],
        }
        pasted, _report, inserted = insert_fake_clipboard_table(table)
        self.assertTrue(inserted)
        self.assertIn(("border", 6, 6, 6, 6), pasted.events)

    def test_clipboard_table_replaces_each_cell_after_applying_style(self):
        table = {
            "table_key": "T001",
            "cell_contract": True,
            "source_sheet": "Sheet1",
            "source_range": "A1:C1",
            "row_count": 1,
            "col_count": 3,
            "matrix": [[
                {"display_text": "배너", "role": "banner_horizontal"},
                {"display_text": "42.0", "role": "value"},
                {"display_text": "항목", "role": "stub"},
            ]],
        }

        hwp, _report, inserted = insert_fake_clipboard_table(table)
        self.assertTrue(inserted)

        self.assertEqual(
            [event for event in hwp.events if event[0] in {"style", "text"}],
            [
                ("style", 13), ("text", "배너"),
                ("style", 14), ("text", "42.0"),
                ("style", 12), ("text", "항목"),
            ],
        )
        self.assertEqual(sum(event == ("run", "Delete") for event in hwp.events), 3)
        self.assertLess(hwp.events.index(("select_table",)), hwp.events.index(("run", "Delete")))
        self.assertIn(("run", "ShapeObjTableSelCell"), hwp.events)
        self.assertIn(("set_pos", 0, 7, 4), hwp.events)
        self.assertNotIn(("run", "MoveDocEnd"), hwp.events)
        self.assertIsNone(hwp.ParentCtrl)

    def test_leaves_current_table_after_anchor_not_document_end(self):
        hwp = FakeHwp()
        hwp.ParentCtrl = FakeTableControl()
        report = {"warnings": []}
        leave_hwp_table(hwp, hwp.ParentCtrl, report)
        self.assertEqual(hwp.GetPos(), (0, 7, 4))
        self.assertNotIn(("run", "MoveDocEnd"), hwp.events)
        with patch.object(FakeHwp, "SetPos", return_value=False):
            with self.assertRaisesRegex(HwpWriterError, "표 뒤"):
                leave_hwp_table(hwp, FakeTableControl(), report)

    def test_clipboard_cell_entry_failure_undoes_paste_before_fallback(self):
        table = {
            "table_key": "T001", "cell_contract": True,
            "source_sheet": "Sheet1", "source_range": "A1:C1",
            "row_count": 1, "col_count": 3,
            "matrix": [[{"display_text": "항목", "role": "stub"}] * 3],
        }
        with patch.object(FakeHwp, "FindCtrl", return_value=False):
            hwp, report, inserted = insert_fake_clipboard_table(table)
        self.assertFalse(inserted)
        self.assertIn(("run", "Undo"), hwp.events)
        self.assertFalse(any(event[0] == "text" for event in hwp.events))
        self.assertEqual(report["tables_written"], 0)

    def test_clipboard_empty_cell_does_not_delete_cell_boundary(self):
        table = {
            "table_key": "T001", "cell_contract": True,
            "source_sheet": "Sheet1", "source_range": "A1:C1",
            "row_count": 1, "col_count": 3,
            "matrix": [[{"display_text": "", "role": "blank"}] * 3],
        }
        with patch.object(FakeHwp, "GetPos", lambda self: (self.list_id, 0, 0)):
            hwp, _report, inserted = insert_fake_clipboard_table(table)
        self.assertTrue(inserted)
        self.assertNotIn(("run", "Delete"), hwp.events)

    def test_clipboard_entry_exception_rolls_back_and_failed_undo_aborts(self):
        table = {
            "table_key": "T001", "cell_contract": True,
            "source_sheet": "Sheet1", "source_range": "A1:C1",
            "row_count": 1, "col_count": 3,
            "matrix": [[{"display_text": "항목", "role": "stub"}] * 3],
        }
        with patch.object(FakeHwp, "FindCtrl", side_effect=RuntimeError("COM failure")):
            hwp, _report, inserted = insert_fake_clipboard_table(table)
        self.assertFalse(inserted)
        self.assertIn(("run", "Undo"), hwp.events)
        original_run = FakeAction.Run
        with (
            patch.object(FakeHwp, "FindCtrl", return_value=False),
            patch.object(FakeAction, "Run", lambda action, name: False if name == "Undo" else original_run(action, name)),
        ):
            with self.assertRaisesRegex(HwpWriterError, "되돌리지"):
                insert_fake_clipboard_table(table)

    def test_clipboard_text_selection_failure_aborts_before_replacement(self):
        table = {
            "table_key": "T001", "cell_contract": True,
            "source_sheet": "Sheet1", "source_range": "A1:C1",
            "row_count": 1, "col_count": 3,
            "matrix": [[{"display_text": "항목", "role": "stub"}] * 3],
        }
        original_run = FakeAction.Run
        with patch.object(FakeAction, "Run", lambda action, name: False if name == "MoveSelListEnd" else original_run(action, name)):
            with self.assertRaisesRegex(HwpWriterError, "선택하지"):
                insert_fake_clipboard_table(table)

    def test_clipboard_table_skips_covered_merged_slots(self):
        table = {
            "table_key": "T001",
            "cell_contract": True,
            "source_sheet": "Sheet1",
            "source_range": "A1:C1",
            "row_count": 1,
            "col_count": 3,
            "matrix": [[
                {"display_text": "병합 배너", "role": "banner_horizontal", "colspan": 2},
                {"display_text": "", "role": "blank", "covered_by": "A1"},
                {"display_text": "다음 실제 셀", "role": "value"},
            ]],
        }

        hwp, _report, inserted = insert_fake_clipboard_table(table)

        self.assertTrue(inserted)
        self.assertEqual(
            [event for event in hwp.events if event[0] in {"style", "text"}],
            [("style", 13), ("text", "병합 배너"), ("style", 14), ("text", "다음 실제 셀")],
        )
        self.assertEqual(sum(event == ("run", "Delete") for event in hwp.events), 2)
        self.assertEqual(sum(event == ("run", "TableRightCell") for event in hwp.events), 1)

    def test_clipboard_vertical_merges_skip_revisited_cell_lists(self):
        table = {
            "table_key": "T001", "cell_contract": True,
            "source_sheet": "Sheet1", "source_range": "A1:B3",
            "row_count": 3, "col_count": 2,
            "matrix": [
                [{"display_text": "A", "role": "banner_horizontal", "rowspan": 2}, {"display_text": "B", "role": "banner_horizontal", "rowspan": 2}],
                [{"covered_by": "A1"}, {"covered_by": "B1"}],
                [{"display_text": "C", "role": "stub"}, {"display_text": "D", "role": "value"}],
            ],
        }
        original_run = FakeAction.Run
        moves = iter([3, 2, 3, 4, 5])

        def run(action, name):
            result = original_run(action, name)
            if name == "TableRightCell":
                action.owner.list_id = next(moves)
                action.owner.position = (action.owner.list_id, 0, 0)
            return result

        with (
            patch.object(FakeAction, "Run", run),
            patch.object(FakeTableControl, "Properties", SimpleNamespace(Rows=3, Cols=2)),
        ):
            hwp, _report, inserted = insert_fake_clipboard_table(table)
        self.assertTrue(inserted)
        self.assertEqual([event[1] for event in hwp.events if event[0] == "text"], ["A", "B", "C", "D"])
        self.assertEqual(sum(event == ("run", "TableRightCell") for event in hwp.events), 5)
        self.assertEqual(hwp.contents, {2: "A", 3: "B", 4: "C", 5: "D"})

    def test_clipboard_navigation_failure_and_cycles_abort(self):
        table = {
            "table_key": "T001", "cell_contract": True,
            "source_sheet": "Sheet1", "source_range": "A1:C1",
            "row_count": 1, "col_count": 3,
            "matrix": [[{"display_text": "항목", "role": "stub"}] * 3],
        }
        original_run = FakeAction.Run
        with patch.object(FakeAction, "Run", lambda action, name: False if name == "TableRightCell" else original_run(action, name)):
            with self.assertRaisesRegex(HwpWriterError, "이동하지"):
                insert_fake_clipboard_table(table)
        with patch.object(FakeHwp, "GetPos", return_value=(2, 0, 0)):
            with self.assertRaisesRegex(HwpWriterError, "반복되어"):
                insert_fake_clipboard_table(table)
        with patch.object(FakeHwp, "GetPos", side_effect=RuntimeError("COM disconnected")):
            with self.assertRaisesRegex(HwpWriterError, "중복 표"):
                insert_fake_clipboard_table(table)

        def leave_table(action, name):
            result = original_run(action, name)
            if name == "TableRightCell":
                action.owner.ParentCtrl = None
            return result

        with patch.object(FakeAction, "Run", leave_table):
            with self.assertRaisesRegex(HwpWriterError, "표 밖"):
                insert_fake_clipboard_table(table)

    def test_temporary_files_follow_hwp_lifecycle(self):
        cases = (
            ("ready", False, False, False),
            ("ready", False, True, True),
            ("failed", True, False, True),
        )
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            for index, (status, keep_on_error, keep_after_save, retained) in enumerate(cases):
                root = parent / str(index)
                root.mkdir()
                (root / "styled_template.hwpx").write_bytes(b"template")
                hwp = SimpleNamespace(Clear=lambda *_args: None, Quit=lambda: None)
                report = {"status": status, "com": {"closed": False}, "warnings": []}

                finalize_hwp_resources(root, hwp, report, keep_on_error, keep_after_save)

                self.assertEqual(root.exists(), retained)
                self.assertEqual(report["temporary_files_retained"], retained)
                self.assertEqual(report["com"]["closed"], not retained)

            failed_cleanup = parent / "cleanup-failure"
            failed_cleanup.mkdir()
            report = {"status": "ready", "com": {"closed": False}, "warnings": []}
            hwp = SimpleNamespace(Clear=lambda *_args: None, Quit=lambda: None)
            with patch("report_automation_engine.hwp_com_writer.shutil.rmtree", side_effect=OSError("locked")):
                finalize_hwp_resources(failed_cleanup, hwp, report, False, False)
            self.assertTrue(report["temporary_files_retained"])
            self.assertIn("locked", report["warnings"][-1])

    def test_style_registration_failure_does_not_create_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            package = root / "package.json"
            preflight = root / "preflight.json"
            template = root / "template.hwpx"
            output = root / "output.hwpx"
            package.write_text('{"meta": {}, "sections": [], "tables": [], "charts": [], "qa": []}', encoding="utf-8")
            preflight.write_text('{"status": "ready"}', encoding="utf-8")
            template.write_bytes(b"not-used-because-registration-is-mocked")

            with (
                patch("report_automation_engine.hwp_com_writer.platform.system", return_value="Windows"),
                patch(
                    "report_automation_engine.hwp_com_writer.register_named_styles",
                    side_effect=ValueError("invalid style registry"),
                ),
                patch("report_automation_engine.hwp_com_writer.create_hwp_object") as create_hwp,
            ):
                with self.assertRaises(HwpWriterError):
                    write_hwp_document(package, preflight, template, output)

            self.assertFalse(output.exists())
            create_hwp.assert_not_called()


if __name__ == "__main__":
    unittest.main()
