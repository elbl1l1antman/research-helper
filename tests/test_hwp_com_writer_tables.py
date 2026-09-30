from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from report_automation_engine.hwp_com_writer import (
    find_hwp_table_control,
    is_hwp_table_control,
    same_hwp_control,
    table_cell_role,
    table_rows_for_hwp,
)


class FakeControl:
    def __init__(self, ctrl_id, instance_id=0):
        self.CtrlID = ctrl_id
        self.instance_id = instance_id

    def GetCtrlInstID(self):
        return self.instance_id


class HwpComWriterTableTests(unittest.TestCase):
    def test_table_rows_for_hwp_uses_cell_contract_display_text(self):
        table = {
            "matrix": [
                [
                    {"display_text": "구분", "raw_value": "구분"},
                    {"display_text": "비율", "raw_value": "비율"},
                ],
                [
                    {"display_text": "전체", "raw_value": "전체"},
                    {"display_text": "63.3%", "raw_value": 0.633335353},
                ],
            ]
        }

        self.assertEqual(table_rows_for_hwp(table), [["구분", "비율"], ["전체", "63.3%"]])

    def test_table_rows_for_hwp_does_not_truncate_wide_table(self):
        table = {
            "matrix": [
                [{"display_text": f"배너 {col}"} for col in range(1, 37)],
                [{"display_text": f"{col}.0"} for col in range(1, 37)],
            ]
        }

        rows = table_rows_for_hwp(table)

        self.assertEqual(len(rows), 2)
        self.assertEqual(len(rows[0]), 36)
        self.assertEqual(rows[0][-1], "배너 36")
        self.assertEqual(rows[1][-1], "36.0")

    def test_hwp_table_control_requires_tbl_control_id(self):
        self.assertTrue(is_hwp_table_control(FakeControl("tbl")))
        self.assertFalse(is_hwp_table_control(FakeControl("gso")))
        self.assertFalse(is_hwp_table_control(None))

    def test_find_hwp_table_control_checks_current_parent_and_last_controls(self):
        class FakeHwp:
            CurSelectedCtrl = FakeControl("gso")
            ParentCtrl = FakeControl("tbl")
            LastCtrl = FakeControl("gso")

        self.assertEqual(find_hwp_table_control(FakeHwp()).CtrlID, "tbl")

    def test_same_hwp_control_uses_control_instance_id(self):
        self.assertTrue(same_hwp_control(FakeControl("tbl", 10), FakeControl("tbl", 10)))
        self.assertFalse(same_hwp_control(FakeControl("tbl", 10), FakeControl("tbl", 11)))

    def test_table_cell_role_prefers_matrix_contract(self):
        table = {"matrix": [[{"display_text": "전체", "role": "base"}]]}

        self.assertEqual(table_cell_role(table, 0, 0), "base")


if __name__ == "__main__":
    unittest.main()
