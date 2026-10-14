import copy
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from report_automation_engine.hwp_table_layout import plan_table_parts


def sample_table():
    matrix = []
    for row in range(3):
        matrix.append([
            {"row": row + 1, "col": col + 1, "rowspan": 1, "colspan": 1,
             "covered_by": "", "source_cell": f"r{row}c{col}",
             "display_text": ("item" if col == 0 else f"{col}.0") if row else f"header {col}",
             "role": "banner_horizontal" if row == 0 else "stub" if col == 0 else "value"}
            for col in range(10)
        ])
    matrix[0][1].update(display_text="group", colspan=9)
    for col in range(2, 10):
        matrix[0][col].update(display_text="", covered_by="r0c1")
    return {"table_key": "T1", "matrix": matrix, "row_count": 3, "col_count": 10, "merged_ranges": [{"row": 1, "col": 2, "rowspan": 1, "colspan": 9}]}


class TableLayoutTests(unittest.TestCase):
    def test_header_only_merged_banner_keeps_existing_split_behavior(self):
        table = sample_table()
        table["matrix"] = table["matrix"][:1]
        table["row_count"] = 1
        parts = plan_table_parts(table, 18000, 8.5)
        self.assertEqual(len(parts), 3)
        self.assertEqual([col for part in parts for col in part["source_columns"]], list(range(1, 11)))

    def test_mixed_header_repeats_vertical_label_and_clips_horizontal_banner(self):
        table = sample_table()
        table["matrix"][0][0].update(role="banner_vertical", rowspan=3, display_text="Category")
        for row in table["matrix"][1:]:
            row[0].update(role="blank", covered_by="r0c0", display_text="")
        parts = plan_table_parts(table, 18000, 8.5)
        self.assertEqual(len(parts), 3)
        for part in parts:
            self.assertEqual(part["source_columns"][0], 1)
            self.assertEqual(part["matrix"][0][0]["rowspan"], 3)
            self.assertEqual(part["matrix"][0][1]["display_text"], "group")
            self.assertEqual(part["matrix"][0][1]["colspan"], part["col_count"] - 1)

    def test_split_repeats_stub_and_clipped_banner_without_losing_values(self):
        table = sample_table()
        before = copy.deepcopy(table)
        parts = plan_table_parts(table, 18000, 8.5)
        self.assertGreater(len(parts), 1)
        originals = []
        for part in parts:
            self.assertLessEqual(sum(part["column_widths_hwpunit"]), 18000)
            self.assertEqual(part["source_columns"][0], 1)
            self.assertEqual(part["matrix"][0][1]["display_text"], "group")
            self.assertEqual(part["matrix"][0][1]["colspan"], part["col_count"] - 1)
            originals.extend(part["source_columns"][1:])
        self.assertEqual(originals, list(range(2, 11)))
        self.assertEqual(table, before)

    def test_numeric_base_and_vertical_merges_are_repeated(self):
        table = sample_table()
        table["matrix"][0][1].update(role="base", colspan=1, display_text="N")
        for col in range(2, 10):
            table["matrix"][0][col].update(covered_by="", display_text=f"H{col}")
        table["matrix"][1][0].update(rowspan=2)
        table["matrix"][2][0].update(covered_by="r1c0", display_text="")
        parts = plan_table_parts(table, 20000, 8.5)
        for part in parts:
            self.assertEqual(part["source_columns"][:2], [1, 2])
            self.assertEqual(part["matrix"][1][0]["rowspan"], 2)
            self.assertTrue(part["matrix"][2][0]["covered_by"])

    def test_body_merge_cannot_be_silently_cut_to_fit(self):
        table = sample_table()
        table["matrix"][1][1].update(colspan=9)
        for col in range(2, 10):
            table["matrix"][1][col].update(covered_by="r1c1", display_text="")
        with self.assertRaisesRegex(ValueError, "병합"):
            plan_table_parts(table, 18000, 8.5)

    def test_unfit_repeated_columns_and_invalid_width_are_blocked(self):
        with self.assertRaises(ValueError):
            plan_table_parts(sample_table(), 0, 8.5)
        with self.assertRaises(ValueError):
            plan_table_parts(sample_table(), 1000, 8.5)

    def test_body_merge_crossing_repeated_prefix_is_blocked(self):
        table = sample_table()
        table["matrix"][1][0].update(colspan=2)
        table["matrix"][1][1].update(covered_by="r1c0", display_text="")
        with self.assertRaisesRegex(ValueError, "병합"):
            plan_table_parts(table, 18000, 8.5)

    def test_merged_body_label_before_first_numeric_row_cannot_be_cut(self):
        table = sample_table()
        label = copy.deepcopy(table["matrix"][1])
        label[1].update(role="stub", colspan=9, display_text="Body section")
        for col in range(2, 10):
            label[col].update(role="blank", covered_by="r1c1", display_text="")
        table["matrix"].insert(1, label)
        with self.assertRaises(ValueError):
            plan_table_parts(table, 18000, 8.5)

    def test_base_role_in_body_is_repeated_in_every_part(self):
        table = sample_table()
        for row in table["matrix"][1:]:
            row[1]["role"] = "base"
        for part in plan_table_parts(table, 20000, 8.5):
            self.assertEqual(part["source_columns"][:2], [1, 2])


if __name__ == "__main__":
    unittest.main()
