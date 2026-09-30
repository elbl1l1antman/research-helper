import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from report_automation_engine.report_table_matrix import (
    build_table_matrix_from_cells,
    sanitize_table_display_text,
)


class TableDisplayTextTests(unittest.TestCase):
    def test_preserves_display_punctuation_and_units(self):
        self.assertEqual(sanitize_table_display_text("(3,232)"), "(3,232)")
        self.assertEqual(sanitize_table_display_text("-3232"), "-3232")
        self.assertEqual(sanitize_table_display_text("● 만족도 3.3%"), "만족도 3.3%")
        self.assertEqual(
            sanitize_table_display_text("&#x20;'-' \"문장\" ~ ㎡ ㎏"),
            "'-' \"문장\" ~ ㎡ ㎏",
        )

    def test_symbols_only_cell_is_empty_and_warns_with_source_address(self):
        table = build_table_matrix_from_cells(
            [
                {
                    "table_key": "T001",
                    "title": "만족도",
                    "row": 1,
                    "col": 1,
                    "role": "stub",
                    "display_text": "●◆",
                    "raw_value": "●◆",
                    "source_cell": "D12",
                    "source_sheet": "Sheet1",
                    "source_range": "D12",
                }
            ]
        )

        self.assertEqual(table["matrix"][0][0]["display_text"], "")
        self.assertEqual(table["matrix"][0][0]["original_display_text"], "●◆")
        self.assertEqual(table["matrix"][0][0]["removed_symbols"], "●◆")
        self.assertTrue(any("D12" in item["message"] for item in table["qa"]))


if __name__ == "__main__":
    unittest.main()
