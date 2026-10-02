import copy
from pathlib import Path
import tempfile
import unittest

import openpyxl

from report_automation_engine import hwp_table_layout as layout
from report_automation_engine import report_table_matrix as matrix_tools
from report_automation_engine.report_package import build_report_package
from tests.test_hwp_table_layout import sample_table


class TableMetadataTests(unittest.TestCase):
    def resolve(self, table, metadata=None):
        self.assertTrue(hasattr(matrix_tools, "resolve_table_metadata"))
        return matrix_tools.resolve_table_metadata(table, metadata)

    def test_explicit_conditions_and_units_override_inferred_headers(self):
        table = {"base_label": "[BASE : 전체응답자(n=454)]", "unit": "명, %, 점", "matrix": []}
        self.assertEqual(self.resolve(table), {"base_label": "전체응답자(n=454)", "unit": "명, %, 점"})
        self.assertEqual(self.resolve(table, {"base_label": "BASE:이용자", "unit": "건"}), {"base_label": "이용자", "unit": "건"})

    def test_header_metadata_is_parsed_without_guessing_a_population(self):
        table = {"matrix": [[{"role": "base", "display_text": "[BASE : 경험자(n=120), 단위 : 명, %, 점]"}]]}
        self.assertEqual(self.resolve(table), {"base_label": "경험자(n=120)", "unit": "명, %, 점"})
        table = {"matrix": [[{"role": "base", "display_text": "사례수"}]], "rows": []}
        self.assertEqual(self.resolve(table), {"base_label": "확인 필요", "unit": "확인 필요"})
        table = {"base_label": "전체", "matrix": [[{"role": "note", "display_text": "(단위 : 명, %)"}]]}
        self.assertEqual(self.resolve(table), {"base_label": "전체", "unit": "명, %"})

    def test_package_reads_table_list_metadata_without_changing_source_matrix(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.xlsx"
            from tests.test_report_package_table_matrix import make_cell_contract_workbook
            make_cell_contract_workbook(path)
            workbook = openpyxl.load_workbook(path)
            listing = workbook.create_sheet("보고서_표목록")
            listing.append(["table_key", "base_label", "unit"])
            listing.append(["T001", "BASE : 이용자", "명, %"])
            workbook.save(path)
            workbook.close()
            table = build_report_package(path)["tables"][0]
            self.assertEqual(table.get("base_label"), "이용자")
            self.assertEqual(table.get("unit"), "명, %")
            self.assertEqual(table["row_count"], 2)
            self.assertEqual(table["matrix"][1][1]["display_text"], "63.3%")

    def test_render_row_is_merged_and_source_table_remains_unchanged(self):
        self.assertTrue(hasattr(layout, "prepare_report_table"))
        table = layout.plan_table_parts(sample_table(), 18000, 8.5)[0]
        table.update(base_label="전체", unit="명")
        before = copy.deepcopy(table)
        rendered = layout.prepare_report_table(table)
        self.assertEqual(table, before)
        self.assertEqual(rendered["row_count"], table["row_count"] + 1)
        self.assertEqual(rendered["matrix"][0][0]["display_text"], "[BASE : 전체 | 단위 : 명]")
        self.assertEqual(rendered["matrix"][0][0]["role"], "metadata")
        self.assertEqual(rendered["matrix"][0][0]["colspan"], table["col_count"])
        self.assertEqual(rendered["header_row_count"], 2)
        self.assertEqual(rendered["matrix"][2][1]["display_text"], table["matrix"][1][1]["display_text"])
        self.assertEqual(rendered["merged_ranges"][1]["row"], table["merged_ranges"][0]["row"] + 1)

    def test_body_crossing_header_merge_does_not_repeat_body_values(self):
        self.assertTrue(hasattr(layout, "prepare_report_table"))
        table = sample_table()
        table["matrix"][0][0].update(role="banner_vertical", rowspan=3)
        for row in table["matrix"][1:]:
            row[0].update(role="blank", covered_by="r0c0", display_text="")
        rendered = layout.prepare_report_table(table)
        self.assertEqual(rendered["header_row_count"], 1)
        self.assertTrue(rendered.get("header_repeat_warning"))


if __name__ == "__main__":
    unittest.main()
