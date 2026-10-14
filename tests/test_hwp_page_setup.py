import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from report_automation_engine.hwp_page_setup import normalize_page_setup, apply_page_setup
from report_automation_engine.hwp_com_writer import HwpWriterError, write_hwp_document


SETUP = dict(paper_width_mm=210, paper_height_mm=297, orientation="portrait",
             top_mm=20, bottom_mm=15, left_mm=24.7, right_mm=25,
             gutter_mm=0, header_mm=10, footer_mm=15)


class PageSetupTests(unittest.TestCase):
    def test_validation_and_body_geometry(self):
        self.assertEqual(normalize_page_setup(SETUP), SETUP)
        for field, value in [("left_mm", -1), ("paper_width_mm", 0),
                             ("top_mm", float("nan")), ("gutter_mm", True),
                             ("left_mm", 210), ("header_mm", 21),
                             ("orientation", "unknown")]:
            with self.subTest(field=field, value=value), self.assertRaises(ValueError):
                normalize_page_setup(dict(SETUP, **{field: value}))
        with self.assertRaises(ValueError):
            normalize_page_setup(dict(SETUP, typo=1))
        missing = dict(SETUP)
        del missing["footer_mm"]
        with self.assertRaises(ValueError):
            normalize_page_setup(missing)
        with self.assertRaises(ValueError):
            normalize_page_setup(dict(SETUP, paper_width_mm=1e20))
        with self.assertRaises(ValueError):
            normalize_page_setup(dict(SETUP, paper_width_mm=49.700001))

    def test_apply_current_section_and_verify(self):
        page = SimpleNamespace()
        params = SimpleNamespace(PageDef=page)
        params.HSet = SimpleNamespace(SetItem=lambda key, val: setattr(params, key, val))
        hwp = SimpleNamespace(HParameterSet=SimpleNamespace(HSecDef=params),
                              HAction=SimpleNamespace(GetDefault=lambda *args: True,
                                                      Execute=lambda *args: True))
        result = apply_page_setup(hwp, dict(SETUP, orientation="landscape"))
        self.assertEqual(params.ApplyTo, 2)
        self.assertEqual(page.GutterType, 0)
        self.assertEqual(page.PaperWidth, round(210 * 7200 / 25.4))
        self.assertEqual(page.Landscape, 1)
        self.assertEqual(result["body_width_hwpunit"], page.PaperHeight - page.LeftMargin - page.RightMargin)
        self.assertEqual(result["body_height_hwpunit"], page.PaperWidth - page.TopMargin - page.BottomMargin)
        hwp.HAction.Execute = lambda *args: False
        with self.assertRaises(ValueError):
            apply_page_setup(hwp, SETUP)
        hwp.HAction.Execute = lambda *args: True
        hwp.HAction.GetDefault = lambda *args: setattr(page, "LeftMargin", 0)
        with self.assertRaises(ValueError):
            apply_page_setup(hwp, SETUP)

    def test_invalid_config_fails_before_com_and_writes_report(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for filename, data in [("package.json", {}), ("preflight.json", {"status": "ready"}),
                                   ("page.json", dict(SETUP, left_mm=999))]:
                (root / filename).write_text(json.dumps(data), encoding="utf-8")
            with patch("report_automation_engine.hwp_com_writer.create_hwp_object") as com:
                with self.assertRaises(HwpWriterError):
                    write_hwp_document(root / "package.json", root / "preflight.json",
                                       root / "template.hwpx", root / "output.hwpx",
                                       page_setup_path=root / "page.json")
                com.assert_not_called()
            report = json.loads((root / "output_hwp_writer_report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["action"], "PageSetup")
            self.assertFalse((root / "output.hwpx").exists())

    def test_dry_run_records_requested_settings_without_com(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for filename, data in [("package.json", {"sections": [], "tables": []}),
                                   ("preflight.json", {"status": "ready"}), ("page.json", SETUP)]:
                (root / filename).write_text(json.dumps(data), encoding="utf-8")
            with patch("report_automation_engine.hwp_com_writer.create_hwp_object") as com:
                for config in [None, root / "page.json"]:
                    write_hwp_document(root / "package.json", root / "preflight.json",
                                       root / "template.hwpx", root / "output.hwpx",
                                       dry_run=True, page_setup_path=config)
                    report = json.loads((root / "output_hwp_writer_report.json").read_text(encoding="utf-8"))
                    self.assertFalse(report["page_setup"]["applied"])
                    if config is None:
                        self.assertEqual(report["page_setup"]["mode"], "template")
                    else:
                        self.assertEqual(report["page_setup"]["requested"], SETUP)
                com.assert_not_called()


if __name__ == "__main__":
    unittest.main()
