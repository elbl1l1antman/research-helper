from __future__ import annotations

import copy
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from report_automation_engine.hwp_style_config import (
    DEFAULT_HWP_STYLE_CONFIG,
    REQUIRED_STYLE_NAMES,
    load_hwp_style_config,
    style_name_for_cell_role,
    validate_hwp_style_config,
)


class HwpStyleConfigTests(unittest.TestCase):
    def test_default_contract_and_role_mapping(self) -> None:
        self.assertEqual(
            REQUIRED_STYLE_NAMES,
            ("보고서 본문1", "보고서 본문2", "표보기", "표배너", "표숫자"),
        )
        self.assertEqual(DEFAULT_HWP_STYLE_CONFIG["schema_version"], "1.0")
        self.assertEqual(DEFAULT_HWP_STYLE_CONFIG["preset_name"], "기본 보고서")
        common_body = {
            "font_family": "맑은 고딕",
            "font_size_pt": 10.0,
            "bold": False,
            "bullet": "",
            "line_spacing_percent": 160,
            "left_indent_mm": 0.0,
            "first_line_indent_mm": 0.0,
            "alignment": "left",
        }
        common_table = {
            "font_family": "맑은 고딕",
            "font_size_pt": 8.5,
            "bold": False,
            "bullet": "",
            "line_spacing_percent": 130,
            "left_indent_mm": 0.0,
            "first_line_indent_mm": 0.0,
            "alignment": "left",
        }
        styles = DEFAULT_HWP_STYLE_CONFIG["paragraph_styles"]
        self.assertEqual(styles["보고서 본문1"], common_body)
        self.assertEqual(
            styles["보고서 본문2"],
            {**common_body, "bullet": "-", "left_indent_mm": 5.0, "first_line_indent_mm": -5.0},
        )
        self.assertEqual(styles["표보기"], common_table)
        self.assertEqual(styles["표배너"], {**common_table, "bold": True, "alignment": "center"})
        self.assertEqual(styles["표숫자"], {**common_table, "alignment": "center"})
        self.assertEqual(
            DEFAULT_HWP_STYLE_CONFIG["table_cell_styles"],
            {
                "표보기": {"fill_color": "FFFFFF", "vertical_alignment": "center"},
                "표배너": {"fill_color": "E7E7E7", "vertical_alignment": "center"},
                "표숫자": {"fill_color": "FFFFFF", "vertical_alignment": "center"},
            },
        )
        self.assertEqual(
            DEFAULT_HWP_STYLE_CONFIG["table_border"],
            {"inner_width_mm": 0.12, "outer_width_mm": 0.4, "color": "000000"},
        )

        expected = {
            "title": "표배너",
            "base": "표배너",
            "header": "표배너",
            "banner_horizontal": "표배너",
            "banner_vertical": "표보기",
            "stub": "표보기",
            "note": "표보기",
            "source": "표보기",
            "unknown": "표보기",
            "blank": "표보기",
            "value": "표숫자",
        }
        self.assertEqual({role: style_name_for_cell_role(role) for role in expected}, expected)

    def test_validation_normalizes_a_deep_copy_and_rejects_bad_fields(self) -> None:
        source = copy.deepcopy(DEFAULT_HWP_STYLE_CONFIG)
        normalized = validate_hwp_style_config(source)
        normalized["paragraph_styles"]["보고서 본문1"]["font_family"] = "변경"
        self.assertEqual(source["paragraph_styles"]["보고서 본문1"]["font_family"], "맑은 고딕")

        invalid_cases = (
            ("missing style", lambda config: config["paragraph_styles"].pop("표숫자"), "표숫자"),
            (
                "bad color",
                lambda config: config["table_cell_styles"]["표배너"].__setitem__("fill_color", "GGGGGG"),
                "table_cell_styles.표배너.fill_color",
            ),
            (
                "large font",
                lambda config: config["paragraph_styles"]["보고서 본문1"].__setitem__("font_size_pt", 31),
                "paragraph_styles.보고서 본문1.font_size_pt",
            ),
            (
                "large spacing",
                lambda config: config["paragraph_styles"]["보고서 본문2"].__setitem__("line_spacing_percent", 301),
                "paragraph_styles.보고서 본문2.line_spacing_percent",
            ),
        )
        for label, mutate, field in invalid_cases:
            with self.subTest(label=label):
                config = copy.deepcopy(DEFAULT_HWP_STYLE_CONFIG)
                mutate(config)
                with self.assertRaisesRegex(ValueError, field.replace(".", r"\.")):
                    validate_hwp_style_config(config)

        range_fields = (
            (("paragraph_styles", "보고서 본문1"), "left_indent_mm", -30, 100),
            (("paragraph_styles", "보고서 본문1"), "first_line_indent_mm", -30, 100),
            (("table_border",), "inner_width_mm", 0.1, 5.0),
            (("table_border",), "outer_width_mm", 0.1, 5.0),
        )
        for path, field, minimum, maximum in range_fields:
            for value in (minimum, maximum):
                with self.subTest(field=field, accepted=value):
                    config = copy.deepcopy(DEFAULT_HWP_STYLE_CONFIG)
                    target = config
                    for key in path:
                        target = target[key]
                    target[field] = value
                    validate_hwp_style_config(config)
            for value in (minimum - 0.01, maximum + 0.01):
                error_field = ".".join((*path, field))
                with self.subTest(field=field, rejected=value):
                    config = copy.deepcopy(DEFAULT_HWP_STYLE_CONFIG)
                    target = config
                    for key in path:
                        target = target[key]
                    target[field] = value
                    with self.assertRaisesRegex(ValueError, error_field.replace(".", r"\.")):
                        validate_hwp_style_config(config)

    def test_loads_default_or_json_file(self) -> None:
        loaded_default = load_hwp_style_config(None)
        self.assertEqual(loaded_default, DEFAULT_HWP_STYLE_CONFIG)
        self.assertIsNot(loaded_default, DEFAULT_HWP_STYLE_CONFIG)

        custom = copy.deepcopy(DEFAULT_HWP_STYLE_CONFIG)
        custom["preset_name"] = "사용자 1"
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "style.json"
            path.write_text(json.dumps(custom, ensure_ascii=False), encoding="utf-8")
            self.assertEqual(load_hwp_style_config(path)["preset_name"], "사용자 1")


if __name__ == "__main__":
    unittest.main()
