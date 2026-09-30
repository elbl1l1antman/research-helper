import unittest
from pathlib import Path
import sys
from unittest.mock import patch

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from report_automation_engine.hwp_com_writer import insert_narrative_blocks
from report_automation_engine.report_package import read_sections, split_narrative_blocks


class NarrativeBlockTests(unittest.TestCase):
    def test_splits_all_newline_forms_in_order(self):
        blocks = split_narrative_blocks("첫째\r\n둘째\n셋째\r넷째")

        self.assertEqual([block["text"] for block in blocks], ["첫째", "둘째", "셋째", "넷째"])
        self.assertEqual(
            [block["style"] for block in blocks],
            ["보고서 본문1", "보고서 본문2", "보고서 본문2", "보고서 본문2"],
        )

    def test_recovers_only_specified_lost_transition_boundaries(self):
        text = "가장 높게 나타남다음으로 A 순임나타남그다음으로 B 순임나타남반면, C는 낮음"
        blocks = split_narrative_blocks(text)

        self.assertEqual(
            [block["text"] for block in blocks],
            ["가장 높게 나타남", "다음으로 A 순임나타남", "그다음으로 B 순임나타남", "반면, C는 낮음"],
        )

    def test_single_sentence_remains_body1(self):
        self.assertEqual(
            split_narrative_blocks("한 문장만 있음"),
            [{"style": "보고서 본문1", "text": "한 문장만 있음"}],
        )

    def test_read_sections_preserves_original_and_adds_three_styled_blocks(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "보고서_분석문"
        sheet.append(["table_key", "문항/표 제목", "최종 사용문"])
        narrative = "첫 문장\n둘째 문장\n셋째 문장"
        sheet.append(["T001", "만족도", narrative])

        section = read_sections(workbook, [])[0]

        self.assertEqual(section["narrative_final"], narrative)
        self.assertEqual(
            [block["style"] for block in section["narrative_blocks"]],
            ["보고서 본문1", "보고서 본문2", "보고서 본문2"],
        )

    def test_writer_inserts_one_paragraph_per_block_without_literal_body2_bullet(self):
        section = {
            "narrative_blocks": [
                {"style": "보고서 본문1", "text": "첫 문장"},
                {"style": "보고서 본문2", "text": "- 다음 문장"},
            ]
        }
        inserted = []
        actions = []

        with (
            patch("report_automation_engine.hwp_com_writer.insert_text", side_effect=lambda _h, text, _r: inserted.append(text)),
            patch("report_automation_engine.hwp_com_writer.run_action", side_effect=lambda _h, action, _r, _s: actions.append(action)),
        ):
            insert_narrative_blocks(object(), section, {})

        self.assertEqual(inserted, ["첫 문장", "다음 문장"])
        self.assertEqual(actions, ["BreakPara", "BreakPara"])


if __name__ == "__main__":
    unittest.main()
