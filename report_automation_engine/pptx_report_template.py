"""Build the bundled PPTX report template (``templates/report_template_basic.pptx``).

The template follows the page frame of public-sector survey reports: a header
band with the section number and title, the part name on the right, a key
message line, a unit label, a table/chart area, the source line and a slide
number. Slides: cover, contents, part divider, repeat (content) slide, QA.

Run ``python -m report_automation_engine.pptx_report_template [output]`` to
rebuild it. Token meanings are documented in ``document_writer.py``.
"""

from __future__ import annotations

import sys
import uuid
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

FONT = "맑은 고딕"
PRIMARY = RGBColor(0x0B, 0x4E, 0xA2)
PRIMARY_LIGHT = RGBColor(0xE8, 0xF0, 0xFB)
INK = RGBColor(0x22, 0x22, 0x22)
MUTED = RGBColor(0x6B, 0x72, 0x80)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xB8, 0xC4, 0xD6)
DEFAULT_OUTPUT = Path(__file__).resolve().parents[1] / "templates" / "report_template_basic.pptx"


def set_font(run, size, color, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        node = rpr.find(qn(tag))
        if node is None:
            node = etree.SubElement(rpr, qn(tag))
        node.set("typeface", FONT)


def text(slide, value, left, top, width, height, size, color=INK, *, bold=False, name=None,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    if name:
        shape.name = name
    frame = shape.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    frame.margin_left = frame.margin_right = Inches(0.05)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    set_font(paragraph.add_run(), size, color, bold)
    paragraph.runs[0].text = value
    return shape


def panel(slide, left, top, width, height, color, name=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    if name:
        shape.name = name
    return shape


def block(slide, token, left, top, width, height):
    """Dashed outline holding {{TABLE}}/{{CHART}}; the writer replaces it with a native object."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.name = "RA_" + token
    shape.fill.background()
    shape.line.color.rgb = LINE
    shape.line.dash_style = MSO_LINE.DASH
    shape.shadow.inherit = False
    frame = shape.text_frame
    set_font(frame.paragraphs[0].add_run(), 12, MUTED)
    frame.paragraphs[0].runs[0].text = "{{" + token + "}}"
    frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    return shape


def slide_number(slide, left, top, width, height):
    shape = text(slide, "", left, top, width, height, 9, MUTED, align=PP_ALIGN.RIGHT, name="RA_SLIDE_NUMBER")
    paragraph = shape.text_frame.paragraphs[0]._p
    run = paragraph.find(qn("a:r"))
    field = etree.SubElement(paragraph, qn("a:fld"), id="{" + str(uuid.uuid4()).upper() + "}", type="slidenum")
    field.append(run.find(qn("a:rPr")))
    etree.SubElement(field, qn("a:t")).text = "‹#›"
    paragraph.remove(run)
    paragraph.insert(0, field)
    return shape


def build(output: str | Path = DEFAULT_OUTPUT) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]  # "Blank"

    def add_slide():
        slide = prs.slides.add_slide(blank)
        for shape in list(slide.placeholders):
            shape._element.getparent().remove(shape._element)
        return slide

    cover = add_slide()
    text(cover, "{{REPORT_TITLE}}", 0.8, 2.3, 7.4, 1.5, 34, PRIMARY, bold=True, anchor=MSO_ANCHOR.BOTTOM)
    text(cover, "{{PROJECT_NAME}}", 0.8, 1.7, 7.4, 0.5, 16, MUTED)
    text(cover, "{{GENERATED_AT}}", 0.8, 4.0, 7.4, 0.4, 12, MUTED)
    panel(cover, 8.9, 0, 4.433, 7.5, PRIMARY)

    contents = add_slide()
    panel(contents, 0, 0, 4.4, 7.5, PRIMARY_LIGHT)
    text(contents, "Contents", 0.8, 2.9, 3.4, 1.0, 36, PRIMARY, bold=True)
    text(contents, "{{TOC}}", 5.2, 1.6, 7.3, 4.6, 20, INK, name="RA_TOC")

    divider = add_slide()
    text(divider, "PART", 0.8, 2.0, 3.0, 0.5, 22, PRIMARY)
    text(divider, "{{PART_NUMBER}}", 0.8, 2.5, 4.0, 1.4, 80, PRIMARY, bold=True)
    text(divider, "{{PART_TITLE}}", 0.8, 4.0, 7.2, 0.8, 28, INK, bold=True)
    panel(divider, 8.6, 0, 4.733, 7.5, PRIMARY)
    text(divider, "{{PART_ITEMS}}", 9.1, 1.6, 3.8, 4.8, 14, WHITE, name="RA_PART_ITEMS", anchor=MSO_ANCHOR.MIDDLE)

    page = add_slide()
    panel(page, 0, 0.25, 13.333, 0.75, PRIMARY, name="RA_HEADER_BAND")
    text(page, "{{SECTION_NUMBER}}", 0.35, 0.3, 0.9, 0.65, 26, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(page, "{{SECTION_TITLE}}", 1.25, 0.3, 8.6, 0.65, 22, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(page, "{{PART_TITLE}}", 9.9, 0.3, 3.1, 0.65, 11, WHITE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    panel(page, 0.42, 1.27, 0.12, 0.12, PRIMARY, name="RA_KEY_MARK")  # removed when there is no key message
    text(page, "{{KEY_MESSAGE}}", 0.6, 1.12, 12.3, 0.45, 15, INK, bold=True, name="RA_KEY_MESSAGE")
    text(page, "{{NARRATIVE}}", 0.6, 1.57, 12.3, 0.55, 11, MUTED, name="RA_NARRATIVE")
    text(page, "{{UNIT}}", 9.9, 2.1, 3.03, 0.3, 10, MUTED, align=PP_ALIGN.RIGHT, name="RA_UNIT")
    block(page, "TABLE", 0.4, 2.42, 6.25, 4.38)
    block(page, "CHART", 6.85, 2.42, 6.08, 4.38)
    text(page, "{{SOURCE}}", 0.4, 6.95, 10.5, 0.3, 9, MUTED, name="RA_SOURCE")
    slide_number(page, 11.4, 6.95, 1.53, 0.3)

    qa = add_slide()
    text(qa, "QA 검토 항목 (내부 검토용 · 숨김 슬라이드)", 0.6, 0.4, 12.0, 0.6, 22, PRIMARY, bold=True)
    text(qa, "{{QA_SUMMARY}}", 0.6, 1.2, 12.1, 5.6, 12, INK, name="RA_QA_SUMMARY")

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))
    return output


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUTPUT))
