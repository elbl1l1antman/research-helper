from __future__ import annotations

import copy
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from report_automation_engine.hwp_style_config import DEFAULT_HWP_STYLE_CONFIG, REQUIRED_STYLE_NAMES
from report_automation_engine.hwpx_style_registry import register_named_styles


HH = "http://www.hancom.co.kr/hwpml/2011/head"
HC = "http://www.hancom.co.kr/hwpml/2011/core"
HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"


HEADER_XML = f"""<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="{HH}" xmlns:hc="{HC}" xmlns:hp="{HP}" version="1.5" secCnt="1">
  <hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>
  <hh:refList>
    <hh:fontfaces itemCnt="2">
      <hh:fontface lang="HANGUL" fontCnt="1"><hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0"/></hh:fontface>
      <hh:fontface lang="LATIN" fontCnt="1"><hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0"/></hh:fontface>
    </hh:fontfaces>
    <hh:borderFills itemCnt="1"><hh:borderFill id="1"/></hh:borderFills>
    <hh:charProperties itemCnt="1">
      <hh:charPr id="0" height="1000" textColor="#000000" shadeColor="none" borderFillIDRef="1">
        <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
      </hh:charPr>
    </hh:charProperties>
    <hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties>
    <hh:numberings itemCnt="0"/>
    <hh:paraProperties itemCnt="1">
      <hh:paraPr id="0" tabPrIDRef="0" condense="0" textDir="LTR">
        <hh:align horizontal="LEFT" vertical="BASELINE"/>
        <hh:heading type="NONE" idRef="0" level="0"/>
        <hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/></hh:margin>
        <hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>
      </hh:paraPr>
    </hh:paraProperties>
    <hh:styles itemCnt="1"><hh:style id="0" type="PARA" name="바탕글" engName="Normal" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/></hh:styles>
  </hh:refList>
</hh:head>
"""


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def child(parent: ET.Element, name: str) -> ET.Element:
    return next(node for node in list(parent) if local_name(node.tag) == name)


def descendants(parent: ET.Element, name: str) -> list[ET.Element]:
    return [node for node in parent.iter() if local_name(node.tag) == name]


def write_fixture(path: Path, header: bytes | None = HEADER_XML.encode("utf-8")) -> list[str]:
    members = ["mimetype", "META-INF/container.xml", "Contents/header.xml", "Contents/section0.xml", "tail.bin"]
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("mimetype", "application/hwp+zip", compress_type=zipfile.ZIP_STORED)
        archive.writestr("META-INF/container.xml", b"<container/>")
        if header is not None:
            archive.writestr("Contents/header.xml", header)
        archive.writestr("Contents/section0.xml", b"<section/>")
        archive.writestr("tail.bin", b"unchanged payload")
    return [name for name in members if header is not None or name != "Contents/header.xml"]


class HwpxStyleRegistryTests(unittest.TestCase):
    def test_registers_five_styles_and_preserves_template(self) -> None:
        config = copy.deepcopy(DEFAULT_HWP_STYLE_CONFIG)
        with tempfile.TemporaryDirectory() as temp_dir:
            template = Path(temp_dir) / "template.hwpx"
            output = Path(temp_dir) / "working.hwpx"
            member_order = write_fixture(template)
            source_bytes = template.read_bytes()

            mapping = register_named_styles(template, output, config)

            self.assertEqual(template.read_bytes(), source_bytes)
            self.assertEqual(list(mapping), list(REQUIRED_STYLE_NAMES))
            self.assertEqual(len(set(mapping.values())), 5)
            with zipfile.ZipFile(output) as archive:
                self.assertEqual(archive.namelist(), member_order)
                self.assertEqual(archive.read("tail.bin"), b"unchanged payload")
                root = ET.fromstring(archive.read("Contents/header.xml"))

            ref_list = child(root, "refList")
            collections = {
                name: child(ref_list, name)
                for name in ("fontfaces", "charProperties", "paraProperties", "bullets", "styles")
            }
            for name, collection in collections.items():
                self.assertEqual(int(collection.attrib["itemCnt"]), len(list(collection)), name)
            for fontface in list(collections["fontfaces"]):
                self.assertEqual(int(fontface.attrib["fontCnt"]), len(list(fontface)))
                self.assertIn("맑은 고딕", {node.attrib["face"] for node in list(fontface)})

            styles = list(collections["styles"])
            by_name = {node.attrib["name"]: node for node in styles}
            self.assertEqual([styles[index].attrib["name"] for index in mapping.values()], list(REQUIRED_STYLE_NAMES))
            self.assertEqual(len([node for node in styles if node.attrib.get("name") in REQUIRED_STYLE_NAMES]), 5)

            char_prs = {node.attrib["id"]: node for node in collections["charProperties"]}
            para_prs = {node.attrib["id"]: node for node in collections["paraProperties"]}
            for style_name, expected in config["paragraph_styles"].items():
                style = by_name[style_name]
                char_pr = char_prs[style.attrib["charPrIDRef"]]
                para_pr = para_prs[style.attrib["paraPrIDRef"]]
                self.assertEqual(int(char_pr.attrib["height"]), round(expected["font_size_pt"] * 100))
                self.assertEqual(any(local_name(node.tag) == "bold" for node in list(char_pr)), expected["bold"])
                self.assertNotEqual(child(char_pr, "fontRef").attrib["hangul"], "0")
                self.assertTrue(all(node.attrib["value"] == str(expected["line_spacing_percent"]) for node in descendants(para_pr, "lineSpacing")))
                self.assertEqual(child(para_pr, "align").attrib["horizontal"], expected["alignment"].upper())
                margin = child(para_pr, "margin")
                self.assertEqual(int(child(margin, "left").attrib["value"]), round(expected["left_indent_mm"] * 7200 / 25.4))
                self.assertEqual(int(child(margin, "intent").attrib["value"]), round(expected["first_line_indent_mm"] * 7200 / 25.4))

            body2_para = para_prs[by_name["보고서 본문2"].attrib["paraPrIDRef"]]
            heading = child(body2_para, "heading")
            self.assertEqual(heading.attrib["type"], "BULLET")
            bullets = list(collections["bullets"])
            self.assertEqual(len(bullets), 1)
            self.assertEqual(bullets[0].attrib["id"], "1")
            self.assertEqual(bullets[0].attrib["id"], heading.attrib["idRef"])
            self.assertEqual(bullets[0].attrib["char"], "-")

    def test_updates_existing_named_style_without_duplication(self) -> None:
        config = copy.deepcopy(DEFAULT_HWP_STYLE_CONFIG)
        config["paragraph_styles"]["보고서 본문1"].update({"font_size_pt": 12.0, "bold": True})
        root = ET.fromstring(HEADER_XML)
        ref_list = child(root, "refList")
        styles = child(ref_list, "styles")
        existing = copy.deepcopy(list(styles)[0])
        existing.attrib.update({"id": "9", "name": "보고서 본문1"})
        styles.insert(0, existing)
        styles.attrib["itemCnt"] = str(len(list(styles)))

        with tempfile.TemporaryDirectory() as temp_dir:
            template = Path(temp_dir) / "template.hwpx"
            first = Path(temp_dir) / "first.hwpx"
            second = Path(temp_dir) / "second.hwpx"
            write_fixture(template, ET.tostring(root, encoding="utf-8", xml_declaration=True))

            first_mapping = register_named_styles(template, first, config)
            second_mapping = register_named_styles(first, second, config)

            self.assertEqual(first_mapping, second_mapping)
            self.assertEqual(first_mapping["보고서 본문1"], 0)
            with zipfile.ZipFile(second) as archive:
                second_root = ET.fromstring(archive.read("Contents/header.xml"))
            second_ref_list = child(second_root, "refList")
            second_styles = child(second_ref_list, "styles")
            names = [node.attrib.get("name") for node in list(second_styles)]
            self.assertEqual(names.count("보고서 본문1"), 1)
            self.assertEqual(sum(name in REQUIRED_STYLE_NAMES for name in names), 5)
            body_style = next(node for node in list(second_styles) if node.attrib.get("name") == "보고서 본문1")
            char_prs = {node.attrib["id"]: node for node in child(second_ref_list, "charProperties")}
            body_char_pr = char_prs[body_style.attrib["charPrIDRef"]]
            self.assertEqual(body_char_pr.attrib["height"], "1200")
            self.assertTrue(any(local_name(node.tag) == "bold" for node in list(body_char_pr)))

    def test_rejects_missing_or_invalid_header(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            missing = temp / "missing.hwpx"
            invalid = temp / "invalid.hwpx"
            output = temp / "output.hwpx"
            write_fixture(missing, None)
            write_fixture(invalid, b"not xml")

            with self.assertRaisesRegex(ValueError, r"Contents/header\.xml.*없"):
                register_named_styles(missing, output, DEFAULT_HWP_STYLE_CONFIG)
            with self.assertRaisesRegex(ValueError, r"Contents/header\.xml.*XML"):
                register_named_styles(invalid, output, DEFAULT_HWP_STYLE_CONFIG)


if __name__ == "__main__":
    unittest.main()
