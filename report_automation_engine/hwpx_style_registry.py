from __future__ import annotations

import copy
import os
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict
from xml.etree import ElementTree as ET

from .hwp_style_config import REQUIRED_STYLE_NAMES, validate_hwp_style_config


HH = "http://www.hancom.co.kr/hwpml/2011/head"
HC = "http://www.hancom.co.kr/hwpml/2011/core"
HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
_LANG_ATTR = {
    "HANGUL": "hangul",
    "LATIN": "latin",
    "HANJA": "hanja",
    "JAPANESE": "japanese",
    "OTHER": "other",
    "SYMBOL": "symbol",
    "USER": "user",
}

ET.register_namespace("hh", HH)
ET.register_namespace("hc", HC)
ET.register_namespace("hp", HP)


def _tag(name: str) -> str:
    return f"{{{HH}}}{name}"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _child(parent: ET.Element, name: str) -> ET.Element | None:
    return next((node for node in list(parent) if _local_name(node.tag) == name), None)


def _required_child(parent: ET.Element, name: str) -> ET.Element:
    node = _child(parent, name)
    if node is None:
        raise ValueError(f"Contents/header.xml에 필수 요소 {name}이 없습니다.")
    return node


def _next_id(collection: ET.Element) -> int:
    ids = [int(node.attrib["id"]) for node in list(collection) if node.attrib.get("id", "").isdigit()]
    return max(ids, default=-1) + 1


def _set_counts(*collections: ET.Element) -> None:
    for collection in collections:
        collection.attrib["itemCnt"] = str(len(list(collection)))


def _ensure_fonts(fontfaces: ET.Element, families: list[str]) -> Dict[str, Dict[str, int]]:
    result: Dict[str, Dict[str, int]] = {family: {} for family in families}
    for fontface in list(fontfaces):
        language = fontface.attrib.get("lang", "")
        attribute = _LANG_ATTR.get(language)
        fonts = list(fontface)
        if not fonts:
            raise ValueError(f"Contents/header.xml fontface({language})에 기준 font가 없습니다.")
        for family in families:
            font = next((node for node in fonts if node.attrib.get("face") == family), None)
            if font is None:
                font = copy.deepcopy(fonts[0])
                font.attrib.update({"id": str(_next_id(fontface)), "face": family})
                fontface.append(font)
                fonts.append(font)
            if attribute is not None:
                result[family][attribute] = int(font.attrib["id"])
        fontface.attrib["fontCnt"] = str(len(list(fontface)))
    fontfaces.attrib["itemCnt"] = str(len(list(fontfaces)))
    return result


def _referenced_count(styles: ET.Element, attribute: str, value: str) -> int:
    return sum(node.attrib.get(attribute) == value for node in list(styles))


def _property_for_style(
    collection: ET.Element,
    styles: ET.Element,
    style: ET.Element | None,
    reference_attribute: str,
) -> ET.Element:
    by_id = {node.attrib.get("id"): node for node in list(collection)}
    reference = style.attrib.get(reference_attribute) if style is not None else None
    current = by_id.get(reference)
    if current is not None and _referenced_count(styles, reference_attribute, reference or "") == 1:
        return current
    base = current if current is not None else next(iter(list(collection)), None)
    if base is None:
        raise ValueError(f"Contents/header.xml에 복제할 {collection.tag.rsplit('}', 1)[-1]} 기준 속성이 없습니다.")
    cloned = copy.deepcopy(base)
    cloned.attrib["id"] = str(_next_id(collection))
    collection.append(cloned)
    return cloned


def _apply_char_style(char_pr: ET.Element, settings: Dict[str, Any], font_ids: Dict[str, int]) -> None:
    char_pr.attrib["height"] = str(round(settings["font_size_pt"] * 100))
    font_ref = _child(char_pr, "fontRef")
    if font_ref is None:
        font_ref = ET.Element(_tag("fontRef"))
        char_pr.insert(0, font_ref)
    font_ref.attrib.update({name: str(font_id) for name, font_id in font_ids.items()})

    for node in [node for node in list(char_pr) if _local_name(node.tag) == "bold"]:
        char_pr.remove(node)
    if settings["bold"]:
        bold = ET.Element(_tag("bold"))
        insert_at = next(
            (index for index, node in enumerate(list(char_pr)) if _local_name(node.tag) in {"underline", "strikeout", "outline", "shadow"}),
            len(list(char_pr)),
        )
        char_pr.insert(insert_at, bold)


def _hwpunit(mm: float) -> str:
    return str(round(mm * 7200 / 25.4))


def _apply_para_style(para_pr: ET.Element, settings: Dict[str, Any], bullet_id: int | None) -> None:
    align = next((node for node in para_pr.iter() if _local_name(node.tag) == "align"), None)
    if align is None:
        align = ET.Element(_tag("align"), {"vertical": "BASELINE"})
        para_pr.insert(0, align)
    align.attrib["horizontal"] = settings["alignment"].upper()

    headings = [node for node in para_pr.iter() if _local_name(node.tag) == "heading"]
    if not headings:
        heading = ET.Element(_tag("heading"))
        para_pr.insert(1, heading)
        headings = [heading]
    for heading in headings:
        heading.attrib.update(
            {"type": "BULLET" if bullet_id is not None else "NONE", "idRef": str(bullet_id or 0), "level": "0"}
        )

    margins = [node for node in para_pr.iter() if _local_name(node.tag) == "margin"]
    if not margins:
        margins = [ET.SubElement(para_pr, _tag("margin"))]
    for margin in margins:
        for name, value in (
            ("intent", settings["first_line_indent_mm"]),
            ("left", settings["left_indent_mm"]),
        ):
            field = _child(margin, name)
            if field is None:
                field = ET.SubElement(margin, f"{{{HC}}}{name}")
            field.attrib.update({"value": _hwpunit(value), "unit": "HWPUNIT"})

    line_spacings = [node for node in para_pr.iter() if _local_name(node.tag) == "lineSpacing"]
    if not line_spacings:
        line_spacings = [ET.SubElement(para_pr, _tag("lineSpacing"))]
    for line_spacing in line_spacings:
        line_spacing.attrib.update(
            {"type": "PERCENT", "value": str(round(settings["line_spacing_percent"])), "unit": "HWPUNIT"}
        )


def _ensure_bullets(ref_list: ET.Element, para_properties: ET.Element) -> ET.Element:
    bullets = _child(ref_list, "bullets")
    if bullets is None:
        bullets = ET.Element(_tag("bullets"), {"itemCnt": "0"})
        ref_list.insert(list(ref_list).index(para_properties), bullets)
    return bullets


def _bullet_for_para(bullets: ET.Element, para_pr: ET.Element, bullet_char: str) -> int:
    headings = [node for node in para_pr.iter() if _local_name(node.tag) == "heading"]
    existing_id = next((node.attrib.get("idRef") for node in headings if node.attrib.get("type") == "BULLET"), None)
    bullet = next((node for node in list(bullets) if node.attrib.get("id") == existing_id), None)
    if bullet is None:
        bullet_id = max(_next_id(bullets), 1)
        bullet = ET.SubElement(bullets, _tag("bullet"), {"id": str(bullet_id)})
        ET.SubElement(
            bullet,
            _tag("paraHead"),
            {
                "level": "0",
                "align": "LEFT",
                "useInstWidth": "1",
                "autoIndent": "1",
                "widthAdjust": "0",
                "textOffsetType": "PERCENT",
                "textOffset": "50",
                "numFormat": "DIGIT",
                "charPrIDRef": "4294967295",
                "checkable": "0",
            },
        )
    bullet.attrib.update({"char": bullet_char, "useImage": "0"})
    for para_head in (node for node in list(bullet) if _local_name(node.tag) == "paraHead"):
        para_head.attrib["numFormat"] = "DIGIT"
    bullets.attrib["itemCnt"] = str(len(list(bullets)))
    return int(bullet.attrib["id"])


def _update_header(header: bytes, config: Dict[str, Any]) -> tuple[bytes, Dict[str, int]]:
    try:
        root = ET.fromstring(header)
    except ET.ParseError as exc:
        raise ValueError(f"Contents/header.xml XML을 읽을 수 없습니다: {exc}") from exc
    if root.tag != _tag("head"):
        namespace = root.tag[1:].split("}", 1)[0] if root.tag.startswith("{") else "없음"
        raise ValueError(f"Contents/header.xml 루트의 HWPX head 네임스페이스가 올바르지 않습니다: {namespace}")

    ref_list = _required_child(root, "refList")
    fontfaces = _required_child(ref_list, "fontfaces")
    char_properties = _required_child(ref_list, "charProperties")
    para_properties = _required_child(ref_list, "paraProperties")
    styles = _required_child(ref_list, "styles")
    bullets = _ensure_bullets(ref_list, para_properties)
    style_config = config["paragraph_styles"]
    families = list(dict.fromkeys(style_config[name]["font_family"] for name in REQUIRED_STYLE_NAMES))
    font_ids = _ensure_fonts(fontfaces, families)

    seen_names = set()
    for style in list(styles):
        name = style.attrib.get("name")
        if name not in REQUIRED_STYLE_NAMES:
            continue
        if name in seen_names:
            styles.remove(style)
        else:
            seen_names.add(name)

    for name in REQUIRED_STYLE_NAMES:
        settings = style_config[name]
        style = next((node for node in list(styles) if node.attrib.get("name") == name), None)
        char_pr = _property_for_style(char_properties, styles, style, "charPrIDRef")
        para_pr = _property_for_style(para_properties, styles, style, "paraPrIDRef")
        bullet_id = _bullet_for_para(bullets, para_pr, settings["bullet"]) if settings["bullet"] else None
        _apply_char_style(char_pr, settings, font_ids[settings["font_family"]])
        _apply_para_style(para_pr, settings, bullet_id)
        if style is None:
            style_id = _next_id(styles)
            style = ET.SubElement(styles, _tag("style"), {"id": str(style_id)})
        else:
            style_id = int(style.attrib.get("id", "0"))
        style.attrib.update(
            {
                "type": "PARA",
                "name": name,
                "engName": name,
                "paraPrIDRef": para_pr.attrib["id"],
                "charPrIDRef": char_pr.attrib["id"],
                "nextStyleIDRef": str(style_id),
                "langID": "1042",
                "lockForm": "0",
            }
        )

    _set_counts(char_properties, para_properties, bullets, styles)
    mapping = {
        node.attrib["name"]: index
        for index, node in enumerate(list(styles))
        if node.attrib.get("name") in REQUIRED_STYLE_NAMES
    }
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), mapping


def register_named_styles(template: Path, output: Path, config: Dict[str, Any]) -> Dict[str, int]:
    """Write a styled HWPX copy and return each fixed style's document index."""
    template = Path(template).resolve()
    output = Path(output).resolve()
    if template == output:
        raise ValueError("원본 HWPX와 출력 HWPX 경로는 달라야 합니다.")
    if not zipfile.is_zipfile(template):
        raise ValueError(f"HWPX ZIP 패키지가 아닙니다: {template}")
    normalized = validate_hwp_style_config(config)

    with zipfile.ZipFile(template, "r") as source:
        infos = source.infolist()
        if "Contents/header.xml" not in source.namelist():
            raise ValueError("HWPX에 Contents/header.xml이 없습니다.")
        payloads = {info.filename: source.read(info.filename) for info in infos}
    payloads["Contents/header.xml"], mapping = _update_header(payloads["Contents/header.xml"], normalized)

    output.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary_name = tempfile.mkstemp(suffix=".hwpx", dir=output.parent)
    os.close(handle)
    temporary = Path(temporary_name)
    try:
        with zipfile.ZipFile(temporary, "w") as target:
            for info in infos:
                target.writestr(info, payloads[info.filename])
        os.replace(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)
    return mapping
