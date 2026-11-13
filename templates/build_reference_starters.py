"""Build clean HWPX starters from the project's own editable basic template."""

from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parent
HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
VARIANTS = {
    "guardian": {
        "labels": ("조사 결과 보고서", "01  조사 결과", "02  검토 요약"),
        "top": "5102", "side": "4252", "bottom": "4252", "header": "1984", "footer": "1417",
        "accent": "#455C6D", "cover_blank_lines": 3, "title_size": "3000",
    },
    "screening": {
        "labels": ("조사 연구 보고서", "제 1 장  주요 조사 결과", "제 2 장  검토 요약"),
        "top": "3969", "side": "5102", "bottom": "3685", "header": "3969", "footer": "3402",
        "accent": "#22669E", "cover_blank_lines": 5, "title_size": "3200",
    },
}


def read_xml(payload):
    for _, (prefix, uri) in ET.iterparse(BytesIO(payload), events=("start-ns",)):
        ET.register_namespace(prefix, uri)
    return ET.fromstring(payload)


def build(variant, target):
    with ZipFile(ROOT / "report_template_basic.hwpx") as source:
        section = read_xml(source.read("Contents/section0.xml"))
        header = read_xml(source.read("Contents/header.xml"))
        paragraphs = section.findall(f"{{{HP}}}p")
        label, body, qa = variant["labels"]
        replacements = {7: label, 12: body, 14: qa}
        for idx, text in replacements.items():
            paragraphs[idx].find(".//{*}t").text = text
        for idx in range(1 + variant["cover_blank_lines"], 7):
            section.remove(paragraphs[idx])
        for idx in range(7, 11):
            paragraphs[idx].set("paraPrIDRef", paragraphs[12].get("paraPrIDRef"))
        for paragraph in paragraphs:
            for child in paragraph.findall(f"{{{HP}}}linesegarray"):
                paragraph.remove(child)
        margin = section.find(".//{*}pagePr/{*}margin")
        margin.attrib.update(left=variant["side"], right=variant["side"],
                             **{key: variant[key] for key in ("top", "bottom", "header", "footer")})
        for cp in header.findall(".//{*}charPr"):
            if cp.get("id") in {"8", "10"}:
                cp.set("textColor", variant["accent"])
            elif cp.get("id") == "7":
                cp.set("height", variant["title_size"])
                cp.set("textColor", "#20292F")
        preview = "\r\n".join(t.text for t in section.findall(".//{*}t") if t.text) + "\r\n"
        changes = {"Contents/section0.xml": ET.tostring(section, encoding="utf-8", xml_declaration=True),
                   "Contents/header.xml": ET.tostring(header, encoding="utf-8", xml_declaration=True),
                   "Preview/PrvText.txt": preview.encode("utf-8")}
        with ZipFile(target, "w") as output:
            for entry in source.infolist():
                output.writestr(entry, changes[entry.filename] if entry.filename in changes else source.read(entry))


if __name__ == "__main__":
    for name, settings in VARIANTS.items():
        build(settings, ROOT / f"report_template_{name}.hwpx")
