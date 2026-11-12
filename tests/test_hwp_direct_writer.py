import json
import zipfile
from pathlib import Path

from report_automation_engine.hwp_direct_writer import write_hwpx_direct


def make_template(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("mimetype", "application/hwp+zip")
        zf.writestr("Contents/section0.xml", "<?xml version='1.0' encoding='UTF-8'?><doc><p>{{BODY}}</p></doc>")


def make_package(path: Path) -> None:
    package = {
        "schema_version": "1.0",
        "meta": {"source_file_name": "sample.xlsx"},
        "sections": [{"table_key": "T001", "title": "만족도", "narrative_final": "만족도는 63.3%로 나타남."}],
        "tables": [
            {
                "table_key": "T001",
                "title": "만족도",
                "matrix": [
                    [{"display_text": "항목", "role": "header"}, {"display_text": "비율", "role": "header"}],
                    [{"display_text": "전체", "role": "stub"}, {"display_text": "63.3%", "role": "value", "raw_value": 63.25}],
                ],
            }
        ],
        "qa": [],
    }
    path.write_text(json.dumps(package, ensure_ascii=False), encoding="utf-8")


def test_write_hwpx_direct_replaces_body_and_uses_display_text(tmp_path):
    template = tmp_path / "template.hwpx"
    package = tmp_path / "package.json"
    preflight = tmp_path / "preflight.json"
    output = tmp_path / "output.hwpx"
    make_template(template)
    make_package(package)
    preflight.write_text(json.dumps({"status": "ready"}, ensure_ascii=False), encoding="utf-8")
    write_hwpx_direct(package, preflight, template, output)
    with zipfile.ZipFile(output) as zf:
        content = "\n".join(zf.read(name).decode("utf-8") for name in zf.namelist() if name.endswith(".xml"))
    assert "{{BODY}}" not in content
    assert "만족도는 63.3%로 나타남." in content
    assert "63.3%" in content
    assert "63.25" not in content
