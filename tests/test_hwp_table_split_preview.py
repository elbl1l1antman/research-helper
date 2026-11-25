import json
from pathlib import Path
import sys
import zipfile

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from report_automation_engine.hwp_table_layout_options import normalize_table_layout
from report_automation_engine.hwp_table_split_preview import (
    build_preview,
    page_from_hwpx,
    render_text,
    resolve_page,
    run_cli,
)
from test_hwp_table_layout_policies import merge, wide_table

TEMPLATE = ROOT / "templates" / "report_template_basic.hwpx"
A4_SETUP = {"paper_width_mm": 210, "paper_height_mm": 297, "orientation": "portrait", "top_mm": 20, "bottom_mm": 15,
            "left_mm": 25, "right_mm": 25, "gutter_mm": 0, "header_mm": 10, "footer_mm": 10}


def package(*tables):
    return {"meta": {}, "sections": [{"table_key": t["table_key"], "title": f"{t['table_key']} 문항"} for t in tables],
            "tables": list(tables), "charts": [], "qa": []}


def keyed(table, key):
    table["table_key"] = key
    return table


def test_template_body_width_is_read_from_the_body_section():
    page = page_from_hwpx(TEMPLATE)
    assert (page["width"], page["left"], page["right"], page["landscape"]) == (59528, 7087, 7087, False)


def test_body_section_is_selected_when_a_cover_section_comes_first(tmp_path):
    cover = '<hp:pagePr landscape="NARROWLY" width="59528" height="84186" gutterType="LEFT_ONLY"><hp:margin left="1000" right="1000" gutter="0"/></hp:pagePr>'
    body = '<hp:pagePr landscape="WIDELY" width="59528" height="84186" gutterType="LEFT_ONLY"><hp:margin left="7000" right="7000" gutter="500"/></hp:pagePr>'
    path = tmp_path / "two.hwpx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("Contents/section0.xml", f"<root>{cover}<p>표지</p>{body}<p>{{{{BODY}}}}</p></root>")
    page = page_from_hwpx(path)
    assert (page["left"], page["gutter"], page["landscape"]) == (7000, 500, False)


def test_page_setup_wins_over_the_template():
    page = resolve_page(TEMPLATE, A4_SETUP)
    assert page["source"] == "page_setup"
    assert page["left"] == round(25 * 7200 / 25.4)


def test_preview_lists_parts_columns_and_allowed_cuts():
    preview = build_preview(package(keyed(wide_table(20), "Q1"), keyed(wide_table(4), "Q2")), template=TEMPLATE)
    assert preview["page"]["body_width_mm"] == 160.0
    first, second = preview["tables"]
    assert first["status"] == "ok" and len(first["parts"]) > 1
    assert first["repeated_columns"] == [1]
    assert first["allowed_cuts"] == list(range(3, 21))
    assert first["current_cuts"] and all(cut in first["allowed_cuts"] for cut in first["current_cuts"])
    assert [c["col"] for c in first["columns"]] == list(range(1, 21))
    assert first["columns"][1]["label"] == "보기2"
    assert all(column["part"] for column in first["columns"][1:])
    assert len(second["parts"]) == 1
    assert preview["summary"] == {"tables": 2, "split": 1, "landscape": 0, "blocked": 0}
    assert "Q1" in render_text(preview)


def test_blocked_table_still_reports_columns_and_blocked_boundaries():
    table = keyed(merge(wide_table(20), 2, 2, colspan=19), "M")
    preview = build_preview(package(table), template=TEMPLATE)
    item = preview["tables"][0]
    assert item["status"] == "blocked" and item["error"]
    assert item["allowed_cuts"] == []
    assert [b["col"] for b in item["blocked_cuts"]] == list(range(3, 21))
    assert len(item["columns"]) == 20


def test_saved_overrides_change_the_preview():
    table = keyed(merge(wide_table(20), 2, 2, colspan=19), "M")
    layout = normalize_table_layout({"schema_version": "1.0", "tables": {
        "M": {"body_merge_policy": "split_repeat", "cuts": [8, 14], "column_widths_mm": {"1": 30}}}})
    item = build_preview(package(table), template=TEMPLATE, table_layout=layout)["tables"][0]
    assert item["status"] == "ok"
    assert item["current_cuts"] == [8, 14]
    assert item["columns"][0]["fixed"] is True
    assert item["columns"][0]["min_width_mm"] == 30.0
    assert any("반복 표시" in w for w in item["warnings"])


def test_landscape_policy_is_previewed():
    table = keyed(merge(wide_table(16), 2, 2, colspan=15), "W")
    layout = normalize_table_layout({"schema_version": "1.0", "defaults": {"wide_table_policy": "landscape_if_needed"}})
    item = build_preview(package(table), template=TEMPLATE, table_layout=layout)["tables"][0]
    assert item["orientation"] == "landscape"
    assert item["planned_body_width_mm"] == item["rotated_body_width_mm"]


def test_table_without_matrix_is_skipped():
    preview = build_preview(package({"table_key": "old", "rows": []}), template=TEMPLATE)
    assert preview["tables"][0]["status"] == "skipped"


def test_cli_writes_json_and_reports_bad_input(tmp_path, capsys):
    package_path = tmp_path / "package.json"
    package_path.write_text(json.dumps(package(keyed(wide_table(20), "Q1"))), encoding="utf-8")
    output = tmp_path / "preview.json"
    assert run_cli(["--package", str(package_path), "--template", str(TEMPLATE), "--output", str(output), "--table-key", "Q1"]) == 0
    assert json.loads(output.read_text(encoding="utf-8"))["tables"][0]["table_key"] == "Q1"

    layout = tmp_path / "layout.json"
    layout.write_text('{"schema_version": "9"}', encoding="utf-8")
    assert run_cli(["--package", str(package_path), "--table-layout", str(layout)]) == 2
    assert "schema_version" in capsys.readouterr().err


def test_one_broken_table_does_not_hide_the_others():
    layout = normalize_table_layout({"schema_version": "1.0", "tables": {"BAD": {"column_widths_mm": {"9": 20}}}})
    preview = build_preview(package(keyed(wide_table(5), "BAD"), keyed({"matrix": [[]]}, "EMPTY"), keyed(wide_table(4), "OK")),
                            template=TEMPLATE, table_layout=layout)
    status = {t["table_key"]: t["status"] for t in preview["tables"]}
    assert status == {"BAD": "blocked", "EMPTY": "blocked", "OK": "ok"}


def test_preview_minimum_widths_follow_display_formatting():
    from report_automation_engine.hwp_style_config import load_hwp_style_config

    table = keyed(wide_table(4), "F")
    for row in table["matrix"][1:]:
        for item in row[1:]:
            item["display_text"] = "1234.5678"
    style = load_hwp_style_config(None)
    style["table_output"]["value_decimals"] = 0
    narrow = build_preview(package(table), template=TEMPLATE, style_config=style)["tables"][0]
    wide = build_preview(package(table), template=TEMPLATE)["tables"][0]
    assert narrow["columns"][1]["min_width_mm"] <= wide["columns"][1]["min_width_mm"]
