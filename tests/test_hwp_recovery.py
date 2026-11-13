import json
import zipfile

import pytest

from report_automation_engine.hwp_recovery import Recovery
from report_automation_engine.hwp_com_writer import write_json


def test_recovery_keeps_last_committed_document(tmp_path):
    source = tmp_path / "input.json"
    source.write_text("one")
    manifest = tmp_path / "resume.json"
    job = Recovery(manifest, (source,), 2)
    with zipfile.ZipFile(job.next_snapshot(), "w") as archive:
        archive.writestr("Contents/section0.xml", f"<t>{job.marker}</t>")
    job.commit({"sections_written": 1, "tables_written": 3}, write_json)
    job.next_snapshot().write_bytes(b"interrupted save")
    restored = Recovery(manifest, (source,), 2)
    assert restored.state["report"] == {"sections_written": 1, "tables_written": 3}
    assert restored.marker == job.marker
    assert restored.next_slot == 1
    with pytest.raises(ValueError):
        Recovery(manifest, (source,), 3)
    restored.snapshot(0).write_bytes(b"corrupted snapshot")
    with pytest.raises(ValueError):
        Recovery(manifest, (source,), 2)
    assert json.loads(manifest.read_text())["slot"] == 0


def test_writer_reports_stale_recovery_as_verify_failure(tmp_path):
    from unittest.mock import patch
    from report_automation_engine.hwp_com_writer import HwpWriterError, write_hwp_document

    package = tmp_path / "package.json"
    preflight = tmp_path / "preflight.json"
    template = tmp_path / "template.hwpx"
    output = tmp_path / "output.hwpx"
    report = tmp_path / "writer_report.json"
    package.write_text('{"meta": {}, "sections": [], "tables": [], "charts": [], "qa": []}', encoding="utf-8")
    preflight.write_text('{"status": "ready"}', encoding="utf-8")
    template.write_bytes(b"template")
    manifest = tmp_path / ".job.recovery.json"
    manifest.write_text('["not", "a", "manifest"]', encoding="utf-8")
    with (
        patch("report_automation_engine.hwp_com_writer.platform.system", return_value="Windows"),
        patch("report_automation_engine.hwp_com_writer.acquire_writer_lock", return_value=None),
        patch("report_automation_engine.hwp_com_writer.create_hwp_object") as create_hwp,
    ):
        with pytest.raises(HwpWriterError) as raised:
            write_hwp_document(package, preflight, template, output, report_path=report, recovery_path=manifest)
    assert (raised.value.stage, raised.value.action) == ("recovery", "verify")
    saved = json.loads(report.read_text(encoding="utf-8"))
    assert saved["status"] == "failed" and saved["stage"] == "recovery"
    create_hwp.assert_not_called()
    assert not output.exists()
