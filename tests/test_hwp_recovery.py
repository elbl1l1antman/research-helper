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
