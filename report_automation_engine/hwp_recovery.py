"""Durable HWPX snapshots; a progress log alone is not a recovery point."""

import hashlib
import json
from pathlib import Path
import uuid
import zipfile


COUNTERS = ("sections_written", "source_tables_written", "tables_written", "charts_deferred",
            "table_results", "table_layouts", "warnings", "style_application_counts", "placeholders")


def file_digest(path):
    with Path(path).open("rb") as source:
        return hashlib.file_digest(source, "sha256").hexdigest()


class Recovery:
    def __init__(self, path, sources, max_sections):
        self.path = Path(path)
        self.inputs = [file_digest(source) if source else None for source in sources] + [max_sections]
        self.state = None
        if self.path.exists():
            state = json.loads(self.path.read_text(encoding="utf-8"))
            if not isinstance(state, dict):
                raise ValueError("자동 재개 기록 형식이 올바르지 않습니다.")
            if state.get("inputs") != self.inputs:
                raise ValueError("자동 재개 입력 또는 설정이 이전 실행과 다릅니다.")
            if state.get("slot") not in (0, 1) or state.get("sha256") != file_digest(self.snapshot(state["slot"])):
                raise ValueError("자동 재개용 한글 사본이 손상되었습니다.")
            if (not isinstance(state.get("marker"), str) or not state["marker"].startswith("RARESUME")
                    or not isinstance(state.get("report"), dict)):
                raise ValueError("자동 재개 위치가 올바르지 않습니다.")
            self.state = state
        self.marker = self.state["marker"] if self.state else "RARESUME" + uuid.uuid4().hex.upper()

    def snapshot(self, slot):
        return self.path.with_name(f"{self.path.stem}_{slot}.hwpx")

    def next_snapshot(self):
        return self.snapshot(self.next_slot)

    @property
    def next_slot(self):
        return 1 - self.state["slot"] if self.state else 0

    def commit(self, report, write_json):
        slot = self.next_slot
        document = self.snapshot(slot)
        with zipfile.ZipFile(document) as archive:
            if not any(self.marker.encode("ascii") in archive.read(name)
                       for name in archive.namelist() if name.startswith("Contents/section") and name.endswith(".xml")):
                raise ValueError("자동 재개 위치가 저장된 문서에 없습니다.")
        state = {"inputs": self.inputs, "slot": slot, "sha256": file_digest(document), "marker": self.marker,
                 "report": {key: report[key] for key in COUNTERS if key in report}}
        write_json(self.path, state)
        self.state = state

    def clear(self):
        self.path.unlink(missing_ok=True)
        for slot in (0, 1):
            self.snapshot(slot).unlink(missing_ok=True)
