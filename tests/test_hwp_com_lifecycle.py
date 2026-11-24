from pathlib import Path
from types import SimpleNamespace
import sys
from unittest.mock import patch

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from report_automation_engine import hwp_com_lifecycle as lifecycle
from report_automation_engine.hwp_com_writer import (
    HwpWriterError,
    open_document,
    settle_closed_hwp,
)


class FakeComError(Exception):
    """Mimic pywintypes.com_error(hresult, text, excepinfo, argerror)."""


RPC_UNAVAILABLE = FakeComError(-2147023174, "RPC 서버를 사용할 수 없습니다.", None, None)
DISP_WITH_REJECTED = FakeComError(-2147352567, "예외가 발생했습니다.", (0, "Hwp", "busy", None, 0, -2147418111), None)
PERMANENT = FakeComError(-2147024894, "파일을 찾을 수 없습니다.", None, None)


class FakeClock:
    def __init__(self):
        self.now = 0.0

    def __call__(self):
        return self.now

    def sleep(self, seconds):
        self.now += seconds


@pytest.fixture(autouse=True)
def clean_pending(monkeypatch):
    # Never touch the real %LOCALAPPDATA% record while testing on Windows.
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    lifecycle.clear_pending_exit()
    yield
    lifecycle.clear_pending_exit()


def test_tasklist_csv_is_parsed_and_other_images_are_ignored():
    output = '"Hwp.exe","4120","Console","1","120,000 K"\n"hwp.exe","88","Console","1","1 K"\n"Other.exe","9","C","1","1 K"\n'
    runner = lambda *_args, **_kwargs: SimpleNamespace(returncode=0, stdout=output)
    assert lifecycle.list_hwp_processes(runner) == {4120, 88}


def test_tasklist_failure_or_no_match_returns_empty_set():
    assert lifecycle.list_hwp_processes(lambda *a, **k: SimpleNamespace(returncode=1, stdout="")) == set()
    info = "INFO: No tasks are running which match the specified criteria.\n"
    assert lifecycle.list_hwp_processes(lambda *a, **k: SimpleNamespace(returncode=0, stdout=info)) == set()

    def broken(*_args, **_kwargs):
        raise OSError("tasklist missing")

    assert lifecycle.list_hwp_processes(broken) == set()


@pytest.mark.parametrize("error, transient", [
    (RPC_UNAVAILABLE, True),
    (DISP_WITH_REJECTED, True),
    (FakeComError(0x800706BA, "unsigned", None, None), True),
    (PERMANENT, False),
    (ValueError("plain error"), False),
])
def test_transient_com_errors_are_recognized(error, transient):
    assert lifecycle.is_transient_com_error(error) is transient


def test_retry_com_retries_only_transient_errors():
    calls, retries, sleeps = [], [], []

    def flaky():
        calls.append(1)
        if len(calls) < 3:
            raise RPC_UNAVAILABLE
        return "opened"

    assert lifecycle.retry_com(flaky, sleep=sleeps.append, on_retry=lambda n, e: retries.append(n)) == "opened"
    assert retries == [1, 2]
    assert sleeps == list(lifecycle.RETRY_DELAYS[:2])

    def permanent():
        raise PERMANENT

    with pytest.raises(FakeComError):
        lifecycle.retry_com(permanent, sleep=sleeps.append)
    assert len(sleeps) == 2


def test_retry_com_gives_up_after_all_delays():
    attempts = []

    def always_busy():
        attempts.append(1)
        raise RPC_UNAVAILABLE

    with pytest.raises(FakeComError):
        lifecycle.retry_com(always_busy, delays=(0.1, 0.1), sleep=lambda _s: None)
    assert len(attempts) == 3


def test_wait_for_exit_polls_until_processes_disappear():
    clock = FakeClock()
    snapshots = iter([{10, 11, 99}, {11, 99}, {99}])
    remaining = lifecycle.wait_for_exit({10, 11}, 5, lister=lambda: next(snapshots), clock=clock, sleep=clock.sleep, poll=0.5)
    assert remaining == set()
    assert clock.now == 1.0


def test_wait_for_exit_returns_leftovers_after_timeout():
    clock = FakeClock()
    remaining = lifecycle.wait_for_exit({10}, 1, lister=lambda: {10}, clock=clock, sleep=clock.sleep, poll=0.25)
    assert remaining == {10}
    assert clock.now >= 1


def test_only_new_processes_are_owned():
    report = {"com": {}}
    assert lifecycle.remember_new_processes(report, {5}, {5, 7}) == {7}
    assert report["com"]["process_ids"] == [7]
    assert report["com"]["attached_existing_instance"] is False
    report = {"com": {}}
    assert lifecycle.remember_new_processes(report, {5}, {5}) == set()
    assert report["com"]["attached_existing_instance"] is True


def test_settle_waits_for_owned_process_and_records_it():
    report = {"com": {"process_ids": [42], "closed": True}, "warnings": []}
    clock = FakeClock()
    snapshots = iter([{42}, set()])
    settle_closed_hwp(report, lister=lambda: next(snapshots), clock=clock, sleep=clock.sleep, timeout=5, poll=0.5)
    assert report["com"]["exit_wait"]["still_running"] == []
    assert lifecycle.pending_exit_pids() == set()
    assert report["warnings"] == []


def test_slow_exit_is_carried_to_the_next_connection():
    report = {"com": {"process_ids": [42], "closed": True}, "warnings": []}
    clock = FakeClock()
    settle_closed_hwp(report, lister=lambda: {42}, clock=clock, sleep=clock.sleep, timeout=1, poll=0.5)
    assert lifecycle.pending_exit_pids() == {42}
    assert report["warnings"]

    next_report = {"com": {}, "warnings": []}
    snapshots = iter([{42}, set()])
    remaining = lifecycle.wait_for_previous_instances(next_report, lister=lambda: next(snapshots), clock=clock, sleep=clock.sleep, poll=0.5)
    assert remaining == set()
    assert next_report["com"]["previous_exit_wait"]["waited_for"] == [42]
    assert lifecycle.pending_exit_pids() == set()


def test_settle_does_nothing_for_an_attached_existing_instance():
    report = {"com": {"process_ids": []}, "warnings": []}
    settle_closed_hwp(report, lister=lambda: pytest.fail("must not list processes"))
    assert "exit_wait" not in report["com"]


def test_open_document_retries_rpc_rejection_after_previous_quit(tmp_path):
    path = tmp_path / "template.hwpx"
    calls = []

    def open_(*args):
        calls.append(args)
        if len(calls) == 1:
            raise RPC_UNAVAILABLE
        return True

    report = {"com": {}, "warnings": []}
    with patch.object(lifecycle.time, "sleep"):
        open_document(SimpleNamespace(Open=open_), path, report)
    assert report["document_opened"] is True
    assert len(calls) == 2
    assert [step["status"] for step in report["com"]["steps"]] == ["started", "retry", "ready"]


def test_open_document_does_not_retry_a_permanent_error(tmp_path):
    calls = []

    def open_(*args):
        calls.append(args)
        raise PERMANENT

    with patch.object(lifecycle.time, "sleep") as sleep, pytest.raises(HwpWriterError):
        open_document(SimpleNamespace(Open=open_), tmp_path / "t.hwpx", {"com": {}, "warnings": []})
    sleep.assert_not_called()
    assert len(calls) == 2  # HWPX open, then plain Open fallback


def test_slow_exit_is_shared_with_the_next_writer_process(tmp_path):
    path = str(tmp_path / "pending.json")
    report = {"com": {"process_ids": [42], "closed": True}, "warnings": []}
    clock = FakeClock()
    lifecycle.wait_after_quit(report, lister=lambda: {42}, clock=clock, sleep=clock.sleep, timeout=0.5, poll=0.5, pending_path=path)
    lifecycle.clear_pending_exit()  # a new Python process starts with no memory

    snapshots = iter([{42}, set()])
    next_report = {"com": {}, "warnings": []}
    lifecycle.wait_for_previous_instances(next_report, lister=lambda: next(snapshots), clock=clock, sleep=clock.sleep, poll=0.5, pending_path=path)
    assert next_report["com"]["previous_exit_wait"]["waited_for"] == [42]
    assert not Path(path).exists()


def test_stale_pending_file_is_ignored(tmp_path):
    path = tmp_path / "pending.json"
    path.write_text('{"recorded_at": 0, "process_ids": [42]}', encoding="utf-8")
    report = {"com": {}, "warnings": []}
    assert lifecycle.wait_for_previous_instances(report, lister=lambda: pytest.fail("no wait"), pending_path=str(path)) == set()
