"""Hancom COM process lifecycle helpers for back-to-back writer runs.

아래한글은 ``Quit()``가 반환된 뒤에도 잠시 동안 프로세스와 COM 등록(ROT)이
남는다. 같은 Python 프로세스에서 바로 다음 문서를 만들면 새 Dispatch가 종료 중인
인스턴스에 붙거나, 첫 ``Open``이 RPC 오류로 실패할 수 있다. 이 모듈은

1. 작성기가 새로 띄운 Hwp.exe PID를 기록하고,
2. 종료한 PID가 실제로 사라질 때까지 기다리며,
3. 연결 직후 상태 확인과 일시적 COM 오류 재시도를 제공한다.

사용자가 이미 열어 둔 한글 프로세스는 기록하지 않으므로 강제로 종료하지 않는다.
"""

from __future__ import annotations

import csv
import gc
import io
import json
import os
import platform
import subprocess
import time
from typing import Any, Callable, Dict, Iterable, Set

HWP_IMAGE_NAME = "Hwp.exe"
EXIT_WAIT_SECONDS = 15.0
POLL_SECONDS = 0.25
RETRY_DELAYS = (0.5, 1.0, 2.0)

# HRESULTs that mean "the server is busy, starting, or shutting down; try again".
TRANSIENT_HRESULTS = {
    -2147418111,  # RPC_E_CALL_REJECTED
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
    -2147417848,  # RPC_E_DISCONNECTED
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE (0x800706BA)
    -2147023170,  # RPC_S_CALL_FAILED (0x800706BE)
    -2146959355,  # CO_E_SERVER_EXEC_FAILURE
    -2147220995,  # CO_E_OBJNOTCONNECTED
    -2147467259,  # E_FAIL, returned by Hancom while a previous instance is closing
}

# PIDs this Python process asked to quit and that may still be shutting down.
_PENDING_EXIT: Set[int] = set()
# Launcher runs start a new Python process per document, so slow exits are also
# recorded on disk for a short time. Old entries are ignored to avoid PID reuse.
PENDING_FILE_MAX_AGE_SECONDS = 120.0


def pending_file() -> str | None:
    root = os.environ.get("LOCALAPPDATA")
    if not root:
        return None
    return os.path.join(root, "ResearchHelper", "hwp_pending_exit.json")


def _read_pending_file(path: str | None, wall_clock: Callable[[], float] = time.time) -> Set[int]:
    if not path:
        return set()
    try:
        with open(path, encoding="utf-8") as stream:
            data = json.load(stream)
        if wall_clock() - float(data.get("recorded_at", 0)) > PENDING_FILE_MAX_AGE_SECONDS:
            return set()
        return {int(pid) for pid in data.get("process_ids", [])}
    except Exception:
        return set()


def _write_pending_file(path: str | None, pids: Set[int], wall_clock: Callable[[], float] = time.time) -> None:
    if not path:
        return
    try:
        if not pids:
            if os.path.exists(path):
                os.remove(path)
            return
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as stream:
            json.dump({"recorded_at": wall_clock(), "process_ids": sorted(pids)}, stream)
    except Exception:
        pass


def _windows() -> bool:
    return platform.system() == "Windows"


def list_hwp_processes(runner: Callable[..., Any] | None = None) -> Set[int]:
    """Return running Hwp.exe PIDs, or an empty set when they cannot be listed."""
    if runner is None:
        if not _windows():
            return set()
        runner = subprocess.run
    try:
        completed = runner(
            ["tasklist", "/FI", f"IMAGENAME eq {HWP_IMAGE_NAME}", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return set()
    if getattr(completed, "returncode", 1) != 0:
        return set()
    pids = set()
    for row in csv.reader(io.StringIO(getattr(completed, "stdout", "") or "")):
        if len(row) >= 2 and row[0].strip().lower() == HWP_IMAGE_NAME.lower():
            try:
                pids.add(int(row[1]))
            except ValueError:
                continue
    return pids


def com_hresults(exc: BaseException) -> Set[int]:
    """Read HRESULTs from pywintypes.com_error: the call result and excepinfo scode."""
    codes = set()
    args = getattr(exc, "args", ()) or ()
    for value in (getattr(exc, "hresult", None), args[0] if args else None):
        if isinstance(value, int) and not isinstance(value, bool):
            codes.add(value)
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 6 and isinstance(args[2][5], int):
        codes.add(args[2][5])
    # Normalize unsigned values (0x800706BA) to the signed form pywin32 reports.
    return {code - (1 << 32) if code > 0x7FFFFFFF else code for code in codes}


def is_transient_com_error(exc: BaseException) -> bool:
    if com_hresults(exc) & TRANSIENT_HRESULTS:
        return True
    text = str(exc).lower()
    return any(token in text for token in ("rpc 서버", "rpc server", "call was rejected", "호출이 거부", "disconnected"))


def retry_com(
    call: Callable[[], Any],
    *,
    delays: Iterable[float] = RETRY_DELAYS,
    sleep: Callable[[float], None] | None = None,
    on_retry: Callable[[int, BaseException], None] | None = None,
) -> Any:
    """Run ``call`` and retry only transient COM failures with backoff."""
    sleep = sleep or time.sleep
    delays = list(delays)
    for attempt in range(len(delays) + 1):
        try:
            return call()
        except Exception as exc:
            if attempt >= len(delays) or not is_transient_com_error(exc):
                raise
            if on_retry is not None:
                on_retry(attempt + 1, exc)
            sleep(delays[attempt])
    raise AssertionError("unreachable")


def wait_for_exit(
    pids: Iterable[int],
    timeout: float = EXIT_WAIT_SECONDS,
    *,
    lister: Callable[[], Set[int]] | None = None,
    clock: Callable[[], float] | None = None,
    sleep: Callable[[float], None] | None = None,
    poll: float = POLL_SECONDS,
) -> Set[int]:
    """Wait until the given PIDs disappear. Return the PIDs still running."""
    lister, clock, sleep = lister or list_hwp_processes, clock or time.monotonic, sleep or time.sleep
    remaining = set(pids)
    if not remaining:
        return set()
    deadline = clock() + timeout
    while True:
        remaining &= lister()
        if not remaining or clock() >= deadline:
            return remaining
        sleep(poll)


def wait_for_previous_instances(report: Dict[str, Any], **kwargs: Any) -> Set[int]:
    """Before a new Dispatch, let instances an earlier run quit finish exiting."""
    path = kwargs.pop("pending_path", pending_file())
    pending = set(_PENDING_EXIT) | _read_pending_file(path)
    if not pending:
        return set()
    started = time.monotonic()
    remaining = wait_for_exit(pending, **kwargs)
    _PENDING_EXIT.difference_update(pending - remaining)
    _write_pending_file(path, remaining)
    report.setdefault("com", {})["previous_exit_wait"] = {
        "waited_for": sorted(pending),
        "still_running": sorted(remaining),
        "seconds": round(time.monotonic() - started, 2),
    }
    if remaining:
        report.setdefault("warnings", []).append(
            "이전 작성에서 종료한 한글 프로세스가 아직 남아 있습니다(PID "
            + ", ".join(map(str, sorted(remaining)))
            + "). 새 연결이 실패하면 작업 관리자에서 해당 Hwp.exe를 종료한 뒤 다시 실행하세요."
        )
    return remaining


def remember_new_processes(report: Dict[str, Any], before: Set[int], after: Set[int]) -> Set[int]:
    owned = set(after) - set(before)
    report.setdefault("com", {})["process_ids"] = sorted(owned)
    report["com"]["attached_existing_instance"] = bool(after) and not owned and bool(before)
    return owned


def probe_hwp(hwp) -> None:
    """Touch a cheap property so a dying or half-started server fails here, not in Open."""
    documents = getattr(hwp, "XHwpDocuments", None)
    if documents is not None:
        # Reading the property is the RPC round trip; the value itself is irrelevant.
        documents.Count
        return
    getattr(hwp, "Version", None)


def wait_after_quit(report: Dict[str, Any], **kwargs: Any) -> Set[int]:
    """After Quit, wait for the processes this writer started and mark leftovers."""
    path = kwargs.pop("pending_path", pending_file())
    owned = set(report.get("com", {}).get("process_ids") or [])
    if not owned:
        return set()
    started = time.monotonic()
    remaining = wait_for_exit(owned, **kwargs)
    _PENDING_EXIT.update(remaining)
    if remaining:
        _write_pending_file(path, set(_PENDING_EXIT))
    report["com"]["exit_wait"] = {
        "process_ids": sorted(owned),
        "still_running": sorted(remaining),
        "seconds": round(time.monotonic() - started, 2),
    }
    if remaining:
        report.setdefault("warnings", []).append(
            "한글 종료 요청 후 프로세스가 제한 시간 안에 끝나지 않았습니다. 다음 작성 전에 다시 기다립니다."
        )
    return remaining


def release_com_references() -> None:
    """Drop Python-side COM proxies so the server's reference count can reach zero."""
    gc.collect()
    try:
        import pythoncom  # type: ignore
    except Exception:
        return
    try:
        pythoncom.CoFreeUnusedLibraries()
    except Exception:
        pass


def pending_exit_pids() -> Set[int]:
    return set(_PENDING_EXIT)


def clear_pending_exit() -> None:
    _PENDING_EXIT.clear()
