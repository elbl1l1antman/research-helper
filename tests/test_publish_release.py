import importlib.util
import json
from pathlib import Path
import subprocess

import pytest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("publish_release", ROOT / "report_automation_launcher" / "scripts" / "publish_release.py")
release = importlib.util.module_from_spec(spec)
spec.loader.exec_module(release)


@pytest.fixture
def repo(tmp_path):
    (tmp_path / "VERSION").write_text("0.1.0\n", encoding="utf-8")
    (tmp_path / "VERSIONING.md").write_text("`0.1.0` adds the first workflow.\n", encoding="utf-8")
    exe = tmp_path / "bin" / "ReportAutomationLauncher.exe"
    exe.parent.mkdir()
    exe.write_bytes(b"MZ" + b"\0" * 64)
    return tmp_path, exe


def test_publish_creates_an_immutable_folder_with_only_release_files(repo):
    root, exe = repo
    target = release.publish(root, exe)
    assert target == root / "releases" / "v0.1.0"
    assert {p.name for p in target.iterdir()} == {"ReportAutomationLauncher.exe", "SHA256SUMS.txt"}
    digest = release.sha256(exe)
    assert (target / "SHA256SUMS.txt").read_text(encoding="utf-8") == f"{digest} *ReportAutomationLauncher.exe\n"
    index = json.loads((root / "releases" / "index.json").read_text(encoding="utf-8"))
    assert index["latest"] == "0.1.0" and index["releases"][0]["notes"] == "adds the first workflow."
    assert "`0.1.0`" in (root / "releases" / "README.md").read_text(encoding="utf-8")
    assert release.verify(root) == []
    with pytest.raises(release.ReleaseError, match="이미 배포"):
        release.publish(root, exe)


def test_unchanged_binary_needs_explicit_permission(repo):
    root, exe = repo
    release.publish(root, exe)
    (root / "VERSION").write_text("0.1.1", encoding="utf-8")
    with pytest.raises(release.ReleaseError, match="0.1.0 배포본과 같습니다"):
        release.publish(root, exe)
    assert not (root / "releases" / "v0.1.1").exists()
    release.publish(root, exe, allow_unchanged=True, notes="재배포")
    index = json.loads((root / "releases" / "index.json").read_text(encoding="utf-8"))
    assert [item["version"] for item in index["releases"]] == ["0.1.1", "0.1.0"]
    assert index["latest"] == "0.1.1" and index["releases"][0]["notes"] == "재배포"


@pytest.mark.parametrize("content, message", [(None, "빌드하세요"), (b"not an exe", "Windows 실행 파일")])
def test_missing_or_invalid_exe_is_refused(repo, content, message):
    root, exe = repo
    if content is None:
        exe.unlink()
    else:
        exe.write_bytes(content)
    with pytest.raises(release.ReleaseError, match=message):
        release.publish(root, exe)
    assert not (root / "releases" / "v0.1.0").exists()


def test_verify_reports_tampering_and_stray_files(repo):
    root, exe = repo
    target = release.publish(root, exe)
    (target / "ReportAutomationLauncher.exe").write_bytes(b"MZ changed")
    (root / "releases" / "notes.txt").write_text("x", encoding="utf-8")
    problems = release.verify(root)
    assert any("해시" in p for p in problems)
    assert any("notes.txt" in p for p in problems)
    (target / "ReportAutomationLauncher.exe").write_bytes(exe.read_bytes())
    (target / "debug.log").write_text("x", encoding="utf-8")
    assert any("만 있어야" in p for p in release.verify(root))


def test_source_commit_is_the_commit_that_holds_the_exe(repo):
    root, exe = repo
    run = lambda *args: subprocess.run(["git", *args], cwd=root, check=True, capture_output=True)
    run("init", "-q")
    run("-c", "user.name=t", "-c", "user.email=t@example.invalid", "add", ".")
    run("-c", "user.name=t", "-c", "user.email=t@example.invalid", "commit", "-qm", "build")
    head = subprocess.run(["git", "rev-parse", "HEAD"], cwd=root, capture_output=True, text=True).stdout.strip()
    source = release.build_source(root, exe)
    assert source == {**source, "source_commit": head, "source_dirty": False}


def test_repository_release_folders_match_their_index():
    assert release.verify(ROOT) == []
