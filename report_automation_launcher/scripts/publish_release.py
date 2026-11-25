"""Copy a built launcher into ``releases/v<VERSION>/`` and keep the release index.

Each release folder holds only what a user installs and runs:

    releases/v0.0.57/ReportAutomationLauncher.exe
    releases/v0.0.57/SHA256SUMS.txt

The EXE is self-contained (Python runtime, engine, Excel add-in, templates and
license notices are embedded), so no other file is needed. Published folders
are never overwritten; ``releases/index.json`` and ``releases/README.md``
list every version with its source commit and SHA-256.

Usage (Windows, after building and testing the launcher)::

    .venv\\Scripts\\python.exe report_automation_launcher\\scripts\\publish_release.py
    .venv\\Scripts\\python.exe report_automation_launcher\\scripts\\publish_release.py --verify
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

EXE_NAME = "ReportAutomationLauncher.exe"
SUMS_NAME = "SHA256SUMS.txt"
ALLOWED_FILES = {EXE_NAME, SUMS_NAME}
GITHUB_FILE_LIMIT = 100 * 1024 * 1024
VERSION_PATTERN = re.compile(r"^\d+\.\d+\.\d+$")
FOLDER_PATTERN = re.compile(r"^v(\d+)\.(\d+)\.(\d+)$")


class ReleaseError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def git(repository: Path, *args: str) -> str:
    try:
        return subprocess.run(["git", *args], cwd=repository, capture_output=True, text=True, check=True).stdout.strip()
    except (OSError, subprocess.CalledProcessError):
        return ""


def build_source(repository: Path, exe: Path) -> dict:
    """Which commit the EXE came from: its last commit if unchanged, else the working tree."""
    relative = exe.resolve().relative_to(repository.resolve()).as_posix() if exe.resolve().is_relative_to(repository.resolve()) else ""
    if relative and not git(repository, "status", "--porcelain", "--", relative):
        commit = git(repository, "log", "-1", "--format=%H", "--", relative)
        if commit:
            return {"source_commit": commit, "source_dirty": False,
                    "built_at": git(repository, "log", "-1", "--format=%cI", commit)}
    head = git(repository, "rev-parse", "HEAD")
    dirty = bool(git(repository, "status", "--porcelain", "--untracked-files=no", "--", ".", ":(exclude)releases", ":(exclude)" + relative)) if relative else True
    return {"source_commit": head, "source_dirty": dirty, "built_at": datetime.now().astimezone().isoformat(timespec="seconds")}


def release_notes(repository: Path, version: str) -> str:
    path = repository / "VERSIONING.md"
    if not path.is_file():
        return ""
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith(f"`{version}`"):
            return line[len(version) + 2:].strip()
    return ""


def version_key(folder_name: str):
    match = FOLDER_PATTERN.match(folder_name)
    return tuple(int(part) for part in match.groups()) if match else None


def load_index(releases: Path) -> list:
    path = releases / "index.json"
    if not path.is_file():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not isinstance(data.get("releases"), list):
        raise ReleaseError("releases/index.json 형식이 올바르지 않습니다.")
    return data["releases"]


def write_index(releases: Path, entries: list) -> None:
    entries = sorted(entries, key=lambda item: tuple(int(p) for p in item["version"].split(".")), reverse=True)
    (releases / "index.json").write_text(json.dumps({"schema_version": "1.0", "latest": entries[0]["version"] if entries else "",
                                                     "releases": entries}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    lines = [
        "# 배포 파일",
        "",
        "버전별 폴더에는 설치·실행에 필요한 파일만 둡니다. 실행 파일 하나에 Python 런타임, 엔진, Excel 추가기능, 기본 템플릿, 라이선스 고지가 모두 들어 있습니다.",
        "",
        "- `ReportAutomationLauncher.exe`: 원하는 폴더에 복사해 바로 실행합니다(별도 설치 없음). 필요 환경은 [독립 실행 런처](../docs/standalone_launcher.md)를 보세요.",
        "- `SHA256SUMS.txt`: 받은 파일이 배포본과 같은지 확인합니다. PowerShell: `Get-FileHash .\\ReportAutomationLauncher.exe`",
        "",
        "폴더는 한 번 올리면 바꾸지 않습니다. 새 버전은 Windows에서 빌드·검증한 뒤 `report_automation_launcher/scripts/publish_release.py`로 추가합니다.",
        "",
        f"최신 버전: `{entries[0]['version']}`" if entries else "아직 배포한 버전이 없습니다.",
        "",
        "| 버전 | 빌드 기준 커밋 | 빌드 일시 | 크기 | SHA-256 | 변경 내용 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for item in entries:
        size = f"{item['size_bytes'] / 1024 / 1024:.1f} MB"
        commit = item["source_commit"][:7] + (" (수정본)" if item.get("source_dirty") else "")
        notes = item.get("notes", "").replace("|", "\\|") or "-"
        lines.append(f"| [`{item['version']}`](v{item['version']}/) | `{commit}` | {item['built_at'][:10]} | {size} | `{item['sha256'][:16]}…` | {notes} |")
    (releases / "README.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def publish(repository: Path, exe: Path, version: str | None = None, *, allow_unchanged: bool = False, notes: str | None = None) -> Path:
    repository = repository.resolve()
    version = version or (repository / "VERSION").read_text(encoding="utf-8").strip()
    if not VERSION_PATTERN.match(version):
        raise ReleaseError(f"버전 형식이 올바르지 않습니다: {version}")
    if not exe.is_file():
        raise ReleaseError(f"실행 파일이 없습니다: {exe}. 먼저 런처를 빌드하세요.")
    with exe.open("rb") as stream:
        if stream.read(2) != b"MZ":
            raise ReleaseError(f"Windows 실행 파일이 아닙니다: {exe}")
    size = exe.stat().st_size
    if size >= GITHUB_FILE_LIMIT:
        raise ReleaseError("실행 파일이 GitHub 파일 크기 제한(100 MiB)을 넘습니다.")

    releases = repository / "releases"
    target = releases / f"v{version}"
    if target.exists():
        raise ReleaseError(f"{target.name}은 이미 배포했습니다. 배포한 폴더는 바꾸지 않습니다. VERSION을 올린 뒤 다시 빌드하세요.")
    entries = load_index(releases)
    digest = sha256(exe)
    same = next((item for item in entries if item["sha256"] == digest), None)
    if same and not allow_unchanged:
        raise ReleaseError(f"실행 파일이 {same['version']} 배포본과 같습니다. 이 버전의 코드로 다시 빌드했는지 확인하세요"
                           " (실행 파일 변경이 없는 버전이면 배포 폴더를 만들지 않아도 됩니다).")
    source = build_source(repository, exe)
    if source["source_dirty"]:
        print("경고: 커밋하지 않은 변경이 있는 작업 트리에서 빌드했습니다. 빌드 기준 커밋에 '수정본'으로 표시합니다.", file=sys.stderr)

    releases.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=f".v{version}-", dir=releases))
    try:
        shutil.copy2(exe, staging / EXE_NAME)
        if sha256(staging / EXE_NAME) != digest:
            raise ReleaseError("복사한 실행 파일의 SHA-256이 원본과 다릅니다.")
        (staging / SUMS_NAME).write_text(f"{digest} *{EXE_NAME}\n", encoding="utf-8", newline="\n")
        staging.rename(target)
    except BaseException:
        shutil.rmtree(staging, ignore_errors=True)
        raise
    entries.append({"version": version, "file": f"v{version}/{EXE_NAME}", "sha256": digest, "size_bytes": size,
                    **source, "notes": release_notes(repository, version) if notes is None else notes})
    write_index(releases, entries)
    return target


def verify(repository: Path) -> list[str]:
    """Return problems; an empty list means every release folder matches the index."""
    releases = repository.resolve() / "releases"
    problems = []
    if not releases.is_dir():
        return problems
    try:
        entries = {item["version"]: item for item in load_index(releases)}
    except (ReleaseError, ValueError, KeyError) as exc:
        return [str(exc)]
    folders = set()
    for path in sorted(releases.iterdir()):
        if path.is_file():
            if path.name not in {"index.json", "README.md"}:
                problems.append(f"releases/{path.name}: 버전 폴더 밖에는 index.json과 README.md만 둡니다.")
            continue
        if version_key(path.name) is None:
            problems.append(f"releases/{path.name}: 'v주.부.수' 형식의 버전 폴더가 아닙니다.")
            continue
        version = path.name[1:]
        folders.add(version)
        names = {child.name for child in path.iterdir()}
        if names != ALLOWED_FILES:
            problems.append(f"{path.name}: {EXE_NAME}와 {SUMS_NAME}만 있어야 합니다 (현재: {sorted(names)}).")
            continue
        digest = sha256(path / EXE_NAME)
        expected = (path / SUMS_NAME).read_text(encoding="utf-8").split()
        if not expected or expected[0] != digest:
            problems.append(f"{path.name}: {SUMS_NAME}와 실행 파일 해시가 다릅니다.")
        if version not in entries:
            problems.append(f"{path.name}: index.json에 없습니다.")
        elif entries[version]["sha256"] != digest or entries[version]["size_bytes"] != (path / EXE_NAME).stat().st_size:
            problems.append(f"{path.name}: index.json의 해시나 크기가 실제 파일과 다릅니다.")
    for version in sorted(set(entries) - folders):
        problems.append(f"index.json의 {version}에 해당하는 폴더가 없습니다.")
    return problems


def main(argv=None) -> int:
    repository = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(description="Publish the built launcher into releases/v<VERSION>/ or verify releases.")
    parser.add_argument("--exe", default=str(repository / "report_automation_launcher" / "bin" / EXE_NAME))
    parser.add_argument("--version", help="Defaults to the VERSION file.")
    parser.add_argument("--allow-unchanged", action="store_true", help="Publish even if the EXE equals an earlier release.")
    parser.add_argument("--notes", help="Release summary for the index (default: this version's VERSIONING.md line).")
    parser.add_argument("--verify", action="store_true", help="Check every release folder against index.json.")
    args = parser.parse_args(argv)
    if args.verify:
        problems = verify(repository)
        for problem in problems:
            print(problem, file=sys.stderr)
        print("릴리스 폴더 검증 통과" if not problems else f"릴리스 폴더 문제 {len(problems)}건")
        return 1 if problems else 0
    try:
        target = publish(repository, Path(args.exe), args.version, allow_unchanged=args.allow_unchanged, notes=args.notes)
    except (ReleaseError, OSError) as exc:
        print(f"배포 중단: {exc}", file=sys.stderr)
        return 2
    print(f"배포 폴더: {target}")
    print("다음 단계: git add releases && git commit, 필요하면 태그 v" + target.name[1:])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
