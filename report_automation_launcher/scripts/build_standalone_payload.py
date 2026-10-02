"""Build and smoke-test the launcher payload from the installed build environment."""

import argparse
import csv
import importlib.metadata as metadata
import json
import os
import struct
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path, PurePosixPath, PureWindowsPath

from packaging.requirements import Requirement
from packaging.tags import parse_tag, sys_tags
from packaging.utils import canonicalize_name


ROOT_DISTRIBUTIONS = ("openpyxl", "python-pptx", "Pillow", "pywin32")
PYTHON_TAG = "{}{}".format(*sys.version_info[:2])
EXCLUDED = {"site-packages", "test", "tests", "__pycache__", "outputs", "output", "env", "venv", "research", "user_research"}
RUNTIME_PTH = "runtime/Lib/site-packages/report_automation_runtime.pth"
# Keep the handle alive: closing it removes the portable pywin32 DLL search path.
RUNTIME_PTH_TEXT = (
    "import os, sys; sys._report_automation_dll_handles = "
    "[os.add_dll_directory(os.path.join(sys.prefix, 'Lib', 'site-packages', 'pywin32_system32'))]\n"
)


def safe_name(name):
    path = PurePosixPath(name)
    if not path.parts or "\\" in name or ":" in name or path.is_absolute() or PureWindowsPath(name).drive or ".." in path.parts:
        raise ValueError("Unsafe payload path: " + name)
    return path


def excluded(path):
    return any(part.lower() in EXCLUDED or part.startswith((".", "~$")) for part in path.parts) or path.suffix.lower() in {".pyc", ".pyo", ".tmp"}


def required_distributions(site_paths):
    available = {}
    for distribution in metadata.distributions(path=[str(p) for p in site_paths]):
        available.setdefault(canonicalize_name(distribution.metadata["Name"]), distribution)
    selected = {}
    pending = [Requirement(name) for name in ROOT_DISTRIBUTIONS]
    while pending:
        requirement = pending.pop()
        if requirement.marker and not requirement.marker.evaluate({"extra": ""}):
            continue
        name = canonicalize_name(requirement.name)
        if name not in available:
            raise ValueError("Required installed distribution not found: " + name)
        distribution = available[name]
        if requirement.specifier and not requirement.specifier.contains(distribution.version, prereleases=True):
            raise ValueError("Incompatible installed dependency: " + str(requirement))
        if name not in selected:
            selected[name] = distribution
            pending.extend(Requirement(value) for value in distribution.requires or [])
    return selected


def collect_payload_files(repository, base_prefix, distributions):
    files = {}

    def add(source, root, destination):
        safe_name(destination)
        source.relative_to(root)
        for part in (source, *source.parents):
            if part.is_symlink() or part.is_junction():
                raise ValueError("Linked payload source is not allowed: " + str(source))
            if part == root:
                break
        if not source.is_file() or not source.resolve().is_relative_to(root.resolve()):
            raise FileNotFoundError("Missing or external payload source: " + str(source))
        if destination in files and files[destination] != source:
            raise ValueError("Duplicate payload path: " + destination)
        files[destination] = source

    def tree(root, source_root, prefix, accept=lambda p: True):
        if not root.is_dir():
            raise FileNotFoundError("Payload directory not found: " + str(root))
        pending = [root]
        while pending:
            directory = pending.pop()
            if directory.is_symlink() or directory.is_junction():
                raise ValueError("Linked payload directory is not allowed: " + str(directory))
            for source in sorted(directory.iterdir()):
                relative = source.relative_to(root)
                if excluded(relative):
                    continue
                if source.is_symlink() or source.is_junction():
                    raise ValueError("Linked payload source is not allowed: " + str(source))
                if source.is_dir():
                    pending.append(source)
                elif accept(relative):
                    add(source, source_root, prefix + relative.as_posix())

    for name in ("python.exe", "python" + PYTHON_TAG + ".dll", "LICENSE.txt"):
        add(base_prefix / name, base_prefix, "runtime/" + name)
    for pattern in ("python*.exe", "python*.dll", "vcruntime*.dll"):
        for source in base_prefix.glob(pattern):
            add(source, base_prefix, "runtime/" + source.name)
    tree(base_prefix / "Lib", base_prefix, "runtime/Lib/")
    tree(base_prefix / "DLLs", base_prefix, "runtime/DLLs/",
         lambda p: p.name.lower() != "_tkinter.pyd" and not (p.suffix.lower() == ".dll" and p.name.lower().startswith(("tcl", "tk"))))

    for distribution in distributions.values():
        root = Path(distribution.locate_file(""))
        record_text = distribution.read_text("RECORD")
        if not record_text:
            raise ValueError("Distribution has no installed file record: " + distribution.metadata["Name"])
        for record in csv.reader(record_text.splitlines()):
            name = record[0]
            # Wheel console scripts are build-environment entry points, not runtime libraries.
            if name.startswith("../../Scripts/") and ".." not in PurePosixPath(name).parts[2:]:
                continue
            relative = safe_name(name)
            if excluded(relative) or relative.name == "direct_url.json":
                continue
            add(root / name, root, "runtime/Lib/site-packages/" + name)

    tree(repository / "report_automation_engine", repository, "report_automation_engine/", lambda p: p.suffix == ".py")
    for name in ("default_style_schema.json", "default_hwp_style_config.json"):
        path = "report_automation_engine/config/" + name
        add(repository / path, repository, path)
    addin = "report_automation_addin/dev/ReportAutomationAddin_dev.xlam"
    add(repository / addin, repository, addin)
    for name in ("report_template_basic.pptx", "chart_review_template_basic.pptx"):
        add(repository / "templates" / name, repository, "templates/" + name)
    add(repository / "LICENSE", repository, "LICENSE")
    license_root = repository / "report_automation_launcher/licenses"
    if license_root.is_dir():
        tree(license_root, repository, "licenses/", lambda p: p.suffix.lower() in {".txt", ".rst", ".python"})
    for name in ("THIRD_PARTY_NOTICES.md", "docs/template_authoring_guide.md", "docs/hwp_page_setup.md", "report_automation_engine/README.md"):
        if (repository / name).exists():
            add(repository / name, repository, name)
    return files


def pe_machine(path):
    with path.open("rb") as stream:
        if stream.read(2) != b"MZ":
            raise ValueError("Not a Windows native binary: " + str(path))
        stream.seek(0x3C)
        offset = struct.unpack("<I", stream.read(4))[0]
        stream.seek(offset)
        if stream.read(4) != b"PE\0\0":
            raise ValueError("Invalid PE header: " + str(path))
        return struct.unpack("<H", stream.read(2))[0]


def validate_licenses(files, distributions):
    if "THIRD_PARTY_NOTICES.md" not in files:
        raise ValueError("Root THIRD_PARTY_NOTICES.md is required for standalone distribution")
    for name, distribution in distributions.items():
        if name in {"openpyxl", "et-xmlfile"}:
            stem = "licenses/" + name.replace("-", "_") + "-" + distribution.version + "-LICENCE"
            notices = [stem + ".rst"] + ([stem + ".python"] if name == "et-xmlfile" else [])
        else:
            notices = ["runtime/Lib/site-packages/" + str(p) for p in distribution.files or []
                       if any(word in str(p).lower() for word in ("license", "licence", "copying"))]
        if not notices or any(p not in files or not files[p].stat().st_size for p in notices):
            raise ValueError("Packaged license missing for " + name + " " + distribution.version)


def validate_runtime(files, distributions):
    if sys.platform != "win32" or sys.version_info[:2] < (3, 12) or struct.calcsize("P") != 8:
        raise ValueError("Build requires an installed Windows x64 Python >= 3.12 environment")
    validate_licenses(files, distributions)
    supported_tags = set(sys_tags())
    for distribution in distributions.values():
        wheel = distribution.read_text("WHEEL") or ""
        tags = set().union(*(parse_tag(line[5:].strip()) for line in wheel.splitlines() if line.startswith("Tag: ")))
        if not tags or not tags.intersection(supported_tags):
            raise ValueError("Incompatible wheel ABI: " + distribution.metadata["Name"])
    for name, source in files.items():
        if source.suffix.lower() in {".exe", ".dll", ".pyd"} and pe_machine(source) != 0x8664:
            raise ValueError("Payload binary is not x64: " + name)
    for name in ("pythoncom" + PYTHON_TAG + ".dll", "pywintypes" + PYTHON_TAG + ".dll"):
        if "runtime/Lib/site-packages/pywin32_system32/" + name not in files:
            raise ValueError("Required pywin32 DLL missing: " + name)


def write_payload(output, files, python_version, distributions):
    if RUNTIME_PTH in files or "bundle_manifest.json" in files:
        raise ValueError("Reserved payload path")
    manifest = {"python_version": python_version, "architecture": "x64",
                "dependencies": {name: d.version for name, d in sorted(distributions.items())},
                "files": sorted([*files, RUNTIME_PTH])}
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for name, source in sorted(files.items()):
            safe_name(name)
            archive.write(source, name)
        # Generated entries use stable dates so an unchanged payload reuses the launcher cache.
        archive.writestr(zipfile.ZipInfo(RUNTIME_PTH), RUNTIME_PTH_TEXT, compress_type=zipfile.ZIP_DEFLATED)
        archive.writestr(zipfile.ZipInfo("bundle_manifest.json"), json.dumps(manifest, indent=2) + "\n", compress_type=zipfile.ZIP_DEFLATED)
    return manifest


SMOKE_CODE = r"""
import sys, json, struct
from pathlib import Path
root = Path(sys.argv[1]).resolve()
manifest = json.loads((root / 'bundle_manifest.json').read_text(encoding='utf-8'))
assert sys.flags.isolated and struct.calcsize('P') == 8
assert sys.version.split()[0] == manifest['python_version']
assert Path(sys.prefix).resolve() == root / 'runtime'
import ssl, sqlite3, bz2, lzma, ctypes, xml.etree.ElementTree, encodings
import win32api, win32com, win32com.client, pythoncom, pywintypes, openpyxl, pptx, PIL
from PIL import Image
from pptx import Presentation
from openpyxl import Workbook
from io import BytesIO
for module in (ssl, sqlite3, win32api, win32com, pythoncom, pywintypes, openpyxl, pptx, PIL):
    assert Path(module.__file__).resolve().is_relative_to(root), module.__file__
assert ctypes.CDLL(pythoncom.__file__) and ctypes.CDLL(pywintypes.__file__)
assert ssl.create_default_context() and sqlite3.connect(':memory:').execute('select 1').fetchone() == (1,)
image = BytesIO(); Image.new('RGB', (1, 1)).save(image, format='PNG')
workbook = BytesIO(); Workbook().save(workbook)
slides = BytesIO(); Presentation().save(slides)
assert image.getvalue() and workbook.getvalue() and slides.getvalue()
print(json.dumps({'python_version': sys.version.split()[0], 'architecture': 'x64', 'imports': 'ok', 'com_created': False}))
"""


def smoke_payload(payload):
    with tempfile.TemporaryDirectory(prefix="ReportAutomationPayload-") as directory:
        root = Path(directory)
        with zipfile.ZipFile(payload) as archive:
            for name in archive.namelist():
                safe_name(name)
            archive.extractall(root)
        unrelated = root / "unrelated-working-directory"
        unrelated.mkdir()
        environment = {key: value for key, value in os.environ.items() if key.upper() not in {"PYTHONHOME", "PYTHONPATH"}}
        windows = environment.get("SystemRoot", r"C:\Windows")
        environment["PATH"] = os.pathsep.join([str(Path(windows) / "System32"), windows])
        result = subprocess.run([str(root / "runtime/python.exe"), "-I", "-B", "-c", SMOKE_CODE, str(root)],
                                cwd=unrelated, env=environment, capture_output=True, text=True, timeout=120)
        if result.returncode or result.stderr.strip():
            raise RuntimeError("Standalone payload smoke failed:\n" + result.stdout + result.stderr)
        print(result.stdout.strip())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repository", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    site_paths = [Path(p) for p in sys.path if Path(p).name == "site-packages"]
    distributions = required_distributions(site_paths)
    files = collect_payload_files(args.repository.absolute(), Path(sys.base_prefix), distributions)
    validate_runtime(files, distributions)
    print("Validated Python " + sys.version.split()[0] + " x64; packaging " + str(len(files)) + " files", flush=True)
    write_payload(args.output, files, sys.version.split()[0], distributions)
    size = args.output.stat().st_size
    if size >= 100 * 1024 * 1024:
        raise ValueError("Compressed payload exceeds the GitHub 100 MiB file limit")
    print("Payload compressed; checking extracted isolated runtime", flush=True)
    smoke_payload(args.output)
    print("Created standalone payload: " + str(args.output) + " ({:.2f} MiB)".format(size / (1024 * 1024)))


if __name__ == "__main__":
    main()
