"""Auto-resume check: start the writer with --recovery, kill ONLY its own writer
and Hancom process after N sections, rerun the same command and confirm it resumes.

Usage:
  resume_test.py <package_prefix> <out_dir> [--kill-after 5] [--max-sections 12]
  resume_test.py <package_prefix> <out_dir> --continue   # rerun only (finish an existing recovery)

<package_prefix> is the path before "_report_package.json"; the matching
_preflight_report.json and _hwp_style_config.json must sit next to it.
Hancom processes that existed before the test are never touched.
"""
import argparse
import json
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
parser = argparse.ArgumentParser()
parser.add_argument("prefix")
parser.add_argument("out_dir")
parser.add_argument("--template", default=str(ROOT / "templates/report_template_newspaper.hwpx"))
parser.add_argument("--kill-after", type=int, default=5)
parser.add_argument("--max-sections", default="12")
parser.add_argument("--continue", dest="resume_only", action="store_true")
args = parser.parse_args()

P, out = args.prefix, Path(args.out_dir)
out.mkdir(parents=True, exist_ok=True)
rec = out / "resume.recovery.json"
report = out / "resume_writer_report.json"
cmd = [str(ROOT / ".venv/Scripts/python.exe"), "-m", "report_automation_engine.hwp_com_writer",
       "--package", f"{P}_report_package.json", "--preflight", f"{P}_preflight_report.json",
       "--template", args.template, "--style-config", f"{P}_hwp_style_config.json",
       "--output", str(out / "resume.hwpx"), "--report-output", str(report),
       "--recovery", str(rec), "--max-sections", args.max_sections]


def hwp_pids():
    result = subprocess.run(["powershell", "-NoProfile", "-c", "(Get-Process Hwp -ErrorAction SilentlyContinue).Id"],
                            capture_output=True, text=True)
    return {int(x) for x in result.stdout.split()}


def committed():
    return json.loads(rec.read_text(encoding="utf-8"))["report"].get("sections_written", 0) if rec.exists() else 0


before = hwp_pids()
print("hwp before (left alone):", sorted(before))
if not args.resume_only:
    start = time.time()
    with open(out / "run1.log", "w") as log:
        proc = subprocess.Popen(cmd, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT)
    while committed() < args.kill_after:
        if proc.poll() is not None:
            sys.exit(f"writer finished before the kill point (exit {proc.returncode})")
        if time.time() - start > 900:
            sys.exit("no recovery point after 15 min")
        time.sleep(2)
    mine = hwp_pids() - before
    print(f"kill at {committed()} sections after {time.time() - start:.0f}s; writer {proc.pid}, hwp {sorted(mine)}")
    proc.kill()
    for pid in mine:
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
    time.sleep(3)

print("recovery point:", committed(), "sections")
start = time.time()
try:
    with open(out / "run2.log", "w") as log:
        code = subprocess.run(cmd, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT).returncode
finally:
    for pid in hwp_pids() - before:
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
w = json.loads(report.read_text(encoding="utf-8"))
print(json.dumps({"exit": code, "secs": round(time.time() - start), "status": w["status"],
                  "resumed_sections": w.get("resumed_sections"), "sections_written": w["sections_written"],
                  "tables_written": w["tables_written"], "errors": w["errors"],
                  "recovery_left": rec.exists()}, ensure_ascii=False))
