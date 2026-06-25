# Versioning

Current version: `0.0.1`

## Policy

- Use semantic versioning: `MAJOR.MINOR.PATCH`.
- Every requested push or publish should include an explicit version check.
- If code, docs, or packaged binaries changed, bump `VERSION` before pushing.
- Create a matching Git tag for release points: `v0.0.1`, `v0.0.2`, ...
- Default bump:
  - Patch: fixes, docs, small UX improvements, internal hardening.
  - Minor: new user-facing workflow or output format.
  - Major: breaking file format, CLI, template, or workflow change.

## Current Baseline

`0.0.1` marks the first launcher-based alpha baseline:

- Excel report output sheets
- Python draft text generation
- `report_package.json` and `preflight_report.json`
- template inspection/factory/autofix
- planning for direct HWPX/PPTX output
