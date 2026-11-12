# Versioning

Current version: `0.0.37`

## Policy

- Use semantic versioning: `MAJOR.MINOR.PATCH`.
- Push completed local commits to their remote work branch by default. Merging into `main` and publishing a GitHub Release remain separate decisions.
- Check the version before each push or publish.
- If code, docs, or packaged binaries changed, bump `VERSION` before pushing.
- Create a matching Git tag for release points: `v0.0.1`, `v0.0.2`, ...
- Keep deprecated assets under `old/legacy/` with a short note explaining what replaced them.
- Default bump:
  - Patch: fixes, docs, small UX improvements, internal hardening.
  - Minor: new user-facing workflow or output format.
  - Major: breaking file format, CLI, template, or workflow change.

## Current Baseline

`0.0.37` adds live narrative and table examples to the HWP style editor and
requires saving edits before applying a changed preset.

`0.0.36` removes manual Excel add-in selection from the launcher, loads the bundled
add-in automatically, and keeps connection diagnostics available on failure.

`0.0.35` packages the editable HWPX starter and two built-in table presets,
improves horizontal and vertical banner detection, and generates Excel output
sheets automatically before HWPX preflight.

`0.0.3` adds the `old/legacy/` archive for version-managed retired assets.

`0.0.2` updates the GitHub README with the project structure and current alpha workflow.

`0.0.1` marks the first launcher-based alpha baseline:

- Excel report output sheets
- Python draft text generation
- `report_package.json` and `preflight_report.json`
- template inspection/factory/autofix
- planning for direct HWPX/PPTX output
