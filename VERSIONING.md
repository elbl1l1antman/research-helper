# Versioning

Current version: `0.0.47`

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

`0.0.47` inserts survey charts into HWPX as verified original EMFs, preserving caption slots and chart state across recovery.

`0.0.46` adds launcher registration of a common PPTX chart style and an editable survey chart review export.

`0.0.45` verifies keyboard selection and preservation of the original chart EMF after reopening and saving the HWPX probe in Hancom. No launcher changes.

`0.0.44` adds the HWPX figure-container probe and records the observed WMF conversion. Automatic chart insertion is not enabled.

`0.0.43` documents the chart review and HWPX EMF integration plans. No executable changes.

`0.0.42` adds two editable HWPX survey starters and three role-based body and
table presets, including a Malgun Gothic option.

`0.0.41` preserves fully covered merge rows and adds native HWP captions,
table number styles, inline table placement, and numeric display options.

`0.0.40` recovers supported survey measures from explicit Excel cell evidence
and resumes interrupted HWPX jobs from the last saved section.

`0.0.39` closes Excel before QA, extends the table preview beyond 5,000 rows,
and uses HWP progress checkpoints to distinguish slow work from a stall.

`0.0.38` fixes HWP report section selection and respondent profile summaries,
and adds a separate destination for report outputs.

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
