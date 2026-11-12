# HWP Named Styles and Table Cleanup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Excel display text, clean decorative symbols, split report narratives into styled paragraphs, and register/apply five user-configurable named styles in generated HWPX documents.

**Architecture:** Normalize text and narrative blocks in the existing Excel-to-package boundary, then patch the HWPX working template with deterministic named style definitions before COM opens it. The existing COM writer applies the returned style indexes to body paragraphs and table-cell paragraphs; the WinForms launcher persists up to three user presets and emits one selected style-config JSON per run.

**Tech Stack:** VBA, Python 3 standard library (`html`, `unicodedata`, `zipfile`, `xml.etree.ElementTree`, `json`), pywin32/HWP Automation, C# WinForms/.NET Framework.

**Spec:** `docs/superpowers/specs/2026-09-30-hwp-named-style-and-table-sanitization-design.md`

## Global Constraints

- Original Excel workbooks and HWP/HWPX templates must never be modified.
- Required style names are fixed: `보고서 본문1`, `보고서 본문2`, `표보기`, `표배너`, `표숫자`.
- The built-in default preset cannot be deleted; users may save at most three additional presets.
- `display_text` is authoritative for HWP output; Python must not convert it back to a number.
- No new runtime dependency may be added.
- PPTX behavior is out of scope.
- Every task ends with a commit and push to the current branch.

## Review Focus

- Existing templates containing one or more required style names must update those styles without producing duplicate names; Task 4 tests this.
- A cell containing only removed symbols must become empty and emit a cell-address QA warning; Task 1 tests this.
- Genuine negative numbers must remain negative while parenthesized display strings stay parenthesized; Task 1 tests both.
- Corrupt preset storage, duplicate names, and a fourth user preset must recover or fail visibly without losing valid presets; Task 6 tests this.
- Narratives with three or more lines and old concatenated transition words must preserve order and apply body2 styling to every follow-up block; Task 2 tests this.

---

### Task 1: Preserve Excel Display Text and Sanitize Table Cells

**Files:**
- Modify: `report_automation_addin/src/ReportAutomationTableCells.bas`
- Modify: `report_automation_engine/report_table_matrix.py`
- Create: `tests/test_report_table_text.py`

**Interfaces:**
- Produces: `sanitize_table_display_text(value: Any) -> str`
- Produces: matrix cells with optional `original_display_text` and `removed_symbols`
- Consumes: existing `ReportAutomation_CellDisplayText`, `build_table_matrix_from_cells`

- [ ] **Step 1: Write failing stdlib unittest cases**

Cover these assertions in `tests/test_report_table_text.py`:

- `(3,232)` remains `(3,232)`.
- `-3232` remains `-3232` when it is the real display text.
- `● 만족도 3.3%` becomes `만족도 3.3%`.
- `&#x20;'-' "문장" ~ ㎡ ㎏` preserves the allowed punctuation and units after entity decoding.
- `●◆` becomes an empty string and the matrix QA contains the source cell address.

- [ ] **Step 2: Run the new test and confirm failure**

Run: `.venv\Scripts\python.exe tests\test_report_table_text.py`

Expected: failure because `sanitize_table_display_text` does not exist.

- [ ] **Step 3: Fix VBA text preservation at the source boundary**

In `ReportAutomation_WriteTableCells`, set column H to text format before rows are written. In `ReportAutomation_WriteTableRangeCells`, write `display_text` with `Value2 = CStr(...)`; write string `raw_value` into a text-formatted target cell and numeric `raw_value` as numeric.

- [ ] **Step 4: Implement table text sanitization**

Add `sanitize_table_display_text(value: Any) -> str` to `report_table_matrix.py` using `html.unescape` and `unicodedata.category`. Preserve punctuation, mathematical/currency symbols, and the unit whitelist from the spec; remove control characters and decorative `So` characters outside that whitelist.

Apply it only to `display_text` in both matrix builders. When text changes, retain the original value and removed characters in the cell and add one QA warning containing `source_cell`.

- [ ] **Step 5: Run focused verification**

Run:

```powershell
.venv\Scripts\python.exe tests\test_report_table_text.py
.venv\Scripts\python.exe -m compileall -q report_automation_engine
```

Expected: all assertions pass and compilation exits 0.

- [ ] **Step 6: Build the VBA add-in smoke artifact**

Run the repository's existing add-in build/import smoke command documented in `README.md` and confirm `ReportAutomationTableCells.bas` imports without compile errors.

- [ ] **Step 7: Commit and push**

```bash
git add report_automation_addin/src/ReportAutomationTableCells.bas report_automation_engine/report_table_matrix.py tests/test_report_table_text.py
git commit -m "fix: preserve and clean HWP table display text"
git push
```

### Task 2: Normalize Narrative Blocks and Restore Lost Line Breaks

**Files:**
- Modify: `report_automation_engine/report_package.py`
- Modify: `report_automation_engine/hwp_com_writer.py`
- Create: `tests/test_report_narrative_blocks.py`

**Interfaces:**
- Produces: `split_narrative_blocks(text: str) -> List[Dict[str, str]]`
- Produces: `sections[].narrative_blocks`
- Consumes: existing `sections[].narrative_final`

- [ ] **Step 1: Write failing stdlib unittest cases**

Test:

- CRLF, LF, and CR each create separate ordered blocks.
- `가장 높게 나타남다음으로 ...` splits before `다음으로`.
- `나타남그다음으로` and `나타남반면,` split at the transition.
- A three-line narrative yields one `보고서 본문1` block and two `보고서 본문2` blocks.
- A single sentence remains one body1 block.

- [ ] **Step 2: Run the new test and confirm failure**

Run: `.venv\Scripts\python.exe tests\test_report_narrative_blocks.py`

Expected: failure because `split_narrative_blocks` does not exist.

- [ ] **Step 3: Implement package normalization**

Add `split_narrative_blocks(text: str) -> List[Dict[str, str]]` to `report_package.py`. Normalize newline forms first, then use only the transition boundaries listed in the spec as legacy recovery. Add `narrative_blocks` in `read_sections` while preserving `narrative_final` unchanged.

- [ ] **Step 4: Make the writer emit one paragraph per block**

Add `insert_narrative_blocks(hwp, section, report, style_indexes=None)` in `hwp_com_writer.py`. Insert each block separately with `BreakPara`; remove a legacy literal body2 bullet prefix so Task 5 can supply the real style bullet without duplication. Until Task 5 supplies style indexes, the function must still produce correct line breaks.

- [ ] **Step 5: Run focused verification**

Run:

```powershell
.venv\Scripts\python.exe tests\test_report_narrative_blocks.py
.venv\Scripts\python.exe -m compileall -q report_automation_engine
```

Expected: all assertions pass.

- [ ] **Step 6: Commit and push**

```bash
git add report_automation_engine/report_package.py report_automation_engine/hwp_com_writer.py tests/test_report_narrative_blocks.py
git commit -m "fix: split HWP narratives into report paragraphs"
git push
```

### Task 3: Define and Validate HWP Style Presets

**Files:**
- Create: `report_automation_engine/hwp_style_config.py`
- Create: `report_automation_engine/config/default_hwp_style_config.json`
- Create: `tests/test_hwp_style_config.py`

**Interfaces:**
- Produces: `load_hwp_style_config(path: Path | None) -> Dict[str, Any]`
- Produces: `validate_hwp_style_config(config: Dict[str, Any]) -> Dict[str, Any]`
- Produces: `style_name_for_cell_role(role: str) -> str`
- Produces: constants `REQUIRED_STYLE_NAMES`, `DEFAULT_HWP_STYLE_CONFIG`

- [ ] **Step 1: Write failing stdlib unittest cases**

Assert the exact five required style names, spec default values, numeric ranges, missing-style rejection, and role mapping. Include malformed color, 31pt font, 301% line spacing, and `blank`/`unknown` role cases.

- [ ] **Step 2: Run the new test and confirm failure**

Run: `.venv\Scripts\python.exe tests\test_hwp_style_config.py`

Expected: import failure for `hwp_style_config`.

- [ ] **Step 3: Implement the minimal config module and default JSON**

Use standard-library JSON only. Validation returns a normalized deep copy and raises `ValueError` with the preset field name on invalid input. Keep style names fixed and permit only their values to change.

- [ ] **Step 4: Run focused verification**

Run:

```powershell
.venv\Scripts\python.exe tests\test_hwp_style_config.py
.venv\Scripts\python.exe -m compileall -q report_automation_engine
```

Expected: all assertions pass.

- [ ] **Step 5: Commit and push**

```bash
git add report_automation_engine/hwp_style_config.py report_automation_engine/config/default_hwp_style_config.json tests/test_hwp_style_config.py
git commit -m "feat: add validated HWP style presets"
git push
```

### Task 4: Register Named Styles in the HWPX Working Template

**Files:**
- Create: `report_automation_engine/hwpx_style_registry.py`
- Create: `tests/test_hwpx_style_registry.py`

**Interfaces:**
- Consumes: normalized config from `load_hwp_style_config`
- Produces: `register_named_styles(template: Path, output: Path, config: Dict[str, Any]) -> Dict[str, int]`
- Produces: style-name to HWP style-index mapping

- [ ] **Step 1: Write a minimal HWPX fixture and failing tests**

The test creates a ZIP fixture containing `mimetype`, `Contents/header.xml`, and required manifest files. Assert:

- all five styles are added with unique names and indexes;
- `itemCnt` values match child counts;
- font size, bold, line spacing, alignment, and indents match config;
- rerunning on a file containing `보고서 본문1` updates it without duplication;
- input ZIP bytes are unchanged;
- invalid/missing `Contents/header.xml` raises a clear error.

- [ ] **Step 2: Run the new test and confirm failure**

Run: `.venv\Scripts\python.exe tests\test_hwpx_style_registry.py`

Expected: import failure for `hwpx_style_registry`.

- [ ] **Step 3: Implement deterministic HWPX XML registration**

Use `zipfile` and `ElementTree`. Copy the template to a separate output ZIP, update `fontfaces`, `charProperties`, `paraProperties`, `bullets`, and `styles`, and preserve all unrelated members and ZIP ordering. Clone the nearest valid base `font`, `charPr`, and `paraPr` elements before changing required attributes. Create or update a real HWP bullet definition for `보고서 본문2` and point its `paraPr` heading to that bullet ID.

Return style indexes in document style order. Existing matching local names are updated in place; missing styles are appended.

- [ ] **Step 4: Run focused verification**

Run:

```powershell
.venv\Scripts\python.exe tests\test_hwpx_style_registry.py
.venv\Scripts\python.exe -m compileall -q report_automation_engine
```

Expected: all assertions pass and generated fixtures reopen as ZIP files.

- [ ] **Step 5: Commit and push**

```bash
git add report_automation_engine/hwpx_style_registry.py tests/test_hwpx_style_registry.py
git commit -m "feat: register report styles in HWPX templates"
git push
```

### Task 5: Apply Named Styles and Cell Appearance in the COM Writer

**Files:**
- Modify: `report_automation_engine/hwp_com_writer.py`
- Modify: `tests/test_hwp_com_writer_tables.py`
- Create: `tests/test_hwp_com_writer_styles.py`

**Interfaces:**
- Consumes: `--style-config PATH`
- Consumes: `register_named_styles(...) -> Dict[str, int]`
- Consumes: `style_name_for_cell_role(role: str) -> str`
- Produces: writer report style counts and style-index map

- [ ] **Step 1: Write failing writer tests with fake COM objects**

Test:

- `Style` Action receives the expected `Apply` index before each paragraph insertion;
- body1 is applied once and body2 to every follow-up paragraph;
- table roles map to `표보기`, `표배너`, and `표숫자`;
- a failed style registration blocks generation instead of silently using direct formatting;
- the configured body2 bullet is a real style bullet and an existing literal prefix is not duplicated.

- [ ] **Step 2: Run the focused tests and confirm failure**

Run:

```powershell
.venv\Scripts\python.exe tests\test_hwp_com_writer_styles.py
.venv\Scripts\python.exe tests\test_hwp_com_writer_tables.py
```

Expected: new style tests fail before implementation.

- [ ] **Step 3: Add CLI and working-template preparation**

Add `--style-config`. For HWPX templates, create a temporary patched HWPX and open that file. For HWP templates, use the existing COM save path to create a temporary HWPX, patch it, close it, and reopen it before body generation. Never patch the source template or the final output in place.

- [ ] **Step 4: Apply styles to narrative paragraphs and table cells**

Add `apply_named_style(hwp, style_index: int, report, style_name: str) -> None` using `HParameterSet.HStyle.Apply` and the `Style` Action. Apply a style before inserting each narrative block and each cell's display text.

Use `CellBorderFill`, table cell selection, and existing table navigation to apply the preset's fill, border, and vertical alignment after cell creation. Record per-style paragraph/cell counts.

- [ ] **Step 5: Run unit and compile verification**

Run:

```powershell
.venv\Scripts\python.exe tests\test_hwp_com_writer_styles.py
.venv\Scripts\python.exe tests\test_hwp_com_writer_tables.py
.venv\Scripts\python.exe -m compileall -q report_automation_engine
```

Expected: all checks pass.

- [ ] **Step 6: Run one-section HWPX COM integration**

Generate one report section with the default style config, reopen it with HWP COM, and verify:

- the document opens;
- one table control exists;
- all five style names occur once in `Contents/header.xml` after HWPX save;
- page text contains separate body1/body2 lines;
- writer report shows nonzero application counts.

- [ ] **Step 7: Commit and push**

```bash
git add report_automation_engine/hwp_com_writer.py tests/test_hwp_com_writer_tables.py tests/test_hwp_com_writer_styles.py
git commit -m "feat: apply named HWP styles to reports"
git push
```

### Task 6: Add Three-Slot Style Preset UI and Persistence

**Files:**
- Modify: `report_automation_launcher/src/ReportAutomationLauncher.cs`
- Modify: `report_automation_launcher/README.md`

**Interfaces:**
- Produces: `%LOCALAPPDATA%\ResearchHelper\hwp_style_presets.json`
- Produces: selected run file `hwp_style_config.json`
- Passes: `--style-config <path>` to `hwp_com_writer.py`

- [ ] **Step 1: Add pure preset-store self-checks inside the launcher source**

Create a small `HwpStylePresetStore` class in the existing source file. Its self-check must cover default loading, three user presets, fourth-preset rejection, duplicate-name overwrite, deletion, and corrupt-JSON recovery while retaining a backup of the corrupt file.

- [ ] **Step 2: Implement preset storage**

Use `JavaScriptSerializer` and `%LOCALAPPDATA%`. Store only user presets; merge the immutable built-in default at load time. Validate before writing and write through a temporary file followed by replace/move.

- [ ] **Step 3: Add the HWP formatting dialog**

Add a preset combo and `서식 설정` button to the HWPX options. The modal dialog edits the five fixed style rows plus table border values, validates the spec ranges, and exposes save, overwrite, delete, and reset actions. Keep the existing one-file launcher build pattern.

- [ ] **Step 4: Emit and pass the selected style config**

Before HWP generation, serialize the selected normalized preset as `hwp_style_config.json` beside `report_package.json`. Add `HwpStyleConfigPath` and selected preset name to `LauncherOptions`, launcher audit config, and the writer command line.

- [ ] **Step 5: Build and run launcher checks**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File report_automation_launcher\scripts\build_report_automation_launcher.ps1
report_automation_launcher\bin\ReportAutomationLauncher.exe --self-check-hwp-style-presets
```

Expected: build succeeds and the self-check exits 0.

- [ ] **Step 6: Commit and push**

```bash
git add report_automation_launcher/src/ReportAutomationLauncher.cs report_automation_launcher/README.md
git commit -m "feat: add reusable HWP style presets"
git push
```

### Task 7: Real-Workbook Regression and Documentation

**Files:**
- Modify: `README.md`
- Modify: `report_automation_engine/README.md`
- Modify: `docs/next_development_plan.md`

**Interfaces:**
- Consumes: all previous task outputs
- Produces: verified KISDI three-table HWPX and companion writer report under ignored `outputs/`

- [ ] **Step 1: Rebuild the add-in and launcher**

Run the existing add-in build plus:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File report_automation_launcher\scripts\build_report_automation_launcher.ps1
```

Expected: both builds succeed.

- [ ] **Step 2: Generate a three-table KISDI working copy**

Use `3. 통계표_kisdi2025 디지털산업 실태조사 보고서 Table(%).xlsx`, including at least one table containing parenthesized base counts and one containing decorative symbols. Keep source files unchanged.

- [ ] **Step 3: Generate and reopen the HWPX output**

Verify with HWP COM and HWPX XML inspection:

- three independent table controls;
- separate body1/body2 paragraphs;
- `(3,232)`-style values preserved;
- decorative symbols removed while punctuation and units remain;
- each required style name exists exactly once;
- writer report has zero errors and expected style application counts.

- [ ] **Step 4: Update user documentation**

Document default values, three-slot preset behavior, style names, special-character policy, and recovery behavior. Mark the four reported issues complete in `docs/next_development_plan.md`.

- [ ] **Step 5: Run final verification**

Run all new stdlib tests, existing writer checks, `compileall`, launcher build, `git diff --check`, and `git status --short`. Confirm no HWP or Excel automation processes remain.

- [ ] **Step 6: Commit and push**

```bash
git add README.md report_automation_engine/README.md docs/next_development_plan.md
git commit -m "docs: document HWP report style controls"
git push
```
