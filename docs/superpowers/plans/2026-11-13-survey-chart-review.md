# Survey Chart Review Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox `- [ ]` syntax for tracking.

**Goal:** 기존 집계표 산출 결과에서 사용자가 그래프를 선택·검토한 뒤 편집 가능한 차트 전용 PPTX를 만든다.

**Architecture:** 기존 package에 손대지 않고 차트 후보를 별도 계획 JSON으로 해석한다. 런처와 PPTX writer가 같은 계획을 읽어 미리보기와 실제 결과의 분류·순서·강조를 일치시킨다. 다음 단계의 한글 writer도 같은 계획을 사용한다.

**Tech Stack:** Windows, Python 3.12, python-pptx, C# WinForms, pytest, PowerShell.

**Spec:** [조사 보고서 차트 출력 명세](../specs/2026-11-13-survey-chart-output.md)

## Global Constraints

- Windows, 기존 Python 3.12·WinForms·`python-pptx`·`pywin32`, 신규 외부 의존성 없음.
- 원본 Excel·HWPX·PPTX 및 기존 로컬 진단 로그는 수정하거나 커밋하지 않는다. 시험용 산출물은 `outputs/`에 둔다.
- 초기 추천은 최대 10개지만 실제 내보내기는 사용자의 선택을 조용히 자르지 않는다. 차트 내부 제목 없음. 기본 글꼴 맑은 고딕, 기본색 `357B99`, 강조색 `C65B39`.
- `pptx_report`, 기관 대시보드, 한글 본문·표와 기존 CLI는 동작을 유지한다. PPTX 레이아웃 복제, 시계열/다중 시리즈는 별도 작업.

## Review Focus

- 같은 문항에 서로 다른 배너 또는 단위가 섞여도 합산하지 않는다 (Task 1).
- 일부 보기만 선택하거나 값이 우연히 100이 되어도 전체 구성으로 오인하지 않는다 (Task 1).
- 원본이나 선택이 바뀐 뒤 이전 차트 계획을 내보내지 않는다 (Task 1, 3).
- 양수·음수·0·동률과 정확한 한글/단위 표시를 보존한다 (Task 1, 2).
- 유효한 차트가 없거나 입력 package가 달라도 빈 PPT를 완료로 알리지 않는다 (Task 2, 3).

---

기준: `origin/main`의 `04468b2`. 구현 및 문서는 `codex/chart-development`에서 관리한다. 차트 계약은 한글 삽입과 공유하지만, 아래 작업은 한글 COM 없이 완료할 수 있다.

### Task 1: 후보 분리와 차트 계획

**Files:** Create `report_automation_engine/survey_chart_plan.py`, `tests/test_survey_chart_plan.py`; Modify `report_automation_engine/report_package.py` (차트 후보 QA만).

**Interfaces:** `package_digest(package: dict) -> str`; `build_chart_plan(package: dict, selection: dict | None = None) -> dict`; `validate_chart_plan(package: dict, plan: dict) -> None`; CLI `python -m report_automation_engine.survey_chart_plan --package P [--selection S] --output chart_plan.json`. `ValueError`는 계약 또는 식별값 불일치, 개별 차트 오류는 `issues`에 기록한다. 새 `preflight_report.json`에는 `package_sha256`을 추가한다.

**계약:** `chart_selection.json`은 `package_sha256`과 `charts[]`를 가진다. 각 설정은 `chart_id`, `enabled`, `selected_point_ids`, `kind`, `confirmed_binary_share`, `title`, `font`, `base_color`, `highlight_color`. `chart_plan.json`은 `schema_version: 1`, `package_sha256`, `charts[]`, `issues[]`. 각 차트는 원본 `table_key`, `series_group`, `measure`, 선택·확정 설정, `kind_resolved`, `source`(확인된 경우 `sheet`, `range`, `base_label`), `points[]`(`point_id`, `label`, `value`, `display_value`, `source_cell`, `highlight`)를 가진다. ID는 그룹 키 및 포인트의 원본 위치/순번을 정규 JSON으로 인코딩해 SHA-256으로 만든다. `package_sha256`은 정렬된 키를 가진 UTF-8 정규 JSON 해시이며, 잘못된 수치는 거부 또는 `null`+진단 후 해시한다. 계획 검증은 원본과 계획의 선택 설정으로 계획을 다시 계산해 비교하며 파생 값 수정을 허용하지 않는다.

- [ ] **실패 테스트 작성:** `test_groups_are_isolated`에서 T001의 서로 다른 배너 및 `%`/`명`이 별도 3개로 남는지 검사. `test_pie_requires_full_confirmed_distribution`에서 `[63.3, 36.7]` 기본 `column`, 이분형 확인 후 `pie`, `[63.3, 20.0, 16.7]` 중 2개 선택과 `[120, 150]` 매출의 원형 거부를 검사.
- [ ] **실패 테스트 추가:** `test_values_and_selection_are_preserved`에서 `-2, 0, 5, 5` 중 두 최대값 강조, `0` 보존, `-`·누락·NaN 오류, 11개를 직접 선택했을 때 11개 보존. `test_invalid_selection_and_changed_package`에서 중복 보기명, 다른 배너의 보기 ID, 존재하지 않는 ID, 변경된 값과 변조된 계획 거부. `test_chart_only_qa_does_not_block_table_only_output`에서 잘못된 후보 수치가 기존 본문·표만 생성하는 경로를 차단하지 않고 채택된 차트의 오류로 남는지 검사.
- [ ] **RED:** `.venv/Scripts/python.exe -m pytest tests/test_survey_chart_plan.py -q` 실행, 미구현 계약에서 실패 확인.
- [ ] **구현:** 위 함수와 CLI 구현. `include_chart`는 기본 체크 상태로만 사용하고, 원형 판정은 잘리기 전 전체 후보를 확인한다. 명시적 원형 부적격은 오류, `auto`는 막대로 결정한다. 선택된 항목의 값을 평균·정규화·자동 보정하지 않는다. `build_preflight`에서 `package_digest`를 기록하고, 차트 전용 QA를 본문·표 차단 조건과 분리하며, 채택 차트의 오류는 내보내기 단계에 귀속한다.
- [ ] **GREEN:** `.venv/Scripts/python.exe -m pytest tests/test_survey_chart_plan.py tests/test_report_package_table_matrix.py -q` 모두 통과.
- [ ] **커밋:** `git add report_automation_engine/survey_chart_plan.py report_automation_engine/report_package.py tests/test_survey_chart_plan.py` 후 `git commit -m "Define validated survey chart plans"`.

### Task 2: 편집 가능한 차트 검토 PPTX

**Files:** Modify `report_automation_engine/document_writer.py` (`write_document`, CLI); Create `tests/test_survey_chart_review.py`.

**Interfaces:** 기존 `write_document(..., template_path=None, chart_plan_path: str | Path | None = None) -> Path`에 마지막 옵션 추가, CLI `--chart-plan`. `chart_review`에서 옵션이 주어지면 Task 1 검증 후 새 경로를 사용한다. 옵션 없는 두 기존 모드는 유지한다.

- [ ] **실패 테스트 작성:** `test_editable_review_charts`는 네 가짜 문항(이분 응답, 5점 척도, 매출과 예상 매출, 두 배너에 같은 보기명)의 `pie`/`column` 타입과 배너별 개별 슬라이드, 네이티브 Chart 및 내장 워크북, 내부 제목 없음, `63.3%` 레이블, 지정 글꼴과 최대값 두 개 색상, 원본 파일 해시 불변을 검사한다. `test_empty_or_invalid_plan_does_not_save`는 오류/빈 선택/차단된 preflight/변경된 패키지/해시 없는 이전 preflight에서 출력 파일이 생기지 않는지 검사한다. `test_legacy_cli_without_plan`은 기존 출력 유지.
- [ ] **RED:** `.venv/Scripts/python.exe -m pytest tests/test_survey_chart_review.py -q` 실행, 미구현 옵션에서 실패 확인.
- [ ] **구현:** 새 경로는 `preflight.package_sha256 == plan.package_sha256`일 때만 출력한다. 이전 preflight는 같은 원본에서 재산출하도록 안내한다. 채택 차트당 한 슬라이드에 문항/배너 제목, 출처, 편집 가능한 `ChartData` 차트를 배치한다. 값 레이블은 PowerPoint에서 값 편집 후에도 맞도록 숫자 서식을 사용하고, 원본 표시문을 같은 방식으로 표현할 수 없으면 별도의 검토 목록에 원문과 단위를 표시한다. 축·범례·레이블에 선택 글꼴을 적용한다. 실패 시 기존 파일을 남기고 임시 결과는 정리한다.
- [ ] **GREEN:** 같은 pytest 명령 통과. 실제 PowerPoint에서 네 유형을 열고 차트 데이터 편집·레이블 갱신·한글 축 가독성을 확인한다. 결과는 `outputs/chart_review/`에 보관한다.
- [ ] **커밋:** `git add report_automation_engine/document_writer.py tests/test_survey_chart_review.py` 후 `git commit -m "Export editable survey chart review decks"`.

### Task 3: 런처 차트 검토와 배포 점검

**Files:** Create `report_automation_launcher/src/SurveyChartReviewDialog.cs`; Modify `report_automation_launcher/src/ReportAutomationLauncher.cs`, `report_automation_launcher/scripts/build_report_automation_launcher.ps1`, `tests/test_standalone_launcher.ps1`, `README.md`.

**Interfaces:** 기존 실행 결과의 `report_package.json`과 동일 실행의 `preflight_report.json` 경로를 사용한다. `SurveyChartReviewDialog(string packagePath, string preflightPath)`에서 `차트 검토 PPT 만들기` 실행. 엔진에는 Task 1/2 CLI를 호출한다.

- [ ] **실패 테스트 작성:** `tests/test_standalone_launcher.ps1`에 `--self-check-survey-chart-review`를 추가한다. 자체검사에서 package 없음은 버튼 비활성, 서로 다른 실행의 preflight는 안내 후 생성 안 됨, 선택/유형 변경 후 최종 미리보기와 생성 계획의 포인트 수·유형 일치, 차트 없는 선택은 생성 안 됨을 확인한다.
- [ ] **RED:** 개발 빌드로 `--self-check-survey-chart-review` 테스트를 실행해 진입점 미구현 실패 확인.
- [ ] **구현:** 결과 화면에 차트 검토 버튼을 추가한다. 문항/배너/지표 목록과 보기 체크, 유형·이분형 확인, 제목·글꼴·색, 데이터/유형/최댓값/출처 미리보기를 같은 창에서 제공한다. Python 실행은 기존 비동기 호출 흐름을 사용한다. `chart_selection.json` 저장, 재검증, 저장 경로 확인, 성공/실패 안내 순서. 실제로 연결되지 않는 기존 차트 출력 콤보의 `PPTX 차트 객체` 문구는 중복 동작으로 남기지 않는다.
- [ ] **GREEN:** `tests/test_standalone_launcher.ps1 -ExePath <검증 빌드>` 및 `.venv/Scripts/python.exe -m pytest -q` 통과. GUI에서 엑셀 산출 완료 → 검토 버튼 → 선택 변경 → PPTX 열기까지 직접 확인한다.
- [ ] **커밋:** `git add report_automation_launcher/src report_automation_launcher/scripts/build_report_automation_launcher.ps1 tests/test_standalone_launcher.ps1 README.md` 후 `git commit -m "Connect chart review to the report launcher"`. 차트 검토를 별도 배포할 경우 버전과 단일 EXE는 검증 후 갱신한다.
