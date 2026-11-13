# HWPX Survey Chart EMF Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox `- [ ]` syntax for tracking.

**Goal:** 선택한 조사 차트를 한글 보고서에 메타파일로 넣되, 기존 표·본문과 중단 후 재개를 보호한다.

**Architecture:** [차트 검토 계획](2026-11-13-survey-chart-review.md)의 동일한 `chart_plan.json`을 재사용한다. Excel COM으로 한 번의 가짜 차트를 그려 메타파일 경로를 먼저 실증하고, 검증된 경우에만 한글 writer에 연결한다. 네이티브 HWP 차트는 만들지 않는다.

**Tech Stack:** Windows, Python 3.12, pywin32, Microsoft Excel, 아래한글, HWPX ZIP/XML, C# WinForms, pytest.

**Spec:** [조사 보고서 차트 출력 명세](../specs/2026-11-13-survey-chart-output.md)

## Global Constraints

- Windows, 기존 Python 3.12·WinForms·`python-pptx`·`pywin32`, 신규 외부 의존성 없음. PowerPoint 설치는 한글 경로의 필수 조건으로 만들지 않는다.
- 한글에는 확인된 EMF만 붙여넣기. PNG/JPEG, OLE 및 형식 미확인 붙여넣기로 대체 금지. 표 경로는 그대로 유지.
- 선택된 차트만 `표 → 차트 → 출처` 순으로, 한글 본문 폭 이내의 독립 문단에 삽입한다. 기존 본문·표 실행에 차트 옵션을 자동 적용하지 않는다.
- 원본 Excel·HWPX·PPTX 및 기존 로컬 진단 로그는 수정하거나 커밋하지 않는다. 시험용 산출물은 `outputs/`에 둔다.

## Review Focus

- 복사 성공으로 보이지만 저장 후 EMF가 PNG/OLE로 바뀐다 (Task 1).
- 클립보드가 잠기거나 선택 붙여넣기가 대화상자를 띄우면 무한 대기하지 않는다 (Task 1).
- 실패한 삽입이 유령 객체나 추가 문단을 남기면 안전한 저장점으로 돌아간다 (Task 2).
- 출력 양식 본문 폭이 좁거나 한 문항이 여러 표로 분할되어도 차트가 넘치거나 반복되지 않는다 (Task 2).
- 중간 재개 시 차트가 중복 삽입되지 않고, 설정 변경 시 예전 체크포인트를 사용하지 않는다 (Task 2, 3).

---

### Task 1: Excel → 한글 EMF 실증

**Files:** Create `report_automation_engine/hwp_chart_emf.py`, `tests/test_hwp_chart_emf.py`, `tests/probe_hwp_chart_emf.py`; Modify `report_automation_engine/README.md` (재현 절차).

**Interfaces:** `copy_chart_metafile(excel, chart: dict) -> None`, `paste_chart_metafile(hwp, body_width_hwpunit: int) -> dict`; 전자는 검증된 계획의 단일 차트를 소비한다. 후자의 성공 필드는 `width_hwpunit`, `insert_mode: "emf"`. `ChartEmfError(message: str, document_unchanged: bool)`에서 `True`는 붙여넣기 전 실패 또는 변경 없음이 확인된 경우만 쓴다. 증명 스크립트: `python tests/probe_hwp_chart_emf.py --out-dir outputs/chart_emf_probe`.

- [ ] **실패 테스트 작성:** `test_requires_enhanced_metafile_and_retries`는 클립보드 잠김이 최대 3회 재시도에 그치고 `CF_ENHMETAFILE`이 없으면 한글에 붙이지 않는지 확인. `test_unverified_paste_fails_closed`는 OLE/래스터/타임아웃/검증 불가를 성공으로 돌려주지 않는지 확인한다. COM 호출의 응답 없음은 동일 writer를 방치하지 않고 진단용 자식 프로세스의 30초 제한으로 실패를 기록한다.
- [ ] **RED:** `.venv/Scripts/python.exe -m pytest tests/test_hwp_chart_emf.py -q`에서 미구현 실패 확인.
- [ ] **구현:** Excel 작업용 임시 통합문서에서 실제 차트 객체를 만들고 Excel 그림 복사 후 클립보드 EMF 존재를 검사한다. 제공된 한글 API 문서를 참조해 대화상자 없이 메타파일을 명시적으로 선택할 수 있는 경로를 실증한다. 기본 붙여넣기나 확장자 이름만으로 형식을 인정하지 않는다. 서명 검사에서 EMF `iType=1`, 오프셋 40의 ` EMF`를 확인하고 HWPX 이미지 연결이 그 BinData를 가리키는지 검사한다. 실제 스크립트는 고정 가짜 값·한글 레이블을 사용하고 표와 본문은 추가하지 않는다.
- [ ] **GREEN 및 중단 기준:** pytest 통과 후 실증 스크립트에서 아래한글에 다시 열리는 1페이지 HWPX와 검사 JSON을 생성한다. `%` 한 자리, 최대값 강조, 폰트, 폭을 한글 캡처/PDF로 확인한다. 명시 선택 붙여넣기와 저장 후 EMF를 모두 증명하지 못하면 Task 2~3의 기능 노출을 보류하고 실패 이유를 기록한다.
- [ ] **커밋:** `git add report_automation_engine/hwp_chart_emf.py report_automation_engine/README.md tests/test_hwp_chart_emf.py tests/probe_hwp_chart_emf.py` 후 `git commit -m "Verify Excel chart EMF paste in Hancom"`.

### Task 2: 한글 작성 및 체크포인트 연동

**Files:** Modify `report_automation_engine/hwp_com_writer.py` (`write_hwp_document`, section 반복, CLI), `report_automation_engine/hwp_recovery.py`, `tests/test_hwp_recovery.py`; Create `tests/test_hwp_com_writer_charts.py`.

**Interfaces:** `write_hwp_document(..., chart_plan_path: str | Path | None = None) -> Path`, CLI `--chart-plan`. 기존 writer report에 `chart_results[]`(`chart_id`, `table_key`, `status: inserted|skipped|failed`, `insert_mode`, `reason`)를 추가. Task 1 함수 및 검증된 `chart_plan`을 사용한다. 계획 없이 실행하면 기존 `charts_deferred` 동작 유지.

- [ ] **실패 테스트 작성:** `test_insert_after_last_table_part`에서 가로 분할표 3개 뒤 차트 1개, 그 뒤 출처와 본문 폭 준수 검사. `test_failed_paste_cannot_damage_saved_document`에서 `document_unchanged=True`이면 건너뛰고 리포트 기록, `False`이면 현재 구간 중단 및 마지막 안전 저장점 보존. `test_resume_does_not_duplicate_chart`에서 2개 중 첫 문항 저장 뒤 중단·재개 시 차트가 총 2개이고, 계획 해시가 바뀌면 기존 저장점을 사용하지 않는지 확인한다.
- [ ] **RED:** `.venv/Scripts/python.exe -m pytest tests/test_hwp_com_writer_charts.py tests/test_hwp_recovery.py -q`에서 미구현 실패 확인.
- [ ] **구현:** `validate_chart_plan`과 `preflight.package_sha256 == plan.package_sha256`로 일치와 오류 여부를 COM 실행 전 확인한다. 이미 연 Excel COM 인스턴스가 있으면 그 인스턴스의 임시 통합문서를 사용하며, 기존 표가 Excel을 열지 않은 경우 차트 실행에서만 지연 초기화한다. 표 분할이 끝난 후 문항별로 차트를 한 번 삽입하고 출처를 쓴 뒤 체크포인트를 남긴다. 차트 계획 해시를 재개 입력에 포함하고 성공 여부와 각 단계 실패를 writer report에 남긴다. 저장된 HWPX에서 차트 EMF 연결을 재확인하고 불일치 결과를 성공 파일로 확정하지 않는다.
- [ ] **GREEN:** 같은 pytest 명령과 `tests/test_hwp_com_writer_tables.py` 통과. 가짜 가로 분할표를 넣은 한글 보고서에서 표·병합·표시값 불변과 차트 저장·재개를 확인한다.
- [ ] **커밋:** `git add report_automation_engine/hwp_com_writer.py report_automation_engine/hwp_recovery.py tests/test_hwp_recovery.py tests/test_hwp_com_writer_charts.py` 후 `git commit -m "Insert verified EMF charts into HWPX reports"`.

### Task 3: 런처 선택 연결과 실사용 검증

**Files:** Modify `report_automation_launcher/src/ReportAutomationLauncher.cs` (`LauncherOptions`, `BuildHwpWriterArguments`, `HwpJobPath`), `report_automation_launcher/src/SurveyChartReviewDialog.cs`, `tests/test_standalone_launcher.ps1`, `README.md`.

**Interfaces:** 검토 화면에서 검증된 `chart_plan.json`을 확정해 HWPX 생성 선택에 연결한다. `LauncherOptions.HwpChartPlanPath`는 미선택 시 `null`; `BuildHwpWriterArguments`에서만 `--chart-plan`으로 전달. `HwpJobPath`는 활성 계획의 SHA-256을 포함한다.

- [ ] **실패 테스트 작성:** 런처 자체검사에서 기본 꺼짐, 계획 누락/오류 시 차트 삽입 비활성, 새 계획 선택 시 HWP 작업 식별값 변경, 기존 본문·표 실행에서 CLI 인자 미전달을 검사한다.
- [ ] **RED:** 개발 빌드 자체검사에서 해당 차트 옵션이 연결되지 않아 실패하는지 확인한다.
- [ ] **구현:** `한글 보고서에 선택한 차트 삽입`을 명시 옵션으로 추가한다. 기존 출력 폴더와 보고서 저장 흐름은 재사용한다. 결과에서 `삽입/건너뜀/실패` 건수와 문항별 사유를 열 수 있게 하고 차트 실패를 무조건 성공으로 표시하지 않는다.
- [ ] **GREEN:** 독립 실행 테스트, `.venv/Scripts/python.exe -m pytest -q` 통과. 한글에서 가짜 4문항과 사용이 허용된 실제 집계표 1~3문항을 다시 열고 EMF, 표 구조, 한글 폰트, 페이지 폭을 확인한다. 차트 OFF, 중단 후 재개, 강제 클립보드 실패도 확인한다. 전체 KISDI 산출은 별도 승인 전 실행하지 않는다.
- [ ] **커밋:** `git add report_automation_launcher/src/ReportAutomationLauncher.cs report_automation_launcher/src/SurveyChartReviewDialog.cs tests/test_standalone_launcher.ps1 README.md` 후 `git commit -m "Enable opt-in HWPX chart insertion"`. 통합 전 배포 버전과 단일 EXE를 갱신하고 차트/비차트 검증 결과를 PR에 첨부한다.
