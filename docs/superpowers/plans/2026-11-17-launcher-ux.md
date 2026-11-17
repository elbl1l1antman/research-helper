# Research Helper Launcher UX Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 첫 실행부터 한글 보고서가 주 작업임을 알 수 있고, 준비 사항·진행 상태·다음 행동을 해당 화면에서 확인한다.

**Architecture:** WinForms `MainForm`의 탭, 기존 컨트롤, 핸들러와 `LauncherUi`를 재사용한다. 화면 구조 변경과 출력 기능 변경을 섞지 않는다.

**Tech Stack:** C# WinForms, 기존 PowerShell 빌드, 독립 실행 EXE 자체 검사, Windows UI 검수.

**Spec:** `docs/superpowers/specs/2026-11-17-launcher-ux.md`

## Global Constraints

- 제안 상태: 아래 C#은 적용할 구조와 테스트의 예시다. 현재 실행 파일은 변경하지 않는다.
- 차트 브랜치의 동작을 기준선으로 삼고, 실제 구현은 그 브랜치가 포함된 기준점에서 별도 `codex/launcher-ux` 브랜치로 진행한다. 기준 브랜치에 변경이 들어오면 첫 구현 전 `VERSION`과 UI 상태를 다시 확인한다.
- 대상은 `report_automation_launcher/src/ReportAutomationLauncher.cs`, `tests/test_standalone_launcher.ps1`, `docs/launcher_ui_design_system.md`다. `StandaloneBundle.cs`의 검사 진입점, Excel/VBA, Python writer, COM 설정, 기존 데이터/경로 계약은 최대한 그대로 둔다.
- 먼저 현재 동작에서 실패하는 assertion을 추가하고, 변경 후 성공을 확인한다. 작업 단위별 diff를 리뷰한다. 커밋을 쪼개되 원격에 올릴 때 저장소 `VERSIONING.md`를 적용한다.

## Review Focus

1. 화면 문자열을 바꿀 때 `ReadOptions()`, CLI의 `--output`, QA 분기와 재개 키까지 달라지는 회귀. 작업 1에서 같은 계약값을 검증한다.
2. 원본을 연달아 바꾸거나 양식을 고칠 때 이전 파일의 성공 상태가 남는 경쟁 조건. 작업 2에서 A→B 역순 완료와 직접 경로 편집을 검증한다.
3. 차트를 끄거나 관리를 닫았을 때 서식 선택이 초기화되거나 기본 자산이 덮어써지는 문제. 작업 3에서 기본/개인/외부 파일과 한글 실행 옵션을 검사한다.
4. 한글 작성 실패가 완료로 보이거나 자동 재개 기록을 잃는 문제. 작업 4에서 이전 실행의 산출물을 다음 실행의 산출물로 잘못 표시하지 않는지 검증한다.
5. 최소 창/고배율에서 액션이나 포커스가 잘리는 문제. 작업 5의 실제 화면 검사까지 통과해야 한다.

## Task 1: 작업 목적과 주 동작

**Files:** `report_automation_launcher/src/ReportAutomationLauncher.cs` (`MainForm` 생성자, `BuildFileGroup`, `BuildOutputGroup`, `ReadOptions`, `UpdateWorkflowStatus`, `CheckWorkflowLayout`).

- [ ] RED: `CheckWorkflowLayout()`에서 `한글 보고서(.hwpx)`가 원본 탭에 있고, `PowerPoint 보고서`는 메인 출력 선택에 없으며, 한글과 Excel 선택 시 각각 기존 `OutputType` 계약값을 보낸다는 검사를 추가한다. 기관 대시보드에서는 보고서 실행 버튼이 숨겨지는지도 확인한다.
- [ ] GREEN: 기존 `outputTypeCombo`를 원본 탭으로 옮긴다. HWPX를 기본 목적에 두고 변경한 표시명은 엔진 값과 분리한다. `BuildOutputGroup()`의 한글 전용 설정은 Excel 모드에서 숨긴다. PPT 양식 관리가 필요하면 별도 도구로 유지한다. 탭/콤보 선택에 숨어 있는 인덱스 비교는 `rg 'outputTypeCombo|workflowTabs.SelectedIndex'`로 찾아 두 출력 종류와 이름 있는 탭 참조로 치환한다.
- [ ] GREEN: 기존 `runButton`은 실제 생성에만 사용하고, 앞 단계에는 필요한 경우 `표·배너 확인`, `보고서 설정으로`를 표시한다. `RunButton_Click`의 최종 옵션 검증과 자동 양식 검사를 삭제하지 않는다. `첫 문항 / 처음 3문항 / 전체 문항`은 기존 선택을 유지한 채 한글 모드의 산출 요약에 노출한다.

```csharp
// ComboBox items and ReadOptions keep the existing output contract.
outputTypeCombo.DisplayMember = "Value";
outputTypeCombo.Items.Add(new KeyValuePair<string, string>("HWPX 보고서", "한글 보고서(.hwpx)"));
outputTypeCombo.Items.Add(new KeyValuePair<string, string>("Excel 산출 시트", "Excel 산출 시트"));
outputTypeCombo.SelectedIndex = 0;
```

`MainForm.SelectedOutputType`을 한 곳에 둔다. `SelectedItem`이 위 타입일 때 `Key`를 사용하고 초기 값이 없으면 빈 문자열을 반환한다. `ReadOptions`, 준비 상태, `Refresh...` 핸들러가 이 값을 공유한다. `LauncherOptions`와 CLI의 기존 Excel 기본값은 바꾸지 않는다.

- [ ] 검증: 자체 검사 외 원본을 선택하지 않은 상태에서 한글/Excel을 각각 고른다. 출력 폴더와 현재 저장 방식이 어디에 보이는지 확인한다. UI와 `ReadOptions()`의 목적이 일치하면 커밋한다.

## Task 2: 표·배너와 준비 상태

**Files:** `ReportAutomationLauncher.cs` (`BrowseWorkbook`, `LoadWorkbookPreviewAsync`, `BuildDataReviewGroup`, `BuildBannerPanel`, `UpdateWorkflowStatus`, `UpdateReadinessChecklist`, `ResetTemplateStatus`, `InspectTemplate`, `EngineRunner.RequireHwpTemplate`).

- [ ] RED: 양식 미검사, 지원하지 않는 `.hwp`, 검사 경고, 필수 필드 누락, 검사 후 경로 변경에 따라 준비 상태가 달라지는 검사를 추가한다. 탐지 중인 원본을 다른 경로로 바꿀 때 이전 결과가 적용되지 않아야 한다.
- [ ] GREEN: 현재 로딩 요청 토큰(같은 경로 재탐지 포함)과 UI에 표시할 원본 경로를 캡처한다. 성공·실패·finally 모두 현재 토큰과 일치할 때만 UI를 갱신한다. 원본이 같을 때 새로고침은 기존 규칙을 유지하고, 다른 원본을 고르면 기존 표 선택과 범위 override를 새 파일에 적용하지 않는다. 자동 탐지 실패의 안내는 살려 두되 기존 수동 지정/기존 산출 시트 사용까지 차단하지 않는다.
- [ ] GREEN: 메인 양식 검사 상태는 경로와 함께 보관한다. 경로 편집·양식 보정 사본 적용 시 검사 상태를 초기화한다. 양식이 미검사면 현재 런처의 자동 검사를 실행하도록 표시하고, 알려진 필수 필드 오류는 양식 선택/재검사로 유도한다. `EngineRunner.RequireHwpTemplate()`의 쓰기 전 검사와 `RequireHwpPreflight()`는 그대로 남긴다.

```csharp
// Apply the same guard in every workbook-preview UI callback.
int request = ++workbookPreviewRequest;
string path = workbookPathText.Text.Trim();
// ... background BannerInspector.ReadPreview(path), then on UI thread:
if (request != workbookPreviewRequest ||
    !string.Equals(path, workbookPathText.Text.Trim(), StringComparison.OrdinalIgnoreCase)) return;
// Apply preview tables, banners, and final state together.
```

- [ ] 검증: 가짜 표 두 파일을 연속 선택한다. 탐지된 표 수, 실제 분석 배너, 원본 범위·보정 범위가 서로 혼용되지 않는지 확인한다. `선택 배너 없음(전체 기준)`과 `탐지 실패`가 구별되면 커밋한다.

## Task 3: 차트 선택과 관리 분리

**Files:** `ReportAutomationLauncher.cs` (`BuildSurveyChartStyleGroup`, `RefreshSurveyChartStyle`, `ActiveSurveyChartStyle`, `BuildRulesPage`, `ReadOptions`). 미리보기 파일이 필요할 때만 기본 테마와 함께 `templates/` 및 standalone 패키징 목록을 갱신한다.

- [ ] RED: 차트 미포함일 때 선택 UI는 보이지 않지만, 다시 포함하면 기존 서식 선택이 살아 있고 `ReadOptions().HwpChartStylePath`가 두 기본 테마와 개인 서식을 그대로 가리키는지 검사한다.
- [ ] GREEN: 한글 설정의 `표·그림` 화면에 `차트 서식` 선택과 미리보기를 둔다. 기존 관리 핸들러를 `차트 서식 관리…` 창에 배치한다. 관리창을 다시 열어도 선택 이벤트가 중복 구독되지 않고, 등록·삭제가 끝나기 전에 관리창을 닫을 경우에는 기존 명령 완료/오류 처리 흐름이 유실되지 않게 한다. UI를 가리는 것만으로 `SelectedIndex`를 바꾸지 않는다.
- [ ] 검증: 기본 테마 두 개의 서식과 실제 차트 모양이 구별되는지, 개인 등록과 삭제 확인·취소가 작동하는지, 원본 PPTX는 그대로인지 확인한다. 미리보기는 원본 업무 자료를 포함하지 않는다.

## Task 4: 실행 중·성공·검토·실패 구분

**Files:** `ReportAutomationLauncher.cs` (`BuildRunGroup`, `RunButton_Click`, `Log`, `PopulateResultFiles`, `EngineRunner.RunWorkflowCore`).

- [ ] RED: 한글 파일 없음, 파일은 있으나 차트 교체 보류, 정상 저장, QA 차단, 재개 각각에서 대표 문구·열기 버튼·진단 정보가 맞는지 fixture 검사한다. 이전 실행의 경로가 현재 결과 버튼에 남으면 실패시킨다.
- [ ] GREEN: `산출 전 확인`은 설정에서 보여주고 결과에는 현재 진행·저장 위치·최근 결과를 둔다. `원본→Excel→QA→HWPX`의 상태를 자연어 로그를 파싱하지 않고 기존 각 단계 호출 위치에서 UI에 알린다. 기존 `EngineRunner.RunWorkflow`의 비 UI 호출에는 상태 전달을 생략할 수 있게 한다. `재개` 시 건너뛴 단계는 완료로 꾸미지 않고 `이전 산출물 사용`으로 표시한다. 임의의 진행률이나 취소 버튼을 새로 만들지 않는다.
- [ ] GREEN: 결과물은 `한글 보고서 열기` 또는 `산출 Excel 열기`를 주 동작으로 두고, 초안·로그·QA 기록은 별도 묶음으로 유지한다. 성공 확인 팝업을 대체하는 경우 차트 검토 경고와 오류 상태도 요약 영역에 지속해서 노출한다. `BuildResultSummary(...).Contains("수동 교체:")`에 UI의 경고 판정을 의존시키지 않는다. 저장된 writer report의 기존 결과 항목과 실제 파일 존재 여부를 사용한다.
- [ ] 검증: 실행 중 중복 클릭, 실패 후 로그 열기, 재시도, 수치 편집 검토 PPTX → EMF 저장 흐름을 확인한다. 편집한 PPTX와 report package가 외부 파일인 경로도 회귀 검사한다.

## Task 5: 네이티브 레이아웃과 인수 확인

**Files:** `ReportAutomationLauncher.cs` (`LauncherUi`, `CheckWorkflowLayout`), `tests/test_standalone_launcher.ps1`, `docs/launcher_ui_design_system.md`.

- [ ] 창 그리드와 간격을 `LauncherUi` 토큰으로 정리한다. 그룹과 포커스 순서를 맞추고, 편집 가능한 텍스트와 로그에 키보드로 접근할 수 있게 한다. 출력 폴더, 양식 오류, 선택 범위가 글자 잘림 없이 보이는지 확인한다.
- [ ] 아래 빌드 및 자체 검사를 수행한다. 기존 사용 EXE 대신 무시되는 산출 폴더에 테스트용 EXE를 빌드한다.

```powershell
.\report_automation_launcher\scripts\build_report_automation_launcher.ps1 -OutputDir "$PWD\outputs\launcher_ux_check"
.\tests\test_standalone_launcher.ps1 -ExePath "$PWD\outputs\launcher_ux_check\ReportAutomationLauncher.exe"
```

- [ ] Windows에서 최소 창·고배율·키보드만으로 원본 선택→한글 양식 확인→1문항 생성→파일 열기를 1회 진행한다. 다른 파일의 QA 실패→기록 열기→다시 실행과 Excel 전용 작업도 1회 확인한다. 한글 결과는 가짜 집계표를 사용해 양식, 표, 차트가 이전과 동일한지 확인한다. 실데이터로 병행 재시험할 필요는 없다.
- [ ] 완료 기준: 기본 한글 경로에서 지원하지 않는 옵션과 사용하지 않는 차트 관리 명령이 먼저 보이지 않는다. 결과에 실패 이유, 보존 파일, 다음 동작이 보이고, 로그 또는 QA가 없는 상태에서 누를 수 있는 것처럼 표시되지 않는다. 새 EXE 단독 실행이 통과한 후에만 교체한다.
