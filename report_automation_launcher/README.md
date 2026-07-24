# Report Automation Launcher

보고서 자동화 프로젝트의 주 실행 프로그램입니다.

알파 단계부터 사용자는 Excel add-in을 직접 조작하지 않고, 이 런처에서 파일 등록 → 데이터 확인 → 작성 규칙 선택 → 실행/결과 확인 흐름으로 작업합니다. VBA add-in은 런처가 내부적으로 호출하는 Excel 산출 엔진으로 유지합니다.

## 실행 파일

- `bin/ReportAutomationLauncher.exe`

GUI에서 집계표 Excel 파일과 자동화 추가기능(`.xlam`)을 선택하면 표 목록과 배너 목록을 먼저 확인할 수 있습니다. 이후 실행하면 선택한 통합문서에 보고서 산출 시트를 생성하고, 선택 시 Python 보조 엔진으로 문장 초안 TXT도 함께 생성합니다. 생성된 초안은 결과 탭에서 미리보기, 문장별 수정, QA 경고 확인, 열기, 복사, 검토본 저장을 할 수 있습니다.

현재 알파에서 실제 실행되는 출력은 `Excel 산출 시트`, `문장 초안 TXT`, `report_package.json`, `preflight_report.json`, HWPX 초본입니다. HWPX 초본은 Windows에서 아래한글 COM 객체를 사용할 수 있을 때 생성하며, PowerPoint 보고서 출력은 아직 선택/검증 중심입니다.

## 기본 동작

- 기본값으로 원본 집계표 옆에 작업 복사본을 만들고 복사본에 산출 시트를 생성합니다.
- 집계표 선택 후 탐지된 표 목록과 배너 목록을 미리 확인합니다.
- 배너 목록은 엑셀 집계표에서 발견된 가로배너를 기준으로 만들고, 추천 선택/전체 선택/선택 해제/순서 이동/목록 삭제를 지원합니다.
- 보고서 유형과 문체 선택값을 실행 설정 파일에 기록합니다.
- `추출 배너 목록`은 선택한 배너 순서대로 자동 반영되며, `제목 제거 접두어`는 필요할 때만 사용자가 입력합니다.
- 출력 형식, 템플릿, 구성요소, 소수점 자리, 차트 출력, 삽입표 방식, LLM 제공자/모델/API 키 입력 여부를 GUI에서 선택할 수 있습니다.
- 템플릿 도구에서 HWPX/PPTX 템플릿을 검사하고, 기본 템플릿을 생성하고, 원본을 보존한 자동 보정 사본을 만들 수 있습니다.
- API 키 값은 실행 중 옵션으로만 들고 있으며, `*_launcher_config.txt`에는 키 자체가 아니라 입력 여부만 기록합니다.
- 기본값으로 `문장 초안 TXT(Python)`를 생성합니다. VBA가 만든 `보고서_분석문*` 시트가 있으면 그 내용을 우선 사용하고, 없으면 원본 표 블록을 직접 분석합니다.
- 실행 후 `report_package.json`과 `preflight_report.json`을 생성해 HWPX/PPTX 문서 생성 전 준비 상태를 `ready`, `ready_with_warnings`, `blocked`로 표시합니다.
- 실행 완료 후 결과 탭에서 산출 엑셀 열기, 초안 TXT 열기, HWPX 초본 열기, HWPX writer report 열기, 초안 미리보기 복사, 문장별 수정, 검토본 저장을 할 수 있습니다.
- HWPX 보고서 출력은 실행 전 아래한글 COM 환경을 먼저 점검하고, 실패 시 writer report에 실패 단계와 오류를 남깁니다.
- QA 경고 탭에서는 출처 없음, 제목 없음, 문장 짧음, 수치 없음, 종결 표현 확인 항목을 필터링해 검토할 수 있습니다.
- 실행 설정은 산출 파일 옆에 `*_launcher_config.txt`로 기록됩니다.
- `REPORT_AUTOMATION_ADDIN` 환경변수를 지정하면 기본 추가기능 경로로 사용합니다.
- `REPORT_AUTOMATION_PYTHON` 환경변수를 지정하면 문장 초안 생성에 사용할 Python 경로로 사용합니다.
- `REPORT_AUTOMATION_ENGINE` 환경변수를 지정하면 `excel_report_generator.py` 경로로 사용합니다.

## 템플릿 도구

산출 방식 영역의 버튼은 다음 역할을 합니다.

- `검사`: 선택한 HWPX/HWP 또는 PPTX 파일의 placeholder와 `RA_` shape 이름을 확인합니다.
- `기본 HWPX`: `report_template_basic.hwpx` 시작 템플릿을 생성합니다.
- `기본 PPTX`: `report_template_basic.pptx` 보고서 템플릿을 생성합니다.
- `차트 PPTX`: `chart_review_template_basic.pptx` 차트 검토용 템플릿을 생성합니다.
- `자동 보정`: 원본 템플릿을 덮어쓰지 않고 `_template_ready` 사본에 최소 placeholder를 삽입합니다.
- `가이드`: 사용자가 직접 템플릿을 만들 때 지켜야 할 최소 placeholder 기준을 엽니다.

사용자가 직접 만드는 최소 기준은 다음과 같습니다.

- HWPX: 본문 시작 위치에 `{{BODY}}` 한 줄 유지
- PPTX 보고서: 반복 슬라이드에 `{{SECTION_TITLE}}`, `{{NARRATIVE}}`, `{{TABLE}}`, `{{CHART}}` 유지
- 차트 검토 PPTX: 반복 슬라이드에 `{{CHART_TITLE}}`, `{{CHART}}`, `{{CHART_NOTE}}` 유지

디자인, 글꼴, 색상, 로고, 배경, 마스터 슬라이드, placeholder 위치와 크기는 수정해도 됩니다. 중괄호 placeholder 텍스트와 `RA_`로 시작하는 shape/bookmark 이름은 유지해야 합니다.

## 빌드

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\report_automation_launcher\scripts\build_report_automation_launcher.ps1
```

빌드에는 Windows .NET Framework의 `csc.exe`를 사용합니다.

## CLI 검증

GUI 없이 자동화 호출만 확인할 때 사용할 수 있습니다.

```powershell
.\report_automation_launcher\bin\ReportAutomationLauncher.exe `
  --workbook "C:\path\table.xlsx" `
  --addin ".\report_automation_addin\dev\ReportAutomationAddin_dev.xlam" `
  --no-copy
```

사용 가능한 옵션:

- `--workbook <path>`: 집계표 Excel 파일
- `--addin <path>`: 보고서 자동화 추가기능
- `--banner <text>`: 추출 배너 목록
- `--prefixes <text>`: 제목 제거 접두어
- `--decimal-places <0|1|2>`: 수치 소수점 자리
- `--chart-output <text>`: 차트 출력 방식
- `--table-insert <text>`: 삽입표 방식
- `--use-llm`: LLM 문장 고도화 사용
- `--llm-provider <text>`: LLM 제공자
- `--llm-model <text>`: LLM 모델명
- `--no-copy`: 원본 파일에 직접 산출
- `--keep-open`: 완료 후 Excel 창 유지
- `--no-draft`: Python 문장 초안 TXT 생성 생략
- `--output-type "HWPX 보고서"`: Excel 산출 후 HWPX 초본 생성
- `--hwp-template <path>`: HWPX 보고서 템플릿
- `--hwp-visible`: 아래한글 창 표시 상태로 HWPX 생성
- `--hwp-keep-open-on-error`: HWPX 생성 실패 시 열린 문서 유지

배너 목록 탐지만 확인할 때는 다음 옵션을 사용할 수 있습니다.

```powershell
.\report_automation_launcher\bin\ReportAutomationLauncher.exe `
  --list-banners `
  --workbook "C:\path\table.xlsx" `
  --out "$env:TEMP\banners.txt"
```
