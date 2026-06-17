# Report Automation Launcher

보고서 자동화 프로젝트의 주 실행 프로그램입니다.

알파 단계부터 사용자는 Excel add-in을 직접 조작하지 않고, 이 런처에서 파일 등록 → 데이터 확인 → 작성 규칙 선택 → 실행/결과 확인 흐름으로 작업합니다. VBA add-in은 런처가 내부적으로 호출하는 Excel 산출 엔진으로 유지합니다.

## 실행 파일

- `bin/ReportAutomationLauncher.exe`

GUI에서 집계표 Excel 파일과 자동화 추가기능(`.xlam`)을 선택하면 표 목록과 배너 목록을 먼저 확인할 수 있습니다. 이후 실행하면 선택한 통합문서에 보고서 산출 시트를 생성하고, 선택 시 Python 보조 엔진으로 문장 초안 TXT도 함께 생성합니다. 생성된 초안은 결과 탭에서 미리보기, 문장별 수정, QA 경고 확인, 열기, 복사, 검토본 저장을 할 수 있습니다.

현재 알파에서 실제 지원하는 출력은 `Excel 산출 시트`와 `문장 초안 TXT`입니다. HWP/HWPX 및 PowerPoint 출력은 다음 알파 단계에서 활성화합니다.

## 기본 동작

- 기본값으로 원본 집계표 옆에 작업 복사본을 만들고 복사본에 산출 시트를 생성합니다.
- 집계표 선택 후 탐지된 표 목록과 배너 목록을 미리 확인합니다.
- 배너 목록은 엑셀 집계표에서 발견된 가로배너를 기준으로 만들고, 추천 선택/전체 선택/선택 해제/순서 이동/목록 삭제를 지원합니다.
- 보고서 유형과 문체 선택값을 실행 설정 파일에 기록합니다.
- `추출 배너 목록`은 선택한 배너 순서대로 자동 반영되며, `제목 제거 접두어`는 필요할 때만 사용자가 입력합니다.
- 기본값으로 `문장 초안 TXT(Python)`를 생성합니다. VBA가 만든 `보고서_분석문*` 시트가 있으면 그 내용을 우선 사용하고, 없으면 원본 표 블록을 직접 분석합니다.
- 실행 완료 후 결과 탭에서 산출 엑셀 열기, 초안 TXT 열기, 초안 미리보기 복사, 문장별 수정, 검토본 저장을 할 수 있습니다.
- QA 경고 탭에서는 출처 없음, 제목 없음, 문장 짧음, 수치 없음, 종결 표현 확인 항목을 필터링해 검토할 수 있습니다.
- 실행 설정은 산출 파일 옆에 `*_launcher_config.txt`로 기록됩니다.
- `REPORT_AUTOMATION_ADDIN` 환경변수를 지정하면 기본 추가기능 경로로 사용합니다.
- `REPORT_AUTOMATION_PYTHON` 환경변수를 지정하면 문장 초안 생성에 사용할 Python 경로로 사용합니다.
- `REPORT_AUTOMATION_ENGINE` 환경변수를 지정하면 `excel_report_generator.py` 경로로 사용합니다.

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
- `--no-copy`: 원본 파일에 직접 산출
- `--keep-open`: 완료 후 Excel 창 유지
- `--no-draft`: Python 문장 초안 TXT 생성 생략

배너 목록 탐지만 확인할 때는 다음 옵션을 사용할 수 있습니다.

```powershell
.\report_automation_launcher\bin\ReportAutomationLauncher.exe `
  --list-banners `
  --workbook "C:\path\table.xlsx" `
  --out "$env:TEMP\banners.txt"
```
