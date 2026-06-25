# ReportAutomation

Current version: `0.0.3`

`ReportAutomation`은 엑셀 집계표를 기반으로 조사 보고서 작성용 산출물을 자동 생성하는 런처 기반 프로젝트입니다.

현재 알파 버전은 HWPX/PPTX 완성 보고서를 바로 만드는 단계가 아니라, 안정적인 중간 산출물과 사전검증을 만드는 단계입니다. 목표 흐름은 다음과 같습니다.

```text
Excel 집계표
  -> Excel 산출 시트
  -> report_package.json
  -> preflight_report.json
  -> HWPX/PPTX 초본 보고서
```

## 프로젝트 구성

### `report_automation_launcher`

사용자가 직접 실행하는 Windows WinForms 런처입니다.

주요 역할:

- 집계표 Excel 파일 선택
- Excel add-in 경로 선택
- 표 목록과 가로배너 목록 미리보기
- 분석에 사용할 배너 선택, 순서 이동, 제외
- 보고서 유형, 문체, 소수점 자리, 차트/표 삽입 방식 선택
- HWPX/PPTX 템플릿 선택, 검사, 자동 보정
- Excel 산출 실행
- 문장 초안 TXT 미리보기와 문장별 검토
- `report_package.json`, `preflight_report.json` 생성 결과 표시

주요 문서:

- `report_automation_launcher/README.md`
- `report_automation_launcher/ALPHA.md`
- `report_automation_launcher/NEXT_DOCUMENT_OUTPUT_PLAN.md`

### `report_automation_addin`

런처가 호출하는 Excel VBA 산출 엔진입니다.

사용자가 직접 매크로를 조작하는 구조가 아니라, 런처가 Excel COM으로 add-in을 열고 실행합니다.

생성하는 주요 시트:

- `보고서_분석문`
- `보고서_차트데이터`
- `보고서_삽입표`
- `보고서_QA`
- `보고서_출처`
- `보고서_수정이력`
- `보고서_메타`

역할:

- 원본 집계표에서 표 블록 탐지
- 전체 기준 주요 수치 추출
- 분석문 기본 문장 생성
- 차트용 데이터 정규화
- 보고서 삽입용 표 데이터 생성
- QA/출처/메타 정보 기록

### `report_automation_engine`

Python 기반 보조 엔진입니다.

주요 역할:

- Excel 산출 시트 기반 문장 초안 TXT 생성
- 원본 집계표 직접 분석 fallback
- 문체/표 해석 설정 JSON 로드
- HWPX 구조 분석 보조
- 템플릿 검사, 기본 템플릿 생성, 자동 보정
- `report_package.json`과 `preflight_report.json` 생성

주요 모듈:

- `excel_report_generator.py`: 문장 초안 TXT 생성
- `report_package.py`: Excel 산출 시트를 중간 JSON 계약으로 변환하고 preflight 수행
- `template_inspector.py`: HWPX/PPTX 템플릿 placeholder 검사
- `template_factory.py`: 기본 HWPX/PPTX 템플릿 생성
- `template_autofix.py`: 원본 보존 방식의 템플릿 자동 보정
- `hwpx_report_writer.py`: HWPX 분석 보조

### `old/legacy`

더 이상 현재 런처 기반 워크플로우에서 사용하지 않는 파일을 보관하는 archive 영역입니다.

원칙:

- 대체 구현이 커밋된 뒤에만 이동
- 원래 경로, retired version, 대체 기능을 README나 메모로 남김
- 현재 실행/빌드/배포 경로에서는 참조하지 않음

## 현재 알파에서 가능한 기능

- Excel 집계표 선택 및 작업 복사본 생성
- 탐지된 표 목록 확인
- 탐지된 가로배너 목록 확인 및 선택
- Excel 산출 시트 생성
- Python 문장 초안 TXT 생성
- 문장별 수정/복사/검토본 저장
- QA 경고 확인
- HWPX/PPTX 템플릿 검사
- 기본 HWPX/PPTX/차트 검토 PPTX 템플릿 생성
- 템플릿 자동 보정 사본 생성
- `report_package.json` 생성
- `preflight_report.json` 생성
- 문서 생성 준비 상태 표시

## 아직 개발 중인 기능

다음 기능은 계획과 기반 작업은 있으나, 현재 알파에서 완성 기능으로 열지 않습니다.

- HWPX 보고서 직접 생성
- PPTX 보고서 직접 생성
- 차트 검토 PPTX 직접 생성
- HWP 바이너리 직접 편집
- HWPX/PPTX 템플릿에 본문, 표, 차트를 완전 자동 삽입
- PowerPoint 편집 가능한 차트 객체 생성

다음 개발 계획은 `report_automation_launcher/NEXT_DOCUMENT_OUTPUT_PLAN.md`를 기준으로 진행합니다.

## 핵심 산출물

### Excel 산출 시트

VBA add-in이 생성하는 1차 산출물입니다. 사람이 Excel에서 직접 검토할 수 있고, Python 엔진이 후속 JSON/package 생성에 사용합니다.

### `report_package.json`

HWPX/PPTX writer가 직접 읽을 중간 데이터 계약입니다.

포함 데이터:

- `sections`: 표 단위 본문 블록
- `tables`: 삽입용 집계표 데이터
- `charts`: 차트 후보 데이터
- `qa`: QA 메시지
- `meta`: 원본 파일, 생성 시각, 보고서 유형, 문체, 배너 등

### `preflight_report.json`

문서 생성 전 검증 리포트입니다.

상태값:

- `ready`: 문서 생성 가능
- `ready_with_warnings`: 생성 가능하지만 검토 필요
- `blocked`: HWPX/PPTX 생성 차단

차단 예:

- 최종 문장 없음
- `table_key` 누락 또는 중복
- 차트 값이 숫자가 아님
- 필수 템플릿 placeholder 없음
- HWP 바이너리처럼 v1에서 구조 분석 불가

## 템플릿 기준

### 최소 HWPX 템플릿

한글 문서 본문이 들어갈 위치에 다음 placeholder가 있어야 합니다.

```text
{{BODY}}
```

### 최소 PPTX 보고서 템플릿

반복 슬라이드에 다음 placeholder가 있어야 합니다.

```text
{{SECTION_TITLE}}
{{NARRATIVE}}
{{TABLE}}
{{CHART}}
```

### 최소 차트 검토 PPTX 템플릿

반복 슬라이드에 다음 placeholder가 있어야 합니다.

```text
{{CHART_TITLE}}
{{CHART}}
{{CHART_NOTE}}
```

`RA_`로 시작하는 shape 이름은 자동화용 식별자로 취급합니다.

## 실행 파일

런처 빌드 산출물:

```text
report_automation_launcher/bin/ReportAutomationLauncher.exe
```

빌드:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\report_automation_launcher\scripts\build_report_automation_launcher.ps1
```

## CLI 예시

문장 초안 생성:

```powershell
python report_automation_engine\excel_report_generator.py `
  --excel "C:\path\table.xlsx" `
  --config "report_automation_engine\config\default_style_schema.json" `
  --output "C:\path\table_draft.txt"
```

Report package/preflight 생성:

```powershell
python -m report_automation_engine.report_package `
  --excel "C:\path\table_report_alpha.xlsx" `
  --package-output "C:\path\report_package.json" `
  --preflight-output "C:\path\preflight_report.json"
```

템플릿 검사:

```powershell
python -m report_automation_engine.template_inspector `
  --template "C:\path\template.pptx" `
  --type pptx_report `
  --output "C:\path\template_report.json"
```

## 버전 관리

현재 버전은 `VERSION` 파일에 기록합니다.

정책:

- semantic versioning 사용: `MAJOR.MINOR.PATCH`
- push/publish 요청 시 버전 확인
- 코드, 문서, 빌드 산출물이 바뀌면 필요한 경우 버전 증가
- 릴리스 기준점은 Git tag로 기록: `v0.0.1`, `v0.0.2`, ...

자세한 정책은 `VERSIONING.md`를 참고합니다.

## 개발 방향

현재 안정화 우선순위:

1. Excel 산출 시트 안정화
2. `report_package.json` 계약 안정화
3. `preflight_report.json` 차단/경고 정확도 개선
4. 차트 검토 PPTX writer
5. PPTX 보고서 writer
6. HWPX 보고서 writer
