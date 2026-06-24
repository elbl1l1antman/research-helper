# Next Alpha Plan: HWPX/PPTX Direct Report Output

## 목표

다음 알파의 목표는 `report_package.json`과 `preflight_report.json`을 기준으로 HWPX/PPTX 초본 보고서를 직접 생성하는 것이다. Excel 산출 시트를 다시 해석하지 않고, 이미 검증된 package만 writer 입력으로 사용한다.

핵심 원칙:

- Excel → `report_package.json` → HWPX/PPTX writer 순서로 고정한다.
- `preflight_report.json`이 `blocked`이면 문서 생성 버튼을 비활성화한다.
- writer는 원본 템플릿을 덮어쓰지 않고 새 파일만 만든다.
- v1은 완성본이 아니라 “편집 가능한 초본”을 만든다.

## 산출 기능

### HWPX 보고서 생성

입력:

- `report_package.json`
- `preflight_report.json`
- HWPX 템플릿 경로

출력:

- `<원본파일명>_draft.hwpx`

기능:

- 템플릿 사본 생성
- `{{REPORT_TITLE}}`, `{{PROJECT_NAME}}`, `{{GENERATED_AT}}` 치환
- `{{BODY}}` 위치에 section 반복 삽입
- section 단위로 `SECTION_TITLE`, `NARRATIVE`, `TABLE`, `CHART`, `SOURCE` 블록 생성
- 표는 HWPX XML 표 삽입이 안정화되기 전까지 텍스트 표 또는 탭 구분 표로 삽입
- 차트는 v1에서 이미지/렌더링 삽입하지 않고 chart placeholder와 데이터 출처를 남김

제외:

- HWP 바이너리 직접 편집
- 한글 COM 기반 실시간 편집
- 차트 이미지 삽입

### PPTX 보고서 생성

입력:

- `report_package.json`
- `preflight_report.json`
- PPTX 보고서 템플릿 경로

출력:

- `<원본파일명>_draft.pptx`

기능:

- 템플릿 사본 생성
- 제목 슬라이드 placeholder 치환
- 반복 슬라이드를 section 수만큼 복제
- `{{SECTION_TITLE}}`, `{{NARRATIVE}}`, `{{TABLE}}`, `{{CHART}}`, `{{SOURCE}}` 치환
- 차트는 `include_chart = true`인 데이터만 사용
- 기본 차트 규칙:
  - 항목 수가 2개인 선택지는 원그래프
  - 그 외는 세로 막대그래프
  - 가장 큰 값은 강조색 적용
- 표는 PowerPoint 표 객체로 삽입하되, 실패 시 텍스트 표로 fallback

### 차트 검토 PPTX 생성

입력:

- `report_package.json`
- 차트 검토 PPTX 템플릿 경로

출력:

- `<원본파일명>_chart_review.pptx`

기능:

- chart 후보 1개당 슬라이드 1장 생성
- `{{CHART_TITLE}}`, `{{CHART}}`, `{{CHART_NOTE}}`, `{{SOURCE}}` 치환
- 차트 유형, 강조 항목, 원본 table_key 표시
- 생성 실패 차트는 “검토 필요” 슬라이드에 목록화

## 개발 요소

### 1. Writer 모듈 추가

신규 Python 모듈:

- `document_writer.py`

CLI:

```powershell
python -m report_automation_engine.document_writer `
  --package "report_package.json" `
  --preflight "preflight_report.json" `
  --template "report_template_basic.pptx" `
  --type pptx_report `
  --output "draft_report.pptx"
```

지원 type:

- `hwpx_report`
- `pptx_report`
- `chart_review`

### 2. 템플릿 처리 규칙

- placeholder 텍스트와 `RA_` shape 이름을 모두 인식한다.
- placeholder가 여러 개 있으면 첫 번째 항목만 치환하고 경고를 남긴다.
- 반복 슬라이드는 `RA_REPORT_SLIDE` 또는 `{{SECTION_TITLE}}`가 있는 슬라이드로 판단한다.
- 원본 템플릿은 절대 수정하지 않는다.

### 3. 런처 UX

신규 버튼:

- `HWPX 초본 생성`
- `PPTX 초본 생성`
- `차트 검토 PPTX 생성`

버튼 활성 조건:

- `preflight.status == ready` 또는 `ready_with_warnings`
- 해당 템플릿 파일 존재
- 해당 템플릿 검사 결과가 `ready` 또는 `usable_with_warnings`

`blocked` 상태:

- 버튼 비활성화
- 차단 사유를 결과 탭에 표시

`ready_with_warnings` 상태:

- 생성 전 경고 요약 표시
- 사용자가 확인하면 생성 진행

### 4. 오류 처리

- writer 실패 시 원본 템플릿과 package는 보존한다.
- 실패 파일은 삭제하고 `.error.txt`를 output 옆에 저장한다.
- 각 section 처리 실패는 전체 중단하지 않고 QA 슬라이드/문단에 남긴다.
- package schema가 맞지 않으면 즉시 중단한다.

## 테스트 계획

### Package/Preflight

- `ready` package로 HWPX/PPTX 생성
- `ready_with_warnings` package로 경고 확인 후 생성
- `blocked` package에서 버튼 비활성화

### HWPX

- `{{BODY}}`만 있는 최소 템플릿
- 권장 placeholder가 모두 있는 템플릿
- `{{BODY}}` 누락 템플릿 차단
- 생성 파일 원본 템플릿 불변 확인

### PPTX

- 기본 보고서 PPTX 템플릿
- `RA_` shape 이름 기반 템플릿
- 반복 슬라이드 복제
- 원그래프/막대그래프 자동 선택
- 최대값 강조색 적용

### 차트 검토 PPTX

- chart 후보 0개
- chart 후보 1개
- chart 후보 여러 개
- 비숫자 value가 preflight에서 차단되는지 확인

## 구현 순서

1. `document_writer.py`에 package/preflight 로더와 공통 검증 추가
2. PPTX 차트 검토 writer 먼저 구현
3. PPTX 보고서 writer 구현
4. HWPX writer는 텍스트 기반 본문 삽입부터 구현
5. 런처 버튼과 활성/비활성 조건 연결
6. 샘플 package 기반 회귀 테스트 추가

## 기본 결정

- 다음 알파의 첫 writer는 `chart_review` PPTX로 한다.
- HWPX는 이미지/차트 삽입 없이 본문과 표 중심으로 먼저 연다.
- PowerPoint 차트는 편집 가능한 차트 객체를 우선한다.
- 한글 COM 자동화는 HWPX XML 방식이 막힐 때만 후순위로 검토한다.
