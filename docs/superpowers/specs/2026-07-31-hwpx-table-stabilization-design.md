# HWPX Table Stabilization Design

## Goal

HWPX 보고서 생성을 프로젝트의 1순위 산출물로 고정하고, Excel 집계표가 HWPX 표로 들어갈 때 구조와 표시값이 손실되지 않도록 중간 데이터 계약과 검증 레이어를 강화한다.

## Current Problems

- `report_package.json`의 `tables[].rows`는 `category`, `percent`, `weighted_n`, `raw_n` 중심이라 실제 보고서 집계표의 다단 헤더, 가로/세로 배너, 병합 구조, 출처 범위를 보존하지 못한다.
- HWP writer는 현재 단순 표를 새로 만드는 방식이라 실제 보고서틀의 반복 블록 구조와 맞지 않는다.
- Excel 값은 raw number로 읽히면 `3.333535353` 같은 값이 문서에 들어갈 수 있다. 보고서에는 Excel 화면 표시값인 `3.3`이 들어가야 한다.
- COM writer 진단은 유용하지만, COM dispatch 지연이 반복되는 환경에서는 핵심 개발 경로로 삼기 어렵다.
- PPTX와 대시보드 기능은 유지하되, 현 단계 신규 개발 우선순위에서는 제외한다.

## Product Direction

다음 개발 단위는 HWPX 표 안정화에만 집중한다.

```text
Excel 산출 시트
  -> report_package.json v1 호환 + table_matrix v2
  -> preflight_report.json table checks
  -> hwp_direct_writer 또는 hwp_com_writer fallback
  -> HWPX 초본
```

PPTX 작성기와 대시보드 작성기는 기존 기능을 유지하지만, 이번 단계에서는 문서와 런처 우선순위에서 후순위로 표기한다.

## Data Contract

기존 `tables[].rows`는 호환성을 위해 그대로 유지한다. 새 writer와 preflight는 `tables[].matrix`와 `tables[].cells`를 우선 사용한다.

### Table Fields

- `table_key`: section과 table을 연결하는 고유 키
- `title`: 보고서 표 제목
- `source_sheet`: 원본 산출 시트명
- `source_range`: 원본 범위 또는 산출 행 범위
- `row_count`: matrix 행 수
- `col_count`: matrix 열 수
- `matrix`: 행/열 형태의 cell 객체 배열
- `cells`: matrix를 1차원으로 펼친 cell 객체 배열
- `merged_ranges`: 향후 원본 병합 구조 보존용 배열
- `roles`: role별 cell 수 요약
- `style_hints`: 표 폭/정렬/글자 크기 기본 힌트
- `qa`: table 단위 warning/error

### Cell Fields

- `row`: 1-based row index
- `col`: 1-based column index
- `rowspan`: 기본값 `1`
- `colspan`: 기본값 `1`
- `role`: `header`, `banner_horizontal`, `banner_vertical`, `stub`, `value`, `base`, `note`, `source`
- `display_text`: HWPX에 실제 삽입할 문자열
- `raw_value`: 계산/QA용 원본 값
- `number_format`: Excel number format 또는 산출 포맷 설명
- `source_cell`: Excel 셀 주소 또는 산출 행 식별자
- `align`: `left`, `center`, `right`
- `is_numeric`: 숫자 여부

## Display Value Policy

HWPX 삽입값 우선순위는 다음으로 고정한다.

1. 산출 시트의 explicit display column
2. Excel/VBA의 `Range.Text` 기반 표시값
3. Python formatter가 number format과 `decimal_places`로 만든 표시값
4. 문자열화한 raw value

raw value는 HWPX에 직접 삽입하지 않는다. raw/display 차이가 확인되면 QA warning으로 남긴다.

## Table Shape Policy

v1은 원본 Excel의 모든 병합 구조를 완전 재현하지 않는다. 대신 보고서용 재구성 표를 안정적으로 만든다.

- 기본 헤더: `항목`, `비율`, `가중 N`, `원 N`
- 행 본문: 기존 `tables[].rows`를 matrix로 변환
- 향후 원본 구조 보존을 위해 `merged_ranges`, `source_range`, `source_cell` 필드는 지금부터 유지
- cell role은 writer 스타일 적용과 preflight 판단에 사용

## Preflight Policy

차단 조건:

- table matrix가 없거나 행/열 수가 0
- cell에 `display_text`가 없고 raw value만 있음
- 병합 좌표가 충돌함
- 값 영역이 전부 비어 있음
- section이 참조하는 table_key가 없음

경고 조건:

- 셀 텍스트가 길어 줄바꿈이 필요함
- 열 수가 많아 본문 폭 초과 가능성이 있음
- raw value와 display_text의 반올림 결과가 다름
- source_range가 없음
- 표 제목이 없음
- 소수점 자리 설정과 표시값이 불일치할 가능성이 있음

## Writer Policy

이번 단계의 writer는 두 가지 목표를 분리한다.

- `hwp_com_writer.py`: 기존 COM writer를 fallback으로 유지한다.
- `hwp_direct_writer.py`: package/preflight/template/output을 받아 템플릿 원본을 보존하고, `{{BODY}}`가 있는 HWPX 패키지에 section 본문과 matrix 기반 표 payload를 직접 쓴다.

직접 writer의 v1 완료 기준은 다음이다.

- 원본 템플릿을 수정하지 않음
- output HWPX zip이 유효함
- `{{BODY}}`가 output 내부에서 제거됨
- section 제목, 분석문, table matrix의 `display_text`가 output 내부 XML에 존재함
- raw 숫자값이 display_text 대신 들어가지 않음
- 실패 시 `hwp_direct_writer_report.json`에 stage/action/error를 남김

한글에서 시각적으로 완전한 표 객체로 열리는 품질은 다음 단계에서 실제 HWPX 표 XML 생성 검증으로 강화한다.

## Testing Strategy

테스트는 실제 HWP 설치 없이 통과 가능한 계층부터 만든다.

- matrix 단위 테스트: display value, role, long text warning, source metadata
- package/preflight 통합 테스트: fake Excel 산출 시트에서 package와 preflight 생성
- direct writer 테스트: 최소 HWPX zip 템플릿에서 placeholder 제거와 payload 삽입 확인
- 기존 기능 회귀: `python -m compileall -q report_automation_engine`

## Non-Goals

- PPTX writer 신규 개선
- 대시보드 PPT 디자인 개선
- HWP 바이너리 직접 편집
- HWPX 차트 자동 삽입
- 외부 라이브러리 `rhwp` vendoring
- 런처 대규모 UI 재작성

## Acceptance Criteria

- `report_package.json`의 각 table에 `matrix`, `cells`, `roles`, `qa`가 생성된다.
- 기존 `tables[].rows` 소비자는 깨지지 않는다.
- `preflight_report.json` summary에 table matrix 관련 warning/error 수가 반영된다.
- `hwp_direct_writer.py --dry-run` 또는 기본 실행으로 writer report가 생성된다.
- 테스트용 HWPX 출력 zip에서 `{{BODY}}`가 제거되고 `display_text` 값이 보존된다.
- Python 엔진 compile 검증이 통과한다.
