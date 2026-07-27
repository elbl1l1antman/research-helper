# 다음 개발 계획

작성 기준 버전: `0.0.28`

## 목표

한도 초기화 또는 새 작업 세션 이후에는 기능을 새로 넓히기보다, 현재 알파의 실패 지점을 줄이는 순서로 진행한다.

우선순위는 다음과 같다.

1. VBA Excel 산출 애드인의 모듈 분리 마무리
2. Excel 산출 시트에서 `report_package.json` 생성 안정화
3. HWP/HWPX 보고서 템플릿의 표 서식 인식
4. 아래한글 COM 기반 HWPX 초본 writer 실사용 검증
5. PPTX 조사 보고서 writer와 대시보드 writer의 템플릿/디자인 안정화

## 1. VBA 모듈 분리 마무리

현재 상태:

- `ReportAutomationAddin.bas`: 진입점, Ribbon 콜백, 실행 흐름
- `ReportAutomationBannerCharts.bas`: 선택 배너 그룹 인식과 배너별 차트 데이터 행 생성
- `ReportAutomationConstants.bas`: 공통 상수
- `ReportAutomationTables.bas`: 원본 집계표 표 블록 탐지
- `ReportAutomationOutputSheets.bas`: 산출 시트 생성
- `ReportAutomationNarratives.bas`: 분석문, 전체 기준 차트/삽입표, QA 산출 흐름
- `ReportAutomationNarrativePoints.bas`: 비율/점수형 핵심 수치 추출, 포인트 표시값 포맷
- `ReportAutomationNarrativeText.bas`: 분석문 문장 조립
- `ReportAutomationUtils.bas`: 공통 유틸리티
- `ReportAutomationSettings.bas`: 설정 읽기/쓰기
- `ReportAutomationOperation.bas`: 실행 상태와 로그

완료된 분리:

- `ReportAutomationNarrativePoints.bas`
  - 점수형/비율형 핵심 수치 추출
  - `ExtractKeyPoints`, `ExtractScorePoints`, `ExtractSimplePoints`, `ExtractWidePoints`
- `ReportAutomationBannerCharts.bas`
  - 배너 그룹 인식
  - 선택 배너별 차트 데이터 행 생성

주의:

- VBA 모듈 간 호출은 `Private` 함수 접근이 불가하므로, 분리 시 공개 API를 최소화해야 한다.
- 동작 변경이 없는 분리만 진행한다.
- 분리할 때마다 dev add-in 재빌드와 실제 Excel 스모크 테스트를 반드시 수행한다.

## 2. Report Package 안정화

목표:

- VBA 산출 시트와 Python writer 사이의 계약을 `report_package.json`으로 고정한다.
- HWPX/PPTX writer는 Excel을 직접 다시 읽지 않고 package만 읽는다.

개발 항목:

- `보고서_분석문` 헤더명 기반 파싱 보강
- `보고서_삽입표`의 표 행 구조 정규화
- `보고서_차트데이터`의 숫자/단위/포함 여부 검증
- `보고서_QA`를 package의 `qa` 배열로 보존
- QA warning을 정상 검토 경고, 개선 필요, 정보 bucket으로 분류
- package 생성 실패 시 원인 JSON 저장

검증 항목:

- 최종 분석문 누락
- 중복 `table_key`
- 숫자 아님
- 출처 범위 누락
- 차트 후보 없음

## 3. HWP/HWPX 표 템플릿 인식

목표:

- 사용자가 제공한 HWP/HWPX 보고서 틀에서 실제 보고서 표 서식을 인식한다.
- 집계표가 들어갈 위치와 표 스타일을 자동화 writer가 재사용할 수 있게 한다.

개발 항목:

- HWPX XML에서 표 객체, 문단 스타일, 셀 배경, 선 스타일, 폰트 크기 추출
- HWP 바이너리는 직접 분석하지 않고 아래한글 COM으로 열린 문서에서 표 속성 검사
- 표 후보별 리포트 생성
  - 표 위치
  - 행/열 수
  - 제목 행 여부
  - 외곽선/내부선
  - 제목 셀 배경
  - 본문/주석 폰트 크기
- 사용자가 “이 표 서식을 기본 삽입표 스타일로 사용”할 수 있도록 설정 JSON 저장

산출 파일:

- `hwp_table_style_report.json`
- `hwp_table_style_profile.json`

## 4. 아래한글 COM 기반 HWPX Writer

목표:

- 한 장이라도 실제 HWPX 보고서 초본을 자동 생성해 가시적으로 확인한다.

v1 범위:

- `{{BODY}}` placeholder 탐색
- 표 제목, 분석문, 삽입용 집계표, 출처 순서 삽입
- 차트는 HWPX v1에서 직접 삽입하지 않고 QA 문구 또는 표 대체
- 원본 템플릿 불변
- 새 `.hwpx` 출력 저장

검증:

- 아래한글 실행 가능 여부
- COM 객체 생성 가능 여부
- template 열기/저장 가능 여부
- `{{BODY}}` 존재 여부
- 생성 파일이 아래한글에서 다시 열리는지 확인

## 5. PPTX Writer 및 대시보드 디자인 안정화

목표:

- 조사 보고서형 PPTX와 기업/기관 세로형 대시보드 PPTX를 분리 유지한다.
- 디자인은 placeholder가 깨지지 않는 객체 구조를 우선한다.

개발 항목:

- 항목명 텍스트 상자와 값 텍스트 상자 분리
- 차트 내부 제목 일괄 삭제
- 사용자 지정 기본 폰트 저장
- PowerPoint chart object 생성 후 폰트 적용 루틴 분리
- 대시보드 디자인 프리셋 2~3개 유지
- 사용자가 편집한 기본 템플릿의 `RA_` shape 이름 인식

검증:

- 한글 글자 깨짐 없음
- 긴 항목명 줄바꿈
- 차트 데이터 편집 가능
- A4/B5 세로형 슬라이드 비율 유지
- 단일 기관/여러 기관 생성 회귀 테스트

## 다음 실행 순서

1. HWPX writer의 표 선/배경/셀 여백 적용 범위를 COM action 단위로 확장한다.
2. 실제 사용자 템플릿에서 인식한 반복 결과 블록과 writer 삽입 위치를 연결한다.
3. 외부 프로그램 전환을 위해 VBA와 Python 엔진이 같은 package 계약을 생성하는 비교 테스트를 추가한다.

## 0.0.20 확인 결과

- VBA `ReportAutomationNarrativePoints.bas` 분리 완료
- dev add-in 빌드 통과
- 실제 Excel 집계표 스모크 테스트 통과
- `report_package.json` 생성 통과
- `preflight_report.json` 상태: `ready_with_warnings`
- package 요약: sections 84개, tables 82개, charts 437개, QA warning 85건, QA error 0건

## 0.0.21 확인 결과

- VBA `ReportAutomationBannerCharts.bas` 분리 완료
- dev add-in 빌드 통과
- 실제 Excel 집계표 스모크 테스트 통과
- `report_package.json` 생성 통과
- `preflight_report.json` 상태: `ready_with_warnings`
- package 요약: sections 84개, tables 82개, charts 437개, QA warning 85건, QA error 0건

## 0.0.22 확인 결과

- `report_package.py` QA warning 분류 필드 추가
- 각 QA에 `category`, `review_action`, `review_bucket`, `review_bucket_label` 추가
- `preflight_report.json` summary에 `qa_warning_buckets`, `qa_warning_categories` 추가
- 실제 Excel 산출물 기준 분류 결과:
  - `normal_review_warning`: 82건
  - `improvement_needed`: 2건
  - `info`: 1건
  - `base_check_needed`: 82건
  - `no_numeric_points`: 2건
  - `run_summary`: 1건

## 0.0.23 확인 결과

- HWPX 표 스타일 인식 보강
- `hwp_template_probe.py`가 표별 `style_summary` 추출
  - table/cell `borderFillIDRef`
  - cell fill color
  - charPr/font height
  - cell margin/size
  - repeat header/cell spacing
- `hwp_template_table_recognizer.py`가 추가 파일 생성
  - `hwp_table_style_report.json`
  - `hwp_table_style_profile.json`
- 실제 국립암센터 HWPX 보고서틀 검증 결과
  - recognition status: `ready`
  - 전체 표: 121개
  - 결과표 후보: 11개
  - style profile: `ready`
  - 대표 스타일 원본: table_index 9, 44행 x 8열
  - 주요 글자 크기: 9.0pt
  - 주요 배경색: none, #E7E7E7, #F3F3F3

## 0.0.26 확인 결과

- HWP COM writer에 render plan 생성 추가
  - `--render-plan-output`
  - `--dry-run`
  - `--max-sections`
- 런처 HWPX 옵션에 초본 문항 수 추가
  - `1개 검증`
  - `3개 검증`
  - `전체`
- 런처 HWPX 생성 시 `_hwp_render_plan.json` 저장
- dry-run 검증 결과:
  - status: `ready`
  - section_count_selected: 1
  - table_preview_rows 생성 확인
  - chart_deferred 표시 확인

## 0.0.27 확인 결과

- HWP COM writer에 표 스타일 profile 입력 추가
  - `--table-style-profile`
  - render plan의 `table_style_profile` 요약 생성
  - writer report의 `table_style_profile_path`, `table_style_profile`, `table_style_applied` 기록
- 런처 HWPX 설정에 `HWP 표 스타일` JSON 선택 필드 추가
- 현재 실제 적용 범위:
  - 대표 글자 크기 `dominant_font_height`
  - 국립암센터 템플릿 style profile 기준 900 = 9.0pt 적용 확인
- 실제 HWPX 생성 검증:
  - status: `ready`
  - sections_written: 1
  - tables_written: 1
  - text_table_fallbacks: 0
  - charts_deferred: 1
  - table_style_applied.dominant_font_pt: 9.0

## 0.0.28 확인 결과

- 런처 UI styling 보강
  - disabled 버튼 색상/테두리 상태 적용
  - 실행 중 버튼 텍스트를 `실행 중...`으로 표시
  - 하단 실행 버튼 영역과 대시보드 명령 영역을 command bar 형태로 정리
  - 표 목록, 문장 리뷰, QA, 대시보드 열 미리보기에 줄무늬 행 적용
- 런처 빌드 통과

## 0.0.29 확인 결과

- HWP COM writer의 표 스타일 profile 처리 보강
  - `table_style_apply_plan` 생성
  - 대표 글자 크기, 헤더 배경색, 주요 셀 선 스타일, 셀 여백 요약을 정규화
  - 현재 적용 가능한 항목과 다음 COM action 구현 대상 분리
    - supported: `dominant_font_height`
    - planned: `header_fill`, `cell_border`, `cell_margin`
- HWP COM writer checkpoint report 저장 보강
  - render plan 생성 직후 writer report 저장
  - 템플릿 복사, COM 객체 생성 진입, 문서 열기, 본문 작성, 저장 단계별 report 갱신
  - COM이 멈춰도 마지막 `stage`/`action`으로 위치 추적 가능
- 검증 결과
  - `py_compile` 통과
  - dry-run 통과
  - style profile 기준 `header_fill_color: #E7E7E7`, 주요 border, cell margin 요약 확인
  - 실제 COM 호출은 현재 PC에서 `create_hwp_object` 단계 지연 확인

## 다음 개발 우선순위

1. HWP COM 환경 진단을 `create_hwp_object`, 보안 모듈 등록, 파일 열기 단계로 분리한다.
2. HWP COM 호출이 일정 시간 이상 멈출 때 사용자가 원인을 볼 수 있도록 런처에 checkpoint report 표시를 추가한다.
3. `CellBorderFill`, `TablePropertyDialog` action을 작은 샘플 문서에서 별도 검증한 뒤 표 배경/선/셀 여백 실제 적용으로 확장한다.
