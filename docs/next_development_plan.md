# 다음 개발 계획

작성 기준 버전: `0.0.22`

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

1. HWPX 표 스타일 인식기를 먼저 만들고, 실제 제공된 보고서 틀 1개로 표 서식 리포트를 생성한다.
2. 아래한글 COM writer로 `{{BODY}}` 위치에 표 1개와 분석문 1개를 삽입하는 최소 실사용 테스트를 수행한다.
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
