# 다음 개발 계획

작성 기준 버전: `0.0.28`

## 2026-10-02 안정화 진행 상황

현재 버전은 `0.0.34`를 유지합니다. 아래 과거 계획보다 HWPX 실제 표 출력 안정화를 우선합니다.

- 완료: 셀 표시값·병합·역할 기반 서식, 최대 3개 사용자 서식 저장.
- 완료: 현재 구역의 본문 폭 계산, 보기/BASE 반복을 유지한 가로 분할, 임시 Excel 범위의 열 폭 조정.
- 완료: 신규 개체 ID 기반 Paste 검증, 무변경 Paste의 잘못된 Undo 방지, 기존 개체 소실 차단.
- 완료: TableCreate 기본 셀 여백 중복 계산 수정, 배경색 적용 후 테두리 적용, 단 기준 왼쪽 배치.
- 완료: 표 안의 최상단 병합 BASE 행(오른쪽 정렬·흰색·무테), 좌우 외곽선 제거 기본값과 저장 가능한 표시 옵션.
- 완료: BASE 행·다단 헤더의 세로 쪽 나눔 반복, 병합 후 헤더 플래그 적용. 본문까지 이어지는 헤더 병합은 BASE만 반복하고 경고.
- 완료: 테두리 적용 시 선택 셀 배경 파라미터도 유지해 배경색 소실 수정. BASE/단위는 표 목록 우선이며 불명확한 정보는 추정하지 않음.
- 완료: 생성 전 편집용지 크기·방향과 상하좌우/제본/머리말/꼬리말 mm 설정. 기본은 템플릿 유지, 직접 지정은 BODY 구역에만 적용한 뒤 실제 본문 폭으로 표 배치. [사용 안내](hwp_page_setup.md).
- 검증: 실제 3개 표를 11개 표로 생성해 2,279개 셀과 83개 병합을 대조하고 한글 재열기·PDF 출력을 확인. 강제 fallback 214개 셀도 대조.
- 추가 검증: BASE 행을 포함한 11개 표의 2,290개 셀·94개 병합·363개 괄호형 값을 대조. clipboard 9개와 실제 실패 후 fallback 2개 모두 통과. 22쪽 PDF의 표가 있는 21쪽에서 BASE 행을 확인하고 9개 이어진 페이지의 전체 헤더 문자를 대조.
- 근거: `outputs/hwp_header_metadata_20261002/KISDI_header_metadata.hwpx`, `.pdf`, `.verification.json`. Python 검증 78개와 런처 저장소/CLI/최소 창 크기 검증 통과. 원본과 결과 파일은 Git에서 제외.
- 편집용지 실측 검증: `outputs/hwp_page_setup_20261002/`의 실제 A4/B5/A4 가로 HWPX 및 PDF. 본문/표 폭은 각각 45,439/44,900, 39,686/39,000, 68,683/67,900 HWPUNIT으로 확인. 가로 출력은 제본 5mm도 적용. 원본 Excel/템플릿 해시 불변, 모두 아래한글 재열기 통과.
- 편집용지 회귀: Python 70개 + 기존 계약 검증 12개 통과. 각 용지별 142개 표시값·병합 셀을 대조해 총 426개 검증. 설정 미지정/dry-run/잘못된 입력 차단도 포함.
- 런처 회귀: 편집용지 저장/불러오기·CLI 전달·템플릿 유지 모드·최소 크기 배치와 기존 서식 저장소/CLI 자체 검증 통과. 기존 배포용 런처도 다시 빌드.
- 잔여 점검: 같은 테스트 프로세스에서 한글 종료 직후 다음 COM 문서 열기가 한 차례 실패. 케이스를 별도 프로세스로 실행하면 모두 통과했으며, 빠른 연속 생성의 COM 종료/참조 해제 타이밍은 후속 회귀 대상으로 남김.

다음 순서:

1. 헤더/본문 경계를 가로지르는 병합 표와 세로 배너의 쪽별 표기 정책을 실제 템플릿으로 추가 검증한다.
2. 사용자가 분할 경계·보기 열 폭을 확인/보정하는 미리보기를 기존 표 목록에 연결한다.
3. 서로 다른 실제 HWPX 템플릿 2개 이상에서 페이지 폭·글꼴·여백 회귀를 수행한다.
4. 한 페이지에 들어가지 않는 본문 병합 표의 가로 용지/명시적 분할 옵션을 검토한다. 현재는 데이터 손실 없이 차단한다.
5. PPT 기능은 유지보수하며 HWPX 표 품질 검증 이후 차트 삽입을 재개한다.

다단 본문은 아직 지원하지 않습니다. 원자료와 생성 문서는 `outputs/`에만 저장하며 Git에는 코드·테스트·검증 설명만 기록합니다.

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

## 0.0.30 확인 결과

- HWP COM 환경 진단을 단계별 checkpoint로 세분화
  - `import_win32com`
  - `dispatch`
  - `register_file_path_checker`
  - `set_visible`
  - `open_document`
  - `save_as`
- `writer_report.json`의 COM 진단 필드 확장
  - `com.current_prog_id`
  - `com.steps[]`
- 실제 COM 호출 지연 시 마지막 단계 확인 가능
  - 현재 PC 검증 결과: `stage=com`, `action=dispatch`, `current_prog_id=HWPFrame.HwpObject`
- 검증 결과
  - `py_compile` 통과
  - dry-run 통과
  - 실제 COM 호출 checkpoint report 생성 확인

## 다음 개발 우선순위

1. 런처 HWP 실행 결과 영역에 `writer_report.json`의 마지막 COM step을 표시한다.
2. `EnsureDispatch` 지연 회피를 위해 `Dispatch`, `DispatchEx`, `EnsureDispatch`를 선택 가능한 dispatch mode로 분리한다.
3. dispatch mode별 환경 테스트 결과를 `hwp_environment_report.json`에 기록한다.

## 0.0.31 확인 결과

- HWP COM dispatch mode 선택 옵션 추가
  - `ensure_dispatch`: 기존 기본값 유지
  - `dispatch`
  - `dispatch_ex`
- 적용 위치
  - HWPX 문서 생성 CLI: `--dispatch-mode`
  - COM 환경 진단 CLI: `--check-environment --dispatch-mode`
  - `writer_report.json`의 `com.dispatch_mode`, `com.steps[].dispatch_mode`
- 검증 결과
  - `py_compile` 통과
  - dry-run 통과
  - `--dispatch-mode dispatch` 환경 진단 checkpoint 생성 확인
  - 현재 PC에서는 `dispatch`도 `stage=com`, `action=dispatch`, `current_prog_id=HWPFrame.HwpObject` 단계에서 지연

## 다음 개발 우선순위

1. 런처 HWP 설정 영역에 dispatch mode 선택 UI를 추가한다.
2. 런처 실행 결과 영역에 마지막 COM step과 dispatch mode를 표시한다.
3. 별도 환경 진단 명령에서 dispatch mode별 결과를 하나의 `hwp_environment_report.json`으로 비교 저장한다.
4. COM dispatch 단계가 장시간 지연될 때 Python 프로세스를 분리 실행하고 timeout 후 종료하는 보호 실행기를 검토한다.

## 0.0.32 확인 결과

- 런처 HWPX 옵션 영역에 dispatch mode 선택 UI 추가
  - `ensure_dispatch`
  - `dispatch`
  - `dispatch_ex`
- 런처에서 HWPX writer 실행 및 COM 환경 점검 시 `--dispatch-mode` 전달
- 런처 실행 결과 요약에 `writer_report.json`의 HWP COM 상태 표시
  - status
  - dispatch mode
  - current ProgID
  - stage/action
  - last COM step
- launcher config에 `HwpDispatchMode` 기록
- 검증 결과
  - 런처 빌드 통과

## 다음 개발 우선순위

1. 별도 환경 진단 명령에서 dispatch mode별 결과를 하나의 `hwp_environment_report.json`으로 비교 저장한다.
2. 런처에 HWP COM 환경 진단 전용 버튼을 추가한다.
3. COM dispatch 단계가 장시간 지연될 때 Python 프로세스를 분리 실행하고 timeout 후 종료하는 보호 실행기를 검토한다.

## 0.0.33 확인 결과

- 런처 HWPX 옵션 영역에 `COM diag` 버튼 추가
- 버튼 실행 시 dispatch mode별 환경 진단 수행
  - `ensure_dispatch`
  - `dispatch`
  - `dispatch_ex`
- 각 mode는 별도 Python 프로세스로 실행
  - mode별 timeout: 15초
  - timeout 발생 시 해당 프로세스를 종료하고 다음 mode 계속 진행
- 통합 진단 리포트 생성
  - `hwp_environment_report_YYYYMMDD_HHMMSS.json`
  - mode별 개별 writer report도 함께 저장
  - dispatch mode, status, stage/action, current ProgID, last COM step 기록
- 런처 실행 결과 영역에 통합 진단 요약 표시
- 검증 결과
  - 런처 빌드 통과

## 다음 개발 우선순위

1. Report Package table matrix v2 계약 안정화
2. HWPX table preflight 차단/경고 정밀화
3. HWPX 직접 writer 최소 경로 검증
4. 실제 사용자 HWPX 템플릿 1장 생성 검수
5. COM writer는 fallback/진단용으로 정리

## 0.0.34 확인 결과

- VBA 셀 계약 `보고서_삽입표셀` 추가
  - Excel `Range.Text` 표시값과 `Value2` 원시값 분리
  - 병합, 셀 역할, 정렬, 원본 시트/범위 보존
- 런처 표 범위 검토 기능 추가
  - 자동 탐지 범위와 최종 사용 범위 표시
  - Excel 기본 범위 선택 창으로 표별 수동 범위 보정
  - 작업 복사본의 `table_range.Txxxx` 설정으로 전달
- `report_package.json` table matrix v2 우선 읽기 및 source 범위 검증
- HWP COM writer에 Excel Range.Copy/Paste 경로 추가
  - 새 HWP 표 객체와 행·열 크기 검증
  - 실패 시 matrix 기반 `TableCreate` fallback
  - 표별 insert mode와 fallback 사유 기록
- 검증 결과
  - Python 테스트 14개 통과
  - VBA add-in 빌드 및 Excel COM smoke 통과
  - 수동 범위 `A2:B3`, 표시값 `3.3`, 원시값 `3.333535353` 보존 확인
  - 런처 빌드 통과
  - 현재 PC의 HWP COM dispatch는 객체 생성 단계에서 지연되어 실제 HWPX 저장 검증은 미완료

## 다음 개발 우선순위

1. HWP COM이 정상 동작하는 환경에서 한 문항 HWPX fixture 검증
2. clipboard 표의 역할별 배경색·선·여백 적용 검증
3. 병합 표 `TableCreate` fallback 셀 병합 지원
4. 표 폭 초과 시 축소 또는 분할 정책 구현

## HWP 이름 있는 스타일 및 표 정리 현황 (Task7)

기준 분기: `codex/excel-range-hwpx-tables`, Tasks1-6 구현·리뷰 기준 커밋: `a779ff4`. 위 버전별 확인 결과는 당시 검증 이력으로 보존하며 이번 문서 정리는 버전을 올리지 않습니다.

| 보고 이슈 | 구현 상태 | 반영 내용 |
| --- | --- | --- |
| 첫 분석문과 후속 문장이 붙어 출력됨 | IMPLEMENTED | 줄바꿈/지정된 기존 연결 경계를 `narrative_blocks`로 분리하고 첫 문단은 본문1, 모든 후속 문단은 본문2 적용 |
| `(3,232)`가 `-3232`로 변환됨 | IMPLEMENTED | VBA 표시문자 열을 텍스트로 저장하고 `display_text`를 출력 기준으로 유지, 실제 음수는 보존 |
| 장식용 특수기호가 표에 출력됨 | IMPLEMENTED | HTML entity 해제, 장식 기호 정리, 구두점·수학/통화 기호·허용 단위 보존, 원본 셀 주소 QA 기록 |
| 역할별 서식을 선택하고 재사용할 수 없음 | IMPLEMENTED | 이름 있는 스타일 5개 등록/갱신, 본문2 실제 글머리표, 셀 외형 적용, 내장 기본값과 사용자 프리셋 최대 3개 |

기본값은 `report_automation_engine/config/default_hwp_style_config.json`입니다. 고정 이름 `보고서 본문1`/`보고서 본문2`는 맑은 고딕 10pt, 줄간격 160%, 왼쪽 정렬이며 본문2만 `-` 글머리표와 왼쪽 5mm/첫 줄 -5mm 들여쓰기를 사용합니다. 본문1 들여쓰기는 0/0mm입니다. `표보기`/`표배너`/`표숫자`는 맑은 고딕 8.5pt, 줄간격 130%, 들여쓰기 0/0mm이며 표보기는 왼쪽, 나머지는 가운데 정렬, 표배너만 굵게입니다. 표 스타일에는 글머리표가 없습니다.

표배너 배경은 `E7E7E7`, 나머지는 `FFFFFF`, 세로 정렬은 모두 가운데입니다. 테두리는 검정 `000000`, 내부 0.12mm/외곽 0.4mm이며 한글 지원 두께로 변환합니다. 내장 `기본 보고서`는 수정·삭제할 수 없고 사용자 서식 저장소는 `%LOCALAPPDATA%\ResearchHelper\hwp_style_presets.json`입니다. 손상 시 `.corrupt.<timestamp>.<id>.bak` 백업과 유효 항목 보존 복구를 수행합니다.

GUI는 HWPX 옵션의 `서식`과 `서식 설정`에서 저장 프리셋을 선택합니다. 런처 CLI는 `--hwp-style-config`로 입력하고 package 옆 선택 실행 설정 `hwp_style_config.json`을 기록해 writer의 `--style-config`로 전달합니다. 기본값·문자 정책·전체 실행 순서는 루트 README와 엔진 README를 참고합니다.

### 실제 KISDI 3개 표 회귀: 데이터·서식 검증 완료

2026-10-01 실제 파일 `3. 통계표_kisdi2025 디지털산업 실태조사 보고서 Table(%).xlsx`를 읽기 전용으로 열어 작업 복사본을 만들고 VBA 셀 계약 추출 → package/preflight → HWP COM 생성을 실행했습니다. 기존 기록의 contract fallback과 달리 이번 결과는 3개 표 모두 `clipboard`이며 다른 환경의 성공까지 보장하지 않습니다.

- 표 크기: 35×5, 36×36, 35×10. 아래한글 재열기에서 독립 표 컨트롤 3개 확인
- 표시문자 1,691개, 괄호형 값 99개, 병합 시작 셀 25개를 계약과 XML로 대조
- 필수 이름 있는 스타일 5개가 각각 한 번 등록되고 역할별 스타일 적용 확인
- 스타일 호출: 본문1 15회, 본문2 3회, 표보기 111회, 표배너 95회, 표숫자 1,485회
- 본문1/본문2 분리 및 글머리표 확인. 원본 Excel/템플릿 SHA-256 불변
- writer 오류 0건, Excel/한글 테스트 COM 종료 및 임시 작업 폴더 정리 확인
- HWPX와 동일한 바이트의 임시 사본으로 재열기하고 한글 문서 API로 PDF 내보내기 확인

로컬 근거: `outputs/hwp_style_regression_20261001/KISDI_three_tables_styled.hwpx`, `KISDI_three_tables_preview.pdf`, `hwp_writer_report.json`, `verification.json`. 개인/업무 원자료와 결과물은 Git에 포함하지 않습니다.

실제 검증에서 수정한 사항:

- 상대 workbook 경로를 package 생성 시 절대 경로로 고정하고 openpyxl 핸들을 닫음
- 붙여넣은 표 앵커로 셀에 진입하고 셀 블록이 아닌 텍스트만 교체
- 세로 병합으로 다시 나타나는 셀을 건너뛰고 이동 오류/반복은 중단
- Undo 실패 또는 편집 중 오류 시 중복 fallback 차단
- 현재 표 바로 뒤로 복귀해 후속 section이 템플릿 QA 뒤로 이동하지 않도록 수정
- 실선 enum을 API로 조회하고 병합 셀의 끝 좌표로 외곽선 두께 판정
- 중복 스타일 참조 복구, 공유 글머리표 보존, writer 직접 파일 실행과 저장소 복구 경고 보완

### 다음 우선순위: 넓은 표의 인쇄 레이아웃

이번 검증은 데이터·서식 회귀의 완료이며 인쇄 완성도는 미완료입니다. Excel 열 너비가 그대로 붙어 본문 폭을 초과하고 특히 36열 표는 PDF에서 오른쪽 열이 잘립니다. 표 전체 너비 API는 실제 셀 폭을 바꾸지 않았고 셀별 비례 조정 시험에서는 병합 셀 이동이 중단되어 정식 코드에 적용하지 않았습니다.

1. 표 생성 전 본문 가용 폭과 실제 열 폭을 비교해 초과량을 명시하고 인쇄용 출력 차단/경고 기준 확정
2. 가로 배너 단위로 분할하고 각 조각에 보기·세로 배너·BASE 열을 반복하는 계약 추가
3. 허용 최소 글자 크기·보기 열 최소 폭을 지키는 범위에서만 축소하고, 불가능하면 분할 또는 차단
4. 표 전체 한글/PDF 검수로 열 잘림·고정 행 높이·쪽 나눔·반복 헤더를 확인한 뒤 정식 기능으로 연결
