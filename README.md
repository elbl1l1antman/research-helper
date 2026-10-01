# ReportAutomation

Current version: `0.0.34`

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
- 기업/기관 대시보드 PPT용 원자료 Excel 선택, sheet/기관/열 선택, KPI/차트 슬롯 매핑
- 기업/기관 대시보드 PPT용 디자인 프리셋과 폰트 선택
- 토큰 기반 WinForms UI 스타일 적용
- 정보 카드형 헤더, owner-draw 탭, 문서형 그룹 패널 스타일 적용
- HWPX writer render plan 생성과 문항 수 제한 실행
- HWP 표 스타일 profile 연결, 대표 글자 크기 적용, 선/배경/여백 적용 계획 기록
- 고정된 HWP 이름 있는 스타일 5종과 기본 서식/사용자 서식 최대 3개 선택 및 저장
- disabled 버튼, command bar, 줄무늬 목록 행 등 런처 UI styling 보강
- `report_package.json`, `preflight_report.json` 생성 결과 표시

주요 문서:

- `report_automation_launcher/README.md`
- `report_automation_launcher/ALPHA.md`
- `report_automation_launcher/NEXT_DOCUMENT_OUTPUT_PLAN.md`
- `docs/launcher_ui_design_system.md`

### `report_automation_addin`

런처가 호출하는 Excel VBA 산출 엔진입니다.

사용자가 직접 매크로를 조작하는 구조가 아니라, 런처가 Excel COM으로 add-in을 열고 실행합니다.

VBA 소스는 기능별 표준 모듈로 분리되어 있습니다.

- `ReportAutomationAddin.bas`: 공개 진입점, Ribbon 콜백, 전체 실행 흐름
- `ReportAutomationConstants.bas`: 공통 상수와 tableRec 인덱스
- `ReportAutomationBannerCharts.bas`: 선택 배너 그룹 인식과 배너별 차트 데이터 행 생성
- `ReportAutomationTables.bas`: 원본 집계표 시트 탐색, 표 블록 탐지, 표 제목/범위/유형 파싱
- `ReportAutomationOutputSheets.bas`: 산출 시트 생성, 설정/표목록/출처/메타 작성
- `ReportAutomationNarratives.bas`: 분석문, 핵심 포인트, 차트 데이터, 삽입표, QA 산출
- `ReportAutomationNarrativePoints.bas`: 집계표에서 비율/점수형 핵심 수치 추출, 포인트 표시값 포맷
- `ReportAutomationNarrativeText.bas`: 분석문 문장 조립, 제목 정규화, 척도형/점수형 문장 생성
- `ReportAutomationUtils.bas`: 문자열/숫자/시트명/공통 서식 유틸리티
- `ReportAutomationSettings.bas`: 설정 시트 읽기/쓰기
- `ReportAutomationOperation.bas`: Excel 실행 상태 관리와 내부 로그

생성하는 주요 시트:

- `보고서_분석문`
- `보고서_차트데이터`
- `보고서_삽입표`
- `보고서_삽입표셀`: 셀 표시값, 원시값, 병합, 역할, 원본 범위
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
- `report_package.py`: Excel 산출 시트를 중간 JSON 계약으로 변환하고 QA warning 유형 분류와 preflight 수행
- `document_writer.py`: `report_package.json` 기반 PPTX 초본 생성 시작점
- `dashboard_package.py`: 기업/기관 가로형 원자료를 대시보드 JSON 계약으로 변환하고 preflight 수행
- `dashboard_writer.py`: 대시보드 JSON 계약을 세로형 A4/B5 PPTX로 생성
- `hwp_com_writer.py`: 아래한글 COM으로 HWPX 템플릿 사본에 본문/표 초본 생성
- `hwp_style_config.py`: 고정된 HWP 스타일 5종의 서식 JSON 검증
- `hwpx_style_registry.py`: 작업용 HWPX에 이름 있는 스타일 등록/갱신
- `hwp_template_probe.py`: HWP/HWPX 보고서틀의 표 구조와 결과표 후보 분석
- `template_blueprint.py`: HWP/HWPX 보고서틀의 반복 결과 블록 후보 생성
- `hwp_template_table_recognizer.py`: 사용자 제공 HWP/HWPX 템플릿의 자동화 표 후보와 표 스타일 프로필 인식
- `hwp_table_mapping.py`: `report_package.json`의 section/table과 HWP 템플릿 표 후보를 writer 계약으로 연결
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
- 표별 자동 탐지 범위 확인 및 Excel에서 수동 범위 재선택
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
- HWP/HWPX 보고서틀 표 후보 인식 및 `hwp_table_mapping.json` 생성
- 문서 생성 준비 상태 표시
- CLI 기반 PPTX 보고서/차트 검토 초본 생성
- PowerPoint 편집 가능한 차트 객체 생성
- 기업/기관 세로형 A4/B5 대시보드 PPTX 생성
- 대시보드 PPTX 디자인 프리셋 3종: 모던 블루, 모던 민트, 그래파이트
- 대시보드 PPTX 하단 인포그래픽 일러스트 자동 삽입
- 아래한글 COM 기반 HWPX 초본 생성
  - Windows + 아래한글 설치 + `pywin32` 환경에서 동작
  - 원본 템플릿을 보존하고 출력 경로에 사본 저장
  - `hwp_render_plan.json`으로 실제 삽입 전 문항/표/차트 대체 계획 확인
  - 알파 검증용으로 1개/3개/전체 문항 수 제한 가능
  - `hwp_table_style_profile.json`을 선택하면 대표 표 글자 크기를 HWP 표 생성 전에 적용
  - 표 스타일 profile에서 헤더 배경색, 주요 선 스타일, 셀 여백을 정규화해 `table_style_apply_plan`으로 기록
  - HWP COM 호출 전후 checkpoint report를 저장해 멈춤 위치를 추적
  - COM 생성/ProgID dispatch/보안 모듈 등록/파일 열기/저장 단계별 진단 기록
  - HWP COM dispatch mode 선택: `ensure_dispatch`, `dispatch`, `dispatch_ex`
  - 런처에서 HWP COM dispatch mode 선택 및 마지막 COM step 요약 표시
  - 런처 `COM diag` 버튼으로 dispatch mode별 환경 진단 및 `hwp_environment_report.json` 생성
  - `{{BODY}}` 위치에 제목, 분석문, 표, 출처를 반복 삽입
  - Excel 표시값, 원시값, 병합, 역할을 셀 계약으로 분리 보존
  - Excel 범위를 한글 표로 붙여넣고 표 객체와 크기를 검증한 뒤 실패 시 `TableCreate`로 전환
  - 분석문을 문단별로 분리하고 이름 있는 본문/표 스타일과 셀 배경·테두리·정렬 적용

계약 기반 fallback은 matrix의 `display_text`로 `TableCreate` 표를 채우고 `merged_ranges`의 셀 병합을 적용합니다. 표 생성 또는 필수 스타일 적용 실패 시 writer 오류 리포트에 중단 사유를 기록합니다.

## HWP 보고서 서식

기본값은 `report_automation_engine/config/default_hwp_style_config.json`의 `기본 보고서`입니다. 아래한글 스타일 이름은 다음 5개로 고정되며 이름을 바꾸거나 추가하지 않고 속성만 변경합니다. 모든 기본 글꼴은 `맑은 고딕`입니다.

| 스타일 | 용도 | 크기 | 굵게 | 줄간격 | 왼쪽/첫 줄 들여쓰기 | 정렬 | 글머리표 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 보고서 본문1 | 첫 분석문 문단 | 10pt | 아니오 | 160% | 0/0mm | 왼쪽 | 없음 |
| 보고서 본문2 | 모든 후속 분석문 문단 | 10pt | 아니오 | 160% | 5/-5mm | 왼쪽 | `-` |
| 표보기 | 보기·세로 배너·주석 등 | 8.5pt | 아니오 | 130% | 0/0mm | 왼쪽 | 없음 |
| 표배너 | 가로 배너·BASE·헤더·제목 셀 | 8.5pt | 예 | 130% | 0/0mm | 가운데 | 없음 |
| 표숫자 | 값 셀 | 8.5pt | 아니오 | 130% | 0/0mm | 가운데 | 없음 |

`보고서 본문2`의 `-`는 문장에 붙이는 문자열이 아니라 실제 스타일 글머리표입니다. 후속 문단에서 선택한 글머리표와 공백/탭이 붙은 접두어만 제거해 중복을 피합니다. 글머리표를 비워 두면 원문 접두어를 유지하고 실제 음수 `-3232`도 보존합니다. 분석문의 CRLF/LF/CR 줄바꿈을 별도 문단으로 만들고, 기존 문장에서 붙어 있는 `나타남다음으로`, `나타남그다음으로`, `나타남반면,` 경계도 복구합니다.

표 셀 배경은 `표보기`/`표숫자`가 흰색 `FFFFFF`, `표배너`가 회색 `E7E7E7`이며 세로 정렬은 모두 가운데입니다. 기본 테두리는 검정 `000000`, 내부선 0.12mm, 외곽선 0.4mm입니다. 셀 배경·테두리·세로 정렬은 문단 스타일과 별도로 같은 프리셋에서 적용하며, 선 두께는 한글이 지원하는 가장 가까운 값으로 변환합니다.

### GUI 선택과 저장

1. 런처에서 Excel 집계표, HWP/HWPX 템플릿과 `HWPX 보고서` 출력을 선택합니다.
2. HWPX 옵션의 `서식`에서 `기본 보고서` 또는 저장된 사용자 서식을 선택합니다. `서식 설정`에서 5개 스타일과 표 셀 배경·정렬·테두리를 편집할 수 있습니다.
3. 다른 이름으로 `새로 저장`하거나 사용자 서식을 `덮어쓰기`한 뒤 `선택 후 닫기`를 누릅니다. 저장하지 않은 편집값은 선택만으로 적용되지 않습니다. `기본값으로 초기화`는 편집값을 기본값으로 되돌리며 저장된 서식은 별도로 저장해야 변경됩니다.
4. 초본 문항 수를 `1개 검증`, `3개 검증`, `전체` 중 선택하고 실행합니다. Excel 작업 복사본과 package/preflight를 만든 뒤, 선택 서식을 `report_package.json`과 같은 폴더의 `hwp_style_config.json`으로 저장해 writer의 `--style-config`로 전달합니다.
5. 결과 탭에서 HWPX와 writer report를 확인합니다. 실행 기록에는 `HwpStylePreset`, `HwpStyleConfig`, writer report에는 `style_preset_name`, `style_config_path`, `style_index_map`, `style_application_counts`가 남습니다. 적용 횟수에는 제목·출처 등의 본문 스타일 호출도 포함됩니다.

내장 `기본 보고서`는 수정·덮어쓰기·삭제할 수 없으며 사용자 서식은 최대 3개입니다. 같은 이름은 명시적인 덮어쓰기 확인이 필요하고 네 번째 새 서식 저장은 거부됩니다. 사용자 서식만 `%LOCALAPPDATA%\ResearchHelper\hwp_style_presets.json`에 저장하며 내장 기본값은 로드할 때 합칩니다.

저장소가 손상되거나 잘못된/중복/초과 항목이 있으면 원본을 `hwp_style_presets.json.corrupt.<timestamp>.<id>.bak`으로 백업하고 읽을 수 있는 유효한 사용자 서식 최대 3개를 보존해 복구합니다. JSON 자체를 읽을 수 없으면 사용자 서식 없이 내장 기본값으로 복구합니다. 복구 경고와 백업 경로는 GUI 실행 로그 또는 CLI 오류 출력에 표시합니다. 저장은 임시 파일을 쓴 뒤 교체하는 방식입니다.

### 표 표시문자 정책

`보고서_삽입표셀`의 `display_text`는 Excel `Range.Text`를 텍스트 형식으로 보존한 값이며 HWP 셀 출력의 기준입니다. `raw_value`와 분리해 `(3,232)`를 `-3232`로 재해석하지 않고 실제 표시값 `-3232`는 그대로 유지합니다.

package 생성 시 HTML entity를 해제하고 공백을 정리합니다. 괄호, 쉼표, 마침표, `%`, `-`, `~`, 따옴표 등 구두점과 수학·통화 기호는 유지합니다. 단위 기호 `℃ ° ㎡ ㎢ ㎥ ㎏ ㎎ ㎞ ㎝ ㎜ ㏄`도 유지합니다. `● ■ ◆ ▶ ※`, 허용 단위 외 Unicode `So` 기호, 제어/형식 문자, variation selector와 emoji 피부색 modifier는 제거합니다. 제어 공백은 공백으로 바꿉니다.

정리된 셀에는 `original_display_text`, `removed_symbols`와 원본 셀 주소를 포함한 QA warning을 남깁니다. 기호만 있던 셀은 빈 문자열이 되며 원시값으로 되채우지 않습니다. 값 영역 전체가 비면 별도의 QA error가 발생할 수 있습니다.

### 검증 상태

분석문 연결, 괄호 표시값 변환, 장식 기호 정리, 재사용 가능한 역할별 서식은 구현 및 자동 테스트를 마쳤습니다. 2026-10-01 실제 KISDI 3개 표(35×5, 36×36, 35×10)를 clipboard 경로로 생성하고 아래한글 재열기·XML 비교·PDF 내보내기를 검증했습니다. 병합 종속 셀을 제외한 1,691개 셀의 표시문자·병합·역할별 스타일과 독립 표 3개를 확인했으며 원본 Excel/템플릿은 변경하지 않았습니다.

결과는 로컬 `outputs/hwp_style_regression_20261001/`의 `KISDI_three_tables_styled.hwpx`, `KISDI_three_tables_preview.pdf`, `hwp_writer_report.json`, `verification.json`에 보관합니다. 원자료와 결과 파일은 Git에 포함하지 않습니다. 표별 `table_results[].insert_mode`와 `fallback_reason`으로 clipboard와 계약 기반 fallback 경로를 구분하며, 특정 자료의 통과를 모든 Excel/한글 환경의 지원으로 확대하지 않습니다.

**남은 레이아웃 제한:** 붙여넣기는 Excel의 열 너비를 유지합니다. 36열처럼 넓은 표는 본문 폭을 초과해 PDF에서 오른쪽 열이 잘릴 수 있으며 자동 축소·가로 분할은 아직 지원하지 않습니다. 셀별 폭 조정 시험에서는 병합 표의 이동이 중단되어 정식 경로에 적용하지 않았습니다. 현재 결과는 데이터·서식 검증용 초본이며 인쇄 완성본이 아닙니다.

## 아직 개발 중인 기능

다음 기능은 계획과 기반 작업은 있으나, 현재 알파에서 완성 기능으로 열지 않습니다.

- 런처에서 PPTX 보고서 직접 생성
- 런처에서 차트 검토 PPTX 직접 생성
- HWP 바이너리 직접 편집
- HWPX 템플릿의 반복 결과 블록을 복제해 실제 보고서 표 서식을 보존한 삽입
- HWPX 템플릿에 차트를 EMF/한글 차트 객체로 자동 삽입
- PPTX 템플릿 디자인을 보존한 placeholder 기반 정밀 치환

다음 개발 계획은 `report_automation_launcher/NEXT_DOCUMENT_OUTPUT_PLAN.md`를 기준으로 진행합니다.
HWP/HWPX 보고서틀의 실제 표 삽입 구조는 `docs/hwp_template_table_insertion_plan.md`를 기준으로 진행합니다.

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

### 대시보드 PPTX 작업 템플릿

기업/기관 대시보드 PPTX는 사용자가 편집한 첫 슬라이드의 `RA_DASH_*` shape 위치와 글꼴을 재사용할 수 있습니다. KPI는 항목명과 값을 반드시 별도 텍스트 상자로 두는 것을 권장합니다.

```text
RA_DASH_KPI_1_LABEL
RA_DASH_KPI_1_VALUE
...
RA_DASH_KPI_6_LABEL
RA_DASH_KPI_6_VALUE
RA_DASH_NARRATIVE_TEXT
RA_DASH_CHART_1_TITLE
...
RA_DASH_CHART_4_TITLE
```

위 이름이 없으면 카드 내부 텍스트 상자를 추정해 사용하지만, 생성된 결과물에서는 항목명과 값이 다시 분리된 텍스트 상자로 삽입됩니다.

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

PPTX 초본 생성:

```powershell
python -m report_automation_engine.document_writer `
  --package "C:\path\report_package.json" `
  --preflight "C:\path\preflight_report.json" `
  --type chart_review `
  --output "C:\path\chart_review_draft.pptx"
```

HWPX 초본 생성:

```powershell
python -m report_automation_engine.hwp_com_writer `
  --package "C:\path\report_package.json" `
  --preflight "C:\path\preflight_report.json" `
  --template "C:\path\report_template.hwpx" `
  --output "C:\path\report_draft.hwpx" `
  --style-config "C:\path\hwp_style_config.json" `
  --table-style-profile "C:\path\hwp_table_style_profile.json" `
  --dispatch-mode dispatch `
  --visible false
```

HWPX writer는 아래한글 COM을 사용하므로 Windows와 아래한글 설치가 필요합니다. 실패 시 출력 파일 옆 또는 `--report-output` 경로에 `hwp_writer_report.json`을 남깁니다.
writer의 `--style-config`를 생략하면 기본 JSON을 사용합니다. `--table-style-profile`은 기존 템플릿 표 속성 요약/적용 계획용이며 이름 있는 스타일 프리셋을 대신하지 않습니다. HWPX 작업 사본에 스타일을 등록/갱신하고 HWP 템플릿은 임시 HWPX로 변환해 같은 처리를 합니다. 설정 검증 또는 스타일 등록 실패 시 기본 서식으로 조용히 대체하지 않고 생성 실패를 기록합니다.

런처 CLI에서는 `--output-type "HWPX 보고서" --hwp-template "C:\path\report_template.hwpx" --hwp-style-config "C:\path\custom_style.json"`을 기존 Excel 실행 인자에 추가합니다. 입력 JSON은 선택 실행 파일 `hwp_style_config.json`으로 정규화됩니다. `--hwp-style-preset`만으로 저장된 사용자 서식을 불러오지는 않으며 실제 설정 JSON의 `preset_name`이 우선합니다.
실사용 환경에서 COM 실행이 지연되면 `writer_report.json`의 `stage`, `action`, `com.current_prog_id`, `com.steps`로 멈춘 지점을 확인합니다.

아래한글 COM 환경만 먼저 확인할 수도 있습니다.

```powershell
python -m report_automation_engine.hwp_com_writer `
  --check-environment `
  --dispatch-mode dispatch `
  --report-output "C:\path\hwp_writer_report.json"
```

HWPX 직접 writer 최소 초본 생성:

```powershell
python -m report_automation_engine.hwp_direct_writer `
  --package "C:\path\report_package.json" `
  --preflight "C:\path\preflight_report.json" `
  --template "C:\path\template.hwpx" `
  --output "C:\path\report_draft.hwpx"
```

Python 의존성 설치 예:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip setuptools wheel
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

런처는 저장소 루트의 `.venv\Scripts\python.exe`를 우선 사용합니다. 아래한글 COM writer를 사용할 때는 이 가상환경에 `pywin32`가 설치되어 있어야 합니다.

```powershell
python -m pip install -r requirements.txt
```

## 버전 관리

현재 버전은 `VERSION` 파일에 기록합니다.

정책:

- semantic versioning 사용: `MAJOR.MINOR.PATCH`
- push/publish 요청 시 버전 확인
- 코드, 문서, 빌드 산출물이 바뀌면 필요한 경우 버전 증가
- 릴리스 기준점은 Git tag로 기록: `v0.0.1`, `v0.0.2`, ...

자세한 정책은 `VERSIONING.md`를 참고합니다.

## 라이선스

이 프로젝트는 `ReportAutomation Personal and Internal Use License`를 사용합니다.

허용:

- 개인 개발, 개인 사용, 테스트
- 회사/기관 내부 개발과 내부 업무 자동화

금지:

- 이 코드를 포함하거나 변형한 유료 프로그램 판매
- 유료 add-in, 유료 템플릿, 유료 SaaS, 상업용 패키지로 제공
- 유료 납품물에 이 코드 또는 실질적 파생물을 포함해 제3자에게 제공

이 라이선스는 OSI 기준 오픈소스 라이선스가 아니라 source-available 성격의 제한 라이선스입니다.

### 외부 HWP/HWPX 도구 검토

`edwardkim/rhwp`는 장기적으로 HWP/HWPX 구조 분석과 대체 writer 후보로 검토합니다. 현재 저장소에는 rhwp 코드를 포함하지 않습니다. 향후 clone, submodule, vendoring, WASM/npm 의존성으로 포함할 경우 MIT 라이선스 고지와 제3자 라이선스 문서를 별도로 추가해야 하며, rhwp의 MIT 권리와 이 프로젝트의 custom license를 분리 표기해야 합니다.

## 개발 방향

현재 신규 개발 우선순위는 HWPX 표 안정화입니다. PPTX 보고서와 대시보드 PPTX는 기존 기능을 유지하지만, 새 기능 개발은 HWPX 표 계약과 writer 안정화 이후로 둡니다.

현재 안정화 우선순위:

1. Excel 산출 시트 안정화
2. `report_package.json` 계약 안정화
3. `preflight_report.json` 차단/경고 정확도 개선
4. HWPX COM writer 안정화
5. 차트 검토 PPTX writer
6. PPTX 보고서 writer
