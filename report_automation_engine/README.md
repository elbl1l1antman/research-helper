# Report Automation Engine

이 폴더는 보고서 자동화 프로젝트의 Python 보조 엔진입니다.

현재 역할은 기존 VBA add-in을 대체하는 것이 아니라, 엑셀/HWPX 집계표를 읽어 보고서 본문 초안을 생성하는 것입니다. VBA add-in은 엑셀 파일 안에 `보고서_분석문`, `보고서_차트데이터`, `보고서_삽입표` 같은 산출 시트를 만드는 역할을 계속 맡고, 이 엔진은 외부 EXE 런처에서 선택적으로 호출합니다.

## 파일 구성

- `style_config_loader.py`
  - 문장 패턴, 제외 키워드, 표 해석 규칙을 JSON에서 읽습니다.
  - 코드 수정 없이 문체와 일부 작성 규칙을 바꾸기 위한 레이어입니다.

- `excel_report_generator.py`
  - 엑셀 집계표를 읽어 표 블록을 탐지하고, 전체/지역/응답자 특성별 본문 초안을 생성합니다.
  - 런처가 만든 `보고서_분석문*` 시트가 있으면 해당 시트의 최종 분석문 열을 우선 사용해 TXT 초안을 빠르게 생성합니다.
  - 제공받은 `excel_report_generator_with_style.py`를 프로젝트용으로 가져오면서 기본 설정 fallback과 주석을 보강했습니다.

- `report_package.py`
  - Excel 산출 시트를 헤더명 기반으로 읽어 `report_package.json`과 `preflight_report.json`을 생성합니다.
  - HWPX/PPTX 문서 생성 전 문장, 표, 차트, QA, 템플릿 상태를 `ready`, `ready_with_warnings`, `blocked`로 검증합니다.
  - QA warning을 `normal_review_warning`, `improvement_needed`, `info` bucket으로 분류합니다.
  - `narrative_final`을 보존하며 줄바꿈과 지정된 기존 연결 경계를 `narrative_blocks`로 분리합니다. 첫 블록은 `보고서 본문1`, 모든 후속 블록은 `보고서 본문2`입니다.

- `report_table_matrix.py`
  - `report_package.json`의 삽입표 데이터를 HWPX writer가 사용할 행/열 matrix 계약으로 정규화합니다.
  - 셀 역할, 표시값, 원본 값, 정렬, style hint, table matrix QA를 함께 생성합니다.
  - `display_text`만 정리하고 변경 전 표시문자와 제거 기호, 원본 셀 주소 QA를 남깁니다.

- `document_writer.py`
  - `report_package.json`을 읽어 PPTX 초본을 생성합니다.
  - PowerPoint에서 편집 가능한 차트 객체와 표 객체를 생성합니다.

- `hwp_com_writer.py`
  - `report_package.json`을 읽어 아래한글 COM으로 HWPX 초본을 생성합니다.
  - 원본 템플릿은 수정하지 않고 출력 경로에 사본을 만든 뒤 `{{BODY}}` 위치에 제목, 분석문, 표, 출처를 삽입합니다.
  - 실패 시 `hwp_writer_report.json`에 실패 단계, COM action, placeholder 상태, 경고를 기록합니다.
  - `--dry-run`, `--max-sections`, `--render-plan-output`으로 실제 아래한글 실행 전 삽입 계획을 확인할 수 있습니다.
  - `--table-style-profile`로 템플릿에서 인식한 표 서식 profile을 읽고, v1에서는 대표 글자 크기를 우선 적용합니다.
  - `--style-config`로 이름 있는 본문/표 스타일과 셀 외형을 적용합니다. 기존 profile의 대표 글자 크기는 표 생성 전 설정이며 셀 글꼴·문단 속성은 이름 있는 스타일 설정으로 적용합니다.

- `hwp_style_config.py`, `config/default_hwp_style_config.json`
  - `schema_version: "1.0"`의 HWP 서식 JSON을 검증하고 기본 `기본 보고서` 서식을 제공합니다. 문장 패턴용 `default_style_schema.json`과는 별도입니다.

- `hwpx_style_registry.py`
  - 원본을 보존한 작업용 HWPX의 `Contents/header.xml`에 고정 스타일 5개를 등록/갱신합니다. 같은 이름을 중복 등록하지 않고 기존 필수 이름의 중복도 정리합니다.

- `hwp_direct_writer.py`
  - `report_package.json`과 `preflight_report.json`을 읽어 HWPX zip 패키지 내부의 `{{BODY}}` placeholder를 최소 XML payload로 치환합니다.
  - direct writer v1은 자동 검증을 위한 최소 HWPX package payload를 쓰는 경로입니다.
  - 시각적 충실도는 table XML validation으로 표 계약이 안정화된 뒤 강화합니다.

- `hwp_com_smoke.py`
  - 아래한글 COM으로 최소 HWPX 템플릿과 샘플 package/preflight를 만든 뒤 실제 writer를 실행합니다.
  - 생성된 HWPX 내부 XML에서 본문 텍스트 삽입과 `{{BODY}}` 제거 여부를 확인합니다.

- `hwp_com_demo.py`
  - 사용자가 직접 열어볼 수 있는 한 장짜리 HWPX 데모 보고서를 생성합니다.
  - fake survey package를 만들고, HWPX 템플릿 생성부터 writer 실행과 미리보기 TXT 생성까지 한 번에 수행합니다.

- `hwp_template_probe.py`
  - HWP/HWPX 보고서틀을 분석해 표 객체, 주변 문단, 결과표 후보를 JSON으로 추출합니다.
  - HWP 원본은 수정하지 않고 아래한글 COM으로 HWPX 사본을 만든 뒤 분석합니다.

- `template_blueprint.py`
  - `hwp_template_probe_report.json`을 읽어 문항별 반복 결과 블록 후보를 점수화합니다.
  - 표 제목, 분석문, 그림/차트 캡션, Base/단위, 결과표의 반복 구조를 writer용 blueprint로 정규화합니다.

- `hwp_template_table_recognizer.py`
  - 사용자가 제공한 HWP/HWPX 보고서 템플릿에서 자동화에 사용할 표 후보를 인식합니다.
  - 결과표 후보, 스타일 원본 표, 레이아웃용 표를 분리한 사용자/런처용 JSON 리포트를 생성합니다.
  - `hwp_table_style_report.json`과 `hwp_table_style_profile.json`에 표 선, 셀 배경, 글자 크기, 셀 여백 요약을 저장합니다.

- `hwp_table_mapping.py`
  - `report_package.json`의 section/table과 `hwp_template_table_recognition.json`의 템플릿 표 후보를 연결합니다.
  - 다음 HWPX writer가 직접 읽을 `hwp_table_mapping.json` 계약을 생성하고, 누락 표/문장/템플릿 후보를 차단합니다.

- `dashboard_package.py`
  - 기관/기업 1행, 지표 여러 열의 가로형 Excel 원자료를 읽어 `dashboard_package.json`과 `dashboard_preflight_report.json`을 생성합니다.
  - Excel 검사 모드에서는 sheet, 열, 예시값, 열 유형, 상위 30행 미리보기를 JSON으로 저장합니다.

- `dashboard_writer.py`
  - `dashboard_package.json`을 읽어 기업/기관별 세로형 A4/B5 대시보드 PPTX를 생성합니다.
  - KPI 카드는 텍스트/도형, 차트는 PowerPoint에서 편집 가능한 chart object로 생성합니다.
  - mapping의 `style_preset`과 `font_family`로 디자인 프리셋과 본문 폰트를 지정합니다.
  - `--template`을 지정하면 사용자가 편집한 PPTX 첫 슬라이드의 `RA_DASH_*` 위치와 폰트를 재사용합니다.

- `hwpx_report_writer.py`
  - HWPX 내부 XML을 읽어 표와 문단 흐름을 분석합니다.
  - 현재는 HWPX에 직접 삽입하기보다, 기존 HWPX 표 구조를 분석하고 문장 생성 로직을 검증하는 보조 도구로 봅니다.

- `template_inspector.py`
  - HWPX/PPTX/HWP 템플릿의 placeholder와 `RA_` shape/bookmark 후보를 검사합니다.
  - 검사 결과를 `ready`, `usable_with_warnings`, `needs_autofix`, `unsupported` 상태 JSON으로 저장합니다.

- `template_factory.py`
  - 사용자가 디자인만 바꿔 쓸 수 있는 기본 HWPX/PPTX 템플릿 파일을 생성합니다.
  - 기본 PPTX는 OpenXML 기반의 편집 가능한 시작 파일입니다.

- `template_autofix.py`
  - 기존 템플릿 원본을 보존하고 `_template_ready` 사본에 최소 placeholder를 삽입합니다.
  - PPTX는 슬라이드 XML에 텍스트 박스를 추가하고, HWPX는 자동 보정용 XML 파트를 추가합니다.

- `config/default_style_schema.json`
  - 기본 문체/표 해석 설정입니다.
  - GUI에서 별도 JSON을 선택하지 않으면 이 파일을 사용합니다.

## 실행 예시

번들 Python 또는 일반 Python 환경에서 실행할 수 있습니다.

```powershell
python report_automation_engine\excel_report_generator.py
```

실행하면 스타일 JSON 선택 창이 먼저 뜹니다. 취소하면 `config/default_style_schema.json`을 사용합니다. 이후 분석할 엑셀 파일을 선택하면 원본 파일 옆에 `_자동생성.txt`가 생성됩니다.

런처 또는 자동 검증에서는 CLI 인자를 사용합니다.

```powershell
python report_automation_engine\excel_report_generator.py `
  --excel "C:\path\table.xlsx" `
  --config "C:\path\default_style_schema.json" `
  --output "C:\path\table_draft.txt" `
  --max-tables 30
```

주요 옵션:

- `--excel <path>`: 분석할 엑셀 파일입니다.
- `--config <path>`: 문체/표 해석 JSON입니다. 생략하면 기본 설정을 사용합니다.
- `--output <path>`: 생성할 TXT 경로입니다. 생략하면 원본 옆에 `_자동생성.txt`를 만듭니다.
- `--sheet <name>`: 원본 표 블록 직접 분석 시 특정 시트만 분석합니다.
- `--max-tables <count>`: 원본 표 블록 직접 분석 시 처리할 최대 표 수입니다.
- `--raw-tables`: `보고서_분석문*` 산출 시트를 무시하고 원본 표 블록을 직접 분석합니다.

Excel 산출 시트가 생성된 뒤에는 package/preflight를 만들 수 있습니다.

```powershell
python report_automation_engine\report_package.py `
  --excel "C:\path\table_report_alpha.xlsx" `
  --package-output "C:\path\report_package.json" `
  --preflight-output "C:\path\preflight_report.json"
```

HWPX 분석기는 명령행 인자를 받을 수 있습니다.

```powershell
python report_automation_engine\hwpx_report_writer.py "input.hwpx" -o "draft.txt"
```

템플릿 도구는 런처에서 호출하거나 CLI로 직접 검증할 수 있습니다.

```powershell
python report_automation_engine\template_inspector.py `
  --template "C:\path\user_template.pptx" `
  --type pptx_report `
  --output "C:\path\template_report.json"
```

```powershell
python report_automation_engine\template_factory.py `
  --type pptx_report `
  --output "C:\path\report_template_basic.pptx"
```

```powershell
python report_automation_engine\template_autofix.py `
  --template "C:\path\user_template.pptx" `
  --type chart_review `
  --output "C:\path\user_template_ready.pptx"
```

템플릿 최소 기준:

- HWPX 보고서: `{{BODY}}`
- PPTX 보고서: `{{SECTION_TITLE}}`, `{{NARRATIVE}}`, `{{TABLE}}`, `{{CHART}}`
- 차트 검토 PPTX: `{{CHART_TITLE}}`, `{{CHART}}`, `{{CHART_NOTE}}`

PPTX 초본 writer는 다음처럼 실행합니다.

```powershell
python -m report_automation_engine.document_writer `
  --package "C:\path\report_package.json" `
  --preflight "C:\path\preflight_report.json" `
  --type chart_review `
  --output "C:\path\chart_review_draft.pptx"
```

HWPX 초본 writer는 아래한글이 설치된 Windows 환경에서 실행합니다.
프로젝트 루트의 `.venv\Scripts\python.exe`에 `pywin32`를 설치하면 런처가 해당 가상환경을 우선 사용합니다.

```powershell
python -m report_automation_engine.hwp_com_writer `
  --check-environment `
  --report-output "C:\path\hwp_writer_report.json"
```

```powershell
python -m report_automation_engine.hwp_com_writer `
  --package "C:\path\report_package.json" `
  --preflight "C:\path\preflight_report.json" `
  --template "C:\path\report_template.hwpx" `
  --output "C:\path\report_draft.hwpx" `
  --render-plan-output "C:\path\hwp_render_plan.json" `
  --table-style-profile "C:\path\hwp_table_style_profile.json" `
  --style-config "C:\path\hwp_style_config.json" `
  --max-sections 1 `
  --visible false
```

사용자 제공 HWP/HWPX 보고서틀에서 반복 결과표 후보를 인식하고 package와 연결하려면 다음 순서로 실행합니다.

```powershell
python -m report_automation_engine.hwp_template_table_recognizer `
  --template "C:\path\report_template.hwpx" `
  --output-dir "C:\path\hwp_template_recognition"
```

```powershell
python -m report_automation_engine.hwp_table_mapping `
  --package "C:\path\report_package.json" `
  --recognition "C:\path\hwp_template_recognition\hwp_template_table_recognition.json" `
  --output "C:\path\hwp_table_mapping.json"
```

주요 동작:

- `--check-environment`로 pywin32와 아래한글 COM 객체 생성 가능 여부를 먼저 확인할 수 있습니다.
- `preflight.status == blocked`이면 아래한글을 열기 전에 중단합니다.
- `--dry-run`을 사용하면 아래한글을 열지 않고 `hwp_render_plan.json`과 writer report만 생성합니다.
- `--max-sections 1` 또는 `--max-sections 3`으로 알파 검증용 짧은 초본을 생성할 수 있습니다.
- `--table-style-profile`을 지정하면 render plan과 writer report에 style source, 대표 글자 크기, 배경색/선 요약이 기록됩니다.
- `--style-config`를 생략하면 `config/default_hwp_style_config.json`을 사용합니다. 지정 파일의 검증 실패는 `stage=style`, `action=load_style_config`로 기록하고 중단합니다.
- 임시 HWPX 작업 사본에 스타일을 등록한 뒤 COM으로 엽니다. HWP 템플릿은 COM으로 임시 HWPX를 저장한 뒤 같은 등록/재열기 흐름을 사용하며 원본은 수정하지 않습니다. 등록 실패도 생성 중단 사유입니다.
- `{{BODY}}`를 찾지 못하면 생성하지 않고 writer report에 실패 사유를 남깁니다.
- 셀 계약이 있으면 Excel 범위 clipboard 삽입을 먼저 시도합니다. 사용할 수 없거나 붙여넣기 검증이 실패하면 matrix 기반 `TableCreate`/셀 병합으로 전환합니다. 표 생성 실패 시 텍스트 표를 성공 산출물로 저장하지 않고 중단합니다.
- clipboard로 붙인 셀도 계약의 `display_text`로 다시 채우고 역할별 스타일과 외형을 적용합니다. 병합에 가려진 `covered_by` 셀은 건너뜁니다.
- 차트는 v1에서 직접 삽입하지 않고 `[차트 삽입 필요]` 문구로 표시합니다.

## HWP 서식 계약과 표시문자

필수 스타일 이름은 `보고서 본문1`, `보고서 본문2`, `표보기`, `표배너`, `표숫자`로 고정입니다. 기본 글꼴은 모두 `맑은 고딕`이며 정확한 기본값은 다음과 같습니다.

| 스타일 | 크기/굵게 | 줄간격 | 왼쪽/첫 줄 들여쓰기 | 정렬 | 글머리표 |
| --- | --- | --- | --- | --- | --- |
| 보고서 본문1 | 10pt/아니오 | 160% | 0/0mm | left | 없음 |
| 보고서 본문2 | 10pt/아니오 | 160% | 5/-5mm | left | `-` |
| 표보기 | 8.5pt/아니오 | 130% | 0/0mm | left | 없음 |
| 표배너 | 8.5pt/예 | 130% | 0/0mm | center | 없음 |
| 표숫자 | 8.5pt/아니오 | 130% | 0/0mm | center | 없음 |

셀 역할 `banner_horizontal`, `base`, `header`, `title`은 `표배너`, `value`는 `표숫자`, 나머지(`blank`, `unknown` 포함)는 `표보기`로 매핑합니다. `table_cell_styles`의 기본 배경은 `표배너=E7E7E7`, 나머지 `FFFFFF`, 세로 정렬은 모두 `center`입니다. `table_border`는 내부선 0.12mm, 외곽선 0.4mm, 색 `000000`이며 두께는 한글 지원 값 중 가장 가까운 값으로 변환합니다. 배경·선·세로 정렬은 문단 스타일 외에 COM 셀 action으로 적용합니다.

JSON 검증 범위는 글자 크기 6~30pt, 줄간격 80~300%, 두 들여쓰기 -30~100mm, 선 두께 0.1~5mm입니다. 색은 `#` 없는 6자리 RGB hex, 가로 정렬은 `left/center/right/justify`, 세로 정렬은 `top/center/bottom`입니다. 필수 스타일 누락/추가나 잘못된 값은 오류입니다.

분석문의 CRLF/LF/CR과 `나타남다음으로`, `나타남그다음으로`, `나타남반면,` 경계를 분리해 각 블록을 `BreakPara`로 삽입합니다. 본문2 말머리는 실제 `BULLET` 스타일 속성이며 선택한 말머리와 공백/탭이 붙은 원문 접두어만 제거합니다. 말머리가 비어 있으면 접두어를 제거하지 않으며 `-3232` 같은 음수 문자열도 유지합니다.

Excel 셀 계약의 `display_text`가 출력 기준이며 `raw_value`로 숫자 재변환하지 않습니다. `(3,232)`, 실제 음수 `-3232`, 빈칸, `0`, `-`를 구분합니다. package 생성 시 HTML entity를 해제하고 공백을 정리하되 구두점(괄호·쉼표·`%`·따옴표·`-`·`~` 등), 수학/통화 기호와 단위 `℃ ° ㎡ ㎢ ㎥ ㎏ ㎎ ㎞ ㎝ ㎜ ㏄`는 유지합니다. `● ■ ◆ ▶ ※`, 허용 단위 외 `So` 기호, Unicode `C` 범주 문자, variation selector와 emoji 피부색 modifier는 제거하고 제어 공백은 공백으로 바꿉니다.

변경 셀에는 `original_display_text`, `removed_symbols`와 원본 셀 주소 QA warning을 남깁니다. 기호만 있던 셀은 빈 문자열로 출력하며 원시값으로 복원하지 않습니다. 값 영역 전체가 비면 별도의 QA error가 발생할 수 있습니다.

GUI는 HWPX 옵션의 `서식`/`서식 설정`에서 선택·편집합니다. 내장 `기본 보고서`는 수정/삭제 불가, 저장 가능한 사용자 서식은 최대 3개이며 같은 이름은 명시적 덮어쓰기가 필요합니다. `새로 저장`/`덮어쓰기` 후 `선택 후 닫기`로 적용할 저장 서식을 선택합니다. 사용자 저장소 `%LOCALAPPDATA%\ResearchHelper\hwp_style_presets.json`이 손상되면 `.corrupt.<timestamp>.<id>.bak` 백업을 남기고 읽을 수 있는 유효 항목을 보존합니다. JSON을 읽을 수 없으면 내장 기본값으로 복구합니다.

런처는 선택 설정을 package 옆 `hwp_style_config.json`으로 저장한 뒤 writer에 `--style-config`로 전달합니다. 런처 CLI에서 `--hwp-style-config <path>`를 주면 해당 JSON의 `preset_name`이 사용되고, 경로 없이 `--hwp-style-preset <name>`을 주면 저장소에서 해당 서식을 조회합니다. 옵션을 생략하면 내장 기본값을 사용합니다. 저장소 복구 경고와 백업 경로는 GUI 로그 또는 CLI 오류 출력에 표시합니다. 실행 기록의 `HwpStylePreset`/`HwpStyleConfig`와 writer report의 `style_preset_name`/`style_config_path`/`style_index_map`/`style_application_counts`를 확인합니다. counts는 제목·출처 등을 포함한 성공한 스타일 적용 호출 횟수이며 순수 분석문 문단 수와 같지 않습니다.

본문 분리·표시값 보존·기호 정리·역할별 서식은 자동 테스트를 통과했습니다. 2026-10-01 실제 KISDI 3개 표를 clipboard 경로로 생성해 독립 표 3개, 표시문자 1,691개, 괄호형 값 99개, 병합 시작 셀 25개와 역할별 스타일을 XML로 대조하고 아래한글 재열기·PDF 내보내기를 확인했습니다. 결과와 근거는 로컬 `outputs/hwp_style_regression_20261001/`에 보관하며 원자료는 Git에 포함하지 않습니다. 표별 `insert_mode`, `style_status`, `fallback_reason`으로 실제 clipboard와 계약 기반 fallback 경로를 구분합니다. `--dry-run`은 COM 열기/스타일 등록/실제 표 생성 검증이 아닙니다.

표 편집 후에는 현재 표의 앵커 바로 뒤로 복귀합니다. 문서 끝으로 이동하지 않으므로 템플릿의 QA/후속 내용 앞에서 section 순서를 유지합니다. 세로 병합 셀이 다시 나타나는 경우 이미 처리한 셀 목록을 건너뛰며 이동 반복/실패는 생성을 중단합니다. Undo 실패 또는 셀 편집 중 오류가 나면 중복 표를 만들 수 있는 fallback을 실행하지 않습니다.

넓은 표는 `hwp_table_layout.py`에서 본문 폭에 맞는 독립 표로 가로 분할합니다. 보기·세로 배너·BASE 열은 반복하고, 가로 헤더 병합은 분할 경계에 맞춰 반복합니다. 본문 병합은 잘라내지 않으며 표현할 수 없으면 차단합니다. `matrix`가 없는 레거시 표에는 이 분할 정책을 적용하지 않습니다.

각 조각의 표시문자를 텍스트 형식의 임시 Excel 범위에 넣고, 실제 포인트 폭을 측정해 열 폭을 설정합니다. 원본 통합문서는 읽기 전용이며 임시 통합문서는 성공/실패 시 모두 닫습니다. 실제 한글 표 폭을 재검사하고 문단 들여쓰기와 무관하게 본문 단 왼쪽에 배치합니다. 다단 본문은 지원하지 않습니다.

`table_layouts[]`는 원본 표별 분할 계획입니다. `table_results[]`에는 `part_index`, `part_count`, `source_columns`, `body_width_hwpunit`, `actual_width_hwpunit`을 기록합니다. `tables_written`은 생성한 조각 수, `source_tables_written`은 완성한 원본 표 수입니다. 원본 package는 변경하지 않습니다.

Paste 검증은 `LastCtrl`이 아니라 전후 신규 개체 ID로 판정합니다. 무변경 실패에는 Undo를 실행하지 않고, 기존 개체 제거·불명확한 변경은 차단합니다. 배경색 적용 후 직접/선택 셀의 테두리를 적용해 Excel에서 가져온 무테 표에도 설정이 유지되도록 합니다. `TableCreate.ColWidth`에는 기본 좌우 셀 여백을 뺀 폭을 전달합니다.

실제 3개 집계표 → 11개 독립 표, 2,279개 표시문자·83개 병합·역할별 스타일·실선 테두리·본문 이내 폭·한글 재열기·PDF 출력 검증 근거는 `outputs/hwp_width_regression_20261001/verification.json`에 있습니다. 원본 열과 파일은 보존했습니다. 강제 fallback 214개 셀/병합도 대조했습니다. 긴 표의 세로 쪽 나눔 시 헤더/배너 반복과 GUI 분할 경계 편집은 후속 작업입니다.

HWPX writer 회귀 확인은 다음 명령으로 실행합니다.

```powershell
python -m report_automation_engine.hwp_com_smoke `
  --output-dir "outputs\hwp_com_smoke" `
  --visible false
```

성공하면 `minimal_template.hwpx`, `report_package.json`, `preflight_report.json`, `draft_output.hwpx`, `hwp_writer_report.json`, `hwp_com_smoke_report.json`이 생성됩니다. 이 테스트는 Windows, 아래한글, pywin32가 모두 준비된 환경에서만 통과합니다.

실제 자동화 결과를 눈으로 확인할 수 있는 한 장짜리 HWPX 데모는 다음처럼 생성합니다.

```powershell
python -m report_automation_engine.hwp_com_demo `
  --output-dir "outputs\hwp_com_demo" `
  --visible false
```

성공하면 `demo_hwp_report.hwpx`를 아래한글에서 열어 확인할 수 있습니다. 함께 생성되는 `demo_preview.txt`는 HWPX 내부에 삽입된 주요 텍스트만 빠르게 확인하기 위한 파일입니다.

HWP/HWPX 보고서틀의 표 구조를 분석할 때는 다음 명령을 사용합니다.

```powershell
python -m report_automation_engine.hwp_template_probe `
  --output-dir "outputs\hwp_template_probe" `
  "C:\path\report_template.hwpx" `
  "C:\path\legacy_template.hwp"
```

성공하면 `hwp_template_probe_report.json`과 `hwp_template_probe_summary.json`을 기준으로 템플릿 안의 결과표 후보, 목차/표지용 레이아웃 표, 그림/표 캡션 후보를 검토할 수 있습니다.

probe 결과에서 반복 결과 블록 후보를 만들 때는 다음 명령을 사용합니다.

```powershell
python -m report_automation_engine.template_blueprint `
  --probe "outputs\hwp_template_probe\hwp_template_probe_report.json" `
  --output "outputs\hwp_template_probe\template_blueprint.json"
```

생성된 blueprint는 HWPX writer가 실제 표를 삽입하기 전, 사용자가 어느 보고서틀 블록을 반복 사용할지 선택하는 중간 계약입니다.

사용자가 제공한 HWP/HWPX 템플릿에서 표 후보를 바로 인식하려면 다음 명령을 사용합니다.

```powershell
python -m report_automation_engine.hwp_template_table_recognizer `
  --template "C:\path\report_template.hwpx" `
  --output-dir "outputs\hwp_template_recognition"
```

성공하면 `hwp_template_table_recognition.json`에 결과표 후보, 스타일 원본 표 후보, 레이아웃 표 목록이 저장됩니다. HWP 입력은 원본을 수정하지 않고 HWPX 사본으로 변환해 분석합니다.

기업/기관 대시보드 PPTX는 원자료를 먼저 검사하고, 사용자가 선택한 데이터/매핑 JSON을 기준으로 생성합니다.

대시보드 mapping JSON에서 사용할 수 있는 디자인 옵션:

- `style_preset`: `modern_blue`, `modern_mint`, `graphite`
- `font_family`: `Malgun Gothic`, `Noto Sans CJK KR`, `Arial` 등 PowerPoint에서 사용할 글꼴명

대시보드 작업용 PPTX 템플릿을 직접 편집할 때는 항목명과 값을 한 텍스트 상자에 합치지 말고 별도 텍스트 상자로 둡니다.
자동화는 다음 shape 이름을 우선 인식합니다.

- KPI 항목명: `RA_DASH_KPI_1_LABEL` ~ `RA_DASH_KPI_6_LABEL`
- KPI 값: `RA_DASH_KPI_1_VALUE` ~ `RA_DASH_KPI_6_VALUE`
- 본문 요약: `RA_DASH_NARRATIVE_TEXT`
- 차트 카드 제목: `RA_DASH_CHART_1_TITLE` ~ `RA_DASH_CHART_4_TITLE`

기존 템플릿에 위 이름이 없으면 카드 내부의 텍스트 상자 위치를 추정해 사용하며, 생성 시에는 기존 카드 내부 텍스트를 제거한 뒤 항목/값을 분리된 텍스트 상자로 다시 삽입합니다.

```powershell
python -m report_automation_engine.dashboard_package `
  --excel "C:\path\company_data.xlsx" `
  --inspect-output "C:\path\dashboard_excel_inspect.json"
```

```powershell
python -m report_automation_engine.dashboard_package `
  --excel "C:\path\company_data.xlsx" `
  --selection "C:\path\dashboard_data_selection.json" `
  --mapping "C:\path\dashboard_mapping.json" `
  --page-size A4 `
  --package-output "C:\path\dashboard_package.json" `
  --preflight-output "C:\path\dashboard_preflight_report.json"
```

```powershell
python -m report_automation_engine.dashboard_writer `
  --package "C:\path\dashboard_package.json" `
  --preflight "C:\path\dashboard_preflight_report.json" `
  --template "C:\path\dashboard_template.pptx" `
  --output "C:\path\organization_dashboard.pptx"
```

## 코드리뷰 포인트

- 이 엔진은 아직 완성된 보고서 편집기가 아닙니다.
- 엑셀 표 구조가 프로젝트별로 달라질 수 있으므로, `config/default_style_schema.json`의 제외어와 키워드를 먼저 조정하는 방식이 안전합니다.
- HWPX 기본 템플릿 생성은 placeholder 검증용 시작 파일입니다. 한글에서 완전히 안정적으로 열리는 고품질 HWPX 생성은 한글 COM API 또는 정식 HWPX writer 레이어를 붙일 때 별도로 강화해야 합니다.
- `edwardkim/rhwp`는 장기 HWP/HWPX 구조 분석 및 대체 writer 후보입니다. 현재는 저장소에 포함하지 않으며, 도입 시 MIT 라이선스 고지와 제3자 라이선스 문서를 먼저 추가해야 합니다.
- 외부 EXE 런처와 연결할 때는 Python 스크립트를 별도 프로세스로 실행하고, 입력 파일/스타일 JSON/출력 경로를 인자로 넘기는 방식이 기존 VBA 매크로와 충돌이 가장 적습니다.
