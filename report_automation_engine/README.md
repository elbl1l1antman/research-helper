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

- `document_writer.py`
  - `report_package.json`을 읽어 PPTX 초본을 생성합니다.
  - PowerPoint에서 편집 가능한 차트 객체와 표 객체를 생성합니다.

- `hwp_com_writer.py`
  - `report_package.json`을 읽어 아래한글 COM으로 HWPX 초본을 생성합니다.
  - 원본 템플릿은 수정하지 않고 출력 경로에 사본을 만든 뒤 `{{BODY}}` 위치에 제목, 분석문, 표, 출처를 삽입합니다.
  - 실패 시 `hwp_writer_report.json`에 실패 단계, COM action, placeholder 상태, 경고를 기록합니다.

- `hwp_com_smoke.py`
  - 아래한글 COM으로 최소 HWPX 템플릿과 샘플 package/preflight를 만든 뒤 실제 writer를 실행합니다.
  - 생성된 HWPX 내부 XML에서 본문 텍스트 삽입과 `{{BODY}}` 제거 여부를 확인합니다.

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
  --visible false
```

주요 동작:

- `--check-environment`로 pywin32와 아래한글 COM 객체 생성 가능 여부를 먼저 확인할 수 있습니다.
- `preflight.status == blocked`이면 아래한글을 열기 전에 중단합니다.
- 템플릿 사본을 출력 경로에 만든 뒤 사본만 수정합니다.
- `{{BODY}}`를 찾지 못하면 생성하지 않고 writer report에 실패 사유를 남깁니다.
- HWP 표 객체 생성을 우선 시도하고, COM action이 실패하면 탭 구분 텍스트 표로 대체합니다.
- 차트는 v1에서 직접 삽입하지 않고 `[차트 삽입 필요]` 문구로 표시합니다.

HWPX writer 회귀 확인은 다음 명령으로 실행합니다.

```powershell
python -m report_automation_engine.hwp_com_smoke `
  --output-dir "outputs\hwp_com_smoke" `
  --visible false
```

성공하면 `minimal_template.hwpx`, `report_package.json`, `preflight_report.json`, `draft_output.hwpx`, `hwp_writer_report.json`, `hwp_com_smoke_report.json`이 생성됩니다. 이 테스트는 Windows, 아래한글, pywin32가 모두 준비된 환경에서만 통과합니다.

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
