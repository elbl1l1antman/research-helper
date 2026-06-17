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

- `hwpx_report_writer.py`
  - HWPX 내부 XML을 읽어 표와 문단 흐름을 분석합니다.
  - 현재는 HWPX에 직접 삽입하기보다, 기존 HWPX 표 구조를 분석하고 문장 생성 로직을 검증하는 보조 도구로 봅니다.

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

HWPX 분석기는 명령행 인자를 받을 수 있습니다.

```powershell
python report_automation_engine\hwpx_report_writer.py "input.hwpx" -o "draft.txt"
```

## 코드리뷰 포인트

- 이 엔진은 아직 완성된 보고서 편집기가 아닙니다.
- 엑셀 표 구조가 프로젝트별로 달라질 수 있으므로, `config/default_style_schema.json`의 제외어와 키워드를 먼저 조정하는 방식이 안전합니다.
- HWPX 직접 삽입은 한글 COM API 또는 HWPX XML 생성 레이어를 붙일 때 별도로 구현해야 합니다.
- 외부 EXE 런처와 연결할 때는 Python 스크립트를 별도 프로세스로 실행하고, 입력 파일/스타일 JSON/출력 경로를 인자로 넘기는 방식이 기존 VBA 매크로와 충돌이 가장 적습니다.
