# 조사 보고서 PPTX 차트 서식

1. 런처에서 **보고서 설정 → PPTX 차트 서식 → 편집용 기본 PPTX 저장**을 선택합니다.
2. PowerPoint에서 세로 막대와 도넛의 글꼴, 색, 막대 간격, 축, 범례, 데이터 레이블을 편집합니다. 각 차트의 첫 번째 보기/조각을 기본색으로, 두 번째를 최대값 강조색으로 설정합니다. 차트 객체 이름 `RA_CHART_COLUMN`, `RA_CHART_DOUGHNUT`, 차트 종류와 데이터 계열 한 개는 유지합니다.
3. **편집한 PPTX 등록**을 누릅니다. 등록 파일은 로컬의 `%LOCALAPPDATA%\ResearchHelper\survey_chart_style.pptx`에 복사되고 모든 프로젝트에 적용됩니다. **기본 서식으로 되돌리기**는 등록 사본만 삭제합니다.
4. 집계표에서 Excel 산출 시트 또는 한글 보고서를 실행한 뒤 **산출 결과 → 차트 검토 PPTX 만들기**를 누릅니다. 한 슬라이드에 차트 한 개를 넣고, 문항 제목을 차트 밖에 표시합니다. 차트 데이터는 PowerPoint에서 편집할 수 있습니다.

`전체` 후보가 있는 집계표에 대표 차트 한 개를 생성합니다. 응답자 특성표를 제외하며, 비율을 다른 지표보다 우선하고 보기 순서는 원본 셀 주소 기준입니다. 확실한 이분형은 도넛, 그 외는 세로 막대입니다. 전체 값 누락, 숫자 누락, 단일 비교값은 검토 슬라이드에 표시합니다.

공통 서식은 슬라이드 전체나 기존 업무용 PPTX를 복제하는 기능이 아닙니다. 예제 데이터·특정 보기에 직접 입력한 레이블 문구는 가져오지 않습니다. 지정한 축 범위가 새 값을 잘라내면 해당 경계를 자동으로 바꿉니다. 글꼴은 해당 PC에 설치되어야 합니다.

## 기존 그래프틀에서 추출

1. **보고서 설정 → PPTX 차트 서식 → 그래프틀에서 CRTX 추출**에서 기존 PPTX와 출력 폴더를 고릅니다.
2. 새 하위 폴더의 `chart_style_inventory.csv`에서 슬라이드·개체 번호, 차트 종류, 계열 수, 글꼴과 크기, 막대 간격·겹침, 축과 범례 설정, 결과 파일을 비교합니다. `chart_style_inventory.json`에는 색상 참조, 숫자 서식, 테마 재정의 여부도 기록됩니다. `slide_002_shape_003.crtx`는 2쪽의 개체 ID 3입니다.
3. `prototype`이 `RA_CHART_COLUMN` 또는 `RA_CHART_DOUGHNUT`이고 추출 상태가 `exported`인 서식을 **CRTX로 기본 PPTX 만들기**에서 고릅니다. 결과 파일을 확인한 뒤 **편집한 PPTX 등록**을 누릅니다. 새 PPTX를 만들기만 하면 현재 등록 서식은 그대로입니다.

PowerPoint의 CRTX에는 차트 내부의 축·간격·색·글꼴·범례·레이블 서식을 보존합니다. 원본 값, 직접 입력한 차트 제목·레이블 문구, 차트 밖의 개별 도형과 슬라이드 배치는 가져오지 않습니다. 그룹에 포함된 편집 가능한 차트는 읽지만, 그림과 외부 연결 차트는 추출하지 않습니다. 추출과 CRTX 적용에는 Windows와 PowerPoint가 필요합니다.

누적·가로 막대·꺾은선·복합·다계열도 Office에서 재사용할 CRTX로 추출할 수 있습니다. 현재 자동 보고서의 기본 슬롯은 단일 계열 세로 막대와 도넛만 지원합니다. 한 유형만 가져오면 다른 유형은 현재 등록 서식 또는 내장 기본형을 유지합니다. 원본에서 숨긴 세로 막대의 항목 축은 기본 PPTX에서 다시 표시하며, 항목명이 들어갈 자리는 자동 배치합니다. 글꼴은 사용 PC에 설치되어 있어야 합니다.

## 검토한 차트를 EMF로 저장

1. **차트 검토 PPTX 만들기**로 생성한 파일을 PowerPoint에서 편집하고 저장합니다. 줄바꿈, 글꼴, 색, 막대 간격 등을 바꿀 수 있습니다. `RA_SURVEY_`로 시작하는 차트 개체 이름과 데이터 계열 1개는 유지합니다. 필요 없는 차트의 슬라이드는 삭제해도 됩니다.
2. 런처 **산출 결과 → 편집한 차트 EMF 저장**에서 편집본과 저장 폴더를 선택합니다. 런처를 다시 열었다면 해당 PPTX를 만든 실행의 `report_package.json`도 지정합니다. 이전 버전의 공통 이름만 있는 검토 PPTX는 자동 매칭하지 않습니다.
3. 현재 집계표의 수치·보기명·순서와 대조합니다. 줄바꿈과 공백만 무시합니다. 수치 변경, 이름 중복, 링크된 차트는 제외합니다. 선택 폴더 안 새 하위 폴더에 문항별 EMF와 `emf_export_report.json`을 저장합니다. 슬라이드 제목이나 별도 도형은 내보내지 않습니다.
4. 한글의 해당 1×1 그림 자리에서 **그림 넣기**로 EMF를 삽입합니다. 기존 차트가 있으면 그 그림만 교체하고 위쪽 캡션은 유지합니다. 새 그림의 번호 종류는 **없음**으로 설정합니다.

원본 PPTX와 HWPX는 건드리지 않습니다. EMF 내보내기에는 Windows와 PowerPoint가 필요합니다. 직접 입력한 데이터 레이블의 텍스트, 단위 표시, 글꼴 대체, 보기 겹침은 EMF를 눈으로 확인하세요.

## 한글 차트 삽입

- 한글 출력에서 **한글에 차트 삽입**은 기본으로 켜져 있습니다. **그림 번호 앞자리**는 사용자가 지정합니다. 예: `3` → `<그림 3-1>`. 차트 데이터 체크를 끄면 자동 삽입도 꺼집니다.
- 분석문과 집계표 사이에 테두리 없는 1×1 표를 넣고, 위에는 그림 캡션을 달며 셀 안의 그림에는 번호를 매기지 않습니다. 자리와 그림은 페이지 안쪽 크기에 맞춥니다.
- 차트가 삽입된 문항의 집계표를 가로 분할해야 하면 첫 분할표부터 다음 쪽에 배치합니다. 차트가 없거나 표가 한 조각이면 이 쪽 나눔을 적용하지 않습니다.
- 자동 삽입은 PowerPoint에서 **차트 객체만 EMF로 내보내고**, 한글 COM에서 셀 안에 그림으로 넣습니다. 일반 메타파일 붙여넣기가 WMF로 변환되는 것과 달리, 이 경로는 저장 후 원본 EMF와 바이트가 일치하는지 확인합니다. PowerPoint·한글 설치가 필요합니다.
- 값 누락, 긴 보기명, PowerPoint 준비 실패 시 그림 캡션과 빈 셀을 남깁니다. 산출 결과에 보류 사유를 표시합니다. 기존 문서를 수동으로 보정할 때는 PowerPoint 차트만 EMF로 저장해 해당 셀에서 **그림 넣기**를 사용하고, 개체 번호 종류를 **없음**으로 설정하세요. 이 환경의 **골라 붙이기 → 그림 (메타파일)**은 WMF로 저장됩니다.
- 차트별 선택 화면과 실제 렌더링 기반 겹침 검사는 아직 없습니다. 보수적인 보기 길이/개수 기준으로 검토 필요를 분류하고, 차트 검토 PPTX에서 확인할 수 있습니다.

```powershell
.venv/Scripts/python.exe -m report_automation_engine.survey_chart_style create --output 'basic chart style.pptx'
.venv/Scripts/python.exe -m report_automation_engine.survey_chart_style register --input 'edited chart style.pptx' --output 'registered style.pptx'
.venv/Scripts/python.exe -m report_automation_engine.survey_chart_style review --input 'report_package.json' --style 'registered style.pptx' --output 'chart review.pptx'
.venv/Scripts/python.exe -m report_automation_engine.survey_chart_style export-reviewed --input 'edited review.pptx' --package 'report_package.json' --output 'outputs'
.venv/Scripts/python.exe -m report_automation_engine.chart_style_library extract --input 'existing charts.pptx' --output 'outputs'
.venv/Scripts/python.exe -m report_automation_engine.chart_style_library make-default --input 'slide_002_shape_003.crtx' --output 'imported chart style.pptx'
```
