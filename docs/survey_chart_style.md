# 조사 보고서 PPTX 차트 서식

1. 런처에서 **보고서 설정 → PPTX 차트 서식 → 편집용 기본 PPTX 저장**을 선택합니다.
2. PowerPoint에서 세로 막대와 도넛의 글꼴, 색, 막대 간격, 축, 범례, 데이터 레이블을 편집합니다. 각 차트의 첫 번째 보기/조각을 기본색으로, 두 번째를 최대값 강조색으로 설정합니다. 차트 객체 이름 `RA_CHART_COLUMN`, `RA_CHART_DOUGHNUT`, 차트 종류와 데이터 계열 한 개는 유지합니다.
3. **편집한 PPTX 등록**을 누릅니다. 등록 파일은 로컬의 `%LOCALAPPDATA%\ResearchHelper\survey_chart_style.pptx`에 복사되고 모든 프로젝트에 적용됩니다. **기본 서식으로 되돌리기**는 등록 사본만 삭제합니다.
4. 집계표에서 Excel 산출 시트 또는 한글 보고서를 실행한 뒤 **산출 결과 → 차트 검토 PPTX 만들기**를 누릅니다. 한 슬라이드에 차트 한 개를 넣고, 문항 제목을 차트 밖에 표시합니다. 차트 데이터는 PowerPoint에서 편집할 수 있습니다.

`전체` 후보가 있는 집계표에 대표 차트 한 개를 생성합니다. 응답자 특성표를 제외하며, 비율을 다른 지표보다 우선하고 보기 순서는 원본 셀 주소 기준입니다. 확실한 이분형은 도넛, 그 외는 세로 막대입니다. 전체 값 누락, 숫자 누락, 단일 비교값은 검토 슬라이드에 표시합니다.

공통 서식은 슬라이드 전체나 기존 업무용 PPTX를 복제하는 기능이 아닙니다. 예제 데이터·특정 보기에 직접 입력한 레이블 문구는 가져오지 않습니다. 지정한 축 범위가 새 값을 잘라내면 해당 경계를 자동으로 바꿉니다. 글꼴은 해당 PC에 설치되어야 합니다.

## 한글 차트 삽입

- 한글 출력에서 **한글에 차트 삽입**은 기본으로 켜져 있습니다. **그림 번호 앞자리**는 사용자가 지정합니다. 예: `3` → `<그림 3-1>`. 차트 데이터 체크를 끄면 자동 삽입도 꺼집니다.
- 분석문과 집계표 사이에 테두리 없는 1×1 표를 넣고, 위에는 그림 캡션을 달며 셀 안의 그림에는 번호를 매기지 않습니다. 자리와 그림은 페이지 안쪽 크기에 맞춥니다.
- 자동 삽입은 PowerPoint에서 **차트 객체만 EMF로 내보내고**, 한글 COM에서 셀 안에 그림으로 넣습니다. 일반 메타파일 붙여넣기가 WMF로 변환되는 것과 달리, 이 경로는 저장 후 원본 EMF와 바이트가 일치하는지 확인합니다. PowerPoint·한글 설치가 필요합니다.
- 값 누락, 긴 보기명, PowerPoint 준비 실패 시 그림 캡션과 빈 셀을 남깁니다. 산출 결과에 보류 사유를 표시합니다. 기존 문서를 수동으로 보정할 때는 PowerPoint 차트만 EMF로 저장해 해당 셀에서 **그림 넣기**를 사용하고, 개체 번호 종류를 **없음**으로 설정하세요. 이 환경의 **골라 붙이기 → 그림 (메타파일)**은 WMF로 저장됩니다.
- 차트별 선택 화면과 실제 렌더링 기반 겹침 검사는 아직 없습니다. 보수적인 보기 길이/개수 기준으로 검토 필요를 분류하고, 차트 검토 PPTX에서 확인할 수 있습니다.

```powershell
.venv/Scripts/python.exe -m report_automation_engine.survey_chart_style create --output 'basic chart style.pptx'
.venv/Scripts/python.exe -m report_automation_engine.survey_chart_style register --input 'edited chart style.pptx' --output 'registered style.pptx'
.venv/Scripts/python.exe -m report_automation_engine.survey_chart_style review --input 'report_package.json' --style 'registered style.pptx' --output 'chart review.pptx'
```
