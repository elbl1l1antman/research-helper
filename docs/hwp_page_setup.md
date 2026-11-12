# HWPX 생성 전 편집용지 설정

런처의 HWPX 옵션에서 **편집용지**를 열어 설정합니다. 기본은 **템플릿 설정 유지**입니다.
직접 지정할 때는 A4/B5/A3/Letter 또는 사용자 지정 크기, 세로/가로 방향과 다음 mm 값을 입력합니다.
B5 기본값은 ISO 규격 176 x 250mm입니다. JIS B5가 필요하면 사용자 지정으로 182 x 257mm를 입력합니다.

| 항목 | 직접 지정 초기값 |
| --- | ---: |
| 용지 너비 / 높이 | 210 / 297 |
| 위쪽 / 아래쪽 | 20 / 15 |
| 왼쪽 / 오른쪽 | 24.7 / 25 |
| 제본 | 0 |
| 머리말 / 꼬리말 | 10 / 15 |

## 적용 범위와 순서

1. 원본 템플릿을 작업 사본으로 열고 `{{BODY}}`를 찾습니다.
2. 직접 지정한 경우 현재 본문 구역에 `PageSetup`을 적용합니다. 다른 구역의 표지·편집용지는 유지합니다.
3. 적용값을 다시 읽어 요청값과 일치하는지 확인합니다. 실패하면 저장을 중단하고 writer 리포트에 기록합니다.
4. 실제 본문 폭을 읽어 열 폭과 가로 분할을 계산하고 표를 삽입합니다.

표 너비는 본문 폭에 맞추되 약 1.8mm의 안전 간격을 둡니다. 세로로 페이지 전체를 채우도록 표를 늘리지는 않습니다.
행 높이는 내용에 따라 늘어나며, 긴 표는 페이지를 넘어가고 BASE/헤더가 반복됩니다.
너무 많은 열은 기존 가로 분할 정책을 사용합니다. 편집용지를 변경해도 다단 본문 지원이 추가되지는 않습니다.

제본은 한쪽 편집의 왼쪽 제본으로 적용합니다. 본문 폭은 용지 폭에서 좌우·제본 여백을 뺀 값입니다.
가로 방향이면 용지 너비/높이를 계산에서 바꿔 사용합니다. 머리말·꼬리말 간격은 위/아래 여백에 포함되므로 다시 빼지 않습니다.
머리말 간격이 위 여백보다 크거나 꼬리말 간격이 아래 여백보다 큰 값, 음수·비유한 값, 본문 공간이 없는 값은 차단합니다.

원본 Excel과 템플릿은 수정하지 않습니다. 같은 구역에 있는 기존 요소는 편집용지 변경 영향을 받으므로 결과를 확인해야 합니다.
별도 구역을 가진 템플릿은 본문 이후 다른 구역의 여백까지 일괄 변경하지 않습니다.

## 설정 파일과 CLI

런처는 직접 지정한 설정을 package 옆의 `hwp_page_setup.json`에 저장해 writer에 전달합니다.
템플릿 유지 모드에서는 이 파일이 남아 있어도 전달하지 않습니다.
현재 창에서 설정을 유지하며, 별도 실행에서 재사용하려면 CLI의 `--hwp-page-setup`으로 파일을 지정합니다.

```json
{
  "paper_width_mm": 210,
  "paper_height_mm": 297,
  "orientation": "portrait",
  "top_mm": 20,
  "bottom_mm": 15,
  "left_mm": 24.7,
  "right_mm": 25,
  "gutter_mm": 0,
  "header_mm": 10,
  "footer_mm": 15
}
```

```powershell
python -m report_automation_engine.hwp_com_writer `
  --package report_package.json --preflight preflight_report.json `
  --template template.hwpx --output report_draft.hwpx `
  --page-setup hwp_page_setup.json
```

`--page-setup`이 없으면 기존 동작 그대로 템플릿을 따릅니다.
writer 리포트의 `page_setup`에는 요청값, 적용 여부, 실제 HWPUNIT 값과 본문 폭/높이를 기록합니다.
dry-run은 입력을 검증하고 요청값만 기록하며 실제 적용 완료로 표시하지 않습니다.

## 검증

```powershell
.venv\Scripts\python.exe -m unittest discover -s tests -p test_hwp_page_setup.py -v
ReportAutomationLauncher.exe --self-check-hwp-page-setup
```

기준 API: 사용자가 제공한 `ParameterSetTable_2504.pdf`의 PageDef/Section 설정과 PageSetup 예제.
`ApplyTo=2`는 현재 구역, mm 변환은 `7200 / 25.4` HWPUNIT을 사용합니다.
