# Report Automation Add-in Legacy Boundary

이 폴더의 VBA add-in은 더 이상 사용자가 직접 조작하는 주 프로그램이 아닙니다.

알파 단계부터 제품의 중심은 `report_automation_launcher`입니다. 이 add-in은 런처가 Excel COM으로 호출하는 내부 산출 엔진으로 유지합니다.

## 유지하는 역할

- Excel 원본 집계표에서 표 블록을 탐지합니다.
- `보고서_분석문`, `보고서_차트데이터`, `보고서_삽입표`, `보고서_QA`, `_ReportMeta` 시트를 생성합니다.
- 런처의 `ReportAutomation_RunWithOptionsSilent` 호출을 처리합니다.

## 레거시로 보는 역할

- 사용자가 Excel 리본/매크로 메뉴에서 직접 옵션을 선택해 실행하는 흐름
- VBA UserForm 중심의 GUI 확장
- HWP/PPT 산출까지 VBA 안에서 직접 처리하려는 방향

## 개발 원칙

- 이 add-in은 "Excel 산출 엔진"으로만 유지합니다.
- 새 UX, 파일 선택, 실행 설정, 결과 요약, HWPX/PPTX 출력 흐름은 런처에서 담당합니다.
- VBA 변경은 런처가 필요로 하는 안정적인 산출 시트와 메타데이터를 만드는 범위로 제한합니다.
