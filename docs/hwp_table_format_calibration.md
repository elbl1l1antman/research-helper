# 한글 3×3 표 서식 실측

2026-11-11, 한컴오피스 한글 COM `13.0.0.711`. 범피스 실행 없이 프로젝트의 기존 한글
작성기와 설치된 한글만 사용했다. 사용자 문서를 열지 않았으며, 합성 기준표 한 개와
비교본 11개를 저장·재열기했다. 각 비교본의 3×3 논리 좌표와 `(3,232)`, `3.3`, `0`,
빈칸 등 모든 표시값을 검사했다. 실험 전후 기준본 SHA-256이 일치했다.

## 동작과 결과

| 대상 | 실행 전 선택 및 액션 | 저장된 HWPX 결과 |
| --- | --- | --- |
| 위쪽 외곽선 | 표 전체, `CellBorderFill`의 `BorderTypeTop`, `BorderWidthTop`, `BorderColorTop` | 첫 행 세 셀의 위쪽만 `SOLID`, 0.4mm, `#1F4E79` |
| 내부 가로선 | 표 전체, `TypeHorz`, `WidthHorz`, `ColorHorz` | 내부 경계 양쪽의 위/아래 선이 0.4mm로 바뀜. 외곽·세로선은 유지 |
| 내부 세로선 | 표 전체, `TypeVert=0`, `WidthVert`, `ColorVert` | 내부 경계 양쪽의 좌/우 선이 `NONE`. 외곽·가로선은 유지 |
| 머리행 아래선 | 첫 행 세 셀, `CellBorderFill`의 `Border*Bottom` | 첫 행 아래쪽과 다음 행 위쪽이 함께 0.4mm로 저장됨 |
| 배경 | 가운데 셀, `CellFill`의 `FillAttr` | `(2,2)`만 `#DDEBF7`. 다른 셀의 선·배경은 유지 |
| 배경 후 선 | 가운데 배경 적용 후 첫 행 아래선 변경 | 가운데 배경을 유지한 채 공유 경계가 변경됨. 같은 셀을 혼합 선택하는 경우까지 보장하지 않음 |
| 셀 안 여백 | 가운데 셀, `TablePropertyDialog`의 `ShapeTableCell.HasMargin`와 `MarginLeft/Right/Top/Bottom` | `(2,2)`만 `hasMargin=1`. 좌우 1mm는 각각 283, 상하 0.5mm는 각각 141 HWPUNIT. 비대칭 1/2/0.8/1.2mm는 283/566/226/340으로 저장 |

한글의 `MiliToHwpUnit`을 사용했다. 시험 표에서는 가운데 셀의 여백이 바뀌어도 셀
외곽 크기가 그대로였다. 폭 계획에서는 **전체 폭과 글자가 들어갈 폭을 분리**해야 한다.
병합·긴 줄바꿈·페이지 분할에 미치는 영향은 다음 단계에서 확인한다.

셀 여백 속성은 [한컴의 웹한글 셀 여백 답변](https://forum.developer.hancom.com/t/topic/2528)을
참고하되, 웹 API 동작이 데스크톱과 같다고 가정하지 않고 위 COM 실험으로 확인했다.
역설계 세션에서 관찰한 내부/외곽선 액션은 [원본 분석](codex://threads/6af3be0d-d230-8191-9353-3b8a3497f08d?hostId=durable)에 남긴다.

## 선 종류 매핑

| `HwpLineType`에 전달 | 반환값 | 이 버전의 HWPX 저장값 |
| --- | ---: | --- |
| `None` | 0 | `NONE` |
| `Solid` | 1 | `SOLID` |
| `Dash` | 2 | `DOT` |
| `Dot` | 3 | `DASH` |
| `DoubleSlim` | 8 | `DOUBLE_SLIM` |
| `Double`, `DoubleLine` | 호출 거부 | 미검증 |

**입력 영문 이름을 저장값이나 UI 표시명으로 그대로 사용하지 않는다.** 새 프리셋은
저장 결과에 근거해 의미를 정하고, 지원할 한글 버전에서 다시 확인한다. 공유 경계는
한 셀의 선만 보고 적용 성공이라고 판단하지 않는다.

## 재현 및 범위

한글에 열려 있는 문서를 모두 닫은 뒤 저장소 루트에서 실행한다. 기존 폴더를 덮어쓰지 않는다.

```powershell
if (@(Get-Process -Name Hwp,HwpFrame -ErrorAction SilentlyContinue).Count -gt 0) { throw 'Hancom already running.' }
.\.venv\Scripts\python.exe -X utf8 tests/probe_hwp_table_format.py `
  --output-dir outputs/hwp_format_calibration/new-run
```

결과: `outputs/hwp_format_calibration/2026-11-11-final/`의 `calibration.json`, `*.hwpx`,
한글 내장 `Preview/PrvImage.png` 사본. 기준본, 배경+머리행 아래선, 이중선의 미리보기를
육안 대조했다. 저장 파일의 COM 재열기는 확인했지만, 실제 보고서의 편집 화면 검수는 아니다.
범피스 결과 일치 여부, 병합·다페이지·Excel 붙여넣기, 다른 한글 버전은 미검증이다.

다음 작업은 개정 계획 Task 2의 **역할별 프리셋과 선택한 표 서식 가져오기**다.
실측 결과를 입력·검증 계약에 반영한 뒤 제품 코드에 연결한다.
