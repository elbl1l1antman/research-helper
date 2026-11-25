# 한글 표 분할 미리보기와 보정

`0.0.57`부터 한글 보고서를 만들기 전에 표가 어떻게 나뉠지 확인하고 고칠 수 있습니다.
여기서 고칠 수 있는 항목은 다음과 같습니다.

- 표를 나누는 위치
- 열 폭
- 병합 셀 처리 방식
- 가로 용지 배치

미리보기는 한글 writer와 같은 계산 함수(`hwp_table_planning.plan_hwp_table`)를 씁니다.
그래서 화면에 보이는 분할 결과가 실제 생성에 그대로 쓰입니다.

## 런처에서 사용하기

1. 한 번 실행해 `..._report_package.json`을 만듭니다. Excel 산출 시트만 만들어도 됩니다.
2. **원본 → 표·배너** 화면에서 **한글 표 분할 미리보기…**를 누릅니다.
   직전 실행의 package가 없으면 파일을 고르는 창이 열립니다.
3. 위쪽 **모든 표 기본값**에서 정책을 고릅니다. 표마다 다르게 하려면 표를 고른 뒤 오른쪽 정책 목록에서 바꿉니다.
   '기본값 따름'을 고르면 위쪽 기본값을 따릅니다.
4. 열 목록에서 다음을 지정할 수 있습니다.
   - **여기서 새 표**: 이 열에서 새 표를 시작합니다. 회색 칸은 나눌 수 없는 위치입니다.
     반복 열이거나, 본문 병합 셀이 가로지르는 위치입니다. 마우스를 올리면 이유가 보입니다.
   - **고정 폭(mm)**: 이 열 폭을 고정합니다. 비워 두면 자동입니다. 남은 폭은 고정하지 않은 열이 나눠 씁니다.
5. **다시 계산**으로 결과를 확인하고 **저장**합니다.
   - 저장 위치: package와 같은 폴더의 `<원본 Excel 이름>_table_layout.json`
   - 같은 원본 Excel로 다시 실행해도 그대로 적용됩니다.
   - 다음 한글 생성부터 writer에 `--table-layout`으로 자동 전달됩니다.

미리보기의 본문 폭은 추정값입니다. 런처에서 편집용지를 지정했으면 그 값을 쓰고, 아니면 HWPX 템플릿의 `{{BODY}}` 구역에서 읽습니다.
실제 생성 때는 한글에서 본문 폭을 다시 측정합니다. 템플릿 여백이 바뀌었다면 생성 기록(`_hwp_writer_report.json`)의 `table_layouts`를 확인하세요.

## 정책

| 항목 | 값 | 동작 |
| --- | --- | --- |
| 넓은 표 (`wide_table_policy`) | `split` (기본) | 본문 폭에 맞게 값 열을 나눕니다. 보기·BASE 열은 각 표에 반복합니다. 나눌 수 없으면 생성이 멈춥니다. |
| | `landscape_if_needed` | 세로 용지에서 배치할 수 없을 때만 표를 가로 용지 구역에 넣습니다. |
| | `landscape_if_split` | 세로 용지에서 나뉘거나 배치할 수 없으면, 가로 용지에서 표 수가 줄어들 때 가로 용지를 씁니다. |
| 본문 병합 (`body_merge_policy`) | `block` (기본) | 본문 병합 셀이 걸친 경계에서는 나누지 않습니다. |
| | `split_repeat` | 병합 셀을 경계에서 나누고, 나뉜 각 표에 같은 값을 반복 표시합니다. 생성 기록에 경고를 남깁니다. |
| 헤더 병합 (`header_merge_policy`) | `base_only` (기본) | 헤더 병합이 본문까지 이어지면 다음 쪽에서 BASE 행만 반복합니다. |
| | `split_at_header` | 본문까지 이어진 머리글 병합을 헤더 경계에서 잘라 헤더 행 전체를 반복합니다. 값·BASE 셀이 걸쳐 있으면 자르지 않습니다. |
| 세로 배너 (`vertical_banner_policy`) | `merge` (기본) | 세로 배너를 병합해 둡니다. 쪽이 넘어가면 배너 글자는 첫 쪽에만 보입니다. |
| | `repeat_per_row` | 병합을 풀고 배너 글자를 행마다 씁니다. 어느 쪽에서 보아도 배너를 알 수 있습니다. |

### 가로 용지 구역

가로 용지가 필요한 표는 다음 순서로 배치합니다.

1. 표 앞에서 **구역 나누기**를 하고, 새 구역의 용지 방향만 바꿉니다. 문항 제목과 분석문은 원래 쪽에 남습니다.
2. 다음 문항을 시작할 때 다시 구역을 나누고 원래 방향으로 되돌립니다.
3. 마지막 문항이 가로 용지일 때:
   - `{{BODY}}` 뒤에 템플릿 글(예: 검토 요약)이 있으면 그 앞에서 원래 방향으로 되돌립니다.
   - 뒤에 아무것도 없으면 빈 쪽을 만들지 않고 가로 구역으로 끝냅니다.

자동 재개 기록에도 '가로 구역이 열린 상태'가 저장됩니다. 그래서 중간에 멈췄다가 다시 시작해도 방향이 맞게 복원됩니다.

## table_layout.json 형식

```json
{
  "schema_version": "1.0",
  "defaults": {
    "wide_table_policy": "landscape_if_needed",
    "body_merge_policy": "block",
    "header_merge_policy": "base_only",
    "vertical_banner_policy": "merge"
  },
  "tables": {
    "Q12": {"cuts": [7, 12], "column_widths_mm": {"1": 28}, "body_merge_policy": "split_repeat"}
  }
}
```

- `cuts`: 새 표를 시작할 원본 열 번호입니다(1부터, 반복 열 다음부터).
  - `[]`이면 나누지 않습니다. 폭을 넘으면 생성 전에 멈춥니다.
  - 항목이 없으면 자동으로 나눕니다.
- `column_widths_mm`: 열 번호별 고정 폭(2~400mm)입니다. 고정 폭의 합이 본문 폭을 넘으면 차단합니다.
- `col_count`: 보정을 저장할 때의 열 수입니다. 런처가 자동으로 기록합니다.
  - 배너 선택이 바뀌어 열 수가 달라지면 `cuts`와 `column_widths_mm`는 적용하지 않고 경고를 남깁니다. 정책 항목은 그대로 적용합니다.
- 알 수 없는 항목이나 범위를 벗어난 값이 있으면 writer가 생성 전에 이유와 함께 멈춥니다(`validate` / `table_layout`).

## 명령행

```powershell
python -m report_automation_engine.hwp_table_split_preview `
  --package "C:\path\설문_report_alpha_..._report_package.json" `
  --template "C:\path\report_template_basic.hwpx" `
  --table-layout "C:\path\설문_table_layout.json" `
  --format text
```

- `--page-setup`: 템플릿 대신 편집용지 설정 JSON을 기준으로 계산합니다.
- `--table-key`: 일부 표만 봅니다. 여러 번 지정할 수 있습니다.
- `--output`: JSON을 파일로 저장합니다.

한글 writer에는 `--table-layout <path>`를 추가합니다.
