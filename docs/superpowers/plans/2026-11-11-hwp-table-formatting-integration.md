# HWPX 표·서식 안정화 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 집계표의 표시값·병합·서식을 보존하는 편집 가능한 HWPX 표를 만들고, 잘못된 셀 입력을 차단한다.

**Architecture:** 기존 `report_package → hwp_table_layout → hwp_com_writer`와 이름 있는 서식 프리셋을 유지한다. COM 액션의 성공과 대상 표를 확인한 후에만 다음 입력으로 진행하며, 저장 결과의 구조는 기존 스모크 검사에서 별도로 읽는다. 검증이 끝난 뒤 선 표시 옵션만 기존 설정에 추가한다.

**Tech Stack:** Windows, 한컴오피스 한글 COM/pywin32, Python 3.12+, 표준 ZIP/XML 라이브러리, C# WinForms, pytest.

**Spec:** [한글 표·서식 개선 설계](../specs/2026-11-11-hwp-table-formatting-integration.md), 첨부 `범피스_한글기능_역설계.md`, [역설계 세션](codex://threads/6af3be0d-d230-8191-9353-3b8a3497f08d?hostId=durable).

## Global Constraints

- Windows + 설치된 한글 + 기존 pywin32, HWPX 출력, 기존 HWP 입력의 HWPX 변환 경로를 유지한다. Python 배포 빌드 조건은 기존 Windows x64 3.12 이상이다. 새 의존성과 새 문서 백엔드를 추가하지 않는다.
- 원본 Excel·템플릿 불변, 사전검증 차단 준수, 작업 사본 사용, 표의 `display_text` 보존, 그림/OLE 대신 편집 가능한 표 객체, 본문 한 단, 기본 좌우 외곽선 숨김, `[BASE : 전체 | 단위 : 명]` 형식을 유지한다. 현재 보안·권한 설정을 바꾸지 않는다.
- 기존 편집용지 기본값, 본문 구역 적용, 서식 스키마 `1.0`과 사용자 프리셋 3개 제한을 유지한다. 원본 실행 파일의 코드나 리소스를 포함하지 않는다.
- 작업별 검증과 커밋을 분리한다. 버전 변경·푸시는 별도 요청에 따른다. EXE 갱신 시 고정 파일명 `ReportAutomationLauncher.exe`만 사용한다.

## Review Focus

- 텍스트 입력·셀 이동이 `False`를 반환해도 다음 셀이나 저장으로 넘어가면 안 된다. Task 1.
- 커서가 다른 `tbl`에 진입하거나 병합 선택이 실패하면 대상 표 안에서라도 잘못 쓸 수 있다. Task 1.
- 손상된 저장 결과에서 문서 어딘가에 문자열이 있다는 이유로 셀 일치 판정을 하면 안 된다. Task 2.
- 범위가 넓거나 세로로 긴 표의 분할·머리행·표지 보존은 XML 검사와 화면 검사를 구분해야 한다. Task 2.
- 예전 프리셋을 열어도 선이 달라지면 안 되며, 무테두리 BASE 행과 숨긴 좌우선이 새 옵션보다 우선한다. Task 3.

---

## File Structure

| 파일 | 담당 |
| --- | --- |
| `report_automation_engine/hwp_com_writer.py` | COM 적용 순서, 표별 실패 기록, 경계별 선 옵션 적용 |
| `tests/test_hwp_com_writer_styles.py` | 기존 Fake COM을 확장해 실패와 서식 순서 검증 |
| `report_automation_engine/hwp_com_smoke.py`, `tests/test_hwp_com_smoke_tables.py` | 저장된 합성 문서의 셀 계약 검사, XML 검사 자체의 회귀 테스트 |
| `report_automation_engine/hwp_style_config.py`, `report_automation_engine/config/default_hwp_style_config.json`, `tests/test_hwp_style_config.py` | 프리셋 옵션 검증·기본값·이전 버전 호환 |
| `report_automation_launcher/src/ReportAutomationLauncher.cs` | 프리셋 검증과 편집 UI, 기존 CLI 자체 검사 확장 |
| `docs/hwp_template_table_insertion_plan.md` | 저장 구조/화면 검증의 적용 범위 및 선 옵션 안내 |

작업 시작 시 `bd99f90` 이후 변경을 확인한다. 계획 작성 시점의 기준선은 HWP 관련 테스트 **57 passed, 32 subtests passed**다. 아래 명령은 저장소 루트 PowerShell에서 실행한다.

### Task 1: 잘못된 커서 위치·부분 생성 차단

**Files:** Modify `report_automation_engine/hwp_com_writer.py` (`write_table_part`, `insert_text`, `insert_hwp_table`, `apply_table_merges`, `apply_table_matrix_styles`, `apply_cell_appearance`); Test `tests/test_hwp_com_writer_styles.py`.

**Interfaces:** 기존 `insert_text(hwp, text: str, report: dict) -> None`, `insert_hwp_table(...) -> bool` 유지. 후자의 `False`는 빈 입력처럼 아직 TableCreate를 호출하지 않은 경우만 허용한다. 신규 내부 함수 `require_table_action(hwp, action: str, report: dict, stage: str) -> None`, `require_current_table(hwp, control, action: str, stage: str) -> None`는 실패 시 `HwpWriterError`를 발생시킨다.

- [ ] **Step 1: 기존 Fake COM에 실패 시나리오를 추가한다.** `[["A", "B"]]`를 입력하면서 `TableRightCell=False` 또는 다른 개체로 이동을 주입한다. 다른 표는 다른 `GetCtrlInstID()`를 반환한다. 빈 문자열 입력 검사는 이벤트가 없는 새 `FakeHwp`를 사용한다. 저장 방지 검사는 HWPX 입력과 2번 분할 표를 사용하고 실제 `write_hwp_document`가 호출하는 `save_as_hwpx`를 패치한다.

```python
# test_contract_table_stops_when_move_fails / test_contract_table_rejects_another_table
assert exc.value.action == "TableRightCell"
assert [e[1] for e in hwp.events if e[0] == "text"] == ["A"]
assert report["tables_written"] == 0
# test_insert_text_false_is_fatal_and_empty_is_noop
assert exc.value.action == "InsertText"
assert empty_hwp.events == []
# test_merge_selection_failure_does_not_merge
assert exc.value.action == "TableCellBlockExtend"
assert ("run", "TableMergeCell") not in hwp.events
# test_table_create_failure_reports_part_without_saving (실제 writer 경로, fake COM)
assert report["table_failure"] == {"table_key": "T001", "part_index": 2, "action": "TableCreate"}
assert report["status"] == "failed"
assert output.read_bytes() == b"existing-output"
assert not save_mock.called
```

- [ ] **Step 2: 실패를 확인한다.** `.\.venv\Scripts\python.exe -m pytest tests/test_hwp_com_writer_styles.py -k 'contract_table or insert_text_false or merge_selection_failure or table_create_failure' -q`에서 대상 테스트가 이전 코드의 무시된 반환값 또는 누락된 실패 기록으로 실패해야 한다.
- [ ] **Step 3: 검증을 구현한다.** 빈 텍스트는 호출하지 않고 비어 있지 않은 `InsertText`의 `False`는 중단한다. 확정된 이동·선택·병합에 `require_table_action`을 사용한다. 이 함수는 `HAction.Run`을 한 번만 호출한다. 실패한 작업이 이미 일부 반영됐을 수 있으므로 `hwp.Run`으로 재시도하지 않는다. 현재 개체 검사는 `is_hwp_table_control`과 `same_hwp_control`을 재사용한다. 이동 후, 셀 텍스트 삭제·입력 전, 병합 적용 전에 원래 개체인지 확인한다. 마지막 셀은 이동하지 않는다. 표 시작점을 찾는 기존 제한된 `TableUpperCell` 탐색의 경계 `False`는 허용하되, 다른 개체에 진입한 경우와 병합 좌표를 향한 필수 이동 실패는 허용하지 않는다. `apply_table_merges`에 기존 대상 `control`을 전달하고, 병합 성공 후 `Cancel`로 선택을 해제한다. `write_table_part`에서 `HwpWriterError`를 받을 때 `table_failure`를 기록하고 다시 발생시킨다. 실제 `TableCreate` 호출 이후 오류를 `False`나 텍스트 대체 경고로 바꾸지 않는다. 정상 클립보드 Undo/대체 조건은 그대로 둔다.
- [ ] **Step 4: 기존 계약을 포함해 검증한다.** `.\.venv\Scripts\python.exe -m pytest tests/test_hwp_com_writer_styles.py tests/test_hwp_com_writer_tables.py -q`가 통과해야 한다. 기존 `test_clipboard_noop_does_not_undo_caption_or_edit_existing_table`, `test_clipboard_vertical_merges_skip_revisited_cell_lists`도 유지한다.
- [ ] **Step 5: 커밋한다.** `git add report_automation_engine/hwp_com_writer.py tests/test_hwp_com_writer_styles.py` 후 `git commit -m "fix(hwp): stop on failed table edits"`.

### Task 2: 저장된 표의 좌표·값·병합을 스모크 검사로 검증

**Files:** Modify `report_automation_engine/hwp_com_smoke.py`; Create `tests/test_hwp_com_smoke_tables.py`; Modify `docs/hwp_template_table_insertion_plan.md`의 검증 항목.

**Interfaces:** `inspect_smoke_tables(path: Path) -> list[dict[str, Any]]`를 스모크 모듈에 추가한다. 문서 순서로 각 표의 `rows`, `cols`, `width_hwpunit`, `repeat_header`, `cells`를 반환한다. `cells`는 `(row, col)` 1기준 키와 `(rowspan, colspan, text, header)` 값의 사전이다. `check_smoke_tables(path: Path, body_widths: list[int], report: dict) -> None`는 고정 표 계약과 비교해 기존 `checks`/`errors`에 결과를 추가한다. 사용자 출력의 일반 writer API는 변경하지 않는다.

- [ ] **Step 1: 작은 HWPX ZIP을 만든 뒤 검사를 고정한다.** `tests/test_hwp_com_smoke_tables.py`에서 아래 계약의 표 두 개를 `Contents/section0.xml`에 넣고, 각 변형은 정상 입력에서 한 속성만 바꾼다. 테스트 표 폭은 30000, 허용 본문 폭은 `[40000, 40000]`으로 둔다. 같은 텍스트를 표 밖에도 두어 문자열 검색만으로 통과하지 못하게 한다. 정상 입력이 오류 없이 통과하는 것도 확인한다.

```python
# test_smoke_table_reads_logical_grid_and_display_text
assert (table["rows"], table["cols"]) == (4, 3)
assert table["cells"][1, 1] == (1, 3, "[BASE : 전체 | 단위 : 명]", True)
assert table["cells"][2, 2] == (1, 2, "결과", True)
assert (2, 3) not in table["cells"]
assert table["cells"][3, 2] == (1, 1, "(3,232)", False)
assert table["cells"][3, 3] == (1, 1, "3.3", False)
assert table["cells"][4, 2][2] == "0"
assert table["cells"][4, 3][2] == ""
# test_smoke_table_mismatch_fails_even_if_document_contains_text
# 입력 변형: 값 위치 교환 / rowspan 오기 / 머리행 누락 / 표 1개 누락 / 폭 초과
assert any(check["name"] == "table_contract" and not check["ok"] for check in report["checks"])
```

- [ ] **Step 2: 실패를 확인한다.** `.\.venv\Scripts\python.exe -m pytest tests/test_hwp_com_smoke_tables.py -q`에서 검사 함수 부재로 실패해야 한다.
- [ ] **Step 3: 좌표 기반 검사를 구현한다.** `hwp_template_probe.local_name`을 재사용하고 `Contents/sectionN.xml`의 N을 숫자로 정렬한다. `tbl`의 `rowCnt`/`colCnt`, `sz.width`, `repeatHeader`와 직접 자식 `tr/tc`의 `cellAddr`/`cellSpan`/`header`를 읽는다. `rowAddr`/`colAddr`는 0기준이므로 반환할 때 1을 더한다. `header`와 `repeatHeader`는 저장된 `0`/`1`로 비교한다. 이 구조는 기존 `outputs/hwp_com_smoke_v009_final/draft_output.hwpx`에서 확인했다. 텍스트 런 분할은 합치되 단위·괄호·숫자·공백을 재서식하지 않는다. 논리 좌표, 크기, 중복 앵커, 병합 상태 또는 필수 구조가 이상하면 `add_check(report, "table_contract", False, detail)`로 기록한다. 기존 탐색기의 `sample_rows`를 늘리거나 임의 문서의 결과 표를 추측하는 기능은 만들지 않는다.
- [ ] **Step 4: 기존 스모크에 고정 입력을 연결한다.** `sample_package()`에 `cell_contract=True`, `base_label="전체"`, `unit="명"`, 3x3 원본 matrix를 넣는다. 첫 행은 `[구분, 결과, ""]`, 두 번째는 `[그렇다, (3,232), 3.3]`, 세 번째는 `[미응답, 0, ""]`; 헤더 B1:C1을 병합하고 B2의 raw 값은 `-3232`, C2는 `3.333535353`이다. 두 번째 표는 다른 `table_key`와 항목명 `다음 표`를 사용한다. `source_workbook` 없이 기본 스모크는 계약 대체를 실행한다. 저장 결과의 두 표 개수와 순서, 위 계약, `repeat_header`, `0 < width_hwpunit <= body_width_hwpunit`, 후행 QA 문단, 템플릿 SHA-256 불변을 확인한다. 본문 폭은 writer 리포트 `table_layouts`의 해당 값을 읽는다. 기존 샘플 문자열·`{{BODY}}` 제거 검사도 유지한다.
- [ ] **Step 5: 구조 테스트를 실행한다.** `.\.venv\Scripts\python.exe -m pytest tests/test_hwp_com_smoke_tables.py tests/test_hwp_table_layout.py tests/test_hwp_table_metadata.py -q`가 통과해야 한다.
- [ ] **Step 6: 설치된 한글에서 확인한다.** `.\.venv\Scripts\python.exe -m report_automation_engine.hwp_com_smoke --output-dir outputs/hwp_table_integration_smoke`를 실행한다. `hwp_com_smoke_report.json`이 `ready`이고 두 실제 표가 검증돼야 한다. 사용자 보고서가 열린 인스턴스는 사용하지 않는다.
- [ ] **Step 7: 실제 보고서 대조 기록을 남긴다.** 검토한 HWPX 템플릿 사본과 변형 집계표로 기존 런처를 실행한다. 표 2개 이상, 넓은 표의 가로 분할, 두 페이지를 넘기는 긴 표와 머리행 반복, 한글+영문 혼용, 미설치 글꼴 대체를 확인한다. 별도 구역의 표지·머리말·꼬리말이 본문 편집용지 변경으로 바뀌지 않는지 확인한다. 같은 표를 Excel 클립보드와 계약 대체에서 한 번씩 산출해 값·병합을 대조하고, 실제 적용된 `insert_mode`를 기록한다. 확인 불가한 PC나 클립보드 환경은 검증 보류로 남긴다. 재사용할 수 있는 테스트에는 합성 데이터만 포함한다.
- [ ] **Step 8: 커밋한다.** `git add report_automation_engine/hwp_com_smoke.py tests/test_hwp_com_smoke_tables.py docs/hwp_template_table_insertion_plan.md` 후 `git commit -m "test(hwp): verify saved table geometry and display values"`.

### Task 3: 내부선·외곽선 표시 설정을 기존 프리셋에 추가

**Files:** Modify `report_automation_engine/hwp_style_config.py`, `report_automation_engine/config/default_hwp_style_config.json`, `report_automation_engine/hwp_com_writer.py`, `report_automation_launcher/src/ReportAutomationLauncher.cs` (`HwpStylePresetStore`, 서식 편집 대화상자), `docs/hwp_template_table_insertion_plan.md`; Test `tests/test_hwp_style_config.py`, `tests/test_hwp_com_writer_styles.py` 및 기존 EXE 자체 검사.

**Interfaces:** `validate_hwp_style_config(config)`의 입력/출력에 `table_border.inner_type`과 `table_border.outer_type`을 추가한다. 두 필드는 `str` 열거 `solid`/`none`, 기본 `solid`. `apply_cell_appearance`와 런처의 `--hwp-style-config` 전달 방식은 그대로다.

- [ ] **Step 1: 프리셋 및 선 회귀 테스트를 추가한다.** 옛 설정의 두 키를 삭제해 로드한다. 타입·값 오류와 경계별 우선순위를 검증한다. 예시의 튜플 순서는 좌·우·상·하다.

```python
# test_legacy_border_types_default_to_solid
assert normalized["table_border"]["inner_type"] == "solid"
assert normalized["table_border"]["outer_type"] == "solid"
# test_border_types_reject_invalid_values: True, None, "dashed" 각각 ValueError
# test_hidden_inner_edges_keep_outer_and_metadata_rules: inner_type=none, outer_type=solid
assert ("border_type", 0, 0, 1, 0) in first_header_hwp.events
assert ("border_type", 0, 0, 0, 0) in metadata_hwp.events
# test_none_outer_edges_keep_inner_edges: 가운데 열 첫 헤더, outer_type=none
assert ("border_type", 1, 1, 0, 1) in header_hwp.events
```

- [ ] **Step 2: 실패를 확인한다.** `.\.venv\Scripts\python.exe -m pytest tests/test_hwp_style_config.py tests/test_hwp_com_writer_styles.py -k 'border_type or inner_edges or outer_edges' -q`에서 새 기본값 또는 새 선 적용 단언이 실패해야 한다.
- [ ] **Step 3: Python 설정과 적용을 구현한다.** 검증 단계에서 누락 기본값을 채우고 허용 문자열 외에는 필드명을 포함한 `ValueError`를 낸다. 기존 변환 `HwpLineType("Solid")`와 기존 없음 코드 `0`을 사용한다. 좌/우가 표 끝이거나 상단이 `data_start_row`, 하단이 마지막 행이면 외곽선이며 나머지는 내부선이다. 병합은 span을 포함한 끝 좌표로 판정한다. `metadata`와 `show_side_borders=false`가 우선한다. `params`와 `SelCellsBorderFill`에 동일한 선 값을 넣고 채움 후 테두리 적용 순서를 유지한다.
- [ ] **Step 4: 런처 프리셋에 연결한다.** `HwpStylePresetStore`의 기본값·검증·저장 불러오기에 동일 두 키를 반영한다. 기존 서식 대화상자의 테두리 영역에 `내부선`, `외곽선` 두 선택을 추가하고 표시명 `실선`, `없음`을 저장값 `solid`, `none`으로 매핑한다. 3개 슬롯, 굵기·색상·좌우선 저장값을 유지한다. `--self-check-hwp-style-presets`에 옛 설정 기본값, `none` 저장/로드, 잘못된 값 거부를 추가한다.
- [ ] **Step 5: 회귀 검사를 실행한다.** `.\.venv\Scripts\python.exe -m pytest tests -q`에서 새 실패가 없어야 한다. `powershell -NoProfile -File report_automation_launcher/scripts/build_report_automation_launcher.ps1 -DevelopmentOnly -OutputDir outputs/hwp_table_plan_check`로 검증본을 만들고 해당 EXE의 `--self-check-hwp-style-presets`가 종료 코드 0인지 확인한다.
- [ ] **Step 6: 실물 결과를 확인한다.** Task 2와 같은 문서에 기본 설정·내부선 없음·외곽선 없음 각각 적용한다. 병합 헤더의 공유 경계, BASE 행, 이어지는 페이지의 머리행과 배경색이 저장 후에도 기대대로인지 한글에서 다시 열어 비교한다. 템플릿 후보 스타일은 읽기 전용 참고값으로 남는다는 점과 선 우선순위를 안내 문서에 추가한다.
- [ ] **Step 7: 단일 배포 EXE를 갱신한다.** 실물 검증이 통과한 경우에만 `powershell -NoProfile -File report_automation_launcher/scripts/build_report_automation_launcher.ps1` 후 `powershell -NoProfile -File tests/test_standalone_launcher.ps1`을 실행한다. 로컬 Excel 캐시 접근 검증은 별도이며 이 작업이 해결했다고 기록하지 않는다.
- [ ] **Step 8: 커밋한다.** Task 3의 파일과 갱신된 고정 배포 EXE만 스테이징하고 `git commit -m "feat(hwp): configure inner and outer table line visibility"`.

## 이후 후보

점선·이중선은 한컴 API 값과 저장된 공유 경계를 실측한 후 확장한다. 셀 여백은 Excel 임시 표와 TableCreate 폭 계산을 함께 바꾸어야 하므로 별도 계획으로 다룬다. 기관별 양식이 여러 개 확보되면 템플릿에서 인식한 값을 곧바로 적용하기보다 사용자 프리셋으로 미리보기·선택하도록 검토한다.
