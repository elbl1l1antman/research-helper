# 0.0.57 로컬 검증 체크리스트

이 버전의 코드는 Linux 클라우드 환경에서 작성했습니다. 그곳에서 확인한 것은 다음과 같습니다.

- Python 단위 테스트
- Mono `mcs -langversion:5`로 런처 컴파일
- 런처 자체 검사 중 UI·경로 관련 2개

한글 COM, Excel, Windows용 런처 빌드는 아래 순서로 로컬 Windows에서 확인해 주세요.
실제 원자료와 결과물은 기존 규칙대로 `outputs/` 또는 `test-data/`에만 두세요.

## 0. 준비

```powershell
git fetch origin claude/brave-keller-w1a57p
git checkout claude/brave-keller-w1a57p
```

```powershell
.\.venv\Scripts\python.exe -m pytest -q tests
```

- 기대 결과: 모두 통과
- Linux에서만 실패하던 다음 5개도 Windows에서는 통과해야 합니다.
  - `test_standalone_payload` 4개
  - `test_style_registration_failure_does_not_create_output`
- 새 테스트 파일:
  - `test_hwp_com_lifecycle.py`
  - `test_hwp_table_layout_policies.py`
  - `test_hwp_table_landscape.py`
  - `test_hwp_table_split_preview.py`
  - `test_document_writer_pptx.py`

## 1. 런처 빌드와 자체 검사

```powershell
.\report_automation_launcher\scripts\build_report_automation_launcher.ps1
```

```powershell
.\report_automation_launcher\bin\ReportAutomationLauncher.exe --self-check-table-split-preview
```

```powershell
.\report_automation_launcher\bin\ReportAutomationLauncher.exe --self-check-hwp-style-cli
```

```powershell
.\report_automation_launcher\bin\ReportAutomationLauncher.exe --self-check-hwp-page-setup
```

```powershell
powershell -ExecutionPolicy Bypass -File tests\test_standalone_launcher.ps1
```

- `bin\ReportAutomationLauncher.exe`는 이 브랜치에서 다시 빌드하지 않았습니다. 로컬 빌드 결과를 커밋해 주세요.
- `--self-check-hwp-page-setup`은 Mono에서 원래 코드도 실패합니다(컨트롤 겹침 판정). 그래서 Windows 결과로만 판단해야 합니다.

## 2. 연속 생성 안정성 (작업 3)

1. 같은 PowerShell 창에서 첫 문항만 생성하는 HWPX writer를 3번 연속 실행합니다.

   ```powershell
   1..3 | % { .\.venv\Scripts\python.exe -m report_automation_engine.hwp_com_writer --package P --preflight F --template T --output "outputs\seq_$_.hwpx" --max-sections 1 }
   ```

   - 기대 결과: 3번 모두 성공합니다.
   - 각 `_hwp_writer_report.json`의 `com.process_ids`에 이번 실행이 띄운 Hwp.exe PID가 있어야 합니다.
   - 각 기록의 `com.exit_wait.still_running`은 `[]`이어야 합니다.
2. 한 Python 프로세스 안에서 `write_hwp_document`를 연속으로 두 번 호출합니다. 예전에 실패하던 경우입니다.
   - 기대 결과: 두 번째 기록에 `previous_exit_wait`가 있거나, 대기 없이 성공합니다.
   - 일시적인 RPC 거부가 있었다면 `com.steps`에 `"status": "retry"`가 남습니다.
3. 한글을 미리 열어 둔 상태에서 실행합니다.
   - 기대 결과: `attached_existing_instance`가 `true`여도 사용자가 열어 둔 한글 창은 강제 종료되지 않습니다.
   - 단, `Quit`은 기존 동작과 같습니다.
4. 종료가 늦은 경우: Hwp.exe가 15초 안에 끝나지 않으면 경고가 남습니다.
   - 이때 `%LOCALAPPDATA%\ResearchHelper\hwp_pending_exit.json`이 생깁니다.
   - 다음 실행은 이 파일의 PID를 먼저 기다립니다. 2분이 지난 기록은 무시합니다.

## 3. 표 분할 미리보기와 정책 (작업 1·2)

합성 자료 또는 `test-data/`의 실제 집계표로 확인합니다.

1. Excel 산출을 실행한 뒤 **원본 → 표·배너 → 한글 표 분할 미리보기…**를 엽니다.
   - 본문 폭이 템플릿 기준(기본 템플릿 160.0mm)으로 보이는지 확인합니다.
   - 넓은 표의 분할 위치가 실제 생성 결과(`_hwp_writer_report.json`의 `table_layouts`)와 같은지 확인합니다.
2. 경계를 바꾸고 고정 폭을 넣어 **다시 계산** 후 **저장**합니다.
   - 저장 위치: `<원본 이름>_table_layout.json`
   - 같은 원본으로 다시 실행했을 때 writer 명령에 `--table-layout`이 붙는지 런처 기록에서 확인합니다.
3. 각 정책을 한글로 생성해 PDF와 한글 재열기로 확인합니다.

   | 확인 항목 | 기대 결과 |
   | --- | --- |
   | `landscape_if_needed` + 본문 병합 때문에 차단되던 표 | 표만 가로 구역에 배치. 앞 쪽(제목·분석문)과 다음 문항은 세로. 빈 쪽 없음. 머리말/꼬리말·쪽 번호 유지 |
   | `landscape_if_split` + 2개로 나뉘던 표 | 가로 구역에서 1개 표 |
   | `split_repeat` | 나뉜 각 표에 병합 값이 반복 표시. 생성 기록에 '반복 표시' 경고 |
   | `split_at_header` (구분 셀이 본문까지 병합된 표) | 쪽이 넘어가도 헤더 행 전체 반복 |
   | `repeat_per_row` (세로 배너가 쪽 경계에 걸친 긴 표) | 다음 쪽에도 배너 글자 표시 |
   | 차트(EMF) 뒤 가로 구역 표 | 차트 쪽 다음에 가로 쪽이 시작되고 한글이 멈추지 않음 |

4. 가로 구역이 열린 상태에서 생성을 중단한 뒤 다시 시작해(자동 재개) 다음 문항이 세로로 돌아오는지 확인합니다.
5. 확인할 한글 API 가정: `BreakSection` 다음에 `PageSetup`(`ApplyTo=2`)을 실행하면 새 구역에만 적용되어야 합니다. 기존 편집용지 기능과 같은 값입니다.
   - 실제로 앞 구역까지 바뀌면 `hwp_com_writer.begin_section`의 `ApplyTo` 값을 조정해야 합니다.

## 4. PPTX 초본 (작업 4)

1. 결과 화면 → **PPT 보고서 초안 만들기…**
   - 선택한 PPT 템플릿이 있으면 그 템플릿을 씁니다. 없으면 기본 양식을 씁니다.
2. PowerPoint에서 확인합니다.
   - 제목 슬라이드의 `{{...}}`가 모두 채워짐
   - 문항마다 슬라이드 1장
   - 표 병합 유지
   - 차트를 '데이터 편집'으로 열 수 있음
   - 2개 항목은 원그래프, 배너별은 묶은 막대
   - 최대값 강조
   - QA 슬라이드는 마지막
3. 사용자 템플릿 확인: `RA_CHART` 같은 이름의 큰 도형을 둔 템플릿에서는 차트가 그 도형 위치·크기를 그대로 써야 합니다.
4. QA가 `ready_with_warnings`이면 확인 창이 뜨고, `blocked`이면 생성하지 않아야 합니다.
5. 저장 위치의 PPTX를 PowerPoint로 열어 둔 채 다시 생성합니다.
   - 기대 결과: 실패 알림이 뜨고 `<파일>.pptx.error.txt`가 생깁니다.
   - 열려 있던 PPTX 원본은 그대로여야 합니다.
