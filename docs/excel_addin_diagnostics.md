# Excel 추가기능 연결 진단

## 사용 방법

런처의 파일 영역에서 **연결 검사**를 누르면 선택한 추가기능을 별도 Excel에서
읽기 전용으로 열고 검사 결과를 표시한다. 원본 집계표를 열거나 작업 복사본을 만들지 않는다.
결과 JSON은 선택한 집계표 폴더에 저장한다. 집계표를 아직 선택하지 않았다면 현재 작업 폴더에 저장한다.

CLI는 집계표 입력 없이도 사용할 수 있다.

```powershell
.\report_automation_launcher\bin\ReportAutomationLauncher.exe `
  --check-excel-addin --out "D:\reports\excel_addin_report.json"
```

`--addin "D:\path\ReportAutomationAddin_dev.xlam"`을 추가하면 직접 지정한 파일을 검사한다.
생략하면 EXE에 내장된 추가기능 경로를 검사한다. 진단은 Office 신뢰 폴더를 등록하거나
회사 정책·레지스트리를 바꾸지 않는다. 별도 검사 인스턴스에서는 매크로 실행을 강제 차단한다.
검사 종료 시 자신이 만든 Excel만 닫고 기존 사용자 Excel에는 접근하지 않는다.

## 결과 해석

- `status: opened`: Excel에서 읽기 전용으로 열기 성공. 매크로 실행 성공을 의미하지 않는다.
- `status: blocked`, `stage: file_check`: 런처에서 파일 존재를 확인하지 못함.
- `status: blocked`, `stage: excel_create`: Excel 설치/COM 연결 단계 실패.
- `status: blocked`, `stage: addin_open`: 파일은 있지만 Excel 프로세스가 열지 못함.
- `macro_execution_verified`는 이 검사에서 항상 `false`이다.

`addin_path`, `file_exists`, `excel_version`, `hresult`, `error`, `message`를 함께 기록한다.
검사 성공 종료 코드는 0, 차단 종료 코드는 1이다. 진단 기록에 원본 셀 데이터나 API 키는 넣지 않는다.

일반 산출에서도 추가기능을 먼저 열어 확인한다. 실패하면 원본 집계표를 열거나 작업 복사본을
만들기 전에 중단한다. CLI는 `--out` 경로에 `.excel_addin.json`을 덧붙인 오류 기록을,
GUI는 집계표 폴더에 `excel_addin_error_날짜시각.json`을 저장한다.
작업 중 예외가 발생하면 열린 집계표의 부분 변경을 종료 과정에서 저장하지 않는다.
성공 경로의 명시적인 저장 동작은 유지한다.

## 2026-10-14 재현 결과

현재 PC에서는 캐시의 기존 XLAM과 같은 폴더에 둔 가짜 XLSX 모두 Excel 열기에 실패했다.
해당 폴더로 Excel 자체의 사본 저장도 실패했다. 동일 바이트의 XLAM은 C 임시 폴더와
D 프로젝트 경로에서 매크로 비활성 상태로 열렸다. 임시 폴더의 173자/200자 경로도 열렸다.
따라서 파일 손상, 매크로 실행, 경로 길이만으로는 현재 현상을 설명할 수 없다.

확인된 실패 범위는 **현재 캐시 폴더에 대한 Excel 프로세스의 접근**이다. 구체적인 OS/Office
정책 원인까지 확정한 것은 아니다. 매크로 보안을 해제하거나 캐시를 임의의 신뢰 폴더로
옮기는 자동 우회는 구현하지 않는다. 사용자는 자신이 검토한 추가기능 경로를 기존 파일
선택란에서 직접 지정할 수 있다. 다른 PC에서도 같은 현상이 발생하는지는 추가 확인이 필요하다.

## 회귀 검사

```powershell
powershell -NoProfile -File tests/test_excel_addin_diagnostics.ps1
powershell -NoProfile -File tests/test_standalone_launcher.ps1
```

첫 검사는 Excel 설치 없이 누락 파일/COM 실패/정상 반환의 구분을 검증한다.
두 번째 검사는 EXE 하나만 복사해 실행하며 진단 CLI의 누락 파일 결과도 확인한다.
실제 Office 연결 검사는 위 CLI로 별도 실행해야 한다.

이번 변경 검증 결과: 전체 Python 208개 및 subtest 32개 통과(심볼릭 링크 권한 검사
1개 건너뜀), 진단 회귀 검사와 EXE 단독 실행 검사 통과. 실제 Excel에서 캐시 접근
실패는 `addin_open / 0x800A03EC`, D 경로는 `opened`로 분류되었다. 캐시 실패 시
입력 파일 해시 불변·작업 복사본 미생성을 확인했다. D 추가기능을 지정한 정상 실행은
분석문 3개·삽입표 3개·차트 후보 6개와 `ready_with_warnings` 사전검증 결과를 생성했다.

진단에서 사용하는 매크로 강제 차단 옵션의 의미는
[Microsoft AutomationSecurity 문서](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)를 따른다.
