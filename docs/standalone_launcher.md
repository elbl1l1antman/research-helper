# 독립 실행 런처

## 사용 방법

배포 파일은 `report_automation_launcher/bin/ReportAutomationLauncher.exe` 하나입니다.
이 파일만 원하는 폴더로 옮겨 실행할 수 있습니다. 저장소, 가상환경, 별도 Python 설치,
실행 중 패키지 다운로드는 필요하지 않습니다.

필수 환경:

- Windows x64와 .NET Framework 실행 환경
- 집계표 처리용 Microsoft Excel
- HWPX 보고서 생성용 아래한글과 사용 가능한 HWP COM 자동화 환경
- 사용자 로컬 캐시와 선택한 산출 폴더에 대한 쓰기 권한

Excel 매크로 신뢰 설정이나 아래한글 보안 정책을 우회하지 않습니다.
기존 환경 검사 결과가 차단이면 설치·신뢰 설정을 확인해야 합니다.
PowerPoint는 PPTX 파일 생성 자체에 필요하지 않지만 파일 검토·편집에는 사용할 수 있습니다.
글꼴은 설치된 시스템 글꼴을 사용하며 글꼴 파일은 배포하지 않습니다.

## 포함하는 파일

- 현재 빌드 환경의 CPython x64 런타임과 필요한 라이브러리
- `report_automation_engine` Python 코드
- 기본 서식 JSON 2개
- `ReportAutomationAddin_dev.xlam`
- 기본 보고서/차트 검토 PPTX 템플릿
- 프로젝트 라이선스와 제3자 라이선스 고지

원본 집계표, 산출물, 사용자 설정, API 키, `.env`, 개발용 테스트와 가상환경 전체는
포함하지 않습니다. 외부 Python 환경변수는 무시하므로 다른 Python 설치의 영향을 받지 않습니다.
사용자가 GUI에서 직접 선택한 Excel 파일·템플릿·추가기능 경로는 기존처럼 사용할 수 있습니다.

## 로컬 캐시

첫 실행에서 EXE에 내장된 파일을
`%LOCALAPPDATA%\ResearchHelper\bundles\<내장 ZIP의 SHA256>`에 풉니다.
동일한 EXE를 다시 실행하면 캐시를 재사용하고, 필수 파일이 누락되면 재구성합니다.
추출 중 중단된 파일은 완료된 캐시로 사용하지 않습니다.
서로 다른 EXE가 동시에 실행될 가능성 때문에 이전 구성요소 캐시는 자동 삭제하지 않습니다.
캐시 정리가 필요하면 런처와 실행 중인 작업을 모두 닫은 뒤 `bundles` 폴더만 삭제할 수 있습니다.
다음 실행 시 다시 준비되며 `hwp_style_presets.json` 등 사용자 설정은 별도로 유지됩니다.

캐시는 내부 실행용 파일이며 버전별 배포 EXE를 별도로 저장하는 방식이 아닙니다.
최종 배포 파일은 언제나 고정 이름 `ReportAutomationLauncher.exe` 하나입니다.

## 빌드 및 검증

개발 PC의 저장소 `.venv`에 `requirements.txt`의 의존성과 빌드용 `packaging`이
설치되어 있어야 합니다. 이미 설치된 패키지를 읽으며 빌드 중 새 패키지를 설치하지 않습니다.
Python 3.12 이상 Windows x64 환경을 사용하고 의존성 버전을 변경할 때는
[제3자 고지](../THIRD_PARTY_NOTICES.md)도 함께 갱신해야 합니다.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\report_automation_launcher\scripts\build_report_automation_launcher.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\test_standalone_launcher.ps1
```

빌드는 필요한 의존성만 ZIP으로 묶고, 외부 Python 경로 없이 압축 해제된 런타임을
검증한 뒤 C# EXE의 리소스로 포함합니다. 같은 폴더의 임시 파일을 원자적으로 교체하며
빌드·검증·교체가 실패하면 기존 실행 파일을 유지합니다. GitHub 파일 크기 제한도 확인합니다.

검증 스크립트는 EXE만 공백이 포함된 별도 폴더에 복사하고 Python/엔진 관련 환경변수를
잘못된 경로로 지정합니다. 내장 의존성 import, 기본 PPTX 템플릿 생성·검사를 확인하며
실제 런처 폼의 숨김 초기화, 기존 서식·편집용지 검사와 캐시 재사용도 확인합니다.
사용자 Excel/아래한글 문서를 열거나 COM 객체를 생성하지 않습니다.
이 검증은 실제 Excel/HWP 자동화 통합 테스트를 대체하지 않습니다.

런타임 경로와 검사 결과 JSON은 다음 명령으로 확인할 수 있습니다.

```powershell
.\report_automation_launcher\bin\ReportAutomationLauncher.exe --self-check-standalone --out "C:\path\standalone_report.json"
```

빠른 소스 개발에만 `-DevelopmentOnly`를 사용할 수 있습니다. 이 빌드는 내장 런타임이
없으며 환경변수/저장소 경로에 의존합니다. 이름은 같으므로 사용자 배포 전 반드시 기본
빌드로 다시 생성하고 독립 실행 검증을 수행해야 합니다.
