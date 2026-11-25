# 배포 파일

버전별 폴더에는 설치·실행에 필요한 파일만 둡니다. 실행 파일 하나에 Python 런타임, 엔진, Excel 추가기능, 기본 템플릿, 라이선스 고지가 모두 들어 있습니다.

- `ReportAutomationLauncher.exe`: 원하는 폴더에 복사해 바로 실행합니다(별도 설치 없음). 필요 환경은 [독립 실행 런처](../docs/standalone_launcher.md)를 보세요.
- `SHA256SUMS.txt`: 받은 파일이 배포본과 같은지 확인합니다. PowerShell: `Get-FileHash .\ReportAutomationLauncher.exe`

폴더는 한 번 올리면 바꾸지 않습니다. 새 버전은 Windows에서 빌드·검증한 뒤 `report_automation_launcher/scripts/publish_release.py`로 추가합니다.

최신 버전: `0.0.56`

| 버전 | 빌드 기준 커밋 | 빌드 일시 | 크기 | SHA-256 | 변경 내용 |
| --- | --- | --- | --- | --- | --- |
| [`0.0.56`](v0.0.56/) | `15fae8e` | 2026-11-18 | 31.1 MB | `e8acc8a78d6c82d8…` | HWPX 우선 런처 작업 흐름, 표·양식 준비 상태 표시, 차트 서식 관리 분리, 진행·저장·경고·실패 구분. 0.0.55에서 빌드한 실행 파일과 동일(0.0.56은 실행 파일 변경 없음). |
