# PowerPoint 차트의 한글 메타파일 붙여넣기 실증

기준: [사용자 합의](../specs/2026-11-13-survey-chart-workflow-approved.md). 기존 Excel 기반 실증안 대신 가짜 차트만 담은 PowerPoint를 사용했다. `tests/probe_hwp_chart_emf.py`는 검증 전용이며 런처와 연결하지 않는다.

| 경로 | 사용자 조작 | HWPX 저장 결과 |
| --- | --- | --- |
| 빈 자리만 생성 (`empty`) | 없음 | 글자처럼 취급하는 1×1 표에 위쪽 그림 캡션, 앞자리 `3`, 자동 번호 `1, 2, 3` |
| PowerPoint 차트 복사 → 클립보드 형식을 EMF로 제한 → 한글 일반 붙이기 (`emf-only`) | 없음 | 셀 안에 `numberingType=NONE`인 그림 2개. 저장 시 두 파일이 WMF로 변환됨. 엄격한 EMF 검사 실패 |
| PowerPoint 차트 복사 → 한글 골라 붙이기 → **그림 (메타파일)** (`special`) | 사람이 항목을 선택해야 진행 | 두 그림 역시 WMF. 실제 한글 재열기 및 PDF에서 차트, 긴 보기명 줄바꿈, `3-1`, `3-2`, 빈 `3-3` 확인 |

재현: `.venv/Scripts/python.exe -B -X utf8 tests/probe_hwp_chart_emf.py --out-dir outputs/chart_paste_new --method emf-only`. 각 재실행은 새 폴더를 사용한다. `special`은 수동 검증에만 사용한다. 자동 모드에 대화상자나 키 입력을 넣지 않는다. 둘 다 자동화를 제품으로 노출하는 승인 조건을 충족하지 않았다.

다음 단계는 **저장 형식 정책 결정**이다. 한글의 `그림 (메타파일)` 결과인 WMF도 허용한다면 두 경로의 도형·텍스트, 크기와 번호를 추가 비교해 무인 검증 경로를 구현한다. 저장 형식을 EMF로 고정한다면 EMF 보존 경로를 별도 검증하기 전까지 자동 삽입을 보류한다. 비트맵 또는 OLE로 대체하지 않는다. 원본 조사 자료는 사용하거나 커밋하지 않았다.
