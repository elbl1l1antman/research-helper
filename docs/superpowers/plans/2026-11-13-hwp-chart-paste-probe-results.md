# PowerPoint 차트의 한글 메타파일 붙여넣기 실증

기준: [사용자 합의](../specs/2026-11-13-survey-chart-workflow-approved.md). 기존 Excel 기반 실증안 대신 가짜 차트만 담은 PowerPoint를 사용했다. `tests/probe_hwp_chart_emf.py`는 검증 전용이며 런처와 연결하지 않는다.

| 경로 | 사용자 조작 | HWPX 저장 결과 |
| --- | --- | --- |
| 빈 자리만 생성 (`empty`) | 없음 | 글자처럼 취급하는 1×1 표에 위쪽 그림 캡션, 앞자리 `3`, 자동 번호 `1, 2, 3` |
| PowerPoint 차트 복사 → 클립보드 형식을 EMF로 제한 → 한글 일반 붙이기 (`emf-only`) | 없음 | 셀 안에 `numberingType=NONE`인 그림 2개. 저장 시 두 파일이 WMF로 변환됨. 엄격한 EMF 검사 실패 |
| PowerPoint 차트 복사 → 한글 골라 붙이기 → **그림 (메타파일)** (`special`) | 사람이 항목을 선택해야 진행 | 두 그림 역시 WMF. 실제 한글 재열기 및 PDF에서 차트, 긴 보기명 줄바꿈, `3-1`, `3-2`, 빈 `3-3` 확인 |
| 골라 붙이기에서 키보드 선택 → PowerPoint 원본 EMF 보존 (`keyboard`) | 없음 | 붙여넣기 직후 한글은 WMF로 저장. 별도 최종 HWPX의 차트 미디어를 원본 EMF로 교체하고 한글에서 재열기·재저장해도 두 EMF의 SHA-256과 그림 번호가 동일함을 확인 |

재현: `.venv/Scripts/python.exe -B -X utf8 tests/probe_hwp_chart_emf.py --out-dir outputs/chart_paste_new --method keyboard`. 각 재실행은 새 폴더를 사용한다. 키보드 모드는 시험 문서와 같은 한글 프로세스의 대화상자에서 형식 목록에 아래 방향키를 두 번 보내고, 결과 설명이 `그림 (메타파일)`이며 창이 전면에 있을 때만 Enter를 누른다. `special`은 수동 검증에만 사용한다.

최종 저장 형식은 **EMF만 허용**한다. `chart_probe.hwpx`는 한글의 중간 WMF 산출물이고, `chart_probe_emf.hwpx`가 검사에 통과한 최종 예제다. PNG·WMF·OLE를 최종 결과로 대신 저장하지 않는다. 기존 HWPX는 덮어쓰지 않으며 원본 조사 자료는 사용하거나 커밋하지 않았다.

아직 런처나 실사용 보고서 작성기에는 연결하지 않았다. 다음 단계에서는 포커스 변경, 대화상자 제한 시간, 한글 버전별 목록 차이, 클립보드 개입 및 안전한 체크포인트를 검증한 뒤 통합한다. 현재 결과는 이 PC의 한글 2024와 PowerPoint에서 가짜 차트 2개로 확인한 범위다.
