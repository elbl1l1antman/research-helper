# 런처 UI 디자인 시스템

## 목적

런처 UI는 기능이 계속 늘어나는 구조이므로 화면마다 색상, 폰트, 버튼 모양을 직접 지정하지 않는다. 모든 화면은 `LauncherUi` 토큰을 기준으로 만든다.

## 토큰 구조

### Primitive

- 배경: `ColorBackground`
- 표면: `ColorSurface`, `ColorSurfaceAlt`
- 선: `ColorBorder`
- 본문: `ColorText`, `ColorMutedText`
- 주요색: `ColorPrimary`, `ColorPrimaryHover`
- 상태색: `ColorSuccess`, `ColorWarning`, `ColorDanger`
- 간격: `SpaceXs`, `SpaceSm`, `SpaceMd`, `SpaceLg`, `SpaceXl`
- 기본 컨트롤 높이: `ControlHeight`

### Semantic

- 배경 화면: `ColorBackground`
- 입력/목록 표면: `ColorSurface`
- 보조 영역: `ColorSurfaceAlt`
- 실행 가능/완료: `ColorSuccess`
- 확인 필요: `ColorWarning`
- 오류: `ColorDanger`

### Component

- 버튼: `LauncherButtonKind.Primary`, `Secondary`, `Ghost`
- 상태 라벨: `LauncherUi.StyleStatusLabel`
- 목록 행: `LauncherUi.StyleListItem`
- 폼 전체 스타일: `LauncherUi.ApplyToForm`
- 하위 컨트롤 일괄 적용: `LauncherUi.ApplyTree`

## 화면 규칙

- 새 버튼은 직접 색상을 지정하지 않고 `LauncherUi.StyleButton`을 사용한다.
- 버튼은 disabled 상태에서도 배경/글자색이 명확히 구분되어야 한다.
- 새 상태 문구는 성공/경고/오류 의미에 맞는 semantic 색상을 사용한다.
- 새 입력 컨트롤은 `ApplyTree` 적용 범위 안에 두고, 개별 색상 지정은 피한다.
- 새 탭은 기존 `CreateStepPage`를 사용한다.
- 표와 목록은 `ListView` 또는 `CheckedListBox`를 우선 사용하고 배경은 흰색으로 유지한다.
- 긴 목록은 줄무늬 행을 적용해 스캔성을 확보한다.
- 설명 문구는 `ColorMutedText`를 사용한다.

## 현재 적용 범위

- 메인 폼 배경/크기/기본 글꼴
- 정보 카드형 상단 헤더
- 워크플로우 단계 상태 라벨
- 탭 헤더 owner-draw 스타일
- 문서형 그룹 패널 owner-draw 스타일
- 버튼 3종 스타일
- disabled 버튼 스타일
- command bar형 버튼 묶음
- TextBox, ComboBox, ListView, CheckedListBox 기본 스타일
- 주요 ListView 줄무늬 행
- 준비 체크리스트 상태색

## 다음 개선 대상

## 보고서 흐름 (0.0.55)

- 원본 화면에서 한글 보고서 또는 Excel 산출 시트를 선택한다. 저장 방식과 실제 저장 폴더는 함께 표시한다.
- 표·분석 배너에서는 읽기 중, 탐지 실패, 전체 기준을 구분한다. 표 범위 수정과 복원은 해당 표를 선택한 경우에만 활성화한다.
- 보고서 설정에서 양식과 출력 문항 수를 확인한다. 변환 및 연결 진단은 고급 양식 도구에서 제공한다.
- 표·그림에서 기본 차트 서식 예시를 확인한다. 등록·삭제·CRTX 변환은 차트 서식 관리 창에서 수행한다.
- 결과 화면의 대표 명령은 실제 저장된 한글 보고서 또는 산출 Excel 열기다. QA와 진단 파일, PPTX·EMF 차트 작업은 별도 탭에서 제공한다.
