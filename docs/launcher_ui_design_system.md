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
- 폼 전체 스타일: `LauncherUi.ApplyToForm`
- 하위 컨트롤 일괄 적용: `LauncherUi.ApplyTree`

## 화면 규칙

- 새 버튼은 직접 색상을 지정하지 않고 `LauncherUi.StyleButton`을 사용한다.
- 새 상태 문구는 성공/경고/오류 의미에 맞는 semantic 색상을 사용한다.
- 새 입력 컨트롤은 `ApplyTree` 적용 범위 안에 두고, 개별 색상 지정은 피한다.
- 새 탭은 기존 `CreateStepPage`를 사용한다.
- 표와 목록은 `ListView` 또는 `CheckedListBox`를 우선 사용하고 배경은 흰색으로 유지한다.
- 설명 문구는 `ColorMutedText`를 사용한다.

## 현재 적용 범위

- 메인 폼 배경/크기/기본 글꼴
- 정보 카드형 상단 헤더
- 워크플로우 단계 상태 라벨
- 탭 헤더 owner-draw 스타일
- 문서형 그룹 패널 owner-draw 스타일
- 버튼 3종 스타일
- TextBox, ComboBox, ListView, CheckedListBox 기본 스타일
- 준비 체크리스트 상태색

## 다음 개선 대상

- 긴 화면을 기능별 카드형 섹션으로 재배치
- 대시보드 PPT 탭의 KPI/차트 매핑을 wizard 방식으로 분리
- 오류/경고를 별도 alert 컴포넌트로 통합
- 버튼 텍스트 기반 자동 분류를 명시적 variant 지정 방식으로 교체
