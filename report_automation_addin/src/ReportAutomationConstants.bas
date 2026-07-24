Attribute VB_Name = "ReportAutomationConstants"
' ============================================================
' 모듈명  : ReportAutomationConstants
' 설  명  : 보고서 자동화 공통 상수
' ============================================================
Option Explicit

' 단순 분포표(좁은 표)와 배너 교차표를 구분하는 열 수 기준값.
Public Const SIMPLE_TABLE_MAX_COL As Long = 6
' 단순 분포표에서 % 열 위치를 헤더에서 못 찾을 때 쓰는 기본 열 번호.
Public Const DEFAULT_PCT_COL As Long = 5
' 분석문 본문에 포함하는 상위 항목 수 (1위 제외 보조 문장용).
Public Const NARRATIVE_TOP_ITEMS As Long = 4
' 수치 요약 문자열에 포함하는 상위 항목 수.
Public Const SUMMARY_TOP_ITEMS As Long = 5
' 차트 데이터 시트에서 "포함" 기본값으로 처리하는 최대 항목 순위.
Public Const CHART_INCLUDE_TOP_N As Long = 8
' tableRec 배열 인덱스 상수.
Public Const IDX_TABLE_KEY  As Long = 0
Public Const IDX_TABLE_NO   As Long = 1
Public Const IDX_TITLE      As Long = 2
Public Const IDX_BASIS      As Long = 3
Public Const IDX_BASE_LABEL As Long = 4
Public Const IDX_START_ROW  As Long = 5
Public Const IDX_END_ROW    As Long = 6
Public Const IDX_LAST_COL   As Long = 7
Public Const IDX_TABLE_TYPE As Long = 8
Public Const IDX_WARNING    As Long = 9
