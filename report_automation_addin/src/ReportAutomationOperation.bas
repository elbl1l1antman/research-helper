Attribute VB_Name = "ReportAutomationOperation"
' ============================================================
' 모듈명  : ReportAutomationOperation
' 설  명  : 보고서 자동화 실행 상태 관리와 내부 로그
' ============================================================
Option Explicit

' 사용자 통합문서 안에 생성하는 내부 로그 시트명. 일반 사용자에게는 보이지 않도록 VeryHidden 처리한다.
Private Const REPORT_LOG_SHEET As String = "_ReportAutomation_Log"
' Application 상태를 일시 변경하기 전 원래 값을 저장하는 모듈 전역 변수들.
' 오류가 발생해도 ReportAutomation_EndOperation에서 최대한 원복한다.
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents As Boolean
Private mPrevDisplayAlerts As Boolean
Private mPrevCalculation As XlCalculation
Private mPrevStatusBar As Variant
Private mPrevEnableCancelKey As XlEnableCancelKey

' ============================================================
' 프로시저 : ReportAutomation_BeginOperation
' 설  명  : 긴 작업 전 Excel 상태를 저장하고 화면갱신/이벤트/계산을 일시 중단한다.
Public Sub ReportAutomation_BeginOperation(ByVal statusText As String)
    On Error Resume Next
    mPrevScreenUpdating = Application.ScreenUpdating
    mPrevEnableEvents = Application.EnableEvents
    mPrevDisplayAlerts = Application.DisplayAlerts
    mPrevCalculation = Application.Calculation
    mPrevStatusBar = Application.StatusBar

    mPrevEnableCancelKey = Application.EnableCancelKey
    ' 수천 행 집계표에서 시트 생성/서식 적용을 반복하므로 성능 옵션을 임시로 낮춘다.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlErrorHandler
    Application.StatusBar = statusText
    On Error GoTo 0
End Sub
' ============================================================
' 프로시저 : ReportAutomation_EndOperation
' 설  명  : BeginOperation에서 변경한 Excel Application 상태를 원래대로 복원한다.
Public Sub ReportAutomation_EndOperation()
    On Error Resume Next
    Application.ScreenUpdating = mPrevScreenUpdating
    Application.EnableEvents = mPrevEnableEvents
    Application.DisplayAlerts = mPrevDisplayAlerts
    Application.Calculation = mPrevCalculation
    Application.EnableCancelKey = mPrevEnableCancelKey
    Application.StatusBar = mPrevStatusBar
    On Error GoTo 0
End Sub
' ============================================================
' 프로시저 : ReportAutomation_LogEvent
' 설  명  : 사용자 통합문서의 VeryHidden 로그 시트에 실행 결과를 누적 기록한다.
Public Sub ReportAutomation_LogEvent(ByVal wb As Workbook, ByVal actionName As String, ByVal targetName As String, ByVal statusText As String, ByVal detailText As String)
    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(REPORT_LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        ' 최초 실행 시에만 로그 시트를 만들고 헤더를 구성한다.
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = REPORT_LOG_SHEET
        ws.Range("A1:F1").Value = Array("logged_at", "version", "action", "target", "status", "detail")
        ReportAutomation_StyleHeader ws.Range("A1:F1")
        ws.Visible = xlSheetVeryHidden
    End If

    ' 로그는 append-only로 누적해 반복 실행/오류 이력을 보존한다.
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = REPORT_AUTOMATION_VERSION
    ws.Cells(nextRow, 3).Value = actionName
    ws.Cells(nextRow, 4).Value = targetName
    ws.Cells(nextRow, 5).Value = statusText
    ws.Cells(nextRow, 6).Value = detailText
End Sub
