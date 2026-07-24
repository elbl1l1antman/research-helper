Attribute VB_Name = "ReportAutomationSettings"
' ============================================================
' 모듈명  : ReportAutomationSettings
' 설  명  : 보고서 자동화 설정 시트 읽기/쓰기
' ============================================================
Option Explicit

' ============================================================
' 함수명 : ReportAutomation_FindLatestOutputSheet
' 설  명 : 특정 접두어로 시작하는 산출 시트 중 가장 뒤에 있는 최신 시트를 찾는다.
Public Function ReportAutomation_FindLatestOutputSheet(ByVal wb As Workbook, ByVal baseName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Left$(ws.Name, Len(baseName)) = baseName Then
            Set ReportAutomation_FindLatestOutputSheet = ws
        End If
    Next ws
End Function
' ============================================================
' 함수명 : ReportAutomation_SettingValue
' 설  명 : 이전 설정 시트에서 같은 라벨의 값을 읽고 없으면 기본값을 반환한다.
Public Function ReportAutomation_SettingValue(ByVal labelText As String, ByVal defaultValue As Variant, ByVal priorSettingsWs As Worksheet) As Variant
    If priorSettingsWs Is Nothing Then
        ReportAutomation_SettingValue = defaultValue
        Exit Function
    End If

    Dim r As Long
    For r = 3 To priorSettingsWs.Cells(priorSettingsWs.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(priorSettingsWs.Cells(r, 1).Value)) = labelText Then
            ReportAutomation_SettingValue = priorSettingsWs.Cells(r, 2).Value
            Exit Function
        End If
    Next r

    ReportAutomation_SettingValue = defaultValue
End Function
' ============================================================
' 프로시저 : ReportAutomation_SetSettingValue
' 설  명  : 생성된 설정 시트에서 특정 라벨의 값을 덮어쓴다.
Public Sub ReportAutomation_SetSettingValue(ByVal ws As Worksheet, ByVal labelText As String, ByVal valueText As String)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ws.Cells(r, 1).Value)) = labelText Then
            ws.Cells(r, 2).Value = valueText
            Exit Sub
        End If
    Next r
End Sub
' ============================================================
' 함수명 : ReportAutomation_ReadBannerSetting
' 설  명 : 설정 시트의 추출 배너 목록을 읽고 값이 없으면 전체를 반환한다.
Public Function ReportAutomation_ReadBannerSetting(ByVal ws As Worksheet) As String
    If ws Is Nothing Then
        ReportAutomation_ReadBannerSetting = "전체"
        Exit Function
    End If

    Dim r As Long
    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "추출 배너 목록" Then
            Dim val As String
            val = Trim$(CStr(ws.Cells(r, 2).Value))
            ReportAutomation_ReadBannerSetting = IIf(Len(val) > 0, val, "전체")
            Exit Function
        End If
    Next r
    ReportAutomation_ReadBannerSetting = "전체"
End Function
' ============================================================
' 함수명 : ReportAutomation_ReadTitlePrefixes
' 설  명 : 설정 시트의 제목 제거 접두어를 쉼표 기준 배열로 반환한다.
Public Function ReportAutomation_ReadTitlePrefixes(ByVal ws As Worksheet) As String()
    Dim defaultArr() As String
    ReDim defaultArr(0 To 0)
    defaultArr(0) = ""

    If ws Is Nothing Then
        ReportAutomation_ReadTitlePrefixes = defaultArr
        Exit Function
    End If

    Dim r As Long
    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "제목 제거 접두어" Then
            Dim val As String
            val = Trim$(CStr(ws.Cells(r, 2).Value))
            If Len(val) = 0 Then
                ReportAutomation_ReadTitlePrefixes = defaultArr
            Else
                ReportAutomation_ReadTitlePrefixes = Split(val, ",")
            End If
            Exit Function
        End If
    Next r
    ReportAutomation_ReadTitlePrefixes = defaultArr
End Function
