Attribute VB_Name = "ReportAutomationTables"
' ============================================================
' 모듈명  : ReportAutomationTables
' 설  명  : 원본 집계표 시트 탐색, 표 블록 탐지, 표 제목/범위/유형 파싱
' ============================================================
Option Explicit

' ============================================================
Public Function ReportAutomation_ResolveDataSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    ' 실제 보고서용으로 배너를 정리한 시트가 함께 있으면 원본보다 조정본을 우선 사용한다.
    For Each ws In wb.Worksheets
        If InStr(1, ws.Name, "배너조정", vbTextCompare) > 0 Then
            If ReportAutomation_CountTableTitleRows(ws) > 0 Then
                Set ReportAutomation_ResolveDataSheet = ws
                Exit Function
            End If
        End If
    Next ws

    ' 사용자가 보고 있던 시트를 가장 먼저 신뢰한다.
    If TypeName(ActiveSheet) = "Worksheet" Then
        Set ws = ActiveSheet
        If Not ws.Parent Is Nothing Then
            If ws.Parent Is wb And ReportAutomation_CountTableTitleRows(ws) > 0 Then
                Set ReportAutomation_ResolveDataSheet = ws
                Exit Function
            End If
        End If
    End If

    ' SPSS/조사 집계표 샘플은 Sheet1에 본표가 있는 경우가 많아 두 번째 후보로 본다.
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("Sheet1")
    On Error GoTo 0
    If Not ws Is Nothing Then
        If ReportAutomation_CountTableTitleRows(ws) > 0 Then
            Set ReportAutomation_ResolveDataSheet = ws
            Exit Function
        End If
    End If

    ' 시트명이 바뀐 파일도 처리하기 위해 모든 워크시트를 마지막 후보로 순회한다.
    For Each ws In wb.Worksheets
        If ReportAutomation_CountTableTitleRows(ws) > 0 Then
            Set ReportAutomation_ResolveDataSheet = ws
            Exit Function
        End If
    Next ws
End Function
' ============================================================
Private Function ReportAutomation_CountTableTitleRows(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ReportAutomation_LastUsedRow(ws)
    If lastRow < 1 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).Value
    Dim r As Long
    For r = 1 To lastRow
        If ReportAutomation_IsTableTitle(arr(r, 1)) Then
            ReportAutomation_CountTableTitleRows = ReportAutomation_CountTableTitleRows + 1
        End If
    Next r
End Function
' ============================================================
Public Function ReportAutomation_CollectTables(ByVal ws As Worksheet) As Collection
    Dim titleRows As New Collection
    Dim lastRow As Long
    lastRow = ReportAutomation_LastUsedRow(ws)

    Dim r As Long
    For r = 1 To lastRow
        If ReportAutomation_IsTableTitle(ws.Cells(r, 1).Value) Then titleRows.Add r
    Next r

    ' 각 제목행부터 다음 제목행 직전까지를 하나의 표 블록으로 본다.
    Dim tables As New Collection
    Dim i As Long
    For i = 1 To titleRows.Count
        Dim startRow As Long, endRow As Long, lastCol As Long
        startRow = CLng(titleRows(i))
        If i < titleRows.Count Then
            endRow = CLng(titleRows(i + 1)) - 1
        Else
            endRow = ReportAutomation_LastNonEmptyRowInBlock(ws, startRow, lastRow)
        End If
        lastCol = ReportAutomation_LastNonEmptyColInBlock(ws, startRow, endRow)

        ' 제목 문자열에서 표번호, 분석 제목, 기준 문구를 분리한다.
        Dim titleText As String, tableNo As String, cleanTitle As String, basis As String
        titleText = CStr(ws.Cells(startRow, 1).Value)
        tableNo = ReportAutomation_ParseTableNo(titleText)
        cleanTitle = ReportAutomation_ParseTableTitle(titleText)
        basis = ReportAutomation_ParseBasis(titleText)

        Dim baseLabel As String
        If startRow + 1 <= endRow Then baseLabel = CStr(ws.Cells(startRow + 1, 1).Value)

        ' 배열 기반 레코드로 저장한다. 후속 단계에서 rec(index)로 빠르게 접근한다.
        tables.Add Array( _
            "T" & Format$(i, "0000"), tableNo, cleanTitle, basis, baseLabel, _
            startRow, endRow, lastCol, _
            ReportAutomation_ClassifyTable(ws, startRow, endRow, lastCol), _
            ReportAutomation_TableWarning(ws, startRow, endRow, lastCol))
    Next i

    Set ReportAutomation_CollectTables = tables
End Function
' ============================================================
Private Function ReportAutomation_LastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then
        ReportAutomation_LastUsedRow = 1
    Else
        ReportAutomation_LastUsedRow = lastCell.Row
    End If
End Function
' ============================================================
Private Function ReportAutomation_IsTableTitle(ByVal valueData As Variant) As Boolean
    Dim text As String
    text = Trim$(CStr(valueData))
    ReportAutomation_IsTableTitle = (Left$(text, 3) = "[표 " Or Left$(text, 3) = "[ 표")
End Function
' ============================================================
Private Function ReportAutomation_LastNonEmptyRowInBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal maxRow As Long) As Long
    Dim r As Long, c As Long
    For r = maxRow To startRow Step -1
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then
                ReportAutomation_LastNonEmptyRowInBlock = r
                Exit Function
            End If
        Next c
    Next r
    ReportAutomation_LastNonEmptyRowInBlock = startRow
End Function
' ============================================================
Private Function ReportAutomation_LastNonEmptyColInBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long, c As Long, lastCol As Long
    For r = startRow To endRow
        c = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If c > lastCol Then lastCol = c
    Next r
    If lastCol < 1 Then lastCol = 1
    ReportAutomation_LastNonEmptyColInBlock = lastCol
End Function
' ============================================================
Private Function ReportAutomation_ParseTableNo(ByVal titleText As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, titleText, "[표 ", vbTextCompare)
    If startPos = 0 Then startPos = InStr(1, titleText, "[ 표", vbTextCompare)
    endPos = InStr(1, titleText, ">", vbTextCompare)
    If startPos > 0 And endPos > startPos Then
        ReportAutomation_ParseTableNo = Trim$(Mid$(titleText, startPos + 3, endPos - startPos - 3))
        Exit Function
    End If

    If startPos > 0 Then
        endPos = InStr(startPos + 1, titleText, "]", vbTextCompare)
        If endPos > startPos Then
            ReportAutomation_ParseTableNo = Trim$(Replace(Mid$(titleText, startPos + 2, endPos - startPos - 2), "표", ""))
        End If
    End If
End Function
' ============================================================
Private Function ReportAutomation_ParseTableTitle(ByVal titleText As String) As String
    Dim endPos As Long
    endPos = InStr(1, titleText, ">", vbTextCompare)
    If endPos > 0 Then
        Dim parsedTitle As String
        parsedTitle = Trim$(Mid$(titleText, endPos + 1))
        If Right$(parsedTitle, 1) = "]" Then parsedTitle = Left$(parsedTitle, Len(parsedTitle) - 1)
        ReportAutomation_ParseTableTitle = Trim$(parsedTitle)
    Else
        Dim text As String
        text = ReportAutomation_CleanText(titleText)

        If Left$(text, 3) = "[ 표" Then
            endPos = InStr(1, text, "]", vbTextCompare)
            If endPos > 0 Then text = Trim$(Mid$(text, endPos + 1))

            If Left$(text, 1) = "[" Then
                endPos = InStr(1, text, "]", vbTextCompare)
                If endPos > 0 Then text = Trim$(Mid$(text, endPos + 1))
            End If

            Dim dividerPos As Long
            dividerPos = InStr(1, text, "─", vbTextCompare)
            If dividerPos > 1 Then text = Trim$(Left$(text, dividerPos - 1))

            Dim questionPos As Long
            questionPos = InStr(1, text, "[ 문", vbTextCompare)
            If questionPos > 1 Then text = Trim$(Left$(text, questionPos - 1))

            Dim cleanedText As String
            cleanedText = ReportAutomation_CleanText(text)
            If Left$(text, 1) = "'" And Left$(cleanedText, 1) <> "'" Then cleanedText = "'" & cleanedText
            ReportAutomation_ParseTableTitle = cleanedText
        Else
            ReportAutomation_ParseTableTitle = titleText
        End If
    End If
End Function
' ============================================================
Private Function ReportAutomation_ParseBasis(ByVal titleText As String) As String
    Dim openPos As Long, closePos As Long
    openPos = InStr(1, titleText, "> [", vbTextCompare)
    If openPos > 0 Then
        closePos = InStr(openPos + 3, titleText, "]", vbTextCompare)
        If closePos > openPos Then
            ReportAutomation_ParseBasis = Mid$(titleText, openPos + 2, closePos - openPos - 1)
        End If
    End If
End Function
' ============================================================
Private Function ReportAutomation_ClassifyTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As String
    If lastCol <= SIMPLE_TABLE_MAX_COL Then
        ReportAutomation_ClassifyTable = "단일 전체 분포표"
    ElseIf ReportAutomation_BlockContainsText(ws, startRow, endRow, lastCol, "▶평균◀") Then
        ReportAutomation_ClassifyTable = "평균 포함 표"
    ElseIf ReportAutomation_BlockContainsText(ws, startRow, startRow + 3, lastCol, "N") Then
        ReportAutomation_ClassifyTable = "N/% 쌍 표"
    Else
        ReportAutomation_ClassifyTable = "배너 교차표"
    End If
End Function
' ============================================================
Private Function ReportAutomation_BlockContainsText(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal needle As String) As Boolean
    If endRow > ws.Rows.Count Then endRow = ws.Rows.Count
    Dim arr As Variant
    arr = ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, lastCol)).Value
    If Not IsArray(arr) Then
        If InStr(1, CStr(arr), needle, vbTextCompare) > 0 Then ReportAutomation_BlockContainsText = True
        Exit Function
    End If
    Dim r As Long, c As Long
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If InStr(1, CStr(arr(r, c)), needle, vbTextCompare) > 0 Then
                ReportAutomation_BlockContainsText = True
                Exit Function
            End If
        Next c
    Next r
End Function
' ============================================================
Private Function ReportAutomation_TableWarning(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As String
    Dim warningText As String
    Dim baseCellValue As String
    If startRow + 1 <= endRow Then baseCellValue = CStr(ws.Cells(startRow + 1, 1).Value)
    If startRow + 1 > endRow Or _
       (InStr(1, baseCellValue, "BASE:", vbTextCompare) = 0 And InStr(1, baseCellValue, "Base :", vbTextCompare) = 0) Then
        warningText = "BASE 행 확인 필요"
    End If
    If lastCol <= 1 Then
        If Len(warningText) > 0 Then warningText = warningText & "; "
        warningText = warningText & "사용 열 확인 필요"
    End If
    ReportAutomation_TableWarning = warningText
End Function
