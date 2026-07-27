Attribute VB_Name = "ReportAutomationBannerCharts"
' ============================================================
' 모듈명  : ReportAutomationBannerCharts
' 설  명  : 사용자가 선택한 배너 그룹을 차트 데이터 행으로 변환
' ============================================================
Option Explicit

' ============================================================
Private Function ReportAutomation_FindBannerGroups(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal lastCol As Long) As Collection
    Dim groups As New Collection
    Dim headerStart As Long, headerEnd As Long
    Dim c As Long, r As Long
    Dim cellText As String
    Dim colGroupName() As String
    Dim curName As String
    Dim colList As Collection
    Dim lastSeenGroupName As String

    If totalRow = 0 Then
        Set ReportAutomation_FindBannerGroups = groups
        Exit Function
    End If

    ' 헤더 행 범위: 제목 다음 행 ~ 전체 행 위 행
    headerStart = startRow + 1
    headerEnd = totalRow - 1
    If headerStart > headerEnd Then
        Set ReportAutomation_FindBannerGroups = groups
        Exit Function
    End If

    ' 열별로 가장 위에 있는 배너 그룹명을 수집한다.
    ReDim colGroupName(1 To lastCol)
    For c = 1 To lastCol
        For r = headerStart To headerEnd
            cellText = ReportAutomation_CellText(ws, r, c)
            If Len(cellText) > 0 And cellText <> "%" And UCase$(cellText) <> "N" And cellText <> "사례수" Then
                colGroupName(c) = cellText
                lastSeenGroupName = cellText
                Exit For
            End If
        Next r
        If Len(colGroupName(c)) = 0 And ReportAutomation_IsPercentMeasureColumn(ws, startRow, totalRow, c) Then
            colGroupName(c) = lastSeenGroupName
        End If
    Next c

    ' 연속된 같은 그룹명 구간을 하나의 배너 그룹으로 묶는다.
    curName = ""
    Set colList = New Collection
    For c = 1 To lastCol
        If colGroupName(c) <> curName Then
            If Len(curName) > 0 And colList.Count > 0 Then
                groups.Add Array(curName, ReportAutomation_CollectionToArray(colList))
            End If
            curName = colGroupName(c)
            Set colList = New Collection
        End If
        If ReportAutomation_IsPercentMeasureColumn(ws, startRow, totalRow, c) Then
            colList.Add c
        End If
    Next c
    If Len(curName) > 0 And colList.Count > 0 Then
        groups.Add Array(curName, ReportAutomation_CollectionToArray(colList))
    End If

    Set ReportAutomation_FindBannerGroups = groups
End Function

' 프로시저 : ReportAutomation_AppendBannerChartRows
' 설  명  : 사용자가 선택한 배너 그룹의 데이터를 차트 시트에 추가 기록한다.
' 인  자  : chartRow — 참조 전달(ByRef), 기록 후 다음 빈 행으로 이동한다.
'           bannerSetting — 쉼표 구분 배너 목록 문자열 (예: "전체,성별")
' 주  의  : "전체"는 WriteNarratives에서 이미 기록하므로 이 함수에서는 건너뛴다.
' ============================================================
Public Sub ReportAutomation_AppendBannerChartRows(ByVal wsChart As Worksheet, ByRef chartRow As Long, _
                                                    ByVal ws As Worksheet, ByVal tableRec As Variant, _
                                                    ByVal bannerSetting As String)
    Dim startRow As Long, endRow As Long, totalRow As Long
    Dim bannerGroups As Collection
    Dim bg As Variant
    Dim bgName As String
    Dim pctCols As Variant
    Dim colIdx As Long, r As Long
    Dim pctCol As Long
    Dim subLabel As String
    Dim sortOrder As Long
    Dim category As String

    startRow = CLng(tableRec(IDX_START_ROW))
    endRow   = CLng(tableRec(IDX_END_ROW))
    totalRow = ReportAutomation_FindTotalRow(ws, startRow, endRow)
    If totalRow = 0 Then Exit Sub

    Set bannerGroups = ReportAutomation_FindBannerGroups(ws, startRow, totalRow, CLng(tableRec(IDX_LAST_COL)))

    For Each bg In bannerGroups
        bgName = CStr(bg(0))

        ' "전체"는 이미 기록됐으므로 건너뛴다.
        If bgName = "전체" Then GoTo NextGroup

        ' 사용자가 선택하지 않은 배너는 건너뛴다.
        If Not ReportAutomation_BannerRequested(bgName, bannerSetting) Then GoTo NextGroup

        pctCols = bg(1)
        If Not IsArray(pctCols) Then GoTo NextGroup

        ' 각 % 열마다 응답 항목 행을 순회해 기록한다.
        For colIdx = LBound(pctCols) To UBound(pctCols)
            pctCol = CLng(pctCols(colIdx))

            ' 이 % 열의 하위 범주 레이블을 헤더에서 읽는다.
            subLabel = ReportAutomation_HeaderLabel(ws, startRow + 1, totalRow - 1, pctCol)
            If Len(subLabel) = 0 Then subLabel = bgName

            ' 항목 행을 먼저 수집한 뒤 값 기준 내림차순으로 정렬한다.
            Dim rowCandidates As Collection
            Set rowCandidates = New Collection
            For r = startRow + 3 To endRow
                If r = totalRow Then GoTo NextRow

                category = Trim$(CStr(ws.Cells(r, 2).Value))
                If Len(category) = 0 Then category = Trim$(CStr(ws.Cells(r, 1).Value))
                If Len(category) = 0 Then GoTo NextRow
                If Left$(category, 1) = "■" Or Left$(category, 1) = "▣" Then GoTo NextRow

                If IsNumeric(ws.Cells(r, pctCol).Value) Then
                    rowCandidates.Add Array(category, CDbl(ws.Cells(r, pctCol).Value))
                End If
NextRow:
            Next r

            If rowCandidates.Count > 0 Then
                Dim rowsArr As Variant, rowIdx As Long, rowRec As Variant
                rowsArr = ReportAutomation_CollectionToArray(rowCandidates)
                ReportAutomation_SortArrayDescending rowsArr, 1

                For rowIdx = LBound(rowsArr) To UBound(rowsArr)
                    rowRec = rowsArr(rowIdx)
                    sortOrder = rowIdx - LBound(rowsArr) + 1
                    wsChart.Cells(chartRow, 1).Value = tableRec(IDX_TABLE_KEY)
                    wsChart.Cells(chartRow, 2).Value = subLabel
                    wsChart.Cells(chartRow, 3).Value = rowRec(0)
                    wsChart.Cells(chartRow, 4).Value = "%"
                    wsChart.Cells(chartRow, 5).Value = CDbl(rowRec(1))
                    wsChart.Cells(chartRow, 6).Value = ReportAutomation_FormatPercent(CDbl(rowRec(1)))
                    wsChart.Cells(chartRow, 7).Value = sortOrder
                    wsChart.Cells(chartRow, 8).Value = IIf(sortOrder <= CHART_INCLUDE_TOP_N, "Y", "N")
                    chartRow = chartRow + 1
                Next rowIdx
            End If
        Next colIdx
NextGroup:
    Next bg
End Sub
' ============================================================
Private Function ReportAutomation_BannerRequested(ByVal bannerName As String, ByVal setting As String) As Boolean
    Dim tokens() As String
    tokens = Split(setting, ",")
    Dim t As Long
    For t = LBound(tokens) To UBound(tokens)
        If Trim$(tokens(t)) = bannerName Then
            ReportAutomation_BannerRequested = True
            Exit Function
        End If
    Next t
End Function
