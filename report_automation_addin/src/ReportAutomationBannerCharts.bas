Attribute VB_Name = "ReportAutomationBannerCharts"
' ============================================================
' 모듈명  : ReportAutomationBannerCharts
' 설  명  : 사용자가 선택한 배너 그룹을 차트 데이터 행으로 변환
' ============================================================
Option Explicit

' ============================================================
' 부모 배너만 인정한다. 응답 항목의 N/% 한 쌍은 별도 배너가 아니다.
Public Function ReportAutomation_FindBannerGroups(ByVal ws As Worksheet, ByVal firstHeader As Long, ByVal lastHeader As Long, ByVal lastCol As Long, Optional ByVal firstCol As Long = 1) As Collection
    Dim groups As New Collection, names() As String, r As Long, c As Long, edge As Long
    Dim text As String, pctCount As Long, k As Long, lowerLevel As Boolean
    ReDim names(firstCol To lastCol)
    For r = firstHeader To lastHeader - 1
        c = firstCol
        Do While c <= lastCol
            text = ReportAutomation_CellText(ws, r, c)
            edge = c
            Do While edge < lastCol
                If ReportAutomation_CellText(ws, r, edge + 1) <> text Then Exit Do
                edge = edge + 1
            Loop
            If Len(text) > 0 And Not IsNumeric(text) Then
                If Not ReportAutomation_IsCountHeader(text) And Not ReportAutomation_IsMeasureHeader(text) _
                    And Not ReportAutomation_IsTotalLabel(text) And Not ReportAutomation_IsMetadata(text) _
                    And Not ReportAutomation_IsTableTitle(text) And Not ReportAutomation_IsStubHeader(text) Then
                    pctCount = 0: lowerLevel = False
                    For k = c To edge
                        If ReportAutomation_IsPercentMeasureColumn(ws, r, lastHeader + 1, k) Then pctCount = pctCount + 1
                        If Len(ReportAutomation_CellText(ws, r + 1, k)) > 0 Then lowerLevel = True
                    Next k
                    ' 기존 표식은 그대로 보존하고 무표식 부모는 여러 측정 쌍을 요구한다.
                    If lowerLevel And (pctCount >= 2 Or (Left$(text, 1) = "◐" And pctCount >= 1)) Then
                        For k = c To edge
                            If Len(names(k)) = 0 Then names(k) = text
                        Next k
                    End If
                End If
            End If
            c = edge + 1
        Loop
    Next r
    Dim curName As String, cols As Collection
    Set cols = New Collection
    For c = firstCol To lastCol
        If names(c) <> curName Then
            If Len(curName) > 0 And cols.Count > 0 Then groups.Add Array(curName, ReportAutomation_CollectionToArray(cols))
            curName = names(c)
            Set cols = New Collection
        End If
        If Len(curName) > 0 And ReportAutomation_IsPercentMeasureColumn(ws, firstHeader - 1, lastHeader + 1, c) Then cols.Add c
    Next c
    If Len(curName) > 0 And cols.Count > 0 Then groups.Add Array(curName, ReportAutomation_CollectionToArray(cols))
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
    Dim startRow As Long, endRow As Long, totalRow As Long, firstCol As Long, bodyRow As Long
    Dim bannerGroups As Collection
    Dim bg As Variant
    Dim bgName As String
    Dim pctCols As Variant
    Dim colIdx As Long, r As Long
    Dim pctCol As Long
    Dim subLabel As String
    Dim sortOrder As Long
    Dim category As String
    Dim percentValue As Variant
    Dim declaredStub As Long

    startRow = ReportAutomation_HeaderRowIndex(tableRec) - 1
    endRow   = CLng(tableRec(IDX_END_ROW))
    firstCol = ReportAutomation_FirstCol(tableRec)
    bodyRow = ReportAutomation_BodyRowIndex(tableRec)
    declaredStub = ReportAutomation_StubColumn(ws, startRow + 1, bodyRow - 1, firstCol, CLng(tableRec(IDX_LAST_COL)))
    totalRow = ReportAutomation_FindTotalRow(ws, bodyRow, endRow, firstCol, CLng(tableRec(IDX_LAST_COL)))
    If totalRow = 0 Then Exit Sub

    Set bannerGroups = ReportAutomation_FindBannerGroups(ws, startRow + 1, bodyRow - 1, CLng(tableRec(IDX_LAST_COL)), firstCol)

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
            subLabel = ReportAutomation_HeaderLabel(ws, startRow + 1, bodyRow - 1, pctCol, firstCol)
            If Len(subLabel) = 0 Then subLabel = bgName

            ' 항목 행을 먼저 수집한 뒤 값 기준 내림차순으로 정렬한다.
            Dim rowCandidates As Collection
            Set rowCandidates = New Collection
            For r = bodyRow To endRow
                If r = totalRow Then GoTo NextRow

                category = ReportAutomation_RowLabel(ws, r, firstCol, CLng(tableRec(IDX_LAST_COL)), declaredStub)
                If Len(category) = 0 Then GoTo NextRow
                If Left$(category, 1) = "■" Or Left$(category, 1) = "▣" Then GoTo NextRow

                percentValue = ReportAutomation_PercentValue(ws.Cells(r, pctCol))
                If Not IsEmpty(percentValue) Then
                    rowCandidates.Add Array(category, CDbl(percentValue))
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
