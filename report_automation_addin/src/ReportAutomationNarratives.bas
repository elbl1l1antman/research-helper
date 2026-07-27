Attribute VB_Name = "ReportAutomationNarratives"
' ============================================================
' 모듈명  : ReportAutomationNarratives
' 설  명  : 분석문, 핵심 포인트, 차트 데이터, 삽입표, QA 산출
' ============================================================
Option Explicit

' ============================================================
Public Sub ReportAutomation_WriteNarratives(ByVal wsNarr As Worksheet, ByVal wsChart As Worksheet, ByVal wsInsert As Worksheet, _
                                             ByVal wsQA As Worksheet, ByVal tables As Collection, ByVal dataWs As Worksheet, _
                                             ByVal wsSettings As Worksheet, ByRef qaCount As Long)
    wsNarr.Range("A1:I1").Value = Array("table_key", "문항/표 제목", "분석문_기본", "분석문_세부", "분석문_LLM", _
                                        "주요 수치 요약", "검토 상태", "사용자 수정문", "최종 사용문")
    wsChart.Range("A1:H1").Value = Array("table_key", "series_group", "category", "measure", "value", "display_value", "sort_order", "include_chart")
    wsInsert.Range("A1:H1").Value = Array("table_key", "title", "category", "weighted_n", "raw_n", "percent", "unit", "source_cell")
    wsQA.Range("A1:F1").Value = Array("table_key", "qa_type", "severity", "message", "source_range", "checked_at")

    ReportAutomation_StyleHeader wsNarr.Range("A1:I1")
    ReportAutomation_StyleHeader wsChart.Range("A1:H1")
    ReportAutomation_StyleHeader wsInsert.Range("A1:H1")
    ReportAutomation_StyleHeader wsQA.Range("A1:F1")

    ' 사용자가 설정 시트에서 선택한 배너 목록과 제목 제거 접두어를 읽는다.
    Dim bannerSetting As String
    bannerSetting = ReportAutomation_ReadBannerSetting(wsSettings)

    Dim titlePrefixes() As String
    titlePrefixes = ReportAutomation_ReadTitlePrefixes(wsSettings)

    Dim narrRow As Long: narrRow = 2
    Dim chartRow As Long: chartRow = 2
    Dim insertRow As Long: insertRow = 2
    Dim qaRow As Long: qaRow = 2

    Dim i As Long, rec As Variant
    For i = 1 To tables.Count
        rec = tables(i)
        Application.StatusBar = "보고서 자동화 처리 중: " & i & " / " & tables.Count & "번째 표  (Esc로 취소)"

        ' points는 분석문/차트/삽입표에 공통으로 쓰는 핵심 수치 목록이다.
        ' 각 원소는 Array(category, percent, weighted_n, raw_n, source_cell) 구조를 따른다.
        Dim points As Collection
        Set points = ReportAutomation_ExtractKeyPoints(dataWs, rec)

        ' 보고서 본문용 문장과 리뷰용 수치 요약을 각각 생성한다.
        Dim narrative As String, summary As String
        narrative = ReportAutomation_BuildNarrative(CStr(rec(IDX_TITLE)), CStr(rec(IDX_BASIS)), points, titlePrefixes)
        summary = ReportAutomation_FormatPointSummary(points)

        wsNarr.Cells(narrRow, 1).Value = rec(IDX_TABLE_KEY)
        wsNarr.Cells(narrRow, 2).Value = rec(IDX_TITLE)
        wsNarr.Cells(narrRow, 3).Value = narrative
        wsNarr.Cells(narrRow, 6).Value = summary
        wsNarr.Cells(narrRow, 7).Value = IIf(points.Count = 0, "확인 필요", "자동 생성")
        wsNarr.Cells(narrRow, 9).Value = narrative
        narrRow = narrRow + 1

        ' 전체 기준 수치를 차트 데이터와 삽입용 표 데이터로 펼친다.
        Dim p As Long, pointRec As Variant
        For p = 1 To points.Count
            pointRec = points(p)

            wsChart.Cells(chartRow, 1).Value = rec(IDX_TABLE_KEY)
            wsChart.Cells(chartRow, 2).Value = "전체"
            wsChart.Cells(chartRow, 3).Value = pointRec(0)
            wsChart.Cells(chartRow, 4).Value = ReportAutomation_PointMeasure(pointRec)
            wsChart.Cells(chartRow, 5).Value = pointRec(1)
            wsChart.Cells(chartRow, 6).Value = ReportAutomation_FormatPointValue(pointRec)
            wsChart.Cells(chartRow, 7).Value = p
            wsChart.Cells(chartRow, 8).Value = IIf(p <= CHART_INCLUDE_TOP_N, "Y", "N")
            chartRow = chartRow + 1

            wsInsert.Cells(insertRow, 1).Value = rec(IDX_TABLE_KEY)
            wsInsert.Cells(insertRow, 2).Value = rec(IDX_TITLE)
            wsInsert.Cells(insertRow, 3).Value = pointRec(0)
            wsInsert.Cells(insertRow, 4).Value = pointRec(2)
            wsInsert.Cells(insertRow, 5).Value = pointRec(3)
            wsInsert.Cells(insertRow, 6).Value = pointRec(1)
            wsInsert.Cells(insertRow, 7).Value = ReportAutomation_PointMeasure(pointRec)
            wsInsert.Cells(insertRow, 8).Value = pointRec(4)
            insertRow = insertRow + 1
        Next p

        ' 사용자가 선택한 배너가 있고 배너 교차표인 경우에만 배너별 수치를 차트 데이터에 추가한다.
        If CLng(rec(IDX_LAST_COL)) > SIMPLE_TABLE_MAX_COL Then
            Dim bannerGroups As Collection
            Set bannerGroups = ReportAutomation_FindBannerGroups(dataWs, CLng(rec(IDX_START_ROW)), _
                                                                  ReportAutomation_FindTotalRow(dataWs, CLng(rec(IDX_START_ROW)), CLng(rec(IDX_END_ROW))), _
                                                                  CLng(rec(IDX_LAST_COL)))
            ReportAutomation_AppendBannerChartRows wsChart, chartRow, dataWs, rec, bannerGroups, bannerSetting
        End If

        ' 분석문 생성 실패 또는 구조 경고가 있는 표는 QA 시트에 남겨 수동 검토 대상으로 만든다.
        If points.Count = 0 Or Len(CStr(rec(IDX_WARNING))) > 0 Then
            wsQA.Cells(qaRow, 1).Value = rec(IDX_TABLE_KEY)
            wsQA.Cells(qaRow, 2).Value = IIf(points.Count = 0, "수치 QA", "표 구조 QA")
            wsQA.Cells(qaRow, 3).Value = "warning"
            wsQA.Cells(qaRow, 4).Value = IIf(points.Count = 0, "분석문 생성에 사용할 전체 기준 비율 값을 찾지 못했습니다.", rec(IDX_WARNING))
            wsQA.Cells(qaRow, 5).Value = dataWs.Range(dataWs.Cells(CLng(rec(IDX_START_ROW)), 1), dataWs.Cells(CLng(rec(IDX_END_ROW)), CLng(rec(IDX_LAST_COL)))).Address(False, False)
            wsQA.Cells(qaRow, 6).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
            qaRow = qaRow + 1
        End If
    Next i

    ' QA 요약 행 - 총 처리 표 수와 경고 건수를 한 눈에 파악할 수 있도록 마지막에 기록한다.
    qaCount = qaRow - 2
    If qaCount > 0 Then
        wsQA.Cells(qaRow + 1, 1).Value = "[요약]"
        wsQA.Cells(qaRow + 1, 4).Value = "총 " & tables.Count & "개 표 처리 / 경고 " & qaCount & "건"
        wsQA.Cells(qaRow + 1, 6).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
        ReportAutomation_StyleHeader wsQA.Range(wsQA.Cells(qaRow + 1, 1), wsQA.Cells(qaRow + 1, 6))
    End If

    wsNarr.Columns("A:I").AutoFit
    wsNarr.Columns("C:I").ColumnWidth = 36
    wsNarr.Columns("C:I").WrapText = True
    wsChart.Columns("A:H").AutoFit
    wsInsert.Columns("A:H").AutoFit
    wsQA.Columns("A:F").AutoFit
    wsNarr.Rows(1).AutoFilter
    wsChart.Rows(1).AutoFilter
    wsInsert.Rows(1).AutoFilter
    wsQA.Rows(1).AutoFilter
End Sub
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
Private Sub ReportAutomation_AppendBannerChartRows(ByVal wsChart As Worksheet, ByRef chartRow As Long, _
                                                    ByVal ws As Worksheet, ByVal tableRec As Variant, _
                                                    ByVal bannerGroups As Collection, ByVal bannerSetting As String)
    Dim startRow As Long, endRow As Long, totalRow As Long
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
