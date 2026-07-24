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
Private Function ReportAutomation_ExtractKeyPoints(ByVal ws As Worksheet, ByVal tableRec As Variant) As Collection
    Dim points As New Collection
    Dim startRow As Long, endRow As Long, lastCol As Long
    startRow = CLng(tableRec(IDX_START_ROW))
    endRow = CLng(tableRec(IDX_END_ROW))
    lastCol = CLng(tableRec(IDX_LAST_COL))

    If ReportAutomation_IsScoreTable(ws, CStr(tableRec(IDX_TITLE)), startRow, endRow, lastCol) Then
        ReportAutomation_ExtractScorePoints ws, startRow, endRow, lastCol, points
        ReportAutomation_SortPointsDescending points
        Set ReportAutomation_ExtractKeyPoints = points
        Exit Function
    End If

    ' 열 수가 적은 표는 일반 분포표, 넓은 표는 배너 교차표로 보고 다른 파서를 적용한다.
    If lastCol <= SIMPLE_TABLE_MAX_COL Then
        ReportAutomation_ExtractSimplePoints ws, startRow, endRow, lastCol, points
    Else
        ReportAutomation_ExtractWidePoints ws, startRow, endRow, lastCol, points
    End If

    ReportAutomation_SortPointsDescending points
    Set ReportAutomation_ExtractKeyPoints = points
End Function
' ============================================================
Private Function ReportAutomation_IsScoreTable(ByVal ws As Worksheet, ByVal titleText As String, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As Boolean
    Dim hintText As String
    hintText = titleText & " " & ReportAutomation_HeaderBlockText(ws, startRow, endRow, lastCol)

    Dim totalRow As Long
    totalRow = ReportAutomation_FindTotalRow(ws, startRow, endRow)

    Dim scoreColCount As Long, pctColCount As Long
    If totalRow > 0 Then
        scoreColCount = ReportAutomation_CountScoreMeasureColumns(ws, startRow, totalRow, lastCol)
        pctColCount = ReportAutomation_CountPercentMeasureColumns(ws, startRow, totalRow, lastCol)
    End If

    If pctColCount > 0 Then Exit Function

    If InStr(1, titleText, "100점", vbTextCompare) > 0 Then
        ReportAutomation_IsScoreTable = True
        Exit Function
    End If

    If scoreColCount >= 2 Then
        ReportAutomation_IsScoreTable = True
        Exit Function
    End If

    If scoreColCount = 1 And pctColCount = 0 Then
        If InStr(1, hintText, "100점", vbTextCompare) > 0 _
            Or InStr(1, hintText, "점수", vbTextCompare) > 0 Then
            ReportAutomation_IsScoreTable = True
        End If
    End If
End Function
' ============================================================
Private Function ReportAutomation_HeaderBlockText(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As String
    Dim limitRow As Long
    limitRow = startRow + 5
    If limitRow > endRow Then limitRow = endRow

    Dim r As Long, c As Long, text As String
    For r = startRow + 1 To limitRow
        For c = 1 To lastCol
            text = Trim$(CStr(ws.Cells(r, c).Value))
            If Len(text) > 0 Then
                If Len(ReportAutomation_HeaderBlockText) > 0 Then ReportAutomation_HeaderBlockText = ReportAutomation_HeaderBlockText & " "
                ReportAutomation_HeaderBlockText = ReportAutomation_HeaderBlockText & text
            End If
        Next c
    Next r
End Function
' ============================================================
Private Sub ReportAutomation_ExtractScorePoints(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal points As Collection)
    Dim scoreCol As Long
    scoreCol = ReportAutomation_FindScoreColumn(ws, startRow, endRow, lastCol)

    If scoreCol > 0 And lastCol <= SIMPLE_TABLE_MAX_COL Then
        Dim r As Long, category As String
        For r = startRow + 1 To endRow
            category = ReportAutomation_RowCategory(ws, r)
            If Len(category) > 0 And Left$(category, 1) <> "■" And Left$(category, 1) <> "▣" Then
                If IsNumeric(ws.Cells(r, scoreCol).Value) Then
                    points.Add Array(category, CDbl(ws.Cells(r, scoreCol).Value), ws.Cells(r, 3).Value, "", ws.Cells(r, scoreCol).Address(False, False), "score_100")
                End If
            End If
        Next r
        Exit Sub
    End If

    Dim totalRow As Long
    totalRow = ReportAutomation_FindTotalRow(ws, startRow, endRow)
    If totalRow = 0 Then Exit Sub

    Dim c As Long, label As String, scoreValue As Variant
    For c = 1 To lastCol
        If ReportAutomation_IsScoreMeasureColumn(ws, startRow, totalRow, c) Then
            scoreValue = ws.Cells(totalRow, c).Value
            If IsNumeric(scoreValue) Then
                If CDbl(scoreValue) >= 0 And CDbl(scoreValue) <= 100 Then
                    label = ReportAutomation_ScoreHeaderLabel(ws, startRow + 1, totalRow - 1, c)
                    If Len(label) > 0 And label <> "계" And label <> "사례수" Then
                        points.Add Array(label, CDbl(scoreValue), "", ws.Cells(totalRow, 3).Value, ws.Cells(totalRow, c).Address(False, False), "score_100")
                    End If
                End If
            End If
        End If
    Next c
End Sub
' ============================================================
Private Function ReportAutomation_IsScoreMeasureColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal colIndex As Long) As Boolean
    Dim r As Long, text As String
    For r = startRow + 1 To totalRow - 1
        text = Trim$(CStr(ws.Cells(r, colIndex).Value))
        If InStr(1, text, "100점", vbTextCompare) > 0 _
            Or InStr(1, text, "점수", vbTextCompare) > 0 Then
            ReportAutomation_IsScoreMeasureColumn = True
            Exit Function
        End If
    Next r
End Function
' ============================================================
Private Function ReportAutomation_CountScoreMeasureColumns(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = 1 To lastCol
        If ReportAutomation_IsScoreMeasureColumn(ws, startRow, totalRow, c) Then
            ReportAutomation_CountScoreMeasureColumns = ReportAutomation_CountScoreMeasureColumns + 1
        End If
    Next c
End Function
' ============================================================
Private Function ReportAutomation_CountPercentMeasureColumns(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = 1 To lastCol
        If ReportAutomation_IsPercentMeasureColumn(ws, startRow, totalRow, c) Then
            ReportAutomation_CountPercentMeasureColumns = ReportAutomation_CountPercentMeasureColumns + 1
        End If
    Next c
End Function
' ============================================================
Private Function ReportAutomation_ScoreHeaderLabel(ByVal ws As Worksheet, ByVal firstHeaderRow As Long, ByVal lastHeaderRow As Long, ByVal colIndex As Long) As String
    Dim r As Long, c As Long, text As String
    For r = firstHeaderRow To lastHeaderRow
        For c = colIndex To 1 Step -1
            text = Trim$(CStr(ws.Cells(r, c).Value))
            If ReportAutomation_IsScoreHeaderLabelCandidate(text) Then
                ReportAutomation_ScoreHeaderLabel = ReportAutomation_CleanScoreItemLabel(text)
                Exit Function
            End If
        Next c
    Next r
End Function
' ============================================================
Private Function ReportAutomation_IsScoreHeaderLabelCandidate(ByVal text As String) As Boolean
    If Len(text) = 0 Then Exit Function
    If text = "%" Or UCase$(text) = "N" Or text = "사례수" Or text = "빈도" Then Exit Function
    If InStr(1, text, "100점", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "점수", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "평균", vbTextCompare) > 0 Then Exit Function
    If Left$(text, 1) = "▶" Then Exit Function
    ReportAutomation_IsScoreHeaderLabelCandidate = True
End Function
' ============================================================
Private Function ReportAutomation_CleanScoreItemLabel(ByVal text As String) As String
    text = ReportAutomation_CleanText(text)
    Do While Len(text) > 0 And InStr(1, "○□", Left$(text, 1), vbBinaryCompare) > 0
        text = Trim$(Mid$(text, 2))
    Loop
    ReportAutomation_CleanScoreItemLabel = text
End Function
' ============================================================
Private Function ReportAutomation_FindScoreColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As Long
    Dim r As Long, c As Long, headerText As String
    Dim limitRow As Long
    limitRow = startRow + 5
    If limitRow > endRow Then limitRow = endRow

    For r = startRow + 1 To limitRow
        For c = 1 To lastCol
            headerText = Trim$(CStr(ws.Cells(r, c).Value))
            If InStr(1, headerText, "100점", vbTextCompare) > 0 _
                Or InStr(1, headerText, "점수", vbTextCompare) > 0 Then
                ReportAutomation_FindScoreColumn = c
                Exit Function
            End If
        Next c
    Next r
End Function
' ============================================================
Private Function ReportAutomation_RowCategory(ByVal ws As Worksheet, ByVal rowIndex As Long) As String
    ReportAutomation_RowCategory = Trim$(CStr(ws.Cells(rowIndex, 2).Value))
    If Len(ReportAutomation_RowCategory) = 0 Then
        ReportAutomation_RowCategory = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
    End If
End Function
' ============================================================
Private Sub ReportAutomation_ExtractSimplePoints(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal points As Collection)
    Dim pctCol As Long
    pctCol = ReportAutomation_FindHeaderColumn(ws, startRow, endRow, lastCol, "%")
    If pctCol = 0 Then pctCol = DEFAULT_PCT_COL

    Dim r As Long
    For r = startRow + 3 To endRow
        Dim category As String
        ' SPSS 표 구조에 따라 항목명이 B열 또는 A열에 올 수 있어 B열 우선, A열 보조로 읽는다.
        category = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(category) = 0 Then category = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(category) > 0 And Left$(category, 1) <> "■" And Left$(category, 1) <> "▣" Then
            If IsNumeric(ws.Cells(r, pctCol).Value) Then
                points.Add Array(category, CDbl(ws.Cells(r, pctCol).Value), ws.Cells(r, 3).Value, ws.Cells(r, 4).Value, ws.Cells(r, pctCol).Address(False, False))
            End If
        End If
    Next r
End Sub
' ============================================================
Private Sub ReportAutomation_ExtractWidePoints(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal points As Collection)
    Dim totalRow As Long
    totalRow = ReportAutomation_FindTotalRow(ws, startRow, endRow)
    If totalRow = 0 Then Exit Sub

    ' 전체 행의 N/% 쌍 중 % 열만 읽어 응답 범주별 비율 후보로 만든다.
    Dim c As Long
    For c = 4 To lastCol
        If ReportAutomation_IsPercentMeasureColumn(ws, startRow, totalRow, c) And IsNumeric(ws.Cells(totalRow, c).Value) Then
            Dim category As String
            category = ReportAutomation_HeaderLabel(ws, startRow + 1, totalRow - 1, c)
            If ReportAutomation_IsReportItemLabel(category) And Not ReportAutomation_IsNeutralCompositeLabel(category) Then
                points.Add Array(category, CDbl(ws.Cells(totalRow, c).Value), "", ws.Cells(totalRow, 3).Value, ws.Cells(totalRow, c).Address(False, False))
            End If
        End If
    Next c

    ' 7점/5점 척도형 표에는 분포율과 함께 100점 평균 열이 붙는 경우가 많다.
    ' 이 값은 순위 후보가 아니라 본문 첫 문장에 붙이는 보조 지표로 별도 kind를 부여한다.
    For c = 4 To lastCol
        If ReportAutomation_IsScoreMeasureColumn(ws, startRow, totalRow, c) And IsNumeric(ws.Cells(totalRow, c).Value) Then
            points.Add Array("100점 평균", CDbl(ws.Cells(totalRow, c).Value), "", ws.Cells(totalRow, 3).Value, ws.Cells(totalRow, c).Address(False, False), "scale_score_100")
            Exit For
        End If
    Next c
End Sub
' ============================================================
Private Function ReportAutomation_FindHeaderColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal headerText As String) As Long
    Dim r As Long, c As Long
    For r = startRow To Application.WorksheetFunction.Min(startRow + 5, endRow)
        For c = 1 To lastCol
            If Trim$(CStr(ws.Cells(r, c).Value)) = headerText Then
                ReportAutomation_FindHeaderColumn = c
                Exit Function
            End If
        Next c
    Next r
End Function
' ============================================================
Private Function ReportAutomation_FindTotalRow(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long
    For r = startRow To endRow
        If Left$(Trim$(CStr(ws.Cells(r, 1).Value)), 1) = "■" Then
            ReportAutomation_FindTotalRow = r
            Exit Function
        End If
    Next r
End Function
' ============================================================
Private Function ReportAutomation_IsPercentMeasureColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal colIndex As Long) As Boolean
    Dim r As Long
    For r = startRow + 1 To totalRow - 1
        If Trim$(CStr(ws.Cells(r, colIndex).Value)) = "%" Then
            ReportAutomation_IsPercentMeasureColumn = True
            Exit Function
        End If
    Next r
End Function
' ============================================================
Private Function ReportAutomation_HeaderLabel(ByVal ws As Worksheet, ByVal firstHeaderRow As Long, ByVal lastHeaderRow As Long, ByVal colIndex As Long) As String
    Dim r As Long, c As Long, text As String

    ' 먼저 현재 열 자체의 병합 헤더를 확인한다. 계 열은 왼쪽 항목으로 대체하지 않고 계로 반환해 후속 필터에서 제외한다.
    For r = lastHeaderRow To firstHeaderRow Step -1
        text = ReportAutomation_CellText(ws, r, colIndex)
        If text = "계" Then
            ReportAutomation_HeaderLabel = "계"
            Exit Function
        End If
        If ReportAutomation_IsHeaderLabelCandidate(text) Then
            ReportAutomation_HeaderLabel = ReportAutomation_CleanHeaderItemLabel(text)
            Exit Function
        End If
    Next r

    For r = lastHeaderRow To firstHeaderRow Step -1
        For c = colIndex To 1 Step -1
            text = ReportAutomation_CellText(ws, r, c)
            If text = "계" Then
                ReportAutomation_HeaderLabel = "계"
                Exit Function
            End If
            If ReportAutomation_IsHeaderLabelCandidate(text) Then
                ReportAutomation_HeaderLabel = ReportAutomation_CleanHeaderItemLabel(text)
                Exit Function
            End If
        Next c
    Next r
End Function
' ============================================================
Private Function ReportAutomation_IsHeaderLabelCandidate(ByVal text As String) As Boolean
    text = ReportAutomation_CleanText(text)
    If Len(text) = 0 Then Exit Function
    If text = "%" Or UCase$(text) = "N" Or text = "사례수" Or text = "빈도" Or text = "계" Then Exit Function
    If InStr(1, text, "BASE", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "100점", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "점수", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "평균", vbTextCompare) > 0 Then Exit Function
    ReportAutomation_IsHeaderLabelCandidate = True
End Function
' ============================================================
Private Function ReportAutomation_CleanHeaderItemLabel(ByVal text As String) As String
    text = ReportAutomation_CleanText(text)
    Do While Len(text) > 0 And InStr(1, "◐○□", Left$(text, 1), vbBinaryCompare) > 0
        text = Trim$(Mid$(text, 2))
    Loop
    ReportAutomation_CleanHeaderItemLabel = text
End Function
' ============================================================
Private Function ReportAutomation_IsReportItemLabel(ByVal text As String) As Boolean
    text = ReportAutomation_CleanText(text)
    If Len(text) = 0 Then Exit Function
    If text = "계" Or text = "빈도" Or text = "사례수" Or text = "%" Then Exit Function
    ReportAutomation_IsReportItemLabel = True
End Function
' ============================================================
Private Sub ReportAutomation_SortPointsDescending(ByVal points As Collection)
    If points.Count <= 1 Then Exit Sub

    Dim arr() As Variant
    ReDim arr(1 To points.Count)

    Dim i As Long, j As Long
    For i = 1 To points.Count
        arr(i) = points(i)
    Next i

    ' 데이터 수가 표별로 크지 않으므로 단순 비교 정렬로 충분하다.
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            Dim leftAux As Boolean, rightAux As Boolean
            leftAux = (ReportAutomation_PointKind(arr(i)) = "scale_score_100")
            rightAux = (ReportAutomation_PointKind(arr(j)) = "scale_score_100")
            If (leftAux And Not rightAux) Or _
               (leftAux = rightAux And CDbl(arr(j)(1)) > CDbl(arr(i)(1))) Then
                Dim temp As Variant
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    Do While points.Count > 0
        points.Remove 1
    Loop
    For i = LBound(arr) To UBound(arr)
        points.Add arr(i)
    Next i
End Sub
' ============================================================
Private Function ReportAutomation_BuildNarrative(ByVal titleText As String, ByVal basisText As String, ByVal points As Collection, ByRef titlePrefixes() As String) As String
    Dim normalizedTitle As String
    normalizedTitle = ReportAutomation_NormalizeAnalysisTitle(titleText, titlePrefixes)
    If Len(normalizedTitle) = 0 Then normalizedTitle = titleText

    If points.Count = 0 Then
        ReportAutomation_BuildNarrative = normalizedTitle & " 문항의 수치형 응답을 찾지 못했습니다."
        Exit Function
    End If

    Dim p1 As Variant
    p1 = points(1)

    Dim text As String
    If ReportAutomation_PointKind(p1) = "score_100" Then
        text = ReportAutomation_BuildScoreNarrative(normalizedTitle, points)
        ReportAutomation_BuildNarrative = text
        Exit Function
    End If

    Dim scaleScore As Variant
    scaleScore = ReportAutomation_FindPointByKind(points, "scale_score_100")
    If Not IsEmpty(scaleScore) And ReportAutomation_IsScaleSummaryPoint(CStr(p1(0))) Then
        text = normalizedTitle & "에 대해 조사한 결과, " & _
               ReportAutomation_Quoted(ReportAutomation_CleanScaleLabel(CStr(p1(0)))) & " 응답은 " & _
               ReportAutomation_FormatPercent(CDbl(p1(1))) & ", 100점 평균은 " & _
               ReportAutomation_FormatScore(CDbl(scaleScore(1))) & "으로 나타남"

        Dim contrastPoint As Variant
        contrastPoint = ReportAutomation_FindScaleContrastPoint(points)
        If Not IsEmpty(contrastPoint) Then
            text = text & vbLf & "반면, " & _
                   ReportAutomation_Quoted(ReportAutomation_CleanScaleLabel(CStr(contrastPoint(0)))) & _
                   ReportAutomation_TopicParticle(CStr(contrastPoint(0))) & " " & _
                   ReportAutomation_FormatPercent(CDbl(contrastPoint(1))) & "로 나타남"
        End If

        ReportAutomation_BuildNarrative = text
        Exit Function
    End If

    ' 복수응답 문항은 해석 기준을 명시해 단일응답 문항과 구분한다.
    If ReportAutomation_IsMultiResponse(titleText, basisText) Then
        text = normalizedTitle & ReportAutomation_TopicParticle(normalizedTitle) & " 복수응답 기준으로 " & _
               ReportAutomation_Quoted(CStr(p1(0))) & ReportAutomation_SubjectParticle(CStr(p1(0))) & " " & _
               ReportAutomation_FormatPercent(CDbl(p1(1))) & "로 가장 높게 나타남"
    Else
        text = normalizedTitle & ReportAutomation_TopicParticle(normalizedTitle) & " " & _
               ReportAutomation_Quoted(CStr(p1(0))) & ReportAutomation_SubjectParticle(CStr(p1(0))) & " " & _
               ReportAutomation_FormatPercent(CDbl(p1(1))) & "로 가장 높게 나타남"
    End If

    ' 2~4위 항목은 별도 줄로 구성해 본문 편집 시 첫 문장과 보조 문장을 나누기 쉽게 한다.
    Dim othersText As String, i As Long, maxItems As Long, pointRec As Variant
    maxItems = points.Count
    If maxItems > NARRATIVE_TOP_ITEMS Then maxItems = NARRATIVE_TOP_ITEMS
    For i = 2 To maxItems
        pointRec = points(i)
        If Len(othersText) > 0 Then othersText = othersText & ", "
        othersText = othersText & ReportAutomation_Quoted(CStr(pointRec(0))) & _
                     "(" & ReportAutomation_FormatPercent(CDbl(pointRec(1))) & ")"
    Next i
    If Len(othersText) > 0 Then
        text = text & vbLf & "다음으로 " & othersText & " 순으로 나타남"
    End If

    ReportAutomation_BuildNarrative = text
End Function
' ============================================================
Private Function ReportAutomation_BuildScoreNarrative(ByVal normalizedTitle As String, ByVal points As Collection) As String
    Dim p1 As Variant
    p1 = points(1)

    Dim displayTitle As String
    displayTitle = ReportAutomation_CleanScoreTitle(normalizedTitle)

    Dim text As String
    text = "‘" & displayTitle & "’" & ReportAutomation_ObjectParticle(displayTitle) & _
           " 100점 환산 기준으로 분석한 결과, " & _
           ReportAutomation_Quoted(CStr(p1(0))) & ReportAutomation_SubjectParticle(CStr(p1(0))) & " " & _
           ReportAutomation_FormatScore(CDbl(p1(1))) & "으로 가장 높게 나타남"

    Dim othersText As String, i As Long, maxItems As Long, pointRec As Variant
    maxItems = points.Count
    If maxItems > NARRATIVE_TOP_ITEMS Then maxItems = NARRATIVE_TOP_ITEMS
    For i = 2 To maxItems
        pointRec = points(i)
        If Len(othersText) > 0 Then othersText = othersText & ", "
        othersText = othersText & ReportAutomation_Quoted(CStr(pointRec(0))) & _
                     "(" & ReportAutomation_FormatScore(CDbl(pointRec(1))) & ")"
    Next i
    If Len(othersText) > 0 Then
        text = text & vbLf & "그다음으로는 " & othersText & " 등의 순으로 높게 나타남"
    End If

    ReportAutomation_BuildScoreNarrative = text
End Function
' ============================================================
Private Function ReportAutomation_FindPointByKind(ByVal points As Collection, ByVal pointKind As String) As Variant
    Dim i As Long, rec As Variant
    For i = 1 To points.Count
        rec = points(i)
        If ReportAutomation_PointKind(rec) = pointKind Then
            ReportAutomation_FindPointByKind = rec
            Exit Function
        End If
    Next i
End Function
' ============================================================
Private Function ReportAutomation_FindScaleContrastPoint(ByVal points As Collection) As Variant
    Dim i As Long, rec As Variant, label As String
    For i = 1 To points.Count
        rec = points(i)
        If ReportAutomation_PointKind(rec) <> "scale_score_100" Then
            label = CStr(rec(0))
            If InStr(1, label, "부정", vbTextCompare) > 0 _
                Or InStr(1, label, "비인지", vbTextCompare) > 0 _
                Or InStr(1, label, "불만족", vbTextCompare) > 0 _
                Or InStr(1, label, "비동의", vbTextCompare) > 0 _
                Or InStr(1, label, "불필요", vbTextCompare) > 0 Then
                ReportAutomation_FindScaleContrastPoint = rec
                Exit Function
            End If
        End If
    Next i
End Function
' ============================================================
Private Function ReportAutomation_IsScaleSummaryPoint(ByVal text As String) As Boolean
    If InStr(1, text, "부정", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "비인지", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "불만족", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "비동의", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "불필요", vbTextCompare) > 0 Then Exit Function

    If InStr(1, text, "긍정", vbTextCompare) > 0 _
        Or InStr(1, text, "만족", vbTextCompare) > 0 _
        Or InStr(1, text, "인지", vbTextCompare) > 0 _
        Or InStr(1, text, "필요함", vbTextCompare) > 0 _
        Or InStr(1, text, "동의", vbTextCompare) > 0 Then
        ReportAutomation_IsScaleSummaryPoint = True
    End If
End Function
' ============================================================
Private Function ReportAutomation_CleanScaleLabel(ByVal text As String) As String
    text = ReportAutomation_CleanText(text)
    text = Replace(text, " (", "(")
    text = Replace(text, "(합)", "(합)")
    ReportAutomation_CleanScaleLabel = text
End Function
' ============================================================
Private Function ReportAutomation_IsNeutralCompositeLabel(ByVal text As String) As Boolean
    text = Replace(ReportAutomation_CleanText(text), " ", "")
    ReportAutomation_IsNeutralCompositeLabel = (InStr(1, text, "보통(", vbTextCompare) > 0)
End Function
' ============================================================
Private Function ReportAutomation_CleanScoreTitle(ByVal titleText As String) As String
    Dim text As String
    text = ReportAutomation_CleanText(titleText)
    text = Replace(text, "(100점)", "")
    text = Replace(text, "（100점）", "")
    ReportAutomation_CleanScoreTitle = ReportAutomation_CleanText(text)
End Function
' ============================================================
Private Function ReportAutomation_FormatPointSummary(ByVal points As Collection) As String
    Dim i As Long, maxItems As Long, text As String, rec As Variant
    maxItems = points.Count
    If maxItems > SUMMARY_TOP_ITEMS Then maxItems = SUMMARY_TOP_ITEMS

    For i = 1 To maxItems
        rec = points(i)
        If Len(text) > 0 Then text = text & "; "
        text = text & CStr(rec(0)) & "=" & ReportAutomation_FormatPointValue(rec) & "(" & CStr(rec(4)) & ")"
    Next i
    ReportAutomation_FormatPointSummary = text
End Function
' ============================================================
Private Function ReportAutomation_NormalizeAnalysisTitle(ByVal titleText As String, ByRef userPrefixes() As String) As String
    Dim text As String
    Dim closePos As Long
    Dim pIdx As Long
    Dim pfx As String
    Dim marker As Variant
    Dim pos As Long

    text = ReportAutomation_CleanText(titleText)

    If Left$(text, 1) = "[" Then
        closePos = InStr(1, text, "]", vbTextCompare)
        ' 제목 앞의 [리스트 기준], [응답 기준] 등은 본문 문장에서는 제거한다.
        If closePos > 0 And closePos < Len(text) Then text = ReportAutomation_CleanText(Mid$(text, closePos + 1))
    End If

    ' 설정 시트 "제목 제거 접두어"에서 읽어온 목록을 순회해 제거한다.
    ' 사용자가 설정 시트 B열에 쉼표 구분으로 접두어를 추가하면 이 배열에 반영된다.
    For pIdx = LBound(userPrefixes) To UBound(userPrefixes)
        pfx = Trim$(userPrefixes(pIdx))
        If Len(pfx) > 0 Then
            If Left$(text, Len(pfx)) = pfx Then
                text = ReportAutomation_CleanText(Mid$(text, Len(pfx) + 1))
                Exit For
            End If
        End If
    Next pIdx

    ' 표 머리글이 제목에 붙어 들어온 경우, 분석 제목 뒤쪽의 메타 컬럼명을 잘라낸다.
    For Each marker In Array("사례수", "전체", "진흥지구명", "사업체 특성", "단위", "Base :")
        pos = InStr(1, text, CStr(marker), vbTextCompare)
        If pos > 1 Then
            text = ReportAutomation_CleanText(Left$(text, pos - 1))
            Exit For
        End If
    Next marker

    ReportAutomation_NormalizeAnalysisTitle = text
End Function

Private Function ReportAutomation_IsMultiResponse(ByVal titleText As String, ByVal basisText As String) As Boolean
    Dim text As String
    text = titleText & " " & basisText
    ReportAutomation_IsMultiResponse = (InStr(1, text, "복수", vbTextCompare) > 0 Or _
                                        InStr(1, text, "중복", vbTextCompare) > 0 Or _
                                        InStr(1, text, "multiple", vbTextCompare) > 0)
End Function

Private Function ReportAutomation_PointKind(ByVal pointRec As Variant) As String
    On Error GoTo Fallback
    If UBound(pointRec) >= 5 Then
        ReportAutomation_PointKind = CStr(pointRec(5))
        If Len(ReportAutomation_PointKind) > 0 Then Exit Function
    End If
Fallback:
    ReportAutomation_PointKind = "percent"
End Function
' ============================================================
Private Function ReportAutomation_PointMeasure(ByVal pointRec As Variant) As String
    If ReportAutomation_PointKind(pointRec) = "score_100" Then
        ReportAutomation_PointMeasure = "100점"
    Else
        ReportAutomation_PointMeasure = "%"
    End If
End Function
' ============================================================
Private Function ReportAutomation_FormatPointValue(ByVal pointRec As Variant) As String
    If ReportAutomation_PointKind(pointRec) = "score_100" Then
        ReportAutomation_FormatPointValue = ReportAutomation_FormatScore(CDbl(pointRec(1)))
    Else
        ReportAutomation_FormatPointValue = ReportAutomation_FormatPercent(CDbl(pointRec(1)))
    End If
End Function
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
