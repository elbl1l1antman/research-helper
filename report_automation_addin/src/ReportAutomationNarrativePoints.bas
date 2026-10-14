Attribute VB_Name = "ReportAutomationNarrativePoints"
' ============================================================
' 모듈명  : ReportAutomationNarrativePoints
' 설  명  : 집계표에서 분석문/차트/삽입표에 사용할 핵심 수치 추출
' ============================================================
Option Explicit

' ============================================================
Public Function ReportAutomation_ExtractKeyPoints(ByVal ws As Worksheet, ByVal tableRec As Variant) As Collection
    Dim points As New Collection
    If InStr(1, CStr(tableRec(IDX_WARNING)), "좌우 표 경계", vbTextCompare) > 0 Then
        Set ReportAutomation_ExtractKeyPoints = points
        Exit Function
    End If
    Dim startRow As Long, endRow As Long, lastCol As Long, firstCol As Long, bodyRow As Long
    startRow = ReportAutomation_HeaderRowIndex(tableRec) - 1
    endRow = CLng(tableRec(IDX_END_ROW))
    lastCol = CLng(tableRec(IDX_LAST_COL))
    firstCol = ReportAutomation_FirstCol(tableRec)
    bodyRow = ReportAutomation_BodyRowIndex(tableRec)

    If ReportAutomation_IsScoreTable(ws, CStr(tableRec(IDX_TITLE)), startRow, bodyRow - 1, lastCol, firstCol) Then
        ReportAutomation_ExtractScorePoints ws, startRow, endRow, lastCol, points, firstCol, bodyRow
        ReportAutomation_SortPointsDescending points
        Set ReportAutomation_ExtractKeyPoints = points
        Exit Function
    End If

    ' 열 수가 적은 표는 일반 분포표, 넓은 표는 배너 교차표로 보고 다른 파서를 적용한다.
    If lastCol - firstCol + 1 <= SIMPLE_TABLE_MAX_COL Then
        ReportAutomation_ExtractSimplePoints ws, startRow, endRow, lastCol, points, firstCol, bodyRow
    Else
        ReportAutomation_ExtractWidePoints ws, startRow, endRow, lastCol, points, firstCol, bodyRow
    End If

    ReportAutomation_SortPointsDescending points
    Set ReportAutomation_ExtractKeyPoints = points
End Function
' ============================================================
Private Function ReportAutomation_IsScoreTable(ByVal ws As Worksheet, ByVal titleText As String, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal firstCol As Long) As Boolean
    Dim hintText As String
    hintText = titleText & " " & ReportAutomation_HeaderBlockText(ws, startRow, endRow, lastCol, firstCol)

    Dim totalRow As Long
    totalRow = endRow + 1

    Dim scoreColCount As Long, pctColCount As Long
    If totalRow > 0 Then
        scoreColCount = ReportAutomation_CountScoreMeasureColumns(ws, startRow, totalRow, lastCol, firstCol)
        pctColCount = ReportAutomation_CountPercentMeasureColumns(ws, startRow, totalRow, lastCol, firstCol)
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
Private Function ReportAutomation_HeaderBlockText(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal firstCol As Long) As String
    Dim limitRow As Long
    limitRow = endRow

    Dim r As Long, c As Long, text As String
    For r = startRow + 1 To limitRow
        For c = firstCol To lastCol
            text = ReportAutomation_CellText(ws, r, c)
            If Len(text) > 0 Then
                If Len(ReportAutomation_HeaderBlockText) > 0 Then ReportAutomation_HeaderBlockText = ReportAutomation_HeaderBlockText & " "
                ReportAutomation_HeaderBlockText = ReportAutomation_HeaderBlockText & text
            End If
        Next c
    Next r
End Function
' ============================================================
Private Sub ReportAutomation_ExtractScorePoints(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal points As Collection, ByVal firstCol As Long, ByVal bodyRow As Long)
    Dim scoreCol As Long
    scoreCol = ReportAutomation_FindScoreColumn(ws, startRow, bodyRow - 1, lastCol, firstCol)

    If scoreCol > 0 And lastCol - firstCol + 1 <= SIMPLE_TABLE_MAX_COL Then
        Dim r As Long, category As String
        Dim declaredStub As Long
        declaredStub = ReportAutomation_StubColumn(ws, startRow + 1, bodyRow - 1, firstCol, lastCol)
        For r = bodyRow To endRow
            category = ReportAutomation_RowLabel(ws, r, firstCol, lastCol, declaredStub)
            If Len(category) > 0 And Left$(category, 1) <> "■" And Left$(category, 1) <> "▣" Then
                If IsNumeric(ws.Cells(r, scoreCol).Value) Then
                    points.Add Array(category, CDbl(ws.Cells(r, scoreCol).Value), ReportAutomation_RowCount(ws, r, startRow + 1, bodyRow - 1, firstCol, lastCol), "", ws.Cells(r, scoreCol).Address(False, False), "score_100")
                End If
            End If
        Next r
        Exit Sub
    End If

    Dim totalRow As Long
    totalRow = ReportAutomation_FindTotalRow(ws, bodyRow, endRow, firstCol, lastCol)
    If totalRow = 0 Then Exit Sub

    Dim c As Long, label As String, scoreValue As Variant
    For c = firstCol To lastCol
        If ReportAutomation_IsScoreMeasureColumn(ws, startRow, bodyRow, c) Then
            scoreValue = ws.Cells(totalRow, c).Value
            If IsNumeric(scoreValue) Then
                If CDbl(scoreValue) >= 0 And CDbl(scoreValue) <= 100 Then
                    label = ReportAutomation_ScoreHeaderLabel(ws, startRow + 1, bodyRow - 1, c, firstCol)
                    If ReportAutomation_IsReportItemLabel(label) Then
                        points.Add Array(label, CDbl(scoreValue), "", ReportAutomation_RowCount(ws, totalRow, startRow + 1, bodyRow - 1, firstCol, lastCol), ws.Cells(totalRow, c).Address(False, False), "score_100")
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
        text = ReportAutomation_CellText(ws, r, colIndex)
        If InStr(1, text, "100점", vbTextCompare) > 0 _
            Or InStr(1, text, "점수", vbTextCompare) > 0 Then
            ReportAutomation_IsScoreMeasureColumn = True
            Exit Function
        End If
    Next r
End Function
' ============================================================
Private Function ReportAutomation_CountScoreMeasureColumns(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal lastCol As Long, ByVal firstCol As Long) As Long
    Dim c As Long
    For c = firstCol To lastCol
        If ReportAutomation_IsScoreMeasureColumn(ws, startRow, totalRow, c) Then
            ReportAutomation_CountScoreMeasureColumns = ReportAutomation_CountScoreMeasureColumns + 1
        End If
    Next c
End Function
' ============================================================
Private Function ReportAutomation_CountPercentMeasureColumns(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal lastCol As Long, ByVal firstCol As Long) As Long
    Dim c As Long
    For c = firstCol To lastCol
        If ReportAutomation_IsPercentMeasureColumn(ws, startRow, totalRow, c) Then
            ReportAutomation_CountPercentMeasureColumns = ReportAutomation_CountPercentMeasureColumns + 1
        End If
    Next c
End Function
' ============================================================
Private Function ReportAutomation_ScoreHeaderLabel(ByVal ws As Worksheet, ByVal firstHeaderRow As Long, ByVal lastHeaderRow As Long, ByVal colIndex As Long, ByVal firstCol As Long) As String
    Dim r As Long, c As Long, text As String
    For r = firstHeaderRow To lastHeaderRow
        For c = colIndex To firstCol Step -1
            text = ReportAutomation_CellText(ws, r, c)
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
    If ReportAutomation_IsPercentHeader(text) Or ReportAutomation_IsCountHeader(text) _
        Or ReportAutomation_IsTotalLabel(text) Or ReportAutomation_IsMetadata(text) Then Exit Function
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
Private Function ReportAutomation_FindScoreColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal firstCol As Long) As Long
    Dim r As Long, c As Long, headerText As String
    Dim limitRow As Long
    limitRow = endRow

    For r = startRow + 1 To limitRow
        For c = firstCol To lastCol
            headerText = ReportAutomation_CellText(ws, r, c)
            If InStr(1, headerText, "100점", vbTextCompare) > 0 _
                Or InStr(1, headerText, "점수", vbTextCompare) > 0 Then
                ReportAutomation_FindScoreColumn = c
                Exit Function
            End If
        Next c
    Next r
End Function
' ============================================================
Private Sub ReportAutomation_ExtractSimplePoints(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal points As Collection, ByVal firstCol As Long, ByVal bodyRow As Long)
    Dim pctCol As Long
    pctCol = ReportAutomation_FindHeaderColumn(ws, startRow + 1, bodyRow - 1, lastCol, "%", firstCol)
    ' 측정 헤더가 없으면 임의의 숫자 열을 비율로 해석하지 않는다.
    If pctCol = 0 Then Exit Sub

    Dim r As Long
    Dim percentValue As Variant
    Dim declaredStub As Long
    declaredStub = ReportAutomation_StubColumn(ws, startRow + 1, bodyRow - 1, firstCol, lastCol)
    For r = bodyRow To endRow
        Dim category As String
        category = ReportAutomation_RowLabel(ws, r, firstCol, lastCol, declaredStub)
        If Len(category) > 0 And Left$(category, 1) <> "■" And Left$(category, 1) <> "▣" Then
            percentValue = ReportAutomation_PercentValue(ws.Cells(r, pctCol))
            If Not IsEmpty(percentValue) Then
                points.Add Array(category, CDbl(percentValue), _
                    ReportAutomation_RowCount(ws, r, startRow + 1, bodyRow - 1, firstCol, lastCol), _
                    ReportAutomation_RowCount(ws, r, startRow + 1, bodyRow - 1, firstCol, lastCol, 2), ws.Cells(r, pctCol).Address(False, False))
            End If
        End If
    Next r
End Sub
' ============================================================
Private Sub ReportAutomation_ExtractWidePoints(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal points As Collection, ByVal firstCol As Long, ByVal bodyRow As Long)
    Dim totalRow As Long
    totalRow = ReportAutomation_FindTotalRow(ws, bodyRow, endRow, firstCol, lastCol)
    If totalRow = 0 Then Exit Sub

    ' 전체 행의 N/% 쌍 중 % 열만 읽어 응답 범주별 비율 후보로 만든다.
    Dim c As Long
    Dim percentValue As Variant
    For c = firstCol To lastCol
        percentValue = ReportAutomation_PercentValue(ws.Cells(totalRow, c))
        If ReportAutomation_IsPercentMeasureColumn(ws, startRow, bodyRow, c) And Not IsEmpty(percentValue) Then
            Dim category As String
            category = ReportAutomation_HeaderLabel(ws, startRow + 1, bodyRow - 1, c, firstCol)
            If ReportAutomation_IsReportItemLabel(category) And Not ReportAutomation_IsNeutralCompositeLabel(category) Then
                points.Add Array(category, CDbl(percentValue), "", ReportAutomation_RowCount(ws, totalRow, startRow + 1, bodyRow - 1, firstCol, lastCol), ws.Cells(totalRow, c).Address(False, False))
            End If
        End If
    Next c

    ' 7점/5점 척도형 표에는 분포율과 함께 100점 평균 열이 붙는 경우가 많다.
    ' 이 값은 순위 후보가 아니라 본문 첫 문장에 붙이는 보조 지표로 별도 kind를 부여한다.
    For c = firstCol To lastCol
        If ReportAutomation_IsScoreMeasureColumn(ws, startRow, bodyRow, c) And IsNumeric(ws.Cells(totalRow, c).Value) Then
            points.Add Array("100점 평균", CDbl(ws.Cells(totalRow, c).Value), "", ReportAutomation_RowCount(ws, totalRow, startRow + 1, bodyRow - 1, firstCol, lastCol), ws.Cells(totalRow, c).Address(False, False), "scale_score_100")
            Exit For
        End If
    Next c
End Sub
' ============================================================
Private Function ReportAutomation_FindHeaderColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal headerText As String, ByVal firstCol As Long) As Long
    Dim r As Long, c As Long
    For r = startRow To endRow
        For c = firstCol To lastCol
            If ReportAutomation_IsPercentHeader(ReportAutomation_CellText(ws, r, c)) Then
                ReportAutomation_FindHeaderColumn = c
                Exit Function
            End If
        Next c
    Next r
End Function
' ============================================================
Public Function ReportAutomation_FindTotalRow(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, Optional ByVal firstCol As Long = 1, Optional ByVal lastCol As Long = 0) As Long
    If lastCol = 0 Then lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    Dim values As Variant, r As Long, c As Long, hasTotal As Boolean
    values = ws.Range(ws.Cells(startRow, firstCol), ws.Cells(endRow, lastCol)).Value2
    If Not IsArray(values) Then Exit Function
    For r = 1 To UBound(values, 1)
        hasTotal = False
        For c = 1 To UBound(values, 2)
            If ReportAutomation_ArrayNumeric(values, r, c) Then
                If hasTotal Then ReportAutomation_FindTotalRow = startRow + r - 1: Exit Function
                Exit For
            End If
            If ReportAutomation_IsTotalLabel(ReportAutomation_ArrayText(values, r, c)) Then hasTotal = True
        Next c
    Next r
End Function
' ============================================================
Public Function ReportAutomation_IsPercentMeasureColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal colIndex As Long) As Boolean
    Dim r As Long
    For r = startRow + 1 To totalRow - 1
        If ReportAutomation_IsPercentHeader(ReportAutomation_CellText(ws, r, colIndex)) Then
            ReportAutomation_IsPercentMeasureColumn = True
            Exit Function
        End If
    Next r
End Function
' ============================================================
Public Function ReportAutomation_HeaderLabel(ByVal ws As Worksheet, ByVal firstHeaderRow As Long, ByVal lastHeaderRow As Long, ByVal colIndex As Long, Optional ByVal firstCol As Long = 1) As String
    Dim r As Long, c As Long, text As String

    ' 먼저 현재 열 자체의 병합 헤더를 확인한다. 계 열은 왼쪽 항목으로 대체하지 않고 계로 반환해 후속 필터에서 제외한다.
    For r = lastHeaderRow To firstHeaderRow Step -1
        text = ReportAutomation_CellText(ws, r, colIndex)
        If ReportAutomation_IsTotalLabel(text) Then
            ReportAutomation_HeaderLabel = text
            Exit Function
        End If
        If ReportAutomation_IsHeaderLabelCandidate(text) Then
            ReportAutomation_HeaderLabel = ReportAutomation_CleanHeaderItemLabel(text)
            Exit Function
        End If
    Next r

    For r = lastHeaderRow To firstHeaderRow Step -1
        For c = colIndex To firstCol Step -1
            text = ReportAutomation_CellText(ws, r, c)
            If ReportAutomation_IsTotalLabel(text) Then
                ReportAutomation_HeaderLabel = text
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
    If ReportAutomation_IsPercentHeader(text) Or ReportAutomation_IsCountHeader(text) _
        Or ReportAutomation_IsTotalLabel(text) Or ReportAutomation_IsStubHeader(text) Then Exit Function
    If ReportAutomation_IsMetadata(text) Or ReportAutomation_IsTableTitle(text) Then Exit Function
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
    If ReportAutomation_IsTotalLabel(text) Or ReportAutomation_IsCountHeader(text) _
        Or ReportAutomation_IsPercentHeader(text) Or ReportAutomation_IsMetadata(text) Then Exit Function
    ReportAutomation_IsReportItemLabel = True
End Function
' ============================================================
Public Function ReportAutomation_RowCount(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal firstHeader As Long, ByVal lastHeader As Long, ByVal firstCol As Long, ByVal lastCol As Long, Optional ByVal ordinal As Long = 1) As Variant
    Dim c As Long, r As Long, found As Long
    ReportAutomation_RowCount = ""
    For c = firstCol To lastCol
        For r = firstHeader To lastHeader
            If ReportAutomation_IsCountHeader(ReportAutomation_CellText(ws, r, c)) Then
                found = found + 1
                If found = ordinal Then ReportAutomation_RowCount = ws.Cells(rowIndex, c).Value2: Exit Function
                Exit For
            End If
        Next r
    Next c
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
Private Function ReportAutomation_IsNeutralCompositeLabel(ByVal text As String) As Boolean
    text = Replace(ReportAutomation_CleanText(text), " ", "")
    ReportAutomation_IsNeutralCompositeLabel = (InStr(1, text, "보통(", vbTextCompare) > 0)
End Function
' ============================================================
Public Function ReportAutomation_FormatPointSummary(ByVal points As Collection) As String
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
Public Function ReportAutomation_PointKind(ByVal pointRec As Variant) As String
    On Error GoTo Fallback
    If UBound(pointRec) >= 5 Then
        ReportAutomation_PointKind = CStr(pointRec(5))
        If Len(ReportAutomation_PointKind) > 0 Then Exit Function
    End If
Fallback:
    ReportAutomation_PointKind = "percent"
End Function
' ============================================================
Public Function ReportAutomation_PointMeasure(ByVal pointRec As Variant) As String
    If ReportAutomation_PointKind(pointRec) = "score_100" Then
        ReportAutomation_PointMeasure = "100점"
    Else
        ReportAutomation_PointMeasure = "%"
    End If
End Function
' ============================================================
Public Function ReportAutomation_FormatPointValue(ByVal pointRec As Variant) As String
    If ReportAutomation_PointKind(pointRec) = "score_100" Then
        ReportAutomation_FormatPointValue = ReportAutomation_FormatScore(CDbl(pointRec(1)))
    Else
        ReportAutomation_FormatPointValue = ReportAutomation_FormatPercent(CDbl(pointRec(1)))
    End If
End Function
' ============================================================
