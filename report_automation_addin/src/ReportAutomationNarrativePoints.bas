Attribute VB_Name = "ReportAutomationNarrativePoints"
' ============================================================
' 모듈명  : ReportAutomationNarrativePoints
' 설  명  : 집계표에서 분석문/차트/삽입표에 사용할 핵심 수치 추출
' ============================================================
Option Explicit

' ============================================================
Public Function ReportAutomation_ExtractKeyPoints(ByVal ws As Worksheet, ByVal tableRec As Variant) As Collection
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
Public Function ReportAutomation_FindTotalRow(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long
    For r = startRow To endRow
        If Left$(Trim$(CStr(ws.Cells(r, 1).Value)), 1) = "■" Then
            ReportAutomation_FindTotalRow = r
            Exit Function
        End If
    Next r
End Function
' ============================================================
Public Function ReportAutomation_IsPercentMeasureColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal colIndex As Long) As Boolean
    Dim r As Long
    For r = startRow + 1 To totalRow - 1
        If Trim$(CStr(ws.Cells(r, colIndex).Value)) = "%" Then
            ReportAutomation_IsPercentMeasureColumn = True
            Exit Function
        End If
    Next r
End Function
' ============================================================
Public Function ReportAutomation_HeaderLabel(ByVal ws As Worksheet, ByVal firstHeaderRow As Long, ByVal lastHeaderRow As Long, ByVal colIndex As Long) As String
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
