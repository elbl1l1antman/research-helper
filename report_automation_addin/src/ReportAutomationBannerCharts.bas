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
    Dim text As String, child As String, pctCount As Long, k As Long, level As Long, lowerLevel As Boolean
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
                    And Not ReportAutomation_IsTableTitle(text) And Not ReportAutomation_IsStubHeader(text) _
                    And Not ReportAutomation_IsScaleSummaryBanner(text) Then
                    pctCount = 0: lowerLevel = False
                    For k = c To edge
                        If ReportAutomation_IsPercentMeasureColumn(ws, r, lastHeader + 1, k) Then pctCount = pctCount + 1
                        If Not lowerLevel Then
                            For level = r + 1 To lastHeader
                                child = ReportAutomation_CellText(ws, level, k)
                                If child <> text And ReportAutomation_IsReportItemLabel(child) And Not ReportAutomation_IsMeasureHeader(child) _
                                    And Not ReportAutomation_IsStubHeader(child) And Not ReportAutomation_IsNumericText(child) Then
                                    lowerLevel = True: Exit For
                                End If
                            Next level
                        End If
                    Next k
                    ' 기존 표식은 그대로 보존하고 무표식 부모는 여러 측정 쌍을 요구한다.
                    If lowerLevel And (pctCount >= 2 Or ((Left$(text, 1) = "◐" Or Left$(text, 1) = "▣") And pctCount >= 1)) Then
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

Private Function ReportAutomation_BannerCountColumn(ByVal ws As Worksheet, ByVal firstHeader As Long, _
    ByVal bodyRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim c As Long, r As Long
    For c = firstCol To lastCol
        For r = firstHeader To bodyRow - 1
            If ReportAutomation_IsCountHeader(ReportAutomation_CellText(ws, r, c)) Then
                ReportAutomation_BannerCountColumn = c
                Exit Function
            End If
        Next r
    Next c
End Function

Public Function ReportAutomation_FindVerticalBannerRows(ByVal ws As Worksheet, ByVal firstHeader As Long, _
    ByVal bodyRow As Long, ByVal endRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Collection
    Dim groups As New Collection, countCol As Long, values As Variant, r As Long
    Dim raw As String, name As String, current As String
    Set ReportAutomation_FindVerticalBannerRows = groups
    countCol = ReportAutomation_BannerCountColumn(ws, firstHeader, bodyRow, firstCol, lastCol)
    If countCol < firstCol + 2 Or countCol >= lastCol Then Exit Function
    values = ws.Range(ws.Cells(bodyRow, countCol - 2), ws.Cells(endRow, lastCol)).Value2
    For r = 1 To UBound(values, 1)
        raw = ReportAutomation_ArrayText(values, r, 1)
        If Len(raw) > 0 Then current = ""
        If Not ReportAutomation_IsVerticalMember(values, r) Then
            current = ""
        ElseIf Len(raw) > 0 Then
            name = ReportAutomation_BannerName(raw)
            If Len(name) > 0 And Not ReportAutomation_IsTotalLabel(name) And Not ReportAutomation_IsMetadata(name) _
                And Not ReportAutomation_IsTableTitle(name) And Not ReportAutomation_IsNumericText(name) Then
                If Left$(raw, 1) = "▣" Then
                    current = name
                ElseIf r < UBound(values, 1) Then
                    If Len(ReportAutomation_ArrayText(values, r + 1, 1)) = 0 And ReportAutomation_IsVerticalMember(values, r + 1) Then current = name
                End If
            End If
        End If
        If Len(current) > 0 Then
            Dim countText As String: countText = ReportAutomation_BannerCountText(values, r, 3)
            If CDbl(countText) > 0 Then groups.Add Array(current, ReportAutomation_ArrayText(values, r, 2), bodyRow + r - 1, countText)
        End If
    Next r
End Function

Private Function ReportAutomation_IsVerticalMember(ByRef values As Variant, ByVal r As Long) As Boolean
    Dim countText As String: countText = ReportAutomation_BannerCountText(values, r, 3)
    If Len(ReportAutomation_ArrayText(values, r, 2)) = 0 Or Not IsNumeric(countText) Then Exit Function
    If CDbl(countText) < 0 Then Exit Function
    Dim c As Long
    For c = 4 To UBound(values, 2)
        If ReportAutomation_ArrayNumeric(values, r, c) Then ReportAutomation_IsVerticalMember = True: Exit Function
    Next c
End Function

Private Function ReportAutomation_BannerCountText(ByRef values As Variant, ByVal r As Long, ByVal c As Long) As String
    ReportAutomation_BannerCountText = Trim$(Replace(Replace(ReportAutomation_ArrayText(values, r, c), "(", ""), ")", ""))
End Function

Public Function ReportAutomation_BannerPercentColumns(ByVal ws As Worksheet, ByVal firstHeader As Long, _
    ByVal bodyRow As Long, ByVal endRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Collection
    Dim columns As New Collection, c As Long
    Set ReportAutomation_BannerPercentColumns = columns
    For c = firstCol To lastCol
        If ReportAutomation_IsPercentMeasureColumn(ws, firstHeader - 1, bodyRow, c) Then columns.Add c
    Next c
    If columns.Count > 0 Then Exit Function

    ' Unmarked rates need two label columns and counts that reproduce the reported rates.
    Dim countCol As Long, candidates As New Collection, values As Variant
    Dim r As Long, n As Double, rate As Double, hits As Double, verified As Long, hasFraction As Boolean
    Dim valid As Boolean, fraction As Boolean, text As String, countText As String
    countCol = ReportAutomation_BannerCountColumn(ws, firstHeader, bodyRow, firstCol, lastCol)
    If countCol < firstCol + 2 Or lastCol < countCol + 2 Or endRow <= bodyRow Then Exit Function
    values = ws.Range(ws.Cells(bodyRow, countCol), ws.Cells(endRow, lastCol)).Value2
    For c = countCol + 1 To lastCol
        valid = True: verified = 0: fraction = False
        For r = firstHeader To bodyRow - 1
            text = ReportAutomation_CellText(ws, r, c)
            If ReportAutomation_IsTotalLabel(text) Or ReportAutomation_IsMeasureHeader(text) Or InStr(text, "평균") > 0 Then valid = False: Exit For
        Next r
        If Not valid Or Not ReportAutomation_IsReportItemLabel(ReportAutomation_HeaderLabel(ws, firstHeader, bodyRow - 1, c, firstCol)) Then GoTo NextColumn
        For r = 1 To UBound(values, 1)
            countText = ReportAutomation_BannerCountText(values, r, 1)
            If IsNumeric(countText) And IsNumeric(values(r, c - countCol + 1)) _
                And Len(ReportAutomation_ArrayText(values, r, c - countCol + 1)) > 0 Then
                n = CDbl(countText)
                rate = CDbl(values(r, c - countCol + 1))
                If n <= 0 Then GoTo NextRate
                hits = n * rate / 100#
                If n <> Fix(n) Or rate < 0 Or rate > 100 Or Abs(hits - Round(hits, 0)) > 0.0001 Then valid = False
                verified = verified + 1
                If Abs(rate - Round(rate, 0)) > 0.0001 Then fraction = True
            End If
NextRate:
        Next r
        ' A total-only column can be a chart helper, not a measure across the body.
        If verified >= 2 Then
            If Not valid Then Exit Function
            candidates.Add c
            If fraction Then hasFraction = True
        End If
NextColumn:
    Next c
    If candidates.Count >= 2 And hasFraction Then Set ReportAutomation_BannerPercentColumns = candidates
End Function

Private Function ReportAutomation_BannerName(ByVal value As String) As String
    ReportAutomation_BannerName = ReportAutomation_CleanText(Replace(Replace(value, "◐", ""), "▣", ""))
End Function

Private Function ReportAutomation_IsScaleSummaryBanner(ByVal value As String) As Boolean
    Dim key As String
    key = ReportAutomation_BannerName(value)
    If Len(key) = 0 Then Exit Function
    key = Replace(Trim$(Split(key, "(")(0)), " ", "")
    Select Case key
        Case "긍정", "부정", "보통", "인지", "비인지", "만족", "불만족", "동의", "비동의", "찬성", "반대", "도움이됨", "해로움", "필요함", "필요하지않음"
            ReportAutomation_IsScaleSummaryBanner = True
        Case Else
            Static pattern As Object
            If pattern Is Nothing Then
                Set pattern = CreateObject("VBScript.RegExp")
                pattern.IgnoreCase = True
                pattern.Pattern = "^(TOP|BOT|BOTTOM|MID|LOW)[0-9]*$"
            End If
            ReportAutomation_IsScaleSummaryBanner = pattern.Test(key)
    End Select
End Function

' 프로시저 : ReportAutomation_AppendBannerChartRows
' 설  명  : 사용자가 선택한 배너 그룹의 데이터를 차트 시트에 추가 기록한다.
' 인  자  : chartRow — 참조 전달(ByRef), 기록 후 다음 빈 행으로 이동한다.
'           bannerSetting — 쉼표 구분 배너 목록 문자열 (예: "전체,성별")
' 주  의  : "전체"는 WriteNarratives에서 이미 기록하므로 이 함수에서는 건너뛴다.
' ============================================================
Public Sub ReportAutomation_AppendBannerChartRows(ByVal wsChart As Worksheet, ByRef chartRow As Long, _
                                                    ByVal ws As Worksheet, ByVal tableRec As Variant, _
                                                    ByVal bannerSetting As String, ByRef narrative As String)
    If Len(bannerSetting) = 0 Or bannerSetting = "전체" Then Exit Sub
    Dim startRow As Long, endRow As Long, totalRow As Long, firstCol As Long, bodyRow As Long
    Dim bannerGroups As Collection
    Dim bg As Variant
    Dim bgName As String
    Dim pctCols As Variant
    Dim colIdx As Long, r As Long
    Dim pctCol As Long
    Dim subLabel As String
    Dim category As String
    Dim percentValue As Variant
    Dim declaredStub As Long

    startRow = ReportAutomation_HeaderRowIndex(tableRec) - 1
    endRow   = CLng(tableRec(IDX_END_ROW))
    firstCol = ReportAutomation_FirstCol(tableRec)
    bodyRow = ReportAutomation_BodyRowIndex(tableRec)
    declaredStub = ReportAutomation_StubColumn(ws, startRow + 1, bodyRow - 1, firstCol, CLng(tableRec(IDX_LAST_COL)))
    totalRow = ReportAutomation_FindTotalRow(ws, bodyRow, endRow, firstCol, CLng(tableRec(IDX_LAST_COL)))
    Dim lastCol As Long: lastCol = CLng(tableRec(IDX_LAST_COL))
    Dim verticalRows As Collection
    Set verticalRows = ReportAutomation_FindVerticalBannerRows(ws, startRow + 1, bodyRow, endRow, firstCol, lastCol)
    Set bannerGroups = ReportAutomation_FindBannerGroups(ws, startRow + 1, bodyRow - 1, lastCol, firstCol)

    For Each bg In bannerGroups
        bgName = CStr(bg(0))

        ' "전체"는 이미 기록됐으므로 건너뛴다.
        If bgName = "전체" Then GoTo NextGroup

        ' 사용자가 선택하지 않은 배너는 건너뛴다.
        If Not ReportAutomation_BannerRequested(bgName, bannerSetting, "가로") Then GoTo NextGroup

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
                    rowCandidates.Add Array(category, CDbl(percentValue), ws.Cells(r, pctCol).Address(False, False))
                End If
NextRow:
            Next r

            ReportAutomation_WriteBannerPoints wsChart, chartRow, CStr(tableRec(IDX_TABLE_KEY)), _
                "가로: " & ReportAutomation_BannerName(bgName) & " / " & subLabel, rowCandidates, narrative, verticalRows.Count = 0
        Next colIdx
NextGroup:
    Next bg

    If verticalRows.Count = 0 Or InStr(bannerSetting, "세로:") = 0 Then Exit Sub
    Dim member As Variant, columns As Collection, c As Variant
    Set columns = ReportAutomation_BannerPercentColumns(ws, startRow + 1, bodyRow, endRow, firstCol, lastCol)
    For Each member In verticalRows
        If ReportAutomation_BannerRequested(CStr(member(0)), bannerSetting, "세로") Then
            Set rowCandidates = New Collection
            For Each c In columns
                category = ReportAutomation_HeaderLabel(ws, startRow + 1, bodyRow - 1, CLng(c), firstCol)
                If ReportAutomation_IsReportItemLabel(category) And Not ReportAutomation_IsNeutralCompositeLabel(category) Then
                    percentValue = ReportAutomation_PercentValue(ws.Cells(CLng(member(2)), CLng(c)))
                    If Not IsEmpty(percentValue) Then rowCandidates.Add Array(category, CDbl(percentValue), ws.Cells(CLng(member(2)), CLng(c)).Address(False, False))
                End If
            Next c
            ReportAutomation_WriteBannerPoints wsChart, chartRow, CStr(tableRec(IDX_TABLE_KEY)), _
                "세로: " & member(0) & " / " & member(1), rowCandidates, narrative, True, CStr(member(3))
        End If
    Next member
End Sub

Private Sub ReportAutomation_WriteBannerPoints(ByVal wsChart As Worksheet, ByRef chartRow As Long, ByVal key As String, _
    ByVal label As String, ByVal points As Collection, ByRef narrative As String, ByVal addNarrative As Boolean, _
    Optional ByVal sampleCount As String = "")
    If points.Count = 0 Then Exit Sub
    Dim rows As Variant, item As Variant, i As Long, leaders As String
    rows = ReportAutomation_CollectionToArray(points)
    ReportAutomation_SortArrayDescending rows, 1
    For i = 0 To UBound(rows)
        item = rows(i)
        wsChart.Cells(chartRow, 1).Resize(1, 9).Value = Array(key, label, item(0), "%", item(1), _
            ReportAutomation_FormatPercent(CDbl(item(1))), i + 1, IIf(i < CHART_INCLUDE_TOP_N, "Y", "N"), item(2))
        chartRow = chartRow + 1
        If Abs(CDbl(item(1)) - CDbl(rows(0)(1))) < 0.0000001 Then
            If Len(leaders) > 0 Then leaders = leaders & ", "
            leaders = leaders & ReportAutomation_Quoted(CStr(item(0)))
        End If
    Next i
    If addNarrative Then
        narrative = narrative & vbLf & "[" & label & "] " & leaders & " 응답 비율이 " & _
            ReportAutomation_FormatPercent(CDbl(rows(0)(1))) & "로 가장 높게 나타남"
        If Len(sampleCount) > 0 Then narrative = narrative & " (n=" & sampleCount & ")"
    End If
End Sub
' ============================================================
Private Function ReportAutomation_BannerRequested(ByVal bannerName As String, ByVal setting As String, ByVal axis As String) As Boolean
    Dim tokens() As String
    tokens = Split(setting, ",")
    Dim t As Long
    For t = LBound(tokens) To UBound(tokens)
        Dim choice As String, key As String
        choice = Replace(ReportAutomation_BannerName(tokens(t)), " ", "")
        key = Replace(ReportAutomation_BannerName(bannerName), " ", "")
        If choice = axis & ":" & key Or (axis = "가로" And choice = key) Then
            ReportAutomation_BannerRequested = True
            Exit Function
        End If
    Next t
End Function
