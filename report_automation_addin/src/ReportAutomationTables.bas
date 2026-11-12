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
    Dim tables As Collection
    Set tables = ReportAutomation_CollectTables(ws)
    ReportAutomation_CountTableTitleRows = tables.Count
End Function

Public Function ReportAutomation_CollectTables(ByVal ws As Worksheet) As Collection
    Dim tables As New Collection, values As Variant, used As Range
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Set ReportAutomation_CollectTables = tables
    If LCase$(Trim$(ws.Name)) = "index" Or Trim$(ws.Name) = "목차" Then Exit Function
    If Left$(ws.Name, 4) = "보고서_" Or LCase$(Left$(ws.Name, 7)) = "_report" Then Exit Function
    Set used = ws.UsedRange
    lastRow = used.Row + used.Rows.Count - 1
    lastCol = used.Column + used.Columns.Count - 1
    If lastRow < 2 Or lastCol < 2 Then Exit Function
    ' 값 탐색은 한 번의 Value2 배열로 처리하고 병합 정보만 후보에서 읽는다.
    values = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2
    ' 내부 구분열은 유지한다. 좌우 독립 표의 의미는 자동으로 추측하지 않는다.
    ReportAutomation_CollectColumnBlock ws, values, lastRow, 1, lastCol, tables
End Function

Private Sub ReportAutomation_CollectColumnBlock(ByVal ws As Worksheet, ByRef values As Variant, _
    ByVal lastRow As Long, ByVal leftCol As Long, ByVal rightCol As Long, ByVal tables As Collection)
    Dim r As Long, h As Long, b As Long, e As Long, c As Long, titleCol As Long, probe As Long
    Dim titleText As String, tableNo As String, cleanTitle As String, warning As String
    Dim firstCol As Long, finalCol As Long, sourceRow As Long, headerRow As Long
    Dim baseLabel As String, tableType As String, rec As Variant
    Dim explicitTitle As Boolean, plainTitle As Boolean, declaredStub As Long
    Dim containsAverage As Boolean, containsN As Boolean
    r = 1
    Do While r <= lastRow
        titleText = ReportAutomation_RowText(values, r, leftCol, rightCol)
        titleCol = ReportAutomation_FirstTextCol(values, r, leftCol, rightCol)
        explicitTitle = ReportAutomation_IsTableTitle(titleText)
        plainTitle = False
        If titleCol > 0 And Not explicitTitle Then
            plainTitle = ReportAutomation_PlainTitleRow(values, r, leftCol, rightCol)
        End If
        h = 0: b = 0
        If explicitTitle Or plainTitle Then
            h = ReportAutomation_FindHeader(ws, values, r + 1, lastRow, leftCol, rightCol)
        ElseIf ReportAutomation_HeaderRow(values, r, leftCol, rightCol) Then
            h = r
        ElseIf ReportAutomation_HeaderBoundary(values, r, leftCol, rightCol) Then
            If ReportAutomation_UnknownHeaderRow(values, r, lastRow, leftCol, rightCol) Then h = r
        End If
        If h > 0 Then b = ReportAutomation_FindBody(ws, values, h, lastRow, leftCol, rightCol)
        If b = 0 Then GoTo NextCandidate
        declaredStub = ReportAutomation_DeclaredStub(values, h, b - 1, leftCol, rightCol)
        If plainTitle Then
            Dim stubCol As Long, numericCol As Long
            stubCol = ReportAutomation_FirstTextCol(values, b, leftCol, rightCol)
            numericCol = 0
            For c = leftCol To rightCol
                If ReportAutomation_ArrayNumeric(values, b, c) Then numericCol = c: Exit For
            Next c
            ' 문항은 스텁을 포함하고, 측정영역의 부분 병합 부모 배너는 제외한다.
            If ws.Cells(r, titleCol).MergeCells Then
                If titleCol > stubCol Then GoTo NextCandidate
                If ws.Cells(r, titleCol).MergeArea.Column + ws.Cells(r, titleCol).MergeArea.Columns.Count - 1 < numericCol Then GoTo NextCandidate
            ElseIf titleCol > Application.Max(stubCol, 3) Then
                GoTo NextCandidate
            End If
            If Not ReportAutomation_HasCountHeader(values, h, b - 1, leftCol, rightCol) Then
                If b = lastRow Then GoTo NextCandidate
                probe = b + 1
                Do While probe <= lastRow
                    If Len(ReportAutomation_RowText(values, probe, leftCol, rightCol)) > 0 Then Exit Do
                    probe = probe + 1
                Loop
                If probe > lastRow Then GoTo NextCandidate
                If Not ReportAutomation_NumericBodyRow(values, probe, leftCol, rightCol, declaredStub) Then GoTo NextCandidate
            End If
        End If
        sourceRow = h: headerRow = h
        If explicitTitle Or plainTitle Then
            sourceRow = r + 1
            For probe = r + 1 To h
                For c = leftCol To rightCol
                    Dim headerText As String
                    headerText = ReportAutomation_ArrayText(values, probe, c)
                    If Len(headerText) > 0 And Not IsNumeric(headerText) Then
                        If Not ReportAutomation_IsMetadata(headerText) Then
                            headerRow = probe
                            Exit For
                        End If
                    End If
                Next c
                If headerRow = probe Then Exit For
            Next probe
        Else
            ' 제목 없는 표는 상위 병합 헤더만 되짚고 이전 본문은 넘지 않는다.
            For probe = h - 1 To 1 Step -1
                If Len(ReportAutomation_RowText(values, probe, leftCol, rightCol)) = 0 Then
                    sourceRow = probe
                ElseIf ReportAutomation_IsMetadata(ReportAutomation_RowText(values, probe, leftCol, rightCol)) Then
                    Exit For
                ElseIf ReportAutomation_IsMergedHeaderRow(ws, values, probe, leftCol, rightCol) Then
                    headerRow = probe: sourceRow = probe
                Else
                    Exit For
                End If
            Next probe
            Do While sourceRow < headerRow And Len(ReportAutomation_RowText(values, sourceRow, leftCol, rightCol)) = 0
                sourceRow = sourceRow + 1
            Loop
        End If
        Do While sourceRow < b
            If Len(ReportAutomation_RowText(values, sourceRow, leftCol, rightCol)) > 0 Then Exit Do
            sourceRow = sourceRow + 1
        Loop
        e = b
        For probe = b + 1 To lastRow
            If ReportAutomation_NumericBodyRow(values, probe, leftCol, rightCol, declaredStub) Then
                e = probe
            ElseIf Len(ReportAutomation_RowText(values, probe, leftCol, rightCol)) = 0 Then
                Dim nextBody As Long
                nextBody = probe + 1
                Do While nextBody <= lastRow
                    If Len(ReportAutomation_RowText(values, nextBody, leftCol, rightCol)) > 0 Then Exit Do
                    nextBody = nextBody + 1
                Loop
                If nextBody > lastRow Then Exit For
                If Not ReportAutomation_NumericBodyRow(values, nextBody, leftCol, rightCol, declaredStub) Then Exit For
            ElseIf ReportAutomation_IsBodyBanner(ws, values, probe, lastRow, leftCol, rightCol) Then
                ' 배너 뒤의 실제 수치가 확인된 경우에만 본문을 계속한다.
            Else
                Exit For
            End If
        Next probe
        firstCol = rightCol: finalCol = leftCol
        containsAverage = False: containsN = False
        For probe = headerRow To e
            For c = leftCol To rightCol
                If Len(ReportAutomation_ArrayText(values, probe, c)) > 0 Then
                    If InStr(ReportAutomation_ArrayText(values, probe, c), "▶평균◀") > 0 Then containsAverage = True
                    If probe < b And UCase$(ReportAutomation_ArrayText(values, probe, c)) = "N" Then containsN = True
                    If c < firstCol Then firstCol = c
                    If c > finalCol Then finalCol = c
                End If
            Next c
        Next probe
        ' 병합 끝과 BASE 앵커를 포함하되 임의의 빈 여백열은 추가하지 않는다.
        For probe = sourceRow To b - 1
            For c = leftCol To rightCol
                If Len(ReportAutomation_ArrayText(values, probe, c)) > 0 Then
                    Dim anchor As Range
                    Set anchor = ws.Cells(probe, c)
                    If c < firstCol Then firstCol = c
                    If c > finalCol Then finalCol = c
                    If anchor.MergeCells Then
                        If anchor.MergeArea.Column < firstCol Then firstCol = anchor.MergeArea.Column
                        If anchor.MergeArea.Column + anchor.MergeArea.Columns.Count - 1 > finalCol Then _
                            finalCol = anchor.MergeArea.Column + anchor.MergeArea.Columns.Count - 1
                        If anchor.MergeArea.Row < sourceRow Then sourceRow = anchor.MergeArea.Row
                    End If
                End If
            Next c
        Next probe
        tableNo = "": cleanTitle = "": warning = "": baseLabel = ""
        If explicitTitle Then
            ReportAutomation_ParseTitle titleText, tableNo, cleanTitle
        ElseIf plainTitle Then
            cleanTitle = titleText
        End If
        If Len(cleanTitle) = 0 Then
            cleanTitle = ws.Name & " " & ws.Cells(headerRow, firstCol).Address(False, False)
            warning = "제목 미확인: 위치 기반 임시 제목, 의미 확인 필요"
        End If
        For probe = sourceRow To b - 1
            If InStr(1, ReportAutomation_RowText(values, probe, leftCol, rightCol), "BASE", vbTextCompare) > 0 Then _
                baseLabel = ReportAutomation_RowText(values, probe, leftCol, rightCol)
        Next probe
        If Len(baseLabel) = 0 Then warning = ReportAutomation_AppendWarning(warning, "BASE 행 확인 필요")
        Dim markerCount As Long, stubCount As Long
        markerCount = 0: stubCount = 0
        For c = leftCol To rightCol
            If ReportAutomation_IsTableTitle(ReportAutomation_ArrayText(values, r, c)) Then markerCount = markerCount + 1
            If ReportAutomation_IsStubHeader(ReportAutomation_ArrayText(values, h, c)) Then stubCount = stubCount + 1
        Next c
        If markerCount > 1 Or stubCount > 1 Then warning = ReportAutomation_AppendWarning(warning, "좌우 표 경계 모호: 수동 범위 지정 필요")
        If Not ReportAutomation_HeaderRow(values, h, leftCol, rightCol) Then _
            warning = ReportAutomation_AppendWarning(warning, "헤더 의미 미확인: 숫자 열 역할 수동 확인 필요")
        tableType = "배너 교차표"
        If finalCol - firstCol + 1 <= SIMPLE_TABLE_MAX_COL Then
            tableType = "단일 전체 분포표"
        ElseIf containsAverage Then
            tableType = "평균 포함 표"
        ElseIf containsN Then
            tableType = "N/% 쌍 표"
        End If
        If Not ReportAutomation_HasMeasureHeader(values, h, b - 1, leftCol, rightCol) Then _
            warning = ReportAutomation_AppendWarning(warning, "측정 헤더 미확인: 분석문 자동 추출 불가")
        rec = Array("T" & Format$(tables.Count + 1, "0000"), tableNo, cleanTitle, _
            ReportAutomation_ParseBasis(titleText), baseLabel, r, e, finalCol, tableType, warning, _
            firstCol, headerRow, b, sourceRow)
        If Not explicitTitle And Not plainTitle Then rec(IDX_START_ROW) = sourceRow
        tables.Add rec
        r = e
NextCandidate:
        r = r + 1
    Loop
End Sub

Public Function ReportAutomation_ArrayText(ByRef values As Variant, ByVal r As Long, ByVal c As Long) As String
    If IsError(values(r, c)) Or IsEmpty(values(r, c)) Then Exit Function
    ReportAutomation_ArrayText = ReportAutomation_CleanText(CStr(values(r, c)))
End Function

Private Function ReportAutomation_RowText(ByRef values As Variant, ByVal r As Long, ByVal firstCol As Long, ByVal lastCol As Long) As String
    Dim c As Long, text As String
    For c = firstCol To lastCol
        text = ReportAutomation_ArrayText(values, r, c)
        If Len(text) > 0 Then ReportAutomation_RowText = Trim$(ReportAutomation_RowText & " " & text)
    Next c
End Function

Private Function ReportAutomation_FirstTextCol(ByRef values As Variant, ByVal r As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = firstCol To lastCol
        If Len(ReportAutomation_ArrayText(values, r, c)) > 0 Then ReportAutomation_FirstTextCol = c: Exit Function
    Next c
End Function

Public Function ReportAutomation_IsTableTitle(ByVal valueData As Variant) As Boolean
    If IsError(valueData) Or IsEmpty(valueData) Then Exit Function
    Dim text As String
    text = LCase$(ReportAutomation_TitleText(CStr(valueData)))
    If Left$(text, 1) = "[" Or Left$(text, 1) = "<" Then text = Trim$(Mid$(text, 2))
    If Left$(text, 1) <> "표" And Left$(text, 5) <> "table" Then Exit Function
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Pattern = "^\s*[\[<]?\s*(표\s*([A-Za-z0-9]|[>\]])|Table\s+[A-Za-z0-9])"
    End If
    ReportAutomation_IsTableTitle = re.Test(ReportAutomation_TitleText(CStr(valueData)))
End Function

Private Function ReportAutomation_TitleText(ByVal text As String) As String
    text = Replace(Replace(text, "〈", "<"), "〉", ">")
    text = Replace(Replace(text, ChrW(&H3000), " "), ChrW(&H202F), " ")
    ReportAutomation_TitleText = ReportAutomation_CleanText(text)
End Function

Public Function ReportAutomation_IsMetadata(ByVal text As String) As Boolean
    text = LCase$(ReportAutomation_CleanText(text))
    ReportAutomation_IsMetadata = (Left$(text, 4) = "base" Or Left$(text, 2) = "출처" _
        Or Left$(text, 2) = "자료" Or Left$(text, 2) = "주:" Or Left$(text, 2) = "주)" _
        Or Left$(text, 2) = "주 " Or Left$(text, 4) = "note" Or Left$(text, 6) = "source" _
        Or Left$(text, 2) = "단위" Or InStr(text, "목차") > 0 Or InStr(text, "index") > 0)
End Function

Private Function ReportAutomation_PlainTitleRow(ByRef values As Variant, ByVal r As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long, count As Long, text As String
    For c = firstCol To lastCol
        text = ReportAutomation_ArrayText(values, r, c)
        If Len(text) > 0 Then
            count = count + 1
            If IsNumeric(text) Or ReportAutomation_IsMetadata(text) Or ReportAutomation_IsTotalLabel(text) _
                Or ReportAutomation_IsMeasureHeader(text) Or ReportAutomation_IsCountHeader(text) _
                Or ReportAutomation_IsStubHeader(text) Or ReportAutomation_IsBannerLabel(text) Then Exit Function
            If Left$(text, 1) = "▣" Or Left$(text, 1) = "■" Then Exit Function
        End If
    Next c
    ReportAutomation_PlainTitleRow = (count = 1)
End Function

Private Function ReportAutomation_HeaderRow(ByRef values As Variant, ByVal r As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long, text As String
    For c = firstCol To lastCol
        text = ReportAutomation_ArrayText(values, r, c)
        If ReportAutomation_IsMeasureHeader(text) Or ReportAutomation_IsCountHeader(text) Then
            ReportAutomation_HeaderRow = True
            Exit Function
        End If
    Next c
End Function

Private Function ReportAutomation_FindHeader(ByVal ws As Worksheet, ByRef values As Variant, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim r As Long, text As String, meaningful As Long
    For r = firstRow To lastRow
        text = ReportAutomation_RowText(values, r, firstCol, lastCol)
        If Len(text) = 0 Then GoTo NextHeader
        meaningful = meaningful + 1
        If meaningful > 32 Then Exit Function
        If ReportAutomation_IsTableTitle(text) Then Exit Function
        If ReportAutomation_HeaderRow(values, r, firstCol, lastCol) Then ReportAutomation_FindHeader = r: Exit Function
        If ReportAutomation_UnknownHeaderRow(values, r, lastRow, firstCol, lastCol) Then ReportAutomation_FindHeader = r: Exit Function
        If ReportAutomation_IsMetadata(text) Then
            If InStr(1, text, "INDEX", vbTextCompare) > 0 Or InStr(text, "목차") > 0 _
                Or Left$(text, 2) = "출처" Or Left$(text, 2) = "자료" Then Exit Function
        ElseIf Len(text) > 0 Then
            If ReportAutomation_PlainTitleRow(values, r, firstCol, lastCol) Then
                If Not ReportAutomation_IsMergedHeaderRow(ws, values, r, firstCol, lastCol) Then Exit Function
                Dim candidateCol As Long
                candidateCol = ReportAutomation_FirstTextCol(values, r, firstCol, lastCol)
                If candidateCol <= firstCol + 2 Then
                    If ws.Cells(r, candidateCol).MergeArea.Column + ws.Cells(r, candidateCol).MergeArea.Columns.Count - 1 >= lastCol Then Exit Function
                End If
            End If
        End If
NextHeader:
    Next r
End Function

Private Function ReportAutomation_FindBody(ByVal ws As Worksheet, ByRef values As Variant, ByVal headerRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim r As Long, c As Long, k As Long, text As String, meaningful As Long
    For r = headerRow + 1 To lastRow
        text = ReportAutomation_RowText(values, r, firstCol, lastCol)
        If Len(text) = 0 Then GoTo NextBody
        meaningful = meaningful + 1
        If meaningful > 32 Then Exit Function
        If ReportAutomation_IsTableTitle(text) Then Exit Function
        If ReportAutomation_IsMetadata(text) Then
            If Left$(LCase$(text), 4) <> "base" And Left$(text, 2) <> "단위" Then Exit Function
        End If
        If ReportAutomation_PlainTitleRow(values, r, firstCol, lastCol) Then
            c = ReportAutomation_FirstTextCol(values, r, firstCol, lastCol)
            If c <= Application.Max(3, ReportAutomation_FirstTextCol(values, headerRow, firstCol, lastCol)) Then
                If Not ReportAutomation_IsMergedHeaderRow(ws, values, r, firstCol, lastCol) Then Exit Function
                If ws.Cells(r, c).MergeArea.Column + ws.Cells(r, c).MergeArea.Columns.Count - 1 >= lastCol Then Exit Function
            End If
        End If
        If ReportAutomation_NumericBodyRow(values, r, firstCol, lastCol, ReportAutomation_DeclaredStub(values, headerRow, r - 1, firstCol, lastCol)) Then
            For c = firstCol To lastCol
                If ReportAutomation_ArrayNumeric(values, r, c) Then
                    For k = headerRow To r - 1
                        If ReportAutomation_IsMeasureHeader(ReportAutomation_ArrayText(values, k, c)) _
                            Or ReportAutomation_IsCountHeader(ReportAutomation_ArrayText(values, k, c)) Then
                            ReportAutomation_FindBody = r
                            Exit Function
                        End If
                    Next k
                End If
            Next c
            If ReportAutomation_UnknownHeaderRow(values, headerRow, lastRow, firstCol, lastCol) Then
                ReportAutomation_FindBody = r
            End If
            Exit Function
        End If
NextBody:
    Next r
End Function

' 알 수 없는 헤더는 경계와 연속된 두 수치 행으로만 검증하고 열 의미는 추측하지 않는다.
Private Function ReportAutomation_HeaderBoundary(ByRef values As Variant, ByVal r As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    If r = 1 Then ReportAutomation_HeaderBoundary = True: Exit Function
    Dim text As String
    text = ReportAutomation_RowText(values, r - 1, firstCol, lastCol)
    ReportAutomation_HeaderBoundary = (Len(text) = 0 Or ReportAutomation_IsMetadata(text))
End Function

Private Function ReportAutomation_UnknownHeaderRow(ByRef values As Variant, ByVal r As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long, text As String, count As Long, b As Long, secondRow As Long
    For c = firstCol To lastCol
        text = ReportAutomation_ArrayText(values, r, c)
        If Len(text) > 0 Then
            If IsNumeric(text) Or ReportAutomation_IsMetadata(text) Or ReportAutomation_IsTableTitle(text) Then Exit Function
            count = count + 1
        End If
    Next c
    If count < 2 Then Exit Function
    b = r + 1
    Do While b <= lastRow
        text = ReportAutomation_RowText(values, b, firstCol, lastCol)
        If Len(text) > 0 Then
            If Not ReportAutomation_IsMetadata(text) Then Exit Do
            If Left$(LCase$(text), 4) <> "base" And Left$(text, 2) <> "단위" Then Exit Function
        End If
        b = b + 1
    Loop
    If b >= lastRow Then Exit Function
    If Not ReportAutomation_NumericBodyRow(values, b, firstCol, lastCol) Then Exit Function
    secondRow = b + 1
    Do While secondRow <= lastRow
        If Len(ReportAutomation_RowText(values, secondRow, firstCol, lastCol)) > 0 Then Exit Do
        secondRow = secondRow + 1
    Loop
    If secondRow > lastRow Then Exit Function
    ReportAutomation_UnknownHeaderRow = ReportAutomation_NumericBodyRow(values, secondRow, firstCol, lastCol)
End Function

Private Function ReportAutomation_HasCountHeader(ByRef values As Variant, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Dim r As Long, c As Long
    For r = firstRow To lastRow
        For c = firstCol To lastCol
            If ReportAutomation_IsCountHeader(ReportAutomation_ArrayText(values, r, c)) Then ReportAutomation_HasCountHeader = True: Exit Function
        Next c
    Next r
End Function

Private Function ReportAutomation_HasMeasureHeader(ByRef values As Variant, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Dim r As Long, c As Long
    For r = firstRow To lastRow
        For c = firstCol To lastCol
            If ReportAutomation_IsMeasureHeader(ReportAutomation_ArrayText(values, r, c)) Then ReportAutomation_HasMeasureHeader = True: Exit Function
        Next c
    Next r
End Function

Public Function ReportAutomation_ArrayNumeric(ByRef values As Variant, ByVal r As Long, ByVal c As Long) As Boolean
    Dim text As String
    text = ReportAutomation_ArrayText(values, r, c)
    If Len(text) > 0 Then ReportAutomation_ArrayNumeric = ReportAutomation_IsNumericText(text)
End Function

Private Function ReportAutomation_NumericBodyRow(ByRef values As Variant, ByVal r As Long, ByVal firstCol As Long, ByVal lastCol As Long, Optional ByVal declaredStub As Long = 0) As Boolean
    Dim c As Long, labelFound As Boolean, numericFound As Boolean, text As String
    If declaredStub > 0 Then labelFound = (Len(ReportAutomation_ArrayText(values, r, declaredStub)) > 0)
    For c = firstCol To lastCol
        text = ReportAutomation_ArrayText(values, r, c)
        If Len(text) > 0 Then
            If ReportAutomation_IsMetadata(text) Or ReportAutomation_IsTableTitle(text) Then Exit Function
            If ReportAutomation_ArrayNumeric(values, r, c) Then
                If labelFound And c <> declaredStub Then numericFound = True
            Else
                If ReportAutomation_IsMeasureHeader(text) Or ReportAutomation_IsCountHeader(text) Then Exit Function
                labelFound = True
            End If
        End If
    Next c
    ReportAutomation_NumericBodyRow = numericFound
End Function

' 수치 스텁은 N/측정 열보다 앞의 명시된 헤더 열에서만 허용한다.
Public Function ReportAutomation_DeclaredStub(ByRef values As Variant, ByVal firstHeader As Long, ByVal lastHeader As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim c As Long, r As Long, measureCol As Long, text As String
    For c = firstCol To lastCol
        For r = firstHeader To lastHeader
            text = ReportAutomation_ArrayText(values, r, c)
            If ReportAutomation_IsCountHeader(text) Or ReportAutomation_IsMeasureHeader(text) Then measureCol = c: Exit For
        Next r
        If measureCol > 0 Then Exit For
    Next c
    If measureCol = 0 Then Exit Function
    For c = firstCol To measureCol - 1
        For r = firstHeader To lastHeader
            text = ReportAutomation_ArrayText(values, r, c)
            If Len(text) > 0 And Not IsNumeric(text) Then
                If Not ReportAutomation_IsMetadata(text) And Not ReportAutomation_IsTableTitle(text) Then
                    ReportAutomation_DeclaredStub = c
                    Exit Function
                End If
            End If
        Next r
    Next c
End Function

Private Function ReportAutomation_IsMergedHeaderRow(ByVal ws As Worksheet, ByRef values As Variant, ByVal r As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long, text As String
    For c = firstCol To lastCol
        text = ReportAutomation_ArrayText(values, r, c)
        If Len(text) > 0 Then
            If IsNumeric(text) Or ReportAutomation_IsMetadata(text) Or ReportAutomation_IsTableTitle(text) Then Exit Function
            If ws.Cells(r, c).MergeCells Then ReportAutomation_IsMergedHeaderRow = True
        End If
    Next c
End Function

Private Function ReportAutomation_IsBodyBanner(ByVal ws As Worksheet, ByRef values As Variant, ByVal r As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long, text As String, count As Long, labelCol As Long
    If r >= lastRow Then Exit Function
    For c = firstCol To lastCol
        text = ReportAutomation_ArrayText(values, r, c)
        If Len(text) > 0 Then
            count = count + 1: labelCol = c
            If IsNumeric(text) Or ReportAutomation_IsMetadata(text) Or ReportAutomation_IsTableTitle(text) _
                Or ReportAutomation_IsCountHeader(text) Or ReportAutomation_IsMeasureHeader(text) _
                Or ReportAutomation_IsTotalLabel(text) Then Exit Function
        End If
    Next c
    If count <> 1 Then Exit Function
    If Not ReportAutomation_NumericBodyRow(values, r + 1, firstCol, lastCol) Then Exit Function
    text = ReportAutomation_ArrayText(values, r, labelCol)
    If Left$(text, 1) = "▣" Then ReportAutomation_IsBodyBanner = True: Exit Function
    If ws.Cells(r, labelCol).MergeCells Then
        Dim numericCol As Long
        For c = firstCol To lastCol
            If ReportAutomation_ArrayNumeric(values, r + 1, c) Then numericCol = c: Exit For
        Next c
        If numericCol > 0 Then ReportAutomation_IsBodyBanner = _
            (ws.Cells(r, labelCol).MergeArea.Column + ws.Cells(r, labelCol).MergeArea.Columns.Count <= numericCol)
    End If
End Function

Private Sub ReportAutomation_ParseTitle(ByVal text As String, ByRef tableNo As String, ByRef title As String)
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    text = ReportAutomation_TitleText(text)
    re.Pattern = "^\s*[\[<]?\s*(표|Table)\s*([A-Za-z0-9]+([-_.][A-Za-z0-9]+)*)?\s*[>\]]?\s*"
    Set matches = re.Execute(text)
    If matches.Count = 0 Then title = text: Exit Sub
    tableNo = matches(0).SubMatches(1)
    title = Trim$(Mid$(text, matches(0).Length + 1))
    If Left$(title, 1) = "." Or Left$(title, 1) = ":" Then title = Trim$(Mid$(title, 2))
    If Right$(title, 1) = "]" Then title = Trim$(Left$(title, Len(title) - 1))
    If Left$(title, 1) = "[" Then
        Dim closePos As Long
        closePos = InStr(title, "]")
        If closePos > 0 Then title = Trim$(Mid$(title, closePos + 1))
    End If
    Dim divider As Long
    divider = InStr(title, "─")
    If divider > 0 Then title = Trim$(Left$(title, divider - 1))
    divider = InStr(title, "[ 문")
    If divider > 0 Then title = Trim$(Left$(title, divider - 1))
End Sub

Private Function ReportAutomation_ParseBasis(ByVal text As String) As String
    Dim openPos As Long, closePos As Long
    openPos = InStr(text, "> [")
    If openPos > 0 Then
        closePos = InStr(openPos + 3, text, "]")
        If closePos > openPos Then ReportAutomation_ParseBasis = Mid$(text, openPos + 2, closePos - openPos - 1)
    End If
End Function

Private Function ReportAutomation_AppendWarning(ByVal prior As String, ByVal text As String) As String
    If Len(prior) > 0 Then prior = prior & "; "
    ReportAutomation_AppendWarning = prior & text
End Function

' 기존 10필드 외부 레코드는 종전 좌표로 처리한다.
Public Function ReportAutomation_FirstCol(ByVal rec As Variant) As Long
    ReportAutomation_FirstCol = 1
    If UBound(rec) >= IDX_FIRST_COL Then ReportAutomation_FirstCol = CLng(rec(IDX_FIRST_COL))
End Function

Public Function ReportAutomation_HeaderRowIndex(ByVal rec As Variant) As Long
    ReportAutomation_HeaderRowIndex = CLng(rec(IDX_START_ROW)) + 1
    If UBound(rec) >= IDX_HEADER_ROW Then ReportAutomation_HeaderRowIndex = CLng(rec(IDX_HEADER_ROW))
End Function

Public Function ReportAutomation_BodyRowIndex(ByVal rec As Variant) As Long
    ReportAutomation_BodyRowIndex = CLng(rec(IDX_START_ROW)) + 3
    If UBound(rec) >= IDX_BODY_ROW Then ReportAutomation_BodyRowIndex = CLng(rec(IDX_BODY_ROW))
End Function

Public Function ReportAutomation_SourceRowIndex(ByVal rec As Variant) As Long
    ReportAutomation_SourceRowIndex = CLng(rec(IDX_START_ROW)) + 1
    If UBound(rec) >= IDX_SOURCE_ROW Then ReportAutomation_SourceRowIndex = CLng(rec(IDX_SOURCE_ROW))
End Function

' COM 회귀 시험용: 원본 통합문서를 변경하지 않고 레코드 배열을 반환한다.
Public Function ReportAutomation_GetTables(ByVal workbookName As String, ByVal sheetName As String, Optional ByVal settingsName As String = "") As Variant
    Dim ws As Worksheet, tables As Collection, rows() As Variant, rec As Variant, i As Long, j As Long
    Dim settings As Worksheet, points As Collection, groups As Collection, group As Variant, names As String
    Set ws = Application.Workbooks(workbookName).Worksheets(sheetName)
    If Len(settingsName) > 0 Then Set settings = ws.Parent.Worksheets(settingsName)
    Set tables = ReportAutomation_CollectTables(ws)
    If tables.Count = 0 Then ReportAutomation_GetTables = Empty: Exit Function
    ReDim rows(1 To tables.Count, 1 To 20)
    For i = 1 To tables.Count
        rec = tables(i)
        For j = 0 To 13
            rows(i, j + 1) = rec(j)
        Next j
        rows(i, 15) = ReportAutomation_TableSourceRange(ws, rec, settings).Address(False, False)
        Set points = ReportAutomation_ExtractKeyPoints(ws, rec)
        rows(i, 16) = points.Count
        rows(i, 17) = ReportAutomation_FormatPointSummary(points)
        Set groups = ReportAutomation_FindBannerGroups(ws, ReportAutomation_HeaderRowIndex(rec), _
            ReportAutomation_BodyRowIndex(rec) - 1, CLng(rec(IDX_LAST_COL)), ReportAutomation_FirstCol(rec))
        rows(i, 18) = groups.Count
        names = ""
        For Each group In groups
            If Len(names) > 0 Then names = names & "; "
            names = names & CStr(group(0))
        Next group
        rows(i, 19) = names
        rows(i, 20) = ReportAutomation_FindTotalRow(ws, ReportAutomation_BodyRowIndex(rec), _
            CLng(rec(IDX_END_ROW)), ReportAutomation_FirstCol(rec), CLng(rec(IDX_LAST_COL)))
    Next i
    ReportAutomation_GetTables = rows
End Function
