Attribute VB_Name = "ReportAutomationTableCells"
' ============================================================
' 모듈명  : ReportAutomationTableCells
' 설  명  : 원본 집계표 셀의 표시값, 실제값, 병합, 역할을 산출 시트로 내보낸다.
' ============================================================
Option Explicit

Public Sub ReportAutomation_WriteTableCells(ByVal wsOut As Worksheet, ByVal tables As Collection, _
                                             ByVal dataWs As Worksheet, ByVal wsSettings As Worksheet)
    wsOut.Range("A1:P1").Value = Array("table_key", "title", "row", "col", "rowspan", "colspan", _
        "role", "display_text", "raw_value", "number_format", "source_sheet", "source_cell", _
        "source_range", "horizontal_align", "vertical_align", "covered_by")
    ReportAutomation_StyleHeader wsOut.Range("A1:P1")
    wsOut.Columns("H:H").NumberFormat = "@"

    Dim outputRow As Long: outputRow = 2
    Dim i As Long, rec As Variant
    For i = 1 To tables.Count
        rec = tables(i)
        Dim sourceRange As Range
        Set sourceRange = ReportAutomation_TableSourceRange(dataWs, rec, wsSettings)
        ReportAutomation_WriteTableRangeCells wsOut, outputRow, sourceRange, rec
    Next i

    wsOut.Columns("A:P").AutoFit
    wsOut.Columns("H:H").ColumnWidth = 24
    wsOut.Columns("H:H").WrapText = True
    wsOut.Rows(1).AutoFilter
End Sub

Public Function ReportAutomation_TableSourceRange(ByVal dataWs As Worksheet, ByVal tableRec As Variant, _
                                                   ByVal wsSettings As Worksheet) As Range
    Dim defaultRange As Range
    ' 표 제목은 HWP writer가 별도 문단으로 입력하므로 셀 계약에는 BASE/배너/값 영역만 포함한다.
    Dim dataStartRow As Long
    dataStartRow = CLng(tableRec(IDX_START_ROW)) + 1
    If dataStartRow > CLng(tableRec(IDX_END_ROW)) Then dataStartRow = CLng(tableRec(IDX_START_ROW))
    Set defaultRange = dataWs.Range( _
        dataWs.Cells(dataStartRow, 1), _
        dataWs.Cells(CLng(tableRec(IDX_END_ROW)), CLng(tableRec(IDX_LAST_COL))))

    Dim overrideText As String
    overrideText = Trim$(CStr(ReportAutomation_SettingValue( _
        "table_range." & CStr(tableRec(IDX_TABLE_KEY)), "", wsSettings)))
    If Len(overrideText) = 0 Then
        Set ReportAutomation_TableSourceRange = defaultRange
        Exit Function
    End If

    Dim bangPos As Long
    bangPos = InStrRev(overrideText, "!")
    If bangPos = 0 Then
        Set ReportAutomation_TableSourceRange = dataWs.Range(overrideText)
        Exit Function
    End If

    Dim sheetName As String, addressText As String
    sheetName = Replace(Left$(overrideText, bangPos - 1), "'", "")
    addressText = Mid$(overrideText, bangPos + 1)
    Set ReportAutomation_TableSourceRange = dataWs.Parent.Worksheets(sheetName).Range(addressText)
End Function

Private Sub ReportAutomation_WriteTableRangeCells(ByVal wsOut As Worksheet, ByRef outputRow As Long, _
                                                   ByVal sourceRange As Range, ByVal tableRec As Variant)
    Dim sourceAddress As String
    sourceAddress = sourceRange.Address(False, False)

    Dim r As Long, c As Long
    For r = 1 To sourceRange.Rows.Count
        For c = 1 To sourceRange.Columns.Count
            Dim cell As Range
            Set cell = sourceRange.Cells(r, c)

            Dim mergeAnchor As Range, coveredBy As String
            Set mergeAnchor = Nothing
            coveredBy = ""
            Dim rowSpan As Long: rowSpan = 1
            Dim colSpan As Long: colSpan = 1
            If cell.MergeCells Then
                Set mergeAnchor = cell.MergeArea.Cells(1, 1)
                If cell.Address = mergeAnchor.Address Then
                    rowSpan = cell.MergeArea.Rows.Count
                    colSpan = cell.MergeArea.Columns.Count
                Else
                    coveredBy = mergeAnchor.Address(False, False)
                End If
            End If

            wsOut.Cells(outputRow, 1).Value = tableRec(IDX_TABLE_KEY)
            wsOut.Cells(outputRow, 2).Value = tableRec(IDX_TITLE)
            wsOut.Cells(outputRow, 3).Value = r
            wsOut.Cells(outputRow, 4).Value = c
            wsOut.Cells(outputRow, 5).Value = rowSpan
            wsOut.Cells(outputRow, 6).Value = colSpan
            wsOut.Cells(outputRow, 7).Value = ReportAutomation_CellRole(cell, r, c, sourceRange, coveredBy)
            wsOut.Cells(outputRow, 8).Value2 = CStr(ReportAutomation_CellDisplayText(cell))
            Dim rawValue As Variant
            rawValue = cell.Value2
            If VarType(rawValue) = vbString Then
                wsOut.Cells(outputRow, 9).NumberFormat = "@"
                wsOut.Cells(outputRow, 9).Value2 = CStr(rawValue)
            Else
                wsOut.Cells(outputRow, 9).Value2 = rawValue
            End If
            wsOut.Cells(outputRow, 10).Value = cell.NumberFormat
            wsOut.Cells(outputRow, 11).Value = sourceRange.Worksheet.Name
            wsOut.Cells(outputRow, 12).Value = cell.Address(False, False)
            wsOut.Cells(outputRow, 13).Value = sourceAddress
            wsOut.Cells(outputRow, 14).Value = ReportAutomation_HorizontalAlignName(cell.HorizontalAlignment)
            wsOut.Cells(outputRow, 15).Value = ReportAutomation_VerticalAlignName(cell.VerticalAlignment)
            wsOut.Cells(outputRow, 16).Value = coveredBy
            outputRow = outputRow + 1
        Next c
    Next r
End Sub

Private Function ReportAutomation_CellDisplayText(ByVal cell As Range) As String
    If IsError(cell.Value) Then
        ReportAutomation_CellDisplayText = CStr(cell.Text)
    Else
        ReportAutomation_CellDisplayText = CStr(cell.Text)
    End If
End Function

Private Function ReportAutomation_CellRole(ByVal cell As Range, ByVal relativeRow As Long, ByVal relativeCol As Long, _
                                            ByVal sourceRange As Range, ByVal coveredBy As String) As String
    If Len(coveredBy) > 0 Then
        ReportAutomation_CellRole = "blank"
        Exit Function
    End If

    Dim text As String
    text = Trim$(ReportAutomation_CellDisplayText(cell))
    If Len(text) = 0 Then
        ReportAutomation_CellRole = "blank"
    ElseIf relativeRow = 1 And relativeCol = 1 And InStr(1, text, "표", vbTextCompare) > 0 Then
        ReportAutomation_CellRole = "title"
    ElseIf InStr(1, text, "BASE", vbTextCompare) > 0 Or InStr(1, text, "사례수", vbTextCompare) > 0 Then
        ReportAutomation_CellRole = "base"
    ElseIf Left$(text, 2) = "출처" Then
        ReportAutomation_CellRole = "source"
    ElseIf Left$(text, 1) = "주" And Not IsNumeric(text) Then
        ReportAutomation_CellRole = "note"
    ElseIf IsNumeric(cell.Value2) Then
        ReportAutomation_CellRole = "value"
    ElseIf relativeRow <= 3 Then
        ReportAutomation_CellRole = "banner_horizontal"
    ElseIf relativeCol = 1 Then
        ReportAutomation_CellRole = "banner_vertical"
    ElseIf relativeCol = 2 Then
        ReportAutomation_CellRole = "stub"
    Else
        ReportAutomation_CellRole = "unknown"
    End If
End Function

Private Function ReportAutomation_HorizontalAlignName(ByVal alignmentValue As Variant) As String
    Select Case alignmentValue
        Case xlLeft: ReportAutomation_HorizontalAlignName = "left"
        Case xlRight: ReportAutomation_HorizontalAlignName = "right"
        Case xlCenter, xlCenterAcrossSelection: ReportAutomation_HorizontalAlignName = "center"
        Case Else: ReportAutomation_HorizontalAlignName = "general"
    End Select
End Function

Private Function ReportAutomation_VerticalAlignName(ByVal alignmentValue As Variant) As String
    Select Case alignmentValue
        Case xlTop: ReportAutomation_VerticalAlignName = "top"
        Case xlBottom: ReportAutomation_VerticalAlignName = "bottom"
        Case Else: ReportAutomation_VerticalAlignName = "center"
    End Select
End Function
