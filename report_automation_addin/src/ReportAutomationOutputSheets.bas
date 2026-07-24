Attribute VB_Name = "ReportAutomationOutputSheets"
' ============================================================
' 모듈명  : ReportAutomationOutputSheets
' 설  명  : 보고서 자동화 산출 시트 생성과 메타 작성
' ============================================================
Option Explicit

' ============================================================
Public Function ReportAutomation_AddOutputSheet(ByVal wb As Workbook, ByVal baseName As String) As Worksheet
    Set ReportAutomation_AddOutputSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ReportAutomation_AddOutputSheet.Name = ReportAutomation_UniqueSheetName(wb, baseName)
End Function
' ============================================================
Public Function ReportAutomation_GetSettingHint(ByVal labelText As String) As String
    Select Case labelText
        Case "추출 배너 목록":        ReportAutomation_GetSettingHint = "전체 / 전체,성별 / 전체,성별,연령대"
        Case "제목 제거 접두어":      ReportAutomation_GetSettingHint = "쉼표 구분 (예: 사업체 특성별,응답자 특성별)"
        Case "문체 프로필":           ReportAutomation_GetSettingHint = "공공조사 보고서형 / 학술보고서형"
        Case "수치 표기":             ReportAutomation_GetSettingHint = "소수점 1자리 / 정수 / 소수점 2자리"
        Case "QA 실행 여부":          ReportAutomation_GetSettingHint = "수치 QA / 문체 QA / 표 서식 QA / 사용 안 함"
        Case "LLM 문장 고도화":       ReportAutomation_GetSettingHint = "사용 안 함 / 사용"
        Case "LLM 제공자":            ReportAutomation_GetSettingHint = "사용 안 함 / OpenAI / Claude"
        Case "LLM API 키 입력 방식":  ReportAutomation_GetSettingHint = "실행 시 입력 / 환경변수 / 설정파일"
        Case "N 표기 여부":           ReportAutomation_GetSettingHint = "가중 사례수와 실제 사례수 보존 / 가중 사례수만"
        Case "비교 기준":             ReportAutomation_GetSettingHint = "전체 대비 / 배너 간 비교"
    End Select
End Function
' ============================================================
Public Sub ReportAutomation_WriteSettingsSheet(ByVal ws As Worksheet, ByVal wb As Workbook, ByVal dataWs As Worksheet, ByVal priorSettingsWs As Worksheet)
    ws.Range("A1").Value = "보고서 자동화 설정"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("C1").Value = "입력 예시"

    ' 현재 단계에서는 기본값을 먼저 깔아두고, 후속 UI에서 사용자가 수정하는 구조를 염두에 둔다.
    Dim labels As Variant, values As Variant
    labels = Array("프로젝트ID", "프로젝트명", "보고서명", "프로젝트 유형", "산출물 유형", _
                   "문체 프로필", "보고서 작성 프로필", "수치 표기", "N 표기 여부", "비교 기준", "QA 실행 여부", _
                   "출처 관리 여부", "수정 이력 관리 여부", "LLM 문장 고도화", "LLM 제공자", _
                   "LLM API 키 입력 방식", "LLM 모델명", "원본 통합문서", "원본 집계표 시트", _
                   "추출 배너 목록", "제목 제거 접두어")
    values = Array("PRJ_" & Format(Now, "yyyymmdd_hhnn"), "", "", "데이터분석", "Excel 산출 시트", _
                   "공공조사 보고서형", "인식도/만족도 조사형", "소수점 1자리", "가중 사례수와 실제 사례수 보존", _
                   "전체 대비", "수치 QA / 문체 QA / 표 서식 QA", "사용 안 함", "사용", _
                   "사용 안 함", "사용 안 함", "실행 시 입력", "", wb.Name, dataWs.Name, _
                   "전체", "")

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        ws.Cells(i + 3, 1).Value = labels(i)
        If labels(i) = "원본 통합문서" Then
            ws.Cells(i + 3, 2).Value = wb.Name
        ElseIf labels(i) = "원본 집계표 시트" Then
            ws.Cells(i + 3, 2).Value = dataWs.Name
        Else
            ws.Cells(i + 3, 2).Value = ReportAutomation_SettingValue(CStr(labels(i)), values(i), priorSettingsWs)
        End If
        Dim settingHint As String
        settingHint = ReportAutomation_GetSettingHint(CStr(labels(i)))
        If Len(settingHint) > 0 Then
            ws.Cells(i + 3, 3).Value = settingHint
            ws.Cells(i + 3, 3).Font.Color = RGB(89, 89, 89)
            ws.Cells(i + 3, 3).Font.Italic = True
        End If
    Next i

    ws.Columns("A:C").AutoFit
    ws.Range("A3:A" & UBound(labels) + 3).Font.Bold = True
    ReportAutomation_StyleHeader ws.Range("A1:C1")
End Sub
' ============================================================
Public Sub ReportAutomation_WriteTableList(ByVal ws As Worksheet, ByVal tables As Collection, ByVal dataWs As Worksheet)
    ws.Range("A1:L1").Value = Array("선택", "table_key", "table_no", "title", "basis", "base_label", _
                                    "source_sheet", "source_range", "table_type", "row_count", "col_count", "warning")
    ReportAutomation_StyleHeader ws.Range("A1:L1")
    ws.Columns(3).NumberFormat = "@"

    Dim i As Long, rec As Variant
    For i = 1 To tables.Count
        rec = tables(i)
        ' 기본값은 전체 표 사용(Y). 사용자는 이 열을 N으로 바꿔 후속 생성 대상에서 제외할 수 있다.
        ws.Cells(i + 1, 1).Value = "Y"
        ws.Cells(i + 1, 2).Value = rec(IDX_TABLE_KEY)
        ws.Cells(i + 1, 3).Value = rec(IDX_TABLE_NO)
        ws.Cells(i + 1, 4).Value = rec(IDX_TITLE)
        ws.Cells(i + 1, 5).Value = rec(IDX_BASIS)
        ws.Cells(i + 1, 6).Value = rec(IDX_BASE_LABEL)
        ws.Cells(i + 1, 7).Value = dataWs.Name
        ws.Cells(i + 1, 8).Value = dataWs.Range(dataWs.Cells(CLng(rec(IDX_START_ROW)), 1), dataWs.Cells(CLng(rec(IDX_END_ROW)), CLng(rec(IDX_LAST_COL)))).Address(False, False)
        ws.Cells(i + 1, 9).Value = rec(IDX_TABLE_TYPE)
        ws.Cells(i + 1, 10).Value = CLng(rec(IDX_END_ROW)) - CLng(rec(IDX_START_ROW)) + 1
        ws.Cells(i + 1, 11).Value = CLng(rec(IDX_LAST_COL))
        ws.Cells(i + 1, 12).Value = rec(IDX_WARNING)
    Next i

    ws.Columns("A:L").AutoFit
    ws.Rows(1).AutoFilter
End Sub
' ============================================================
Public Sub ReportAutomation_WriteSourceSheet(ByVal ws As Worksheet)
    ws.Range("A1:J1").Value = Array("source_id", "title", "author_org", "year", "url", "accessed_at", "source_type", "reliability", "used_section", "note")
    ReportAutomation_StyleHeader ws.Range("A1:J1")
    ws.Cells(2, 1).Value = "SRC_0001"
    ws.Cells(2, 2).Value = "원본 집계표"
    ws.Cells(2, 7).Value = "Excel"
    ws.Cells(2, 8).Value = "primary"
    ws.Columns("A:J").AutoFit
End Sub
' ============================================================
Public Sub ReportAutomation_WriteRevisionSheet(ByVal ws As Worksheet)
    ws.Range("A1:G1").Value = Array("revision_id", "target", "user_feedback", "before", "after", "created_at", "applied_status")
    ReportAutomation_StyleHeader ws.Range("A1:G1")
    ws.Columns("A:G").AutoFit
End Sub
' ============================================================
Public Sub ReportAutomation_WriteMetaSheet(ByVal ws As Worksheet, ByVal wb As Workbook, ByVal dataWs As Worksheet, ByVal tables As Collection, _
                                            ByVal wsSettings As Worksheet, ByVal wsList As Worksheet, ByVal wsNarr As Worksheet, _
                                            ByVal wsChart As Worksheet, ByVal wsInsert As Worksheet)
    ws.Range("A1:B1").Value = Array("key", "value")
    ReportAutomation_StyleHeader ws.Range("A1:B1")

    Dim rowNo As Long: rowNo = 2
    ReportAutomation_WriteMetaKV ws, rowNo, "generated_at", Format(Now, "yyyy-mm-dd hh:nn:ss")
    ReportAutomation_WriteMetaKV ws, rowNo, "source_workbook", wb.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "source_sheet", dataWs.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "table_count", tables.Count
    ' 산출 시트명은 timestamp가 붙으므로 메타에 정확한 이름을 남겨 후속 스크립트가 찾을 수 있게 한다.
    ReportAutomation_WriteMetaKV ws, rowNo, "settings_sheet", wsSettings.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "table_list_sheet", wsList.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "narrative_sheet", wsNarr.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "chart_data_sheet", wsChart.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "insert_table_sheet", wsInsert.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "style_profile", "formal_korean_research"
    ReportAutomation_WriteMetaKV ws, rowNo, "number_format", "percent_1_decimal"
    ReportAutomation_WriteMetaKV ws, rowNo, "hwp_chart_paste_mode", "metafile_preferred"
    ws.Columns("A:B").AutoFit
End Sub
' ============================================================
Public Sub ReportAutomation_WriteMetaKV(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal keyText As String, ByVal valueText As Variant)
    ws.Cells(rowNo, 1).Value = keyText
    ws.Cells(rowNo, 2).Value = valueText
    rowNo = rowNo + 1
End Sub
