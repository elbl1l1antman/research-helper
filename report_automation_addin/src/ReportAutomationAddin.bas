Attribute VB_Name = "ReportAutomationAddin"
' ============================================================
' 모듈명  : ReportAutomationAddin
' 설  명  : 집계표 기반 보고서 산출 시트 생성 전용 추가기능
' ============================================================
Option Explicit

' 추가기능 버전. 산출 메타 시트와 로그 시트에 함께 기록해 결과 파일의 생성 기준을 추적한다.
Public Const REPORT_AUTOMATION_VERSION As String = "0.1.0"

' UserForm에서 전달한 1회성 실행 옵션. 설정 시트를 생성한 뒤 이 값으로 덮어쓴다.
Private mHasOptionOverrides As Boolean
Private mOverrideBannerSetting As String
Private mOverrideTitlePrefixes As String

' ============================================================
' 프로시저 : ReportAutomation_About
' 설  명  : 추가기능 버전과 현재 제공 기능을 사용자에게 안내한다.
' 호출처  : 리본 About 버튼 또는 매크로 목록
' ============================================================
Public Sub ReportAutomation_About()
    MsgBox "보고서 자동화 추가기능" & vbCrLf & _
           "버전: " & REPORT_AUTOMATION_VERSION & vbCrLf & vbCrLf & _
           "기능:" & vbCrLf & _
           "- 집계표 제목([표 ...]) 자동 탐지" & vbCrLf & _
           "- 보고서용 표 목록, 분석문, 차트데이터, 삽입표 생성" & vbCrLf & _
           "- QA/출처/수정이력/메타 시트 생성", _
           vbInformation, "보고서 자동화"
End Sub

' ============================================================
' 프로시저 : ReportAutomation_GenerateExcelOutputs
' 설  명  : 사용자가 직접 실행하는 공개 매크로.
'           활성 통합문서의 집계표를 읽어 보고서 산출 시트를 생성한다.
' 주  의  : 추가기능 파일이 아니라 데이터 통합문서가 활성화되어 있어야 한다.
' ============================================================
Public Sub ReportAutomation_GenerateExcelOutputs()
    ReportAutomation_GenerateExcelOutputsCore False
End Sub

' ============================================================
' 프로시저 : ReportAutomation_ShowOptions
' 설  명  : 사용자 설정 폼을 열어 옵션을 받은 뒤 보고서 산출을 실행한다.
' 주  의  : 폼이 빌드에 포함되지 않은 환경에서는 기본 실행으로 fallback한다.
' ============================================================
Public Sub ReportAutomation_ShowOptions()
    On Error GoTo Fallback
    ReportAutomationOptionsForm.Show
    Exit Sub

Fallback:
    Dim frmErrNum As Long, frmErrDesc As String
    frmErrNum = Err.Number
    frmErrDesc = Err.Description
    On Error GoTo 0
    If frmErrNum = 18 Then Exit Sub
    If frmErrNum <> 0 And frmErrNum <> 424 And frmErrNum <> 91 Then
        MsgBox "옵션 폼을 여는 중 오류가 발생했습니다(" & frmErrNum & "): " & frmErrDesc & vbCrLf & _
               "기본 실행으로 진행합니다.", vbExclamation, "보고서 자동화"
    End If
    ReportAutomation_GenerateExcelOutputs
End Sub

' ============================================================
' 함수명 : ReportAutomation_RunWithOptions
' 설  명 : UserForm에서 받은 옵션을 1회성 override로 저장하고 산출 생성을 실행한다.
' 반환값 : 실행 요청을 정상 접수했으면 True
' ============================================================
Public Function ReportAutomation_RunWithOptions(ByVal bannerSetting As String, ByVal titlePrefixes As String) As Boolean
    mHasOptionOverrides = True
    mOverrideBannerSetting = Trim$(bannerSetting)
    mOverrideTitlePrefixes = Trim$(titlePrefixes)

    If Len(mOverrideBannerSetting) = 0 Then mOverrideBannerSetting = "전체"
    If Len(mOverrideTitlePrefixes) = 0 Then mOverrideTitlePrefixes = ""

    ReportAutomation_GenerateExcelOutputsCore False
    ReportAutomation_RunWithOptions = True
End Function

' ============================================================
' 함수명 : ReportAutomation_RunWithOptionsSilent
' 설  명 : 외부 런처/테스트 자동화에서 메시지 박스 없이 산출 생성을 실행한다.
' 반환값 : 실행 요청을 정상 접수했으면 True
' ============================================================
Public Function ReportAutomation_RunWithOptionsSilent(ByVal bannerSetting As String, ByVal titlePrefixes As String) As Boolean
    mHasOptionOverrides = True
    mOverrideBannerSetting = Trim$(bannerSetting)
    mOverrideTitlePrefixes = Trim$(titlePrefixes)

    If Len(mOverrideBannerSetting) = 0 Then mOverrideBannerSetting = "전체"
    If Len(mOverrideTitlePrefixes) = 0 Then mOverrideTitlePrefixes = ""

    ReportAutomation_GenerateExcelOutputsCore True
    ReportAutomation_RunWithOptionsSilent = True
End Function

' ============================================================
' 함수명 : ReportAutomation_DefaultBannerSetting
' 설  명 : UserForm 초기값으로 사용할 배너 설정을 반환한다.
' ============================================================
Public Function ReportAutomation_DefaultBannerSetting() As String
    Dim ws As Worksheet
    If Not ActiveWorkbook Is Nothing And Not ActiveWorkbook Is ThisWorkbook Then
        Set ws = ReportAutomation_FindLatestOutputSheet(ActiveWorkbook, "보고서_설정")
    End If
    ReportAutomation_DefaultBannerSetting = CStr(ReportAutomation_SettingValue("추출 배너 목록", "전체", ws))
End Function

' ============================================================
' 함수명 : ReportAutomation_DefaultTitlePrefixes
' 설  명 : UserForm 초기값으로 사용할 제목 제거 접두어 목록을 반환한다.
' ============================================================
Public Function ReportAutomation_DefaultTitlePrefixes() As String
    Dim ws As Worksheet
    If Not ActiveWorkbook Is Nothing And Not ActiveWorkbook Is ThisWorkbook Then
        Set ws = ReportAutomation_FindLatestOutputSheet(ActiveWorkbook, "보고서_설정")
    End If
    ReportAutomation_DefaultTitlePrefixes = CStr(ReportAutomation_SettingValue("제목 제거 접두어", "", ws))
End Function

' ============================================================
' 프로시저 : ReportAutomation_GenerateExcelOutputsSilent
' 설  명  : 자동 테스트/외부 호출용 무인 실행 매크로.
'           MsgBox 대신 Err.Raise로 실패 사유를 호출자에게 전달한다.
' ============================================================
Public Sub ReportAutomation_GenerateExcelOutputsSilent()
    ReportAutomation_GenerateExcelOutputsCore True
End Sub

' ============================================================
' 프로시저 : Ribbon_ReportAutomationGenerate
' 설  명  : 리본 버튼에서 산출 시트 생성을 호출하는 콜백.
' ============================================================
Public Sub Ribbon_ReportAutomationGenerate(control As IRibbonControl)
    ReportAutomation_ShowOptions
End Sub

' ============================================================
' 프로시저 : Ribbon_ReportAutomationAbout
' 설  명  : 리본 버튼에서 추가기능 정보를 표시하는 콜백.
' ============================================================
Public Sub Ribbon_ReportAutomationAbout(control As IRibbonControl)
    ReportAutomation_About
End Sub

' ============================================================
' 프로시저 : ReportAutomation_GenerateExcelOutputsCore
' 설  명  : 보고서 자동화의 메인 실행 루틴.
' 흐  름  : 대상 통합문서 확인 → 집계표 탐지 → 산출 시트 생성
'           → 분석문/차트데이터/삽입표/QA/메타 작성 → 로그 기록
' 인  자  : silent=True이면 UI 메시지 없이 오류를 호출자에게 다시 발생시킨다.
' ============================================================
Private Sub ReportAutomation_GenerateExcelOutputsCore(ByVal silent As Boolean)
    On Error GoTo ErrHandler

    ' 실행 대상은 반드시 사용자가 연 데이터 통합문서여야 한다.
    ' ThisWorkbook은 추가기능 자신이므로 여기에 산출 시트를 만들면 안 된다.
    If ActiveWorkbook Is Nothing Then Exit Sub
    If ActiveWorkbook Is ThisWorkbook Then
        If silent Then
            Err.Raise vbObjectError + 7200, "ReportAutomation_GenerateExcelOutputs", "보고서 산출 대상 통합문서를 먼저 활성화하세요."
        Else
            MsgBox "보고서 산출 대상 통합문서를 먼저 활성화하세요.", vbExclamation, "보고서 자동화"
        End If
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' 우선 활성 시트, 그 다음 Sheet1, 마지막으로 모든 시트를 훑어 집계표 원본 시트를 찾는다.
    Dim dataWs As Worksheet
    Set dataWs = ReportAutomation_ResolveDataSheet(wb)
    If dataWs Is Nothing Then
        If silent Then
            Err.Raise vbObjectError + 7201, "ReportAutomation_GenerateExcelOutputs", "집계표 제목([표 ...])이 있는 시트를 찾지 못했습니다."
        Else
            MsgBox "집계표 제목([표 ...])이 있는 시트를 찾지 못했습니다.", vbExclamation, "보고서 자동화"
        End If
        Exit Sub
    End If

    ' 원본 시트에서 [표 ...] 제목행을 기준으로 표 블록 메타데이터를 수집한다.
    Dim tables As Collection
    Set tables = ReportAutomation_CollectTables(dataWs)
    If tables.Count = 0 Then
        If silent Then
            Err.Raise vbObjectError + 7202, "ReportAutomation_GenerateExcelOutputs", "집계표 제목([표 ...])을 찾지 못했습니다."
        Else
            MsgBox "집계표 제목([표 ...])을 찾지 못했습니다.", vbExclamation, "보고서 자동화"
        End If
        Exit Sub
    End If

    ReportAutomation_BeginOperation "보고서 자동화 산출 시트를 생성하는 중입니다..."

    ' 산출물은 한 번의 실행마다 timestamp가 붙은 새 시트로 생성한다.
    ' 기존 산출 시트를 덮어쓰지 않아 사용자가 이전 결과와 비교할 수 있다.
    Dim wsSettings As Worksheet, wsList As Worksheet, wsNarr As Worksheet
    Dim wsChart As Worksheet, wsInsert As Worksheet, wsQA As Worksheet
    Dim wsSource As Worksheet, wsRevision As Worksheet, wsMeta As Worksheet
    Dim wsPriorSettings As Worksheet

    ' 사용자가 이전 산출 설정 시트를 편집한 경우, 새 실행에서도 그 값을 승계한다.
    Set wsPriorSettings = ReportAutomation_FindLatestOutputSheet(wb, "보고서_설정")

    Set wsSettings = ReportAutomation_AddOutputSheet(wb, "보고서_설정")
    Set wsList = ReportAutomation_AddOutputSheet(wb, "보고서_표목록")
    Set wsNarr = ReportAutomation_AddOutputSheet(wb, "보고서_분석문")
    Set wsChart = ReportAutomation_AddOutputSheet(wb, "보고서_차트데이터")
    Set wsInsert = ReportAutomation_AddOutputSheet(wb, "보고서_삽입표")
    Set wsQA = ReportAutomation_AddOutputSheet(wb, "보고서_QA")
    Set wsSource = ReportAutomation_AddOutputSheet(wb, "보고서_출처")
    Set wsRevision = ReportAutomation_AddOutputSheet(wb, "보고서_수정이력")
    Set wsMeta = ReportAutomation_AddOutputSheet(wb, "_ReportMeta")

    ' 각 산출 시트 작성은 역할별 전용 함수로 분리해 후속 리뷰와 테스트 범위를 좁힌다.
    ReportAutomation_WriteSettingsSheet wsSettings, wb, dataWs, wsPriorSettings
    If mHasOptionOverrides Then
        ReportAutomation_SetSettingValue wsSettings, "추출 배너 목록", mOverrideBannerSetting
        ReportAutomation_SetSettingValue wsSettings, "제목 제거 접두어", mOverrideTitlePrefixes
    End If
    ReportAutomation_WriteTableList wsList, tables, dataWs
    Dim qaCount As Long
    ReportAutomation_WriteNarratives wsNarr, wsChart, wsInsert, wsQA, tables, dataWs, wsSettings, qaCount
    ReportAutomation_WriteSourceSheet wsSource
    ReportAutomation_WriteRevisionSheet wsRevision
    ReportAutomation_WriteMetaSheet wsMeta, wb, dataWs, tables, wsSettings, wsList, wsNarr, wsChart, wsInsert

    ' 메타 시트는 Python/HWP/PPT 후속 자동화가 참조하는 내부 정보이므로 숨김 처리한다.
    wsMeta.Visible = xlSheetVeryHidden

    ReportAutomation_LogEvent wb, "generate_excel_outputs", dataWs.Name, "OK", tables.Count & "개 표 탐지"
    ReportAutomation_EndOperation

    If Not silent Then
        Dim qaMsg As String
        qaMsg = IIf(qaCount > 0, vbCrLf & "QA 경고: " & qaCount & "건  (보고서_QA 시트 참조)", "")
        MsgBox "보고서 자동화 산출 시트를 생성했습니다." & vbCrLf & _
               "원본 시트: " & dataWs.Name & vbCrLf & _
               "탐지 표 수: " & tables.Count & qaMsg, vbInformation, "보고서 자동화"
        If MsgBox("통합문서를 저장하시겠습니까?", vbQuestion + vbYesNo, "보고서 자동화") = vbYes Then
            On Error Resume Next
            wb.Save
            If Err.Number <> 0 Then MsgBox "저장 중 오류: " & Err.Description, vbExclamation, "보고서 자동화"
            On Error GoTo 0
        End If
    End If
    ReportAutomation_ClearOptionOverrides
    Exit Sub

ErrHandler:
    ' EndOperation 내부의 On Error Resume Next가 Err를 소거하기 전에 저장한다.
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    ReportAutomation_EndOperation
    On Error Resume Next
    If errNum = 18 Then
        ReportAutomation_LogEvent wb, "generate_excel_outputs", "", "CANCELLED", "사용자가 작업을 취소했습니다."
        On Error GoTo 0
        If Not silent Then MsgBox "작업이 취소되었습니다.", vbInformation, "보고서 자동화"
        ReportAutomation_ClearOptionOverrides
        Exit Sub
    End If
    ReportAutomation_LogEvent wb, "generate_excel_outputs", "", "ERROR", errDesc
    On Error GoTo 0
    If silent Then
        Err.Raise errNum, errSrc, errDesc
    Else
        MsgBox "보고서 자동화 산출 시트 생성 중 오류가 발생했습니다." & vbCrLf & errDesc, vbExclamation, "보고서 자동화"
    End If
    ReportAutomation_ClearOptionOverrides
End Sub

' ============================================================
' 함수명 : ReportAutomation_ResolveDataSheet
' 설  명 : 집계표 원본 시트를 결정한다.
' 우선순위: 배너조정 시트 → 현재 활성 시트 → Sheet1 → 통합문서 내 첫 번째 집계표 포함 시트
' 반환값 : [표 ...] 제목행이 있는 Worksheet, 없으면 Nothing
' ============================================================
Private Function ReportAutomation_ResolveDataSheet(ByVal wb As Workbook) As Worksheet
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
' 함수명 : ReportAutomation_CountTableTitleRows
' 설  명 : A열에서 [표 ...] 형식의 제목행 개수를 센다.
' 반환값 : 해당 시트가 집계표 원본인지 판단하는 근거 수치
' ============================================================
Private Function ReportAutomation_CountTableTitleRows(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ReportAutomation_LastUsedRow(ws)
    If lastRow < 1 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).Value
    Dim r As Long
    For r = 1 To lastRow
        If ReportAutomation_IsTableTitle(arr(r, 1)) Then
            ReportAutomation_CountTableTitleRows = ReportAutomation_CountTableTitleRows + 1
        End If
    Next r
End Function

' ============================================================
' 함수명 : ReportAutomation_CollectTables
' 설  명 : [표 ...] 제목행을 기준으로 표 블록 범위와 기본 메타데이터를 수집한다.
' 반환값 : Collection(Array(table_key, table_no, title, basis, base_label,
'                          start_row, end_row, last_col, table_type, warning))
' 주  의 : 현재 MVP는 A열 제목행 기준의 연속 블록 구조를 전제로 한다.
' ============================================================
Private Function ReportAutomation_CollectTables(ByVal ws As Worksheet) As Collection
    Dim titleRows As New Collection
    Dim lastRow As Long
    lastRow = ReportAutomation_LastUsedRow(ws)

    Dim r As Long
    For r = 1 To lastRow
        If ReportAutomation_IsTableTitle(ws.Cells(r, 1).Value) Then titleRows.Add r
    Next r

    ' 각 제목행부터 다음 제목행 직전까지를 하나의 표 블록으로 본다.
    Dim tables As New Collection
    Dim i As Long
    For i = 1 To titleRows.Count
        Dim startRow As Long, endRow As Long, lastCol As Long
        startRow = CLng(titleRows(i))
        If i < titleRows.Count Then
            endRow = CLng(titleRows(i + 1)) - 1
        Else
            endRow = ReportAutomation_LastNonEmptyRowInBlock(ws, startRow, lastRow)
        End If
        lastCol = ReportAutomation_LastNonEmptyColInBlock(ws, startRow, endRow)

        ' 제목 문자열에서 표번호, 분석 제목, 기준 문구를 분리한다.
        Dim titleText As String, tableNo As String, cleanTitle As String, basis As String
        titleText = CStr(ws.Cells(startRow, 1).Value)
        tableNo = ReportAutomation_ParseTableNo(titleText)
        cleanTitle = ReportAutomation_ParseTableTitle(titleText)
        basis = ReportAutomation_ParseBasis(titleText)

        Dim baseLabel As String
        If startRow + 1 <= endRow Then baseLabel = CStr(ws.Cells(startRow + 1, 1).Value)

        ' 배열 기반 레코드로 저장한다. 후속 단계에서 rec(index)로 빠르게 접근한다.
        tables.Add Array( _
            "T" & Format$(i, "0000"), tableNo, cleanTitle, basis, baseLabel, _
            startRow, endRow, lastCol, _
            ReportAutomation_ClassifyTable(ws, startRow, endRow, lastCol), _
            ReportAutomation_TableWarning(ws, startRow, endRow, lastCol))
    Next i

    Set ReportAutomation_CollectTables = tables
End Function

' ============================================================
' 함수명 : ReportAutomation_LastUsedRow
' 설  명 : A열뿐 아니라 전체 UsedRange 기준으로 마지막 사용 행을 찾는다.
' ============================================================
Private Function ReportAutomation_LastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then
        ReportAutomation_LastUsedRow = 1
    Else
        ReportAutomation_LastUsedRow = lastCell.Row
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_IsTableTitle
' 설  명 : 셀 값이 집계표 제목행인지 판정한다.
' 기준   : 왼쪽 공백 제거 후 "[표 "로 시작
' ============================================================
Private Function ReportAutomation_IsTableTitle(ByVal valueData As Variant) As Boolean
    Dim text As String
    text = Trim$(CStr(valueData))
    ReportAutomation_IsTableTitle = (Left$(text, 3) = "[표 " Or Left$(text, 3) = "[ 표")
End Function

' ============================================================
' 함수명 : ReportAutomation_LastNonEmptyRowInBlock
' 설  명 : 표 블록 안에서 마지막 값이 있는 행을 뒤에서부터 찾는다.
' 용  도 : 마지막 표는 다음 제목행이 없으므로 실제 끝 행을 추정해야 한다.
' ============================================================
Private Function ReportAutomation_LastNonEmptyRowInBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal maxRow As Long) As Long
    Dim r As Long, c As Long
    For r = maxRow To startRow Step -1
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then
                ReportAutomation_LastNonEmptyRowInBlock = r
                Exit Function
            End If
        Next c
    Next r
    ReportAutomation_LastNonEmptyRowInBlock = startRow
End Function

' ============================================================
' 함수명 : ReportAutomation_LastNonEmptyColInBlock
' 설  명 : 표 블록 내 사용된 마지막 열 번호를 계산한다.
' 반환값 : 최소 1 이상의 열 번호
' ============================================================
Private Function ReportAutomation_LastNonEmptyColInBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long, c As Long, lastCol As Long
    For r = startRow To endRow
        c = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If c > lastCol Then lastCol = c
    Next r
    If lastCol < 1 Then lastCol = 1
    ReportAutomation_LastNonEmptyColInBlock = lastCol
End Function

' ============================================================
' 함수명 : ReportAutomation_ParseTableNo
' 설  명 : "[표 MQT2 > ...]" 같은 제목에서 표 번호(MQT2)를 추출한다.
' 반환값 : 추출 실패 시 빈 문자열
' ============================================================
Private Function ReportAutomation_ParseTableNo(ByVal titleText As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, titleText, "[표 ", vbTextCompare)
    If startPos = 0 Then startPos = InStr(1, titleText, "[ 표", vbTextCompare)
    endPos = InStr(1, titleText, ">", vbTextCompare)
    If startPos > 0 And endPos > startPos Then
        ReportAutomation_ParseTableNo = Trim$(Mid$(titleText, startPos + 3, endPos - startPos - 3))
        Exit Function
    End If

    If startPos > 0 Then
        endPos = InStr(startPos + 1, titleText, "]", vbTextCompare)
        If endPos > startPos Then
            ReportAutomation_ParseTableNo = Trim$(Replace(Mid$(titleText, startPos + 2, endPos - startPos - 2), "표", ""))
        End If
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_ParseTableTitle
' 설  명 : 표 제목 문자열에서 보고서용 제목 부분만 추출한다.
' 예  시 : "[표 MDBR-6 > [리스트 기준] 대분류...]" → "[리스트 기준] 대분류..."
' ============================================================
Private Function ReportAutomation_ParseTableTitle(ByVal titleText As String) As String
    Dim endPos As Long
    endPos = InStr(1, titleText, ">", vbTextCompare)
    If endPos > 0 Then
        Dim parsedTitle As String
        parsedTitle = Trim$(Mid$(titleText, endPos + 1))
        If Right$(parsedTitle, 1) = "]" Then parsedTitle = Left$(parsedTitle, Len(parsedTitle) - 1)
        ReportAutomation_ParseTableTitle = Trim$(parsedTitle)
    Else
        Dim text As String
        text = ReportAutomation_CleanText(titleText)

        If Left$(text, 3) = "[ 표" Then
            endPos = InStr(1, text, "]", vbTextCompare)
            If endPos > 0 Then text = Trim$(Mid$(text, endPos + 1))

            If Left$(text, 1) = "[" Then
                endPos = InStr(1, text, "]", vbTextCompare)
                If endPos > 0 Then text = Trim$(Mid$(text, endPos + 1))
            End If

            Dim dividerPos As Long
            dividerPos = InStr(1, text, "─", vbTextCompare)
            If dividerPos > 1 Then text = Trim$(Left$(text, dividerPos - 1))

            Dim questionPos As Long
            questionPos = InStr(1, text, "[ 문", vbTextCompare)
            If questionPos > 1 Then text = Trim$(Left$(text, questionPos - 1))

            Dim cleanedText As String
            cleanedText = ReportAutomation_CleanText(text)
            If Left$(text, 1) = "'" And Left$(cleanedText, 1) <> "'" Then cleanedText = "'" & cleanedText
            ReportAutomation_ParseTableTitle = cleanedText
        Else
            ReportAutomation_ParseTableTitle = titleText
        End If
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_ParseBasis
' 설  명 : 제목 안의 기준 문구([리스트 기준], [응답 기준] 등)를 추출한다.
' 주  의 : 현재는 "> [" 다음 첫 닫는 대괄호까지를 기준 문구로 본다.
' ============================================================
Private Function ReportAutomation_ParseBasis(ByVal titleText As String) As String
    Dim openPos As Long, closePos As Long
    openPos = InStr(1, titleText, "> [", vbTextCompare)
    If openPos > 0 Then
        closePos = InStr(openPos + 3, titleText, "]", vbTextCompare)
        If closePos > openPos Then
            ReportAutomation_ParseBasis = Mid$(titleText, openPos + 2, closePos - openPos - 1)
        End If
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_ClassifyTable
' 설  명 : 표 폭과 키워드로 표 유형을 대략 분류한다.
' 반환값 : 단일 전체 분포표 / 평균 포함 표 / N/% 쌍 표 / 배너 교차표
' ============================================================
Private Function ReportAutomation_ClassifyTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As String
    If lastCol <= SIMPLE_TABLE_MAX_COL Then
        ReportAutomation_ClassifyTable = "단일 전체 분포표"
    ElseIf ReportAutomation_BlockContainsText(ws, startRow, endRow, lastCol, "▶평균◀") Then
        ReportAutomation_ClassifyTable = "평균 포함 표"
    ElseIf ReportAutomation_BlockContainsText(ws, startRow, startRow + 3, lastCol, "N") Then
        ReportAutomation_ClassifyTable = "N/% 쌍 표"
    Else
        ReportAutomation_ClassifyTable = "배너 교차표"
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_BlockContainsText
' 설  명 : 표 블록 범위 안에 특정 문자열이 포함되어 있는지 검사한다.
' 용  도 : 표 유형 분류, 평균 포함 여부, N/% 구조 판정
' ============================================================
Private Function ReportAutomation_BlockContainsText(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal needle As String) As Boolean
    If endRow > ws.Rows.Count Then endRow = ws.Rows.Count
    Dim arr As Variant
    arr = ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, lastCol)).Value
    If Not IsArray(arr) Then
        If InStr(1, CStr(arr), needle, vbTextCompare) > 0 Then ReportAutomation_BlockContainsText = True
        Exit Function
    End If
    Dim r As Long, c As Long
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If InStr(1, CStr(arr(r, c)), needle, vbTextCompare) > 0 Then
                ReportAutomation_BlockContainsText = True
                Exit Function
            End If
        Next c
    Next r
End Function

' ============================================================
' 함수명 : ReportAutomation_TableWarning
' 설  명 : 표 구조상 검토가 필요한 항목을 warning 문자열로 만든다.
' 현재검사: BASE 행 존재 여부, 사용 열 존재 여부
' ============================================================
Private Function ReportAutomation_TableWarning(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As String
    Dim warningText As String
    Dim baseCellValue As String
    If startRow + 1 <= endRow Then baseCellValue = CStr(ws.Cells(startRow + 1, 1).Value)
    If startRow + 1 > endRow Or _
       (InStr(1, baseCellValue, "BASE:", vbTextCompare) = 0 And InStr(1, baseCellValue, "Base :", vbTextCompare) = 0) Then
        warningText = "BASE 행 확인 필요"
    End If
    If lastCol <= 1 Then
        If Len(warningText) > 0 Then warningText = warningText & "; "
        warningText = warningText & "사용 열 확인 필요"
    End If
    ReportAutomation_TableWarning = warningText
End Function

' ============================================================
' 프로시저 : ReportAutomation_WriteNarratives
' 설  명  : 분석문, 차트 데이터, 삽입용 표, QA 시트를 동시에 작성한다.
' 이유    : 네 산출물은 같은 핵심 포인트(points)를 공유하므로 한 번의 순회로 생성한다.
' ============================================================
Private Sub ReportAutomation_WriteNarratives(ByVal wsNarr As Worksheet, ByVal wsChart As Worksheet, ByVal wsInsert As Worksheet, _
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
' 함수명 : ReportAutomation_ExtractKeyPoints
' 설  명 : 표 유형에 맞춰 핵심 응답 수치 목록을 추출하고 내림차순 정렬한다.
' 반환값 : Collection(Array(category, percent, weighted_n, raw_n, source_cell))
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
' 함수명 : ReportAutomation_IsScoreTable
' 설  명 : 인식도/만족도/동의도 조사에서 자주 쓰이는 100점 환산 또는 평균 점수 표인지 판정한다.
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
' 함수명 : ReportAutomation_HeaderBlockText
' 설  명 : 제목 다음 헤더 영역의 텍스트를 모아 점수형 표 단서를 찾는 데 사용한다.
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
' 프로시저 : ReportAutomation_ExtractScorePoints
' 설  명  : 100점 환산/평균 점수 표에서 항목별 점수를 추출한다.
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
' 함수명 : ReportAutomation_IsScoreMeasureColumn
' 설  명 : 헤더 영역에 100점/점수/평균 표기가 있는 점수 열인지 판정한다.
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
' 함수명 : ReportAutomation_CountScoreMeasureColumns
' 설  명 : 표 헤더에서 점수형 측정 열 개수를 센다.
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
' 함수명 : ReportAutomation_CountPercentMeasureColumns
' 설  명 : 표 헤더에서 % 측정 열 개수를 센다.
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
' 함수명 : ReportAutomation_ScoreHeaderLabel
' 설  명 : N/점수 쌍 구조에서 점수 열 왼쪽의 문항명을 찾아 반환한다.
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
' 함수명 : ReportAutomation_IsScoreHeaderLabelCandidate
' 설  명 : 점수 열 헤더에서 문항명으로 쓸 수 있는 텍스트인지 판정한다.
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
' 함수명 : ReportAutomation_CleanScoreItemLabel
' 설  명 : 점수형 종합표 문항명 앞의 기호를 제거한다.
' ============================================================
Private Function ReportAutomation_CleanScoreItemLabel(ByVal text As String) As String
    text = ReportAutomation_CleanText(text)
    Do While Len(text) > 0 And InStr(1, "○□", Left$(text, 1), vbBinaryCompare) > 0
        text = Trim$(Mid$(text, 2))
    Loop
    ReportAutomation_CleanScoreItemLabel = text
End Function

' ============================================================
' 함수명 : ReportAutomation_FindScoreColumn
' 설  명 : 단순 표에서 100점/평균/점수 열을 찾는다.
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
' 함수명 : ReportAutomation_RowCategory
' 설  명 : SPSS 집계표의 행 범주명을 B열 우선, A열 보조로 읽는다.
' ============================================================
Private Function ReportAutomation_RowCategory(ByVal ws As Worksheet, ByVal rowIndex As Long) As String
    ReportAutomation_RowCategory = Trim$(CStr(ws.Cells(rowIndex, 2).Value))
    If Len(ReportAutomation_RowCategory) = 0 Then
        ReportAutomation_RowCategory = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
    End If
End Function

' ============================================================
' 프로시저 : ReportAutomation_ExtractSimplePoints
' 설  명  : 단일 전체 분포표에서 항목별 비율을 추출한다.
' 전  제  : 항목명은 보통 A/B열, 비율은 % 헤더 열 또는 기본 5열에 있다.
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
' 프로시저 : ReportAutomation_ExtractWidePoints
' 설  명  : 배너 교차표에서 전체 행(■로 시작)의 비율 값을 추출한다.
' 전  제  : 전체 행의 각 % 열이 보고서 전체 요약의 후보 수치가 된다.
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
' 함수명 : ReportAutomation_FindHeaderColumn
' 설  명 : 표 상단 일부 행에서 특정 헤더 텍스트가 있는 열을 찾는다.
' 반환값 : 열 번호, 찾지 못하면 0
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
' 함수명 : ReportAutomation_FindTotalRow
' 설  명 : 배너 교차표에서 전체/합계 역할을 하는 행을 찾는다.
' 기준   : A열 값이 "■"로 시작하는 행
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
' 함수명 : ReportAutomation_IsPercentMeasureColumn
' 설  명 : 특정 열이 N/% 쌍 중 비율(%) 열인지 판정한다.
' 기준   : 전체 행 위쪽 헤더 영역에 "%" 표시가 있는지 확인
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
' 함수명 : ReportAutomation_HeaderLabel
' 설  명 : 병합/다단 헤더 구조에서 특정 열의 응답 범주명을 추정한다.
' 방식   : 현재 열에서 왼쪽으로 거슬러 올라가며 의미 있는 텍스트를 찾는다.
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
' 함수명 : ReportAutomation_IsHeaderLabelCandidate
' 설  명 : 헤더에서 보고서 항목명으로 쓸 수 있는 텍스트인지 판정한다.
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
' 함수명 : ReportAutomation_CleanHeaderItemLabel
' 설  명 : 보고서 문장에 불필요한 헤더 기호를 제거한다.
' ============================================================
Private Function ReportAutomation_CleanHeaderItemLabel(ByVal text As String) As String
    text = ReportAutomation_CleanText(text)
    Do While Len(text) > 0 And InStr(1, "◐○□", Left$(text, 1), vbBinaryCompare) > 0
        text = Trim$(Mid$(text, 2))
    Loop
    ReportAutomation_CleanHeaderItemLabel = text
End Function

' ============================================================
' 함수명 : ReportAutomation_IsReportItemLabel
' 설  명 : 계/빈도/사례수 등 문장 후보가 아닌 항목을 제외한다.
' ============================================================
Private Function ReportAutomation_IsReportItemLabel(ByVal text As String) As Boolean
    text = ReportAutomation_CleanText(text)
    If Len(text) = 0 Then Exit Function
    If text = "계" Or text = "빈도" Or text = "사례수" Or text = "%" Then Exit Function
    ReportAutomation_IsReportItemLabel = True
End Function

' ============================================================
' 프로시저 : ReportAutomation_SortPointsDescending
' 설  명  : 핵심 수치 목록을 percent 값 기준 내림차순으로 정렬한다.
' 주  의  : VBA Collection은 직접 정렬이 어려워 배열로 복사 후 다시 채운다.
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
' 함수명 : ReportAutomation_BuildNarrative
' 설  명 : 핵심 수치 목록을 보고서 본문 문장으로 변환한다.
' 문체   : 첨부 Python 작성기와 맞춰 "가장 높게 나타남 / 다음으로 ... 순" 구조 사용
' 반환값 : 분석문_기본 및 최종 사용문에 들어갈 텍스트
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
' 함수명 : ReportAutomation_BuildScoreNarrative
' 설  명 : 인식도/만족도/동의도 등 100점 환산 점수형 문장을 생성한다.
' 예  시 : '권고안 관련 동의도'를 100점 환산 기준으로 분석한 결과, ...
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
' 함수명 : ReportAutomation_FindPointByKind
' 설  명 : 핵심 수치 목록에서 지정한 kind의 첫 번째 레코드를 찾는다.
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
' 함수명 : ReportAutomation_FindScaleContrastPoint
' 설  명 : 척도형 표에서 부정/비인지/불만족 등 반대 방향 보조 수치를 찾는다.
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
' 함수명 : ReportAutomation_IsScaleSummaryPoint
' 설  명 : 긍정/만족/인지 등 척도형 표의 대표 지표인지 판정한다.
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
' 함수명 : ReportAutomation_CleanScaleLabel
' 설  명 : 척도형 대표 지표명에서 보고서 문장에 불필요한 공백을 정리한다.
' ============================================================
Private Function ReportAutomation_CleanScaleLabel(ByVal text As String) As String
    text = ReportAutomation_CleanText(text)
    text = Replace(text, " (", "(")
    text = Replace(text, "(합)", "(합)")
    ReportAutomation_CleanScaleLabel = text
End Function

' ============================================================
' 함수명 : ReportAutomation_IsNeutralCompositeLabel
' 설  명 : 원항목과 중복되는 보통(합) 계열의 보조 지표인지 판정한다.
' ============================================================
Private Function ReportAutomation_IsNeutralCompositeLabel(ByVal text As String) As Boolean
    text = Replace(ReportAutomation_CleanText(text), " ", "")
    ReportAutomation_IsNeutralCompositeLabel = (InStr(1, text, "보통(", vbTextCompare) > 0)
End Function

' ============================================================
' 함수명 : ReportAutomation_CleanScoreTitle
' 설  명 : 문장 본문에서는 표 제목의 점수 단위 보조 표기를 제거한다.
' ============================================================
Private Function ReportAutomation_CleanScoreTitle(ByVal titleText As String) As String
    Dim text As String
    text = ReportAutomation_CleanText(titleText)
    text = Replace(text, "(100점)", "")
    text = Replace(text, "（100점）", "")
    ReportAutomation_CleanScoreTitle = ReportAutomation_CleanText(text)
End Function

' ============================================================
' 함수명 : ReportAutomation_FormatPointSummary
' 설  명 : 리뷰용으로 핵심 수치와 원본 셀 주소를 압축 표기한다.
' 예  시 : 도소매업=43.1%(E12)
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
' 함수명 : ReportAutomation_NormalizeAnalysisTitle
' 설  명 : 표 제목에서 보고서 문장에 불필요한 기준/메타 문구를 제거한다.
' 예  시 : "[리스트 기준] 대분류..." → "대분류..."
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
' 함수명 : ReportAutomation_PointMeasure
' 설  명 : 차트/삽입표에 기록할 수치 단위명을 반환한다.
' ============================================================
Private Function ReportAutomation_PointMeasure(ByVal pointRec As Variant) As String
    If ReportAutomation_PointKind(pointRec) = "score_100" Then
        ReportAutomation_PointMeasure = "100점"
    Else
        ReportAutomation_PointMeasure = "%"
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_FormatPointValue
' 설  명 : point 유형에 따라 표시값을 % 또는 점수로 포맷한다.
' ============================================================
Private Function ReportAutomation_FormatPointValue(ByVal pointRec As Variant) As String
    If ReportAutomation_PointKind(pointRec) = "score_100" Then
        ReportAutomation_FormatPointValue = ReportAutomation_FormatScore(CDbl(pointRec(1)))
    Else
        ReportAutomation_FormatPointValue = ReportAutomation_FormatPercent(CDbl(pointRec(1)))
    End If
End Function

' ============================================================
' 프로시저 : ReportAutomation_ClearOptionOverrides
' 설  명  : UserForm에서 넘긴 1회성 override 값을 초기화한다.
' ============================================================
Private Sub ReportAutomation_ClearOptionOverrides()
    mHasOptionOverrides = False
    mOverrideBannerSetting = ""
    mOverrideTitlePrefixes = ""
End Sub

' ============================================================
' 함수명 : ReportAutomation_FindBannerGroups
' 설  명 : 배너 교차표 헤더 영역을 스캔해 배너 그룹명과 소속 % 열을 매핑한다.
' 반환값 : Collection
'   각 원소: Array(bannerName As String, pctCols As Variant)
'     - bannerName : 배너 그룹명 (예: "전체", "성별", "연령대")
'     - pctCols    : 해당 그룹에 속한 % 열 번호 Long 배열
' 동작 원리:
'   헤더 행에서 N/% 가 아닌 텍스트를 배너 그룹명 후보로 보고,
'   그 열부터 다음 그룹 시작 전까지의 % 열들을 묶는다.
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
' 함수명 : ReportAutomation_BannerRequested
' 설  명 : 배너 설정 문자열에 특정 배너명이 포함되어 있는지 확인한다.
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
