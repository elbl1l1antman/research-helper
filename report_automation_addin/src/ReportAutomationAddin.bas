Attribute VB_Name = "ReportAutomationAddin"
' ============================================================
' 모듈명  : ReportAutomationAddin
' 설  명  : 집계표 기반 보고서 산출 시트 생성 전용 추가기능
' ============================================================
Option Explicit

' 추가기능 버전. 산출 메타 시트와 로그 시트에 함께 기록해 결과 파일의 생성 기준을 추적한다.
Public Const REPORT_AUTOMATION_VERSION As String = "0.0.23"

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

' 프로시저 : ReportAutomation_ClearOptionOverrides
' 설  명  : UserForm에서 넘긴 1회성 override 값을 초기화한다.
' ============================================================
Private Sub ReportAutomation_ClearOptionOverrides()
    mHasOptionOverrides = False
    mOverrideBannerSetting = ""
    mOverrideTitlePrefixes = ""
End Sub
