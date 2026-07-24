Attribute VB_Name = "ReportAutomationUtils"
' ============================================================
' 모듈명  : ReportAutomationUtils
' 설  명  : 보고서 자동화 공통 유틸리티
' ============================================================
Option Explicit

' ============================================================
' 함수명 : ReportAutomation_CleanText
' 설  명 : 줄바꿈/비분리 공백/중복 공백을 정리해 비교와 문장 생성에 쓸 텍스트를 만든다.
Public Function ReportAutomation_CleanText(ByVal text As String) As String
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    Do While InStr(1, text, "  ", vbBinaryCompare) > 0
        text = Replace(text, "  ", " ")
    Loop
    ReportAutomation_CleanText = Trim$(text)
End Function
' ============================================================
' 함수명 : ReportAutomation_FormatPercent
' 설  명 : 비율 값을 보고서 표준 소수점 한 자리 % 문자열로 변환한다.
Public Function ReportAutomation_FormatPercent(ByVal valueData As Double) As String
    ReportAutomation_FormatPercent = Format(valueData, "0.0") & "%"
End Function
' ============================================================
' 함수명 : ReportAutomation_FormatScore
' 설  명 : 점수 값을 보고서 표준 소수점 한 자리 점수 문자열로 변환한다.
Public Function ReportAutomation_FormatScore(ByVal valueData As Double) As String
    ReportAutomation_FormatScore = Format(valueData, "0.0") & "점"
End Function
' ============================================================
' 함수명 : ReportAutomation_Quoted
' 설  명 : 응답 항목명을 분석문용 작은따옴표 표기로 감싼다.
Public Function ReportAutomation_Quoted(ByVal text As String) As String
    ReportAutomation_Quoted = "'" & text & "'"
End Function
' ============================================================
' 함수명 : ReportAutomation_TopicParticle
' 설  명 : 단어의 받침 유무에 맞춰 은/는 조사를 반환한다.
Public Function ReportAutomation_TopicParticle(ByVal text As String) As String
    ReportAutomation_TopicParticle = IIf(ReportAutomation_HasBatchim(text), "은", "는")
End Function
' ============================================================
' 함수명 : ReportAutomation_ObjectParticle
' 설  명 : 단어의 받침 유무에 맞춰 을/를 조사를 반환한다.
Public Function ReportAutomation_ObjectParticle(ByVal text As String) As String
    ReportAutomation_ObjectParticle = IIf(ReportAutomation_HasBatchim(text), "을", "를")
End Function
' ============================================================
' 함수명 : ReportAutomation_SubjectParticle
' 설  명 : 단어의 받침 유무에 맞춰 이/가 조사를 반환한다.
Public Function ReportAutomation_SubjectParticle(ByVal text As String) As String
    ReportAutomation_SubjectParticle = IIf(ReportAutomation_HasBatchim(text), "이", "가")
End Function
' ============================================================
' 함수명 : ReportAutomation_HasBatchim
' 설  명 : 한글/숫자/일부 영문 끝글자를 기준으로 받침 유무를 판정한다.
Public Function ReportAutomation_HasBatchim(ByVal text As String) As Boolean
    Dim cleaned As String
    cleaned = ReportAutomation_CleanText(text)
    If Right$(cleaned, 1) = ")" Then
        Dim parenPos As Long
        parenPos = InStrRev(cleaned, "(")
        If parenPos > 1 Then cleaned = Trim$(Left$(cleaned, parenPos - 1))
    End If
    Do While Len(cleaned) > 0 And InStr(1, "'""”)]} ", Right$(cleaned, 1), vbBinaryCompare) > 0
        cleaned = Left$(cleaned, Len(cleaned) - 1)
    Loop
    If Len(cleaned) = 0 Then Exit Function

    Dim ch As String
    ch = Right$(cleaned, 1)

    ' 숫자는 한국어로 읽었을 때 받침이 있는 숫자만 받침 있음으로 본다.
    If ch Like "#" Then
        ReportAutomation_HasBatchim = (InStr(1, "013678", ch, vbBinaryCompare) > 0)
        Exit Function
    End If

    ' AscW가 음수로 반환되는 환경을 고려해 0~65535 범위로 보정한다.
    Dim code As Long
    code = AscW(ch)
    If code < 0 Then code = code + 65536

    ' 한글 완성형 음절은 (코드-가) Mod 28 값으로 받침 유무를 판정한다.
    If code >= 44032 And code <= 55203 Then
        ReportAutomation_HasBatchim = (((code - 44032) Mod 28) <> 0)
    Else
        ch = LCase$(ch)
        ReportAutomation_HasBatchim = (InStr(1, "bcklmnpst", ch, vbBinaryCompare) > 0)
    End If
End Function
' ============================================================
' 프로시저 : ReportAutomation_StyleHeader
' 설  명  : 산출 시트 헤더 행에 굵게/배경색/테두리 공통 서식을 적용한다.
Public Sub ReportAutomation_StyleHeader(ByVal targetRange As Range)
    With targetRange
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .Borders.LineStyle = xlContinuous
    End With
End Sub
' ============================================================
' 함수명 : ReportAutomation_UniqueSheetName
' 설  명 : Excel 31자 제한 안에서 중복되지 않는 산출 시트명을 만든다.
Public Function ReportAutomation_UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim name As String
    name = Left$(baseName & "_" & Format(Now, "mmdd_hhnn"), 31)
    Dim candidate As String
    candidate = name

    Dim n As Long
    n = 1
    Do While ReportAutomation_WorksheetExists(wb, candidate)
        n = n + 1
        Dim sfx As String: sfx = "_" & n
        candidate = Left$(name, 31 - Len(sfx)) & sfx
    Loop
    ReportAutomation_UniqueSheetName = candidate
End Function
' ============================================================
' 함수명 : ReportAutomation_CellText
' 설  명 : 병합셀인 경우 병합 영역 좌상단 값을 반환해 헤더 스캔을 안정화한다.
Public Function ReportAutomation_CellText(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    Dim cell As Range
    Set cell = ws.Cells(rowIndex, colIndex)

    If cell.MergeCells Then
        ReportAutomation_CellText = Trim$(CStr(cell.MergeArea.Cells(1, 1).Value))
    Else
        ReportAutomation_CellText = Trim$(CStr(cell.Value))
    End If
End Function
' ============================================================
' 함수명 : ReportAutomation_WorksheetExists
' 설  명 : 통합문서 안에 특정 이름의 워크시트가 있는지 확인한다.
Public Function ReportAutomation_WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    ReportAutomation_WorksheetExists = Not ws Is Nothing
End Function
' ============================================================
' 함수명 : ReportAutomation_CollectionToArray
' 설  명 : VBA Collection을 정렬/반복에 쓰기 쉬운 0 기반 Variant 배열로 변환한다.
Public Function ReportAutomation_CollectionToArray(ByVal col As Collection) As Variant
    If col.Count = 0 Then
        ReportAutomation_CollectionToArray = Array()
        Exit Function
    End If
    Dim arr() As Variant
    ReDim arr(0 To col.Count - 1)
    Dim i As Long
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    ReportAutomation_CollectionToArray = arr
End Function
' ============================================================
' 프로시저 : ReportAutomation_SortArrayDescending
' 설  명  : Array 안의 Array 레코드를 지정 숫자 인덱스 기준 내림차순으로 정렬한다.
Public Sub ReportAutomation_SortArrayDescending(ByRef arr As Variant, ByVal valueIndex As Long)
    Dim i As Long, j As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CDbl(arr(j)(valueIndex)) > CDbl(arr(i)(valueIndex)) Then
                Dim temp As Variant
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub
