Attribute VB_Name = "ReportAutomationNarrativeText"
' ============================================================
' 모듈명  : ReportAutomationNarrativeText
' 설  명  : 핵심 수치 목록을 보고서 분석문 문장으로 변환
' ============================================================
Option Explicit

' ============================================================
' 함수명 : ReportAutomation_BuildNarrative
' 설  명 : 핵심 수치 목록을 보고서 본문 문장으로 변환한다.
' 문체   : "가장 높게 나타남 / 다음으로 ... 순" 구조 사용
' 반환값 : 분석문_기본 및 최종 사용문에 들어갈 텍스트
' ============================================================
Public Function ReportAutomation_BuildNarrative(ByVal titleText As String, ByVal basisText As String, ByVal points As Collection, ByRef titlePrefixes() As String) As String
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
' 함수명 : ReportAutomation_NormalizeAnalysisTitle
' 설  명 : 표 제목에서 보고서 문장에 불필요한 기준/메타 문구를 제거한다.
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

' ============================================================
' 함수명 : ReportAutomation_IsMultiResponse
' 설  명 : 제목/기준 문구에서 복수응답 문항 여부를 판정한다.
' ============================================================
Private Function ReportAutomation_IsMultiResponse(ByVal titleText As String, ByVal basisText As String) As Boolean
    Dim text As String
    text = titleText & " " & basisText
    ReportAutomation_IsMultiResponse = (InStr(1, text, "복수", vbTextCompare) > 0 Or _
                                        InStr(1, text, "중복", vbTextCompare) > 0 Or _
                                        InStr(1, text, "multiple", vbTextCompare) > 0)
End Function
