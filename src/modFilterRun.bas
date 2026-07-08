Attribute VB_Name = "modFilterRun"
' @desc 필터 실행 파이프라인 오케스트레이션 담당
' @depends modCache, modInputParse, modMatch, modApply, modReport, clsColAsset, clsSet, Utilities, modLogger
Option Explicit

' @desc 입력 받아 매칭-적용-리포트 실행 후 결과 딕셔너리 반환, 실패 시 Ok=False
Public Function RunFilter(ByVal rngBody As Range, ByVal rngFilterArea As Range, _
                          ByVal lField As Long, ByVal sRawText As String, _
                          ByVal bExclude As Boolean, ByVal bStack As Boolean, _
                          ByVal bPartial As Boolean, _
                          Optional ByVal bQuickZeroNoop As Boolean = False) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Ok") = False
    result("QuickNoMatch") = False
    result("ErrMsg") = ""
    result("MatchedCount") = 0&
    result("SelfCheckNote") = ""

    On Error GoTo Fail
    Utilities.OptimizedMode True

    ' 1. 자산 취득 (캐시가 세대 검증과 재빌드를 내부에서 처리)
    Dim asset As clsColAsset
    Set asset = modCache.GetColumnAsset(rngBody)

    ' 2. 토큰화
    Dim tokens As Variant
    tokens = modInputParse.Tokenize(sRawText)
    If ArrCount(tokens) = 0 Then
        result("ErrMsg") = "검색할 내용이 없습니다."
        GoTo Cleanup
    End If

    ' 3. 매칭 (2레인)
    Dim outcome As Object
    Set outcome = modMatch.Execute(tokens, asset, bPartial)
    Dim matched As clsSet
    Set matched = outcome("Matched")
    result("MatchedCount") = matched.Count
    result("Missing") = outcome("Missing")

    ' 4. 적용
    Dim applyInfo As Object
    Dim expectedRows As Long
    expectedRows = -1

    If matched.Count = 0 And bQuickZeroNoop Then
        ' 퀵 모드 0건: 필터/화면을 일절 건드리지 않는다
        result("QuickNoMatch") = True
        Set applyInfo = EmptyApplyInfo()
    ElseIf bExclude Then
        If matched.Count = 0 Then
            ' 제외 대상 없음: 이 필드만 해제 (no-op에 가깝게)
            If Not bStack Then modApply.ResetFilters rngFilterArea
            modApply.ClearField rngFilterArea, lField
            Set applyInfo = EmptyApplyInfo()
        Else
            Dim comp As clsSet
            Set comp = asset.Displays.Difference(matched)
            If comp.Count = 0 And Not asset.HasBlanks Then
                modApply.ApplyHideAll rngFilterArea, lField, bStack
                Set applyInfo = EmptyApplyInfo()
                expectedRows = 0
            Else
                Set applyInfo = modApply.ApplyDisplays(rngFilterArea, lField, comp, _
                                                       asset.HasBlanks, bStack)
                expectedRows = asset.ExpectedRows(comp, asset.HasBlanks)
            End If
        End If
    Else
        If matched.Count = 0 Then
            modApply.ApplyHideAll rngFilterArea, lField, bStack
            Set applyInfo = EmptyApplyInfo()
            expectedRows = 0
        Else
            Set applyInfo = modApply.ApplyDisplays(rngFilterArea, lField, matched, False, bStack)
            expectedRows = asset.ExpectedRows(matched, False)
        End If
    End If

    ' 5. 자가검증: Reset 단일 필드에서만 기대 행수가 유효하다
    If expectedRows >= 0 And Not bStack Then
        Dim vis As Long
        vis = modApply.CountVisibleRows(rngFilterArea)
        If vis <> expectedRows Then
            result("SelfCheckNote") = "적용 결과(" & vis & "행)가 기대(" & _
                                      expectedRows & "행)와 다릅니다."
            Log "RunFilter: 자가검증 불일치 vis=" & vis & " expected=" & expectedRows, lvlWarning
        End If
    End If

    ' 6. 리포트 구성
    Dim msgs As Object
    Set msgs = modReport.BuildMessages(outcome, bExclude, applyInfo)
    result("StatusMessage") = msgs("StatusMessage")
    result("MissingMessage") = msgs("MissingMessage")
    result("CorrectionMessage") = msgs("CorrectionMessage")
    result("WarningMessage") = msgs("WarningMessage")

    ' 누락 클립보드 복사 (포함 모드, 퀵 0건 무적용 제외)
    If Not bExclude And Not result("QuickNoMatch") Then
        If ArrCount(outcome("Missing")) > 0 Then modReport.CopyMissing outcome("Missing")
    End If

    result("Ok") = True

Cleanup:
    Utilities.OptimizedMode False
    Set RunFilter = result
    Exit Function

Fail:
    result("ErrMsg") = "오류 " & Err.Number & ": " & Err.Description
    Log "RunFilter: " & result("ErrMsg"), lvlError
    Resume Cleanup
End Function

' @desc 퀵 모드 실행: 완전일치/포함/초기화 고정, 0건이면 무적용 신호 반환
Public Function RunQuickFilter(ByVal rngBody As Range, ByVal rngFilterArea As Range, _
                               ByVal lField As Long, ByVal sClipText As String) As Object
    Set RunQuickFilter = RunFilter(rngBody, rngFilterArea, lField, sClipText, _
                                   False, False, False, True)
End Function

Private Function EmptyApplyInfo() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Mode") = "none"
    d("Applied") = 0&
    d("DroppedLong") = Array()
    Set EmptyApplyInfo = d
End Function

Private Function ArrCount(ByVal v As Variant) As Long
    If Not IsArray(v) Then Exit Function
    ArrCount = UBound(v) - LBound(v) + 1
    If ArrCount < 0 Then ArrCount = 0
End Function
