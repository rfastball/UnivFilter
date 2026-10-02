Attribute VB_Name = "modReport"
' @desc 누락/보정 리포트 구성과 클립보드 복사 담당
' @depends ClipboardAPI, modLogger
Option Explicit

Private Const MISSING_PREVIEW_MAX As Long = 5
Private Const STATUS_CLEAR_SECS As Long = 7

' 상태바 자동 소거 예약 관리 (재표시 시 이전 예약 취소용)
Private mStatusClearAt As Date
Private mStatusPending As Boolean

' @desc 매칭 결과 받아 상태/누락/보정 메시지 딕셔너리 반환
Public Function BuildMessages(ByVal outcome As Object, ByVal bExclude As Boolean, _
                              ByVal applyInfo As Object) As Object
    Dim msgs As Object
    Set msgs = CreateObject("Scripting.Dictionary")

    Dim vMissing As Variant
    vMissing = outcome("Missing")
    Dim nMissing As Long
    nMissing = ArrCount(vMissing)

    Dim corrections As Object
    Set corrections = outcome("Corrections")

    ' 상태 메시지
    Dim s As String
    s = "매칭 " & outcome("Matched").Count & "건"
    If bExclude Then s = s & " 제외"
    If nMissing > 0 Then s = s & ", 누락 " & nMissing & "건"
    If corrections.Count > 0 Then s = s & ", 자동보정 " & corrections.Count & "건"
    msgs("StatusMessage") = s

    ' 누락 메시지 (포함 모드 전용, 클립보드 복사는 호출부가 CopyMissing으로)
    If nMissing > 0 And Not bExclude Then
        Dim m As String, i As Long, shown As Long
        m = "다음 " & nMissing & "건은 목록에 없습니다:" & vbCrLf
        For i = LBound(vMissing) To UBound(vMissing)
            If shown >= MISSING_PREVIEW_MAX Then
                m = m & "... 외 " & (nMissing - shown) & "건" & vbCrLf
                Exit For
            End If
            m = m & "- " & CStr(vMissing(i)) & vbCrLf
            shown = shown + 1
        Next i
        m = m & vbCrLf & "(누락 목록이 클립보드에 복사되었습니다)"
        msgs("MissingMessage") = m
    Else
        msgs("MissingMessage") = ""
    End If

    Dim vKey As Variant
    ' 자동보정 인라인 요약 (상태바용 한 줄 - 보정은 결정론 동작이라 결과가 곧 피드백,
    ' 팝업 대신 상태바로만 알린다)
    If corrections.Count > 0 Then
        Dim inl As String
        For Each vKey In corrections.Keys
            If Len(inl) > 0 Then inl = inl & ", "
            inl = inl & CStr(vKey) & " -> " & CStr(corrections(vKey))
            If Len(inl) > 120 Then
                inl = inl & " ..."
                Exit For
            End If
        Next vKey
        msgs("CorrectionInline") = inl
    Else
        msgs("CorrectionInline") = ""
    End If

    ' 255자 초과 탈락 경고
    Dim vDropped As Variant
    vDropped = applyInfo("DroppedLong")
    If ArrCount(vDropped) > 0 Then
        msgs("WarningMessage") = "255자를 초과하는 값 " & ArrCount(vDropped) & _
                                 "건은 필터에 반영되지 못했습니다."
    Else
        msgs("WarningMessage") = ""
    End If

    Set BuildMessages = msgs
End Function

' @desc 결과 한 줄을 상태바에 표시함, 일정 시간 후 자동 소거 예약
Public Sub ShowStatusBar(ByVal sMsg As String)
    On Error Resume Next
    If mStatusPending Then
        Application.OnTime mStatusClearAt, "modReport.ClearStatusBar", , False
        mStatusPending = False
    End If
    Application.StatusBar = sMsg
    mStatusClearAt = Now + TimeSerial(0, 0, STATUS_CLEAR_SECS)
    Application.OnTime mStatusClearAt, "modReport.ClearStatusBar"
    mStatusPending = True
    On Error GoTo 0
End Sub

' @desc 상태바를 기본 상태로 되돌림 (OnTime 콜백)
Public Sub ClearStatusBar()
    On Error Resume Next
    mStatusPending = False
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' @desc 누락 배열을 줄바꿈 목록으로 클립보드에 복사함
Public Sub CopyMissing(ByVal vMissing As Variant)
    If ArrCount(vMissing) = 0 Then Exit Sub
    Dim s As String
    s = Join(vMissing, vbCrLf)
    If Not ClipboardAPI.SetClipboardText(s) Then
        Log "CopyMissing: 클립보드 복사 실패", lvlWarning
    End If
End Sub

Private Function ArrCount(ByVal v As Variant) As Long
    If Not IsArray(v) Then Exit Function
    ArrCount = UBound(v) - LBound(v) + 1
    If ArrCount < 0 Then ArrCount = 0
End Function

' @desc 결과 딕셔너리 받아 상태바 요약 반환, 보정 있으면 인라인 병기
Public Function StatusLine(ByVal result As Object) As String
    StatusLine = "UnivFilter: " & result("StatusMessage")
    If Len(result("CorrectionInline")) > 0 Then
        StatusLine = StatusLine & " | 보정: " & result("CorrectionInline")
    End If
End Function
