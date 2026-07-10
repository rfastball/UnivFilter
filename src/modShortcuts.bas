Attribute VB_Name = "modShortcuts"
' @desc 전역 단축키 핸들러와 폼 주입 담당
' @depends frmFilter01, modUnivFilter, modFilterRun, modReport, ClipboardAPI, AutoCloseMsgAPI
Option Explicit

' ==========================================================
' [Ctrl + Shift + Q] 퀵 핸들러 (Quick Mode)
' 항상 즉시 실행. 매칭 0건이면 아무것도 적용하지 않고 알림만.
' ==========================================================
Public Sub UnivFilter_QuickHandler()
    Dim frm As New frmFilter01
    Dim sClip As String

    ' 1. 기본 준비 (범위 주입용 폼 인스턴스 재활용)
    If Not SetupFormData(frm) Then Unload frm: Exit Sub

    sClip = ClipboardAPI.GetClipboardText()
    If Trim(sClip) = "" Then
        AutoCloseMsgBox_Popup "클립보드가 비어 있습니다.", 1, "알림"
        Unload frm: Exit Sub
    End If

    ' 2. 즉시 실행 (완전일치/포함/초기화 고정)
    Dim result As Object
    Set result = modFilterRun.RunQuickFilter( _
                    frm.InjectBodyRange, frm.InjectFilterArea, _
                    frm.InjectRelativeCol, sClip)

    ' 3. 결과 알림: 팝업은 실행 가능 정보(오류/0건/누락)만.
    ' 자동보정은 결정론 동작이라 결과가 곧 피드백 - 상태바 한 줄로만 알린다.
    If Not result("Ok") Then
        AutoCloseMsgBox_Popup result("ErrMsg"), 3, "오류"
    ElseIf result("QuickNoMatch") Then
        ' 0건: 필터/폼/클립보드 전부 원상태 유지
        AutoCloseMsgBox_Popup "일치하는 값이 없어 필터를 적용하지 않았습니다.", 2, "알림"
    Else
        modReport.ShowStatusBar StatusLine(result)
        If Len(result("MissingMessage")) > 0 Then
            AutoCloseMsgBox_Popup result("MissingMessage"), 2
        End If
    End If

    ' 퀵 모드는 폼이 없어 무표정 멈춤이 생길 수 있어 최초 저속 스캔을 사후 안내
    If result("Ok") And Len(result("ScanNote")) > 0 Then
        AutoCloseMsgBox_Popup result("ScanNote"), 3, "스캔 안내"
    End If

    Unload frm
End Sub

' ==========================================================
' [Ctrl + Shift + U] UI 핸들러 (Form Mode)
' ==========================================================
Public Sub UnivFilter_UIHandler()
    Dim frm As New frmFilter01
    If SetupFormData(frm) Then
        modUnivFilter.MoveFormToCell_API frm, ActiveCell '폼 띄우기 전 위치이동
        frm.Show
    Else
        Unload frm
    End If
End Sub

' ==========================================================
' [Helper] 결과 요약을 상태바 한 줄로 조립 (퀵/폼 공용)
' ==========================================================
' @desc 결과 딕셔너리 받아 상태바 한 줄 요약 반환, 보정 있으면 인라인 병기
Public Function StatusLine(ByVal result As Object) As String
    Dim s As String
    s = "UnivFilter: " & result("StatusMessage")
    If Len(result("CorrectionInline")) > 0 Then
        s = s & " | 보정: " & result("CorrectionInline")
    End If
    StatusLine = s
End Function

' ==========================================================
' [Helper] 현재 열 정보를 폼에 주입하는 공통 로직
' ==========================================================
Private Function SetupFormData(ByRef frm As frmFilter01) As Boolean
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim targetCell As Range: Set targetCell = ActiveCell ' 현재 포커스된 셀 기준
    Dim lo As ListObject: Set lo = targetCell.ListObject
    Dim filterRng As Range
    Dim relCol As Long

    ' 기본값은 실패(False)로 설정
    SetupFormData = False

    ' 1. 우선순위: 표(Table) 내부인지 확인
    If Not lo Is Nothing Then
        If lo.DataBodyRange Is Nothing Then Exit Function ' 빈 표 방어
        relCol = targetCell.Column - lo.Range.Column + 1
        Set frm.InjectFilterArea = lo.Range
        Set frm.InjectBodyRange = lo.DataBodyRange.Columns(relCol)
        frm.InjectHeaderName = lo.HeaderRowRange.Cells(1, relCol).Value
        frm.InjectRelativeCol = relCol

        SetupFormData = True
        Exit Function
    End If

    ' 2. 차선책: 자동 필터(AutoFilter) 영역 내부인지 확인
    If ws.AutoFilterMode Then
        Set filterRng = ws.AutoFilter.Range
        ' 선택한 셀이 실제 필터 범위와 겹치는지 정밀 검사
        If Not Intersect(targetCell, filterRng) Is Nothing Then
            If filterRng.Rows.Count < 2 Then Exit Function ' 본문 없는 필터 방어
            relCol = targetCell.Column - filterRng.Column + 1
            Set frm.InjectFilterArea = filterRng
            ' 본문 = 필터 영역 자체에서 헤더 1행 제외 (영역 밖 End(xlUp) 재유도 금지:
            ' AutoFilter.Range의 경계가 실제 필터 대상의 진실이다)
            Set frm.InjectBodyRange = filterRng.Columns(relCol).Offset(1, 0) _
                                        .Resize(filterRng.Rows.Count - 1, 1)
            frm.InjectHeaderName = ws.Cells(filterRng.Row, targetCell.Column).Value
            frm.InjectRelativeCol = relCol

            SetupFormData = True
            Exit Function
        End If
    End If

    ' 3. [최종 낙하 지점] 위 조건에 모두 해당하지 않는 경우에만 메시지 출력
    AutoCloseMsgBox_Popup "표(Table) 또는 자동 필터 범위 내에서 실행해 주세요.", 1, "UnivFilter 알림"

End Function
