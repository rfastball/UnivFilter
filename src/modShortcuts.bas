Attribute VB_Name = "modShortcuts"
Option Explicit

' ==========================================================
' [Ctrl + Shift + Q] 퀵 핸들러 (Quick Mode)
' 수정: Hybrid/Error일 경우 폼을 띄우는 방식으로 우회
' ==========================================================
Public Sub UnivFilter_QuickHandler()
    Dim frm As New frmFilter01
    Dim sClip As String
    
    ' 1. 기본 준비
    If Not SetupFormData(frm) Then Unload frm: Exit Sub
    
    sClip = ClipboardAPI.GetClipboardText()
    If Trim(sClip) = "" Then
        AutoCloseMsgBox_Popup "클립보드가 비어 있습니다.", 1, "알림"
        Unload frm: Exit Sub
    End If

    ' 2. 데이터 분석 실행 (여기서 HYBRID면 폼 내부에서 이미 잠금(Lock)이 걸림!)
    frm.StartInitialization
    
    ' 3. 상태에 따른 분기 처리 (정책 단일 소스 = modTypePolicy)
    Select Case modTypePolicy.GetModePolicy(frm.AssetType).RequiresFormInQuickMode
        Case True
            ' [경로 A] 데이터가 지저분함 -> 폼을 띄워서 사용자에게 보여줌 (이미 잠겨있음)
            frm.txtFilterValues.text = sClip  ' 클립보드 내용 입력
            modUnivFilter.MoveFormToCell_API frm, ActiveCell '폼 띄우기 전 위치이동
            frm.Show vbModeless               ' 실행하지 않고 보여주기만 함
            
            ' (옵션) 팝업으로 상황 설명
            AutoCloseMsgBox_Popup "데이터 형식이 혼합(숫자/문자/에러)되어 있습니다." & vbCrLf & vbCrLf & _
                                  "▶ 텍스트 부분 일치 모드로 전환됩니다." & vbCrLf & _
                                  "(완전 일치 검색을 원하시면 서식을 '텍스트'로 통일하세요)", 3, "알림"
        Case Else
            ' [경로 B] 깔끔한 데이터 -> 기존대로 즉시 퀵 실행
            frm.RunQuickFilter sClip
            Unload frm
    End Select
End Sub

' ==========================================================
' [Ctrl + Shift + U] UI 핸들러 (Form Mode)
' 수정: 폼 열기 전 타입 확인 후 잠금 처리
' ==========================================================
Public Sub UnivFilter_UIHandler()
    Dim frm As New frmFilter01
    If SetupFormData(frm) Then
        frm.StartInitialization ' 여기서 HYBRID면 알아서 잠김
        modUnivFilter.MoveFormToCell_API frm, ActiveCell '폼 띄우기 전 위치이동
        frm.Show
    Else
        Unload frm
    End If
End Sub

' 현재 열 정보를 폼에 주입하는 헬퍼 [cite: 137]
' [Helper] 현재 열 정보를 폼에 주입하는 공통 로직 (오타 수정됨)
' ==========================================================
Private Function SetupFormData(ByRef frm As frmFilter01) As Boolean
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim targetCell As Range: Set targetCell = ActiveCell ' 현재 포커스된 셀 기준
    Dim lo As ListObject: Set lo = targetCell.ListObject
    Dim filterRng As Range
    Dim relCol As Long
    
    ' 기본값은 실패(False)로 설정
    SetupFormData = False
    
    ' 1. 우선순위: 표(Table) 내부인지 확인 [cite: 137]
    If Not lo Is Nothing Then
        relCol = targetCell.Column - lo.Range.Column + 1
        Set frm.InjectFilterArea = lo.Range
        Set frm.InjectBodyRange = lo.DataBodyRange.Columns(relCol)
        frm.InjectHeaderName = lo.HeaderRowRange.Cells(1, relCol).Value
        frm.InjectRelativeCol = relCol
        
        SetupFormData = True
        Exit Function ' 찾았으므로 즉시 종료
    End If
    
    ' 2. 차선책: 자동 필터(AutoFilter) 영역 내부인지 확인 [cite: 201]
    If ws.AutoFilterMode Then
        Set filterRng = ws.AutoFilter.Range
        ' 선택한 셀이 실제 필터 범위와 겹치는지 정밀 검사
        If Not Intersect(targetCell, filterRng) Is Nothing Then
            relCol = targetCell.Column - filterRng.Column + 1
            Set frm.InjectFilterArea = filterRng
            ' 헤더 행을 제외한 데이터 본문 범위 설정 [cite: 203]
            Set frm.InjectBodyRange = ws.Range(ws.Cells(filterRng.Row + 1, targetCell.Column), _
                                               ws.Cells(ws.Rows.Count, targetCell.Column).End(xlUp))
            frm.InjectHeaderName = ws.Cells(filterRng.Row, targetCell.Column).Value
            frm.InjectRelativeCol = relCol
            
            SetupFormData = True
            Exit Function ' 찾았으므로 즉시 종료
        End If
    End If
    
    ' 3. [최종 낙하 지점] 위 조건에 모두 해당하지 않는 경우에만 메시지 출력 [cite: 147]
    AutoCloseMsgBox_Popup "표(Table) 또는 자동 필터 범위 내에서 실행해 주세요.", 1, "UnivFilter 알림"
    
End Function
