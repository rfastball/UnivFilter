VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFilter01
   Caption         =   "필터설정"
   ClientHeight    =   3536
   ClientLeft      =   88
   ClientTop       =   424
   ClientWidth     =   4056
   OleObjectBlob   =   "frmFilter01.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmFilter01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc 필터 조건 입력 폼 UI와 위임 담당
' @depends modFilterRun, modReport, FocusAPI, AutoCloseMsgAPI, modUnivFilter
Option Explicit

' ==============================================================================
' [내부 변수 선언]
' ==============================================================================
' 폼 상태 제어 플래그
Private pIsExcludeMode As Boolean   ' True: 제외(Exclude), False: 포함(Include)
Private pIsStackMode As Boolean     ' True: 누적(Stack), False: 초기화(Reset)
Private pIsPartialMatch As Boolean  ' True: 부분일치, False: 완전일치

' 외부 호출부(Main)로부터 주입받는 변수
Public MyFormKey As String          ' 폼 인스턴스 관리를 위한 고유 식별자
Public InjectHeaderName As String   ' 선택된 열의 헤더 이름
Public InjectRelativeCol As Long    ' 필터 영역 내에서의 상대적 열 번호

' 내부 데이터 처리 범위 (자산은 파이프라인이 실행 시점에 캐시에서 취득)
Private m_rngBody As Range           ' 데이터 본문 범위 (분석 대상)
Private m_rngFilterArea As Range     ' 필터가 실제로 적용될 영역

' ==============================================================================
' [Data Injection] 외부 데이터 주입 프로퍼티
' ==============================================================================
' 분석용 본문 범위 주입
Public Property Set InjectBodyRange(ByVal rng As Range)
    Set m_rngBody = rng
    Log "Property Set InjectBodyRange: " & rng.Address(False, False), lvlDebug
End Property

Public Property Get InjectBodyRange() As Range
    Set InjectBodyRange = m_rngBody
End Property

' 필터 적용 대상 범위 주입
Public Property Set InjectFilterArea(ByVal rng As Range)
    Set m_rngFilterArea = rng
    Log "Property Set InjectFilterArea: " & rng.Address(False, False), lvlDebug
End Property

Public Property Get InjectFilterArea() As Range
    Set InjectFilterArea = m_rngFilterArea
End Property

' ==============================================================================
' [Form Life-Cycle] 폼 생명주기 관리
' ==============================================================================
Private Sub UserForm_Initialize()
    Log "UserForm_Initialize: 폼 로드 시작", lvlInfo

    ' 버튼 레이아웃 초기화 (Default/Cancel 버튼 설정)
    btnApply.Visible = True: btnApply.Left = 9000: btnApply.Default = True
    btnCancel.Visible = True: btnCancel.Left = 9000: btnCancel.Cancel = True

    lblInfo.Caption = "대기 중..."
    txtFilterValues.text = ""

    ' 전역 변수에 기록된 사용자의 마지막 작업 상태 복원
    pIsExcludeMode = G_LastExcludeMode
    pIsStackMode = G_LastStackMode
    pIsPartialMatch = False ' 초기 매칭 모드는 완전일치로 고정

    Log "UserForm_Initialize: 상태 복원 (Exclude: " & pIsExcludeMode & ", Stack: " & pIsStackMode & ")", lvlDebug

    ResetHoverEffects
    UpdateUI
End Sub

Private Sub UserForm_Activate()
    ' 헤더 이름 표시
    If InjectHeaderName <> "" Then
        Dim displayHeader As String
        displayHeader = IIf(Len(InjectHeaderName) > 10, Left(InjectHeaderName, 8) & "...", InjectHeaderName)
        lblInfo.Caption = "[" & displayHeader & "]"
        lblInfo.ControlTipText = InjectHeaderName
    Else
        lblInfo.Caption = "로드 실패"
    End If
End Sub

Private Sub UserForm_Terminate()
    Log "UserForm_Terminate: 폼 종료 프로세스 시작", lvlInfo

    ' 폼 종료 시 메모리 누수 방지를 위해 객체 변수 명시적 해제
    On Error Resume Next
    If Not G_OpenForms Is Nothing Then G_OpenForms.Remove MyFormKey
    On Error GoTo 0

    Set m_rngBody = Nothing
    Set m_rngFilterArea = Nothing

    ' 폼이 닫힌 후 워크시트나 다른 폼으로 포커스 복원 (Focus API 활용)
    Dim frm As Object
    For Each frm In VBA.UserForms
        If Not frm Is Me Then
            If frm.Visible Then
                FocusAPI.FocusWindowByPartialTitle frm.Caption
                Exit For
            End If
        End If
    Next frm
End Sub

' ==============================================================================
' [Controller] 적용 버튼 -> 파이프라인 위임
' ==============================================================================
Private Sub btnApply_Click()
    Log "btnApply_Click: 적용 버튼 클릭됨", lvlInfo
    On Error GoTo ErrorHandler

    ' 1. 유효성 검사 (View Validation)
    If m_rngBody Is Nothing Or m_rngFilterArea Is Nothing Then
        Log "btnApply_Click: 필수 객체 누락", lvlError
        MsgBox "대상 정보가 누락되었습니다.", vbCritical: Unload Me: Exit Sub
    End If

    ' 2. 입력값 수집 (View -> Data)
    Dim sInputTxt As String: sInputTxt = Trim(txtFilterValues.text)

    If sInputTxt = "" Then
        AutoCloseMsgBox_Popup "검색할 내용을 입력해주세요!", 1, "알림"
        txtFilterValues.SetFocus
        Exit Sub
    End If

    ' 3. UI 피드백: 전역 커서는 강제 중단 시 복원되지 않으므로 변경하지 않는다.
    Dim oldCaption As String: oldCaption = lblInfo.Caption
    lblInfo.Caption = "처리 중..."
    Me.Repaint

    ' 4. 파이프라인 실행 (위임)
    Dim result As Object
    Set result = modFilterRun.RunFilter( _
                        m_rngBody, _
                        m_rngFilterArea, _
                        InjectRelativeCol, _
                        sInputTxt, _
                        pIsExcludeMode, _
                        pIsStackMode, _
                        pIsPartialMatch)

    lblInfo.Caption = oldCaption

    ' 5. 결과 처리: 팝업은 실행 가능 정보만, 자동보정은 상태바 한 줄 (퀵 모드와 동일)
    If result("Ok") Then
        Log "RunFilter 성공: " & result("StatusMessage"), lvlInfo
        modReport.ShowStatusBar modReport.StatusLine(result)

        If Len(result("WarningMessage")) > 0 Then
            AutoCloseMsgBox_Popup result("WarningMessage"), 3, "경고"
        End If
        If Len(result("SelfCheckNote")) > 0 Then
            AutoCloseMsgBox_Popup result("SelfCheckNote"), 3, "검증 경고"
        End If
        If Len(result("ScanNote")) > 0 Then
            AutoCloseMsgBox_Popup result("ScanNote"), 3, "스캔 안내"
        End If
        If Len(result("MissingMessage")) > 0 Then
            AutoCloseMsgBox_Popup result("MissingMessage"), 2
        End If

        Unload Me
    Else
        If Len(result("ErrMsg")) > 0 Then
            Log "RunFilter 실패: " & result("ErrMsg"), lvlWarning
            AutoCloseMsgBox_Popup result("ErrMsg"), 3, "알림"
            txtFilterValues.SetFocus
        Else
            Unload Me
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Critical Error: " & Err.Description, vbCritical
End Sub

Private Sub btnCancel_Click()
    Log "btnCancel_Click: 취소 버튼 클릭", lvlInfo
    Unload Me
End Sub

' ==============================================================================
' [UI Logic] 상태 업데이트 및 스타일 제어
' ==============================================================================
Private Sub UpdateUI()
    ' 스타일 상수 정의 (배경색, 텍스트색, 테두리)
    Const COL_BG_IDLE As Long = &HF3F3F3
    Const COL_TXT_IDLE As Long = &H808080
    Const COL_BORDER_IDLE As Long = &HD0D0D0

    Const COL_GREEN As Long = &H467321
    Const COL_RED As Long = &HC0
    Const COL_RESET As Long = &H947556
    Const COL_STACK As Long = &H503E2C
    Const COL_MATCH_ON As Long = &H805000

    ' 1. 모드 표시 업데이트 (포함/제외)
    If pIsExcludeMode Then
        SetLabelStyle lblModeIn, COL_BG_IDLE, COL_TXT_IDLE, False, COL_BORDER_IDLE
        SetLabelStyle lblModeOut, COL_RED, vbWhite, True, COL_RED
    Else
        SetLabelStyle lblModeIn, COL_GREEN, vbWhite, True, COL_GREEN
        SetLabelStyle lblModeOut, COL_BG_IDLE, COL_TXT_IDLE, False, COL_BORDER_IDLE
    End If

    ' 2. 필터 방식 표시 업데이트 (초기화/누적)
    If pIsStackMode Then
        SetLabelStyle lblReset, COL_BG_IDLE, COL_TXT_IDLE, False, COL_BORDER_IDLE
        SetLabelStyle lblStack, COL_STACK, vbWhite, True, COL_STACK
    Else
        SetLabelStyle lblReset, COL_RESET, vbWhite, True, COL_RESET
        SetLabelStyle lblStack, COL_BG_IDLE, COL_TXT_IDLE, False, COL_BORDER_IDLE
    End If

    ' 3. 매칭 모드 체크박스 (표시 문자열 엔진에서는 어떤 열이든 자유 토글)
    With lblMatchSwitch
        If pIsPartialMatch Then
            .Caption = ChrW(&H2714)
            .BackColor = COL_MATCH_ON
            .ForeColor = vbWhite
            .BorderStyle = fmBorderStyleNone

            lblMatchText.ForeColor = &H0  ' 검정
            lblMatchText.Font.Bold = True
        Else
            .Caption = ""
            .BackColor = COL_BG_IDLE
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = COL_BORDER_IDLE

            lblMatchText.ForeColor = COL_TXT_IDLE ' 연한 회색
            lblMatchText.Font.Bold = False
        End If
    End With
End Sub

' 공용 스타일 적용 헬퍼
Private Sub SetLabelStyle(lbl As MSForms.Label, bgCol As Long, fontCol As Long, isBold As Boolean, borderCol As Long)
    With lbl
        .BackColor = bgCol: .ForeColor = fontCol: .Font.Bold = isBold
        .BorderStyle = fmBorderStyleSingle: .BorderColor = borderCol
    End With
End Sub

' ==============================================================================
' [Event Handlers] 클릭 및 마우스 인터랙션
' ==============================================================================
Private Sub lblModeIn_Click():  pIsExcludeMode = False: G_LastExcludeMode = False: UpdateUI: Log "Mode Change: Include", lvlDebug: End Sub
Private Sub lblModeOut_Click(): pIsExcludeMode = True:  G_LastExcludeMode = True:  UpdateUI: Log "Mode Change: Exclude", lvlDebug: End Sub
Private Sub lblReset_Click():   pIsStackMode = False:   G_LastStackMode = False:   UpdateUI: Log "Mode Change: Reset", lvlDebug: End Sub
Private Sub lblStack_Click():   pIsStackMode = True:    G_LastStackMode = True:    UpdateUI: Log "Mode Change: Stack", lvlDebug: End Sub

' 매칭 모드 스위치 클릭
Private Sub lblMatchSwitch_Click()
    pIsPartialMatch = Not pIsPartialMatch
    UpdateUI
End Sub

' 매칭 모드 텍스트 라벨 클릭 (스위치와 동일 동작)
Private Sub lblMatchText_Click()
    lblMatchSwitch_Click
End Sub

' 적용/취소 가상 버튼 클릭 처리
Private Sub lblApply_Click(): Call btnApply_Click: End Sub
Private Sub lblCancel_Click(): Call btnCancel_Click: End Sub

' [UI 호버 효과] 마우스가 폼 위를 떠다닐 때 스타일 초기화
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    ' MouseMove는 로그를 남기지 않음 (로그 폭주 방지)
    ResetHoverEffects
End Sub

' 모든 호버 스타일 초기화 (기본 상태로 복구)
Private Sub ResetHoverEffects()
    lblApply.BackColor = &H606060: lblApply.ForeColor = vbWhite
    lblCancel.BackColor = &H8000000F: lblCancel.ForeColor = &H505050
    UpdateUI
End Sub

' [개별 호버 이벤트] 마우스 오버 시 시각적 반응
Private Sub lblModeIn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If pIsExcludeMode Then lblModeIn.BackColor = &HE0E0E0
End Sub
Private Sub lblModeOut_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Not pIsExcludeMode Then lblModeOut.BackColor = &HE0E0E0
End Sub
Private Sub lblReset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If pIsStackMode Then lblReset.BackColor = &HE0E0E0
End Sub
Private Sub lblStack_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Not pIsStackMode Then lblStack.BackColor = &HE0E0E0
End Sub
Private Sub lblMatchSwitch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Not pIsPartialMatch Then lblMatchSwitch.BackColor = &HE0E0E0
End Sub
Private Sub lblMatchText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Not pIsPartialMatch Then lblMatchSwitch.BackColor = &HE0E0E0
End Sub
Private Sub lblApply_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    lblApply.BackColor = &H303030
End Sub
Private Sub lblCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    lblCancel.BackColor = &HE0E0E0
End Sub

' ==============================================================================
' [Shared Logic] 단축키 공용 처리기 (리모컨)
' ==============================================================================
Private Sub RunGlobalShortcuts(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    ' -----------------------------------------------------------
    ' 1. [Tab] 부분일치 토글
    ' -----------------------------------------------------------
    If KeyCode = vbKeyTab Then
        lblMatchSwitch_Click
        KeyCode = 0 ' 포커스 이동 방지
        Exit Sub
    End If

    ' -----------------------------------------------------------
    ' 2. [F1 ~ F4] 필터 모드 변경 / [F5] 실행
    ' -----------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1: pIsExcludeMode = False: UpdateUI: KeyCode = 0
        Case vbKeyF2: pIsExcludeMode = True:  UpdateUI: KeyCode = 0
        Case vbKeyF3: pIsStackMode = False:   UpdateUI: KeyCode = 0
        Case vbKeyF4: pIsStackMode = True:    UpdateUI: KeyCode = 0
        Case vbKeyF5: Call btnApply_Click:              KeyCode = 0
    End Select

End Sub

' 1. 텍스트 박스에서 키를 눌렀을 때
Private Sub txtFilterValues_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RunGlobalShortcuts(KeyCode, Shift)
End Sub

' 2. 폼 배경에서 키를 눌렀을 때
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RunGlobalShortcuts(KeyCode, Shift)
End Sub
