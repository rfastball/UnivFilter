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
' @desc 필터 조건 입력 폼 UI와 위임 담당
' @depends modCache, modProcessManager, modTypePolicy, FocusAPI, AutoCloseMsgAPI, clsColumnAsset, modUnivFilter
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

' 내부 데이터 처리 객체 및 범위
Private m_rngBody As Range           ' 데이터 본문 범위 (분석 대상)
Private m_rngFilterArea As Range     ' 필터가 실제로 적용될 영역
Private m_Asset As clsColumnAsset    ' [신규 엔진] 열 속성 및 데이터 집합 관리 객체

' ==============================================================================
' [Data Injection] 외부 데이터 주입 프로퍼티
' ==============================================================================
' 분석용 본문 범위 주입
Public Property Set InjectBodyRange(ByVal rng As Range)
    Set m_rngBody = rng
    ' [LOG] 데이터 주입 확인
    Log "Property Set InjectBodyRange: " & rng.Address(False, False), lvlDebug
End Property

Public Property Get InjectBodyRange() As Range
    Set InjectBodyRange = m_rngBody
End Property

' 필터 적용 대상 범위 주입
Public Property Set InjectFilterArea(ByVal rng As Range)
    Set m_rngFilterArea = rng
    ' [LOG] 필터 영역 주입 확인
    Log "Property Set InjectFilterArea: " & rng.Address(False, False), lvlDebug
End Property

Public Property Get InjectFilterArea() As Range
    Set InjectFilterArea = m_rngFilterArea
End Property


' ----------------------------------------------------------
' 목적: 단축키 퀵모드 실행 (UI 조작 없이 로직만 수행)
' ----------------------------------------------------------
Public Sub RunQuickFilter(ByVal clipText As String)
    ' [Stage 8, D7] 비즈니스 로직은 modProcessManager로 이동 - 폼은 위임만 (표면 불변)
    Me.StartInitialization
    modProcessManager.ExecuteQuickFilter clipText, m_Asset, m_rngFilterArea, InjectRelativeCol
End Sub
' ==============================================================================
' [Asset Initialization] 데이터 자산화 및 고속 로드
' ==============================================================================
Public Sub StartInitialization()
    ' [Stage 8, D7] 자산 취득과 UI 잠금을 분할 (공개 시그니처 불변)
    If Not LoadColumnAsset() Then Exit Sub
    ApplyTypeLock
End Sub

' [내부 1/2] 캐시 매니저(modCache)에게 자산 요청
' 설명: 데이터 로드, 평탄화(에러 처리 포함), 분석, 캐싱까지
'       모두 modCache.GetColumnAsset이 처리한다.
Private Function LoadColumnAsset() As Boolean
    Log "StartInitialization: 데이터 로드 및 분석 시작", lvlInfo

    ' 데이터 범위가 지정되지 않은 경우 즉시 종료
    If m_rngBody Is Nothing Then
        Log "StartInitialization: m_rngBody가 Nothing이므로 중단", lvlWarning
        LoadColumnAsset = False
        Exit Function
    End If

    Set m_Asset = modCache.GetColumnAsset(m_rngBody)

    ' [LOG] 자산 수령 확인
    Log "StartInitialization: 자산 수령 완료 (Type: " & m_Asset.DetectedType & ", Count: " & m_Asset.data.Count & ")", lvlDebug
    LoadColumnAsset = True
End Function

' [내부 2/2] 데이터 타입에 따른 UI 강제 잠금 (정책 단일 소스 = modTypePolicy)
Private Sub ApplyTypeLock()
    Select Case modTypePolicy.GetModePolicy(m_Asset.DetectedType).LockTag
        Case "HYBRID"
            Log "StartInitialization: 정합성 이슈(" & m_Asset.DetectedType & ") -> Hybrid 잠금 실행", lvlWarning
            Me.EnforceHybridMode

        Case "ROUNDING"
            Log "StartInitialization: 반올림 이슈(" & m_Asset.DetectedType & ") -> Rounding 잠금 실행", lvlWarning
            Me.EnforceRoundingMode

        Case "DATE"
            Log "StartInitialization: 날짜 이슈(" & m_Asset.DetectedType & ") -> Date 잠금 실행", lvlWarning
            Me.EnforceDateMode

        Case Else
            ' 정상 상태: 아무것도 하지 않음 (기본값 유지)
    End Select
End Sub

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
    ' 1. [기본 동작] 헤더 이름 표시 (무조건 표시)
    If InjectHeaderName <> "" Then
        Dim displayHeader As String
        displayHeader = IIf(Len(InjectHeaderName) > 10, Left(InjectHeaderName, 8) & "...", InjectHeaderName)
        lblInfo.Caption = "[" & displayHeader & "]"
        lblInfo.ControlTipText = InjectHeaderName
    Else
        lblInfo.Caption = "로드 실패"
    End If
    
    ' 2. [스타일링] 잠금 상태면 헤더 색상만 '경고색'으로 변경 (내용은 유지)
    If lblMatchSwitch.Enabled = False Then
        lblInfo.ForeColor = &H80 ' (주황빛 나는 빨강) - 평소와 다름을 암시
    Else
        lblInfo.ForeColor = &H80000012 ' (기본 검정)
    End If
End Sub

Private Sub UserForm_Terminate()
    Log "UserForm_Terminate: 폼 종료 프로세스 시작", lvlInfo

    ' 폼 종료 시 메모리 누수 방지를 위해 객체 변수 명시적 해제
    On Error Resume Next
    If Not G_OpenForms Is Nothing Then G_OpenForms.Remove MyFormKey
    On Error GoTo 0
    
    Set m_Asset = Nothing
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
' [Core Engine] 필터링 적용 로직 (Controller 역할)
' ==============================================================================
Private Sub btnApply_Click()
    Log "btnApply_Click: 적용 버튼 클릭됨", lvlInfo
    On Error GoTo ErrorHandler

    ' 1. 유효성 검사 (View Validation)
    If m_rngBody Is Nothing Or m_rngFilterArea Is Nothing Then
        Log "btnApply_Click: 필수 객체 누락", lvlError
        MsgBox "대상 정보가 누락되었습니다. (Asset 초기화 오류)", vbCritical: Unload Me: Exit Sub
    End If
    
    ' 2. 입력값 수집 (View -> Data)
    Dim sInputTxt As String: sInputTxt = Trim(txtFilterValues.text)
    
    If sInputTxt = "" Then
        AutoCloseMsgBox_Popup "검색할 내용을 입력해주세요!", 1, "알림"
        txtFilterValues.SetFocus
        Exit Sub
    End If

    ' 3. UI 피드백 (Wait State)
    Dim oldCaption As String: oldCaption = lblInfo.Caption
    lblInfo.Caption = "처리 중..."
    Application.Cursor = xlWait
    Me.Repaint
    
    ' ==========================================================================
    ' [PROCESS] 오케스트레이터 호출 (위임)
    ' ==========================================================================
    Dim resultDict As Object
    
    ' 모듈의 함수 호출 (순수 로직 실행)
    Set resultDict = modProcessManager.ExecuteFilterProcess( _
                        sInputTxt, _
                        m_Asset, _
                        m_rngFilterArea, _
                        InjectRelativeCol, _
                        pIsExcludeMode, _
                        pIsStackMode, _
                        pIsPartialMatch)
    
    ' ==========================================================================
    
    Application.Cursor = xlDefault
    lblInfo.Caption = oldCaption
    
    ' 4. 결과 처리 (Process Result Handling)
    If resultDict("IsSuccess") Then
        Log "Process Success: " & resultDict("LogMessage"), lvlInfo
        
        ' 4-1. 누락 데이터 사후 처리 (UI 피드백 포함)
        ' 제외 모드가 아닐 때만 누락 데이터를 체크합니다.
        If Not pIsExcludeMode Then
            Dim sMissingMsg As String
            sMissingMsg = modProcessManager.ProcessMissingData(resultDict("MissingData"))
            
            If sMissingMsg <> "" Then
                ' 누락 발생 시 팝업 (클립보드 복사 여부 포함)
                AutoCloseMsgBox_Popup sMissingMsg, 2
            End If
        End If
        
        Unload Me
    Else
        ' 4-2. 실패 처리
        If resultDict.exists("ErrorNumber") Then
            ' 시스템 에러
             MsgBox "오류 발생: " & resultDict("ErrorMsg"), vbCritical
        Else
            ' 로직 중단 (Nothing 반환 등)
            Log "Process Halted: " & resultDict("ErrorMsg"), lvlWarning
            Unload Me
        End If
    End If

    Exit Sub
    
ErrorHandler:
    Application.Cursor = xlDefault
    ' [UI] 사용자 정의 에러 처리
    If Err.Number = vbObjectError + 1000 Then
        AutoCloseMsgBox_Popup Err.Description, 3, "입력 데이터 확인"
        txtFilterValues.SetFocus
    Else
        MsgBox "Critical Error: " & Err.Description, vbCritical
    End If
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
    
    ' [잠금 상태 전용] 회색 스타일
    Const COL_LOCKED_BG As Long = &HE0E0E0      ' 연한 회색 (버튼 배경)
    Const COL_LOCKED_TXT As Long = &H606060     ' 진한 회색 (글자)

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

    ' ==========================================================================
    ' 3. 매칭 모드 체크박스 (Tag 기반 스타일링)
    ' ==========================================================================
    With lblMatchSwitch
        ' [잠금 상태 통합 처리]
        ' 태그에 "LOCKED"가 포함되면 (LOCKED, LOCKED_ROUNDING 모두 해당) -> 회색 처리
        If InStr(1, .tag, "LOCKED") > 0 Then
            .Caption = ChrW(&H2714)       ' 체크 표시 (?)
            .BackColor = COL_LOCKED_BG    ' 회색 배경
            .BorderColor = &H808080
            .BorderStyle = fmBorderStyleNone
            
            ' [텍스트 스타일도 잠금 상태로 통일]
            lblMatchText.ForeColor = COL_LOCKED_TXT
            lblMatchText.Font.Bold = True
            
        ' [정상 상태: 활성화]
        ElseIf pIsPartialMatch Then
            .Caption = ChrW(&H2714)
            .BackColor = COL_MATCH_ON
            .ForeColor = vbWhite
            .BorderStyle = fmBorderStyleNone
            
            lblMatchText.ForeColor = &H0  ' 검정
            lblMatchText.Font.Bold = True
            
        ' [정상 상태: 비활성화]
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
' [Helper] 외부 제어용 속성 및 메서드 (새로 추가됨)
' ==============================================================================

' 1. [잠금 실행 A] Hybrid/Error 모드 (기존)
Public Sub EnforceHybridMode()
    pIsPartialMatch = True
    
    LockControls "LOCKED", "텍스트 포함 검색"
    
    ' 툴팁 설정
    lblMatchText.ControlTipText = "데이터 형식이 혼재되어 [텍스트 검색]으로 고정됩니다."
End Sub

' 2. [잠금 실행 B] Rounding 모드 (신규)
Public Sub EnforceRoundingMode()
    pIsPartialMatch = True ' 텍스트 포함 검색 모드 활성화 (화면에 보이는 값 기준 검색)
    
    LockControls "LOCKED_ROUNDING", "텍스트 포함 검색"
    
    ' 툴팁 설정 (반올림 관련 안내)
    lblMatchText.ControlTipText = "표시되는 값과 실제 정밀도가 달라 [텍스트 검색]으로 고정됩니다."
End Sub

' 3. [잠금 실행 C] Date 모드 (신규)
Public Sub EnforceDateMode()
    pIsPartialMatch = True ' 텍스트 포함 검색 모드 활성화 (화면에 보이는 값 기준 검색)
    
    LockControls "LOCKED_DATE", "텍스트 포함 검색"
    
    ' 툴팁 설정 (반올림 관련 안내)
    lblMatchText.ControlTipText = "날짜 데이터의 정확한 매칭을 위해 [텍스트 검색]으로 고정됩니다."
End Sub
' [공통] 컨트롤 잠금 처리 헬퍼
Private Sub LockControls(sTag As String, sCaption As String)
    ' 스위치 잠금
    With lblMatchSwitch
        .Enabled = True ' 클릭 이벤트를 위해 True 유지
        .tag = sTag
        .MousePointer = 12 ' No Drop 커서
    End With
    
    ' 텍스트 라벨 잠금
    With lblMatchText
        .Enabled = True
        .tag = sTag
        .MousePointer = 12
        .Caption = sCaption
        .Font.Size = 9
    End With
    
    UpdateUI
End Sub


' 2. [속성 개방] 외부(Shortcuts 모듈 등)에서 자산 타입을 확인할 수 있도록 속성 추가
Public Property Get AssetType() As String
    If m_Asset Is Nothing Then
        AssetType = ""
    Else
        AssetType = m_Asset.ColumnType
    End If
End Property

' [Helper] 잠금 상태 클릭 시 보여줄 안내 팝업
Private Sub ShowHybridGuideMsg()
    ' 상황: 서식이 제각각이거나, 데이터 타입이 섞여 있음
    ' 해결: 무엇이든 좋으니 '하나의 서식'으로 통일하면 됨
    
    AutoCloseMsgBox_Popup "데이터의 서식이 서로 다르거나 형식이 혼재되어" & vbCrLf & _
                          "완전일치 검색을 사용할 수 없습니다." & vbCrLf & vbCrLf & _
                          "Tip: 해당 열의 서식을 하나로 통일하면 해제됩니다." & vbCrLf & _
                          "(예: 전체 텍스트 지정 또는 동일한 표시 형식 적용)", 4, "잠금 알림"
End Sub

' [Helper] Rounding 잠금 알림 (신규)
Private Sub ShowRoundingGuideMsg()
    AutoCloseMsgBox_Popup "반올림 또는 백분율 서식이 적용된 데이터입니다." & vbCrLf & _
                          "화면에 보이는 값과 실제 값이 다를 수 있어" & vbCrLf & _
                          "안전을 위해 [텍스트 포함 검색]으로 고정됩니다." & vbCrLf & vbCrLf & _
                          "Tip: 서식을 '일반'이나 '숫자'로 풀면 해제됩니다.", 4, "서식 잠금 알림"
End Sub
' [Helper] Rounding 잠금 알림 (신규)
Private Sub ShowDateGuideMsg()
    AutoCloseMsgBox_Popup "날짜 및 서식이 포함된 열입니다." & vbCrLf & _
                          "정확한 매칭을 위해 [텍스트 포함 검색]으로 고정됩니다." & vbCrLf & vbCrLf & _
                          "Tip: 화면에 보이는 그대로 입력하시면" & vbCrLf & _
                          "가장 쉽고 빠르게 필터링됩니다!", 4, "데이터 검색 안내"
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
    ' 1. 잠금 상태 확인 (태그로 분기)
    If lblMatchSwitch.tag = "LOCKED" Then
        ShowHybridGuideMsg ' Hybrid/Error 전용 메시지
        Exit Sub
    ElseIf lblMatchSwitch.tag = "LOCKED_ROUNDING" Then
        ShowRoundingGuideMsg ' Rounded 전용 메시지 (NEW)
        Exit Sub
    ElseIf lblMatchSwitch.tag = "LOCKED_DATE" Then
    ShowDateGuideMsg ' DATE 전용 메시지 (NEW)
    Exit Sub
    
    End If
    
    ' 2. 정상 토글 동작
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
        If InStr(1, lblMatchSwitch.tag, "LOCKED") = 0 Then
            lblMatchSwitch_Click
        End If
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

' 1. 텍스트 박스에서 키를 눌렀을 때 (여기가 제일 중요합니다!)
Private Sub txtFilterValues_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RunGlobalShortcuts(KeyCode, Shift)
End Sub

' 2. (혹시 모르니) 폼 배경에서 키를 눌렀을 때
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RunGlobalShortcuts(KeyCode, Shift)
End Sub
