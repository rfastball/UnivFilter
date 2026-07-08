Attribute VB_Name = "Test_TypePolicy"
' @desc 타입 정책 진리표 단위테스트 담당
' @depends modTypePolicy, LogShim
Option Explicit

' D5 진리표의 특성화 고정: 8 타입 x 3 필드.
' ROUNDED 비대칭 포함 현행 동작 그대로 - assert "개선" 금지.

Private Sub AssertPolicy(ByVal sType As String, _
                         ByVal bForceText As Boolean, _
                         ByVal sLockTag As String, _
                         ByVal bQuickForm As Boolean)
    Dim p As ModePolicy
    p = modTypePolicy.GetModePolicy(sType)
    LogShim.AssertEq "policy_" & sType & "_forcetext", bForceText, p.ForceTextStrategy
    LogShim.AssertEq "policy_" & sType & "_locktag", sLockTag, p.LockTag
    LogShim.AssertEq "policy_" & sType & "_quickform", bQuickForm, p.RequiresFormInQuickMode
End Sub

' @desc 8개 타입의 정책 3필드가 진리표와 일치하는지 검증함
Public Sub Test_Policy_TruthTable()
    AssertPolicy "NUMBER", False, "", False
    AssertPolicy "PERCENT", False, "", False
    AssertPolicy "STRING", False, "", False

    ' 비대칭: UI 잠금만 (엔진 강제/퀵 우회 없음)
    AssertPolicy "ROUNDED", False, "ROUNDING", False
    AssertPolicy "PERCENT_ROUNDED", False, "ROUNDING", False

    AssertPolicy "DATE", True, "DATE", True
    AssertPolicy "HYBRID", True, "HYBRID", True
    AssertPolicy "ERROR", True, "HYBRID", True
End Sub
