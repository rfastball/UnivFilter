Attribute VB_Name = "AutoCloseMsgAPI"
' @desc 자동 닫힘 팝업 표시 담당
' ==============================================================================
' [외부 API 및 상수 선언]
' ==============================================================================
' 자동 종료 기능이 포함된 Windows 표준 메시지 박스 API
Private Declare PtrSafe Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" ( _
    ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, _
    ByVal wType As Long, ByVal wLanguageId As Integer, ByVal dwMilliseconds As Long) As Long
    
' 메시지 박스 스타일 상수
Private Const MB_ICONINFO As Long = 64
Private Const MB_TOPMOST As Long = &H40000
Private Const MB_SETFOREGROUND As Long = &H10000

    
' ------------------------------------------------------------------------------
' [Utility] 자동 닫힘 메시지 팝업
' ------------------------------------------------------------------------------
Public Sub AutoCloseMsgBox_Popup(msg As String, Optional sec As Long = 1, Optional title As String = "알림")
    MessageBoxTimeout 0, msg, title, MB_ICONINFO + MB_TOPMOST + MB_SETFOREGROUND, 0, sec * 1000
End Sub
    
