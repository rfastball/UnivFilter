Attribute VB_Name = "KeyDownAPI"
' @desc 키 눌림 상태 조회 담당
' ==============================================================================
' [Windows API] 키 상태 감지
' ==============================================================================
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' 가독성을 위한 가상 키 코드 상수 (필요시 추가)
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12    ' Alt Key
Public Const VK_SHIFT = &H10

' [키 상태] 특정 키가 눌려있는지 확인 (GetAsyncKeyState wrapper)
Public Function IsKeyDown(ByVal vKey As Long) As Boolean
    IsKeyDown = (GetAsyncKeyState(vKey) And &H8000) <> 0
End Function
