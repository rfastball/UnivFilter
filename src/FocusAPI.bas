Attribute VB_Name = "FocusAPI"
' [윈도우 제어: 창 찾기 및 포커스]
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long

Private Const GW_HWNDNEXT As Long = 2
Private Const SW_RESTORE As Long = 9

' [윈도우 포커스] 제목 일부만 일치해도 창을 찾아 활성화
Public Function FocusWindowByPartialTitle(partialTitle As String) As Boolean
    Dim hwnd As LongPtr
    Dim winTitle As String
    Dim lenTitle As Long
    
    ' 현재 실행 중인 첫 번째 윈도우 핸들
    hwnd = FindWindow(vbNullString, vbNullString)
    
    Do While hwnd <> 0
        lenTitle = GetWindowTextLength(hwnd)
        If lenTitle > 0 Then
            winTitle = String(lenTitle, vbNullChar)
            GetWindowText hwnd, winTitle, lenTitle + 1
            winTitle = Left(winTitle, lenTitle)
            
            ' 부분 일치 검사
            If InStr(1, winTitle, partialTitle, vbTextCompare) > 0 Then
                ' 최소화 상태면 복구
                If IsIconic(hwnd) Then
                    ShowWindow hwnd, SW_RESTORE
                Else
                    SetForegroundWindow hwnd
                End If
                FocusWindowByPartialTitle = True
                Exit Function
            End If
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
    
    FocusWindowByPartialTitle = False
End Function
