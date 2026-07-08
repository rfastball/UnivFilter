Attribute VB_Name = "Utilities"
Option Explicit


' 최적화 모드 원래 설정값을 저장할 변수
Private originalCalcMode As XlCalculation
Private originalScreenUpdating As Boolean
Private originalEnableEvents As Boolean
Private originalStatusBar As Boolean
Private originalEnableAnimations As Boolean
Private originalPrintCommunication As Boolean
Private originalPageBreaks As Object
Private isOptimized As Boolean


' VBA 코드 최적화 모드 설정/해제 프로시저
Public Sub OptimizedMode(ByVal enable As Boolean)
    Dim ws As Worksheet
    Dim wsRestore As Worksheet

    ' 이미 같은 상태면 중복 실행 방지
    If isOptimized = enable Then Exit Sub
    
    If enable Then
        ' 현재 설정 저장
        originalCalcMode = Application.Calculation
        originalScreenUpdating = Application.ScreenUpdating
        originalEnableEvents = Application.EnableEvents
        originalStatusBar = Application.DisplayStatusBar
        originalEnableAnimations = Application.EnableAnimations
        originalPrintCommunication = Application.PrintCommunication
        Set originalPageBreaks = CreateObject("Scripting.Dictionary")
        
        ' 최적화 설정 적용
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
        Application.EnableAnimations = False
        Application.PrintCommunication = False
        
        ' 모든 워크시트의 페이지 나누기 표시 끄기
        For Each ws In ActiveWorkbook.Worksheets
            ' [Fix v2.7] 시트별 DisplayPageBreaks 원래 값 저장
            If Not originalPageBreaks.Exists(ws.Name) Then originalPageBreaks.Add ws.Name, ws.DisplayPageBreaks
            ws.DisplayPageBreaks = False
        Next ws
    Else
        ' 원래 설정으로 복원
        Application.Calculation = originalCalcMode
        Application.ScreenUpdating = originalScreenUpdating
        Application.EnableEvents = originalEnableEvents
        Application.DisplayStatusBar = originalStatusBar
        Application.EnableAnimations = originalEnableAnimations
        Application.PrintCommunication = originalPrintCommunication
        
        If Not originalPageBreaks Is Nothing Then
            ' [Fix v2.7] 시트별 DisplayPageBreaks 원래 값 복원
            On Error Resume Next
            For Each wsRestore In ActiveWorkbook.Worksheets
                If originalPageBreaks.Exists(wsRestore.Name) Then
                    wsRestore.DisplayPageBreaks = originalPageBreaks(wsRestore.Name)
                End If
            Next wsRestore
            On Error GoTo 0
            Set originalPageBreaks = Nothing
        End If

    End If
    
    isOptimized = enable
End Sub
