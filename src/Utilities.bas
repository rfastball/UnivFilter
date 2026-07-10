Attribute VB_Name = "Utilities"
' @desc 화면 갱신/계산 일시 최적화 토글 담당
Option Explicit

' 필터 파이프라인에 필요한 3속성만 저장-복원한다.
' 구판은 PrintCommunication + 전 시트 DisplayPageBreaks까지 토글해 왕복 47.7ms를
' 매 실행 지불했다(ProbeFixedOverheads 실측). 이 엔진은 읽기 + AutoFilter 적용만
' 하므로 나머지 속성은 무관하고, DisplayStatusBar 끄기는 상태바 리포트 UX와 상충.
Private originalCalcMode As XlCalculation
Private originalScreenUpdating As Boolean
Private originalEnableEvents As Boolean
Private isOptimized As Boolean

' @desc 화면갱신/이벤트/계산 3속성을 일시 최적화하거나 원상 복원함
Public Sub OptimizedMode(ByVal enable As Boolean)
    ' 이미 같은 상태면 중복 실행 방지
    If isOptimized = enable Then Exit Sub

    If enable Then
        originalCalcMode = Application.Calculation
        originalScreenUpdating = Application.ScreenUpdating
        originalEnableEvents = Application.EnableEvents

        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    Else
        Application.Calculation = originalCalcMode
        Application.ScreenUpdating = originalScreenUpdating
        Application.EnableEvents = originalEnableEvents
    End If

    isOptimized = enable
End Sub
