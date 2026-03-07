VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==========================================================
' [ThisWorkbook] UnivFilter 전역 감시 및 설정
' ==========================================================
Option Explicit

' 엑셀 어플리케이션 전체 이벤트를 받기 위한 변수
Private WithEvents AppEvents As Application
Attribute AppEvents.VB_VarHelpID = -1

Private Sub Workbook_Open()
    ' 1. 앱 레벨 이벤트 연결 (엑셀 전체 감시 시작)
    Set AppEvents = Application
    
    ' 2. 단축키 바인딩
    Application.OnKey "^+Q", "UnivFilter_QuickHandler"
    Application.OnKey "^+U", "UnivFilter_UIHandler"
    
    SetupContextMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 1. 단축키 및 메뉴 해제
    Application.OnKey "^+Q"
    Application.OnKey "^+U"
    RemoveContextMenu
    
    ' 2. 감시 해제 및 메모리 정리
    Set AppEvents = Nothing
    modCache.InvalidateCache ' 종료 시에는 전체 초기화
End Sub

Private Sub Workbook_AddinInstall()
    SetupContextMenu
End Sub

Private Sub Workbook_AddinUninstall()
    RemoveContextMenu
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    SetupContextMenu
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    RemoveContextMenu
End Sub

' ==========================================================
' [Core] 전역 데이터 변경 감지 로직
' ==========================================================

' 1. 값 변경 감지 (Cell Value Change)
Private Sub AppEvents_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Add-in 파일 내부의 변경은 무시 (무한 루프 방지)
    If Sh.Parent Is ThisWorkbook Then Exit Sub
    
    ' 변경된 시트(Sh)에 해당하는 캐시만 부분 삭제 요청
    modCache.InvalidateCacheBySheet Sh
End Sub

' 2. 수식 재계산 감지 (Formula Recalculation)
Private Sub AppEvents_SheetCalculate(ByVal Sh As Object)
    ' Add-in 파일 내부의 변경은 무시
    If Sh.Parent Is ThisWorkbook Then Exit Sub
    
    ' 재계산된 시트(Sh)에 해당하는 캐시만 부분 삭제 요청
    modCache.InvalidateCacheBySheet Sh
End Sub
