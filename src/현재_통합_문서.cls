VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' ==========================================================
' [ThisWorkbook] UnivFilter 전역 감시 및 설정
' ==========================================================
' @desc 워크북 단축키/메뉴 바인딩 담당
' @depends modCache, modUnivFilter
Option Explicit

' 캐시 무효화는 이벤트가 아니라 지문 관측이 담당한다(modCache). 앱 이벤트 감시 없음.

Private Sub Workbook_Open()
    ' 단축키 바인딩
    Application.OnKey "^+Q", "UnivFilter_QuickHandler"
    Application.OnKey "^+U", "UnivFilter_UIHandler"
    
    SetupContextMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 1. 단축키 및 메뉴 해제
    Application.OnKey "^+Q"
    Application.OnKey "^+U"
    RemoveContextMenu

    ' 2. 메모리 정리
    modCache.ClearCache ' 종료 시 전체 청소
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
