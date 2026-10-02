Attribute VB_Name = "modShortcuts"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc 전역 단축키 핸들러와 폼 주입 담당
' @depends modUnivFilter, modFilterRun, modReport, ClipboardAPI, AutoCloseMsgAPI
Option Explicit

' ==========================================================
' [Ctrl + Shift + Q] 퀵 핸들러 (Quick Mode)
' 항상 즉시 실행. 매칭 0건이면 아무것도 적용하지 않고 알림만.
' 폼 미사용: 대상 해석은 ResolveFilterTarget - 폼 인스턴스화 고정비(~16ms) 제거.
' ==========================================================
Public Sub UnivFilter_QuickHandler()
    Dim rngBody As Range, rngArea As Range
    Dim lField As Long, sHeader As String
    If Not modUnivFilter.ResolveFilterTarget(ActiveCell, rngBody, rngArea, lField, sHeader) Then
        AutoCloseMsgAPI.AutoCloseMsgBox_Popup "표(Table) 또는 자동 필터 범위 내에서 실행해 주세요.", 1, "UnivFilter 알림"
        Exit Sub
    End If

    Dim sClip As String
    sClip = ClipboardAPI.GetClipboardText()
    If Trim(sClip) = "" Then
        AutoCloseMsgBox_Popup "클립보드가 비어 있습니다.", 1, "알림"
        Exit Sub
    End If

    ' 2. 즉시 실행 (완전일치/포함/초기화 고정)
    Dim result As Object
    Set result = modFilterRun.RunQuickFilter(rngBody, rngArea, lField, sClip)

    ' 3. 결과 알림: 팝업은 실행 가능 정보(오류/0건/누락)만.
    ' 자동보정은 결정론 동작이라 결과가 곧 피드백 - 상태바 한 줄로만 알린다.
    If Not result("Ok") Then
        AutoCloseMsgBox_Popup result("ErrMsg"), 3, "오류"
    ElseIf result("QuickNoMatch") Then
        ' 0건: 필터/폼/클립보드 전부 원상태 유지
        AutoCloseMsgBox_Popup "일치하는 값이 없어 필터를 적용하지 않았습니다.", 2, "알림"
    Else
        modReport.ShowStatusBar modReport.StatusLine(result)
        If Len(result("WarningMessage")) > 0 Then
            AutoCloseMsgAPI.AutoCloseMsgBox_Popup result("WarningMessage"), 3, "경고"
        End If
        If Len(result("SelfCheckNote")) > 0 Then
            AutoCloseMsgAPI.AutoCloseMsgBox_Popup result("SelfCheckNote"), 3, "검증 경고"
        End If
        If Len(result("MissingMessage")) > 0 Then
            AutoCloseMsgBox_Popup result("MissingMessage"), 2
        End If
    End If

    ' 퀵 모드는 폼이 없어 무표정 멈춤이 생길 수 있어 최초 저속 스캔을 사후 안내
    If result("Ok") And Len(result("ScanNote")) > 0 Then
        AutoCloseMsgBox_Popup result("ScanNote"), 3, "스캔 안내"
    End If
End Sub

' @desc 활성 셀의 필터 설정 폼을 엶
Public Sub UnivFilter_UIHandler()
    modUnivFilter.OpenFilterForm ActiveCell
End Sub
