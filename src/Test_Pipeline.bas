Attribute VB_Name = "Test_Pipeline"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc modFilterRun 파이프라인 e2e 계약 테스트 담당
' @depends modFilterRun, modCache, modApply, modUnivFilter, modReport, frmFilter01, clsSet, Utilities, LogShim
Option Explicit

' 2026-07 재설계 스펙: 폼/퀵 공용 진입점의 결과 계약을 실기 시트로 고정한다.

Private Const SHEET_NAME As String = "UF_PipeTest"

' @desc 포함 모드 기본 흐름(매칭-적용-상태 메시지)을 검증함
Public Sub Test_PL_IncludeBasic()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    PutNumberCol ws

    Dim r As Object
    Set r = modFilterRun.RunFilter(Body(ws), Area(ws), 1, "1,000", False, False, False)

    LogShim.AssertEq "pl_inc_ok", True, r("Ok")
    LogShim.AssertEq "pl_inc_matched", CLng(1), CLng(r("MatchedCount"))
    LogShim.AssertEq "pl_inc_visible", CLng(1), modApply.CountVisibleRows(Area(ws))
    LogShim.AssertEq "pl_inc_selfcheck_quiet", "", CStr(r("SelfCheckNote"))
    ' 소규모 열은 스캔이 빨라 안내 없음
    LogShim.AssertEq "pl_inc_scannote_quiet", "", CStr(r("ScanNote"))

    DropSheet
End Sub

' @desc 값 브리지 자동보정과 누락 분리를 검증함
Public Sub Test_PL_CorrectionAndMissing()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    PutNumberCol ws

    Dim priorCorrection As Boolean, frm As frmFilter01
    priorCorrection = modUnivFilter.G_ValueCorrection
    On Error GoTo Fail
    modUnivFilter.G_ValueCorrection = False
    Set frm = New frmFilter01
    Dim panel As MSForms.Frame, box As MSForms.Label, compactHeight As Single, inputTop As Single
    Set panel = frm.Controls("fraOptions")
    Set box = panel.Controls("lblCorrectionSwitch")
    LogShim.AssertEq "pl_corr_ui_default", "", box.Caption
    LogShim.AssertEq "pl_corr_options_collapsed", False, panel.Visible
    compactHeight = frm.Height
    inputTop = frm.txtFilterValues.Top
    frm.ToggleOptions
    LogShim.AssertEq "pl_corr_options_expanded", True, panel.Visible
    LogShim.AssertEq "pl_corr_options_room", True, frm.lblApply.Top >= panel.Top + panel.Height
    frm.ToggleOptions
    LogShim.AssertEq "pl_corr_options_height_restored", compactHeight, frm.Height
    LogShim.AssertEq "pl_corr_input_stable", inputTop, frm.txtFilterValues.Top
    ' 기본 OFF: 퀵 필터는 무적용, 폼 파이프라인은 누락으로 판정한다.
    Dim offResult As Object
    Set offResult = modFilterRun.RunQuickFilter(Body(ws), Area(ws), 1, "1000")
    LogShim.AssertEq "pl_corr_quick_off", True, offResult("QuickNoMatch")
    Set offResult = modFilterRun.RunFilter(Body(ws), Area(ws), 1, "1000", False, False, False)
    LogShim.AssertEq "pl_corr_form_off", CLng(0), offResult("MatchedCount")
    LogShim.AssertEq "pl_corr_off_report", "", CStr(offResult("CorrectionInline"))
    frm.ToggleValueCorrection
    LogShim.AssertEq "pl_corr_ui_checked", ChrW(&H2714), box.Caption
    LogShim.AssertEq "pl_corr_keyboard_reveals_option", True, panel.Visible
    frm.ToggleOptions
    LogShim.AssertEq "pl_corr_ui_shared", True, modUnivFilter.G_ValueCorrection
    Set offResult = modFilterRun.RunQuickFilter(Body(ws), Area(ws), 1, "1000")
    LogShim.AssertEq "pl_corr_quick_on", CLng(1), offResult("MatchedCount")

    ' "1000"은 표시에 없음 -> 값 브리지로 "1,000" 보정. "9999"는 누락.
    Dim r As Object
    Set r = modFilterRun.RunFilter(Body(ws), Area(ws), 1, _
                                   "1000" & vbCrLf & "9999", False, False, False, bValueCorrection:=True)

    LogShim.AssertEq "pl_corr_ok", True, r("Ok")
    LogShim.AssertEq "pl_corr_inline", True, InStr(modReport.StatusLine(r), "1000 -> 1,000") > 0
    LogShim.AssertEq "pl_corr_missing_msg", True, Len(r("MissingMessage")) > 0
    LogShim.AssertEq "pl_corr_visible", CLng(1), modApply.CountVisibleRows(Area(ws))
Cleanup:
    Set box = Nothing: Set panel = Nothing
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    modUnivFilter.G_ValueCorrection = priorCorrection
    DropSheet
    Exit Sub
Fail:
    LogShim.AssertEq "pl_corr_error", "", Err.Description
    Resume Cleanup
End Sub

' @desc 제외 모드(여집합+빈칸 유지)를 검증함
Public Sub Test_PL_Exclude()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "H"
    ws.Range("A2").Value = "AA"
    ' A3 빈칸
    ws.Range("A4").Value = "BB"

    Dim r As Object
    Set r = modFilterRun.RunFilter(ws.Range("A2:A4"), ws.Range("A1:A4"), 1, _
                                   "AA", True, False, False)

    LogShim.AssertEq "pl_exc_ok", True, r("Ok")
    LogShim.AssertEq "pl_exc_visible", CLng(2), modApply.CountVisibleRows(ws.Range("A1:A4"))
    LogShim.AssertEq "pl_exc_blank_kept", False, ws.Rows(3).Hidden

    DropSheet
End Sub

' @desc 퀵 모드 0건 무적용(초기 상태 보존)을 검증함
Public Sub Test_PL_QuickZeroNoop()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    PutNumberCol ws

    Dim r As Object
    Set r = modFilterRun.RunQuickFilter(Body(ws), Area(ws), 1, "zzz_nothing")

    LogShim.AssertEq "pl_quick_ok", True, r("Ok")
    LogShim.AssertEq "pl_quick_nomatch", True, r("QuickNoMatch")
    ' 필터 미적용: 전 행 가시
    LogShim.AssertEq "pl_quick_untouched", CLng(3), modApply.CountVisibleRows(Area(ws))
    LogShim.AssertEq "pl_quick_no_automode", False, ws.AutoFilterMode

    DropSheet
End Sub

' @desc 퀵 모드 정상 매칭 즉시 적용을 검증함
Public Sub Test_PL_QuickHit()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    PutNumberCol ws

    Dim r As Object
    Set r = modFilterRun.RunQuickFilter(Body(ws), Area(ws), 1, "2,000")

    LogShim.AssertEq "pl_quickhit_ok", True, r("Ok")
    LogShim.AssertEq "pl_quickhit_nomatch_false", False, r("QuickNoMatch")
    LogShim.AssertEq "pl_quickhit_visible", CLng(1), modApply.CountVisibleRows(Area(ws))

    DropSheet
End Sub

' @desc 빈 입력의 실패 계약을 검증함
Public Sub Test_PL_EmptyInput()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    PutNumberCol ws

    Dim r As Object
    Set r = modFilterRun.RunFilter(Body(ws), Area(ws), 1, "  ", False, False, False)

    LogShim.AssertEq "pl_empty_not_ok", False, r("Ok")
    LogShim.AssertEq "pl_empty_errmsg", True, Len(r("ErrMsg")) > 0

    DropSheet
End Sub

' @desc 퀵 경로의 긴 값 거부와 부분 적용 경고를 검증함
Public Sub Test_PL_QuickLongValues()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    On Error GoTo Fail
    ws.Range("A1").Value = "H"
    ws.Range("A2").Value = String$(300, "A")
    ws.Range("A3").Value = String$(300, "B")
    ws.Range("A4").Value = String$(300, "C")
    ws.Range("A5").Value = "keep"
    ws.Range("A1:A5").AutoFilter Field:=1, Criteria1:="keep"
    Dim oldCriteria As String
    oldCriteria = CStr(ws.AutoFilter.Filters(1).Criteria1)

    Dim r As Object, raw As String
    raw = ws.Range("A2").Value & vbLf & ws.Range("A3").Value & vbLf & ws.Range("A4").Value
    Set r = modFilterRun.RunQuickFilter(ws.Range("A2:A5"), ws.Range("A1:A5"), 1, raw)
    LogShim.AssertEq "pl_long_not_ok", False, r("Ok")
    LogShim.AssertEq "pl_long_reason", True, InStr(r("ErrMsg"), "기존 필터를 유지") > 0
    LogShim.AssertEq "pl_long_filter_preserved", oldCriteria, CStr(ws.AutoFilter.Filters(1).Criteria1)

    Set r = modFilterRun.RunQuickFilter(ws.Range("A2:A5"), ws.Range("A1:A5"), 1, _
                                       "keep" & vbLf & ws.Range("A2").Value)
    LogShim.AssertEq "pl_long_partial_ok", True, r("Ok")
    LogShim.AssertEq "pl_long_warning", True, Len(r("WarningMessage")) > 0
    LogShim.AssertEq "pl_long_selfcheck", True, Len(r("SelfCheckNote")) > 0
    LogShim.AssertEq "pl_long_partial_visible", CLng(1), modApply.CountVisibleRows(ws.Range("A1:A5"))
Cleanup:
    DropSheet
    Exit Sub
Fail:
    LogShim.AssertEq "pl_long_error", "", Err.Description
    Resume Cleanup
End Sub

' @desc 세로 표/통합문서별 폼 분리와 같은 대상 재사용을 검증함
Public Sub Test_PL_TargetForms()
    Dim ws As Worksheet, wb As Workbook, other As Worksheet
    Dim topForm As frmFilter01, bottomForm As frmFilter01
    Dim otherForm As frmFilter01, again As frmFilter01
    Set ws = PrepSheet()
    On Error GoTo Fail
    ws.Range("A1").Value = "Top": ws.Range("A2").Value = "a"
    ws.Range("A5").Value = "Bottom": ws.Range("A6").Value = "b"
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:A2"), , xlYes)
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A5:A6"), , xlYes)
    Dim priorForms As Long
    If Not modUnivFilter.G_OpenForms Is Nothing Then priorForms = modUnivFilter.G_OpenForms.Count
    Set topForm = modUnivFilter.GetFilterForm(ws.Range("A1"))
    Set bottomForm = modUnivFilter.GetFilterForm(ws.Range("A5"))
    LogShim.AssertEq "pl_target_top", "$A$2", topForm.InjectBodyRange.Address
    LogShim.AssertEq "pl_target_bottom", "$A$6", bottomForm.InjectBodyRange.Address
    LogShim.AssertEq "pl_target_distinct_tables", False, topForm Is bottomForm
    Set again = modUnivFilter.GetFilterForm(ws.Range("A6"))
    LogShim.AssertEq "pl_target_reuse", True, again Is bottomForm
    lo.Resize ws.Range("A5:A7")
    Set again = modUnivFilter.GetFilterForm(ws.Range("A5"))
    LogShim.AssertEq "pl_target_resize_reuse", True, again Is bottomForm
    LogShim.AssertEq "pl_target_resize_body", "$A$6:$A$7", again.InjectBodyRange.Address

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set other = wb.Worksheets(1)
    other.Name = ws.Name
    other.Range("A1").Value = "Other": other.Range("A2").Value = "c"
    other.Range("A1:A2").AutoFilter
    Set otherForm = modUnivFilter.GetFilterForm(other.Range("A1"))
    LogShim.AssertEq "pl_target_distinct_books", False, topForm Is otherForm
    LogShim.AssertEq "pl_target_book", wb.Name, otherForm.InjectBodyRange.Parent.Parent.Name
    Set again = modUnivFilter.GetFilterForm(other.Range("A10"))
    LogShim.AssertEq "pl_target_outside", True, again Is Nothing
Cleanup:
    On Error Resume Next
    If Not topForm Is Nothing Then Unload topForm
    If Not bottomForm Is Nothing Then Unload bottomForm
    If Not otherForm Is Nothing Then Unload otherForm
    Set again = Nothing
    Set topForm = Nothing: Set bottomForm = Nothing: Set otherForm = Nothing
    LogShim.AssertEq "pl_target_closed", priorForms, modUnivFilter.G_OpenForms.Count
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    DropSheet
    Exit Sub
Fail:
    LogShim.AssertEq "pl_target_error", "", Err.Description
    Resume Cleanup
End Sub

' @desc 버튼 생성 실패 후 Excel 계산/이벤트/화면 설정 복원을 검증함
Public Sub Test_PL_ButtonFailureRestoresState()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    Dim calc As XlCalculation, events As Boolean, screen As Boolean
    calc = Application.Calculation
    events = Application.EnableEvents
    screen = Application.ScreenUpdating
    On Error GoTo Fail
    PutNumberCol ws
    ws.Range("A1:A4").AutoFilter
    ws.Protect DrawingObjects:=True
    Dim errNum As Long
    On Error Resume Next
    modUnivFilter.CreateSearchButtons ws.Range("A1")
    errNum = Err.Number
    On Error GoTo Fail
    LogShim.AssertEq "pl_button_error", True, errNum <> 0
    LogShim.AssertEq "pl_button_calc", CLng(calc), CLng(Application.Calculation)
    LogShim.AssertEq "pl_button_events", events, Application.EnableEvents
    LogShim.AssertEq "pl_button_screen", screen, Application.ScreenUpdating
Cleanup:
    On Error Resume Next
    ws.Unprotect
    Application.Calculation = calc
    Application.EnableEvents = events
    Application.ScreenUpdating = screen
    On Error GoTo 0
    DropSheet
    Exit Sub
Fail:
    LogShim.AssertEq "pl_button_restore_error", "", Err.Description
    Resume Cleanup
End Sub

' @desc 정상/실패 반환 시 원래 전역 설정과 중단 정책 보존을 검증함
Public Sub Test_PL_RestoresSettings()
    Dim ws As Worksheet, calc As XlCalculation, events As Boolean, screen As Boolean
    Dim cancel As XlEnableCancelKey, cursor As XlMousePointer, mode As Long, invalid As Long
    Dim r As Object, expectedCalc As XlCalculation, expectedEnabled As Boolean
    Set ws = PrepSheet()
    calc = Application.Calculation: events = Application.EnableEvents
    screen = Application.ScreenUpdating: cancel = Application.EnableCancelKey
    cursor = Application.Cursor
    On Error GoTo Fail
    PutNumberCol ws
    For mode = 0 To 1
        expectedEnabled = (mode = 0)
        If mode = 0 Then expectedCalc = xlCalculationAutomatic Else expectedCalc = xlCalculationManual
        For invalid = 0 To 1
            Application.Calculation = expectedCalc
            Application.EnableEvents = expectedEnabled
            Application.ScreenUpdating = expectedEnabled
            Application.EnableCancelKey = xlDisabled
            Application.Cursor = xlNorthwestArrow
            If invalid = 0 Then
                Set r = modFilterRun.RunFilter(Body(ws), Area(ws), 1, "1,000", False, False, False)
            Else
                Set r = modFilterRun.RunFilter(Nothing, Area(ws), 1, "1,000", False, False, False)
            End If
            LogShim.AssertEq "pl_restore_result", (invalid = 0), r("Ok")
            LogShim.AssertEq "pl_restore_calc", CLng(expectedCalc), CLng(Application.Calculation)
            LogShim.AssertEq "pl_restore_events", expectedEnabled, Application.EnableEvents
            LogShim.AssertEq "pl_restore_screen", expectedEnabled, Application.ScreenUpdating
            LogShim.AssertEq "pl_restore_cancel", CLng(xlDisabled), CLng(Application.EnableCancelKey)
            LogShim.AssertEq "pl_restore_cursor", CLng(xlNorthwestArrow), CLng(Application.Cursor)
        Next invalid
    Next mode
Cleanup:
    Application.EnableCancelKey = xlDisabled
    Utilities.OptimizedMode False
    Application.Calculation = calc
    Application.EnableEvents = events
    Application.ScreenUpdating = screen
    Application.Cursor = cursor
    Application.EnableCancelKey = cancel
    DropSheet
    Exit Sub
Fail:
    LogShim.AssertEq "pl_restore_error", "", Err.Description
    Resume Cleanup
End Sub

' 헤더 + 숫자(#,##0) 3행 열을 만든다: 1,000 / 2,000 / 3,000
Private Sub PutNumberCol(ByVal ws As Worksheet)
    ws.Range("A1").Value = "H"
    ws.Range("A2:A4").NumberFormat = "#,##0"
    ws.Range("A2").Value = 1000
    ws.Range("A3").Value = 2000
    ws.Range("A4").Value = 3000
End Sub

Private Function Body(ByVal ws As Worksheet) As Range
    Set Body = ws.Range("A2:A4")
End Function

Private Function Area(ByVal ws As Worksheet) As Range
    Set Area = ws.Range("A1:A4")
End Function

Private Function PrepSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_NAME
    End If
    ws.AutoFilterMode = False
    ws.Cells.Clear
    ' 시트 재생성 간 결정성 확보 (지문이 내용차를 잡지만 잔여 캐시 청소)
    modCache.ClearCache
    Set PrepSheet = ws
End Function

Private Sub DropSheet()
    Dim prev As Boolean
    prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = prev
End Sub
