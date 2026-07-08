Attribute VB_Name = "Test_Pipeline"
' @desc modFilterRun 파이프라인 e2e 계약 테스트 담당
' @depends modFilterRun, modCache, modApply, clsSet, LogShim
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

    DropSheet
End Sub

' @desc 값 브리지 자동보정과 누락 분리를 검증함
Public Sub Test_PL_CorrectionAndMissing()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    PutNumberCol ws

    ' "1000"은 표시에 없음 -> 값 브리지로 "1,000" 보정. "9999"는 누락.
    Dim r As Object
    Set r = modFilterRun.RunFilter(Body(ws), Area(ws), 1, _
                                   "1000" & vbCrLf & "9999", False, False, False)

    LogShim.AssertEq "pl_corr_ok", True, r("Ok")
    LogShim.AssertEq "pl_corr_msg", True, Len(r("CorrectionMessage")) > 0
    LogShim.AssertEq "pl_corr_missing_msg", True, Len(r("MissingMessage")) > 0
    LogShim.AssertEq "pl_corr_visible", CLng(1), modApply.CountVisibleRows(Area(ws))

    DropSheet
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
    ' 헤드리스에선 AppEvents가 미배선이므로 실환경의 SheetChange 무효화를 수동 재현
    modCache.InvalidateCacheBySheet ws
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
