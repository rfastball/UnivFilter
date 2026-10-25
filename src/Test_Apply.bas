Attribute VB_Name = "Test_Apply"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc modApply 적용 계약과 스캔-매칭-적용 e2e 테스트 담당
' @depends modApply, modColumnScan, modMatch, clsColAsset, clsSet, LogShim
Option Explicit

' 2026-07 재설계 스펙: xlFilterValues 단일 Tier 적용을 실기 시트로 고정한다.

Private Const SHEET_NAME As String = "UF_ApplyTest"

' @desc 포함 모드 배열 적용과 가시 행수를 검증함
Public Sub Test_AP_IncludeArray()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    PutTextCol ws, Array("Apple", "Banana", "Cherry")

    Dim sel As New clsSet
    sel.Add "Apple"
    sel.Add "Cherry"

    Dim info As Object
    Set info = modApply.ApplyDisplays(ws.Range("A1:A4"), 1, sel, False, False)

    LogShim.AssertEq "ap_inc_mode", "array", CStr(info("Mode"))
    LogShim.AssertEq "ap_inc_visible", CLng(2), modApply.CountVisibleRows(ws.Range("A1:A4"))

    DropSheet
End Sub

' @desc 제외 여집합 + 빈칸 유지 적용을 검증함
Public Sub Test_AP_ExcludeWithBlanks()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "H"
    ws.Range("A2").Value = "AA"
    ' A3 빈칸
    ws.Range("A4").Value = "BB"

    ' "AA" 제외 -> 여집합 {BB} + 빈칸
    Dim comp As New clsSet
    comp.Add "BB"

    Dim info As Object
    Set info = modApply.ApplyDisplays(ws.Range("A1:A4"), 1, comp, True, False)

    LogShim.AssertEq "ap_exc_visible", CLng(2), modApply.CountVisibleRows(ws.Range("A1:A4"))
    LogShim.AssertEq "ap_exc_blank_shown", False, ws.Rows(3).Hidden
    LogShim.AssertEq "ap_exc_aa_hidden", True, ws.Rows(2).Hidden

    DropSheet
End Sub

' @desc 전체 숨김과 필드 해제 경로를 검증함
Public Sub Test_AP_HideAllAndClear()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    PutTextCol ws, Array("Apple", "Banana")

    modApply.ApplyHideAll ws.Range("A1:A3"), 1, False
    LogShim.AssertEq "ap_hideall", CLng(0), modApply.CountVisibleRows(ws.Range("A1:A3"))

    modApply.ClearField ws.Range("A1:A3"), 1
    LogShim.AssertEq "ap_clear", CLng(2), modApply.CountVisibleRows(ws.Range("A1:A3"))

    DropSheet
End Sub

' @desc 스택 모드의 타 열 AND 결합을 검증함
Public Sub Test_AP_Stack()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "H1": ws.Range("B1").Value = "H2"
    ws.Range("A2").Value = "X": ws.Range("B2").Value = "P"
    ws.Range("A3").Value = "X": ws.Range("B3").Value = "Q"
    ws.Range("A4").Value = "Y": ws.Range("B4").Value = "P"

    Dim rng As Range
    Set rng = ws.Range("A1:B4")

    Dim s1 As New clsSet
    s1.Add "X"
    modApply.ApplyDisplays rng, 1, s1, False, False

    Dim s2 As New clsSet
    s2.Add "P"
    modApply.ApplyDisplays rng, 2, s2, False, True   ' Stack: 1열 필터 유지

    LogShim.AssertEq "ap_stack_and", CLng(1), modApply.CountVisibleRows(rng)
    LogShim.AssertEq "ap_stack_row2", False, ws.Rows(2).Hidden

    ' Reset: 두 필드 모두 해제 후 재적용
    modApply.ApplyDisplays rng, 2, s2, False, False
    LogShim.AssertEq "ap_reset_clears", CLng(2), modApply.CountVisibleRows(rng)

    DropSheet
End Sub

' @desc 255자 초과 항목의 등호 스칼라 경로를 검증함
Public Sub Test_AP_LongItem()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    Dim longVal As String
    longVal = String$(300, "A")
    ws.Range("A1").Value = "H"
    ws.Range("A2").NumberFormat = "@"
    ws.Range("A2").Value = longVal
    ws.Range("A3").Value = "BB"

    Dim sel As New clsSet
    sel.Add longVal

    Dim info As Object
    Set info = modApply.ApplyDisplays(ws.Range("A1:A3"), 1, sel, False, False)

    LogShim.AssertEq "ap_long_mode", "scalar", CStr(info("Mode"))
    LogShim.AssertEq "ap_long_visible", CLng(1), modApply.CountVisibleRows(ws.Range("A1:A3"))

    ' 2건 스칼라는 유지, 3건부터는 기존 필터를 보존하며 명시적으로 거부.
    sel.Add String$(300, "B")
    Set info = modApply.ApplyDisplays(ws.Range("A1:A3"), 1, sel, False, False)
    LogShim.AssertEq "ap_long_pair_mode", "scalar", CStr(info("Mode"))
    ws.Range("A1:A3").AutoFilter Field:=1, Criteria1:="BB"
    Dim oldCriteria As String
    oldCriteria = CStr(ws.AutoFilter.Filters(1).Criteria1)
    sel.Add String$(300, "C")
    Dim errNum As Long
    On Error Resume Next
    modApply.ApplyDisplays ws.Range("A1:A3"), 1, sel, False, False
    errNum = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "ap_long_rejected", CLng(vbObjectError + 2101), errNum
    LogShim.AssertEq "ap_long_preserved", oldCriteria, CStr(ws.AutoFilter.Filters(1).Criteria1)
    LogShim.AssertEq "ap_long_old_row_visible", False, ws.Rows(3).Hidden

    sel.Clear
    Set info = modApply.ApplyDisplays(ws.Range("A1:A3"), 1, sel, False, False)
    LogShim.AssertEq "ap_empty_preserved", oldCriteria, CStr(ws.AutoFilter.Filters(1).Criteria1)

    DropSheet
End Sub

' @desc 회계 열 스캔-값브리지-적용 e2e를 검증함 (실측 P6 확정 경로)
Public Sub Test_AP_AccountingE2E()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "H"
    ws.Range("A2:A3").NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""_-;_-@_-"
    ws.Range("A2").Value = 1234
    ws.Range("A3").Value = 5678

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A2:A3"), "T")

    ' 사용자는 화면 요지("1,234")를 입력 -> 값 레인이 접힌 키로 브리지
    Dim r As Object
    Set r = modMatch.Execute(Array("1,234"), asset, False, True)
    LogShim.AssertEq "ap_acc_matched", CLng(1), r("Matched").Count
    LogShim.AssertEq "ap_acc_corrected", True, r("Corrections").exists("1,234")

    modApply.ApplyDisplays ws.Range("A1:A3"), 1, r("Matched"), False, False
    LogShim.AssertEq "ap_acc_visible", CLng(1), modApply.CountVisibleRows(ws.Range("A1:A3"))
    LogShim.AssertEq "ap_acc_row2", False, ws.Rows(2).Hidden

    DropSheet
End Sub

' @desc 날짜 열 표시 레인 적용 e2e와 자가검증 일치를 검증함
Public Sub Test_AP_DateE2E()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "H"
    ws.Range("A2").NumberFormat = "yyyy-mm-dd"
    ws.Range("A2").Value = DateSerial(2024, 1, 3)
    ws.Range("A3").NumberFormat = "@"
    ws.Range("A3").Value = "2024-01-03"
    ws.Range("A4").NumberFormat = "yyyy-mm-dd"
    ws.Range("A4").Value = DateSerial(2024, 1, 5)

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A2:A4"), "T")

    Dim r As Object
    Set r = modMatch.Execute(Array("2024-01-03"), asset, False)

    modApply.ApplyDisplays ws.Range("A1:A4"), 1, r("Matched"), False, False

    ' 표시 쌍둥이(실제 날짜 + 텍스트) 2행 -> 자가검증 기대값과 실측 일치
    Dim expected As Long
    expected = asset.ExpectedRows(r("Matched"), False)
    LogShim.AssertEq "ap_date_expected", CLng(2), expected
    LogShim.AssertEq "ap_date_visible", expected, modApply.CountVisibleRows(ws.Range("A1:A4"))

    DropSheet
End Sub

' 헤더+텍스트 열을 만든다.
Private Sub PutTextCol(ByVal ws As Worksheet, ByVal vals As Variant)
    ws.Range("A1").Value = "H"
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        ws.Cells(i + 2, 1).Value = vals(i)
    Next i
End Sub

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
