Attribute VB_Name = "Test_Generation2"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc modCache 지문 신선도 계약(관측 재빌드) 테스트 담당
' @depends modCache, clsColAsset, LogShim
Option Explicit

' 2026-07 지문 재설계: 신선도는 세대 예측이 아니라 지문 관측으로 확정된다.
' 이벤트/수동 무효화 없이도 값·서식 변경이 재취득 시 자동 감지된다.
' 불변식은 동일: 낡은 스냅샷으로는 절대 필터하지 않는다.

Private Const SHEET_NAME As String = "UF_GenTest"

' @desc 무변경 재취득 시 캐시 적중(동일 객체 반환)을 검증함
Public Sub Test_G2_CacheHit()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "AA"
    ws.Range("A2").Value = "BB"

    Dim a1 As clsColAsset, a2 As clsColAsset
    Set a1 = modCache.GetColumnAsset(ws.Range("A1:A2"))
    Set a2 = modCache.GetColumnAsset(ws.Range("A1:A2"))

    LogShim.AssertEq "g2_hit_same_object", True, (a1 Is a2)
    LogShim.AssertEq "g2_hit_content", True, a1.Displays.Contains("AA")

    DropSheet
End Sub

' @desc 값 편집을 이벤트/무효화 없이 지문이 감지해 반영하는지 검증함(증분 패치)
Public Sub Test_G2_StaleRebuild()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "AA"

    Dim a1 As clsColAsset
    Set a1 = modCache.GetColumnAsset(ws.Range("A1:A1"))
    LogShim.AssertEq "g2_before_edit", True, a1.Displays.Contains("AA")

    ' 시트 편집만 (수동 무효화 없음 - 지문이 관측으로 잡아야 한다).
    ' 값만 바뀌면 증분 패치가 같은 자산을 제자리 갱신한다 (불변식은 내용 신선도).
    ws.Range("A1").Value = "ZZ"

    Dim a2 As clsColAsset
    Set a2 = modCache.GetColumnAsset(ws.Range("A1:A1"))

    LogShim.AssertEq "g2_patched_same_object", True, (a1 Is a2)
    LogShim.AssertEq "g2_rebuilt_content", True, a2.Displays.Contains("ZZ")
    LogShim.AssertEq "g2_rebuilt_old_gone", False, a2.Displays.Contains("AA")

    DropSheet
End Sub

' @desc 값/빈칸 전이 편집이 증분 패치로 정확히 반영되는지 검증함(빈칸 델타/인덱스)
Public Sub Test_G2_PatchTransitions()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1:A4").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000
    ws.Range("A2").Value = 2000
    ' A3은 빈칸
    ws.Range("A4").Value = 1000

    Dim a1 As clsColAsset
    Set a1 = modCache.GetColumnAsset(ws.Range("A1:A4"))
    LogShim.AssertEq "g2_patch_pre_blank", CLng(1), a1.BlankCount
    LogShim.AssertEq "g2_patch_pre_rows1k", CLng(2), CLng(a1.RowCountByKey("1,000"))

    ' 값->값(A1), 값->빈칸(A2), 빈칸->값(A3) 동시 편집
    ws.Range("A1").Value = 3000
    ws.Range("A2").ClearContents
    ws.Range("A3").Value = 5000

    Dim a2 As clsColAsset
    Set a2 = modCache.GetColumnAsset(ws.Range("A1:A4"))

    LogShim.AssertEq "g2_patch_same_object", True, (a1 Is a2)
    LogShim.AssertEq "g2_patch_new_3k", True, a2.Displays.Contains("3,000")
    LogShim.AssertEq "g2_patch_new_5k", True, a2.Displays.Contains("5,000")
    LogShim.AssertEq "g2_patch_old_2k_gone", False, a2.Displays.Contains("2,000")
    LogShim.AssertEq "g2_patch_kept_1k", True, a2.Displays.Contains("1,000")
    LogShim.AssertEq "g2_patch_blank", CLng(1), a2.BlankCount
    LogShim.AssertEq "g2_patch_rows1k", CLng(1), CLng(a2.RowCountByKey("1,000"))
    LogShim.AssertEq "g2_patch_rowcount", CLng(4), a2.RowCount

    DropSheet
End Sub

' @desc 서식만 바꿔도(이벤트 미발화) 지문이 재스캔을 유발하는지 검증함(결함#3)
Public Sub Test_G2_FormatEditRebuild()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000

    Dim a1 As clsColAsset
    Set a1 = modCache.GetColumnAsset(ws.Range("A1:A1"))
    LogShim.AssertEq "g2_fmt_before", True, a1.Displays.Contains("1,000")

    ' 값 불변, 서식만 변경 (SheetChange 미발화 상황)
    ws.Range("A1").NumberFormat = "0.00"

    Dim a2 As clsColAsset
    Set a2 = modCache.GetColumnAsset(ws.Range("A1:A1"))
    LogShim.AssertEq "g2_fmt_new_object", False, (a1 Is a2)
    LogShim.AssertEq "g2_fmt_new_key", True, a2.Displays.Contains("1000.00")
    LogShim.AssertEq "g2_fmt_old_gone", False, a2.Displays.Contains("1,000")

    DropSheet
End Sub

' @desc 열별 지문 격리(무관 열은 편집돼도 적중)로 스래싱이 없는지 검증함
Public Sub Test_G2_PerColumnIsolation()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "AA"
    ws.Range("B1").Value = "BB"

    Dim ca1 As clsColAsset, cb1 As clsColAsset
    Set ca1 = modCache.GetColumnAsset(ws.Range("A1:A1"))
    Set cb1 = modCache.GetColumnAsset(ws.Range("B1:B1"))

    ' A열만 편집
    ws.Range("A1").Value = "ZZ"

    Dim ca2 As clsColAsset, cb2 As clsColAsset
    Set ca2 = modCache.GetColumnAsset(ws.Range("A1:A1"))
    Set cb2 = modCache.GetColumnAsset(ws.Range("B1:B1"))

    LogShim.AssertEq "g2_iso_a_fresh", True, ca2.Displays.Contains("ZZ")   ' 편집된 A는 갱신(패치)
    LogShim.AssertEq "g2_iso_b_hit", True, (cb1 Is cb2)        ' 무관한 B는 적중(스래싱 없음)

    DropSheet
End Sub

' @desc 전역 캐시 청소 후 재취득 시 새 자산을 내는지 검증함
Public Sub Test_G2_GlobalInvalidate()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "AA"

    Dim a1 As clsColAsset
    Set a1 = modCache.GetColumnAsset(ws.Range("A1:A1"))

    modCache.ClearCache

    Dim a2 As clsColAsset
    Set a2 = modCache.GetColumnAsset(ws.Range("A1:A1"))
    LogShim.AssertEq "g2_global_new_object", False, (a1 Is a2)

    DropSheet
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
    ws.Cells.Clear
    ' 잔여 캐시 청소 (지문이 내용차를 잡지만, 시트 재생성 간 결정성 확보)
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
