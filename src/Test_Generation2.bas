Attribute VB_Name = "Test_Generation2"
' @desc modCache 세대 계약(투명 재빌드) 테스트 담당
' @depends modCache, clsColAsset, LogShim
Option Explicit

' 2026-07 재설계 스펙: stale 자산은 fail-fast 대신 캐시 내부에서 투명 재스캔된다.
' 불변식은 동일: 낡은 스냅샷으로는 절대 필터하지 않는다.

Private Const SHEET_NAME As String = "UF_GenTest"

' @desc 캐시 적중(동일 객체 반환)을 검증함
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

' @desc 무효화 후 투명 재스캔(새 내용 반영)을 검증함
Public Sub Test_G2_StaleRebuild()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "AA"

    Dim a1 As clsColAsset
    Set a1 = modCache.GetColumnAsset(ws.Range("A1:A1"))
    LogShim.AssertEq "g2_before_edit", True, a1.Displays.Contains("AA")

    ' 시트 편집 + 무효화 (실환경에서는 AppEvents_SheetChange가 호출)
    ws.Range("A1").Value = "ZZ"
    modCache.InvalidateCacheBySheet ws

    Dim a2 As clsColAsset
    Set a2 = modCache.GetColumnAsset(ws.Range("A1:A1"))

    LogShim.AssertEq "g2_rebuilt_new_object", False, (a1 Is a2)
    LogShim.AssertEq "g2_rebuilt_content", True, a2.Displays.Contains("ZZ")
    LogShim.AssertEq "g2_rebuilt_old_gone", False, a2.Displays.Contains("AA")
    LogShim.AssertEq "g2_generation_advanced", True, (a2.Generation > a1.Generation)

    DropSheet
End Sub

' @desc 전역 무효화 경로를 검증함
Public Sub Test_G2_GlobalInvalidate()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Value = "AA"

    Dim a1 As clsColAsset
    Set a1 = modCache.GetColumnAsset(ws.Range("A1:A1"))

    modCache.InvalidateCache

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
