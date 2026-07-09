Attribute VB_Name = "Test_ColumnScan"
' @desc modColumnScan 자산 구축 계약 단위테스트 담당
' @depends modColumnScan, clsColAsset, clsSet, LogShim
Option Explicit

' 2026-07 재설계 스펙: 실측 표시 키 규칙(gap-probe-findings.md)을 고정한다.

Private Const SHEET_NAME As String = "UF_ScanTest"

' @desc 균일 숫자 서식 열의 표시/값/행수 자산을 검증함
Public Sub Test_CS_UniformNumber()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1:A3").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000
    ws.Range("A2").Value = 2000
    ws.Range("A3").Value = 1000
    ' A4는 빈칸

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A4"), "T", 1)

    LogShim.AssertEq "cs_num_displays", CLng(2), asset.Displays.Count
    LogShim.AssertEq "cs_num_has_comma", True, asset.Displays.Contains("1,000")
    LogShim.AssertEq "cs_num_vmap", True, asset.ValueMap.exists(CStr(1000#))
    LogShim.AssertEq "cs_num_blank", CLng(1), asset.BlankCount
    LogShim.AssertEq "cs_num_rows", CLng(4), asset.RowCount
    LogShim.AssertEq "cs_num_rowbykey", CLng(2), CLng(asset.RowCountByKey("1,000"))

    Dim sel As New clsSet
    sel.Add "1,000"
    LogShim.AssertEq "cs_num_expected", CLng(2), asset.ExpectedRows(sel, False)

    DropSheet
End Sub

' @desc 회계(* 패딩) 서식이 접힌형 키로 렌더되는지 검증함
Public Sub Test_CS_Accounting()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1:A2").NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""_-;_-@_-"
    ws.Range("A1").Value = 1234
    ws.Range("A2").Value = 0

    Dim cleanFmt As String
    cleanFmt = Replace(ws.Range("A1").NumberFormatLocal, "!", "")
    Dim expPos As String, expZero As String
    expPos = Application.WorksheetFunction.Text(1234, cleanFmt)
    expZero = Application.WorksheetFunction.Text(0, cleanFmt)

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A2"), "T", 1)

    LogShim.AssertEq "cs_acc_collapsed_pos", True, asset.Displays.Contains(expPos)
    LogShim.AssertEq "cs_acc_collapsed_zero", True, asset.Displays.Contains(expZero)
    ' 폭 패딩된 .Text는 키가 아니다 (매칭 불가 실측)
    LogShim.AssertEq "cs_acc_no_padded", False, asset.Displays.Contains(ws.Range("A1").Text)

    DropSheet
End Sub

' @desc 혼합 서식(값쌍둥이) 열의 표시 분리와 값 지도 통합을 검증함
Public Sub Test_CS_MixedFormats()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").NumberFormat = "0.00"
    ws.Range("A1").Value = 0.5
    ws.Range("A2").NumberFormat = "0%"
    ws.Range("A2").Value = 0.5

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A2"), "T", 1)

    LogShim.AssertEq "cs_mix_displays", CLng(2), asset.Displays.Count
    LogShim.AssertEq "cs_mix_050", True, asset.Displays.Contains("0.50")
    LogShim.AssertEq "cs_mix_50pct", True, asset.Displays.Contains("50%")
    LogShim.AssertEq "cs_mix_vmap_twins", CLng(2), _
        CLng(UBound(Split(asset.ValueMap(CStr(0.5)), vbNullChar)) + 1)

    DropSheet
End Sub

' @desc 좁은 열(#-런)의 AutoFit 재독과 폭 복원을 검증함
Public Sub Test_CS_HashNarrow()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").NumberFormat = "#,##0"
    ws.Range("A1").Value = 9876543
    ws.Columns("A").ColumnWidth = 2

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A1"), "T", 1)

    LogShim.AssertEq "cs_hash_ideal", True, asset.Displays.Contains("9,876,543")
    LogShim.AssertEq "cs_hash_width_restored", CDbl(2), ws.Columns("A").ColumnWidth

    DropSheet
End Sub

' @desc 실제 날짜와 텍스트 쌍둥이가 하나의 표시 키로 합쳐지는지 검증함
Public Sub Test_CS_DateAndText()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").NumberFormat = "yyyy-mm-dd"
    ws.Range("A1").Value = DateSerial(2024, 1, 3)
    ws.Range("A2").NumberFormat = "@"
    ws.Range("A2").Value = "2024-01-03"

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A2"), "T", 1)

    LogShim.AssertEq "cs_date_one_key", CLng(1), asset.Displays.Count
    LogShim.AssertEq "cs_date_two_rows", CLng(2), CLng(asset.RowCountByKey("2024-01-03"))
    ' 값 지도에는 실제 날짜만 (시리얼 키)
    LogShim.AssertEq "cs_date_vmap", True, _
        asset.ValueMap.exists(CStr(CDbl(DateSerial(2024, 1, 3))))

    DropSheet
End Sub

' @desc 블록 구조 혼합 서식 열의 이진 분할 정확성을 검증함
Public Sub Test_CS_BlockMixed()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ' 3블록: A1:A2 #,##0(1000쌍) / A3:A4 0(2000쌍) / A5 yyyy-mm-dd(날짜)
    ws.Range("A1:A2").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000
    ws.Range("A2").Value = 1000
    ws.Range("A3:A4").NumberFormat = "0"
    ws.Range("A3").Value = 2000
    ws.Range("A4").Value = 2000
    ws.Range("A5").NumberFormat = "yyyy-mm-dd"
    ws.Range("A5").Value = DateSerial(2024, 1, 3)

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A5"), "T", 1)

    ' 표시 키: "1,000"(블록1) / "2000"(블록2) / "2024-01-03"(블록3)
    LogShim.AssertEq "cs_block_displays", CLng(3), asset.Displays.Count
    LogShim.AssertEq "cs_block_comma", True, asset.Displays.Contains("1,000")
    LogShim.AssertEq "cs_block_plain", True, asset.Displays.Contains("2000")
    LogShim.AssertEq "cs_block_date", True, asset.Displays.Contains("2024-01-03")
    ' 행수 누적: 블록1 2행, 블록2 2행
    LogShim.AssertEq "cs_block_rows1", CLng(2), CLng(asset.RowCountByKey("1,000"))
    LogShim.AssertEq "cs_block_rows2", CLng(2), CLng(asset.RowCountByKey("2000"))
    ' 값 지도: 같은 값 1000/2000이 서로 다른 표시로 분리
    LogShim.AssertEq "cs_block_vmap_1000", "1,000", CStr(asset.ValueMap(CStr(1000#)))
    LogShim.AssertEq "cs_block_vmap_2000", "2000", CStr(asset.ValueMap(CStr(2000#)))

    DropSheet
End Sub

' @desc 에러 셀과 빈 열의 방어 동작을 검증함
Public Sub Test_CS_ErrorAndEmpty()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Formula = "=NA()"

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A1"), "T", 1)
    LogShim.AssertEq "cs_error_display", True, asset.Displays.Contains("#N/A")

    ' 빈 열: 주입 범위를 신뢰하므로 전 행이 빈칸으로 집계된다
    Set asset = modColumnScan.BuildAsset(ws.Range("C1:C5"), "T", 1)
    LogShim.AssertEq "cs_empty_rows", CLng(5), asset.RowCount
    LogShim.AssertEq "cs_empty_blanks", CLng(5), asset.BlankCount
    LogShim.AssertEq "cs_empty_displays", CLng(0), asset.Displays.Count

    DropSheet
End Sub

' 스크래치 시트를 만들고 초기화한다.
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
    ws.Columns("A:C").ColumnWidth = 12
    Set PrepSheet = ws
End Function

' 스크래치 시트를 제거한다(떠돌이 문서 모듈 방지).
Private Sub DropSheet()
    Dim prev As Boolean
    prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = prev
End Sub
