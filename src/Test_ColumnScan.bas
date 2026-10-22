Attribute VB_Name = "Test_ColumnScan"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc modColumnScan 자산 구축 계약 단위테스트 담당
' @depends modMatch, modColumnScan, clsColAsset, clsColSig, clsSet, LogShim
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
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A4"), "T")

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
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A2"), "T")

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
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A2"), "T")

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
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A1"), "T")

    LogShim.AssertEq "cs_hash_ideal", True, asset.Displays.Contains("9,876,543")
    LogShim.AssertEq "cs_hash_width_restored", CDbl(2), ws.Columns("A").ColumnWidth

    DropSheet
End Sub

' @desc General 좁은 열(과학표기/반올림, #-런 아님)이 캐논형 키로 복구되는지 검증함
Public Sub Test_CS_GeneralNarrow()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").NumberFormat = "General"
    ws.Range("A1").Value = 123456789012#     ' 좁으면 과학표기(1E+11) - # 없음
    ws.Range("A2").NumberFormat = "General"
    ws.Range("A2").Value = 0.123456789012    ' 좁으면 반올림(0.1235) - # 없음
    ws.Columns("A").ColumnWidth = 6

    ' 현재 좁은폭 표시(버그 시 이게 잘못된 키가 됨)
    Dim narrowBig As String
    narrowBig = ws.Range("A1").Text

    ' 캐논형 오라클: AutoFit 후 .Text (실측상 AutoFilter 매칭 대상), 폭 6으로 복원 후 스캔
    ws.Columns("A").AutoFit
    Dim expBig As String, expFrac As String
    expBig = ws.Range("A1").Text
    expFrac = ws.Range("A2").Text
    ws.Columns("A").ColumnWidth = 6

    ' 전제: 좁힘이 실제로 표시를 바꿨다 (안 바뀌면 케이스 무의미)
    LogShim.AssertEq "cs_gennar_premise", True, (narrowBig <> expBig)

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A2"), "T")

    ' 좁은 열이어도 캐논형 키로 저장돼야 한다
    LogShim.AssertEq "cs_gennar_big", True, asset.Displays.Contains(expBig)
    LogShim.AssertEq "cs_gennar_frac", True, asset.Displays.Contains(expFrac)
    ' 잘린 현재폭형이 키로 새면 안 된다
    LogShim.AssertEq "cs_gennar_no_trunc", False, asset.Displays.Contains(narrowBig)
    ' 값 지도도 캐논형으로 연결
    LogShim.AssertEq "cs_gennar_vmap", True, _
        (InStr(vbNullChar & asset.ValueMap(CStr(123456789012#)) & vbNullChar, _
               vbNullChar & expBig & vbNullChar) > 0)
    ' 폭 복원
    LogShim.AssertEq "cs_gennar_width_restored", CDbl(6), ws.Columns("A").ColumnWidth

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
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A2"), "T")

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
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A5"), "T")

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

' @desc 데이터는 균일하나 빈칸만 다른 서식일 때 분할이 올바른 결과를 내는지 검증함
Public Sub Test_CS_DataUniformBlankPollution()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1:A5").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000
    ws.Range("A2").Value = 2000
    ' A3은 빈칸이면서 General 서식 -> 전체 범위 NumberFormatLocal이 Null로 오판
    ws.Range("A3").NumberFormat = "General"
    ws.Range("A4").Value = 1000
    ws.Range("A5").Value = 3000

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A5"), "T")

    ' 데이터 셀만 재판정하면 #,##0 균일 -> 올바른 표시 키
    LogShim.AssertEq "cs_pollute_displays", CLng(3), asset.Displays.Count
    LogShim.AssertEq "cs_pollute_comma", True, asset.Displays.Contains("1,000")
    LogShim.AssertEq "cs_pollute_2k", True, asset.Displays.Contains("2,000")
    LogShim.AssertEq "cs_pollute_rows1k", CLng(2), CLng(asset.RowCountByKey("1,000"))
    LogShim.AssertEq "cs_pollute_blank", CLng(1), asset.BlankCount

    DropSheet
End Sub

' @desc @서식 숫자 셀이 좁은 열에서도 General 캐논형 키로 렌더되는지 검증함(P11)
Public Sub Test_CS_AtNumericCanon()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ' 값 먼저 입력 후 @ 적용 (Double 유지 - 실데이터 SL_V 패턴)
    ws.Range("A1").Value = 212970000
    ws.Range("A2").Value = 1234
    ws.Range("A1:A2").NumberFormat = "@"
    ws.Columns("A").ColumnWidth = 8

    ' 전제: 좁은 폭에서 .Text가 캐논형과 다르게 잘렸다 (P11: 잘린형은 매칭 키가 아님)
    Dim narrowBig As String
    narrowBig = ws.Range("A1").Text
    LogShim.AssertEq "cs_atnum_premise", True, (narrowBig <> "212970000")

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A2"), "T")

    LogShim.AssertEq "cs_atnum_canon", True, asset.Displays.Contains("212970000")
    LogShim.AssertEq "cs_atnum_small", True, asset.Displays.Contains("1234")
    LogShim.AssertEq "cs_atnum_no_clip", False, asset.Displays.Contains(narrowBig)
    ' 값 레인 대칭: RenderValueKey도 General 캐논
    LogShim.AssertEq "cs_atnum_vkey", "212970000", _
        modColumnScan.RenderValueKey(212970000#, "@")

    DropSheet
End Sub

' @desc 에러/불리언 셀이 좁은 열에서도 캐논 문자열 키로 렌더되는지 검증함(P12)
Public Sub Test_CS_ErrBoolCanon()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Formula = "=NA()"
    ws.Range("A2").Formula = "=1/0"
    ws.Range("A3").Value = True
    ws.Range("A4").Value = False
    ws.Columns("A").ColumnWidth = 3

    ' 전제: 좁은 폭에서 에러 표시가 잘렸다 ("##" - P12: 잘린형은 매칭 키가 아님)
    Dim narrowNa As String
    narrowNa = ws.Range("A1").Text
    LogShim.AssertEq "cs_errbool_premise", True, (narrowNa <> "#N/A")

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A4"), "T")

    LogShim.AssertEq "cs_errbool_na", True, asset.Displays.Contains("#N/A")
    LogShim.AssertEq "cs_errbool_div", True, asset.Displays.Contains("#DIV/0!")
    LogShim.AssertEq "cs_errbool_true", True, asset.Displays.Contains("TRUE")
    LogShim.AssertEq "cs_errbool_false", True, asset.Displays.Contains("FALSE")
    LogShim.AssertEq "cs_errbool_no_clip", False, asset.Displays.Contains(narrowNa)

    DropSheet
End Sub

' @desc 에러 셀과 빈 열의 방어 동작을 검증함
Public Sub Test_CS_ErrorAndEmpty()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").Formula = "=NA()"

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A1"), "T")
    LogShim.AssertEq "cs_error_display", True, asset.Displays.Contains("#N/A")

    ' 빈 열: 주입 범위를 신뢰하므로 전 행이 빈칸으로 집계된다
    Set asset = modColumnScan.BuildAsset(ws.Range("C1:C5"), "T")
    LogShim.AssertEq "cs_empty_rows", CLng(5), asset.RowCount
    LogShim.AssertEq "cs_empty_blanks", CLng(5), asset.BlankCount
    LogShim.AssertEq "cs_empty_displays", CLng(0), asset.Displays.Count

    DropSheet
End Sub

' @desc CF(숫자서식) 열이 CF형이 아닌 기본서식 캐논형 키로 렌더되는지 검증함(P8 결론2/P9)
Public Sub Test_CS_CondFormatRender()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ' 혼합 두 셀: #,##0 + General (General은 WText 로컬토큰 실패 케이스 - "General" 매핑 검증)
    ws.Range("A1").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000
    ws.Range("A2").NumberFormat = "General"
    ws.Range("A2").Value = 12.5

    ' 오라클(기본서식 렌더) = AutoFilter 매칭형
    Dim oracleComma As String, oracleGen As String
    oracleComma = ws.Range("A1").Text     ' "1,000"
    oracleGen = ws.Range("A2").Text       ' "12.5"

    ' 숫자서식 CF 추가 -> .Text가 CF형("1000.00"/"12.50")으로 바뀐다
    ws.Range("A1:A2").FormatConditions.Delete
    Dim fc As FormatCondition
    Set fc = ws.Range("A1:A2").FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
    fc.NumberFormat = "0.00"

    ' 전제: .Text가 실제로 CF를 추종해 오라클과 달라졌다
    LogShim.AssertEq "cs_cf_premise", True, (ws.Range("A1").Text <> oracleComma)

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(ws.Range("A1:A2"), "T")

    ' 기본서식 캐논형이 키여야 AutoFilter가 매칭(CF형은 0건 무음실패)
    LogShim.AssertEq "cs_cf_comma_base", True, asset.Displays.Contains(oracleComma)
    LogShim.AssertEq "cs_cf_gen_base", True, asset.Displays.Contains(oracleGen)
    ' CF형이 키로 새면 안 된다
    LogShim.AssertEq "cs_cf_no_cfform_comma", False, asset.Displays.Contains(ws.Range("A1").Text)
    LogShim.AssertEq "cs_cf_no_cfform_gen", False, asset.Displays.Contains(ws.Range("A2").Text)

    ws.Range("A1:A2").FormatConditions.Delete
    DropSheet
End Sub

' @desc 벌크 Evaluate 렌더가 현행 렌더 경로와 전수 동치인지 검증함(지연 우주 Stage A)
Public Sub Test_CS_BulkRenderEquiv()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Columns("A").ColumnWidth = 24

    ' 대표 서식군(winner_table 61케이스 계열 축약) + 함정: 빈칸("0" 함정)/에러(전파)/
    ' General 토큰/* 접힌형/[색] 빈문자열. 각 케이스 = 균일 서식 열 1개.
    EquivCase ws, "gen", "General", _
              Array(5, 12.5, 123456789012#, 0.123456789012, -3.75), True, True
    EquivCase ws, "int0", "0", Array(42, -7, 0), True, False
    EquivCase ws, "dec2", "0.00", Array(0.5, 1234.567, -0.25), False, False
    EquivCase ws, "thou", "#,##0", Array(1000, 9876543, -1234567), True, True
    EquivCase ws, "thou2", "#,##0.00", Array(1234.5, -9999.99), False, False
    EquivCase ws, "pct", "0%", Array(0.5, 0.125, 2), True, False
    EquivCase ws, "sci", "0.00E+00", Array(123456789, 0.0000123), False, False
    EquivCase ws, "frac", "# ?/?", Array(0.5, 2.25), False, False
    EquivCase ws, "curr", "$#,##0.00", Array(1234.5, -1234.5), False, False
    EquivCase ws, "acc_usd", "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""_-;_-@_-", _
              Array(1234, -1234, 0, 5678), True, False
    EquivCase ws, "acc_plain", "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)", _
              Array(1234, -1234, 0), False, False
    EquivCase ws, "acc_won", AccWonFmt(), Array(1234, 0), True, False
    EquivCase ws, "date_iso", "yyyy-mm-dd", _
              Array(DateSerial(2024, 1, 3), DateSerial(1999, 12, 31)), True, False
    EquivCase ws, "date_han", "yyyy""년"" m""월"" d""일""", _
              Array(DateSerial(2024, 1, 3)), False, False
    EquivCase ws, "dtime", "yyyy-mm-dd hh:mm", _
              Array(DateSerial(2024, 1, 3) + TimeSerial(14, 30, 0)), False, False
    EquivCase ws, "time", "hh:mm", Array(TimeSerial(8, 9, 0)), False, False
    EquivCase ws, "color", "[Red]0", Array(5, -5), False, False

    ' 텍스트 열(@): 문자열 자기자신 + 날짜문자열(TEXT 강제 함정 우회) + 숫자 셀(벌크 부적용)
    ws.Columns("A").Clear
    ws.Range("A1:A5").NumberFormat = "@"
    ws.Range("A1").Value = "Apple"
    ws.Range("A2").Value = "1,234,567"
    ws.Range("A3").Value = "2024-01-03"
    ws.Range("A4").Value = 1234
    ' A5 빈칸
    AssertEquiv "textcol", ws.Range("A1:A5")

    ' 불리언 셀(셀별 폴백) + 숫자 블록 혼합
    ws.Columns("A").Clear
    ws.Range("A1").Value = True
    ws.Range("A2").NumberFormat = "0"
    ws.Range("A2").Value = 7
    AssertEquiv "boolcol", ws.Range("A1:A2")

    ' 혼합 블록 열: 값쌍둥이가 블록별 서식으로 각자 렌더
    ws.Columns("A").Clear
    ws.Range("A1:A2").NumberFormat = "0.00"
    ws.Range("A1").Value = 0.5
    ws.Range("A2").Value = 1.5
    ws.Range("A3:A4").NumberFormat = "0%"
    ws.Range("A3").Value = 0.5
    ws.Range("A4").Value = 0.25
    AssertEquiv "mixblocks", ws.Range("A1:A4")

    ' CF 열: 현행은 WText(기본서식), 벌크는 TEXT(기본서식) - 동일 캐논형이어야 함
    ws.Columns("A").Clear
    ws.Range("A1:A3").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000
    ws.Range("A2").Value = 9876543
    ' A3 빈칸
    ws.Range("A1:A3").FormatConditions.Delete
    Dim fc As FormatCondition
    Set fc = ws.Range("A1:A3").FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
    fc.NumberFormat = "0.00"
    AssertEquiv "cfcol", ws.Range("A1:A3")
    ws.Range("A1:A3").FormatConditions.Delete

    ' 좁은 열: 현행은 AutoFit 재독, 벌크는 폭 무관 - 같은 캐논형이어야 함
    ws.Columns("A").Clear
    ws.Range("A1").NumberFormat = "#,##0"
    ws.Range("A1").Value = 9876543
    ws.Columns("A").ColumnWidth = 2
    AssertEquiv "narrow", ws.Range("A1:A1")
    ws.Columns("A").ColumnWidth = 24

    DropSheet
End Sub

' @desc 벌크와 강제 폴백의 후보 렌더링 매칭 계약을 검증함
Public Sub Test_CS_FallbackMatch()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    On Error GoTo Fail
    Dim scenario As Long, rng As Range, sig As clsColSig, content As Variant
    Dim fast As clsColAsset, fallback As clsColAsset
    Dim expected As Object, actual As Object, inputs As Variant, tokens As Variant
    Dim key As Variant, tag As String
    inputs = Array(Array("1.44"), Array("1.4"), Array("9.99"), _
                   Array("1.410"), Array("1.44", "1.4", "9.99"))
    For scenario = 1 To 3
        ws.Cells.Clear
        ws.Range("A1").NumberFormat = "0.0"
        ws.Range("A1").Value2 = 1.41
        Set rng = ws.Range("A1")
        If scenario > 1 Then
            ws.Range("A2").NumberFormat = "0.00"
            ws.Range("A2").Value2 = 1.41
            ws.Range("A3").NumberFormat = "@"
            ws.Range("A3").Value2 = "1"
            ws.Range("A4").NumberFormat = "0"
            Set rng = ws.Range("A1:A4")
        End If
        modColumnScan.G_ForceLegacyRender = False
        Set fast = modColumnScan.BuildAsset(rng, "빠른 경로")
        LogShim.AssertEq "fallback_fast_rendered", True, fast.HasRendered
        Set sig = modColumnScan.ComputeSignature(rng)
        If scenario = 3 Then
            ' 분할 예산 초과 시 생성되는 지문: 블록 목록 없이 셀별 서식만 사용한다.
            content = sig.ContentVec
            sig.Init rng.Rows.Count, content, False, "", Nothing, True
        End If
        modColumnScan.G_ForceLegacyRender = True
        Set fallback = modColumnScan.BuildAssetFromSig(rng, "폴백", sig)
        modColumnScan.G_ForceLegacyRender = False
        For Each tokens In inputs
            tag = "fallback_" & CStr(scenario) & "_" & Join(tokens, "/")
            Set expected = modMatch.Execute(tokens, fast, False)
            Set actual = modMatch.Execute(tokens, fallback, False)
            If CStr(tokens(0)) = "1.44" Then
                LogShim.AssertEq tag & "_oracle", True, expected("Matched").Contains("1.4")
            End If
            LogShim.AssertEq tag & "_count", expected("Matched").Count, actual("Matched").Count
            LogShim.AssertEq tag & "_keys", CLng(0), expected("Matched").Difference(actual("Matched")).Count
            LogShim.AssertEq tag & "_missing", Join(expected("Missing"), "|"), Join(actual("Missing"), "|")
            LogShim.AssertEq tag & "_corrections", expected("Corrections").Count, actual("Corrections").Count
            For Each key In expected("Corrections").Keys
                LogShim.AssertEq tag & "_corrected", CStr(expected("Corrections")(key)), CStr(actual("Corrections")(key))
            Next key
            LogShim.AssertEq tag & "_rows", fast.ExpectedRows(expected("Matched"), False), _
                                                   fallback.ExpectedRows(actual("Matched"), False)
        Next tokens
    Next scenario
Cleanup:
    modColumnScan.G_ForceLegacyRender = False
    DropSheet
    Exit Sub
Fail:
    LogShim.AssertEq "fallback_error", "", Err.Description
    Resume Cleanup
End Sub

' 균일 서식 열 1개를 구성해 동치 대조까지 수행한다(withBlank는 빈칸 행, withErr는 =NA() 행 추가).
Private Sub EquivCase(ws As Worksheet, ByVal cas As String, ByVal fmtEng As String, _
                      vals As Variant, ByVal withBlank As Boolean, ByVal withErr As Boolean)
    ws.Columns("A").Clear
    Dim n As Long
    n = UBound(vals) - LBound(vals) + 1
    Dim total As Long
    total = n + IIf(withBlank, 1, 0) + IIf(withErr, 1, 0)
    ws.Range(ws.Cells(1, 1), ws.Cells(total, 1)).NumberFormat = fmtEng

    Dim r As Long, i As Long
    r = 1
    For i = LBound(vals) To UBound(vals)
        ws.Cells(r, 1).Value = vals(i)
        r = r + 1
    Next i
    If withBlank Then r = r + 1
    If withErr Then ws.Cells(r, 1).Formula = "=NA()"

    AssertEquiv cas, ws.Range(ws.Cells(1, 1), ws.Cells(total, 1))
End Sub

' 셀별 렌더 오라클(.Text/AutoFit/WText 규칙의 BuildPerCell)과 벌크 렌더 키 집계를
' 전수 대조한다(기본이 벌크라 훅으로 오라클 경로를 강제 - 이 동치가 벌크 채택의 근거).
Private Sub AssertEquiv(ByVal cas As String, ByVal rng As Range)
    Dim asset As clsColAsset
    modColumnScan.G_ForceLegacyRender = True
    Set asset = modColumnScan.BuildAsset(rng, "T")
    modColumnScan.G_ForceLegacyRender = False

    Dim sig As clsColSig
    Set sig = modColumnScan.ComputeSignature(rng)
    LogShim.AssertEq cas & "_not_pathological", False, sig.IsPathological

    Dim nRows As Long
    nRows = rng.Rows.Count
    Dim keys() As String
    ReDim keys(1 To nRows)
    Dim blank2 As Long, okBulk As Boolean
    Dim rT As Range
    Set rT = rng.Columns(1)
    Dim v2 As Variant
    v2 = sig.ContentVec

    If sig.IsMixed Then
        Dim blk As Variant
        okBulk = True
        For Each blk In sig.Blocks
            If Not modColumnScan.RenderBlockBulk(rT, v2, CLng(blk(0)), CLng(blk(1)), _
                                                 CStr(blk(2)), keys, blank2) Then okBulk = False
        Next blk
    Else
        okBulk = modColumnScan.RenderBlockBulk(rT, v2, 1, nRows, sig.FmtUniform, keys, blank2)
    End If
    LogShim.AssertEq cas & "_bulk_ok", True, okBulk
    If Not okBulk Then Exit Sub

    ' 집계: 현행 누적 규칙과 동일(빈칸 제외 행수 합산, Double만 값지도, 대소문자 무시 키)
    Dim rbk2 As Object
    Set rbk2 = CreateObject("Scripting.Dictionary")
    rbk2.CompareMode = vbTextCompare
    Dim vm2 As Object
    Set vm2 = CreateObject("Scripting.Dictionary")
    Dim i As Long, vk As String
    For i = 1 To nRows
        If Not IsEmpty(v2(i, 1)) Then
            rbk2(keys(i)) = rbk2(keys(i)) + 1
            If Not IsError(v2(i, 1)) Then
                If VarType(v2(i, 1)) = vbDouble Then
                    vk = CStr(v2(i, 1))
                    If vm2.exists(vk) Then
                        If InStr(1, vbNullChar & vm2(vk) & vbNullChar, _
                                 vbNullChar & keys(i) & vbNullChar) = 0 Then
                            vm2(vk) = vm2(vk) & vbNullChar & keys(i)
                        End If
                    Else
                        vm2(vk) = keys(i)
                    End If
                End If
            End If
        End If
    Next i

    ' 전수 대조: 빈칸 수 / 키 집합+행수 / 값지도
    LogShim.AssertEq cas & "_blank", CLng(asset.BlankCount), CLng(blank2)
    LogShim.AssertEq cas & "_nkeys", CLng(asset.Displays.Count), CLng(rbk2.Count)

    Dim k As Variant, mis As String
    For Each k In rbk2.Keys
        If Not asset.RowCountByKey.exists(k) Then
            If Len(mis) < 150 Then mis = mis & "+[" & k & "]"
        ElseIf CLng(asset.RowCountByKey(k)) <> CLng(rbk2(k)) Then
            If Len(mis) < 150 Then mis = mis & "#[" & k & " " & _
                asset.RowCountByKey(k) & "/" & rbk2(k) & "]"
        End If
    Next k
    LogShim.AssertEq cas & "_keys_equiv", "", mis

    Dim vmMis As String
    LogShim.AssertEq cas & "_vm_n", CLng(asset.ValueMap.Count), CLng(vm2.Count)
    For Each k In vm2.Keys
        If Not asset.ValueMap.exists(k) Then
            If Len(vmMis) < 150 Then vmMis = vmMis & "+[" & k & "]"
        ElseIf CStr(asset.ValueMap(k)) <> CStr(vm2(k)) Then
            If Len(vmMis) < 150 Then vmMis = vmMis & "#[" & k & " " & _
                Replace(CStr(asset.ValueMap(k)), vbNullChar, "|") & "/" & _
                Replace(CStr(vm2(k)), vbNullChar, "|") & "]"
        End If
    Next k
    LogShim.AssertEq cas & "_vm_equiv", "", vmMis
End Sub

' 회계 원화(* 패딩 + fmtLocal ! 이스케이프 유발) 영문 서식을 만든다.
Private Function AccWonFmt() As String
    Dim won As String
    won = ChrW(8361)
    AccWonFmt = "_-" & won & "* #,##0_-;-" & won & "* #,##0_-;_-" & won & "* ""-""_-;_-@_-"
End Function

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
