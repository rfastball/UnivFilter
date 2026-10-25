Attribute VB_Name = "Test_Match"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc modMatch 2레인 매칭 계약 단위테스트 담당
' @depends modMatch, modInputParse, clsColAsset, clsSet, LogShim
Option Explicit

' 2026-07 재설계 스펙: 인메모리 자산으로 시트 없이 매칭 로직을 고정한다.
' 자산: 표시 {"0.50","50%","1,000","Apple"}, 값지도 {0.5->[0.50,50%], 1000->[1,000]}

' @desc 표시 레인 완전일치(대소문자 무시 포함)를 검증함
Public Sub Test_MT_DisplayExact()
    Dim asset As clsColAsset
    Set asset = BuildTestAsset()

    Dim r As Object
    Set r = modMatch.Execute(Array("0.50"), asset, False)
    LogShim.AssertEq "mt_disp_matched", CLng(1), r("Matched").Count
    LogShim.AssertEq "mt_disp_no_corr", CLng(0), r("Corrections").Count
    LogShim.AssertEq "mt_disp_no_missing", CLng(0), MissingCount(r)

    Set r = modMatch.Execute(Array("apple"), asset, False)
    LogShim.AssertEq "mt_disp_ci_hit", CLng(1), r("Matched").Count
End Sub

' @desc 값 레인 브리지(값쌍둥이 전개 + 자동보정 기록)를 검증함
Public Sub Test_MT_ValueBridge()
    Dim asset As clsColAsset
    Set asset = BuildTestAsset()

    ' "0.5"는 표시에 없음 -> 값 0.5의 표시쌍둥이 2개로 전개 (FACT 3 의미론)
    Dim r As Object
    Set r = modMatch.Execute(Array("0.5"), asset, False, True)
    LogShim.AssertEq "mt_bridge_twins", CLng(2), r("Matched").Count
    LogShim.AssertEq "mt_bridge_corr", True, r("Corrections").exists("0.5")

    ' "1000" -> "1,000" 자동보정
    Set r = modMatch.Execute(Array("1000"), asset, False, True)
    LogShim.AssertEq "mt_bridge_comma", True, r("Matched").Contains("1,000")
    LogShim.AssertEq "mt_bridge_comma_corr", "1,000", CStr(r("Corrections")("1000"))

    Set r = modMatch.Execute(Array("1000", "0.5"), asset, False)
    LogShim.AssertEq "mt_bridge_default_off", CLng(0), r("Matched").Count
    LogShim.AssertEq "mt_bridge_off_missing", CLng(2), MissingCount(r)
    LogShim.AssertEq "mt_bridge_off_no_corr", CLng(0), r("Corrections").Count

    Set r = modMatch.Execute(Array("1000"), asset, True, True)
    LogShim.AssertEq "mt_partial_never_corrects", CLng(0), r("Matched").Count

    ' 표시 레인 적중 시 값 레인 미가동 (0.50은 보정 없이 그 표시만)
    Set r = modMatch.Execute(Array("0.50"), asset, False, True)
    LogShim.AssertEq "mt_display_wins", CLng(1), r("Matched").Count
End Sub

' @desc 누락 판정과 혼합 입력 분류를 검증함
Public Sub Test_MT_Missing()
    Dim asset As clsColAsset
    Set asset = BuildTestAsset()

    Dim r As Object
    Set r = modMatch.Execute(Array("999", "Apple", "zzz"), asset, False)
    LogShim.AssertEq "mt_missing_count", CLng(2), MissingCount(r)
    LogShim.AssertEq "mt_missing_matched", CLng(1), r("Matched").Count
    LogShim.AssertEq "mt_missing_first", "999", CStr(r("Missing")(0))
End Sub

' @desc 부분일치 스캔(적용 키는 여전히 정확한 표시 키)을 검증함
Public Sub Test_MT_Partial()
    Dim asset As clsColAsset
    Set asset = BuildTestAsset()

    Dim r As Object
    Set r = modMatch.Execute(Array("app"), asset, True)
    LogShim.AssertEq "mt_partial_hit", CLng(1), r("Matched").Count
    LogShim.AssertEq "mt_partial_key", True, r("Matched").Contains("Apple")

    ' "0"을 포함하는 표시: 0.50 / 50% / 1,000
    Set r = modMatch.Execute(Array("0"), asset, True)
    LogShim.AssertEq "mt_partial_multi", CLng(3), r("Matched").Count

    Set r = modMatch.Execute(Array("zzz"), asset, True)
    LogShim.AssertEq "mt_partial_missing", CLng(1), MissingCount(r)
End Sub

' @desc 혼합 부분일치의 겹침과 누락 순서 및 폴백을 검증함
Public Sub Test_MT_PartialHybrid()
    Dim displays As clsSet, asset As clsColAsset, emptyMap As Object, actual As Object
    Dim expected As clsSet, keys As Variant, tokens() As Variant, hit() As Boolean
    Dim spec As Variant, specs As Variant, width As Long, n As Long, i As Long, j As Long
    Dim missing As String, actualMissing As Variant, specials As Variant, value As Variant
    specs = Array(Array(15, 9), Array(15, 10), Array(64, 19), Array(64, 20), _
                  Array(128, 49), Array(128, 50), Array(15, 11), Array(15, 12), Array(15, 13))
    specials = Array("ushers", "SHE", "aaaa", "ababa", "ABC", "한글ABC", _
                     ChrW(&HFF21) & "BC", "a" & vbNullChar & "b", "", "I", ChrW(&H130), ChrW(&H131))
    For Each spec In specs
        width = spec(0): n = spec(1)
        Set displays = New clsSet
        For i = 1 To 1025
            displays.Add String$(width - 6, "x") & Format$(i, "000000")
        Next i
        For Each value In specials
            displays.Add value
        Next value
        Set emptyMap = CreateObject("Scripting.Dictionary")
        Set asset = New clsColAsset
        asset.Init displays, emptyMap, emptyMap, 0, displays.Count, "혼합 단위검증", Nothing
        ReDim tokens(1 To n)
        For i = 1 To n
            tokens(i) = "MISS" & CStr(i)
        Next i
        tokens(1) = "he": tokens(2) = "she": tokens(3) = "hers"
        tokens(4) = "a": tokens(5) = "aa": tokens(6) = "aba"
        tokens(7) = "ABC": tokens(8) = "abc": tokens(9) = "i"
        If n = 11 Then tokens(10) = "한글"
        If n = 12 Then tokens(10) = ""
        If n = 13 Then tokens(10) = String$(8193, "x")
        keys = displays.KeysArray
        ReDim hit(1 To n)
        Set expected = New clsSet
        For i = LBound(keys) To UBound(keys)
            For j = 1 To n
                If InStr(1, CStr(keys(i)), CStr(tokens(j)), vbTextCompare) > 0 Then
                    expected.Add keys(i)
                    hit(j) = True
                End If
            Next j
        Next i
        missing = ""
        For j = 1 To n
            If Not hit(j) Then missing = missing & CStr(tokens(j)) & vbNullChar
        Next j
        Set actual = modMatch.Execute(tokens, asset, True)
        LogShim.AssertEq "mt_hybrid_count", expected.Count, actual("Matched").Count
        LogShim.AssertEq "mt_hybrid_keys", CLng(0), expected.Difference(actual("Matched")).Count
        actualMissing = actual("Missing")
        actualMissing = Join(actualMissing, vbNullChar)
        If Len(actualMissing) > 0 Then actualMissing = actualMissing & vbNullChar
        LogShim.AssertEq "mt_hybrid_missing_order", missing, CStr(actualMissing)
        LogShim.AssertEq "mt_hybrid_inputs", n, actual("InputCount")
        LogShim.AssertEq "mt_hybrid_rows", CLng(-1), actual("MatchedRows")
        LogShim.AssertEq "mt_hybrid_corrections", CLng(0), actual("Corrections").Count
    Next spec
    Set actual = modMatch.Execute(Array(), asset, True)
    LogShim.AssertEq "mt_hybrid_empty", CLng(0), actual("Matched").Count
End Sub

' @desc 제외 여집합(표시 집합 수준)을 검증함
Public Sub Test_MT_ExcludeComplement()
    Dim asset As clsColAsset
    Set asset = BuildTestAsset()

    Dim r As Object
    Set r = modMatch.Execute(Array("Apple"), asset, False)

    Dim comp As clsSet
    Set comp = asset.Displays.Difference(r("Matched"))
    LogShim.AssertEq "mt_comp_count", CLng(3), comp.Count
    LogShim.AssertEq "mt_comp_excluded", False, comp.Contains("Apple")
    LogShim.AssertEq "mt_comp_kept", True, comp.Contains("1,000")
End Sub

' 공용 테스트 자산을 만든다.
Private Function BuildTestAsset() As clsColAsset
    Dim displays As New clsSet
    displays.Add "0.50"
    displays.Add "50%"
    displays.Add "1,000"
    displays.Add "Apple"

    ' 값 지도 값은 vbNullChar 연결 문자열 (modColumnScan 성능 계약)
    Dim vm As Object
    Set vm = CreateObject("Scripting.Dictionary")
    vm(CStr(0.5)) = "0.50" & vbNullChar & "50%"
    vm(CStr(1000#)) = "1,000"

    Dim rbk As Object
    Set rbk = CreateObject("Scripting.Dictionary")
    rbk.CompareMode = vbTextCompare
    rbk("0.50") = 1
    rbk("50%") = 1
    rbk("1,000") = 1
    rbk("Apple") = 1

    Dim asset As New clsColAsset
    asset.Init displays, vm, rbk, 1, 5, "TEST!A1", Nothing
    Set BuildTestAsset = asset
End Function

Private Function MissingCount(ByVal r As Object) As Long
    Dim arr As Variant
    arr = r("Missing")
    MissingCount = UBound(arr) - LBound(arr) + 1
    If MissingCount < 0 Then MissingCount = 0
End Function
