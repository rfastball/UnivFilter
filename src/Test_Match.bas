Attribute VB_Name = "Test_Match"
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
    Set r = modMatch.Execute(Array("0.5"), asset, False)
    LogShim.AssertEq "mt_bridge_twins", CLng(2), r("Matched").Count
    LogShim.AssertEq "mt_bridge_corr", True, r("Corrections").exists("0.5")

    ' "1000" -> "1,000" 자동보정
    Set r = modMatch.Execute(Array("1000"), asset, False)
    LogShim.AssertEq "mt_bridge_comma", True, r("Matched").Contains("1,000")
    LogShim.AssertEq "mt_bridge_comma_corr", "1,000", CStr(r("Corrections")("1000"))

    ' 표시 레인 적중 시 값 레인 미가동 (0.50은 보정 없이 그 표시만)
    Set r = modMatch.Execute(Array("0.50"), asset, False)
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

    Dim vm As Object
    Set vm = CreateObject("Scripting.Dictionary")

    Dim sHalf As New clsSet
    sHalf.Add "0.50"
    sHalf.Add "50%"
    Set vm(CStr(0.5)) = sHalf

    Dim sThousand As New clsSet
    sThousand.Add "1,000"
    Set vm(CStr(1000#)) = sThousand

    Dim asset As New clsColAsset
    asset.Init displays, vm, True, 5, "TEST!A1", 1
    Set BuildTestAsset = asset
End Function

Private Function MissingCount(ByVal r As Object) As Long
    Dim arr As Variant
    arr = r("Missing")
    MissingCount = UBound(arr) - LBound(arr) + 1
    If MissingCount < 0 Then MissingCount = 0
End Function
