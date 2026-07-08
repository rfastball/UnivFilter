Attribute VB_Name = "Test_InputParse"
' @desc modInputParse 토큰화/값해석 계약 단위테스트 담당
' @depends modInputParse, LogShim
Option Explicit

' 2026-07 재설계 스펙: 실측(gap-probe-findings.md) 기반 신규 계약.

' @desc Tokenize의 구분자/트림/중복제거 계약을 검증함
Public Sub Test_IP_Tokenize()
    Dim arr As Variant

    arr = modInputParse.Tokenize("aa" & vbCrLf & "bb" & vbLf & "cc" & vbTab & "dd")
    LogShim.AssertEq "tok_mixed_sep_count", CLng(4), CLng(UBound(arr) - LBound(arr) + 1)
    LogShim.AssertEq "tok_first", "aa", CStr(arr(0))
    LogShim.AssertEq "tok_last", "dd", CStr(arr(3))

    arr = modInputParse.Tokenize("  x  " & vbCrLf & vbCrLf & "x" & vbCrLf & "X")
    LogShim.AssertEq "tok_trim_dedupe_ci", CLng(1), CLng(UBound(arr) - LBound(arr) + 1)
    LogShim.AssertEq "tok_first_form_kept", "x", CStr(arr(0))

    arr = modInputParse.Tokenize("")
    LogShim.AssertEq "tok_empty_input", CLng(0), CLng(UBound(arr) - LBound(arr) + 1)

    ' 콤마는 구분자가 아니다 ("1,234"는 하나의 토큰)
    arr = modInputParse.Tokenize("1,234")
    LogShim.AssertEq "tok_comma_not_sep", CLng(1), CLng(UBound(arr) - LBound(arr) + 1)
End Sub

' @desc 숫자/콤마/통화/괄호음수 값 해석을 검증함
Public Sub Test_IP_ParseNumber()
    LogShim.AssertEq "parse_plain", True, HasVal("1000", 1000#)
    LogShim.AssertEq "parse_comma", True, HasVal("1,234", 1234#)
    LogShim.AssertEq "parse_decimal", True, HasVal("3.14", 3.14)
    LogShim.AssertEq "parse_negative", True, HasVal("-500", -500#)
    LogShim.AssertEq "parse_won", True, HasVal(ChrW(8361) & "1,234", 1234#)
    LogShim.AssertEq "parse_dollar", True, HasVal("$5,678", 5678#)
    LogShim.AssertEq "parse_paren_neg", True, HasVal("(500)", -500#)
    LogShim.AssertEq "parse_won_paren", True, HasVal("(" & ChrW(8361) & "1,234)", -1234#)
End Sub

' @desc 백분율 값 해석(1/100 스케일)을 검증함
Public Sub Test_IP_ParsePercent()
    LogShim.AssertEq "parse_pct_int", True, HasVal("50%", 0.5)
    LogShim.AssertEq "parse_pct_dec", True, HasVal("12.5%", 0.125)
End Sub

' @desc 날짜/시간/한글날짜 값 해석을 검증함
Public Sub Test_IP_ParseDate()
    LogShim.AssertEq "parse_date_iso", True, HasVal("2024-01-03", CDbl(DateSerial(2024, 1, 3)))
    LogShim.AssertEq "parse_date_slash", True, HasVal("2024/01/03", CDbl(DateSerial(2024, 1, 3)))
    LogShim.AssertEq "parse_date_hangul", True, _
        HasVal("2024" & ChrW(&HB144) & " 1" & ChrW(&HC6D4) & " 3" & ChrW(&HC77C), _
               CDbl(DateSerial(2024, 1, 3)))
    LogShim.AssertEq "parse_time", True, HasVal("08:09", CDbl(TimeSerial(8, 9, 0)))
    LogShim.AssertEq "parse_datetime", True, _
        HasVal("2024-01-03 08:09", CDbl(DateSerial(2024, 1, 3) + TimeSerial(8, 9, 0)))
End Sub

' @desc 해석 불가 토큰이 빈 후보를 반환하는지 검증함
Public Sub Test_IP_ParseNone()
    Dim arr As Variant
    arr = modInputParse.TryParseValues("Apple")
    LogShim.AssertEq "parse_text_none", CLng(0), CandCount(arr)
    arr = modInputParse.TryParseValues("")
    LogShim.AssertEq "parse_empty_none", CLng(0), CandCount(arr)
End Sub

' 후보 배열에 특정 값이 있는지 확인한다.
Private Function HasVal(ByVal sToken As String, ByVal d As Double) As Boolean
    Dim arr As Variant, i As Long
    arr = modInputParse.TryParseValues(sToken)
    For i = LBound(arr) To UBound(arr)
        If Abs(CDbl(arr(i)) - d) < 0.0000001 Then HasVal = True: Exit Function
    Next i
End Function

Private Function CandCount(ByVal arr As Variant) As Long
    CandCount = UBound(arr) - LBound(arr) + 1
    If CandCount < 0 Then CandCount = 0
End Function
