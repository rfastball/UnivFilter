Attribute VB_Name = "Test_Normalize"
' @desc modNormalize 값 정규화 계약 단위테스트 담당
' @depends modNormalize, modIndexKey, LogShim
Option Explicit

' L1 분리(Stage 5) 시점의 동작 계약 고정 - 개선 목적의 어서션 수정 금지.

' @desc NormalizeValue의 숫자 계열 정규화(대시/기호/콤마/공격 추출)를 검증함
Public Sub Test_Norm_NormalizeValue_Number()
    LogShim.AssertEq "norm_num_comma", CDbl(3000), modNormalize.NormalizeValue("3,000", fdNumber)
    LogShim.AssertEq "norm_num_acc_dash_zero", CDbl(0), modNormalize.NormalizeValue("\-", fdNumber)
    LogShim.AssertEq "norm_num_won_strip", CDbl(1000), modNormalize.NormalizeValue(ChrW(8361) & "1,000", fdNumber)
    LogShim.AssertEq "norm_num_dollar_strip", CDbl(1000), modNormalize.NormalizeValue("$1,000", fdNumber)
    LogShim.AssertEq "norm_num_nbsp_strip", CDbl(50), modNormalize.NormalizeValue(ChrW(160) & "50", fdNumber)
    LogShim.AssertEq "norm_num_aggressive", CDbl(-12.5), modNormalize.NormalizeValue("Val:-12.5", fdNumber)

    ' 숫자 추출 불가면 원본 유지 (실패 바스켓 판정은 호출자 소관)
    LogShim.AssertEq "norm_num_fail_passthrough", "abc", modNormalize.NormalizeValue("abc", fdNumber)
End Sub

' @desc NormalizeValue의 퍼센트/문자열/빈 입력 처리를 검증함
Public Sub Test_Norm_NormalizeValue_PercentString()
    LogShim.AssertEq "norm_pct_scaled", CDbl(0.5), modNormalize.NormalizeValue("50%", fdPercent)
    LogShim.AssertEq "norm_pct_raw", CDbl(50), modNormalize.NormalizeValue("50", fdPercent)
    LogShim.AssertEq "norm_string_trim", "abc", modNormalize.NormalizeValue("  abc  ", fdString)
    LogShim.AssertEq "norm_empty_input", "", modNormalize.NormalizeValue("", fdString)
    LogShim.AssertEq "norm_date_as_string", "2025-12-25", modNormalize.NormalizeValue("2025-12-25", fdDate)
End Sub

' @desc ExtractNumberAggressive의 추출/통과 계약을 검증함
Public Sub Test_Norm_ExtractNumberAggressive()
    LogShim.AssertEq "extract_negative_mid", CDbl(-12.5), modNormalize.ExtractNumberAggressive("Val:-12.5")
    LogShim.AssertEq "extract_comma_currency", CDbl(1000), modNormalize.ExtractNumberAggressive("$1,000")
    LogShim.AssertEq "extract_nonnumeric_passthrough", "abc", modNormalize.ExtractNumberAggressive("abc")
End Sub

' @desc SafeText와 modIndexKey.NormalizeIndexKey 계약을 검증함
Public Sub Test_Norm_SafeText_IndexKey()
    LogShim.AssertEq "normsafe_error_marker", "#ERROR", modNormalize.SafeText(CVErr(2042))
    LogShim.AssertEq "normsafe_null_empty", "", modNormalize.SafeText(Null)
    LogShim.AssertEq "normsafe_empty", "", modNormalize.SafeText(Empty)
    LogShim.AssertEq "normsafe_number_cstr", "42", modNormalize.SafeText(42)

    ' 벤더링된 키 계약 (UnivDB modIndexKey): Empty -> "", 그 외 CStr
    LogShim.AssertEq "indexkey_empty", "", modIndexKey.NormalizeIndexKey(Empty)
    LogShim.AssertEq "indexkey_cstr_number", "42", modIndexKey.NormalizeIndexKey(42)
    LogShim.AssertEq "indexkey_composite_prefix", "3:abc", modIndexKey.EncodeCompositeKeyPart("abc")
End Sub
