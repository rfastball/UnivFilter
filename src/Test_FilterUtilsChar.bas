Attribute VB_Name = "Test_FilterUtilsChar"
' @desc 표시 텍스트/정제 함수 특성화 단위테스트 담당
' @depends modDisplayText, modNormalize, LogShim
Option Explicit

' 타입 로직 4분해(Stage 5~6) 전후 동작 불변을 보증하는 특성화 테스트.
' 어서션은 현행 동작의 기록 - assert 이름은 모듈 이동 후에도 유지할 것.

Private Const ACC_FMT_SYM As String = "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""_-;_-@_-"
Private Const ACC_FMT_NOSYM As String = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"

' @desc GetCleanFormat의 회계 기호/소수점/세미콜론 분할 처리를 검증함
Public Sub Test_Char_GetCleanFormat()
    ' 일반/텍스트/빈 서식은 빈 문자열로
    LogShim.AssertEq "cleanfmt_general", "", modDisplayText.GetCleanFormat("General")
    LogShim.AssertEq "cleanfmt_text", "", modDisplayText.GetCleanFormat("@")
    LogShim.AssertEq "cleanfmt_empty", "", modDisplayText.GetCleanFormat("")

    ' 회계 서식(*, _)은 기호 + #,##0 (+소수점)으로 재조립
    LogShim.AssertEq "cleanfmt_acc_symbol", "$#,##0", modDisplayText.GetCleanFormat(ACC_FMT_SYM)
    LogShim.AssertEq "cleanfmt_acc_nosymbol", "#,##0", modDisplayText.GetCleanFormat(ACC_FMT_NOSYM)
    LogShim.AssertEq "cleanfmt_acc_decimal", "#,##0.00", modDisplayText.GetCleanFormat("_-* #,##0.00_-;-* #,##0.00_-")

    ' 세미콜론 분할: 양수부만 사용 (회계 기호 없으면 원본 유지 경로)
    LogShim.AssertEq "cleanfmt_semicolon_positive", "$#,##0", modDisplayText.GetCleanFormat("$#,##0;[Red]-$#,##0")

    ' 단위 등 그 외 서식은 원본 유지
    LogShim.AssertEq "cleanfmt_passthrough_unit", "0"" kg""", modDisplayText.GetCleanFormat("0"" kg""")
End Sub

' @desc GetSafeFormattedText의 4단 폴백(Text/Format/날짜기본/CStr)을 검증함
Public Sub Test_Char_GetSafeFormattedText()
    LogShim.AssertEq "safetext_comma", "3,000", modDisplayText.GetSafeFormattedText(3000, "#,##0")
    LogShim.AssertEq "safetext_percent", "50%", modDisplayText.GetSafeFormattedText(0.5, "0%")
    LogShim.AssertEq "safetext_date_fmt", "2025-12-25", modDisplayText.GetSafeFormattedText(DateSerial(2025, 12, 25), "yyyy-mm-dd")

    ' 서식 없음: 날짜는 ISO 기본, 그 외는 CStr
    LogShim.AssertEq "safetext_date_default", "2025-12-25", modDisplayText.GetSafeFormattedText(DateSerial(2025, 12, 25), "")
    LogShim.AssertEq "safetext_plain_cstr", "123", modDisplayText.GetSafeFormattedText(123, "")

    ' 에러 값은 CStr로 즉시 반환, Empty는 빈 문자열
    LogShim.AssertEq "safetext_error_passthrough", CStr(CVErr(2042)), modDisplayText.GetSafeFormattedText(CVErr(2042), "#,##0")
    LogShim.AssertEq "safetext_empty", "", modDisplayText.GetSafeFormattedText(Empty, "#,##0")
End Sub

' @desc GetAccountingFilterString의 0 대시/음수 괄호/기호 밀착을 검증함
Public Sub Test_Char_GetAccountingFilterString()
    ' 회계 서식의 0은 기호+대시, 비회계 서식의 0은 "0"
    LogShim.AssertEq "accstr_zero_dash", "\-", modDisplayText.GetAccountingFilterString(0, ACC_FMT_SYM)
    LogShim.AssertEq "accstr_zero_plain", "0", modDisplayText.GetAccountingFilterString(0, "#,##0")

    ' 음수: 서식에 괄호가 있으면 괄호형, 없으면 마이너스형 (기호 밀착)
    LogShim.AssertEq "accstr_neg_paren", "(\1,234)", modDisplayText.GetAccountingFilterString(-1234, "$#,##0;($#,##0)")
    LogShim.AssertEq "accstr_neg_minus", "-\1,234", modDisplayText.GetAccountingFilterString(-1234, ACC_FMT_SYM)

    ' 양수: 기호 밀착 + 천단위 콤마
    LogShim.AssertEq "accstr_positive", "\50,000", modDisplayText.GetAccountingFilterString(50000, ACC_FMT_SYM)

    ' 숫자 추출 불가면 원본 그대로
    LogShim.AssertEq "accstr_nonnumeric", "abc", modDisplayText.GetAccountingFilterString("abc", ACC_FMT_SYM)
End Sub

' @desc GetPercentageFilterString의 서식 적용/비숫자 통과를 검증함
Public Sub Test_Char_GetPercentageFilterString()
    LogShim.AssertEq "pctstr_int", "50%", modDisplayText.GetPercentageFilterString(0.5, "0%")
    LogShim.AssertEq "pctstr_decimal", "12.3%", modDisplayText.GetPercentageFilterString(0.123, "0.0%")
    LogShim.AssertEq "pctstr_nonnumeric", "%", modDisplayText.GetPercentageFilterString("%", "0%")
End Sub

' @desc GetUniqueCleanData의 콤마/퍼센트 정규화와 이중 바스켓을 검증함
Public Sub Test_Char_GetUniqueCleanData()
    Dim dict As Object

    ' 숫자: 천단위 콤마 제거 후 CDbl 키, 값은 원본 보존
    Set dict = modNormalize.GetUniqueCleanData(Array("3,000", "50"), fdNumber)
    LogShim.AssertEq "ucd_num_count", 2, dict.Count
    LogShim.AssertEq "ucd_num_comma_key", True, dict.exists(CDbl(3000))
    LogShim.AssertEq "ucd_num_original_kept", "3,000", dict(CDbl(3000))

    ' 회계 대시 표기는 0으로
    Set dict = modNormalize.GetUniqueCleanData(Array("\-"), fdNumber)
    LogShim.AssertEq "ucd_num_acc_dash_zero", True, dict.exists(CDbl(0))

    ' 퍼센트: % 기호가 있으면 /100, 없으면 그대로
    Set dict = modNormalize.GetUniqueCleanData(Array("50%"), fdPercent)
    LogShim.AssertEq "ucd_pct_scaled", True, dict.exists(CDbl(0.5))
    Set dict = modNormalize.GetUniqueCleanData(Array("50"), fdPercent)
    LogShim.AssertEq "ucd_pct_raw", True, dict.exists(CDbl(50))

    ' 변환 실패는 실패 바스켓: Key=원본, Value=False
    Set dict = modNormalize.GetUniqueCleanData(Array("abc"), fdNumber)
    LogShim.AssertEq "ucd_fail_basket_flag", False, dict("abc")

    ' 멀티라인 문자열 입력은 줄 단위로 분해
    Set dict = modNormalize.GetUniqueCleanData("a" & vbLf & "b", fdString)
    LogShim.AssertEq "ucd_multiline_split", 2, dict.Count
End Sub
