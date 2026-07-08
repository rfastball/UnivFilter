Attribute VB_Name = "modDisplayText"
' @desc 값과 서식 받아 필터용 표시 텍스트 재구성 담당
' @depends modNormalize, modLogger
Option Explicit

' ==============================================================================
' [L2b 표시 텍스트 재구성] modFilterUtils에서 이동 (Stage 6)
' 검증된 로직의 이동이므로 알고리즘 수정 금지 - 계약은 Test_FilterUtilsChar가 고정
' ==============================================================================

' ==============================================================================
' 함수명: GetSafeFormattedText
' 설명: 값(vVal)에 서식(sFmt)을 적용하되, 에러 발생 시 단계적으로 덜 정교한 메서드를 시도함
' 우선순위:
'   1. WorksheetFunction.Text (엑셀 고유 서식 완벽 지원)
'   2. VBA Format 함수 (표준 서식 지원)
'   3. 기본 날짜 포맷 (yyyy-mm-dd)
'   4. 단순 문자열 변환 (CStr)
' ==============================================================================
' @desc 값과 서식 받아 표시 문자열 반환, 실패 시 단계 폴백
Public Function GetSafeFormattedText(ByVal vVal As Variant, ByVal sFmt As String) As String
    Dim sResult As String
    Dim bSuccess As Boolean

    bSuccess = False
    sResult = ""

    ' [예외 처리] 값이 에러(#N/A 등)이거나 비어있는 경우
    If IsError(vVal) Then
        GetSafeFormattedText = CStr(vVal)
        Exit Function
    End If
    If IsEmpty(vVal) Then
        GetSafeFormattedText = ""
        Exit Function
    End If

    ' [1단계] 엑셀 워크시트 함수 사용 (가장 정교함)
    If sFmt <> "" Then
        On Error Resume Next
        sResult = Application.WorksheetFunction.text(vVal, sFmt)

        If Err.Number = 0 Then
            bSuccess = True
        Else
            Err.Clear ' 에러 초기화
        End If
        On Error GoTo 0
    End If

    ' [2단계] VBA Format 함수 사용 (1단계 실패 시)
    If Not bSuccess And sFmt <> "" Then
        On Error Resume Next
        sResult = Format(vVal, sFmt)

        If Err.Number = 0 Then
            bSuccess = True
        Else
            ' [LOG] 서식 적용 실패 (드문 경우라 로그 남김)
            Log "GetSafeFormattedText: Format 실패 (Val: " & vVal & ", Fmt: " & sFmt & ")", lvlDebug
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' [3단계 & 4단계] 기본값 처리 (모두 실패하거나 서식이 없는 경우)
    If Not bSuccess Then
        If IsDate(vVal) Then
            sResult = Format(vVal, "yyyy-mm-dd") ' ISO 표준 날짜
        Else
            sResult = CStr(vVal) ' 단순 문자열
        End If
    End If

    GetSafeFormattedText = sResult
End Function


' ==========================================================
' [ GetCleanFormat ]
' 목적: 엑셀 AutoFilter가 '인식 가능한' 표준 문자열 생성을 위한 서식 정제
' ----------------------------------------------------------
' 상세 사유 (Excel Filter Logic 대응):
'   1. 엑셀 필터는 '값'이 아닌 '화면에 표시된 텍스트'와 매칭하여 작동함
'   2. 회계 서식의 특수 기호(*, _)는 실제 텍스트에 불규칙한 공백을 만듦.
'   3. 그런데 엑셀 필터는 또 이런 불규칙한 공백은 포함되면 매칭이 안됨(...)
'   4. 따라서 회계서식이 걸린 필터에 주는 값은 원본서식이 아니라 원본서식에서 불규칙한 공백이 제거된 서식이어야 함
'   5. 본 함수는 이런 '정렬용 노이즈'를 제거하여 유연한 매칭(VisFail 보정)이
'      가능한 깨끗한 서식(예: #,##0.00)만 추출함.
'   6. 이렇게 변환된 값은 GetAccountingFilterString로 넘겨서 필터에 먹일 string으로 변환함
' ==========================================================
' @desc 원시 서식 받아 필터 인식용 정제 서식 반환, General/@은 빈 문자열
Public Function GetCleanFormat(ByVal rawFmt As String) As String
    Dim s As String: s = Trim(rawFmt)

    ' [LOG] 서식 클리닝 진입 (디버그)
    ' Log "GetCleanFormat: " & rawFmt, lvlDebug

    ' [1] 안전장치: 일반/텍스트 서식은 건드리지 않음
    If s = "General" Or s = "@" Or s = "" Then GetCleanFormat = "": Exit Function

    ' [2] 양수/음수 서식 분리: 세미콜론이 있으면 앞부분(양수)만 사용
    If InStr(s, ";") > 0 Then s = Split(s, ";")(0)

    ' [3] 회계/특수 서식 기호(*, _) 처리 로직 (핵심 개선 구간)
    If InStr(s, "*") > 0 Or InStr(s, "_") > 0 Then

        Dim sSymbol As String: sSymbol = ""
        Dim sDec As String: sDec = ""

        ' A. 통화 기호 살리기
        ' -------------------------------------------------------
        If InStr(s, "\") > 0 Or InStr(s, "\") > 0 Then
            sSymbol = "\"
        ElseIf InStr(s, "$") > 0 Then
            sSymbol = "$"
        ElseIf InStr(s, "€") > 0 Then
            sSymbol = "€"
        ElseIf InStr(s, "￡") > 0 Then
            sSymbol = "￡"
        End If
        ' -------------------------------------------------------

        ' B. 소수점 자릿수 확인 (기존 로직 포괄)
        If InStr(s, ".00") > 0 Then
            sDec = ".00"
        ElseIf InStr(s, ".0") > 0 Then
            sDec = ".0"
        End If

        ' C. 서식 재조립
        ' 기호가 없으면 sSymbol은 빈 값이므로, 기존 로직(순수 숫자)과 똑같이 동작함
        GetCleanFormat = sSymbol & "#,##0" & sDec

        Log "GetCleanFormat: 회계 서식 정제 -> " & GetCleanFormat, lvlDebug
        Exit Function
    End If

    ' [4] 그 외 서식(단위 등)은 최대한 원본 유지
    GetCleanFormat = s
End Function

' ==========================================================
' [ GetAccountingFilterString ]
' 목적: 회계 서식 특유의 '시각적 변환'을 엑셀 필터가 인식하는 '패턴'으로 정규화
' ----------------------------------------------------------
' 핵심 기능:
'   1. Zero Value 보정: 숫자 0을 회계용 대시(예: \-) 문자로 변환
'   2. 공백 압축: 기호(Currency)와 숫자 사이의 가변 공백을 제거하여 밀착시킴
'   3. 음수 패턴 매칭: 서식 설정에 따라 괄호형() 또는 마이너스형(-) 패턴 생성
' ==========================================================
' @desc 값과 회계 서식 받아 필터용 문자열 반환, 숫자 추출 실패 시 원본
Public Function GetAccountingFilterString(ByVal vValue As Variant, ByVal sFormat As String) As String
    Dim sSymbol As String: sSymbol = ""
    Dim sCleanNum As String: sCleanNum = modNormalize.ExtractNumberAggressive(CStr(vValue))
    Dim dblVal As Double

    If Not IsNumeric(sCleanNum) Then
        GetAccountingFilterString = CStr(vValue): Exit Function
    End If
    dblVal = CDbl(sCleanNum)

    ' 1. 기호 판별 (로캘 원화 기호인 역슬래시 사용)
    If InStr(sFormat, "$") > 0 Or InStr(sFormat, "\") > 0 Or _
       InStr(sFormat, ChrW(8361)) > 0 Or InStr(sFormat, "[$") > 0 Then
        sSymbol = Chr(92)
    End If

    ' 2. [핵심] 제로(0) 처리: 사용자님이 말씀하신 \- 형태
    If dblVal = 0 Then
        ' 회계 서식이라면 \- , 아니면 그냥 0
        If InStr(sFormat, "*") > 0 Then
            GetAccountingFilterString = sSymbol & "-"
        Else
            GetAccountingFilterString = "0"
        End If
        Log "GetAccountingFilterString: Zero Value -> " & GetAccountingFilterString, lvlDebug
        Exit Function
    End If

    ' 3. 음수/양수 조립 (와일드카드 삭제, 기호 밀착)
    If dblVal < 0 Then
        If InStr(sFormat, "(") > 0 Then
            ' 괄호형: (\1,234)
            GetAccountingFilterString = "(" & sSymbol & GetSafeFormattedText(Abs(dblVal), "#,##0") & ")"
        Else
            ' 마이너스형: -\1,234 (사용자님 스크린샷의 -\1,234 선택 반영)
            GetAccountingFilterString = "-" & sSymbol & GetSafeFormattedText(Abs(dblVal), "#,##0")
        End If
    Else
        ' 양수: \50,000
        GetAccountingFilterString = sSymbol & GetSafeFormattedText(dblVal, "#,##0")
    End If

    ' [LOG] 최종 변환 결과 확인 (필터링 실패 시 원인 파악용)
    Log "GetAccountingFilterString: " & vValue & " -> " & GetAccountingFilterString, lvlDebug
End Function

' @desc 값과 퍼센트 서식 받아 필터용 문자열 반환, 비숫자는 원본
Public Function GetPercentageFilterString(ByVal vValue As Variant, ByVal sFormat As String) As String
    Dim sCleanNum As String: sCleanNum = modNormalize.ExtractNumberAggressive(CStr(vValue))

    ' 1. 숫자가 아니면 원본 반환 (예: 사용자가 "%"만 입력해 검색할 경우)
    If Not IsNumeric(sCleanNum) Then
        GetPercentageFilterString = CStr(vValue)
        Exit Function
    End If

    ' 2. 숫자인 경우, 해당 열의 서식을 적용하여 문자열 생성
    ' 예: 0.5 -> "50%", 0.123 -> "12.3%" (서식에 따라 다름)
    GetPercentageFilterString = GetSafeFormattedText(CDbl(sCleanNum), sFormat)

    ' [LOG] 퍼센트 변환 확인
    Log "GetPercentageFilterString: " & vValue & " -> " & GetPercentageFilterString, lvlDebug
End Function
