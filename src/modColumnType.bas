Attribute VB_Name = "modColumnType"
' @desc 열 서식/값 스캔해 데이터 타입 판별 담당
' @depends modLogger
Option Explicit

' ==============================================================================
' [L2a 서식 인지 타입 판별] modFilterUtils에서 이동 (Stage 6)
' 검증된 로직의 이동이므로 알고리즘 수정 금지
' ==============================================================================

' ==============================================================================
' [섹션 0] 데이터 타입 열거형 정의 (수정됨)
' ==============================================================================
Public Enum FilterDataType
    ' [기본/미확인]
    fdAuto = 0          ' 초기 상태 또는 분석 실패

    ' [기본 타입]
    fdString = 1        ' 문자열 (STRING)
    fdDate = 3          ' 날짜 (DATE)

    ' [숫자 계열 - 통합]
    fdNumber = 2        ' 숫자 (NUMBER) - 정수/실수 통합
    fdPercent = 4       ' 백분율 (PERCENT) - 정수/실수 통합

    ' [주의 필요 - New]
    fdRounded = 5       ' 숨겨진 소수 (ROUNDED) - 값은 실수이나 서식은 정수
    fdPercentRounded = 6 ' 숨겨진 백분율 (PERCENT_ROUNDED) - 값은 소수%이나 서식은 정수%

    ' [특수 상태]
    fdHybrid = 9        ' 복합 데이터 (HYBRID) - 텍스트 취급 권장
    fdError = 99        ' 에러 데이터 (ERROR)
End Enum

' [도우미] 타입 번호에 따른 이름 반환 (수정됨)
' @desc 타입 enum 받아 한국어 타입 이름 반환, 미확인 시 텍스트
Public Function GetTypeName(eType As FilterDataType) As String
    Select Case eType
        Case fdDate:            GetTypeName = "날짜"
        Case fdNumber:          GetTypeName = "숫자"          ' 통합됨
        Case fdPercent:         GetTypeName = "백분율"        ' 통합됨

        Case fdRounded:         GetTypeName = "숫자(숨김)"    ' [New]
        Case fdPercentRounded:  GetTypeName = "백분율(숨김)"  ' [New]

        Case fdHybrid:          GetTypeName = "복합(혼용)"
        Case fdError:           GetTypeName = "오류"
        Case Else:              GetTypeName = "텍스트"
    End Select
End Function

' [도우미] 식별자 문자열(Identity)을 내부 Enum 타입으로 매핑 (구 modMatchingEngine.MapIdentityToEnum)
' @desc 정체성 문자열 받아 FilterDataType 반환, 미확인 시 fdAuto
Public Function IdentityToEnum(sID As String) As FilterDataType
    Select Case UCase(sID) ' 안전을 위해 대문자 통일
        ' [기본 숫자 계열]
        Case "NUMBER":           IdentityToEnum = fdNumber
        Case "PERCENT":          IdentityToEnum = fdPercent

        ' [주의: 숨겨진 값 계열 - New]
        Case "ROUNDED":          IdentityToEnum = fdRounded
        Case "PERCENT_ROUNDED":  IdentityToEnum = fdPercentRounded

        ' [기타 타입]
        Case "DATE":             IdentityToEnum = fdDate
        Case "STRING":           IdentityToEnum = fdString

        ' [특수 상태]
        Case "HYBRID":           IdentityToEnum = fdHybrid
        Case "ERROR":            IdentityToEnum = fdError

        ' [기본값]
        Case Else:               IdentityToEnum = fdAuto
    End Select
End Function

' ------------------------------------------------------------------------------
' 함수명: GetRepresentativeFormat
' 목적: 범위 내에서 데이터가 존재하는 첫 번째 셀을 찾아 표시 형식을 추출
' 포맷이 단일하다고 확인된 열에대해 구체적 포맷값을 얻을 때
' ------------------------------------------------------------------------------
' @desc 범위 받아 첫 유효 셀의 표시 형식 반환, 빈 범위 시 General
Public Function GetRepresentativeFormat(ByVal rngData As Range) As String
    Dim rngSample As Range
    If rngData Is Nothing Then GetRepresentativeFormat = "General": Exit Function

    ' 빈 셀을 건너뛰고 실제 값이 있는 첫 번째 셀 탐색
    Set rngSample = rngData.Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext)

    If rngSample Is Nothing Then
        GetRepresentativeFormat = rngData.Cells(1, 1).NumberFormat
    Else
        GetRepresentativeFormat = rngSample.NumberFormat
    End If
    ' [LOG] 대표 서식 추출 (디버깅용)
    ' Log "GetRepresentativeFormat: " & GetRepresentativeFormat, lvlDebug
End Function

' ------------------------------------------------------------------------------
' 함수명: GetColumnIdentity
' 목적: 열 전체 데이터를 스캔하여 데이터 타입(FLOAT, NUMBER, PERCENT 등)을 최종 판정
' ------------------------------------------------------------------------------
' @desc 열 범위 받아 정체성 문자열 반환, 데이터 없으면 STRING
Public Function GetColumnIdentity(targetCol As Range) As String
    Dim ws As Worksheet: Set ws = targetCol.Worksheet
    Dim headerRow As Long, lastRow As Long
    Dim dataRange As Range, dataArray As Variant
    Dim sampleCell As Range
    Dim fmt As String
    Dim colIdx As Long: colIdx = targetCol.Column

    ' [설정] 부동소수점 오차 허용 범위
    Const EPSILON As Double = 0.000000000001

    ' [LOG] 컬럼 분석 시작
    ' Log "GetColumnIdentity: 분석 시작 (Col: " & colIdx & ")", lvlDebug

    ' 1단계: 헤더 및 데이터 범위 확정
    ' ---------------------------------------------------------
    headerRow = 0
    If Not targetCol.Cells(1).ListObject Is Nothing Then
        headerRow = targetCol.Cells(1).ListObject.HeaderRowRange.Row
    ElseIf ws.AutoFilterMode And Not ws.AutoFilter Is Nothing Then
        If Not Intersect(targetCol, ws.AutoFilter.Range) Is Nothing Then
            headerRow = ws.AutoFilter.Range.Row
        End If
    End If

    If headerRow = 0 Then
        On Error Resume Next
        headerRow = ws.Columns(colIdx).Find(What:="*", After:=ws.Cells(ws.Rows.Count, colIdx), _
                    LookIn:=xlValues, SearchDirection:=xlNext).Row
        On Error GoTo 0
    End If

    lastRow = ws.Cells(ws.Rows.Count, colIdx).End(xlUp).Row

    If headerRow = 0 Or lastRow <= headerRow Then
        GetColumnIdentity = "STRING"
        Exit Function
    End If

    Set dataRange = ws.Range(ws.Cells(headerRow + 1, colIdx), ws.Cells(lastRow, colIdx))

' [2단계] 서식(Format) 통일성 검증 및 사전 분석
    ' ---------------------------------------------------------
    Dim vFmtCheck As Variant
    Dim isPercentCol As Boolean: isPercentCol = False ' [최적화 키]

    vFmtCheck = dataRange.NumberFormat

    If IsNull(vFmtCheck) Then
        GetColumnIdentity = "HYBRID"
        Exit Function
    Else
        ' 서식이 통일되어 있다면, 문자열로 변환하여 미리 분석
        fmt = CStr(vFmtCheck)
        If fmt = "@" Then GetColumnIdentity = "STRING": Exit Function

        ' [최적화] 루프 돌기 전에 미리 "이 컬럼은 백분율 후보인가?" 판단
        If InStr(fmt, "%") > 0 Then isPercentCol = True
    End If

    ' [3단계] 값(Value) 기반 정밀 분석 (조건부 최적화)
    ' ---------------------------------------------------------
    Dim hasDecimal As Boolean: hasDecimal = False
    Dim hasPercentFloat As Boolean: hasPercentFloat = False

    If dataRange.Cells.CountLarge > 0 Then
        ' [안전장치] 단일 셀이든 범위든 무조건 2차원 배열로 변환
        If dataRange.Cells.CountLarge = 1 Then
            ReDim dataArray(1 To 1, 1 To 1)
            dataArray(1, 1) = dataRange.Value
        Else
            dataArray = dataRange.Value
        End If

        Dim i As Long
        Dim hasText As Boolean, hasNum As Boolean, hasErr As Boolean
        Dim currentVal As Variant
        Dim dVal As Double, delta As Double, deltaPct As Double

        For i = 1 To UBound(dataArray, 1)
            currentVal = dataArray(i, 1)

            If IsError(currentVal) Then
                hasErr = True: Exit For
            ElseIf Not IsEmpty(currentVal) Then
                Select Case VarType(currentVal)
                    Case vbDouble, vbSingle, vbCurrency
                        hasNum = True

                        ' 1. 일반 소수점(FLOAT) 체크 (항상 수행)
                        If Not hasDecimal Then
                            dVal = CDbl(currentVal)
                            delta = Abs(dVal - Fix(dVal))
                            If delta > EPSILON And delta < (1# - EPSILON) Then
                                hasDecimal = True
                            End If
                        End If

                        ' 2. 백분율 소수점(PERCENT_FLOAT) 체크
                        ' [최적화] 서식이 백분율(%)인 경우에만 수행! (일반 숫자에서는 Skip)
                        If isPercentCol And Not hasPercentFloat Then
                            ' 일반 소수점 체크 때 썼던 dVal 재사용
                            Dim scaledVal As Double
                            scaledVal = dVal * 100
                            deltaPct = Abs(scaledVal - Fix(scaledVal))
                            If deltaPct > EPSILON And deltaPct < (1# - EPSILON) Then
                                hasPercentFloat = True
                            End If
                        End If

                        ' 둘 다 찾았거나(백분율의 경우),
                        ' 백분율이 아닌데 소수점은 찾았다면(일반 숫자의 경우) -> 더 이상 연산 불필요?
                        ' -> 일반 숫자는 hasDecimal만 True면 더 연산할 필요 없음
                        ' -> 백분율은 hasPercentFloat까지 True여야 함

                    Case vbInteger, vbLong, vbByte
                        hasNum = True

                    Case vbDate
                        hasNum = True

                    Case vbString
                        If Len(currentVal) > 0 Then hasText = True
                End Select
            End If

            If hasNum And hasText Then Exit For
        Next i

        If hasErr Then GetColumnIdentity = "ERROR": Exit Function
        If hasText And hasNum Then GetColumnIdentity = "HYBRID": Exit Function
    End If

' [4단계] 최종 상세 판정 (Logic Updated)
    ' --------------------------------------------------------------------------
    ' 표본 셀 확보
    Set sampleCell = dataRange.Find("*", LookIn:=xlValues, SearchDirection:=xlNext)

    If sampleCell Is Nothing Then
        GetColumnIdentity = "STRING"
        Exit Function
    End If

    ' 서식 내 특수문자([]) 제거
    Dim cleanFmt As String: cleanFmt = fmt
    Do While InStr(cleanFmt, "[") > 0 And InStr(cleanFmt, "]") > InStr(cleanFmt, "[")
        ' [Fix v2.7] mID -> Mid 오타 수정
        cleanFmt = Left(cleanFmt, InStr(cleanFmt, "[") - 1) & Mid(cleanFmt, InStr(cleanFmt, "]") + 1)
    Loop

    ' 불리언/날짜 타입 우선 배제
    If VarType(sampleCell.Value) = vbBoolean Then GetColumnIdentity = "STRING": Exit Function
    If VarType(sampleCell.Value) = vbDate Then GetColumnIdentity = "DATE": Exit Function

    ' 숫자형 데이터 처리
    If IsNumeric(sampleCell.Value2) And VarType(sampleCell.Value2) <> vbString Then

        ' 1. 날짜 서식
        If IsDateFmt(cleanFmt) Then
            GetColumnIdentity = "DATE"

        ' 2. 백분율 처리 (PERCENT_FLOAT 삭제 -> PERCENT 통합, ROUNDED만 분리)
        ElseIf isPercentCol Then
            If hasPercentFloat Then
                ' [중요] 값은 소수점이 있는데(실수), 서식에는 점(.)이 없는 경우
                ' 예: 실제값 12.34% -> 서식 "0%" (보이는 값 12%)
                If cleanFmt <> "General" And cleanFmt <> "" And InStr(cleanFmt, ".") = 0 Then
                    GetColumnIdentity = "PERCENT_ROUNDED" ' [주의] 반올림된 백분율
                Else
                    GetColumnIdentity = "PERCENT"         ' 실수/정수 구분 없이 그냥 백분율
                End If
            Else
                GetColumnIdentity = "PERCENT"             ' 정수형 백분율
            End If

        ' 3. 일반 숫자 처리 (FLOAT 삭제 -> NUMBER 통합, ROUNDED만 분리)
        Else
            If hasDecimal Then
                ' [중요] 값은 소수점이 있는데(실수), 서식에는 점(.)이 없는 경우
                ' 예: 실제값 10.5 -> 서식 "0" (보이는 값 11)
                If cleanFmt <> "General" And cleanFmt <> "" And InStr(cleanFmt, ".") = 0 Then
                    GetColumnIdentity = "ROUNDED"         ' [주의] 반올림된 숫자
                Else
                    GetColumnIdentity = "NUMBER"          ' 실수/정수 구분 없이 그냥 숫자
                End If
            Else
                GetColumnIdentity = "NUMBER"              ' 정수형 숫자
            End If
        End If

    Else
        GetColumnIdentity = "STRING"
    End If

    Log "GetColumnIdentity: 최종 판정 -> " & GetColumnIdentity, lvlDebug

End Function

' ------------------------------------------------------------------------------
' 함수명: IsDateFmt
' 목적: 서식 문자열을 분석하여 날짜/시간 관련 서식 코드 포함 여부 판별
' ------------------------------------------------------------------------------
Private Function IsDateFmt(ByVal sFmt As String) As Boolean
    If sFmt = "General" Or sFmt = "" Or sFmt = "@" Then IsDateFmt = False: Exit Function

    Dim sClean As String, i As Long, bInQuote As Boolean, c As String

    ' 따옴표로 감싸진 리터럴 텍스트 제거 (서식 코드 오인식 방지)
    sClean = ""
    For i = 1 To Len(sFmt)
        ' [Fix v2.7] mID -> Mid 오타 수정
        c = Mid(sFmt, i, 1)
        If c = """" Then bInQuote = Not bInQuote Else If Not bInQuote Then sClean = sClean & c
    Next i

    ' 표준 날짜/시간 기호(y, m, d, h, s) 포함 여부 체크
    sClean = LCase(sClean)
    If InStr(1, sClean, "y") > 0 Or InStr(1, sClean, "d") > 0 Or _
       InStr(1, sClean, "h") > 0 Or InStr(1, sClean, "s") > 0 Or _
       InStr(1, sClean, "m") > 0 Then
        IsDateFmt = True
    Else
        IsDateFmt = False
    End If
End Function
