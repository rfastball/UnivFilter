Attribute VB_Name = "modFilterUtils"
Option Explicit

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
Private Function GetTypeName(eType As FilterDataType) As String
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

' ==============================================================================
' [섹션 1] 메인 데이터 정제 엔진
' ==============================================================================
' ------------------------------------------------------------------------------
' 함수명: GetUniqueCleanData
' 목적: 입력 데이터를 eType에 맞게 정제하고 중복을 제거하여 Dictionary로 반환
' ------------------------------------------------------------------------------
Public Function GetUniqueCleanData(ByVal vInput As Variant, ByVal eType As FilterDataType) As Object
    Dim dict As Object
    Dim vArr As Variant
    Dim i As Long
    Dim vRaw As Variant, vClean As Variant
    Dim bIsValid As Boolean
    
    ' [LOG] 정제 시작
    Log "GetUniqueCleanData: 데이터 정제 시작 (TargetType: " & GetTypeName(eType) & ")", lvlDebug
    
    Set dict = CreateObject("Scripting.Dictionary")
    vArr = ConvertToFlatArray(vInput)
    
    For i = LBound(vArr) To UBound(vArr)
        vRaw = vArr(i)
        
        If Len(Trim(CStr(vRaw))) > 0 Then
            ' 1. 변환 시도 (CleanValue 함수도 업데이트 필요)
            vClean = CleanValue(vRaw, eType)
            
            ' 2. 유효성 검증 (확장된 Enum 반영)
            bIsValid = True
            
            Select Case eType
                ' [숫자 계열] 정수, 실수, 백분율 모두 숫자인지 확인
                Case fdNumber, fdPercent, fdRounded, fdPercentRounded
                    If Not IsNumeric(vClean) Then bIsValid = False
                    
                ' [날짜 계열]
                Case fdDate
                    If Not IsDate(vClean) Then bIsValid = False
                    
                ' [텍스트/기타]
                ' fdString, fdHybrid, fdAuto 등은 기본적으로 True 유지
                Case Else
                    ' 필요시 추가 검증 로직 (예: 길이 제한 등)
            End Select
            
            ' 3. 이중 바스켓 적재 (Double Basket Strategy)
            If bIsValid Then
                ' [성공 바스켓] Key: 정제된 값 / Value: 원본 (시각적 검사용)
                If Not dict.exists(vClean) Then dict(vClean) = vRaw
            Else
                ' [실패 바스켓] Key: 원본 값 / Value: False (플래그)
                If Not dict.exists(vRaw) Then dict(vRaw) = False
                Log "GetUniqueCleanData: 변환 실패 -> 실패 바스켓 적재 (" & vRaw & ")", lvlDebug
            End If
        End If
    Next i
    
    Set GetUniqueCleanData = dict
End Function



' ==============================================================================
' [섹션 2] 데이터 변환 라우터 (Converter & Workers)
' ==============================================================================

' ------------------------------------------------------------------------------
' 함수명: ConvertToFlatArray (Router)
' 목적: 입력된 Variant 데이터의 타입에 맞춰 최적의 배열 변환 워커로 라우팅
' ------------------------------------------------------------------------------
Private Function ConvertToFlatArray(ByVal vInput As Variant) As Variant
    If IsObject(vInput) Then
        ' Excel Range 객체인 경우
        If TypeOf vInput Is Range Then
            Log "ConvertToFlatArray: Range Source Detected", lvlDebug
            ConvertToFlatArray = RangeToArray(vInput)
            Exit Function
        End If
    ElseIf IsArray(vInput) Then
        ' 이미 배열 형태인 경우 그대로 반환
        Log "ConvertToFlatArray: Array Source Detected", lvlDebug
        ConvertToFlatArray = vInput
        Exit Function
    ElseIf VarType(vInput) = vbString Then
        ' 텍스트 박스 등에서 들어온 멀티라인 문자열인 경우
        Log "ConvertToFlatArray: String Source Detected", lvlDebug
        ConvertToFlatArray = StringToArray(CStr(vInput))
        Exit Function
    End If
    
    ' 그 외 단일 값(숫자, 단일 문자열 등)은 1차원 배열로 래핑
    Log "ConvertToFlatArray: Single Value Source Detected", lvlDebug
    ConvertToFlatArray = Array(vInput)
End Function

' [Worker 1] 멀티라인 문자열 -> 배열 변환 (줄바꿈 기준)
' ------------------------------------------------------------------------------
Private Function StringToArray(ByVal strRaw As String) As Variant
    Dim tempVal As String
    Dim arrResult() As String
    Dim vFinal() As Variant
    Dim i As Long, cnt As Long
    
    If Trim(strRaw) = "" Then
        StringToArray = Array(): Exit Function
    End If
    
    ' 줄바꿈 기호(CRLF, CR)를 LF로 통일 후 분리
    tempVal = Replace(Replace(strRaw, vbCrLf, vbLf), vbCr, vbLf)
    arrResult = Split(tempVal, vbLf)
    
    ' 유효 데이터 추출 및 공백 제거
    ReDim vFinal(0 To UBound(arrResult))
    cnt = 0
    For i = LBound(arrResult) To UBound(arrResult)
        tempVal = Trim(arrResult(i))
        If Len(tempVal) > 0 Then
            vFinal(cnt) = tempVal
            cnt = cnt + 1
        End If
    Next i
    
    ' 결과 배열 리사이징
    If cnt > 0 Then
        ReDim Preserve vFinal(0 To cnt - 1)
        StringToArray = vFinal
    Else
        StringToArray = Array()
    End If
End Function

' [Worker 2] 엑셀 범위(Range) -> 배열 변환 (고속 메모리 로드)
' ------------------------------------------------------------------------------
Private Function RangeToArray(ByVal rngTarget As Range) As Variant
    Dim vData As Variant
    Dim vResult() As Variant
    Dim i As Long, j As Long, cnt As Long
    
    If rngTarget Is Nothing Then RangeToArray = Array(): Exit Function
    
    ' 단일 셀 처리 최적화
    If rngTarget.Cells.CountLarge = 1 Then
        ReDim vResult(0 To 0)
        vResult(0) = rngTarget.Value2
        RangeToArray = vResult
        Exit Function
    End If
    
    ' 대량 데이터 처리를 위해 범위를 메모리 배열로 일시 적재
    vData = rngTarget.Value2
    ReDim vResult(0 To rngTarget.Cells.CountLarge - 1)
    
    cnt = 0
    For i = LBound(vData, 1) To UBound(vData, 1)
        For j = LBound(vData, 2) To UBound(vData, 2)
            If Not IsEmpty(vData(i, j)) Then
                vResult(cnt) = vData(i, j)
                cnt = cnt + 1
            End If
        Next j
    Next i
    
    ' 유효 데이터 수에 맞춰 배열 크기 확정
    If cnt > 0 Then
        ReDim Preserve vResult(0 To cnt - 1)
        RangeToArray = vResult
    Else
        RangeToArray = Array()
    End If
End Function

' ==============================================================================
' [섹션 3] 데이터 정규화 알고리즘 (Sanitization)
' ==============================================================================
' ==========================================================
' CleanValue: eType에 따라 '적절히' 정규화된 값 추출
' ==========================================================
Private Function CleanValue(ByVal vRaw As Variant, ByVal eType As FilterDataType) As Variant
    Dim sTemp As String
    
    ' 1. [기본 청소]
    sTemp = CStr(vRaw)
    sTemp = Replace(sTemp, ChrW(160), "")  ' Non-breaking space
    sTemp = Replace(sTemp, ChrW(8361), "") ' \ (Won Sign)
    sTemp = Replace(sTemp, "$", "")        ' $ (Dollar)
    sTemp = Trim(sTemp)
    
    If Len(sTemp) = 0 Then CleanValue = Empty: Exit Function

    Select Case eType
        Case fdDate ' [날짜 타입]
            CleanValue = sTemp '날짜도 문자열로 간주
            
        ' [통합] 일반 숫자 & 숨겨진 숫자(Rounded)
        ' 목표: 화면에 11로 보여도 실제값 10.5를 추출해야 함
        Case fdNumber, fdRounded
            ' 1. 회계 서식 (-) 처리
            If sTemp = "-" Or sTemp = "-\" Or sTemp = "\-" Then
                CleanValue = 0
                Exit Function
            End If
            
            ' 2. 천단위 콤마 제거
            Dim sNumClean As String
            sNumClean = Replace(Replace(sTemp, " ", ""), ",", "")
            
            ' 3. 변환
            If IsNumeric(sNumClean) Then
                CleanValue = CDbl(sNumClean)
            Else
                ' [강력 추출] 숫자와 소수점(.) 보존
                Dim vAggressive As Variant
                vAggressive = ExtractNumberAggressive(sTemp) ' 기존 함수 재사용
                
                If IsNumeric(vAggressive) Then
                    CleanValue = CDbl(vAggressive)
                Else
                    CleanValue = vRaw
                End If
            End If
            
        ' [통합] 백분율 & 숨겨진 백분율
        Case fdPercent, fdPercentRounded
            Dim sPctVal As String
            sPctVal = Replace(sTemp, " ", "")
            
            Dim sPctNum As String
            sPctNum = Replace(sPctVal, "%", "")
            
            If IsNumeric(sPctNum) Then
                ' % 기호 유무에 따른 스케일링
                If InStr(sTemp, "%") > 0 Then
                    CleanValue = CDbl(sPctNum) / 100#
                Else
                    CleanValue = CDbl(sPctNum)
                End If
            Else
                CleanValue = vRaw
            End If
            
        Case fdString ' [문자열 타입]
            CleanValue = sTemp
            
        Case Else ' [자동 감지]
            ' 기존 로직 유지 (숫자 우선, 실패시 문자)
            If IsNumeric(sTemp) Then
                CleanValue = CDbl(sTemp)
            ElseIf IsDate(sTemp) Then
                CleanValue = CDate(sTemp)
            Else
                Dim sAutoCheck As String
                sAutoCheck = Replace(Replace(sTemp, " ", ""), ",", "")
                If IsNumeric(sAutoCheck) And sAutoCheck <> "" Then
                    CleanValue = CDbl(sAutoCheck)
                Else
                    CleanValue = sTemp
                End If
            End If
    End Select
End Function

' ------------------------------------------------------------------------------
' 함수명: ExtractNumberAggressive (RegExp 활용)
' 목적: 오염된 문자열에서 숫자(정수/실수/음수)만 강력하게 추출
' 개선: 마이너스(-) 기호가 문자열 중간에 끼어 있어도(예: "Val:-12.5") 인식하도록 수정
' ------------------------------------------------------------------------------
Private Function ExtractNumberAggressive(ByVal sInput As String) As Variant
    Static reg As Object
    Dim matches As Object
    
    ' [강력 청소] 콤마, 특수공백(160), 원화(8361), 달러($), 일반공백 제거
    ' 주의: 마이너스(-)는 제거하면 안 됨!
    Dim sClean As String
    sClean = sInput
    sClean = Replace(sClean, ",", "")
    sClean = Replace(sClean, ChrW(160), "")
    sClean = Replace(sClean, ChrW(8361), "")
    sClean = Replace(sClean, "$", "")
    sClean = Replace(sClean, " ", "")
    sClean = Trim(sClean)
    
    If reg Is Nothing Then
        Set reg = CreateObject("VBScript.RegExp")
        reg.Global = True
        reg.IgnoreCase = True
        ' [패턴 설명]
        ' [-]?       : 마이너스 기호가 0개 또는 1개 있음
        ' \d+        : 숫자 1개 이상
        ' (\.\d+)?   : (소수점과 숫자들)이 0개 또는 1개 있음
        reg.Pattern = "[-]?\d+(\.\d+)?"
    End If
    
    ' 이제 bIsNegative 변수로 앞머리만 검사할 필요가 없음
    If reg.Test(sClean) Then
        Set matches = reg.Execute(sClean)
        ' 정규식이 "-12.5"를 통째로 잡았으므로 바로 CDbl 변환
        ExtractNumberAggressive = CDbl(matches(0).Value)
    Else
        ' 숫자가 아예 없으면 원본 반환 (CleanValue에서 처리하도록)
        ExtractNumberAggressive = sInput
    End If
End Function

' ==============================================================================
' [섹션 4] 메타데이터 분석 및 아이덴티티 감지
' ==============================================================================
' ------------------------------------------------------------------------------
' 함수명: GetRepresentativeFormat
' 목적: 범위 내에서 데이터가 존재하는 첫 번째 셀을 찾아 표시 형식을 추출
' 포맷이 단일하다고 확인된 열에대해 구체적 포맷값을 얻을 때
' ------------------------------------------------------------------------------
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
        cleanFmt = Left(cleanFmt, InStr(cleanFmt, "[") - 1) & mID(cleanFmt, InStr(cleanFmt, "]") + 1)
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
        c = mID(sFmt, i, 1)
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

' ==============================================================================
' 함수명: GetSafeFormattedText
' 설명: 값(vVal)에 서식(sFmt)을 적용하되, 에러 발생 시 단계적으로 덜 정교한 메서드를 시도함
' 우선순위:
'   1. WorksheetFunction.Text (엑셀 고유 서식 완벽 지원)
'   2. VBA Format 함수 (표준 서식 지원)
'   3. 기본 날짜 포맷 (yyyy-mm-dd)
'   4. 단순 문자열 변환 (CStr)
' ==============================================================================
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
Public Function GetAccountingFilterString(ByVal vValue As Variant, ByVal sFormat As String) As String
    Dim sSymbol As String: sSymbol = ""
    Dim sCleanNum As String: sCleanNum = modFilterUtils.ExtractNumberAggressive(CStr(vValue))
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

Public Function GetPercentageFilterString(ByVal vValue As Variant, ByVal sFormat As String) As String
    Dim sCleanNum As String: sCleanNum = modFilterUtils.ExtractNumberAggressive(CStr(vValue))
    
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


