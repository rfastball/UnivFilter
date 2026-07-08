Attribute VB_Name = "modNormalize"
' @desc 입력 값을 타입별 정규화 값과 배열로 변환 담당
' @depends modIndexKey, modLogger, modColumnType
Option Explicit

' ==============================================================================
' [L1 순수 값 정규화] modFilterUtils에서 이동 (Stage 5)
' 검증된 로직의 이동이므로 알고리즘 수정 금지 - 동작 계약은 Test_Normalize가 고정
' ==============================================================================

' ------------------------------------------------------------------------------
' 함수명: GetUniqueCleanData
' 목적: 입력 데이터를 eType에 맞게 정제하고 중복을 제거하여 Dictionary로 반환
' ------------------------------------------------------------------------------
' @desc 입력과 타입 받아 정제 키/원본 값 Dictionary 반환, 실패 값은 False 플래그
Public Function GetUniqueCleanData(ByVal vInput As Variant, ByVal eType As FilterDataType) As Object
    Dim dict As Object
    Dim vArr As Variant
    Dim i As Long
    Dim vRaw As Variant, vClean As Variant
    Dim bIsValid As Boolean

    ' [LOG] 정제 시작
    Log "GetUniqueCleanData: 데이터 정제 시작 (TargetType: " & modColumnType.GetTypeName(eType) & ")", lvlDebug

    Set dict = CreateObject("Scripting.Dictionary")
    vArr = ConvertToFlatArray(vInput)

    For i = LBound(vArr) To UBound(vArr)
        vRaw = vArr(i)

        If Len(Trim(CStr(vRaw))) > 0 Then
            ' 1. 변환 시도
            vClean = NormalizeValue(vRaw, eType)

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
' [데이터 변환 라우터] (Converter & Workers)
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
' [데이터 정규화 알고리즘] (Sanitization)
' ==============================================================================
' ==========================================================
' NormalizeValue: eType에 따라 '적절히' 정규화된 값 추출 (구 CleanValue)
' ==========================================================
' @desc 원시 값과 타입 받아 정규화 값 반환, 빈 입력 시 Empty 반환
Public Function NormalizeValue(ByVal vRaw As Variant, ByVal eType As FilterDataType) As Variant
    Dim sTemp As String

    ' 1. [기본 청소]
    sTemp = CStr(vRaw)
    sTemp = Replace(sTemp, ChrW(160), "")  ' Non-breaking space
    sTemp = Replace(sTemp, ChrW(8361), "") ' Won Sign
    sTemp = Replace(sTemp, "$", "")        ' $ (Dollar)
    sTemp = Trim(sTemp)

    If Len(sTemp) = 0 Then NormalizeValue = Empty: Exit Function

    Select Case eType
        Case fdDate ' [날짜 타입]
            NormalizeValue = sTemp '날짜도 문자열로 간주

        ' [통합] 일반 숫자 & 숨겨진 숫자(Rounded)
        ' 목표: 화면에 11로 보여도 실제값 10.5를 추출해야 함
        Case fdNumber, fdRounded
            ' 1. 회계 서식 (-) 처리
            If sTemp = "-" Or sTemp = "-\" Or sTemp = "\-" Then
                NormalizeValue = 0
                Exit Function
            End If

            ' 2. 천단위 콤마 제거
            Dim sNumClean As String
            sNumClean = Replace(Replace(sTemp, " ", ""), ",", "")

            ' 3. 변환
            If IsNumeric(sNumClean) Then
                NormalizeValue = CDbl(sNumClean)
            Else
                ' [강력 추출] 숫자와 소수점(.) 보존
                Dim vAggressive As Variant
                vAggressive = ExtractNumberAggressive(sTemp)

                If IsNumeric(vAggressive) Then
                    NormalizeValue = CDbl(vAggressive)
                Else
                    NormalizeValue = vRaw
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
                    NormalizeValue = CDbl(sPctNum) / 100#
                Else
                    NormalizeValue = CDbl(sPctNum)
                End If
            Else
                NormalizeValue = vRaw
            End If

        Case fdString ' [문자열 타입]
            NormalizeValue = sTemp

        Case Else ' [자동 감지]
            ' 기존 로직 유지 (숫자 우선, 실패시 문자)
            If IsNumeric(sTemp) Then
                NormalizeValue = CDbl(sTemp)
            ElseIf IsDate(sTemp) Then
                NormalizeValue = CDate(sTemp)
            Else
                Dim sAutoCheck As String
                sAutoCheck = Replace(Replace(sTemp, " ", ""), ",", "")
                If IsNumeric(sAutoCheck) And sAutoCheck <> "" Then
                    NormalizeValue = CDbl(sAutoCheck)
                Else
                    NormalizeValue = sTemp
                End If
            End If
    End Select
End Function

' ------------------------------------------------------------------------------
' 함수명: ExtractNumberAggressive (RegExp 활용)
' 목적: 오염된 문자열에서 숫자(정수/실수/음수)만 강력하게 추출
' 개선: 마이너스(-) 기호가 문자열 중간에 끼어 있어도(예: "Val:-12.5") 인식하도록 수정
' ------------------------------------------------------------------------------
' @desc 오염 문자열 받아 추출 숫자 반환, 숫자 없으면 원본 반환
Public Function ExtractNumberAggressive(ByVal sInput As String) As Variant
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
        ' 숫자가 아예 없으면 원본 반환 (NormalizeValue에서 처리하도록)
        ExtractNumberAggressive = sInput
    End If
End Function

' ==============================================================================
' [안전 문자열화] modIndexKey 키 계약의 표시용 확장 (구 Core_modNorm 아이디어 흡수)
' ==============================================================================
' @desc Variant 값 받아 안전한 표시 문자열 반환, Error는 #ERROR / Null은 빈 문자열
Public Function SafeText(ByVal vValue As Variant) As String
    If IsError(vValue) Then
        SafeText = "#ERROR"
    ElseIf IsNull(vValue) Then
        SafeText = vbNullString
    Else
        ' Empty -> "" / 그 외 CStr은 벤더링된 키 계약(modIndexKey)에 위임
        SafeText = modIndexKey.NormalizeIndexKey(vValue)
    End If
End Function
