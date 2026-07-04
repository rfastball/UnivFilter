Attribute VB_Name = "modMatchingEngine"
Option Explicit

' ==============================================================================
' [섹션 1] 메인 컨트롤러 (CalculateCriteria)
' 목적: 열의 특성을 분석하여 최적의 매칭 전략(Strategy)을 결정하고 실행합니다.
' ==============================================================================
Public Function CalculateCriteria( _
    ByVal vInput As Variant, _
    ByVal colAsset As clsColumnAsset, _
    ByVal bExact As Boolean, _
    ByVal bExclude As Boolean) As Object
    
    Dim dictResult As Object
    Dim dictInput As Object
    Dim sColIdentity As String
    Dim eType As FilterDataType, eTypeForInput As FilterDataType
    
    ' [LOG] 매칭 프로세스 시작
    Log "CalculateCriteria: 매칭 분석 시작 (Exact=" & bExact & ", Exclude=" & bExclude & ")", lvlInfo
    
    Set dictResult = CreateObject("Scripting.Dictionary")
    
    ' 1. [타입 분석] 데이터 자산(Asset)으로부터 열의 정체성 확인
    sColIdentity = colAsset.ColumnType
    eType = MapIdentityToEnum(sColIdentity)
    
    Log "CalculateCriteria: 열 정체성 확인 (ID: " & sColIdentity & " -> Enum: " & eType & ")", lvlDebug
    
    ' --------------------------------------------------------------------------
    ' 데이터 정합성 오류 시 -> 텍스트 모드로 강제 전환 HYBRID, ERROR, DATE
    ' --------------------------------------------------------------------------
    If colAsset.ColumnType = "HYBRID" Or colAsset.ColumnType = "ERROR" Or colAsset.ColumnType = "DATE" Then
        Log "CalculateCriteria: [WARN] Hybrid/Error 컬럼 감지 -> String 전략으로 강제 우회", lvlWarning
        eType = fdString ' 강제로 텍스트 타입으로 간주
    End If
    ' --------------------------------------------------------------------------
    
    ' 2. [입력 데이터 정제]
    ' 부분 일치(Exact=False) 시에는 입력값을 문자열로 취급하여 유연한 검색을 유도합니다.
    ' 정확 일치(Exact=True) 시에는 열 타입에 맞춰 파싱하여 유니버스 값과 정밀 매칭을 수행합니다.
    If bExact Then
        eTypeForInput = eType
    Else
        eTypeForInput = fdString
    End If
    
    Set dictInput = modFilterUtils.GetUniqueCleanData(vInput, eTypeForInput)
    Log "CalculateCriteria: 입력값 정제 완료 (Count: " & dictInput.Count & ")", lvlDebug
    
    ' 3. [전략 배분] 타입에 따른 전략 분기 (Dispatch)
    If eType = fdString Or eType = fdAuto Then
        ' 텍스트 전용 전략 실행 (String Strategy)
        Log "CalculateCriteria: String Strategy 실행", lvlInfo
        Set dictResult = ExecuteStringStrategy(dictInput, colAsset, bExact, bExclude)
    Else
        ' 숫자/날짜 전용 전략 실행 (Value Strategy)
        Log "CalculateCriteria: Value Strategy 실행", lvlInfo
        Set dictResult = ExecuteValueStrategy(dictInput, colAsset, bExact, bExclude, eType)
    End If
    
    ' 4. [결과 패키징] 메타 정보 기록
    dictResult("ColumnType") = sColIdentity
    dictResult("InputCount") = dictInput.Count
    
    Log "CalculateCriteria: 최종 결과 패키징 완료", lvlDebug
    
    Set CalculateCriteria = dictResult
End Function

' ==============================================================================
' [섹션 2] 값 기반 통합 전략 (Value Strategy)
' 역할: 숫자/날짜 타입의 모든 매칭(정확/부분) 및 포맷 보정 처리
' ==============================================================================
' ==========================================================
' [MODULE] modMatchingEngine
' [수정] ExecuteValueStrategy: 입국 심사(Pre-Filtering) 및 합집합 로직 추가
' ==========================================================
' ==============================================================================
' [섹션 2] 값 기반 통합 전략 (Value Strategy)
' 역할: 숫자/날짜 타입의 모든 매칭(정확/부분) 및 포맷 보정 처리
' 수정사항: FLOAT 예외 로직 제거 -> 모든 숫자(ROUNDED 포함)가 서식 기반 매칭(GetSafeFormattedText)으로 통일됨
' ==============================================================================
Private Function ExecuteValueStrategy( _
    ByVal dictInput As Object, _
    ByVal colAsset As clsColumnAsset, _
    ByVal bExact As Boolean, _
    ByVal bExclude As Boolean, _
    ByVal eType As FilterDataType) As Object
    
    Dim dictRes As Object
    Set dictRes = CreateObject("Scripting.Dictionary")
    
    ' [Check] 서식 정보 확인
    Dim bIsAccounting As Boolean
    If InStr(colAsset.fmt, "*") > 0 Or InStr(colAsset.fmt, "_") > 0 Then bIsAccounting = True
    
    ' [분기 1] 정확 일치 (Exact Match)
    If bExact Then
        ' ----------------------------------------------------------------------
        ' 1. 입국 심사 (Pre-Filtering)
        ' ----------------------------------------------------------------------
        Dim vKey As Variant
        Dim arrCleanInput() As Variant  ' clsSet에 태울 순수 숫자들
        Dim arrTextFail() As Variant    ' 실패한 텍스트들 (즉시 Missing 처리)
        Dim iClean As Long, iFail As Long
        
        ReDim arrCleanInput(0 To dictInput.Count)
        ReDim arrTextFail(0 To dictInput.Count)
        iClean = 0: iFail = 0
        
        For Each vKey In dictInput.Keys
            ' Value가 False면 실패 바스켓(텍스트) -> 열외
            If VarType(dictInput(vKey)) = vbBoolean Then
                If dictInput(vKey) = False Then
                    arrTextFail(iFail) = vKey
                    iFail = iFail + 1
                    GoTo NextItem
                End If
            End If
            
            ' Value가 정상이면 성공 바스켓(숫자) -> 탑승
            arrCleanInput(iClean) = vKey
            iClean = iClean + 1
NextItem:
        Next vKey
        
        ' 배열 크기 최적화 (Trim)
        If iClean > 0 Then ReDim Preserve arrCleanInput(0 To iClean - 1) Else arrCleanInput = Array()
        If iFail > 0 Then ReDim Preserve arrTextFail(0 To iFail - 1) Else arrTextFail = Array()
        
        Log "ExecuteValueStrategy: 입국 심사 완료 (Pass: " & iClean & ", Fail: " & iFail & ")", lvlDebug
        
        ' ----------------------------------------------------------------------
        ' 2. Core_clsSet 집합 연산 (순수 숫자만 처리)
        ' ----------------------------------------------------------------------
        Dim setUniverse As Core_clsSet, setInputClean As Core_clsSet
        Dim vResult As Variant
        Dim setCommon As Core_clsSet, setDiffInput As Core_clsSet
        
        Set setUniverse = colAsset.data
        Set setInputClean = New Core_clsSet
        
        ' 정제된 숫자만 로드
        If iClean > 0 Then setInputClean.LoadFromArray arrCleanInput
        
        ' 비교 연산 (Common, DiffUniverse, DiffInput)
        vResult = setUniverse.Compare(setInputClean, 4)
        Set setCommon = vResult(0)
        ' vResult(1)은 유니버스 차집합(미사용)
        Set setDiffInput = vResult(2) ' 숫자인데 없는 것들
        
        ' ----------------------------------------------------------------------
        ' 3. 결과 병합 (Numeric Missing + Text Fail -> Union)
        ' ----------------------------------------------------------------------
        ' 아까 빼둔 텍스트 실패분(arrTextFail)을 미매칭 목록에 합류
        If iFail > 0 Then
            setDiffInput.LoadFromArray arrTextFail
        End If
        
        ' ----------------------------------------------------------------------
        ' 4. VisFail 검사 (숫자 매칭 성공분에 한해서만 수행)
        ' ROUNDED/PERCENT_ROUNDED 포함 모든 숫자가 서식 검사 수행
        ' ----------------------------------------------------------------------
        Dim dictCriteria As Object, dictVisFail As Object
        Dim sFmt As String, sReconst As String, vRaw As Variant
        
        Set dictCriteria = CreateObject("Scripting.Dictionary")
        Set dictVisFail = CreateObject("Scripting.Dictionary")
        sFmt = modFilterUtils.GetCleanFormat(colAsset.fmt)
        
        For Each vKey In setCommon.KeysArray
            ' 딕셔너리에 있는 원본값(vRaw)을 가져옴 (성공 바스켓은 Value가 원본임)
            If dictInput.exists(vKey) Then
                vRaw = dictInput(vKey)
                
                ' [Logic Change]
                ' eType = fdFloat 체크 로직 삭제됨.
                ' ROUNDED, PERCENT_ROUNDED를 포함한 모든 Number Strategy는
                ' 화면에 보이는 값(SafeFormattedText)을 기준으로 역산(Reconst)합니다.
                sReconst = modFilterUtils.GetSafeFormattedText(vKey, sFmt)
                
                dictCriteria(sReconst) = Empty
                
                ' 원본 입력값(vRaw)과 화면 표시값(sReconst)이 다르면 VisFail 등록
                If StrComp(Trim(CStr(vRaw)), sReconst, vbTextCompare) <> 0 Then
                    dictVisFail(vRaw) = sReconst
                End If
            End If
        Next vKey
        
        ' 5. 필터 기준(Criteria) 생성
        Dim arrFinal As Variant
        Dim targetSet As Core_clsSet
        If bExclude Then
             ' 제외 모드: 전체 - 교집합
             Set targetSet = setUniverse.Difference(setCommon)
        Else
             ' 포함 모드: 교집합
             Set targetSet = setCommon
        End If
        
        If targetSet.Count > 0 Then
            If bIsAccounting Then
                arrFinal = FormatSetToCriteria(targetSet, sFmt)
            Else
                If bExclude Then
                    arrFinal = FormatSetToCriteria(targetSet, sFmt)
                Else
                    arrFinal = dictCriteria.Keys
                End If
            End If
        Else
            arrFinal = Array("###NO_MATCH###")
        End If
        
        dictRes("Criteria") = arrFinal
        dictRes("Missing") = setDiffInput.KeysArray ' 합쳐진 Missing 반환
        Set dictRes("VisFail") = dictVisFail
        
    ' [분기 2] 부분 일치 (Partial Match) - 기존 로직 유지
    Else
        ' 부분 일치는 텍스트 변환 기반이므로, 입력값 전체(dictInput.Keys)를 그대로 사용해도 무방
        Dim setMatched As Core_clsSet
        Dim arrCriteriaPart As Variant
        Dim targetSetPart As Core_clsSet
        
        Set setMatched = GetPartialMatch_Smart(colAsset.data, dictInput.Keys, colAsset.fmt)
        
        If bExclude Then
             Set targetSetPart = colAsset.data.Difference(setMatched)
        Else
             Set targetSetPart = setMatched
        End If
        
        If targetSetPart.Count > 0 Then
            If bIsAccounting Then
                arrCriteriaPart = targetSetPart.KeysArray
            Else
                arrCriteriaPart = FormatSetToCriteria(targetSetPart, colAsset.fmt)
            End If
        Else
            arrCriteriaPart = Array("###NO_MATCH###")
        End If
        
        dictRes("Criteria") = arrCriteriaPart
        dictRes("Missing") = Array() ' 부분 일치는 Missing 리포트 없음
        Set dictRes("VisFail") = CreateObject("Scripting.Dictionary")
    End If
            
    Set ExecuteValueStrategy = dictRes
End Function

' ==============================================================================
' [섹션 3] 문자열 처리 전략 (String Strategy)
' 역할: 순수 텍스트 컬럼 전용 (고속 텍스트 검색 최적화)
' ==============================================================================
Private Function ExecuteStringStrategy( _
    ByVal dictInput As Object, _
    ByVal colAsset As clsColumnAsset, _
    ByVal bExact As Boolean, _
    ByVal bExclude As Boolean) As Object
    
    Dim dictRes As Object
    Dim setUniverse As Core_clsSet, setKeywords As Core_clsSet
    Dim arrCriteria As Variant, vMissing As Variant
    
    Set dictRes = CreateObject("Scripting.Dictionary")
    Set setUniverse = colAsset.data
    
    ' [Patch 1] 유니버스 데이터가 Boolean 타입인지 감지
    Dim isBoolType As Boolean
    If setUniverse.Count > 0 Then
        If VarType(setUniverse.KeysArray(0)) = vbBoolean Then isBoolType = True
    End If
    
    If bExact Then
        ' 텍스트 정확 일치: 집합 연산(Compare)을 통해 처리
        Set setKeywords = New Core_clsSet
        
        ' [Patch 2] Boolean 열인데 입력값이 텍스트면, Boolean으로 변환해서 엔진에 주입
        If isBoolType Then
            Log "ExecuteStringStrategy: Boolean 타입 열 감지 (변환 로직 가동)", lvlDebug
            Dim k As Variant, sKey As String
            For Each k In dictInput.Keys
                sKey = LCase(Trim(CStr(k))) ' 입력값을 소문자로 통일
                If sKey = "true" Or sKey = "t" Or sKey = "1" Then
                    setKeywords.Add True
                ElseIf sKey = "false" Or sKey = "f" Or sKey = "0" Then
                    setKeywords.Add False
                Else
                    ' Boolean이 아닌 잡다한 텍스트는 그냥 둠 (어차피 매칭 안됨)
                    setKeywords.Add k
                End If
            Next k
        Else
            ' 일반 텍스트 열은 그대로 로드
            setKeywords.LoadFromArray dictInput.Keys
        End If

        Dim vResult As Variant
        vResult = setUniverse.Compare(setKeywords, 4)
        
        If bExclude Then arrCriteria = vResult(1).KeysArray Else arrCriteria = vResult(0).KeysArray
        vMissing = vResult(2).KeysArray
        
        Log "ExecuteStringStrategy: [Exact] 매칭 완료 (Criteria: " & (UBound(arrCriteria) - LBound(arrCriteria) + 1) & ", Missing: " & (UBound(vMissing) - LBound(vMissing) + 1) & ")", lvlDebug
        
    Else
        ' 텍스트 부분 일치: 데이터 성격에 따라 경로 결정
        Dim setMatched As Core_clsSet
        
        ' [수정] 날짜(DATE) 타입이거나 특별한 표시 형식이 지정된 경우 Smart Path로 우회
        ' Smart Path는 Raw Value와 Formatted Text를 모두 검사합니다.
        If colAsset.ColumnType = "DATE" Or (colAsset.fmt <> "General" And colAsset.fmt <> "@" And colAsset.fmt <> "") Then
            Log "ExecuteStringStrategy: [Partial] Smart Path 실행 (날짜/서식 대응)", lvlDebug
            ' Smart Path는 유니버스(setUniverse), 키워드 배열, 서식 문자열을 인자로 받습니다.
            Set setMatched = GetPartialMatch_Smart(setUniverse, dictInput.Keys, colAsset.fmt)
        Else
            ' 순수 텍스트 컬럼인 경우 성능을 위해 기존 Fast Path 유지
            Log "ExecuteStringStrategy: [Partial] Fast Path 실행 (순수 텍스트)", lvlDebug
            Set setMatched = GetPartialMatch_Fast(setUniverse, dictInput.Keys)
        End If
        
        If bExclude Then
            arrCriteria = setUniverse.Difference(setMatched).KeysArray
        Else
            arrCriteria = setMatched.KeysArray
        End If
        vMissing = Array()
        
        Log "ExecuteStringStrategy: [Partial] 매칭 완료 (" & (UBound(arrCriteria) - LBound(arrCriteria) + 1) & "건)", lvlDebug
    End If
    
    If IsEmpty(arrCriteria) Then arrCriteria = Array("###NO_MATCH###")
    
    ' [Patch 3] Boolean 결과값을 엑셀 필터가 좋아하는 "TRUE"/"FALSE" 문자열로 최종 변환 (Output Formatting)
    If isBoolType And Not IsEmpty(arrCriteria) Then
        If arrCriteria(0) <> "###NO_MATCH###" Then
            Log "ExecuteStringStrategy: Boolean 결과 문자열 변환 수행", lvlDebug
            Dim vFinal() As String
            Dim i As Long
            ReDim vFinal(LBound(arrCriteria) To UBound(arrCriteria))
            
            For i = LBound(arrCriteria) To UBound(arrCriteria)
                If VarType(arrCriteria(i)) = vbBoolean Then
                    ' 대문자 문자열로 강제 변환 (엑셀 AutoFilter 표준)
                    If arrCriteria(i) Then vFinal(i) = "TRUE" Else vFinal(i) = "FALSE"
                Else
                    vFinal(i) = CStr(arrCriteria(i))
                End If
            Next i
            arrCriteria = vFinal
        End If
    End If

    dictRes("Criteria") = arrCriteria
    dictRes("Missing") = vMissing
    Set dictRes("VisFail") = CreateObject("Scripting.Dictionary")
    
    Set ExecuteStringStrategy = dictRes
End Function

' ==============================================================================
' [섹션 4] 내부 조력 함수 (Helpers)
' ==============================================================================

' [Helper 1] 스마트 부분 검색 (본캐/부캐 이중 검사)
' ------------------------------------------------------------------------------
Private Function GetPartialMatch_Smart( _
    ByVal setUniverse As Core_clsSet, _
    ByVal arrKeywords As Variant, _
    ByVal sFmt As String) As Core_clsSet
    
    Dim newSet As New Core_clsSet
    Dim vRawData As Variant
    Dim sDataRaw As String, sDataFmt As String, sKw As String
    Dim i As Long, k As Long
    Dim isMatch As Boolean
    
    vRawData = setUniverse.KeysArray
    If UBound(vRawData) < LBound(vRawData) Then Set GetPartialMatch_Smart = newSet: Exit Function
    
    ' 유니버스의 각 데이터를 스캔
    For i = LBound(vRawData) To UBound(vRawData)
        sDataRaw = CStr(vRawData(i))             ' 원본 데이터(Raw)
        
        ' 표시 형식이 있다면 포맷팅된 문자열(Formatted) 생성
        If sFmt <> "" And sFmt <> "General" Then
            sDataFmt = GetSafeFormattedText(vRawData(i), sFmt)
        Else
            sDataFmt = sDataRaw
        End If
        
        isMatch = False
        ' 키워드 루프: 원본이나 포맷팅된 값 중 하나라도 매칭되면 OK
        For k = LBound(arrKeywords) To UBound(arrKeywords)
            sKw = CStr(arrKeywords(k))
            If InStr(1, sDataRaw, sKw, vbTextCompare) > 0 Or _
               InStr(1, sDataFmt, sKw, vbTextCompare) > 0 Then
                isMatch = True
                Exit For
            End If
        Next k
        
        If isMatch Then newSet.Add vRawData(i)
    Next i
    
    Set GetPartialMatch_Smart = newSet
End Function

' [Helper 2] 고속 텍스트 부분 검색
' ------------------------------------------------------------------------------
Private Function GetPartialMatch_Fast( _
    ByVal setUniverse As Core_clsSet, _
    ByVal arrKeywords As Variant) As Core_clsSet
    
    Dim newSet As New Core_clsSet
    Dim vRawData As Variant
    Dim sData As String, sKw As String
    Dim i As Long, k As Long
    Dim isMatch As Boolean
    
    vRawData = setUniverse.KeysArray
    If UBound(vRawData) < LBound(vRawData) Then Set GetPartialMatch_Fast = newSet: Exit Function
    
    For i = LBound(vRawData) To UBound(vRawData)
        sData = CStr(vRawData(i))
        isMatch = False
        For k = LBound(arrKeywords) To UBound(arrKeywords)
            sKw = CStr(arrKeywords(k))
            If InStr(1, sData, sKw, vbTextCompare) > 0 Then
                isMatch = True
                Exit For
            End If
        Next k
        If isMatch Then newSet.Add vRawData(i)
    Next i
    
    Set GetPartialMatch_Fast = newSet
End Function

' [Helper 3] 결과 데이터 포맷팅 (Set -> String Array)
' ------------------------------------------------------------------------------
Private Function FormatSetToCriteria(targetSet As Core_clsSet, sFmt As String) As Variant
    Dim vArr As Variant, vRes() As String
    Dim i As Long
    
    If targetSet.Count = 0 Then
        FormatSetToCriteria = Array("###NO_MATCH###"): Exit Function
    End If
    
    vArr = targetSet.KeysArray
    ReDim vRes(LBound(vArr) To UBound(vArr))
    
    ' Log "FormatSetToCriteria: 포맷팅 수행 (Fmt: " & sFmt & ", Count: " & targetSet.Count & ")", lvlDebug
    
    ' 엑셀 필터 기준(Criteria)에 맞게 모든 데이터를 문자열화
    For i = LBound(vArr) To UBound(vArr)
        If sFmt <> "" Then
            vRes(i) = GetSafeFormattedText(vArr(i), sFmt)
        Else
            vRes(i) = CStr(vArr(i))
        End If
    Next i
    
    FormatSetToCriteria = vRes
End Function

' [Helper 4] 식별자 문자열(Identity)을 내부 Enum 타입으로 매핑 (업데이트됨)
' ------------------------------------------------------------------------------
Private Function MapIdentityToEnum(sID As String) As FilterDataType
    Select Case UCase(sID) ' 안전을 위해 대문자 통일
        ' [기본 숫자 계열]
        Case "NUMBER":           MapIdentityToEnum = fdNumber
        Case "PERCENT":          MapIdentityToEnum = fdPercent
        
        ' [주의: 숨겨진 값 계열 - New]
        Case "ROUNDED":          MapIdentityToEnum = fdRounded
        Case "PERCENT_ROUNDED":  MapIdentityToEnum = fdPercentRounded
        
        ' [기타 타입]
        Case "DATE":             MapIdentityToEnum = fdDate
        Case "STRING":           MapIdentityToEnum = fdString
        
        ' [특수 상태]
        Case "HYBRID":           MapIdentityToEnum = fdHybrid
        Case "ERROR":            MapIdentityToEnum = fdError
        
        ' [기본값]
        Case Else:               MapIdentityToEnum = fdAuto
    End Select
End Function

