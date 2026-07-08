Attribute VB_Name = "modCache"
' @desc 열 자산 캐시와 시트 세대 관리 담당
' @depends modColumnType, clsColumnAsset, clsSet, modLogger
Option Explicit

' ========================================================
' Module: modCache
' 설명: 전역 열 캐시 관리자 및 데이터 전처리 담당
' ========================================================

' 전역 저장소 (매크로 실행 종료 후에도 메모리에 생존)
Private pColCache As Object

' [세대 계약 D6] 시트별 세대 카운터: 무효화(편집/재계산)마다 증가
' 자산은 취득 시점 세대를 스탬프하고, 파이프라인 진입점이 불일치 시 fail-fast(9998)
Private pGenBySheet As Object

' ========================================================
' [세대 계약] 내부 도우미
' ========================================================
Private Function EnsureGenStore() As Object
    If pGenBySheet Is Nothing Then
        Set pGenBySheet = CreateObject("Scripting.Dictionary")
        pGenBySheet.CompareMode = vbTextCompare
    End If
    Set EnsureGenStore = pGenBySheet
End Function

' 캐시 키('[통합문서]시트'!$A$2:$A$9 형태)에서 시트 식별 접두부만 추출
Private Function PatternFromKey(ByVal sKey As String) As String
    Dim p As Long
    p = InStr(1, sKey, "!")
    If p > 0 Then PatternFromKey = Left(sKey, p - 1) Else PatternFromKey = sKey
End Function

' @desc 시트 식별 접두부 받아 현재 세대 반환, 기록 없으면 0
Public Function GetSheetGeneration(ByVal sPattern As String) As Long
    Dim d As Object: Set d = EnsureGenStore()
    If d.exists(sPattern) Then GetSheetGeneration = d(sPattern)
End Function

' @desc 컬럼 자산 받아 세대 불일치 여부 반환, 취득 후 시트 편집 발생 시 True
Public Function IsAssetStale(ByVal colAsset As clsColumnAsset) As Boolean
    If colAsset Is Nothing Then Exit Function
    IsAssetStale = (colAsset.Generation <> GetSheetGeneration(PatternFromKey(colAsset.ID)))
End Function

' ========================================================
' [Public] 자산 요청 (Get or Create)
' 설명: 캐시가 있으면 즉시 반환, 없으면 시트 읽어서 생성 후 반환
' ========================================================
Public Function GetColumnAsset(ByVal rngCol As Range) As clsColumnAsset
    Dim sKey As String
    Dim newAsset As clsColumnAsset
    Dim newSet As clsSet
    Dim vRaw As Variant
    Dim vFlat As Variant
    Dim sFmt As String
    Dim sType As String ' <--- [New] 타입 정보를 담을 변수 선언
    
    ' [LOG] 요청 시작 알림
    Log "GetColumnAsset: 자산 요청 시작 (" & rngCol.Address(False, False) & ")", lvlDebug
    
    ' 1. 저장소 없으면 생성
    If pColCache Is Nothing Then
        Set pColCache = CreateObject("Scripting.Dictionary")
        Log "GetColumnAsset: 저장소(Dictionary) 초기화됨", lvlDebug
    End If
    
    ' 2. 캐시 키 생성 (파일경로+시트명+주소 -> 고유값)
    sKey = rngCol.Address(External:=True)
    
    ' 3. [HIT] 캐시가 존재하면 즉시 반환
    If pColCache.exists(sKey) Then
        Set GetColumnAsset = pColCache(sKey)
        Log "GetColumnAsset: [CACHE HIT] 기존 자산 반환 - " & sKey, lvlInfo
        Exit Function
    End If
    
    ' 4. [MISS] 캐시 생성 시작 (I/O 발생)
    Log "GetColumnAsset: [CACHE MISS] 신규 자산 생성 시작", lvlInfo
    
    ' (0) [Critical] 데이터 로딩 전, 신분(Type)부터 확인!
    ' modColumnType에 있는 판별 함수를 호출합니다.
    ' 여기서 'HYBRID'나 'ERROR'가 나와도 일단 자산에는 기록해두고,
    ' 판단(Stop or Go)은 메인 로직에게 맡깁니다.
    sType = modColumnType.GetColumnIdentity(rngCol)
    Log "GetColumnAsset: 컬럼 Identity 판별 -> " & sType, lvlDebug
    
    ' (1) 데이터 읽기
    Dim rTarget As Range
    Dim lLastRow As Long
    Dim lEndRow As Long
    Dim lUsedRow As Long
    
    ' [Robustness] UsedRange 대신 실제 데이터가 있는 마지막 행을 찾아 범위를 설정
    ' [Fix v2] End(xlUp)이 필터링된 상태에서 간혹 보이는 행까지만 인식하는 현상(Edge Case) 방어
    ' UsedRange와 비교하여 더 큰 행 번호를 채택합니다. (데이터 누락 방지 최우선)
    
    ' 1. End(xlUp) 체크
    lEndRow = rngCol.Parent.Cells(rngCol.Parent.Rows.Count, rngCol.Column).End(xlUp).Row
    
    ' 2. UsedRange 체크
    On Error Resume Next
    lUsedRow = rngCol.Parent.UsedRange.Row + rngCol.Parent.UsedRange.Rows.Count - 1
    On Error GoTo 0
    
    ' 3. 더 넓은 범위 채택 (빈 셀을 좀 더 읽더라도 누락보다는 안전함)
    If lUsedRow > lEndRow Then lLastRow = lUsedRow Else lLastRow = lEndRow
    
    Log "GetColumnAsset: [DEBUG] LastRow Check v2 (End: " & lEndRow & ", Used: " & lUsedRow & ") -> Final: " & lLastRow, lvlDebug
    
    If lLastRow < rngCol.Row Then
        Set rTarget = Nothing
    Else
        Set rTarget = rngCol.Parent.Range(rngCol.Parent.Cells(rngCol.Row, rngCol.Column), _
                                          rngCol.Parent.Cells(lLastRow, rngCol.Column))
    End If
    
    If rTarget Is Nothing Then ' 데이터가 아예 없는 경우
        Log "GetColumnAsset: [Warning] 유효 데이터 없음 (Empty Column)", lvlWarning
        vRaw = Array()
        sFmt = "General"
    Else
        Log "GetColumnAsset: 데이터 로드 범위 확정 (" & rTarget.Address(False, False) & ")", lvlDebug
        vRaw = rTarget.Value
        ' 첫 번째 셀의 서식 추출
        On Error Resume Next
        sFmt = rTarget.Cells(1, 1).NumberFormatLocal
        On Error GoTo 0
        Log "GetColumnAsset: 대표 서식 추출 -> " & sFmt, lvlDebug
    End If
    
    ' (2) 2차원 -> 1차원 평탄화 (Flatten)
    vFlat = FlattenArray(vRaw)
    
    ' (3) clsSet 생성 및 적재
    Set newSet = New clsSet
    newSet.LoadFromArray vFlat
    
    ' (4) 자산 객체 포장 (Init 호출 시 sType 전달!)
    Set newAsset = New clsColumnAsset
    ' ---------------------------------------------------------
    ' [수정됨] 인자가 3개에서 4개로 늘어났습니다.
    ' ---------------------------------------------------------
    ' [세대 계약] 취득 시점의 시트 세대를 스탬프
    newAsset.Init newSet, sFmt, sKey, sType, GetSheetGeneration(PatternFromKey(sKey))
    
    ' (5) 저장소 등록
    pColCache.Add sKey, newAsset
    Log "GetColumnAsset: 캐시 등록 완료 (Key: " & sKey & ")", lvlInfo
    
    ' (6) 반환
    Set GetColumnAsset = newAsset
    
    ' 메모리 정리
    Erase vRaw
    Erase vFlat
End Function

' ========================================================
' [Helper] 2차원 배열 -> 1차원 배열 고속 변환
' ========================================================
Private Function FlattenArray(ByVal vArr As Variant) As Variant
    Dim i As Long
    Dim rMin As Long, rMax As Long
    Dim vRes() As Variant
    Dim vVal As Variant
    
    ' [Case 1] 단일 값(Single Value) 처리
    If Not IsArray(vArr) Then
        Log "FlattenArray: 배열 아님 (단일 값 처리)", lvlDebug
        
        ' 여기도 방역 필수! 단일 값이 에러라면 텍스트로 변환
        If IsError(vArr) Then
            FlattenArray = Array(CVErrToString(vArr))
        Else
            FlattenArray = Array(vArr)
        End If
        Exit Function
    End If
    
    ' [Case 2] 배열(Array) 처리
    rMin = LBound(vArr, 1)
    rMax = UBound(vArr, 1)
    
    ReDim vRes(rMin To rMax)
    
    For i = rMin To rMax
        vVal = vArr(i, 1)
        
        ' 배열 내부 값 방역
        If IsError(vVal) Then
            vRes(i) = CVErrToString(vVal)
        Else
            vRes(i) = vVal
        End If
    Next i
    
    FlattenArray = vRes
End Function

' [Helper] 에러 값을 검색용 텍스트로 변환
Private Function CVErrToString(ByVal vErr As Variant) As String
    Select Case vErr
        Case CVErr(xlErrNA):    CVErrToString = "#N/A"
        Case CVErr(xlErrRef):   CVErrToString = "#REF!"
        Case CVErr(xlErrValue): CVErrToString = "#VALUE!"
        Case CVErr(xlErrDiv0):  CVErrToString = "#DIV/0!"
        Case CVErr(xlErrNum):   CVErrToString = "#NUM!"
        Case CVErr(xlErrName):  CVErrToString = "#NAME?"
        Case CVErr(xlErrNull):  CVErrToString = "#NULL!"
        Case Else:              CVErrToString = "#Error"
    End Select
End Function

' ========================================================
' [Public] 캐시 초기화 (데이터 변경 시 호출)
' ========================================================
Public Sub InvalidateCache()
    If Not pColCache Is Nothing Then
        Log "InvalidateCache: 전역 캐시 삭제 (" & pColCache.Count & " items)", lvlInfo
        pColCache.RemoveAll
        Set pColCache = Nothing
    'Else
        'Log "InvalidateCache: 삭제할 캐시 없음", lvlDebug
    End If

    ' [세대 계약] 전체 무효화 = 알려진 모든 시트의 세대 증가
    If Not pGenBySheet Is Nothing Then
        Dim vPattern As Variant
        For Each vPattern In pGenBySheet.Keys
            pGenBySheet(vPattern) = pGenBySheet(vPattern) + 1
        Next vPattern
    End If
End Sub

' ========================================================
' [New] 부분 캐시 삭제 (정교한 파괴)
' 설명: 전체를 날리지 않고, 변경된 시트(Sh)와 관련된 컬럼만 골라서 제거
' ========================================================
Public Sub InvalidateCacheBySheet(ByVal Sh As Object)
    Dim vKey As Variant
    Dim sTargetPattern As String
    Dim colToDelete As New Collection
    Dim i As Long
    
    ' 1. [핵심 수정] 삭제 대상 패턴을 '정확하게' 생성하기
    ' 설명: 우리가 문자열을 조립하지 말고, 해당 시트의 A1 셀 주소를
    '       External 모드로 뽑은 뒤, 뒤쪽 셀 주소($A$1)만 떼어냅니다.
    '       이렇게 하면 띄어쓰기나 작은따옴표(') 처리가 저장할 때와 100% 일치합니다.
    ' 예시:
    '   저장된 키: '[내 파일.xlsx]시트 1'!$A:$A
    '   기존 방식: [내 파일.xlsx]시트 1  <-- 작은따옴표 없어서 매칭 실패!
    '   수정 방식: '[내 파일.xlsx]시트 1'!$A$1  <-- 여기서 ! 앞까지만 추출
    
    On Error Resume Next
    sTargetPattern = Sh.Cells(1, 1).Address(External:=True)
    On Error GoTo 0
    
    ' ! 문자를 기준으로 앞부분(파일+시트명)만 가져옴
    If InStr(1, sTargetPattern, "!") > 0 Then
        sTargetPattern = Left(sTargetPattern, InStr(1, sTargetPattern, "!") - 1)
    End If
    If LenB(sTargetPattern) = 0 Then Exit Sub

    ' 2. [세대 계약] 무효화 = 해당 시트 세대 증가 (캐시 보유 여부와 무관)
    Dim dGen As Object: Set dGen = EnsureGenStore()
    dGen(sTargetPattern) = GetSheetGeneration(sTargetPattern) + 1

    ' 3. 저장소 확인 (이하 캐시 항목 정리)
    If pColCache Is Nothing Then Exit Sub
    If pColCache.Count = 0 Then Exit Sub

    ' 디버깅 로그: 내가 지금 뭘 찾으려는지 확인
    'Log "DEBUG: 삭제 패턴 검색 -> " & sTargetPattern, lvlDebug

    ' 4. 삭제할 키 수집
    For Each vKey In pColCache.Keys
        ' 이제 패턴이 정확하므로 안심하고 검색
        If InStr(1, vKey, sTargetPattern, vbTextCompare) > 0 Then
            colToDelete.Add vKey
        End If
    Next vKey
    
    ' 5. 삭제 실행 및 결과 로그
    If colToDelete.Count > 0 Then
        For i = 1 To colToDelete.Count
            pColCache.Remove colToDelete(i)
        Next i
        ' 삭제되었다면 확실하게 로그 남기기
        Log "InvalidateCacheBySheet: [Success] " & colToDelete.Count & "건 삭제됨 (" & sTargetPattern & ")", lvlInfo
    Else
        ' 변경 이벤트는 왔는데 매칭되는 키가 없는 경우 (이미 지워졌거나, 해당 시트는 캐시된 적 없음)
         Log "InvalidateCacheBySheet: [Skip] 삭제 대상 없음 (" & sTargetPattern & ")", lvlDebug
    End If
    
    Set colToDelete = Nothing
End Sub


' ========================================================
' [Debug] 현재 캐시 상태 전수 조사 (직접 실행용)
' 설명: 직접실행 창(Ctrl+G)에 현재 메모리에 든 모든 키를 출력합니다.
' ========================================================
Public Sub Debug_DumpCache()
    Dim vKey As Variant
    Dim i As Long
    
    Debug.Print "========== [CACHE DUMP START] =========="
    
    If pColCache Is Nothing Then
        Debug.Print "상태: 저장소(Dictionary)가 초기화되지 않음 (Nothing)"
    ElseIf pColCache.Count = 0 Then
        Debug.Print "상태: 저장소는 있으나 비어있음 (Count: 0)"
    Else
        Debug.Print "총 캐시 개수: " & pColCache.Count
        i = 1
        For Each vKey In pColCache.Keys
            Debug.Print Format(i, "00") & ". Key: " & vKey
            i = i + 1
        Next vKey
    End If
    
    Debug.Print "========== [CACHE DUMP END] ============"
End Sub
