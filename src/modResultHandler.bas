Attribute VB_Name = "modResultHandler"
Option Explicit

' ==============================================================================
' [섹션 1] 실행 엔진 (Action: AutoFilter Execution)
' 목적: 계산된 필터 조건(Criteria)을 실제 워크시트에 물리적으로 적용합니다.
' ==============================================================================
' 파일명: modResultHandler.bas
' 함수: ApplyFilterResult

Public Sub ApplyFilterResult( _
    ByVal rngTable As Range, _
    ByVal lColIndex As Long, _
    ByVal dictResult As Object, _
    ByVal bStack As Boolean)
    
    Dim arrCriteria As Variant
    Dim sType As String
    Dim sFmt As String
    Dim i As Long
    
    ' [LOG] 필터 적용 시작
    Log "ApplyFilterResult: 필터 적용 시작 (Col: " & lColIndex & ", Stack: " & bStack & ")", lvlInfo
    
    On Error GoTo ErrorHandler
    Call OptimizedMode(True)
    
    ' 1. 필터 모드 초기화
    If Not bStack Then
        If rngTable.Parent.FilterMode Then
            rngTable.Parent.ShowAllData
            Log "ApplyFilterResult: 기존 필터 초기화(ShowAllData) 완료", lvlDebug
        End If
    End If
    
    ' 2. 기준값 및 서식 정보 추출
    arrCriteria = dictResult("Criteria")
    sType = modFilterUtils.GetColumnIdentity(rngTable.Columns(lColIndex))
    sFmt = modFilterUtils.GetRepresentativeFormat(rngTable.Columns(lColIndex))
    
    ' [LOG] 타겟 컬럼 정보 확인
    Log "ApplyFilterResult: 컬럼 분석 (Type: " & sType & ", Fmt: " & sFmt & ")", lvlDebug
    
    ' 서식 성격 판별
    Dim isAcc As Boolean: isAcc = (InStr(sFmt, "$") > 0) Or (InStr(sFmt, "\") > 0) Or (InStr(sFmt, "*") > 0)
    Dim isPct As Boolean: isPct = (InStr(sFmt, "%") > 0)

    ' ==========================================================================
    ' [CORE LOGIC] 데이터 타입별 지능형 변환 및 최종 문자열 확정 (Unified)
    ' 설명: HYBRID, NUMBER, ROUNDED 등 모든 타입에 대해
    '       "화면에 보이는 값(WYSIWYG)"으로 변환하여 엑셀 필터 오동작을 방지함.
    ' ==========================================================================
    If IsArray(arrCriteria) Then
        
        ' [방어] 더미 데이터(###_DUMMY...)는 변환하지 않고 패스
        Dim bSkip As Boolean: bSkip = False
        If UBound(arrCriteria) >= LBound(arrCriteria) Then
             If Left(CStr(arrCriteria(LBound(arrCriteria))), 3) = "###" Then bSkip = True
        End If
        
    ' [modResultHandler] 내 ApplyFilterResult 수정 구문
    If Not bSkip Then
        For i = LBound(arrCriteria) To UBound(arrCriteria)
            
            ' [0순위] 날짜 타입은 다른 거 묻지도 따지지도 말고 가로채기!
            If sType = "DATE" Then
                arrCriteria(i) = modFilterUtils.GetSafeFormattedText(arrCriteria(i), sFmt)
            ' [1순위] 날짜가 아니라면, 회계/통화 서식인가? (If가 아니라 ElseIf로 연결!)
            ElseIf isAcc Then
                arrCriteria(i) = modFilterUtils.GetAccountingFilterString(arrCriteria(i), sFmt)
            
            ' [2순위] 그것도 아니면, 백분율 서식인가?
            ElseIf isPct Then
                arrCriteria(i) = modFilterUtils.GetSafeFormattedText(arrCriteria(i), sFmt)
            
            ' [3순위] 나머지는 일반 서식 처리
            Else
                If sFmt <> "" And sFmt <> "General" And sFmt <> "@" Then
                    arrCriteria(i) = modFilterUtils.GetSafeFormattedText(arrCriteria(i), sFmt)
                Else
                    arrCriteria(i) = CStr(arrCriteria(i))
                End If
            End If ' <--- 이 End If 하나가 위의 If / ElseIf 체인 전체를 닫아줍니다!
            
        Next i ' 이제 엑셀이 기분 좋게 다음으로 넘어갑니다. [cite: 318]
    End If
        
        ' Log "ApplyFilterResult: [Unified] 배열 데이터 표준화 완료", lvlDebug
    End If
    ' --------------------------------------------------------------------------
    
    ' 3. [개선] 매칭 데이터 부재 및 빈 배열 처리 (Dummy)
    ' --------------------------------------------------------------------------
    Dim isNothingToFilter As Boolean
    isNothingToFilter = False
    
    If IsArray(arrCriteria) Then
        ' 배열이 비어있거나(0개), 원소가 1개인데 '없음' 신호인 경우
        If UBound(arrCriteria) < LBound(arrCriteria) Then
            isNothingToFilter = True
        ElseIf UBound(arrCriteria) = LBound(arrCriteria) Then
            If arrCriteria(LBound(arrCriteria)) = "###NO_MATCH###" Then isNothingToFilter = True
        End If
    Else
        ' 배열이 아니면서 비어있는 경우
        If IsEmpty(arrCriteria) Or arrCriteria = "" Then isNothingToFilter = True
    End If

    ' 매칭 결과가 없으면 즉시 모든 데이터 숨기고 안전하게 종료 (에러 방지)
    If isNothingToFilter Then
        Log "ApplyFilterResult: [SKIP] 매칭된 데이터 없음 -> Dummy Filter 적용", lvlInfo
        rngTable.AutoFilter Field:=lColIndex, Criteria1:="###_DUMMY_DATA_###"
        GoTo Finalize
    End If
    ' --------------------------------------------------------------------------
    
    ' 4. [핵심] 1개면 껍질 까기 (Unwrap) 로직
    ' 이 부분이 있어야 회계 제로(\-) 등이 와일드카드 없이도 정확히 매칭됩니다.
    If IsArray(arrCriteria) Then
        Dim iCount As Long: iCount = UBound(arrCriteria) - LBound(arrCriteria) + 1
        
        If iCount = 1 Then
            ' 단일 값(Scalar) 모드: 특수 기호 매칭에 유리
            Log "ApplyFilterResult: 단일 값 필터 적용 (" & arrCriteria(LBound(arrCriteria)) & ")", lvlDebug
            rngTable.AutoFilter Field:=lColIndex, Criteria1:=arrCriteria(LBound(arrCriteria))
        Else
            ' 배열 모드: 다중 선택 매칭
            Log "ApplyFilterResult: 다중 값 필터 적용 (Count: " & iCount & ")", lvlDebug
            rngTable.AutoFilter Field:=lColIndex, Criteria1:=arrCriteria, Operator:=xlFilterValues
        End If
    Else
        Log "ApplyFilterResult: 단순 값 필터 적용 (" & arrCriteria & ")", lvlDebug
        rngTable.AutoFilter Field:=lColIndex, Criteria1:=arrCriteria
    End If

Finalize:
    Call OptimizedMode(False)
    Log "ApplyFilterResult: 필터 적용 완료", lvlInfo
    Exit Sub

ErrorHandler:
    Call OptimizedMode(False)
    Log "ApplyFilterResult: [ERROR] " & Err.Description, lvlError
    MsgBox "Filter Error: " & Err.Description
End Sub

' ==============================================================================
' [섹션 2] 피드백 생성 엔진 (Feedback: Reporting Analysis)
' 목적: 엔진 연산 과정에서 발생한 누락 데이터 및 서식 보정 내역을 리포트화합니다.
' ==============================================================================
Public Sub GenerateResultReport( _
    ByVal dictResult As Object, _
    ByRef outMsg As String, _
    ByRef outStatus As String)
    
    Dim vMissing As Variant
    Dim dictVisFail As Object
    Dim iMissingCnt As Long, iVisFailCnt As Long
    Dim sReport As String
    
    ' 결과 딕셔너리로부터 리포트용 데이터 추출
    vMissing = dictResult("Missing")
    Set dictVisFail = dictResult("VisFail")
    
    sReport = ""
    
    ' ---------------------------------------------------------
    ' A. 누락 데이터 보고 (Missing Data Analysis)
    ' ---------------------------------------------------------
    If IsArray(vMissing) Then
        iMissingCnt = UBound(vMissing) - LBound(vMissing) + 1
    Else
        iMissingCnt = 0
    End If
    
    If iMissingCnt > 0 Then
        ' [Fix v2.7] 깨진 이모지 문자를 ASCII 기호로 대체
        sReport = sReport & "[!] [데이터 없음] " & iMissingCnt & "건" & vbCrLf & _
                  "    (" & Join(vMissing, ", ") & ")" & vbCrLf & vbCrLf
    End If
    
    ' ---------------------------------------------------------
    ' B. 자동 보정 보고 (Vis Layer Correction)
    ' ---------------------------------------------------------
    iVisFailCnt = dictVisFail.Count
    
    If iVisFailCnt > 0 Then
        ' [Fix v2.7] 깨진 이모지 문자를 ASCII 기호로 대체
        sReport = sReport & "[i] [서식 자동 보정] " & iVisFailCnt & "건" & vbCrLf
        
        ' 리포트 가독성을 위해 상위 3개 항목만 샘플로 노출
        Dim k As Variant, i As Long
        i = 0
        For Each k In dictVisFail.Keys
            sReport = sReport & "    '" & k & "' → '" & dictVisFail(k) & "'" & vbCrLf
            i = i + 1
            If i >= 3 Then
                sReport = sReport & "    (외 " & (iVisFailCnt - 3) & "건...)" & vbCrLf
                Exit For
            End If
        Next k
        sReport = sReport & vbCrLf
    End If
    
    ' ---------------------------------------------------------
    ' C. 결과 문자열 최종 반환
    ' ---------------------------------------------------------
    outMsg = sReport ' 상세 메시지 박스용
    
    ' 하단 상태표시줄(StatusBar)용 요약 메시지 구성
    If iMissingCnt = 0 And iVisFailCnt = 0 Then
        outStatus = "완료: 모든 항목이 정상적으로 검색되었습니다."
    Else
        outStatus = "완료: 누락(" & iMissingCnt & "), 보정(" & iVisFailCnt & ")"
    End If
    
    ' [LOG] 최종 리포트 요약 기록
    Log "GenerateResultReport: " & outStatus, lvlInfo
    
End Sub


