Attribute VB_Name = "modProcessManager"
Option Explicit

' ==============================================================================
' Module: modProcessManager
' 설명: 필터링 실행 프로세스를 총괄하는 오케스트레이터
'       (UI 종속성을 제거하여 매크로/폼 양쪽에서 재사용 가능하도록 설계)
' ==============================================================================

' ------------------------------------------------------------------------------
' 함수명: ExecuteFilterProcess
' 목적: 입력값 -> 엔진계산 -> 필터적용 -> 리포트생성 -> 결과반환의 전과정 수행
' ------------------------------------------------------------------------------
Public Function ExecuteFilterProcess( _
    ByVal vInputData As Variant, _
    ByVal colAsset As clsColumnAsset, _
    ByVal rngFilterArea As Range, _
    ByVal lRelativeCol As Long, _
    ByVal bExclude As Boolean, _
    ByVal bStack As Boolean, _
    ByVal bPartial As Boolean) As Object
    
    Dim dictProcessResult As Object
    Set dictProcessResult = CreateObject("Scripting.Dictionary")
    
    ' 기본값 설정 (실패 가정)
    dictProcessResult("IsSuccess") = False
    
    On Error GoTo ErrorHandler

    ' --------------------------------------------------------------------------
    ' 1. 매칭 엔진 가동 (CalculateCriteria)
    ' --------------------------------------------------------------------------
    Dim dictCriteria As Object
    ' 부분일치 여부에 따라 Not 연산 주의 (기존 로직 유지)
    Set dictCriteria = modMatchingEngine.CalculateCriteria( _
                        vInputData, colAsset, Not bPartial, bExclude)
    
    ' 엔진에서 Nothing을 반환하면(혼종 데이터 등) 중단
    If dictCriteria Is Nothing Then
        dictProcessResult("ErrorMsg") = "매칭 엔진에서 처리를 중단했습니다."
        Set ExecuteFilterProcess = dictProcessResult
        Exit Function
    End If
    
    ' --------------------------------------------------------------------------
    ' 2. 엑셀 필터 물리적 적용 (ApplyFilterResult)
    ' --------------------------------------------------------------------------
    modResultHandler.ApplyFilterResult rngFilterArea, lRelativeCol, dictCriteria, bStack
    
    ' --------------------------------------------------------------------------
    ' 3. 결과 리포트 생성 (GenerateResultReport)
    ' --------------------------------------------------------------------------
    Dim sLogMsg As String, sStatusMsg As String
    modResultHandler.GenerateResultReport dictCriteria, sLogMsg, sStatusMsg
    
    ' --------------------------------------------------------------------------
    ' 4. 누락 데이터 분석 (Missing Data Logic)
    ' --------------------------------------------------------------------------
    Dim vMissing As Variant
    vMissing = dictCriteria("Missing")
    
    ' --------------------------------------------------------------------------
    ' 5. 최종 결과 패키징
    ' --------------------------------------------------------------------------
    dictProcessResult("IsSuccess") = True
    dictProcessResult("LogMessage") = sLogMsg
    dictProcessResult("StatusMessage") = sStatusMsg
    dictProcessResult("MissingData") = vMissing
    
    Set ExecuteFilterProcess = dictProcessResult
    Exit Function

ErrorHandler:
    dictProcessResult("IsSuccess") = False
    dictProcessResult("ErrorNumber") = Err.Number
    dictProcessResult("ErrorMsg") = Err.Description
    Set ExecuteFilterProcess = dictProcessResult
End Function

' ------------------------------------------------------------------------------
' 함수명: ProcessMissingData
' 목적: 누락된 데이터가 있을 경우 클립보드에 복사하고 결과 메시지를 반환
' ------------------------------------------------------------------------------
Public Function ProcessMissingData(ByVal vMissing As Variant) As String
    Dim iMissingCnt As Long
    Dim sResultMsg As String
    
    If IsArray(vMissing) Then
        iMissingCnt = UBound(vMissing) - LBound(vMissing) + 1
    Else
        iMissingCnt = 0
    End If
    
    If iMissingCnt > 0 Then
        Dim copySuccess As Boolean
        Dim sMissingList As String
        
        sMissingList = Join(vMissing, vbCrLf)
        
        On Error Resume Next
        copySuccess = ClipboardAPI.SetClipboardText(sMissingList)
        On Error GoTo 0
        
        If copySuccess Then
            sResultMsg = "완료! (누락 " & iMissingCnt & "건 클립보드 복사됨)"
        Else
            sResultMsg = "완료 (누락 " & iMissingCnt & "건)" & vbCrLf & "클립보드 복사 실패"
        End If
    Else
        sResultMsg = "" ' 누락 없음
    End If
    
    ProcessMissingData = sResultMsg
End Function
