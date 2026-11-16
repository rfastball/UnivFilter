Attribute VB_Name = "ClipboardAPI"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc Win32 클립보드 텍스트 읽기/쓰기 담당
Option Explicit

' ========================================================================================
' 모듈: ClipboardAPI
' 설명: Windows API를 사용하여 유니코드(Unicode) 텍스트를 클립보드에 읽고 쓰는 기능 제공
' ========================================================================================

' ----------------------------------------------------------------------------------------
' [1] Windows API 선언 (64비트/32비트 호환 PtrSafe 사용)
' ----------------------------------------------------------------------------------------

' 1.1 클립보드 제어 함수
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr

' 1.2 메모리 관리 함수 (클립보드 데이터는 전역 메모리를 사용해야 함)
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

' 1.3 문자열 처리 함수 (유니코드/WideChar 전용)
Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

' ----------------------------------------------------------------------------------------
' [2] 상수 선언
' ----------------------------------------------------------------------------------------
Private Const GMEM_ZEROINIT As Long = &H40   ' 빈 문자열과 종료 문자를 0으로 초기화
Private Const CF_UNICODETEXT As Long = 13     ' 유니코드 텍스트 포맷 ID
Private Const GMEM_MOVEABLE As Long = &H2     ' 이동 가능한 메모리 (Windows가 힙 내에서 위치 조정 가능)


' ========================================================================================
' 함수명: SetClipboardText
' 설  명: 텍스트 문자열을 클립보드에 복사합니다. (유니코드 지원)
' 반  환: 성공 시 True, 실패 시 False
' ========================================================================================
Public Function SetClipboardText(ByVal text As String) As Boolean
    Dim hMem As LongPtr, lpData As LongPtr
    Dim opened As Boolean
    On Error GoTo Cleanup

    ' 기존 클립보드를 비우기 전에 종료 문자를 포함한 데이터를 준비한다.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(text) + 2)
    If hMem = 0 Then Exit Function
    lpData = GlobalLock(hMem)
    If lpData = 0 Then GoTo Cleanup
    If LenB(text) > 0 Then
        If lstrcpyW(lpData, StrPtr(text)) = 0 Then GoTo Cleanup
    End If
    GlobalUnlock hMem
    lpData = 0

    opened = (OpenClipboard(Application.hwnd) <> 0)
    If Not opened Then GoTo Cleanup
    If EmptyClipboard() = 0 Then GoTo Cleanup
    If SetClipboardData(CF_UNICODETEXT, hMem) <> 0 Then
        hMem = 0 ' 소유권이 시스템으로 넘어간 메모리는 해제하지 않는다.
        SetClipboardText = True
    End If
Cleanup:
    On Error Resume Next
    If lpData <> 0 Then GlobalUnlock hMem
    If hMem <> 0 Then GlobalFree hMem
    If opened Then CloseClipboard
End Function


' ========================================================================================
' 함수명: GetClipboardText
' 설  명: 클립보드에 있는 텍스트를 가져옵니다. (유니코드 지원)
' 반  환: 클립보드 텍스트 (없으면 빈 문자열)
' ========================================================================================
Public Function GetClipboardText() As String
    Dim hMem As LongPtr     ' 클립보드 데이터 핸들
    Dim lpData As LongPtr   ' 데이터 실제 주소
    Dim nSize As Long       ' 문자열 길이
    Dim strBuffer As String ' 결과를 담을 버퍼
    
    ' [단계 1] 클립보드 열기
    If OpenClipboard(0) = 0 Then Exit Function
    
    ' [단계 2] 유니코드 텍스트 형식의 데이터 핸들 가져오기
    hMem = GetClipboardData(CF_UNICODETEXT)
    
    If hMem <> 0 Then
        ' [단계 3] 메모리 잠금 및 포인터 획득
        lpData = GlobalLock(hMem)
        
        If lpData <> 0 Then
            ' [단계 4] 문자열 길이 계산 (Null 문자 제외한 길이)
            nSize = lstrlenW(lpData)
            
            ' [단계 5] VBA 문자열 버퍼 확보
            strBuffer = String(nSize, vbNullChar)
            
            ' [단계 6] 데이터 복사 (클립보드 -> VBA 변수)
            If nSize > 0 Then lstrcpyW StrPtr(strBuffer), lpData
            
            ' [단계 7] 잠금 해제
            GlobalUnlock hMem
        End If
    End If
    
    ' [단계 8] 클립보드 닫기
    CloseClipboard
    
    ' 결과 반환
    GetClipboardText = strBuffer
End Function
