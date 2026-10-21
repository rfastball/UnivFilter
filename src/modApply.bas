Attribute VB_Name = "modApply"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc 표시 키 집합의 AutoFilter 적용 담당
' @depends clsSet, modLogger
Option Explicit

' 실측 근거 (AutoFilterTest gap-probe-findings.md):
'  - xlFilterValues 배열이 단일 Tier (날짜/회계 접힌형/분수 전부 소화, P2/P5/P6)
'  - 배열 항목 255자 초과는 조용히 0건(P1) -> 등호 스칼라(<=2건) 또는 탈락 기록
'  - 빈칸 포함은 배열 항목 "=" (P3)

Private Const MAX_ITEM_LEN As Long = 255
Private Const NOMATCH_SENTINEL As String = "###UF_NOMATCH###"

' @desc 표시 키 집합 적용해 결과(모드/적용수/탈락) 딕셔너리 반환, 키 없으면 무적용 반환
Public Function ApplyDisplays(ByVal rngFilterArea As Range, ByVal lField As Long, _
                              ByVal displaySet As clsSet, ByVal bIncludeBlanks As Boolean, _
                              ByVal bStack As Boolean) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Mode") = "none"
    result("Applied") = 0&
    result("DroppedLong") = Array()

    ' 항목 분류: 배열 가능(<=255자) / 초과
    Dim shortItems As New Collection, longItems As New Collection
    Dim vKey As Variant
    For Each vKey In displaySet.KeysArray
        If Len(CStr(vKey)) > MAX_ITEM_LEN Then
            longItems.Add CStr(vKey)
        Else
            shortItems.Add CStr(vKey)
        End If
    Next vKey

    Dim nShort As Long, nLong As Long
    nShort = shortItems.Count
    nLong = longItems.Count

    If nShort = 0 And nLong = 0 And Not bIncludeBlanks Then
        Set ApplyDisplays = result
        Exit Function
    End If

    If nShort = 0 And Not bIncludeBlanks And nLong <= 2 Then
        ' 등호 스칼라 경로: 긴 항목만 1~2건 (P1: 등호는 길이 무제한)
        If Not bStack Then ResetFilters rngFilterArea
        If nLong = 1 Then
            rngFilterArea.AutoFilter Field:=lField, Criteria1:="=" & longItems(1)
        Else
            rngFilterArea.AutoFilter Field:=lField, Criteria1:="=" & longItems(1), _
                Operator:=xlOr, Criteria2:="=" & longItems(2)
        End If
        result("Mode") = "scalar"
        result("Applied") = nLong
    Else
        ' 기본 경로: xlFilterValues 배열 (긴 항목은 탈락 기록)
        Dim arr() As Variant
        Dim n As Long
        n = nShort + IIf(bIncludeBlanks, 1, 0)
        If n = 0 Then
            Err.Raise vbObjectError + 2101, "modApply.ApplyDisplays", _
                "255자 초과 값만 3개 이상이어서 적용할 수 없습니다. 기존 필터를 유지합니다."
        End If
        ReDim arr(0 To n - 1)
        Dim i As Long
        ' 숫자 인덱스 반복 조회는 대규모 Collection에서 급격히 느려진다.
        For Each vKey In shortItems
            arr(i) = vKey
            i = i + 1
        Next vKey
        If bIncludeBlanks Then arr(n - 1) = "="

        If Not bStack Then ResetFilters rngFilterArea
        rngFilterArea.AutoFilter Field:=lField, Criteria1:=arr, Operator:=xlFilterValues
        result("Mode") = "array"
        result("Applied") = n
        If nLong > 0 Then
            result("DroppedLong") = ToArray(longItems)
            Log "ApplyDisplays: 255자 초과 항목 " & nLong & "건 탈락", lvlWarning
        End If
    End If

    Set ApplyDisplays = result
End Function

' @desc 매칭 0건 상태를 전체 숨김으로 적용함(포함 모드 폼 경로 전용)
Public Sub ApplyHideAll(ByVal rngFilterArea As Range, ByVal lField As Long, _
                        ByVal bStack As Boolean)
    If Not bStack Then ResetFilters rngFilterArea
    rngFilterArea.AutoFilter Field:=lField, Criteria1:="=" & NOMATCH_SENTINEL
End Sub

' @desc 지정 필드의 필터만 해제함(제외 0건의 no-op 경로)
Public Sub ClearField(ByVal rngFilterArea As Range, ByVal lField As Long)
    On Error Resume Next
    rngFilterArea.AutoFilter Field:=lField
    On Error GoTo 0
End Sub

' @desc 영역의 모든 필터 조건을 해제함(Reset 모드)
Public Sub ResetFilters(ByVal rngFilterArea As Range)
    On Error Resume Next
    If Not rngFilterArea.ListObject Is Nothing Then
        rngFilterArea.ListObject.AutoFilter.ShowAllData
    Else
        rngFilterArea.Parent.ShowAllData
    End If
    On Error GoTo 0
End Sub

' @desc 필터 영역의 가시 데이터 행수 반환, 전부 숨김이면 0 반환
Public Function CountVisibleRows(ByVal rngFilterArea As Range) As Long
    On Error GoTo Fail
    Dim body As Range
    ' 헤더 1행 제외한 본문 첫 열 기준
    Set body = rngFilterArea.Parent.Range( _
        rngFilterArea.Cells(2, 1), _
        rngFilterArea.Cells(rngFilterArea.Rows.Count, 1))
    CountVisibleRows = body.SpecialCells(xlCellTypeVisible).Count
    Exit Function
Fail:
    ' SpecialCells는 가시 셀 0개에서 오류를 낸다
    CountVisibleRows = 0
End Function

Private Function ToArray(ByVal col As Collection) As Variant
    Dim arr() As Variant
    ReDim arr(0 To col.Count - 1)
    Dim i As Long
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    ToArray = arr
End Function
