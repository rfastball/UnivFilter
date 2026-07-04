VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Core_clsSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @desc 대량 데이터용 대소문자 무시 집합 엔진
' @depends Core_Enums
Option Explicit

Private mDict As Object

Private Sub Class_Initialize()
    Set mDict = CreateObject("Scripting.Dictionary")
    mDict.CompareMode = vbTextCompare
End Sub

Private Sub Class_Terminate()
    Set mDict = Nothing
End Sub

' @desc 내부 Dictionary를 비워 재사용 가능한 상태로 전환
Public Sub Clear()
    mDict.RemoveAll
End Sub

' @desc 빈 값이 아닌 단일 값을 집합에 추가
Public Sub Add(ByVal vItem As Variant)
    If Not IsEmpty(vItem) And Not IsNull(vItem) And Not IsError(vItem) Then
        If Not mDict.exists(vItem) Then mDict(vItem) = Empty
    End If
End Sub

' @desc 단일 값을 집합에서 제거
Public Sub Remove(ByVal vItem As Variant)
    If Not IsNull(vItem) And Not IsError(vItem) Then
        If mDict.exists(vItem) Then mDict.Remove vItem
    End If
End Sub

' @desc 단일 값의 집합 포함 여부 반환
Public Function Contains(ByVal vItem As Variant) As Boolean
    If IsNull(vItem) Or IsError(vItem) Then Exit Function
    Contains = mDict.exists(vItem)
End Function

Public Property Get Count() As Long
    Count = mDict.Count
End Property

' @desc 배열 또는 단일 값을 집합에 고속 적재
Public Sub LoadFromArray(ByVal vArr As Variant)
    Dim i As Long
    Dim vVal As Variant

    If Not IsArray(vArr) Then
        Me.Add vArr
        Exit Sub
    End If

    For i = LBound(vArr) To UBound(vArr)
        vVal = vArr(i)
        If Not IsEmpty(vVal) And Not IsNull(vVal) And Not IsError(vVal) Then
            mDict(vVal) = Empty
        End If
    Next i
End Sub

' @desc Single-pass로 교집합과 양방향 차집합을 분류해 반환
Public Function Compare(ByVal otherSet As Core_clsSet, _
                        Optional ByVal mode As CoreSetRetMode = Core_Enums.srAllSplit) As Variant
    Dim setCommon As Core_clsSet, setDiffOther As Core_clsSet, setDiffMe As Core_clsSet
    Dim dictCommon As Object, dictDiffOther As Object, dictDiffMe As Object
    Dim vKey As Variant, arrOther As Variant
    Dim i As Long

    For Each vKey In mDict.Keys
        mDict(vKey) = Empty
    Next vKey

    If mode = Core_Enums.srIntersect Or mode = Core_Enums.srAllSplit Then
        Set setCommon = New Core_clsSet
        Set dictCommon = setCommon.InternalDict
    End If
    If mode = Core_Enums.srDiffOther Or mode = Core_Enums.srAllSplit Then
        Set setDiffOther = New Core_clsSet
        Set dictDiffOther = setDiffOther.InternalDict
    End If
    If mode = Core_Enums.srDiffMe Or mode = Core_Enums.srAllSplit Then
        Set setDiffMe = New Core_clsSet
        Set dictDiffMe = setDiffMe.InternalDict
    End If

    arrOther = otherSet.KeysArray
    If Core_HasArrayItems(arrOther) Then
        For i = LBound(arrOther) To UBound(arrOther)
            vKey = arrOther(i)
            If mDict.exists(vKey) Then
                mDict(vKey) = True
                If Not dictCommon Is Nothing Then dictCommon(vKey) = Empty
            Else
                If Not dictDiffOther Is Nothing Then dictDiffOther(vKey) = Empty
            End If
        Next i
    End If

    If Not dictDiffMe Is Nothing Then
        For Each vKey In mDict.Keys
            If IsEmpty(mDict(vKey)) Then dictDiffMe(vKey) = Empty
        Next vKey
    End If

    Select Case mode
        Case Core_Enums.srIntersect: Set Compare = setCommon
        Case Core_Enums.srDiffOther: Set Compare = setDiffOther
        Case Core_Enums.srDiffMe:    Set Compare = setDiffMe
        Case Core_Enums.srAllSplit:  Compare = Array(setCommon, setDiffMe, setDiffOther)
    End Select
End Function

' @desc 교집합을 새 집합으로 반환
Public Function Intersection(ByVal otherSet As Core_clsSet) As Core_clsSet
    Set Intersection = Compare(otherSet, Core_Enums.srIntersect)
End Function

' @desc 내 차집합을 새 집합으로 반환
Public Function Difference(ByVal otherSet As Core_clsSet) As Core_clsSet
    Set Difference = Compare(otherSet, Core_Enums.srDiffMe)
End Function

' @desc 상대 차집합을 새 집합으로 반환
Public Function DifferenceInput(ByVal otherSet As Core_clsSet) As Core_clsSet
    Set DifferenceInput = Compare(otherSet, Core_Enums.srDiffOther)
End Function

' @desc 합집합을 새 집합으로 반환
Public Function Union(ByVal otherSet As Core_clsSet) As Core_clsSet
    Dim newSet As New Core_clsSet
    Dim targetDict As Object
    Dim vArr As Variant
    Dim i As Long

    Set targetDict = newSet.InternalDict

    If mDict.Count > 0 Then
        vArr = mDict.Keys
        For i = 0 To UBound(vArr)
            targetDict(vArr(i)) = Empty
        Next i
    End If

    If otherSet.Count > 0 Then
        vArr = otherSet.KeysArray
        For i = 0 To UBound(vArr)
            targetDict(vArr(i)) = Empty
        Next i
    End If

    Set Union = newSet
End Function

' @desc 집합 키를 1차원 배열로 반환
Public Function ToArray() As Variant
    If mDict.Count = 0 Then ToArray = Array() Else ToArray = mDict.Keys
End Function

Public Property Get InternalDict() As Object
    Set InternalDict = mDict
End Property

Public Property Get KeysArray() As Variant
    If mDict.Count = 0 Then KeysArray = Array() Else KeysArray = mDict.Keys
End Property

Private Function Core_HasArrayItems(ByVal vArr As Variant) As Boolean
    On Error GoTo EmptyArray
    Core_HasArrayItems = (UBound(vArr) >= LBound(vArr))
    Exit Function
EmptyArray:
    Core_HasArrayItems = False
End Function
