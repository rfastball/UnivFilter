VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 열 표시키 집합/값 지도/지문 보관 담당
Option Explicit

' 지연 우주(docs/tasks/2026-07-10_lazy_universe_design.md): 우주의 원본은 행별
' 렌더키 배열(RenderedVec)이고, 집합/지도는 그것에 대한 인덱스다. 인덱스를 요구하는
' 소비자(제외/부분일치/자가검증)가 처음 나타날 때 1회 지연 구축한다(고카디널 ~180ms).

' 표시 키 집합: AutoFilter가 실제로 매칭하는 문자열의 우주 (빈칸 제외)
Private mDisplays As clsSet
' 값 지도: CStr(Value2) -> 그 값을 렌더링하는 표시 키들의 vbNullChar 연결 문자열
' (숫자/날짜 셀만. clsSet 대신 문자열인 이유 = 값당 객체 생성 비용, modColumnScan 참조)
Private mValueMap As Object
' 행수 지도: 표시 키 -> 해당 키의 행수 (적용 자가검증용)
Private mRowCountByKey As Object
Private mBlankCount As Long
Private mRowCount As Long
Private mID As String
' 취득 시점 지문. GetColumnAsset이 반환 전 현재 지문과 Matches로 대조(신선도 관측).
Private mSignature As clsColSig
' 행별 표시키 배열(1..RowCount, 빈칸 행은 ""). InitRendered 경로에서만 보유.
Private mRenderedVec() As String
Private mHasRendered As Boolean
' 인덱스(집합/지도) 구축 완료 여부 - Init(주입)은 즉시 True, InitRendered는 지연.
Private mIndexesBuilt As Boolean

' @desc 스캔 산출물 주입받아 자산을 초기화함
Public Sub Init(ByVal displaySet As clsSet, ByVal valueMap As Object, _
                ByVal rowCountByKey As Object, ByVal lBlankCount As Long, _
                ByVal lRowCount As Long, ByVal sID As String, ByVal sig As clsColSig)
    Set mDisplays = displaySet
    Set mValueMap = valueMap
    Set mRowCountByKey = rowCountByKey
    mBlankCount = lBlankCount
    mRowCount = lRowCount
    mID = sID
    Set mSignature = sig
    mHasRendered = False
    mIndexesBuilt = True
End Sub

' @desc 행별 렌더키 배열 주입받아 자산 초기화함, 인덱스는 최초 접근 시 지연 구축
Public Sub InitRendered(ByRef renderedVec() As String, ByVal lBlankCount As Long, _
                        ByVal lRowCount As Long, ByVal sID As String, ByVal sig As clsColSig)
    mRenderedVec = renderedVec
    mBlankCount = lBlankCount
    mRowCount = lRowCount
    mID = sID
    Set mSignature = sig
    mHasRendered = True
    mIndexesBuilt = False
End Sub

Public Property Get HasRendered() As Boolean
    HasRendered = mHasRendered
End Property

' @desc 행별 표시키 배열 반환(값 복사), InitRendered 자산에서만 유효
Public Property Get RenderedVec() As String()
    RenderedVec = mRenderedVec
End Property

Public Property Get Displays() As clsSet
    EnsureIndexes
    Set Displays = mDisplays
End Property

Public Property Get ValueMap() As Object
    EnsureIndexes
    Set ValueMap = mValueMap
End Property

Public Property Get RowCountByKey() As Object
    EnsureIndexes
    Set RowCountByKey = mRowCountByKey
End Property

Public Property Get BlankCount() As Long
    BlankCount = mBlankCount
End Property

Public Property Get HasBlanks() As Boolean
    HasBlanks = (mBlankCount > 0)
End Property

Public Property Get RowCount() As Long
    RowCount = mRowCount
End Property

Public Property Get ID() As String
    ID = mID
End Property

Public Property Get Signature() As clsColSig
    Set Signature = mSignature
End Property

' @desc 표시 키 집합 받아 기대 가시 행수 반환, 미지 키는 0으로 집계
Public Function ExpectedRows(ByVal keySet As clsSet, ByVal bWithBlanks As Boolean) As Long
    EnsureIndexes
    Dim n As Long, vKey As Variant
    For Each vKey In keySet.KeysArray
        If mRowCountByKey.exists(vKey) Then n = n + mRowCountByKey(vKey)
    Next vKey
    If bWithBlanks Then n = n + mBlankCount
    ExpectedRows = n
End Function

' 지연 인덱스 구축: RenderedVec + 지문 ContentVec 1패스로 집합/값지도/행수지도 생성.
' 누적 규칙은 구 렌더 경로와 동일(빈칸 제외, Double만 값지도, 대소문자 무시 키).
Private Sub EnsureIndexes()
    If mIndexesBuilt Then Exit Sub
    mIndexesBuilt = True

    Dim rowByKey As Object
    Set rowByKey = CreateObject("Scripting.Dictionary")
    rowByKey.CompareMode = vbTextCompare
    Dim vm As Object
    Set vm = CreateObject("Scripting.Dictionary")

    Dim v2 As Variant
    v2 = mSignature.ContentVec

    Dim i As Long, k As String
    For i = 1 To mRowCount
        If Not IsEmpty(v2(i, 1)) Then
            k = mRenderedVec(i)
            rowByKey(k) = rowByKey(k) + 1
            If VarType(v2(i, 1)) = vbDouble Then AppendVm vm, CStr(v2(i, 1)), k
        End If
    Next i

    Dim displays As New clsSet
    displays.LoadFromArray rowByKey.Keys

    Set mDisplays = displays
    Set mValueMap = vm
    Set mRowCountByKey = rowByKey
End Sub

' 값 지도에 표시 키를 vbNullChar로 이어붙인다(modColumnScan.AppendVm과 동일 규칙).
Private Sub AppendVm(ByVal vm As Object, ByVal vk As String, ByVal key As String)
    If vm.exists(vk) Then
        If InStr(1, vbNullChar & vm(vk) & vbNullChar, vbNullChar & key & vbNullChar) = 0 Then
            vm(vk) = vm(vk) & vbNullChar & key
        End If
    Else
        vm(vk) = key
    End If
End Sub
