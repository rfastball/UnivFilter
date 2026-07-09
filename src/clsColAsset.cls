VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 열 표시키 집합/값 지도/세대 보관 담당
Option Explicit

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
Private mGeneration As Long

' @desc 스캔 산출물 주입받아 자산을 초기화함
Public Sub Init(ByVal displaySet As clsSet, ByVal valueMap As Object, _
                ByVal rowCountByKey As Object, ByVal lBlankCount As Long, _
                ByVal lRowCount As Long, ByVal sID As String, ByVal lGeneration As Long)
    Set mDisplays = displaySet
    Set mValueMap = valueMap
    Set mRowCountByKey = rowCountByKey
    mBlankCount = lBlankCount
    mRowCount = lRowCount
    mID = sID
    mGeneration = lGeneration
End Sub

Public Property Get Displays() As clsSet
    Set Displays = mDisplays
End Property

Public Property Get ValueMap() As Object
    Set ValueMap = mValueMap
End Property

Public Property Get RowCountByKey() As Object
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

Public Property Get Generation() As Long
    Generation = mGeneration
End Property

' @desc 표시 키 집합 받아 기대 가시 행수 반환, 미지 키는 0으로 집계
Public Function ExpectedRows(ByVal keySet As clsSet, ByVal bWithBlanks As Boolean) As Long
    Dim n As Long, vKey As Variant
    For Each vKey In keySet.KeysArray
        If mRowCountByKey.exists(vKey) Then n = n + mRowCountByKey(vKey)
    Next vKey
    If bWithBlanks Then n = n + mBlankCount
    ExpectedRows = n
End Function
