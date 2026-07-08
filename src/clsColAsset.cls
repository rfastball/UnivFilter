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
' 값 지도: CStr(Value2) -> 그 값을 렌더링하는 표시 키 clsSet (숫자/날짜 셀만)
Private mValueMap As Object
Private mHasBlanks As Boolean
Private mRowCount As Long
Private mID As String
Private mGeneration As Long

' @desc 스캔 산출물 주입받아 자산을 초기화함
Public Sub Init(ByVal displaySet As clsSet, ByVal valueMap As Object, _
                ByVal bHasBlanks As Boolean, ByVal lRowCount As Long, _
                ByVal sID As String, ByVal lGeneration As Long)
    Set mDisplays = displaySet
    Set mValueMap = valueMap
    mHasBlanks = bHasBlanks
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

Public Property Get HasBlanks() As Boolean
    HasBlanks = mHasBlanks
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
