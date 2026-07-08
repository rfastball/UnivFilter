VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColumnAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 열 데이터/서식/타입/세대 자산 보관 담당
' @depends clsSet
Option Explicit

' ========================================================
' Class: clsColumnAsset
' 설명: 열(Column) 하나의 자산 단위 (데이터 + 서식)
' ========================================================

Private mSetData As clsSet      ' 최적화된 집합 엔진
Private mFormat As String       ' 엑셀 표시 형식 (예: "#,##0")
Private mID As String           ' 캐시 키 (디버깅용)
Private mDetectedType As String ' [New] 열의 판별된 타입 (NUMBER, STRING, HYBRID 등)
Private mGeneration As Long     ' [세대 계약] 취득 시점의 시트 세대 스탬프


' ========================================================
' [Method] 데이터 주입 (Dependency Injection)
' 설명: modCache에서 만들어진 데이터를 받아 보관만 함
' ========================================================
' ========================================================
' [Method] 데이터 주입 (Init 수정)
' 설명: 기존 인자에 sType을 추가하여 초기화 시점에 타입 정보를 받음
' ========================================================
Public Sub Init(ByVal builtSet As clsSet, _
                ByVal sFormat As String, _
                ByVal sID As String, _
                ByVal sType As String, _
                Optional ByVal lGeneration As Long = 0) ' [세대 계약] 취득 세대
    Set mSetData = builtSet
    mFormat = sFormat
    mID = sID
    mDetectedType = sType ' [New] 타입 저장
    mGeneration = lGeneration
End Sub

' ========================================================
' [Property] 읽기 전용 속성
' ========================================================
Public Property Get data() As clsSet
    Set data = mSetData
End Property
Public Property Get DetectedType() As String
    DetectedType = mDetectedType
End Property
Public Property Get fmt() As String
    fmt = mFormat
End Property

Public Property Get ID() As String
    ID = mID
End Property

' [세대 계약] 취득 시점 시트 세대 (modCache.IsAssetStale이 판정에 사용)
Public Property Get Generation() As Long
    Generation = mGeneration
End Property

' [추가] 외부에서 판별된 타입을 읽어갈 수 있도록 통로 개설
Public Property Get ColumnType() As String
    ColumnType = mDetectedType
End Property

' [추가] 표시 형식도 재활용할 가능성이 높으니 함께 추가하면 좋습니다
Public Property Get ColumnFormat() As String
    ColumnFormat = mFormat
End Property
