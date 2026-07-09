VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColSig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 열 지문(Value2/서식 관측) 저장·신선도대조 담당
Option Explicit

' 지문 = 스캔의 값싼 prefix. 렌더 없이 우주 신선도를 관측으로 확정한다.
'  - ContentVec: 취득 시 Value2 2D 배열 (빈 열이면 Empty). 원소별 대조로 충돌 0.
'  - 서식: 균일이면 NumberFormatLocal 문자열, 혼합이면 블록목록 [(start,end,fmt)] 구조 대조.
'  - 병리적(분할예산 초과): 항상 불일치(재스캔 센티넬).
Private mRowCount As Long
Private mContentVec As Variant
Private mIsMixed As Boolean
Private mIsPathological As Boolean
Private mFmtUniform As String
Private mBlocks As Collection

' @desc 관측값 주입받아 지문을 초기화함
Public Sub Init(ByVal lRowCount As Long, ByRef vContent As Variant, _
                ByVal bMixed As Boolean, ByVal sFmtUniform As String, _
                ByVal oBlocks As Collection, ByVal bPathological As Boolean)
    mRowCount = lRowCount
    mContentVec = vContent
    mIsMixed = bMixed
    mFmtUniform = sFmtUniform
    Set mBlocks = oBlocks
    mIsPathological = bPathological
End Sub

Public Property Get RowCount() As Long
    RowCount = mRowCount
End Property

Public Property Get ContentVec() As Variant
    ContentVec = mContentVec
End Property

Public Property Get IsMixed() As Boolean
    IsMixed = mIsMixed
End Property

Public Property Get IsPathological() As Boolean
    IsPathological = mIsPathological
End Property

Public Property Get FmtUniform() As String
    FmtUniform = mFmtUniform
End Property

Public Property Get Blocks() As Collection
    Set Blocks = mBlocks
End Property

' @desc 다른 지문 받아 우주 동일 여부 반환, 병리적/불일치 시 False
Public Function Matches(ByVal other As clsColSig) As Boolean
    If other Is Nothing Then Exit Function
    ' 병리적: 항상 재스캔 (자신 또는 상대)
    If mIsPathological Then Exit Function
    If other.IsPathological Then Exit Function
    If mRowCount <> other.RowCount Then Exit Function
    If mIsMixed <> other.IsMixed Then Exit Function
    If mIsMixed Then
        If Not BlocksEqual(mBlocks, other.Blocks) Then Exit Function
    Else
        If mFmtUniform <> other.FmtUniform Then Exit Function
    End If
    Dim vOther As Variant
    vOther = other.ContentVec
    If Not VecEqual(mContentVec, vOther) Then Exit Function
    Matches = True
End Function

' 두 Value2 배열(또는 빈 열 Empty)의 원소별 동일 여부. 순서민감(스왑=불필요재스캔, 무해).
Private Function VecEqual(ByRef a As Variant, ByRef b As Variant) As Boolean
    Dim ae As Boolean, be As Boolean
    ae = IsEmpty(a)
    be = IsEmpty(b)
    If ae Then
        VecEqual = be
        Exit Function
    End If
    If be Then Exit Function
    If Not IsArray(a) Then Exit Function
    If Not IsArray(b) Then Exit Function
    If LBound(a, 1) <> LBound(b, 1) Then Exit Function
    If UBound(a, 1) <> UBound(b, 1) Then Exit Function
    Dim i As Long
    For i = LBound(a, 1) To UBound(a, 1)
        If Not VarEq(a(i, 1), b(i, 1)) Then Exit Function
    Next i
    VecEqual = True
End Function

' 셀 값 1쌍 대조: VarType 우선(숫자5 != 텍스트"5"), 에러셀은 강제 재스캔(안전측).
Private Function VarEq(ByRef a As Variant, ByRef b As Variant) As Boolean
    If IsEmpty(a) Then
        VarEq = IsEmpty(b)
        Exit Function
    End If
    If IsEmpty(b) Then Exit Function
    ' 에러 Variant는 = 비교 시 예외 -> 무조건 불일치로 두어 재스캔(정합성 안전)
    If IsError(a) Then Exit Function
    If IsError(b) Then Exit Function
    If VarType(a) <> VarType(b) Then Exit Function
    VarEq = (a = b)
End Function

' 두 블록목록(Array(start,end,fmt))의 구조 동일 여부.
Private Function BlocksEqual(ByVal a As Collection, ByVal b As Collection) As Boolean
    If a Is Nothing Then
        BlocksEqual = (b Is Nothing)
        Exit Function
    End If
    If b Is Nothing Then Exit Function
    If a.Count <> b.Count Then Exit Function
    Dim i As Long, ba As Variant, bb As Variant
    For i = 1 To a.Count
        ba = a(i)
        bb = b(i)
        If CLng(ba(0)) <> CLng(bb(0)) Then Exit Function
        If CLng(ba(1)) <> CLng(bb(1)) Then Exit Function
        If CStr(ba(2)) <> CStr(bb(2)) Then Exit Function
    Next i
    BlocksEqual = True
End Function
