VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColSig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
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
    If Not FormatMatches(other) Then Exit Function
    Dim vOther As Variant
    vOther = other.ContentVec
    If Not VecEqual(mContentVec, vOther) Then Exit Function
    Matches = True
End Function

' @desc 다른 지문과 행수/서식 구조 동일 여부 반환, 병리적이면 False
' Matches의 값 대조 제외판 - 증분 패치 적격성(값만 바뀐 경우) 판정에 사용.
Public Function FormatMatches(ByVal other As clsColSig) As Boolean
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
    FormatMatches = True
End Function

' @desc 다른 지문과 값이 다른 행 번호 컬렉션 반환, 전제 불일치나 상한 초과 시 Nothing
' 전제 = FormatMatches(행수/서식 구조 동일). 증분 재렌더(modColumnScan.TryPatchAsset)
' 전용 - 키는 (값, 서식)의 순수 함수(P11/P12)라 무변경 행의 키 재사용이 건전하다.
Public Function DiffRows(ByVal other As clsColSig, ByVal maxRows As Long) As Collection
    If Not FormatMatches(other) Then Exit Function
    Dim vO As Variant
    vO = other.ContentVec

    Dim result As New Collection
    If IsEmpty(mContentVec) Then
        If IsEmpty(vO) Then Set DiffRows = result
        Exit Function
    End If
    If IsEmpty(vO) Then Exit Function
    If Not IsArray(mContentVec) Then Exit Function
    If Not IsArray(vO) Then Exit Function
    If LBound(mContentVec, 1) <> LBound(vO, 1) Then Exit Function
    If UBound(mContentVec, 1) <> UBound(vO, 1) Then Exit Function

    Dim i As Long
    For i = LBound(mContentVec, 1) To UBound(mContentVec, 1)
        If Not VarEq(mContentVec(i, 1), vO(i, 1)) Then
            result.Add i
            If result.Count > maxRows Then Exit Function   ' 상한 초과: 전량 재스캔이 유리
        End If
    Next i
    Set DiffRows = result
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

' 셀 값 1쌍 대조: VarType 우선(숫자5 != 텍스트"5"), 에러셀은 에러코드끼리 대조.
Private Function VarEq(ByRef a As Variant, ByRef b As Variant) As Boolean
    If IsEmpty(a) Then
        VarEq = IsEmpty(b)
        Exit Function
    End If
    If IsEmpty(b) Then Exit Function
    ' 에러 Variant는 = 비교 시 예외 -> 에러코드 대조로 우회. 구현이 "무조건 불일치"였을 때
    ' 에러 상존 실데이터에서 캐시가 전멸했다(2026-07-10 지출 fixture 실측: 영구 재스캔).
    If IsError(a) Then
        If IsError(b) Then VarEq = ErrEq(a, b)
        Exit Function
    End If
    If IsError(b) Then Exit Function
    If VarType(a) <> VarType(b) Then Exit Function
    VarEq = (a = b)
End Function

' 에러 변이체 1쌍의 에러코드 동일 여부(CStr "Error NNNN" 렌더 대조, 실패 시 안전측 False).
Private Function ErrEq(ByRef a As Variant, ByRef b As Variant) As Boolean
    On Error GoTo Fail
    ErrEq = (CStr(a) = CStr(b))
    Exit Function
Fail:
    ErrEq = False
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
