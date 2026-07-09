Attribute VB_Name = "Test_Signature"
' @desc clsColSig 대조 계약 + ComputeSignature 관측 단위테스트 담당
' @depends modColumnScan, clsColSig, LogShim
Option Explicit

Private Const SHEET_NAME As String = "UF_SigTest"

' ==========================================================================
' clsColSig.Matches 단위 (시트 무관 - 직접 구성)
' ==========================================================================

' @desc 동일 내용/서식 지문이 매칭되는지 검증함
Public Sub Test_Sig_MatchesSame()
    Dim s1 As clsColSig, s2 As clsColSig
    Set s1 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(15)), "General")
    Set s2 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(15)), "General")
    LogShim.AssertEq "sig_same", True, s1.Matches(s2)
End Sub

' @desc 값 1개가 다르면 불일치하는지 검증함
Public Sub Test_Sig_ValueDiff()
    Dim s1 As clsColSig, s2 As clsColSig
    Set s1 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(15)), "General")
    Set s2 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(20)), "General")
    LogShim.AssertEq "sig_valdiff", False, s1.Matches(s2)
End Sub

' @desc 행수가 다르면 불일치하는지 검증함
Public Sub Test_Sig_RowCountDiff()
    Dim s1 As clsColSig, s2 As clsColSig
    Set s1 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(15)), "General")
    Set s2 = UniformSig(Vec2(CDbl(5), CDbl(10)), "General")
    LogShim.AssertEq "sig_rowdiff", False, s1.Matches(s2)
End Sub

' @desc 균일 vs 혼합(kind) 전이가 불일치하는지 검증함
Public Sub Test_Sig_KindDiff()
    Dim s1 As clsColSig, s2 As clsColSig
    Set s1 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(15)), "General")
    Set s2 = New clsColSig
    s2.Init 3, Vec3(CDbl(5), CDbl(10), CDbl(15)), True, "", Blocks2("#,##0", "0"), False
    LogShim.AssertEq "sig_kinddiff", False, s1.Matches(s2)
End Sub

' @desc 타입태그(숫자5 != 텍스트"5")로 불일치하는지 검증함
Public Sub Test_Sig_TypeTag()
    Dim s1 As clsColSig, s2 As clsColSig
    Set s1 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(15)), "General")
    Set s2 = UniformSig(Vec3("5", CDbl(10), CDbl(15)), "General")
    LogShim.AssertEq "sig_typetag", False, s1.Matches(s2)
End Sub

' @desc 병리적 지문은 동일 내용이어도 항상 불일치(재스캔)하는지 검증함
Public Sub Test_Sig_Pathological()
    Dim s1 As clsColSig, s2 As clsColSig
    Set s1 = New clsColSig
    s1.Init 3, Vec3(CDbl(5), CDbl(10), CDbl(15)), False, "", Nothing, True
    Set s2 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(15)), "")
    LogShim.AssertEq "sig_patho_self", False, s1.Matches(s2)
    LogShim.AssertEq "sig_patho_other", False, s2.Matches(s1)
End Sub

' @desc 내용 동일·서식만 다르면 불일치하는지 검증함(서식편집 감지 핵심)
Public Sub Test_Sig_FmtDiff()
    Dim s1 As clsColSig, s2 As clsColSig
    Set s1 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(15)), "General")
    Set s2 = UniformSig(Vec3(CDbl(5), CDbl(10), CDbl(15)), "#,##0")
    LogShim.AssertEq "sig_fmtdiff", False, s1.Matches(s2)
End Sub

' @desc 혼합 블록목록 동일/상이를 대조하는지 검증함
Public Sub Test_Sig_Blocks()
    Dim s1 As clsColSig, s2 As clsColSig, s3 As clsColSig
    Set s1 = New clsColSig
    s1.Init 3, Vec3(CDbl(5), CDbl(10), CDbl(15)), True, "", Blocks2("#,##0", "0"), False
    Set s2 = New clsColSig
    s2.Init 3, Vec3(CDbl(5), CDbl(10), CDbl(15)), True, "", Blocks2("#,##0", "0"), False
    Set s3 = New clsColSig
    s3.Init 3, Vec3(CDbl(5), CDbl(10), CDbl(15)), True, "", Blocks2("#,##0", "0.00"), False
    LogShim.AssertEq "sig_blocks_same", True, s1.Matches(s2)
    LogShim.AssertEq "sig_blocks_diff", False, s1.Matches(s3)
End Sub

' ==========================================================================
' ComputeSignature 관측 (실 시트)
' ==========================================================================

' @desc 균일 열 지문 계산과 무변경 재계산 안정성을 검증함
Public Sub Test_Sig_ComputeUniform()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1:A3").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000
    ws.Range("A2").Value = 2000
    ws.Range("A3").Value = 1000

    Dim s1 As clsColSig, s2 As clsColSig
    Set s1 = modColumnScan.ComputeSignature(ws.Range("A1:A3"))
    Set s2 = modColumnScan.ComputeSignature(ws.Range("A1:A3"))
    LogShim.AssertEq "sig_compute_uniform_kind", False, s1.IsMixed
    LogShim.AssertEq "sig_compute_self", True, s1.Matches(s1)
    LogShim.AssertEq "sig_compute_stable", True, s1.Matches(s2)

    DropSheet
End Sub

' @desc 값 편집을 이벤트 없이 지문이 감지하는지 검증함
Public Sub Test_Sig_DetectsValueEdit()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1:A3").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000
    ws.Range("A2").Value = 2000
    ws.Range("A3").Value = 1000

    Dim s1 As clsColSig
    Set s1 = modColumnScan.ComputeSignature(ws.Range("A1:A3"))
    ws.Range("A2").Value = 9999
    Dim s2 As clsColSig
    Set s2 = modColumnScan.ComputeSignature(ws.Range("A1:A3"))
    LogShim.AssertEq "sig_detect_value", False, s1.Matches(s2)

    DropSheet
End Sub

' @desc 서식만 바꿔도(이벤트 미발화) 지문이 감지하는지 검증함(결함#3)
Public Sub Test_Sig_DetectsFormatEdit()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1:A3").NumberFormat = "#,##0"
    ws.Range("A1").Value = 1000
    ws.Range("A2").Value = 2000
    ws.Range("A3").Value = 1000

    Dim s1 As clsColSig
    Set s1 = modColumnScan.ComputeSignature(ws.Range("A1:A3"))
    ws.Range("A1:A3").NumberFormat = "0.00"      ' 값 불변, 서식만 변경
    Dim s2 As clsColSig
    Set s2 = modColumnScan.ComputeSignature(ws.Range("A1:A3"))
    LogShim.AssertEq "sig_detect_format", False, s1.Matches(s2)

    DropSheet
End Sub

' @desc 혼합 열 지문 계산과 블록 서식편집 감지를 검증함
Public Sub Test_Sig_ComputeMixed()
    Dim ws As Worksheet
    Set ws = PrepSheet()
    ws.Range("A1").NumberFormat = "0.00"
    ws.Range("A1").Value = 0.5
    ws.Range("A2").NumberFormat = "0%"
    ws.Range("A2").Value = 0.5

    Dim s1 As clsColSig, s2 As clsColSig
    Set s1 = modColumnScan.ComputeSignature(ws.Range("A1:A2"))
    Set s2 = modColumnScan.ComputeSignature(ws.Range("A1:A2"))
    LogShim.AssertEq "sig_mixed_kind", True, s1.IsMixed
    LogShim.AssertEq "sig_mixed_stable", True, s1.Matches(s2)

    ws.Range("A2").NumberFormat = "0.0%"         ' 한 블록 서식만 변경
    Dim s3 As clsColSig
    Set s3 = modColumnScan.ComputeSignature(ws.Range("A1:A2"))
    LogShim.AssertEq "sig_mixed_blockedit", False, s1.Matches(s3)

    DropSheet
End Sub

' ==========================================================================
' 헬퍼
' ==========================================================================

' 균일 서식 지문을 직접 구성한다.
Private Function UniformSig(ByRef vContent As Variant, ByVal sFmt As String) As clsColSig
    Dim s As New clsColSig
    s.Init RowsOf(vContent), vContent, False, sFmt, Nothing, False
    Set UniformSig = s
End Function

Private Function RowsOf(ByRef v As Variant) As Long
    If IsArray(v) Then RowsOf = UBound(v, 1) - LBound(v, 1) + 1
End Function

Private Function Vec2(a As Variant, b As Variant) As Variant
    Dim v(1 To 2, 1 To 1) As Variant
    v(1, 1) = a
    v(2, 1) = b
    Vec2 = v
End Function

Private Function Vec3(a As Variant, b As Variant, c As Variant) As Variant
    Dim v(1 To 3, 1 To 1) As Variant
    v(1, 1) = a
    v(2, 1) = b
    v(3, 1) = c
    Vec3 = v
End Function

' 블록 2개(1..2 / 3..3) 목록을 구성한다.
Private Function Blocks2(ByVal f1 As String, ByVal f2 As String) As Collection
    Dim c As New Collection
    c.Add Array(1&, 2&, f1)
    c.Add Array(3&, 3&, f2)
    Set Blocks2 = c
End Function

' 스크래치 시트를 만들고 초기화한다.
Private Function PrepSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_NAME
    End If
    ws.Cells.Clear
    ws.Columns("A:C").ColumnWidth = 12
    Set PrepSheet = ws
End Function

' 스크래치 시트를 제거한다(떠돌이 문서 모듈 방지).
Private Sub DropSheet()
    Dim prev As Boolean
    prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = prev
End Sub
