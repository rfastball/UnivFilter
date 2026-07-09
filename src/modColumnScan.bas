Attribute VB_Name = "modColumnScan"
' @desc 열 1회 스캔해 표시키/값지도 자산 구축 담당
' @depends clsColAsset, clsSet, modLogger
Option Explicit

' 표시 키 규칙 (실측: AutoFilterTest gap-probe-findings.md P4/P6):
'  - 서식에 * 포함: TEXT(Value2, fmtLocal에서 ! 제거) 접힌형이 매칭 키
'  - 그 외: Range.Text가 매칭 키. #-런(좁은 열)은 AutoFit 재독 후 폭 복원
'  - 텍스트/에러/불리언 셀은 .Text 그대로, 빈칸은 키 없음(BlankCount)
'
' 성능 계약:
'  - 균일 서식(NumberFormatLocal 비Null): Value2로 중복 제거해 고유값당 1회만
'    렌더한다(.Text N회 -> U회). RenderBlock이 담당.
'  - 혼합 서식: 실데이터는 셀별 랜덤이 아니라 '서식 블록'이다. NumberFormatLocal이
'    혼합 범위에서 Null인 성질을 이진 분할로 재귀해 블록 경계만 찾고(O(블록수 x logN)
'    프로브), 각 블록을 RenderBlock 고속 경로로 처리한다(TryPartition).
'    병리적(셀마다 다름)이면 프로브 예산 초과 시 셀별 폴백(BuildPerCell).
'  - 값 지도 값은 clsSet 대신 vbNullChar 연결 문자열 - 값당 객체 생성 비용 회피.

Private Const PARTITION_BUDGET As Long = 2000

' @desc 열 범위 받아 표시/값 자산 반환, 빈 열이면 빈 자산 반환
Public Function BuildAsset(ByVal rngCol As Range, ByVal sID As String, _
                           ByVal lGeneration As Long) As clsColAsset
    Dim displays As New clsSet
    Dim vm As Object
    Set vm = CreateObject("Scripting.Dictionary")
    Dim rowByKey As Object
    Set rowByKey = CreateObject("Scripting.Dictionary")
    rowByKey.CompareMode = vbTextCompare

    Dim asset As New clsColAsset
    Dim rTarget As Range
    Set rTarget = ResolveTarget(rngCol)

    If rTarget Is Nothing Then
        asset.Init displays, vm, rowByKey, 0, 0, sID, lGeneration
        Set BuildAsset = asset
        Exit Function
    End If

    Dim nRows As Long
    nRows = rTarget.Rows.Count

    ' 벌크 읽기 (값/빈칸/타입 판별 - 단일 COM 호출)
    Dim v2 As Variant
    v2 = To2D(rTarget.Value2)

    ' 균일/혼합 판정도 단일 COM 호출 (혼합이면 Null)
    Dim vFmt As Variant
    vFmt = rTarget.NumberFormatLocal

    Dim blankCount As Long
    If IsNull(vFmt) Then
        ' 전체 범위가 '혼합'으로 보여도 빈칸(다른 서식) 오염일 수 있다.
        ' 데이터 셀만으로 재판정해 균일이면 분할 없이 단일 고속 경로.
        Dim dataFmt As String
        dataFmt = DataUniformFormat(rTarget)
        If Len(dataFmt) > 0 Then
            blankCount = RenderBlock(rTarget, v2, 1, nRows, dataFmt, vm, rowByKey)
        Else
            blankCount = BuildMixed(rTarget, v2, nRows, vm, rowByKey)
        End If
    Else
        blankCount = RenderBlock(rTarget, v2, 1, nRows, CStr(vFmt), vm, rowByKey)
    End If

    ' 표시 집합: 행수 지도의 키가 곧 고유 표시 키 (블록 누적 후 1회)
    displays.LoadFromArray rowByKey.Keys

    asset.Init displays, vm, rowByKey, blankCount, nRows, sID, lGeneration
    Set BuildAsset = asset
End Function

' 균일 서식 블록 [firstRow,lastRow]를 중복제거 렌더해 vm/rowByKey에 누적. 빈칸 수 반환.
Private Function RenderBlock(ByVal rTarget As Range, ByVal v2 As Variant, _
                            ByVal firstRow As Long, ByVal lastRow As Long, _
                            ByVal fmt As String, ByVal vm As Object, _
                            ByVal rowByKey As Object) As Long
    Dim isPad As Boolean
    isPad = (InStr(fmt, "*") > 0)
    Dim cleanFmt As String
    cleanFmt = Replace(fmt, "!", "")

    ' pass 1: 타입태그 + 값으로 중복 제거 (COM 무접촉, 에러 셀만 별도)
    Dim uniqFirst As Object, uniqCnt As Object
    Set uniqFirst = CreateObject("Scripting.Dictionary")
    Set uniqCnt = CreateObject("Scripting.Dictionary")
    Dim errRows As New Collection
    Dim blankCount As Long

    Dim i As Long, tag As String
    For i = firstRow To lastRow
        If IsEmpty(v2(i, 1)) Then
            blankCount = blankCount + 1
        ElseIf IsError(v2(i, 1)) Then
            errRows.Add i
        Else
            ' 타입태그로 숫자 500과 텍스트 "500"의 값키 충돌 방지
            tag = CStr(VarType(v2(i, 1))) & Chr$(1) & CStr(v2(i, 1))
            If uniqCnt.exists(tag) Then
                uniqCnt(tag) = uniqCnt(tag) + 1
            Else
                uniqFirst(tag) = i
                uniqCnt(tag) = 1
            End If
        End If
    Next i

    ' pass 2: 고유값당 1회 렌더 (첫 등장 셀 .Text 또는 * 접힌형 TEXT)
    Dim keyByTag As Object
    Set keyByTag = CreateObject("Scripting.Dictionary")
    Dim hashTags As New Collection
    Dim vt As Variant, r As Long, val As Variant, key As String
    For Each vt In uniqFirst.Keys
        r = uniqFirst(vt)
        val = v2(r, 1)
        If VarType(val) = vbString Then
            key = CStr(val)                      ' 텍스트 셀: 자기 자신, 렌더 불요
        ElseIf isPad Then
            key = SafeWText(val, cleanFmt)       ' * 패딩: 접힌형 TEXT
            If Len(key) = 0 Then key = rTarget.Cells(r, 1).Text
        Else
            key = rTarget.Cells(r, 1).Text       ' 그 외: 실셀 .Text (충실)
            If IsHashRun(key) Then hashTags.Add vt
        End If
        keyByTag(vt) = key
    Next vt

    ' ### 좁은 열: 1회 AutoFit 후 해당 고유값만 재독, 폭 복원
    If hashTags.Count > 0 Then
        Dim col As Range
        Set col = rTarget.EntireColumn
        Dim wPrev As Double
        wPrev = col.ColumnWidth
        On Error Resume Next
        col.AutoFit
        Dim h As Long
        For h = 1 To hashTags.Count
            keyByTag(hashTags(h)) = rTarget.Cells(uniqFirst(hashTags(h)), 1).Text
        Next h
        col.ColumnWidth = wPrev
        On Error GoTo 0
    End If

    ' 누적: 고유값별 개수를 표시키에 합산(서로 다른 값이 같은 표시가 될 수 있음)
    For Each vt In uniqFirst.Keys
        key = keyByTag(vt)
        rowByKey(key) = rowByKey(key) + uniqCnt(vt)
        val = v2(uniqFirst(vt), 1)
        If VarType(val) = vbDouble Then AppendVm vm, CStr(val), key
    Next vt

    ' 에러 셀(드묾): 개별 렌더
    Dim e As Long
    For e = 1 To errRows.Count
        key = rTarget.Cells(errRows(e), 1).Text
        rowByKey(key) = rowByKey(key) + 1
    Next e

    RenderBlock = blankCount
End Function

' 데이터(비빈칸) 셀만의 서식이 균일하면 그 서식 문자열을, 아니면 "" 반환.
' 빈칸이 데이터와 다른 서식이어서 전체 범위가 Null로 오판될 때 쓴다.
' SpecialCells로 상수/수식 셀을 area로 모아 단일 서식 프로브(빈칸 제외).
Private Function DataUniformFormat(ByVal rTarget As Range) As String
    Dim dc As Range, fc As Range, dataCells As Range
    On Error Resume Next
    Set dc = rTarget.SpecialCells(xlCellTypeConstants)
    Set fc = rTarget.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If dc Is Nothing And fc Is Nothing Then Exit Function   ' 데이터 없음
    If dc Is Nothing Then
        Set dataCells = fc
    ElseIf fc Is Nothing Then
        Set dataCells = dc
    Else
        Set dataCells = Application.Union(dc, fc)
    End If

    Dim f As Variant
    f = dataCells.NumberFormatLocal
    If Not IsNull(f) Then DataUniformFormat = CStr(f)
End Function

' 혼합 서식: 이진 분할로 균일 블록을 찾아 블록별 고속 렌더. 빈칸 수 반환.
Private Function BuildMixed(ByVal rTarget As Range, ByVal v2 As Variant, _
                           ByVal nRows As Long, ByVal vm As Object, _
                           ByVal rowByKey As Object) As Long
    Dim blocks As New Collection
    Dim budget As Long
    budget = PARTITION_BUDGET

    If TryPartition(rTarget, 1, nRows, blocks, budget) Then
        Dim total As Long, blk As Variant
        For Each blk In blocks
            total = total + RenderBlock(rTarget, v2, blk(0), blk(1), CStr(blk(2)), vm, rowByKey)
        Next blk
        Log "BuildMixed: 블록 " & blocks.Count & "개로 분할 (프로브 " & _
            (PARTITION_BUDGET - budget) & "회)", lvlInfo
        BuildMixed = total
    Else
        Log "BuildMixed: 과편성 - 셀별 폴백", lvlWarning
        BuildMixed = BuildPerCell(rTarget, v2, nRows, vm, rowByKey)
    End If
End Function

' [firstRow,lastRow]를 균일 블록 목록으로 이진 분할한다. 예산 초과 시 False.
Private Function TryPartition(ByVal rTarget As Range, ByVal firstRow As Long, _
                             ByVal lastRow As Long, ByVal blocks As Collection, _
                             ByRef budget As Long) As Boolean
    If budget <= 0 Then Exit Function
    budget = budget - 1

    Dim sub_ As Range
    Set sub_ = rTarget.Parent.Range(rTarget.Cells(firstRow, 1), rTarget.Cells(lastRow, 1))
    Dim vFmt As Variant
    vFmt = sub_.NumberFormatLocal

    If Not IsNull(vFmt) Then
        ' 균일(단일 셀 포함) - NumberFormatLocal은 단일 셀에서 항상 비Null
        blocks.Add Array(firstRow, lastRow, CStr(vFmt))
        TryPartition = True
        Exit Function
    End If

    ' 혼합 - 중점 분할
    Dim mid As Long
    mid = (firstRow + lastRow) \ 2
    If TryPartition(rTarget, firstRow, mid, blocks, budget) Then
        TryPartition = TryPartition(rTarget, mid + 1, lastRow, blocks, budget)
    End If
End Function

' 셀별 폴백(과편성 서식 전용): 같은 값도 셀마다 다르게 표시될 수 있어 셀별 렌더.
Private Function BuildPerCell(ByVal rTarget As Range, ByVal v2 As Variant, _
                             ByVal nRows As Long, ByVal vm As Object, _
                             ByVal rowByKey As Object) As Long
    Dim keys() As String
    ReDim keys(1 To nRows)

    Dim i As Long, needHash As Boolean
    Dim c As Range, f As String
    For i = 1 To nRows
        If IsEmpty(v2(i, 1)) Then
            keys(i) = ""
        ElseIf IsError(v2(i, 1)) Then
            keys(i) = rTarget.Cells(i, 1).Text
        Else
            Set c = rTarget.Cells(i, 1)
            f = SafeFmtLocal(c)
            If InStr(f, "*") > 0 And VarType(v2(i, 1)) = vbDouble Then
                keys(i) = SafeWText(v2(i, 1), Replace(f, "!", ""))
                If Len(keys(i)) = 0 Then keys(i) = c.Text
            Else
                keys(i) = c.Text
                If VarType(v2(i, 1)) = vbDouble And IsHashRun(keys(i)) Then needHash = True
            End If
        End If
    Next i
    If needHash Then FixHashKeys rTarget, keys, nRows

    Dim blankCount As Long, k As String
    For i = 1 To nRows
        If IsEmpty(v2(i, 1)) Then
            blankCount = blankCount + 1
        Else
            k = keys(i)
            If Len(k) > 0 Then
                rowByKey(k) = rowByKey(k) + 1
                If VarType(v2(i, 1)) = vbDouble Then AppendVm vm, CStr(v2(i, 1)), k
            End If
        End If
    Next i

    BuildPerCell = blankCount
End Function

' 값 지도에 표시 키를 vbNullChar로 이어붙인다(블록 간 중복 키 방지).
Private Sub AppendVm(ByVal vm As Object, ByVal vk As String, ByVal key As String)
    If vm.exists(vk) Then
        If InStr(1, vbNullChar & vm(vk) & vbNullChar, vbNullChar & key & vbNullChar) = 0 Then
            vm(vk) = vm(vk) & vbNullChar & key
        End If
    Else
        vm(vk) = key
    End If
End Sub

' 주입 범위를 신뢰한다: 끝쪽 빈칸도 필터 대상 행이다.
' (true last row 방어는 범위를 만드는 호출부 소관 - modShortcuts.SetupFormData)
Private Function ResolveTarget(ByVal rngCol As Range) As Range
    If rngCol Is Nothing Then Exit Function
    Set ResolveTarget = rngCol.Columns(1)
End Function

' #-런 키를 AutoFit 재독으로 복구하고 열 폭을 복원한다 (셀별 폴백용).
Private Sub FixHashKeys(ByVal rTarget As Range, ByRef keys() As String, ByVal nRows As Long)
    Dim col As Range
    Set col = rTarget.EntireColumn

    Dim wPrev As Double
    wPrev = col.ColumnWidth

    On Error GoTo Restore
    col.AutoFit

    Dim i As Long
    For i = 1 To nRows
        If IsHashRun(keys(i)) Then keys(i) = rTarget.Cells(i, 1).Text
    Next i

Restore:
    On Error Resume Next
    col.ColumnWidth = wPrev
    On Error GoTo 0
End Sub

' 문자열이 # 문자만으로 이루어져 있는지 확인한다(좁은 열 표시).
Private Function IsHashRun(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) <> "#" Then Exit Function
    Next i
    IsHashRun = True
End Function

' WorksheetFunction.Text를 오류 없이 시도한다(실패 시 빈 문자열).
Private Function SafeWText(ByVal v As Variant, ByVal sFmt As String) As String
    On Error GoTo Fail
    SafeWText = Application.WorksheetFunction.Text(v, sFmt)
    Exit Function
Fail:
    SafeWText = ""
End Function

Private Function SafeFmtLocal(ByVal c As Range) As String
    On Error GoTo Fail
    SafeFmtLocal = c.NumberFormatLocal
    Exit Function
Fail:
    SafeFmtLocal = "G/" & ChrW(&HD45C) & ChrW(&HC900)
End Function

' Value2 반환값을 항상 2차원 배열 형태로 맞춘다(단일 셀 방어).
Private Function To2D(ByVal v As Variant) As Variant
    If IsArray(v) Then
        To2D = v
    Else
        Dim tmp(1 To 1, 1 To 1) As Variant
        tmp(1, 1) = v
        To2D = tmp
    End If
End Function
