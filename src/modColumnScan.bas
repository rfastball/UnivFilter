Attribute VB_Name = "modColumnScan"
' @desc 열 1회 스캔해 표시키/값지도 자산 구축 담당
' @depends clsColAsset, clsColSig, clsSet, modLogger
Option Explicit

' 표시 키 규칙 (실측: AutoFilterTest gap-probe-findings.md P4/P6/P7):
'  - 서식에 * 포함: TEXT(Value2, fmtLocal에서 ! 제거) 접힌형이 매칭 키
'  - 그 외: Range.Text가 매칭 키. 단 폭에 잘린 형은 AutoFit 재독 후 폭 복원:
'    #-런(고정서식은 #####로 잘림) + General 전체(과학표기/반올림으로 잘림, # 없음, P7)
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

' @desc 열 범위 받아 지문(Value2/서식 관측) 반환, 빈 열이면 균일 빈 지문
' 스캔의 값싼 prefix: Value2 벌크 + NumberFormatLocal 1회(+혼합 시 TryPartition). 렌더/AutoFit 없음.
Public Function ComputeSignature(ByVal rngCol As Range) As clsColSig
    Dim sig As New clsColSig
    Dim rTarget As Range
    Set rTarget = ResolveTarget(rngCol)

    If rTarget Is Nothing Then
        sig.Init 0, Empty, False, "", Nothing, False
        Set ComputeSignature = sig
        Exit Function
    End If

    Dim nRows As Long
    nRows = rTarget.Rows.Count

    Dim v2 As Variant
    v2 = To2D(rTarget.Value2)

    Dim vFmt As Variant
    vFmt = rTarget.NumberFormatLocal

    If Not IsNull(vFmt) Then
        sig.Init nRows, v2, False, CStr(vFmt), Nothing, False
    Else
        Dim blocks As New Collection
        Dim budget As Long
        budget = PARTITION_BUDGET
        If TryPartition(rTarget, 1, nRows, blocks, budget) Then
            sig.Init nRows, v2, True, "", blocks, False
        Else
            sig.Init nRows, v2, False, "", Nothing, True
        End If
    End If

    Set ComputeSignature = sig
End Function

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
        blankCount = BuildMixed(rTarget, v2, nRows, vm, rowByKey)
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
    ' General은 좁으면 과학표기/반올림으로 잘린다(#-런 아님). 캐논형 확보 위해 AutoFit 재독 강제.
    Dim isGeneral As Boolean
    isGeneral = IsGeneralFmt(fmt)
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
    Dim refitTags As New Collection
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
            ' 폭에 잘린 형 방어: #-런(고정서식) + General 전체(과학표기/반올림)
            If isGeneral Or IsHashRun(key) Then refitTags.Add vt
        End If
        keyByTag(vt) = key
    Next vt

    ' 폭에 잘린 고유값: 1회 AutoFit 후 해당 값만 캐논형으로 재독, 폭 복원
    If refitTags.Count > 0 Then
        Dim col As Range
        Set col = rTarget.EntireColumn
        Dim wPrev As Double
        wPrev = col.ColumnWidth
        On Error Resume Next
        col.AutoFit
        Dim h As Long
        For h = 1 To refitTags.Count
            keyByTag(refitTags(h)) = rTarget.Cells(uniqFirst(refitTags(h)), 1).Text
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

    Dim refit() As Boolean
    ReDim refit(1 To nRows)
    Dim i As Long, needRefit As Boolean
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
                ' 폭에 잘린 형 방어: General(과학표기/반올림) + #-런(고정서식)
                If VarType(v2(i, 1)) = vbDouble Then
                    If IsGeneralFmt(f) Or IsHashRun(keys(i)) Then
                        refit(i) = True
                        needRefit = True
                    End If
                End If
            End If
        End If
    Next i
    If needRefit Then FixHashKeys rTarget, keys, refit, nRows

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

' 폭에 잘린 키를 AutoFit 재독으로 캐논형 복구하고 열 폭을 복원한다 (셀별 폴백용).
Private Sub FixHashKeys(ByVal rTarget As Range, ByRef keys() As String, _
                       ByRef refit() As Boolean, ByVal nRows As Long)
    Dim col As Range
    Set col = rTarget.EntireColumn

    Dim wPrev As Double
    wPrev = col.ColumnWidth

    On Error GoTo Restore
    col.AutoFit

    Dim i As Long
    For i = 1 To nRows
        If refit(i) Then keys(i) = rTarget.Cells(i, 1).Text
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

' General 서식(로컬 "G/표준" 또는 "General")인지 확인한다. General은 폭에 따라
' 과학표기/반올림으로 잘려 #-런이 아닌 잘못된 키를 낼 수 있어 AutoFit 재독이 필요하다.
Private Function IsGeneralFmt(ByVal fmt As String) As Boolean
    IsGeneralFmt = (fmt = "General") Or (fmt = "G/" & ChrW(&HD45C) & ChrW(&HC900))
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
