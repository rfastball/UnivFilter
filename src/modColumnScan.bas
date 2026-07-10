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
Public Function BuildAsset(ByVal rngCol As Range, ByVal sID As String) As clsColAsset
    Set BuildAsset = BuildAssetFromSig(rngCol, sID, ComputeSignature(rngCol))
End Function

' @desc 열 범위와 미리 계산된 지문 받아 자산 반환, 빈 열이면 빈 자산 반환
' 지문이 이미 읽은 Value2/서식분할을 재사용한다(렌더만 수행, 추가 Value2 COM 없음).
Public Function BuildAssetFromSig(ByVal rngCol As Range, ByVal sID As String, _
                                  ByVal sig As clsColSig) As clsColAsset
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
        asset.Init displays, vm, rowByKey, 0, 0, sID, sig
        Set BuildAssetFromSig = asset
        Exit Function
    End If

    Dim nRows As Long
    nRows = rTarget.Rows.Count

    ' 지문의 Value2 재사용 (COM 무접촉)
    Dim v2 As Variant
    v2 = sig.ContentVec

    ' 조건부서식(CF) 유무 1회 판정(O(1)): CF 있으면 숫자 셀은 .Text(CF 추종) 대신
    ' 기본서식 WText로 렌더한다 (P9: AutoFilter는 CF 아닌 기본서식 캐논형만 매칭).
    Dim hasCF As Boolean
    hasCF = HasConditionalFormat(rTarget)

    Dim blankCount As Long
    Dim blk As Variant
    If sig.IsPathological Then
        ' 과편성 서식: 셀별 폴백 (지문은 항상 재스캔 센티넬)
        blankCount = BuildPerCell(rTarget, v2, nRows, vm, rowByKey, hasCF)
    ElseIf sig.IsMixed Then
        ' 혼합: 지문이 이미 찾은 블록 경계를 그대로 순회 (재분할 없음)
        For Each blk In sig.Blocks
            blankCount = blankCount + RenderBlock(rTarget, v2, _
                CLng(blk(0)), CLng(blk(1)), CStr(blk(2)), vm, rowByKey, hasCF)
        Next blk
    Else
        blankCount = RenderBlock(rTarget, v2, 1, nRows, sig.FmtUniform, vm, rowByKey, hasCF)
    End If

    ' 표시 집합: 행수 지도의 키가 곧 고유 표시 키 (블록 누적 후 1회)
    displays.LoadFromArray rowByKey.Keys

    asset.Init displays, vm, rowByKey, blankCount, nRows, sID, sig
    Set BuildAssetFromSig = asset
End Function

' 균일 서식 블록 [firstRow,lastRow]를 중복제거 렌더해 vm/rowByKey에 누적. 빈칸 수 반환.
Private Function RenderBlock(ByVal rTarget As Range, ByVal v2 As Variant, _
                            ByVal firstRow As Long, ByVal lastRow As Long, _
                            ByVal fmt As String, ByVal vm As Object, _
                            ByVal rowByKey As Object, ByVal hasCF As Boolean) As Long
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
            key = SafeWText(val, cleanFmt)       ' * 패딩: 접힌형 TEXT (CF-안전)
            If Len(key) = 0 Then key = rTarget.Cells(r, 1).Text
        ElseIf hasCF And VarType(val) = vbDouble Then
            ' CF 열의 숫자: .Text는 CF 추종하므로 기본서식 WText로 렌더(폭 독립, AutoFit 불요)
            key = SafeWText(val, RenderFmt(fmt, cleanFmt, isGeneral))
            If Len(key) = 0 Then key = rTarget.Cells(r, 1).Text   ' 잔여: WText 빈문자열 폴백
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

' @desc 균일 서식 블록을 Evaluate(TEXT) 1회로 벌크 렌더해 행별 표시키 배열 채움, 실패 시 False 반환
' 지연 우주 재설계(docs/tasks/2026-07-10_lazy_universe_design.md) Stage A 경로.
' keys(firstRow..lastRow)에 표시키를 채운다(빈칸은 "" + blankCount 가산).
' 규칙은 RenderBlock과 동일 의미론: 텍스트 셀은 자기 자신, * 패딩은 접힌형 TEXT,
' General은 영문 "General" 매핑, ! 제거. CF 무관(TEXT는 기본서식 캐논형 - P9).
' 폭 무관(P7 캐논형)이라 AutoFit 재독 자체가 없다. 실측 91,235행 COM 1회 = 113ms.
' 에러 셀(TEXT가 에러 전파)/[색] 등 TEXT 빈문자열/불리언/@서식 숫자 셀만 셀별 .Text 폴백.
' False 반환(Evaluate 실패/수식 255자 초과) 시 호출부가 RenderBlock으로 폴백한다.
Public Function RenderBlockBulk(ByVal rTarget As Range, ByVal v2 As Variant, _
                               ByVal firstRow As Long, ByVal lastRow As Long, _
                               ByVal fmt As String, ByRef keys() As String, _
                               ByRef blankCount As Long) As Boolean
    Dim cleanFmt As String
    cleanFmt = Replace(fmt, "!", "")
    Dim fmtRender As String
    fmtRender = RenderFmt(fmt, cleanFmt, IsGeneralFmt(fmt))
    ' 텍스트 서식(@) 열의 숫자 셀은 TEXT(v,"@")가 실표시와 다를 수 있어 벌크 부적용(셀별)
    Dim isTextFmt As Boolean
    isTextFmt = (fmt = "@")

    ' 숫자 셀이 하나라도 있을 때만 Evaluate 1회 (텍스트/빈칸뿐이면 렌더 불요)
    Dim needRender As Boolean
    Dim i As Long
    If Not isTextFmt Then
        For i = firstRow To lastRow
            If VarType(v2(i, 1)) = vbDouble Then
                needRender = True
                Exit For
            End If
        Next i
    End If

    Dim rendered As Variant
    If needRender Then
        Dim blockRng As Range
        Set blockRng = rTarget.Worksheet.Range(rTarget.Cells(firstRow, 1), _
                                               rTarget.Cells(lastRow, 1))
        Dim sFormula As String
        sFormula = "TEXT(" & blockRng.Address & "," & Chr$(34) & _
                   Replace(fmtRender, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34) & ")"
        If Len(sFormula) > 255 Then Exit Function        ' Evaluate 수식 한계

        On Error Resume Next
        rendered = rTarget.Worksheet.Evaluate(sFormula)
        On Error GoTo 0
        If Not IsArray(rendered) Then
            If firstRow = lastRow Then
                If IsError(rendered) Or IsEmpty(rendered) Then Exit Function
                Dim tmp1 As Variant
                ReDim tmp1(1 To 1, 1 To 1)
                tmp1(1, 1) = rendered                    ' 단일 행 블록: 스칼라 반환 정규화
                rendered = tmp1
            Else
                Exit Function                            ' 다행인데 비배열 = Evaluate 실패
            End If
        End If
    End If

    Dim val As Variant, k As String, j As Long
    For i = firstRow To lastRow
        val = v2(i, 1)
        If IsEmpty(val) Then
            keys(i) = ""                                 ' TEXT(빈칸)="0" 함정: 렌더값 무시
            blankCount = blankCount + 1
        ElseIf IsError(val) Then
            keys(i) = rTarget.Cells(i, 1).Text           ' 에러: TEXT가 에러 전파 - 셀별
        ElseIf VarType(val) = vbString Then
            keys(i) = CStr(val)                          ' 텍스트: 자기 자신 (TEXT는 날짜문자열 강제 위험)
        ElseIf VarType(val) = vbDouble And needRender Then
            j = i - firstRow + 1
            If IsError(rendered(j, 1)) Then
                keys(i) = rTarget.Cells(i, 1).Text
            Else
                k = CStr(rendered(j, 1))
                If Len(k) = 0 Then k = rTarget.Cells(i, 1).Text   ' [색] 등 TEXT 빈문자열 폴백
                keys(i) = k
            End If
        Else
            keys(i) = rTarget.Cells(i, 1).Text           ' 불리언/@서식 숫자 등: 셀별 (드묾)
        End If
    Next i

    RenderBlockBulk = True
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
                             ByVal rowByKey As Object, ByVal hasCF As Boolean) As Long
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
            ElseIf hasCF And VarType(v2(i, 1)) = vbDouble Then
                ' CF 열의 숫자: 기본서식 WText로 렌더(.Text는 CF 추종, 폭 독립)
                keys(i) = SafeWText(v2(i, 1), RenderFmt(f, Replace(f, "!", ""), IsGeneralFmt(f)))
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

' 범위에 조건부서식(CF)이 걸려 있는지 1회 판정한다(있으면 숫자 렌더를 기본서식 WText로 전환).
Private Function HasConditionalFormat(ByVal rTarget As Range) As Boolean
    On Error Resume Next
    HasConditionalFormat = (rTarget.FormatConditions.Count > 0)
    On Error GoTo 0
End Function

' CF 셀 WText 렌더용 서식 결정: General은 로컬 토큰(G/표준)을 TEXT가 미해석하므로 영문 "General",
' 그 외는 ! 제거한 서식 (P9: 공통 숫자서식은 로컬==영문이라 WText가 오라클을 재현).
Private Function RenderFmt(ByVal localFmt As String, ByVal cleanFmt As String, _
                           ByVal isGeneral As Boolean) As String
    If isGeneral Then RenderFmt = "General" Else RenderFmt = cleanFmt
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
