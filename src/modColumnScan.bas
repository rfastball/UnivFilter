Attribute VB_Name = "modColumnScan"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc 열 1회 스캔해 표시키/값지도 자산 구축 담당
' @depends clsColAsset, clsColSig, clsSet, modLogger
Option Explicit

' 표시 키 규칙 (실측: AutoFilterTest gap-probe-findings.md P4/P6/P7/P11/P12):
'  - 매칭 키는 폭/CF/화면 무관 캐논형 - (값, 서식)의 순수 함수 (P7/P11/P12)
'  - 숫자 셀: TEXT(Value2, 렌더서식). * 패딩은 ! 제거 접힌형(P6), General은 영문
'    "General" 매핑, @서식 숫자 셀은 General 규칙(P11)
'  - 텍스트 셀: 자기 자신(CStr). 불리언: "TRUE"/"FALSE"(P12). 에러: 코드 매핑
'    (고전 7종 영문 불변 - 신형/미지 코드는 셀별 .Text 폴백). 빈칸: 키 없음(BlankCount)
'  - 셀별 .Text는 폴백에서만: 화면 프록시라 폭 잘림 + 행숨김 이력 저하(행위치 비례,
'    ProbeTextDecay 실측)에 노출되므로 캐논 경로가 항상 우선한다
'
' 성능 계약 (지연 우주 - docs/tasks/2026-07-10_lazy_universe_design.md):
'  - 균일/혼합 블록: 블록별 Application.Text(배열) 1회 벌크 렌더(RenderBlockBulk)로
'    행별 표시키 배열(RenderedVec)을 생산한다. 실측 91,235행 = COM 1회 ~100ms.
'    집합/지도 인덱스는 자산(clsColAsset)이 소비자 최초 접근 시 지연 구축.
'  - 혼합 서식: 실데이터는 셀별 랜덤이 아니라 '서식 블록'이다. NumberFormatLocal이
'    혼합 범위에서 Null인 성질을 이진 분할로 재귀해 블록 경계만 찾는다(TryPartition,
'    O(블록수 x logN) 프로브 - 지문 계산이 수행).
'  - 병리적(셀마다 다른 서식, 예산 초과) 또는 벌크 실패(희소): 셀별 렌더 폴백
'    (BuildPerCell)이 즉시 인덱스 자산을 구축한다.
'  - 값 지도 값은 clsSet 대신 vbNullChar 연결 문자열 - 값당 객체 생성 비용 회피.

Private Const PARTITION_BUDGET As Long = 2000

' 테스트 훅: True면 벌크 렌더를 건너뛰고 셀별 렌더(BuildPerCell) 오라클을 강제한다.
' (Test_CS_BulkRenderEquiv가 벌크와 셀별 오라클을 대조하는 용도 - 실사용 코드 사용 금지)
Public G_ForceLegacyRender As Boolean

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

    Dim blankCount As Long
    Dim blk As Variant

    ' 기본 경로(지연 우주 Stage B): 블록별 벌크 Evaluate(TEXT)로 행별 렌더키 배열을
    ' 생산하고 집합/지도는 자산이 지연 구축한다. 벌크 실패(희소) 시 레거시로 폴백.
    If (Not sig.IsPathological) And (Not G_ForceLegacyRender) Then
        Dim keys() As String
        ReDim keys(1 To nRows)
        Dim okBulk As Boolean
        If sig.IsMixed Then
            okBulk = True
            For Each blk In sig.Blocks
                If Not RenderBlockBulk(rTarget, v2, CLng(blk(0)), CLng(blk(1)), _
                                       CStr(blk(2)), keys, blankCount) Then
                    okBulk = False
                    Exit For
                End If
            Next blk
        Else
            okBulk = RenderBlockBulk(rTarget, v2, 1, nRows, sig.FmtUniform, keys, blankCount)
        End If
        If okBulk Then
            asset.InitRendered keys, blankCount, nRows, sID, sig
            Set BuildAssetFromSig = asset
            Exit Function
        End If
        blankCount = 0                       ' 폴백: 레거시 경로가 다시 계수
    End If

    ' 셀별 렌더 경로: 병리적 서식 / 벌크 실패 / 테스트 훅.
    ' 키 규칙은 벌크와 동일한 캐논(fmtLocal + WText) - CF 분기 불요(WText는 CF 무관).
    blankCount = BuildPerCell(rTarget, v2, nRows, vm, rowByKey)

    ' 표시 집합: 행수 지도의 키가 곧 고유 표시 키
    displays.LoadFromArray rowByKey.Keys

    asset.Init displays, vm, rowByKey, blankCount, nRows, sID, sig
    Set BuildAssetFromSig = asset
End Function

' @desc 균일 서식 블록을 Application.Text 1회로 벌크 렌더해 행별 표시키 배열 채움, 실패 시 False 반환
' 지연 우주(docs/tasks/2026-07-10_lazy_universe_design.md)의 기본 렌더 경로.
' keys(firstRow..lastRow)에 표시키를 채운다(빈칸은 "" + blankCount 가산).
' 표시 키 규칙(모듈 헤더)의 벌크판: 텍스트 셀은 자기 자신, * 패딩은 접힌형 TEXT,
' General은 영문 "General" 매핑, ! 제거, @서식 숫자 셀은 General 렌더(P11 캐논).
' CF 무관(TEXT는 기본서식 캐논형 - P9). 폭 무관(P7/P11 캐논형)이라 AutoFit 재독 없음.
' 실측 91,235행 COM 1회 = ~100ms. 에러 셀은 코드 매핑(ErrKey, P12) - 미지 코드/
' TEXT 빈문자열([색])만 셀별 .Text 폴백. 불리언은 "TRUE"/"FALSE"(P12).
' False 반환(리프팅 실패) 시 호출부가 셀별 렌더(BuildPerCell)로 폴백.
' 정합 게이트: Test_CS_BulkRenderEquiv가 셀별 오라클과 전수 동치를 고정.
Public Function RenderBlockBulk(ByVal rTarget As Range, ByVal v2 As Variant, _
                               ByVal firstRow As Long, ByVal lastRow As Long, _
                               ByVal fmt As String, ByRef keys() As String, _
                               ByRef blankCount As Long) As Boolean
    Dim cleanFmt As String
    cleanFmt = Replace(fmt, "!", "")
    ' @ 블록의 숫자 셀은 General 규칙으로 표시된다(P11) - General로 벌크 렌더
    Dim fmtRender As String
    If fmt = "@" Then
        fmtRender = "General"
    Else
        fmtRender = RenderFmt(fmt, cleanFmt, IsGeneralFmt(fmt))
    End If

    ' 숫자 셀이 하나라도 있을 때만 벌크 렌더 1회 (텍스트/빈칸뿐이면 렌더 불요)
    Dim needRender As Boolean
    Dim i As Long
    For i = firstRow To lastRow
        If VarType(v2(i, 1)) = vbDouble Then
            needRender = True
            Exit For
        End If
    Next i

    Dim rendered As Variant
    If needRender Then
        ' Application.Text(배열, 서식) 리프팅 - Evaluate(TEXT(범위)) 동속 실측
        ' (ProbeEvalAlternatives). 수식 문자열이 없어 255자 한계/주소 이스케이프가
        ' 소멸하고, 이미 읽은 v2를 그대로 입력해 시트 상태와 무관하다.
        Dim inputArr As Variant
        If firstRow = 1 And lastRow = UBound(v2, 1) Then
            inputArr = v2                                ' 균일 열: 복사 없이 전체 전달
        Else
            Dim nBlk As Long
            nBlk = lastRow - firstRow + 1
            Dim tmp() As Variant
            ReDim tmp(1 To nBlk, 1 To 1)
            Dim r As Long
            For r = 1 To nBlk
                tmp(r, 1) = v2(firstRow + r - 1, 1)
            Next r
            inputArr = tmp
        End If

        On Error Resume Next
        rendered = Application.Text(inputArr, fmtRender)
        On Error GoTo 0
        If Not IsArray(rendered) Then
            If firstRow = lastRow Then
                If IsError(rendered) Or IsEmpty(rendered) Then Exit Function
                Dim tmp1 As Variant
                ReDim tmp1(1 To 1)
                tmp1(1) = rendered                       ' 단일 행 블록: 스칼라 반환 정규화
                rendered = tmp1
            Else
                Exit Function                            ' 다행인데 비배열 = 리프팅 실패
            End If
        End If
    End If

    ' 리프팅 반환은 크기에 따라 1차원/2차원(Nx1)이 갈린다 - 차원 불문 인덱싱 준비
    Dim rTwoD As Boolean, rBase As Long
    If needRender Then
        rTwoD = ArrIs2D(rendered)
        If rTwoD Then rBase = LBound(rendered, 1) Else rBase = LBound(rendered)
    End If

    Dim val As Variant, rv As Variant, k As String
    For i = firstRow To lastRow
        val = v2(i, 1)
        If IsEmpty(val) Then
            keys(i) = ""                                 ' TEXT(빈칸)="0" 함정: 렌더값 무시
            blankCount = blankCount + 1
        ElseIf IsError(val) Then
            k = ErrKey(val)                              ' 캐논 에러 문자열 (P12)
            If Len(k) = 0 Then k = rTarget.Cells(i, 1).Text   ' 신형/미지 코드만 셀별
            keys(i) = k
        ElseIf VarType(val) = vbString Then
            keys(i) = CStr(val)                          ' 텍스트: 자기 자신 (TEXT는 날짜문자열 강제 위험)
        ElseIf VarType(val) = vbBoolean Then
            If val Then keys(i) = "TRUE" Else keys(i) = "FALSE"   ' 캐논 (P12)
        ElseIf VarType(val) = vbDouble And needRender Then
            If rTwoD Then
                rv = rendered(rBase + i - firstRow, 1)
            Else
                rv = rendered(rBase + i - firstRow)
            End If
            If IsError(rv) Then
                keys(i) = rTarget.Cells(i, 1).Text
            Else
                k = CStr(rv)
                If Len(k) = 0 Then k = rTarget.Cells(i, 1).Text   ' [색] 등 TEXT 빈문자열 폴백
                keys(i) = k
            End If
        Else
            keys(i) = rTarget.Cells(i, 1).Text           ' 방어적 잔여 (Value2는 위 5형뿐)
        End If
    Next i

    RenderBlockBulk = True
End Function

' 배열이 2차원인지 판정한다 (Application.Text 리프팅 반환 정규화용).
Private Function ArrIs2D(ByRef a As Variant) As Boolean
    On Error GoTo OneD
    Dim n As Long
    n = UBound(a, 2)
    ArrIs2D = True
OneD:
End Function

' 에러 Variant의 캐논 표시 문자열 반환 (P12: 매칭 키는 폭 무관 캐논 문자열).
' 고전 7종은 표시 문자열이 로캘 불변임을 실측(ProbeDisplayMap). 조회 키는
' CStr(CVErr(코드))로 런타임 파생 - CStr 에러 문자열의 로캘에 의존하지 않는다.
' 미지/신형 코드(#SPILL! 등 - 표시가 로캘화됨)는 빈 문자열 반환, 호출부가 .Text 폴백.
Private Function ErrKey(ByRef v As Variant) As String
    Static pMap As Object
    On Error GoTo Unknown
    If pMap Is Nothing Then
        Set pMap = CreateObject("Scripting.Dictionary")
        pMap(CStr(CVErr(2000))) = "#NULL!"
        pMap(CStr(CVErr(2007))) = "#DIV/0!"
        pMap(CStr(CVErr(2015))) = "#VALUE!"
        pMap(CStr(CVErr(2023))) = "#REF!"
        pMap(CStr(CVErr(2029))) = "#NAME?"
        pMap(CStr(CVErr(2036))) = "#NUM!"
        pMap(CStr(CVErr(2042))) = "#N/A"
    End If
    Dim sKey As String
    sKey = CStr(v)
    If pMap.exists(sKey) Then ErrKey = pMap(sKey)
    Exit Function
Unknown:
    ErrKey = ""
End Function

' @desc 값과 블록 서식 받아 우주 표시키 반환, 렌더 불가 시 빈 문자열 반환
' 벌크 렌더(RenderBlockBulk)와 동일 규칙의 값 1건판 - 값 레인 후보키 생성용(modMatch).
' "이 값이 이 서식 블록에 있었다면 우주에 어떤 키로 존재하는가"를 재현한다.
Public Function RenderValueKey(ByVal v As Double, ByVal fmt As String) As String
    If fmt = "@" Then
        RenderValueKey = SafeWText(v, "General")   ' @숫자 = General 캐논 (P11, 우주와 대칭)
    Else
        RenderValueKey = SafeWText(v, RenderFmt(fmt, Replace(fmt, "!", ""), IsGeneralFmt(fmt)))
    End If
End Function

' @desc 캐시 자산과 새 지문 받아 변경 행만 재렌더해 패치 성공 여부 반환
' 지문 불일치 중 값만 바뀐 경우(행수/서식 구조 동일)의 증분 경로 - 편집 후
' 재스캔을 전량 재렌더 대신 변경 행 x 캐논 단건 렌더로 처리한다. 변경 행이
' 상한(1,000)을 넘거나 행수/서식 변경/병리적이면 False - 호출부가 전량 재스캔.
Public Function TryPatchAsset(ByVal asset As clsColAsset, ByVal rngCol As Range, _
                              ByVal newSig As clsColSig) As Boolean
    Const PATCH_MAX_ROWS As Long = 1000

    If asset Is Nothing Then Exit Function
    If Not asset.HasRendered Then Exit Function
    If asset.Signature Is Nothing Then Exit Function
    If newSig Is Nothing Then Exit Function

    Dim changed As Collection
    Set changed = asset.Signature.DiffRows(newSig, PATCH_MAX_ROWS)
    If changed Is Nothing Then Exit Function

    Dim rTarget As Range
    Set rTarget = ResolveTarget(rngCol)
    If rTarget Is Nothing Then Exit Function

    Dim n As Long
    n = changed.Count

    Dim rowsArr() As Long, keysArr() As String
    ReDim rowsArr(1 To n + 1)                    ' n=0 방어 (빈 배열 불가)
    ReDim keysArr(1 To n + 1)

    Dim blankDelta As Long
    If n > 0 Then
        Dim vNew As Variant, vOld As Variant
        vNew = newSig.ContentVec
        vOld = asset.Signature.ContentVec

        Dim j As Long, r As Long, k As String
        Dim val As Variant
        For j = 1 To n
            r = CLng(changed(j))
            rowsArr(j) = r
            val = vNew(r, 1)
            ' 키 규칙은 벌크(RenderBlockBulk)와 동일한 캐논 단건판
            If IsEmpty(val) Then
                k = ""
            ElseIf IsError(val) Then
                k = ErrKey(val)
                If Len(k) = 0 Then k = rTarget.Cells(r, 1).Text
            ElseIf VarType(val) = vbString Then
                k = CStr(val)
            ElseIf VarType(val) = vbBoolean Then
                If val Then k = "TRUE" Else k = "FALSE"
            ElseIf VarType(val) = vbDouble Then
                k = RenderValueKey(CDbl(val), BlockFmtForRow(newSig, r))
                If Len(k) = 0 Then k = rTarget.Cells(r, 1).Text
            Else
                k = rTarget.Cells(r, 1).Text
            End If
            keysArr(j) = k

            ' 빈칸 델타는 키가 아니라 값의 IsEmpty로 판정 ([색] 폴백 빈 키와 구분)
            If IsEmpty(vOld(r, 1)) Then
                If Not IsEmpty(val) Then blankDelta = blankDelta - 1
            ElseIf IsEmpty(val) Then
                blankDelta = blankDelta + 1
            End If
        Next j
    End If

    asset.PatchRendered rowsArr, keysArr, n, blankDelta, newSig
    TryPatchAsset = True
End Function

' 지문에서 행이 속한 블록의 서식을 찾는다 (균일 열이면 균일 서식).
Private Function BlockFmtForRow(ByVal sig As clsColSig, ByVal r As Long) As String
    If Not sig.IsMixed Then
        BlockFmtForRow = sig.FmtUniform
        Exit Function
    End If
    Dim blk As Variant
    For Each blk In sig.Blocks
        If r >= CLng(blk(0)) And r <= CLng(blk(1)) Then
            BlockFmtForRow = CStr(blk(2))
            Exit Function
        End If
    Next blk
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

' 셀별 폴백(과편성 서식 전용): 셀마다 서식이 달라 블록 벌크가 불가한 경우.
' 키 규칙은 벌크와 동일한 캐논(fmtLocal + WText) - .Text를 주 경로로 쓰지 않아
' 폭 잘림/CF/행숨김 이력 저하와 무관하다(P7/P11/P12 + ProbeTextDecay 실측).
' fmtLocal은 이력 저하 면역 실측(ProbeDisplayMap). AutoFit 재독은 소멸.
Private Function BuildPerCell(ByVal rTarget As Range, ByVal v2 As Variant, _
                             ByVal nRows As Long, ByVal vm As Object, _
                             ByVal rowByKey As Object) As Long
    Dim keys() As String
    ReDim keys(1 To nRows)

    Dim i As Long
    Dim c As Range, f As String
    For i = 1 To nRows
        If IsEmpty(v2(i, 1)) Then
            keys(i) = ""
        ElseIf IsError(v2(i, 1)) Then
            keys(i) = ErrKey(v2(i, 1))                   ' 캐논 에러 문자열 (P12)
            If Len(keys(i)) = 0 Then keys(i) = rTarget.Cells(i, 1).Text
        ElseIf VarType(v2(i, 1)) = vbString Then
            keys(i) = CStr(v2(i, 1))                     ' 텍스트: 자기 자신
        ElseIf VarType(v2(i, 1)) = vbBoolean Then
            If v2(i, 1) Then keys(i) = "TRUE" Else keys(i) = "FALSE"   ' 캐논 (P12)
        ElseIf VarType(v2(i, 1)) = vbDouble Then
            Set c = rTarget.Cells(i, 1)
            f = SafeFmtLocal(c)
            If f = "@" Then
                keys(i) = SafeWText(v2(i, 1), "General") ' @숫자 = General 캐논 (P11)
            Else
                keys(i) = SafeWText(v2(i, 1), RenderFmt(f, Replace(f, "!", ""), IsGeneralFmt(f)))
            End If
            If Len(keys(i)) = 0 Then keys(i) = c.Text    ' [색] 등 WText 빈문자열 폴백
        Else
            keys(i) = rTarget.Cells(i, 1).Text           ' 방어적 잔여 (Value2는 위 5형뿐)
        End If
    Next i

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
' (true last row 방어는 범위를 만드는 호출부 소관 - modUnivFilter.ResolveFilterTarget)
Private Function ResolveTarget(ByVal rngCol As Range) As Range
    If rngCol Is Nothing Then Exit Function
    Set ResolveTarget = rngCol.Columns(1)
End Function

' General 서식(로컬 "G/표준" 또는 "General")인지 확인한다. General은 TEXT가
' 로컬 토큰을 미해석하므로 영문 "General"로 매핑해 렌더한다(RenderFmt).
Private Function IsGeneralFmt(ByVal fmt As String) As Boolean
    IsGeneralFmt = (fmt = "General") Or (fmt = "G/" & ChrW(&HD45C) & ChrW(&HC900))
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
