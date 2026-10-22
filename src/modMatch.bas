Attribute VB_Name = "modMatch"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc 토큰과 열 자산의 2레인 매칭 판정 담당
' @depends modInputParse, modColumnScan, clsSet, clsColAsset, clsColSig
Option Explicit

' 지연 우주(Stage C): RenderedVec 보유 자산의 완전일치는 인덱스(집합/지도) 없이
' 자산 내부 렌더키 1패스 멤버십(ScanCount, 실측 91k행 ~10ms)으로 매칭하고, 매칭
' 행수를 부산물(MatchedRows)로 반환해 자가검증이 스캔에 공짜로 얹힌다. 값 레인은
' "파싱값이 블록 서식으로 렌더되면 우주에 어떤 키인가"(RenderValueKey) 후보의
' 멤버십 - 셀별 폴백도 같은 경로를 사용한다. 부분일치/수동 인덱스 자산은
' 기존 인덱스 경로 유지(지연 구축 1회, MatchedRows=-1).

' @desc 토큰 배열과 자산 받아 매칭/누락/보정 결과 반환, 토큰 없으면 빈 결과 반환
Public Function Execute(ByVal vTokens As Variant, ByVal asset As clsColAsset, _
                        ByVal bPartial As Boolean) As Object
    If asset.HasRendered And Not bPartial Then
        Set Execute = ExecuteScan(vTokens, asset)
        Exit Function
    End If

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim matched As New clsSet
    Dim corrections As Object
    Set corrections = CreateObject("Scripting.Dictionary")
    Dim missing As New Collection

    Dim i As Long, t As String, keys As Variant
    If IsArray(vTokens) Then
        If bPartial Then
            ' 검색어마다 전체 키 배열을 복사하지 않고 실행 안에서 재사용한다.
            If CountTokens(vTokens) > 0 Then
                keys = asset.Displays.KeysArray
                If TryPartialAho(keys, vTokens, matched, missing) Then GoTo Done
            End If
        End If
        For i = LBound(vTokens) To UBound(vTokens)
            t = CStr(vTokens(i))
            If Not MatchToken(t, asset, bPartial, matched, corrections, keys) Then
                missing.Add t
            End If
        Next i
    End If

Done:
    Set result("Matched") = matched
    Set result("Corrections") = corrections
    result("Missing") = ToStringArray(missing)
    result("InputCount") = CountTokens(vTokens)
    result("MatchedRows") = -1&              ' 인덱스 경로: 계수는 ExpectedRows가 담당
    Set Execute = result
End Function

' 완전일치 직스캔 매칭: 인덱스 미구축 상태를 유지한 채 표시/값 레인을 판정한다.
' pass1 토큰 멤버십+행계수 -> 미스 토큰만 값 레인 후보 -> pass2 후보 멤버십+행계수.
' 후보는 토큰과 겹치지 않게 걸러 두 패스의 행계수가 서로소로 합산된다.
Private Function ExecuteScan(ByVal vTokens As Variant, ByVal asset As clsColAsset) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim matched As New clsSet
    Dim corrections As Object
    Set corrections = CreateObject("Scripting.Dictionary")
    Dim missing As New Collection
    Dim rows As Long

    Dim nTok As Long
    nTok = CountTokens(vTokens)
    If nTok = 0 Then GoTo Done

    ' pass 1: 토큰 소형 dict 멤버십 + 매칭 행수 (실측 91k행 ~10ms)
    Dim tokDict As Object
    Set tokDict = CreateObject("Scripting.Dictionary")
    tokDict.CompareMode = vbTextCompare
    Dim i As Long, t As String
    For i = LBound(vTokens) To UBound(vTokens)
        t = CStr(vTokens(i))
        If Not tokDict.exists(t) Then tokDict(t) = False
    Next i
    Dim found As Boolean
    If tokDict.Count = 1 Then
        rows = asset.ScanOne(CStr(tokDict.Keys()(0)), found)   ' 토큰 1개: StrComp 직비교
        If found Then tokDict(tokDict.Keys()(0)) = True
    Else
        rows = asset.ScanCount(tokDict)
    End If

    ' 값 레인 후보: 미스 토큰의 파싱값 x 블록 서식 -> 우주 후보키
    ' (토큰과 겹치는 후보는 pass1이 이미 판정 - candDict에 넣지 않아 계수 중복 방지)
    Dim fmts As Collection
    Set fmts = BlockFormats(asset)
    Dim candDict As Object                    ' 후보키 -> 발견 여부
    Set candDict = CreateObject("Scripting.Dictionary")
    candDict.CompareMode = vbTextCompare
    Dim candsByTok As Object                  ' 토큰 -> 후보키 vbNullChar 연결
    Set candsByTok = CreateObject("Scripting.Dictionary")

    Dim vTok As Variant, cands As Variant
    Dim j As Long, f As Long, ck As String, lst As String
    For Each vTok In tokDict.Keys
        If Not tokDict(vTok) Then
            cands = modInputParse.TryParseValues(CStr(vTok))
            lst = ""
            For j = LBound(cands) To UBound(cands)
                For f = 1 To fmts.Count
                    ck = modColumnScan.RenderValueKey(CDbl(cands(j)), CStr(fmts(f)))
                    If Len(ck) > 0 Then
                        If Not tokDict.exists(ck) Then
                            If Not candDict.exists(ck) Then candDict(ck) = False
                        End If
                        If InStr(1, vbNullChar & lst & vbNullChar, _
                                 vbNullChar & ck & vbNullChar) = 0 Then
                            If Len(lst) > 0 Then lst = lst & vbNullChar
                            lst = lst & ck
                        End If
                    End If
                Next f
            Next j
            If Len(lst) > 0 Then candsByTok(vTok) = lst
        End If
    Next vTok

    ' pass 2: 후보키 멤버십 + 행수 합산 (실측 값 브리지 ~12ms, 후보 없으면 생략)
    If candDict.Count = 1 Then
        rows = rows + asset.ScanOne(CStr(candDict.Keys()(0)), found)
        If found Then candDict(candDict.Keys()(0)) = True
    ElseIf candDict.Count > 0 Then
        rows = rows + asset.ScanCount(candDict)
    End If

    ' 취합: 원 토큰 순서대로 매칭/보정/누락 분류 (기존 계약과 동일)
    Dim parts() As String, d As Long, corr As String, anyHit As Boolean
    For i = LBound(vTokens) To UBound(vTokens)
        t = CStr(vTokens(i))
        If tokDict(t) Then
            matched.Add t
        ElseIf candsByTok.exists(t) Then
            parts = Split(candsByTok(t), vbNullChar)
            corr = ""
            anyHit = False
            For d = 0 To UBound(parts)
                ck = parts(d)
                If tokDict.exists(ck) Then
                    found = tokDict(ck)       ' 토큰과 겹친 후보는 pass1 판정 재사용
                Else
                    found = candDict(ck)
                End If
                If found Then
                    matched.Add ck
                    If Len(corr) > 0 Then corr = corr & ", "
                    corr = corr & ck
                    anyHit = True
                End If
            Next d
            If anyHit Then
                corrections(t) = corr
            Else
                missing.Add t
            End If
        Else
            missing.Add t
        End If
    Next i

Done:
    Set result("Matched") = matched
    Set result("Corrections") = corrections
    result("Missing") = ToStringArray(missing)
    result("InputCount") = nTok
    result("MatchedRows") = rows
    Set ExecuteScan = result
End Function

' 자산 지문에서 블록 서식을 수집한다. 분할 실패 시 셀별 관측 서식을 사용한다.
Private Function BlockFormats(ByVal asset As clsColAsset) As Collection
    If Not asset.FallbackFormats Is Nothing Then
        Set BlockFormats = asset.FallbackFormats
        Exit Function
    End If
    Dim c As New Collection
    Dim sig As clsColSig
    Set sig = asset.Signature
    If sig.IsMixed Then
        Dim blk As Variant
        For Each blk In sig.Blocks
            c.Add CStr(blk(2))
        Next blk
    Else
        c.Add sig.FmtUniform
    End If
    Set BlockFormats = c
End Function

' 토큰 1개를 표시 레인 -> 값 레인 순으로 매칭한다. 성공 시 True.
Private Function MatchToken(ByVal t As String, ByVal asset As clsColAsset, _
                            ByVal bPartial As Boolean, ByVal matched As clsSet, _
                            ByVal corrections As Object, ByRef keys As Variant) As Boolean
    Dim hit As Boolean

    If bPartial Then
        ' 부분일치: 표시 키 전수 스캔 (적용은 여전히 정확한 키 배열)
        Dim k As Long
        For k = LBound(keys) To UBound(keys)
            If InStr(1, CStr(keys(k)), t, vbTextCompare) > 0 Then
                matched.Add keys(k)
                hit = True
            End If
        Next k
    Else
        ' 표시 레인: 보이는 그대로 입력한 경우
        If asset.Displays.Contains(t) Then
            matched.Add t
            hit = True
        Else
            ' 값 레인: 숫자/날짜 해석 -> 값 지도 -> 표시 키 합류 (자동보정 기록)
            ' 값 지도 값은 vbNullChar 연결 문자열 (성능 계약 - modColumnScan 참조)
            Dim cands As Variant
            cands = modInputParse.TryParseValues(t)

            Dim j As Long, key As String, corr As String
            Dim parts() As String, d As Long
            For j = LBound(cands) To UBound(cands)
                key = CStr(cands(j))
                If asset.ValueMap.exists(key) Then
                    parts = Split(asset.ValueMap(key), vbNullChar)
                    For d = 0 To UBound(parts)
                        matched.Add parts(d)
                        If Len(corr) > 0 Then corr = corr & ", "
                        corr = corr & parts(d)
                    Next d
                    hit = True
                End If
            Next j
            If hit Then corrections(t) = corr
        End If
    End If

    MatchToken = hit
End Function

' Collection의 문자열을 0-기반 배열로 변환한다(비면 빈 배열).
Private Function ToStringArray(ByVal col As Collection) As Variant
    If col.Count = 0 Then
        ToStringArray = Array()
    Else
        Dim arr() As Variant
        ReDim arr(0 To col.Count - 1)
        Dim i As Long
        For i = 1 To col.Count
            arr(i - 1) = col(i)
        Next i
        ToStringArray = arr
    End If
End Function

Private Function CountTokens(ByVal vTokens As Variant) As Long
    If IsArray(vTokens) Then
        On Error Resume Next
        CountTokens = UBound(vTokens) - LBound(vTokens) + 1
        If CountTokens < 0 Then CountTokens = 0
        On Error GoTo 0
    End If
End Function

' shortcut: 합성 벤치의 보수적 기준. 실제 데이터에서 역전되면 기준만 재측정한다.
Private Function TryPartialAho(ByRef keys As Variant, ByRef tokens As Variant, _
                               ByVal matched As clsSet, ByVal missing As Collection) As Boolean
    Dim n As Long, count As Long, i As Long, total As Double, threshold As Long, cap As Long
    n = CountTokens(tokens)
    If n < 10 Then Exit Function
    count = UBound(keys) - LBound(keys) + 1
    If count <= 1024 Then Exit Function
    For i = 0 To 31
        total = total + Len(CStr(keys(LBound(keys) + CLng(Fix(CDbl(i) * (count - 1) / 31)))))
    Next i
    If total <= 16 * 32 Then
        threshold = 10
    ElseIf total <= 64 * 32 Then
        threshold = 20
    Else
        threshold = 50
    End If
    If n < threshold Then Exit Function
    ' ASCII 대소문자 대응이 다른 로캘은 기존 텍스트 비교를 유지한다.
    If StrComp(LCase$("ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "abcdefghijklmnopqrstuvwxyz", vbBinaryCompare) <> 0 Then Exit Function
    For i = LBound(tokens) To UBound(tokens)
        If Len(CStr(tokens(i))) = 0 Then Exit Function
        If Not PartialAscii(CStr(tokens(i))) Then Exit Function
        ' 전이표 용량을 제한하고 큰 검색어 집합은 기존 경로로 보낸다.
        If Len(CStr(tokens(i))) > 8192 - cap Then Exit Function
        cap = cap + Len(CStr(tokens(i)))
    Next i
    PartialAho keys, tokens, matched, missing, cap
    TryPartialAho = True
End Function

Private Sub PartialAho(ByRef keys As Variant, ByRef tokens As Variant, _
                       ByVal result As clsSet, ByVal missing As Collection, ByVal cap As Long)
    Dim i As Long, j As Long, c As Long, s As Long, node As Long
    Dim n As Long, text As String, asciiKey As Boolean
    Dim goToState() As Long, fail() As Long, output() As Long, suffix() As Long
    Dim nextToken() As Long, queue() As Long, terminal As Long, v As Long, f As Long
    Dim head As Long, tail As Long, matched As Boolean, item As Long
    Dim tokenHit() As Boolean
    n = UBound(tokens) - LBound(tokens) + 1
    ReDim tokenHit(LBound(tokens) To UBound(tokens))
    ReDim goToState(0 To cap, 0 To 127)
    ReDim fail(0 To cap)
    ReDim output(0 To cap)
    ReDim suffix(0 To cap)
    ReDim nextToken(1 To n)
    ReDim queue(0 To cap)
    For i = LBound(tokens) To UBound(tokens)
        text = LCase$(CStr(tokens(i)))
        s = 0
        For j = 1 To Len(text)
            c = AscW(Mid$(text, j, 1))
            If goToState(s, c) = 0 Then
                node = node + 1
                goToState(s, c) = node
            End If
            s = goToState(s, c)
        Next j
        item = i - LBound(tokens) + 1
        nextToken(item) = output(s)
        output(s) = item
    Next i
    For c = 0 To 127
        If goToState(0, c) <> 0 Then
            queue(tail) = goToState(0, c)
            tail = tail + 1
        End If
    Next c
    Do While head < tail
        s = queue(head)
        head = head + 1
        For c = 0 To 127
            v = goToState(s, c)
            If v <> 0 Then
                f = goToState(fail(s), c)
                fail(v) = f
                If output(f) <> 0 Then
                    suffix(v) = f
                Else
                    suffix(v) = suffix(f)
                End If
                queue(tail) = v
                tail = tail + 1
            Else
                goToState(s, c) = goToState(fail(s), c)
            End If
        Next c
    Loop
    For i = LBound(keys) To UBound(keys)
        text = CStr(keys(i))
        matched = False
        asciiKey = PartialAscii(text)
        If asciiKey Then
            text = LCase$(text)
            s = 0
            For j = 1 To Len(text)
                c = AscW(Mid$(text, j, 1))
                s = goToState(s, c)
                terminal = s
                Do While terminal <> 0
                    item = output(terminal)
                    Do While item <> 0
                        tokenHit(item + LBound(tokens) - 1) = True
                        matched = True
                        item = nextToken(item)
                    Loop
                    terminal = suffix(terminal)
                Loop
            Next j
        Else
            For j = LBound(tokens) To UBound(tokens)
                If InStr(1, CStr(keys(i)), CStr(tokens(j)), vbTextCompare) > 0 Then
                    tokenHit(j) = True
                    matched = True
                End If
            Next j
        End If
        If matched Then result.Add keys(i)
    Next i
    For i = LBound(tokens) To UBound(tokens)
        If Not tokenHit(i) Then missing.Add CStr(tokens(i))
    Next i
End Sub

Private Function PartialAscii(ByVal text As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1))
        If c < 32 Or c > 126 Then Exit Function
    Next i
    PartialAscii = True
End Function
