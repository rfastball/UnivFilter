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
' 멤버십 - 표시 의미론으로 일관된다. 부분일치/레거시 자산(병리적/벌크 실패)은
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
            If CountTokens(vTokens) > 0 Then keys = asset.Displays.KeysArray
        End If
        For i = LBound(vTokens) To UBound(vTokens)
            t = CStr(vTokens(i))
            If Not MatchToken(t, asset, bPartial, matched, corrections, keys) Then
                missing.Add t
            End If
        Next i
    End If

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

' 자산 지문에서 블록 서식 목록을 수집한다(균일이면 1개, 병리적 자산은 미호출).
Private Function BlockFormats(ByVal asset As clsColAsset) As Collection
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
