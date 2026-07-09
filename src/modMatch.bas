Attribute VB_Name = "modMatch"
' @desc 토큰과 열 자산의 2레인 매칭 판정 담당
' @depends modInputParse, clsSet, clsColAsset
Option Explicit

' @desc 토큰 배열과 자산 받아 매칭/누락/보정 결과 반환, 토큰 없으면 빈 결과 반환
Public Function Execute(ByVal vTokens As Variant, ByVal asset As clsColAsset, _
                        ByVal bPartial As Boolean) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim matched As New clsSet
    Dim corrections As Object
    Set corrections = CreateObject("Scripting.Dictionary")
    Dim missing As New Collection

    Dim i As Long, t As String
    If IsArray(vTokens) Then
        For i = LBound(vTokens) To UBound(vTokens)
            t = CStr(vTokens(i))
            If Not MatchToken(t, asset, bPartial, matched, corrections) Then
                missing.Add t
            End If
        Next i
    End If

    Set result("Matched") = matched
    Set result("Corrections") = corrections
    result("Missing") = ToStringArray(missing)
    result("InputCount") = CountTokens(vTokens)
    Set Execute = result
End Function

' 토큰 1개를 표시 레인 -> 값 레인 순으로 매칭한다. 성공 시 True.
Private Function MatchToken(ByVal t As String, ByVal asset As clsColAsset, _
                            ByVal bPartial As Boolean, ByVal matched As clsSet, _
                            ByVal corrections As Object) As Boolean
    Dim hit As Boolean

    If bPartial Then
        ' 부분일치: 표시 키 전수 스캔 (적용은 여전히 정확한 키 배열)
        Dim keys As Variant, k As Long
        keys = asset.Displays.KeysArray
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
