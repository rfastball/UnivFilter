Attribute VB_Name = "modColumnScan"
' @desc 열 1회 스캔해 표시키/값지도 자산 구축 담당
' @depends clsColAsset, clsSet, modLogger
Option Explicit

' 표시 키 규칙 (실측: AutoFilterTest gap-probe-findings.md P4/P6):
'  - 서식에 * 포함: TEXT(Value2, fmtLocal에서 ! 제거) 접힌형이 매칭 키
'  - 그 외: Range.Text가 매칭 키. #-런(좁은 열)은 AutoFit 재독 후 폭 복원
'  - 텍스트/에러/불리언 셀은 .Text 그대로, 빈칸은 키 없음(BlankCount)

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

    ' 벌크 읽기 (값/빈칸/타입 판별)
    Dim v2 As Variant
    v2 = To2D(rTarget.Value2)

    ' 표시 키 렌더
    Dim keys() As String
    ReDim keys(1 To nRows)
    RenderKeys rTarget, v2, keys, nRows

    ' 집합/지도 구축.
    ' 성능 계약: 행 루프의 늦은 바인딩 Dictionary 호출은 행당 최소로 유지한다
    ' (행마다 clsSet을 꺼내 Add하면 10만 행에서 수초 - modPerfProbe로 계측된 병목).
    ' (값,표시) 쌍은 복합키 대입 1회로 dedup하고 clsSet 조립은 고유 쌍에만 수행.
    Dim pairSeen As Object
    Set pairSeen = CreateObject("Scripting.Dictionary")
    pairSeen.CompareMode = vbTextCompare

    Dim i As Long, blankCount As Long
    Dim k As String
    For i = 1 To nRows
        If IsEmpty(v2(i, 1)) Then
            blankCount = blankCount + 1
        Else
            k = keys(i)
            If Len(k) > 0 Then
                rowByKey(k) = rowByKey(k) + 1   ' 미존재 키는 Empty+1=1
                ' 값 지도 재료: 숫자/날짜 셀만 (Value2는 이들을 Double로 반환)
                If VarType(v2(i, 1)) = vbDouble Then
                    pairSeen(CStr(v2(i, 1)) & vbNullChar & k) = Empty
                End If
            End If
        End If
    Next i


    ' 표시 집합: 행수 지도의 키가 곧 고유 표시 키
    displays.LoadFromArray rowByKey.Keys

    ' 값 지도: 고유 (값,표시) 쌍에서 vbNullChar 연결 문자열로 조립.
    ' 성능 계약: 값마다 New clsSet 금지 - 내부 CreateObject가 개당 ~0.5ms라
    ' 고유값 수천 개면 수초가 된다 (modPerfProbe로 계측된 병목).
    Dim vPair As Variant, p As Long, vk As String
    For Each vPair In pairSeen.Keys
        p = InStr(1, vPair, vbNullChar)
        vk = Left$(vPair, p - 1)
        k = Mid$(vPair, p + 1)
        If vm.exists(vk) Then
            vm(vk) = vm(vk) & vbNullChar & k
        Else
            vm(vk) = k
        End If
    Next vPair


    asset.Init displays, vm, rowByKey, blankCount, nRows, sID, lGeneration
    Set BuildAsset = asset
End Function

' 주입 범위를 신뢰한다: 끝쪽 빈칸도 필터 대상 행이다.
' (true last row 방어는 범위를 만드는 호출부 소관 - modShortcuts.SetupFormData)
Private Function ResolveTarget(ByVal rngCol As Range) As Range
    If rngCol Is Nothing Then Exit Function
    Set ResolveTarget = rngCol.Columns(1)
End Function

' 셀별 표시 키를 실측 규칙대로 렌더링한다.
Private Sub RenderKeys(ByVal rTarget As Range, ByVal v2 As Variant, _
                       ByRef keys() As String, ByVal nRows As Long)
    Dim vFmt As Variant
    vFmt = rTarget.NumberFormatLocal   ' 혼합 서식이면 Null

    Dim i As Long
    Dim uniform As Boolean
    uniform = Not IsNull(vFmt)

    ' 주의: VBA And는 단락평가가 없으므로 Null 가드를 먼저 편다
    Dim padUniform As Boolean
    If uniform Then padUniform = (InStr(CStr(vFmt), "*") > 0)

    If padUniform Then
        ' 균일 * 패딩 서식: 고유값당 접힌형 TEXT 1회 (P6 확정 규칙)
        Dim cleanFmt As String
        cleanFmt = Replace(CStr(vFmt), "!", "")
        Dim rendered As Object
        Set rendered = CreateObject("Scripting.Dictionary")

        For i = 1 To nRows
            If IsEmpty(v2(i, 1)) Then
                keys(i) = ""
            ElseIf IsError(v2(i, 1)) Then
                keys(i) = rTarget.Cells(i, 1).Text
            Else
                Dim ck As String
                ck = CStr(v2(i, 1))
                If Not rendered.exists(ck) Then
                    rendered(ck) = SafeWText(v2(i, 1), cleanFmt)
                End If
                keys(i) = rendered(ck)
                If Len(keys(i)) = 0 Then keys(i) = rTarget.Cells(i, 1).Text
            End If
        Next i
    ElseIf uniform Then
        ' 균일 일반 서식: 셀별 .Text (서식 재검사는 루프 밖에서 이미 끝남)
        Dim needHashFix As Boolean
        For i = 1 To nRows
            If IsEmpty(v2(i, 1)) Then
                keys(i) = ""
            Else
                keys(i) = rTarget.Cells(i, 1).Text
                If VarType(v2(i, 1)) = vbDouble And IsHashRun(keys(i)) Then
                    needHashFix = True
                End If
            End If
        Next i
        If needHashFix Then FixHashKeys rTarget, keys, nRows
    Else
        ' 혼합 서식: 셀별로 서식을 읽어 * 패딩 검사 (드문 경로)
        Dim needHashFix2 As Boolean
        For i = 1 To nRows
            If IsEmpty(v2(i, 1)) Then
                keys(i) = ""
            Else
                Dim c As Range
                Set c = rTarget.Cells(i, 1)
                Dim f As String
                f = SafeFmtLocal(c)

                If InStr(f, "*") > 0 And VarType(v2(i, 1)) = vbDouble Then
                    keys(i) = SafeWText(v2(i, 1), Replace(f, "!", ""))
                    If Len(keys(i)) = 0 Then keys(i) = c.Text
                Else
                    keys(i) = c.Text
                    If VarType(v2(i, 1)) = vbDouble And IsHashRun(keys(i)) Then
                        needHashFix2 = True
                    End If
                End If
            End If
        Next i
        If needHashFix2 Then FixHashKeys rTarget, keys, nRows
    End If
End Sub

' #-런 키를 AutoFit 재독으로 복구하고 열 폭을 복원한다.
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
