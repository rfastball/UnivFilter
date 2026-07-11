Attribute VB_Name = "modPerfProbe"
' @desc 스캔 병목 계측 일회용 프로브 담당
' @depends modColumnScan, modCache, modMatch, clsColSig, clsColAsset, clsSet, Utilities, LogShim
Option Explicit

Private Const SHEET_NAME As String = "UF_PerfProbe"

' @desc 대규모 열에서 BuildAsset 단계별 소요를 로그로 기록함
Public Sub RunScanPerf()
    LogShim.Reset
    Utilities.OptimizedMode True

    On Error GoTo Done
    MeasureFormat "general", "General", 100000
    MeasureFormat "comma", "#,##0", 100000
    MeasureFormat "date", "yyyy-mm-dd", 50000
    MeasureFormat "text", "@", 50000
    MeasureFormat "acc_pad", "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""_-;_-@_-", 50000
    MeasureHighCardinality 100000
    MeasureBlockMixed 50000
    MeasureMixedFormat 50000
    MeasureBlanksSameFmt 50000
    MeasureBlanksDiffFmt 50000
    MeasureRawOps 100000

Done:
    Utilities.OptimizedMode False
    DropSheet
    LogShim.LogEvent "perf_done", Array("ok"), Array(Err.Number = 0)
End Sub

' 지정 서식/행수의 열을 만들어 BuildAsset 전체 소요를 잰다.
Private Sub MeasureFormat(ByVal tag As String, ByVal numFmt As String, ByVal nRows As Long)
    Dim ws As Worksheet
    Set ws = PrepSheet()

    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To 1)
    Dim i As Long
    For i = 1 To nRows
        If numFmt = "@" Then
            arr(i, 1) = "V" & (i Mod 5000)
        ElseIf numFmt = "yyyy-mm-dd" Then
            arr(i, 1) = CDbl(DateSerial(2024, 1, 1) + (i Mod 365))
        Else
            arr(i, 1) = (i Mod 5000) * 7
        End If
    Next i

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, 1))
    rng.NumberFormat = numFmt
    rng.Value2 = arr

    modCache.ClearCache

    Dim t0 As Single
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rng, "PERF")
    LogShim.LogEvent "perf_buildasset", _
        Array("fmt", "rows", "ms", "displays"), _
        Array(tag, nRows, CLng((Timer - t0) * 1000!), asset.Displays.Count)
End Sub

' 고유값 100%(중복 없음) 균일 열: 중복제거 이득이 0인 최악 경우.
Private Sub MeasureHighCardinality(ByVal nRows As Long)
    Dim ws As Worksheet
    Set ws = PrepSheet()

    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To 1)
    Dim i As Long
    For i = 1 To nRows
        arr(i, 1) = i          ' 전부 고유
    Next i
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, 1))
    rng.NumberFormat = "#,##0"
    rng.Value2 = arr

    modCache.ClearCache
    Dim t0 As Single
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rng, "PERF")
    LogShim.LogEvent "perf_buildasset", Array("fmt", "rows", "ms", "displays"), _
        Array("uniq100pct", nRows, CLng((Timer - t0) * 1000!), asset.Displays.Count)
End Sub

' 블록 구조 혼합 열(전반/후반 서로 다른 서식): 이진 분할 이득 측정.
Private Sub MeasureBlockMixed(ByVal nRows As Long)
    Dim ws As Worksheet
    Set ws = PrepSheet()

    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To 1)
    Dim i As Long
    For i = 1 To nRows
        arr(i, 1) = (i Mod 100) * 7   ' 블록당 저카디널리티(실사용 필터 열)
    Next i
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, 1))
    rng.Value2 = arr
    ' 3개 블록: 전반 #,##0 / 중반 0 / 후반 $#,##0
    Dim third As Long
    third = nRows \ 3
    ws.Range(ws.Cells(1, 1), ws.Cells(third, 1)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(third + 1, 1), ws.Cells(2 * third, 1)).NumberFormat = "0"
    ws.Range(ws.Cells(2 * third + 1, 1), ws.Cells(nRows, 1)).NumberFormat = "$#,##0"

    modCache.ClearCache
    Dim t0 As Single
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rng, "PERF")
    LogShim.LogEvent "perf_buildasset", Array("fmt", "rows", "ms", "displays"), _
        Array("block_mixed3", nRows, CLng((Timer - t0) * 1000!), asset.Displays.Count)
End Sub

' 혼합 서식 열(행마다 다른 서식): 셀별 폴백의 실제 비용(병리적 최악).
Private Sub MeasureMixedFormat(ByVal nRows As Long)
    Dim ws As Worksheet
    Set ws = PrepSheet()

    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To 1)
    Dim i As Long
    For i = 1 To nRows
        arr(i, 1) = (i Mod 5000) * 7
    Next i
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, 1))
    rng.Value2 = arr
    ' 행마다 서식을 번갈아 -> NumberFormatLocal이 Null(혼합)
    Dim blk As Long
    For blk = 1 To nRows Step 2
        ws.Cells(blk, 1).NumberFormat = "#,##0"
    Next blk
    For blk = 2 To nRows Step 2
        ws.Cells(blk, 1).NumberFormat = "0"
    Next blk

    modCache.ClearCache
    Dim t0 As Single
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rng, "PERF")
    LogShim.LogEvent "perf_buildasset", Array("fmt", "rows", "ms", "displays"), _
        Array("mixed_fmt", nRows, CLng((Timer - t0) * 1000!), asset.Displays.Count)
End Sub

' 데이터 균일 + 빈칸 같은 서식(열 전체 서식): 오판 없음 기대(단일 고속 경로).
Private Sub MeasureBlanksSameFmt(ByVal nRows As Long)
    Dim ws As Worksheet
    Set ws = PrepSheet()
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, 1))
    rng.NumberFormat = "#,##0"        ' 빈칸 포함 열 전체 동일 서식

    Dim i As Long
    For i = 1 To nRows
        If i Mod 10 <> 0 Then ws.Cells(i, 1).Value = (i Mod 100) * 7  ' 10칸마다 빈칸
    Next i

    modCache.ClearCache
    Dim t0 As Single
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rng, "PERF")
    LogShim.LogEvent "perf_buildasset", Array("fmt", "rows", "ms", "displays"), _
        Array("blank_samefmt", nRows, CLng((Timer - t0) * 1000!), asset.Displays.Count)
End Sub

' 데이터 균일 + 빈칸만 다른 서식(General): NumberFormatLocal Null 오판 -> 분할 폭증.
Private Sub MeasureBlanksDiffFmt(ByVal nRows As Long)
    Dim ws As Worksheet
    Set ws = PrepSheet()
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, 1))
    rng.NumberFormat = "#,##0"

    Dim i As Long
    For i = 1 To nRows
        If i Mod 10 = 0 Then
            ws.Cells(i, 1).NumberFormat = "General"   ' 빈칸만 General
        Else
            ws.Cells(i, 1).Value = (i Mod 100) * 7
        End If
    Next i

    modCache.ClearCache
    Dim t0 As Single
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rng, "PERF")
    LogShim.LogEvent "perf_buildasset", Array("fmt", "rows", "ms", "displays"), _
        Array("blank_difffmt", nRows, CLng((Timer - t0) * 1000!), asset.Displays.Count)
End Sub

' 진단: 빈칸 제외 데이터 셀만의 NumberFormatLocal이 균일값을 값싸게 주는가?
Public Sub DiagDataOnlyFormat()
    LogShim.Reset
    Const N As Long = 50000
    Dim ws As Worksheet
    Set ws = PrepSheet()
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(N, 1))
    rng.NumberFormat = "#,##0"
    Dim i As Long
    For i = 1 To N
        If i Mod 10 = 0 Then
            ws.Cells(i, 1).NumberFormat = "General"
        Else
            ws.Cells(i, 1).Value = (i Mod 100) * 7
        End If
    Next i

    Dim t0 As Single
    ' 전체 범위 서식
    t0 = Timer
    Dim whole As Variant
    whole = rng.NumberFormatLocal
    Dim wStr As String
    If IsNull(whole) Then wStr = "(null)" Else wStr = CStr(whole)
    LogShim.LogEvent "diag", Array("op", "val", "ms"), _
        Array("whole", wStr, CLng((Timer - t0) * 1000!))

    ' 데이터(상수) 셀만
    t0 = Timer
    Dim dc As Range, dfmt As Variant
    On Error Resume Next
    Set dc = rng.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    Dim areas As Long, dStr As String
    If dc Is Nothing Then
        areas = 0
        dStr = "(no cells)"
    Else
        areas = dc.Areas.Count
        dfmt = dc.NumberFormatLocal
        If IsNull(dfmt) Then dStr = "(null)" Else dStr = CStr(dfmt)
    End If
    LogShim.LogEvent "diag", Array("op", "areas", "val", "ms"), _
        Array("constants", areas, dStr, CLng((Timer - t0) * 1000!))

    DropSheet
    LogShim.LogEvent "perf_done", Array("ok"), Array(True)
End Sub

' 원시 연산별 단가를 잰다: 벌크 Value2 / 셀별 .Text / 셀별 fmtLocal / 고유값 TEXT.
Private Sub MeasureRawOps(ByVal nRows As Long)
    Dim ws As Worksheet
    Set ws = PrepSheet()

    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To 1)
    Dim i As Long
    For i = 1 To nRows
        arr(i, 1) = (i Mod 5000) * 7
    Next i
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, 1))
    rng.NumberFormat = "#,##0"
    rng.Value2 = arr

    Dim t0 As Single, v As Variant, s As String

    t0 = Timer
    v = rng.Value2
    LogShim.LogEvent "perf_raw", Array("op", "rows", "ms"), _
        Array("bulk_value2", nRows, CLng((Timer - t0) * 1000!))

    t0 = Timer
    For i = 1 To nRows
        s = rng.Cells(i, 1).Text
    Next i
    LogShim.LogEvent "perf_raw", Array("op", "rows", "ms"), _
        Array("percell_text", nRows, CLng((Timer - t0) * 1000!))

    t0 = Timer
    For i = 1 To 10000
        s = rng.Cells(i, 1).NumberFormatLocal
    Next i
    LogShim.LogEvent "perf_raw", Array("op", "rows", "ms"), _
        Array("percell_fmtlocal_10k", 10000, CLng((Timer - t0) * 1000!))

    t0 = Timer
    Dim uniq As Object
    Set uniq = CreateObject("Scripting.Dictionary")
    Dim k As String
    For i = 1 To nRows
        k = CStr(v(i, 1))
        If Not uniq.exists(k) Then
            uniq(k) = Application.WorksheetFunction.Text(v(i, 1), "#,##0")
        End If
    Next i
    LogShim.LogEvent "perf_raw", Array("op", "rows", "ms"), _
        Array("bulk_unique_text", nRows, CLng((Timer - t0) * 1000!))
End Sub

' @desc BuildAsset 내부 단계를 복제해 단계별 소요를 잰다
Public Sub RunStagePerf()
    LogShim.Reset
    Utilities.OptimizedMode True
    On Error GoTo Done

    Const N As Long = 100000
    Dim ws As Worksheet
    Set ws = PrepSheet()

    Dim arr() As Variant
    ReDim arr(1 To N, 1 To 1)
    Dim i As Long
    For i = 1 To N
        arr(i, 1) = (i Mod 5000) * 7
    Next i
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(N, 1))
    rng.NumberFormat = "#,##0"
    rng.Value2 = arr

    Dim t0 As Single
    Dim v2 As Variant
    t0 = Timer
    v2 = rng.Value2
    LogShim.LogEvent "stage", Array("op", "ms"), Array("value2", CLng((Timer - t0) * 1000!))

    ' 단계 1: 균일성 판정
    t0 = Timer
    Dim vFmt As Variant
    vFmt = rng.NumberFormatLocal
    LogShim.LogEvent "stage", Array("op", "ms"), Array("fmtlocal_range", CLng((Timer - t0) * 1000!))

    ' 단계 2: 셀별 .Text 루프 (RenderKeys 비 * 경로 복제)
    t0 = Timer
    Dim keys() As String
    ReDim keys(1 To N)
    Dim c As Range, f As String
    For i = 1 To N
        If IsEmpty(v2(i, 1)) Then
            keys(i) = ""
        Else
            Set c = rng.Cells(i, 1)
            f = CStr(vFmt)
            If InStr(f, "*") > 0 And VarType(v2(i, 1)) = vbDouble Then
                ' (해당 없음)
            Else
                keys(i) = c.Text
            End If
        End If
    Next i
    LogShim.LogEvent "stage", Array("op", "ms"), Array("renderloop", CLng((Timer - t0) * 1000!))

    ' 단계 3: displays + rowByKey만
    t0 = Timer
    Dim displays As New clsSet
    Dim rowByKey As Object
    Set rowByKey = CreateObject("Scripting.Dictionary")
    rowByKey.CompareMode = vbTextCompare
    Dim k As String
    For i = 1 To N
        k = keys(i)
        If Len(k) > 0 Then
            displays.Add k
            If rowByKey.exists(k) Then
                rowByKey(k) = rowByKey(k) + 1
            Else
                rowByKey(k) = 1
            End If
        End If
    Next i
    LogShim.LogEvent "stage", Array("op", "ms"), Array("sets", CLng((Timer - t0) * 1000!))

    ' 단계 4: ValueMap만
    t0 = Timer
    Dim vm As Object
    Set vm = CreateObject("Scripting.Dictionary")
    Dim vk As String
    For i = 1 To N
        If Not IsEmpty(v2(i, 1)) Then
            If VarType(v2(i, 1)) = vbDouble Then
                vk = CStr(v2(i, 1))
                If Not vm.exists(vk) Then Set vm(vk) = New clsSet
                vm(vk).Add keys(i)
            End If
        End If
    Next i
    LogShim.LogEvent "stage", Array("op", "ms"), Array("valuemap", CLng((Timer - t0) * 1000!))

Done:
    Utilities.OptimizedMode False
    DropSheet
    LogShim.LogEvent "perf_done", Array("ok"), Array(Err.Number = 0)
End Sub

' @desc 지문 신선도 경로(적중/미스/대조/조기탈출)의 접근당 비용을 로그로 기록함
Public Sub RunFingerprintBench()
    LogShim.Reset
    Utilities.OptimizedMode True
    On Error GoTo Done

    BenchSize "comma_10k", "#,##0", 10000, False
    BenchSize "comma_50k", "#,##0", 50000, False
    BenchSize "comma_100k", "#,##0", 100000, False
    BenchSize "blockmix_50k", "", 50000, True

Done:
    Utilities.OptimizedMode False
    DropSheet
    LogShim.LogEvent "perf_done", Array("ok"), Array(Err.Number = 0)
End Sub

' 지정 열에서 미스(재빌드)/적중(지문+대조)/ComputeSignature/Matches 단가를 REP회 평균으로 잰다.
Private Sub BenchSize(ByVal tag As String, ByVal numFmt As String, _
                      ByVal nRows As Long, ByVal blockMixed As Boolean)
    Const REP As Long = 200
    Dim ws As Worksheet
    Set ws = PrepSheet()

    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To 1)
    Dim i As Long
    For i = 1 To nRows
        arr(i, 1) = (i Mod 100) * 7      ' 저카디널리티(실사용 필터 열)
    Next i
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, 1))
    rng.Value2 = arr
    If blockMixed Then
        ' 3블록: 전반 #,##0 / 중반 0 / 후반 $#,##0 (ComputeSignature가 TryPartition 수행)
        Dim third As Long
        third = nRows \ 3
        ws.Range(ws.Cells(1, 1), ws.Cells(third, 1)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(third + 1, 1), ws.Cells(2 * third, 1)).NumberFormat = "0"
        ws.Range(ws.Cells(2 * third + 1, 1), ws.Cells(nRows, 1)).NumberFormat = "$#,##0"
    Else
        rng.NumberFormat = numFmt
    End If

    Dim t0 As Single, j As Long
    Dim asset As clsColAsset
    Dim sig As clsColSig, s1 As clsColSig, s2 As clsColSig, s3 As clsColSig
    Dim b As Boolean

    ' 1) 미스(콜드): ComputeSignature + BuildAssetFromSig 전체 재빌드
    modCache.ClearCache
    t0 = Timer
    Set asset = modCache.GetColumnAsset(rng)
    BenchLog tag, "miss_rebuild", nRows, (Timer - t0) * 1000!, 1

    ' 2) 적중(웜): ComputeSignature + Matches (렌더 없음) - 신 설계의 접근당 비용
    t0 = Timer
    For j = 1 To REP
        Set asset = modCache.GetColumnAsset(rng)
    Next j
    BenchLog tag, "hit_getasset", nRows, (Timer - t0) * 1000!, REP

    ' 3) ComputeSignature 단독 (적중 비용의 지배 항: Value2 벌크 + 서식 판정)
    t0 = Timer
    For j = 1 To REP
        Set sig = modColumnScan.ComputeSignature(rng)
    Next j
    BenchLog tag, "computesig", nRows, (Timer - t0) * 1000!, REP

    ' 4) Matches 단독(동일 지문): ContentVec 원소별 대조 전량 통과
    Set s1 = modColumnScan.ComputeSignature(rng)
    Set s2 = modColumnScan.ComputeSignature(rng)
    t0 = Timer
    For j = 1 To REP
        b = s1.Matches(s2)
    Next j
    BenchLog tag, "matches_same", nRows, (Timer - t0) * 1000!, REP

    ' 5) Matches 조기탈출(1행만 편집): 첫 원소 불일치로 즉시 이탈
    ws.Cells(1, 1).Value = CDbl(ws.Cells(1, 1).Value) + 1
    Set s3 = modColumnScan.ComputeSignature(rng)
    t0 = Timer
    For j = 1 To REP
        b = s1.Matches(s3)
    Next j
    BenchLog tag, "matches_earlyexit", nRows, (Timer - t0) * 1000!, REP
End Sub

Private Sub BenchLog(ByVal tag As String, ByVal metric As String, ByVal nRows As Long, _
                     ByVal totalMs As Double, ByVal rep As Long)
    Dim perOp As Double
    perOp = totalMs / rep
    LogShim.LogEvent "fpbench", _
        Array("case", "metric", "rows", "totalMs", "rep", "perOpMs"), _
        Array(tag, metric, nRows, Format$(totalMs, "0.0"), rep, Format$(perOp, "0.000"))
End Sub

' @desc 값 레인 COM 비용을 미스 수/블록 수 함수로 계측함(관측 (1) 검증 프로브)
' modMatch.ExecuteScan은 미스 숫자 토큰마다 nCands x nBlocks회 RenderValueKey(COM
' WorksheetFunction.Text)를 돈다. numeric_miss(값 레인 전부 가동) vs text_miss(값 레인
' 미가동, COM 0회)의 동일 토큰수 델타가 값 레인 순비용이다. 3블록 혼합 열은 COM 3배 증폭.
Public Sub RunValueLaneProbe()
    LogShim.Reset
    Utilities.OptimizedMode True
    On Error GoTo Done

    MeasureValueLane "uniform_1blk", False, 100000
    MeasureValueLane "mixed_3blk", True, 90000

Done:
    Utilities.OptimizedMode False
    DropSheet
    LogShim.LogEvent "perf_done", Array("ok"), Array(Err.Number = 0)
End Sub

' 지정 열에서 토큰 종류(numeric_miss/text_miss/hit) x 토큰수별 Execute 소요를 잰다.
Private Sub MeasureValueLane(ByVal tag As String, ByVal blockMixed As Boolean, ByVal nRows As Long)
    Dim ws As Worksheet
    Set ws = PrepSheet()

    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To 1)
    Dim i As Long
    For i = 1 To nRows
        arr(i, 1) = (i Mod 5000) * 7          ' 저카디널리티(실사용 필터 열)
    Next i
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, 1))
    rng.Value2 = arr

    Dim nBlk As Long
    If blockMixed Then
        Dim third As Long
        third = nRows \ 3
        ws.Range(ws.Cells(1, 1), ws.Cells(third, 1)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(third + 1, 1), ws.Cells(2 * third, 1)).NumberFormat = "0"
        ws.Range(ws.Cells(2 * third + 1, 1), ws.Cells(nRows, 1)).NumberFormat = "$#,##0"
        nBlk = 3
    Else
        rng.NumberFormat = "#,##0"
        nBlk = 1
    End If

    modCache.ClearCache
    Dim asset As clsColAsset
    Set asset = modCache.GetColumnAsset(rng)   ' 콜드 스캔 1회(측정 밖) - 자산 확보

    Dim sizes As Variant
    sizes = Array(100, 500, 1000, 2000)
    Dim s As Long, n As Long
    For s = LBound(sizes) To UBound(sizes)
        n = CLng(sizes(s))
        MeasureExec tag, nBlk, "numeric_miss", n, BuildNumericMissTokens(n), asset
        MeasureExec tag, nBlk, "text_miss", n, BuildTextMissTokens(n), asset
        If Not blockMixed Then
            MeasureExec tag, nBlk, "hit", n, BuildHitTokens(n), asset
        End If
    Next s
End Sub

' 토큰 배열로 Execute(완전일치)를 REP회 돌려 접근당 소요와 매칭수를 로그한다.
Private Sub MeasureExec(ByVal tag As String, ByVal nBlk As Long, ByVal kind As String, _
                        ByVal n As Long, ByVal tokens As Variant, ByVal asset As clsColAsset)
    Const REP As Long = 3
    Dim t0 As Single, j As Long
    Dim outcome As Object
    t0 = Timer
    For j = 1 To REP
        Set outcome = modMatch.Execute(tokens, asset, False)
    Next j
    Dim totalMs As Double
    totalMs = (Timer - t0) * 1000!
    Dim matched As clsSet
    Set matched = outcome("Matched")
    LogShim.LogEvent "vlane", _
        Array("case", "blocks", "kind", "tokens", "perOpMs", "matched"), _
        Array(tag, nBlk, kind, n, Format$(totalMs / REP, "0.0"), matched.Count)
End Sub

' 숫자꼴이지만 컬럼(<=34,993)에 부재한 대형 수 -> 값 레인 전부 미스(COM 최대 가동).
Private Function BuildNumericMissTokens(ByVal n As Long) As Variant
    Dim a() As Variant
    ReDim a(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        a(i) = CStr(1000000001# + i)
    Next i
    BuildNumericMissTokens = a
End Function

' 숫자 파싱 불가 텍스트 -> 값 레인 미가동(COM 0회). 동일 토큰수 numeric_miss의 대조군.
Private Function BuildTextMissTokens(ByVal n As Long) As Variant
    Dim a() As Variant
    ReDim a(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        a(i) = "ZZ" & i
    Next i
    BuildTextMissTokens = a
End Function

' 실제 표시값(#,##0 렌더) -> 표시 레인 적중(값 레인 미도달). pass1 스캔 순비용 기준선.
Private Function BuildHitTokens(ByVal n As Long) As Variant
    Dim a() As Variant
    ReDim a(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        a(i) = Application.WorksheetFunction.Text((i + 1) * 7, "#,##0")
    Next i
    BuildHitTokens = a
End Function

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
    Set PrepSheet = ws
End Function

Private Sub DropSheet()
    Dim prev As Boolean
    prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = prev
End Sub
