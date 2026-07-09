Attribute VB_Name = "modPerfProbe"
' @desc 스캔 병목 계측 일회용 프로브 담당
' @depends modColumnScan, modCache, Utilities, LogShim
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
    MeasureMixedFormat 50000
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

    modCache.InvalidateCacheBySheet ws

    Dim t0 As Single
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rng, "PERF", 0)
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

    modCache.InvalidateCacheBySheet ws
    Dim t0 As Single
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rng, "PERF", 0)
    LogShim.LogEvent "perf_buildasset", Array("fmt", "rows", "ms", "displays"), _
        Array("uniq100pct", nRows, CLng((Timer - t0) * 1000!), asset.Displays.Count)
End Sub

' 혼합 서식 열(행마다 다른 서식): 셀별 폴백의 실제 비용.
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

    modCache.InvalidateCacheBySheet ws
    Dim t0 As Single
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rng, "PERF", 0)
    LogShim.LogEvent "perf_buildasset", Array("fmt", "rows", "ms", "displays"), _
        Array("mixed_fmt", nRows, CLng((Timer - t0) * 1000!), asset.Displays.Count)
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
