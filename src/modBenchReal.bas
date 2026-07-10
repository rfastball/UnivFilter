Attribute VB_Name = "modBenchReal"
' @desc 실데이터 열을 격리 fixture로 복제하고 스캔경로 접근당 비용 계측 담당
' @depends modCache, modColumnScan, modFilterRun, modInputParse, modMatch, modApply, modReport, clsColAsset, clsColSig, clsSet, frmFilter01, Utilities, LogShim
Option Explicit

' 실데이터 원본(직접 벤치 금지 - 데이터만 뽑아 fixture로 복제 후 그 위에서 측정).
' (Const는 상수식만 허용 - ChrW 조립은 함수로.)
Private Function RealPath() As String
    ' ...\지출 리팩토링\지출원인행위정보 송신 목록.xlsm
    RealPath = "C:\Users\rfast\Desktop\VBA_Projects\" & _
        ChrW(&HC9C0) & ChrW(&HCD9C) & " " & _
        ChrW(&HB9AC) & ChrW(&HD329) & ChrW(&HD1A0) & ChrW(&HB9C1) & "\" & _
        ChrW(&HC9C0) & ChrW(&HCD9C) & ChrW(&HC6D0) & ChrW(&HC778) & ChrW(&HD589) & _
        ChrW(&HC704) & ChrW(&HC815) & ChrW(&HBCF4) & " " & _
        ChrW(&HC1A1) & ChrW(&HC2E0) & " " & ChrW(&HBAA9) & ChrW(&HB85D) & ".xlsm"
End Function

Private Function FixturePath() As String
    FixturePath = Environ$("TEMP") & "\uf_bench_fixture.xlsx"
End Function

' 원본 시트명(코드에 한글 상수로 두면 인코딩 취약 - ChrW로 조립).
Private Function SheetDW() As String
    SheetDW = "DW"
End Function
Private Function SheetSend() As String
    ' 지출원인행위송신목록
    SheetSend = ChrW(&HC9C0) & ChrW(&HCD9C) & ChrW(&HC6D0) & ChrW(&HC778) & _
                ChrW(&HD589) & ChrW(&HC704) & ChrW(&HC1A1) & ChrW(&HC2E0) & _
                ChrW(&HBAA9) & ChrW(&HB85D)
End Function

' @desc 실데이터 대표 열들을 값/서식 그대로 임시 fixture 워크북으로 복제함
Public Sub ExtractFixture()
    Dim prevEv As Boolean, prevAlerts As Boolean, prevSec As Long
    prevEv = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevSec = Application.AutomationSecurity
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AutomationSecurity = 3    ' msoAutomationSecurityForceDisable (원본 매크로 차단)

    LogShim.Reset
    On Error GoTo Fail

    Dim src As Workbook
    Set src = Workbooks.Open(RealPath(), ReadOnly:=True, UpdateLinks:=0)

    Dim dst As Workbook
    Set dst = Workbooks.Add
    Dim wsDW As Worksheet, wsT As Worksheet
    Set wsDW = dst.Sheets(1): wsDW.Name = "DW"
    Set wsT = dst.Sheets.Add(After:=wsDW): wsT.Name = "T"

    ' DW(20만행): 고카디널리티 금액 / 저카디널리티 카운트 / 텍스트 / 긴코드 / 날짜숫자 / 오류셀
    CopyCol src.Sheets(SheetDW), "P", wsDW, 1
    CopyCol src.Sheets(SheetDW), "Z", wsDW, 2
    CopyCol src.Sheets(SheetDW), "D", wsDW, 3
    CopyCol src.Sheets(SheetDW), "M", wsDW, 4
    CopyCol src.Sheets(SheetDW), "C", wsDW, 5
    CopyCol src.Sheets(SheetDW), "AC", wsDW, 6

    ' 송신목록(2.8만행): 전열 텍스트(@) - 금액텍스트 / 결과텍스트
    CopyCol src.Sheets(SheetSend), "V", wsT, 1
    CopyCol src.Sheets(SheetSend), "AH", wsT, 2

    src.Close SaveChanges:=False

    On Error Resume Next
    Kill FixturePath()
    On Error GoTo Fail
    dst.SaveAs FixturePath(), FileFormat:=51    ' xlOpenXMLWorkbook (.xlsx)
    dst.Close SaveChanges:=False

    LogShim.LogEvent "fixture", Array("path", "ok"), Array(FixturePath(), True)

Cleanup:
    Application.EnableEvents = prevEv
    Application.DisplayAlerts = prevAlerts
    Application.AutomationSecurity = prevSec
    Exit Sub
Fail:
    LogShim.LogEvent "fixture_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Resume Cleanup
End Sub

' 원본 열 1개를 값(타입/오류 충실) + 균일서식으로 dst 열에 복제.
Private Sub CopyCol(ByVal wsSrc As Worksheet, ByVal srcCol As String, _
                    ByVal wsDst As Worksheet, ByVal dstCol As Long)
    On Error GoTo Fail
    Dim lastRow As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, srcCol).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    Dim rSrc As Range, rDst As Range
    Set rSrc = wsSrc.Range(wsSrc.Cells(1, srcCol), wsSrc.Cells(lastRow, srcCol))
    Set rDst = wsDst.Range(wsDst.Cells(1, dstCol), wsDst.Cells(lastRow, dstCol))

    rDst.Value2 = rSrc.Value2

    Dim f As Variant
    f = rSrc.NumberFormatLocal
    If IsNull(f) Then f = wsSrc.Cells(2, srcCol).NumberFormatLocal
    rDst.NumberFormatLocal = CStr(f)

    LogShim.LogEvent "copycol", Array("col", "rows", "fmt"), _
        Array(srcCol, lastRow, CStr(f))
    Exit Sub
Fail:
    LogShim.LogEvent "copycol_err", Array("col", "num", "desc"), _
        Array(srcCol, Err.Number, Err.Description)
End Sub

' @desc fixture 각 열에서 미스/적중/지문/대조 접근당 비용을 로그로 기록함
Public Sub BenchReal()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())    ' 읽기쓰기(AutoFit 재독 경로 유효)
    Utilities.OptimizedMode True

    BenchColumn wb.Sheets("DW"), 1, "DW_P_amount_hicard"
    BenchColumn wb.Sheets("DW"), 2, "DW_Z_count_locard"
    BenchColumn wb.Sheets("DW"), 3, "DW_D_text"
    BenchColumn wb.Sheets("DW"), 4, "DW_M_longcode"
    BenchColumn wb.Sheets("DW"), 5, "DW_C_datenum"
    BenchColumn wb.Sheets("DW"), 6, "DW_AC_amount_err"
    BenchColumn wb.Sheets("T"), 1, "SL_V_amtText"
    BenchColumn wb.Sheets("T"), 2, "SL_AH_result"

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    LogShim.LogEvent "bench_done", Array("ok"), Array(True)
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "bench_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' @desc fixture 고유값 전수로 WText==현행표시키 일치와 인프로세스 렌더 소요를 기록함
Public Sub MeasureWTextOracle()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    OracleCol wb.Sheets("DW"), 1, "DW_P_amount"
    OracleCol wb.Sheets("DW"), 2, "DW_Z_count"
    OracleCol wb.Sheets("DW"), 5, "DW_C_datenum"

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    LogShim.LogEvent "oracle_done", Array("ok"), Array(True)
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "oracle_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' 한 열의 고유 숫자값 전수에 대해 WText 렌더를 현행 엔진 표시키(ValueMap)와 대조한다.
Private Sub OracleCol(ByVal ws As Worksheet, ByVal col As Long, ByVal tag As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    Dim body As Range
    Set body = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))

    ' 현행 엔진의 산출물 = 오라클 (ValueMap: CStr(Value2) -> 표시키)
    modCache.ClearCache
    Dim asset As clsColAsset
    Set asset = modCache.GetColumnAsset(body)
    Dim vm As Object
    Set vm = asset.ValueMap
    If vm.Count = 0 Then
        LogShim.LogEvent "oracle", Array("tag", "uniq", "note"), Array(tag, 0, "no doubles")
        Exit Sub
    End If

    ' 렌더 서식 결정 (modColumnScan.RenderFmt 규칙 복제: General -> 영문, 그 외 ! 제거)
    Dim fmt As String
    fmt = CStr(body.NumberFormatLocal)
    Dim isGeneral As Boolean
    isGeneral = (fmt = "General") Or (fmt = "G/" & ChrW(&HD45C) & ChrW(&HC900))
    Dim useFmt As String
    If isGeneral Then useFmt = "General" Else useFmt = Replace(fmt, "!", "")

    ' 전수 대조 + WText 소요
    Dim t0 As Single, nMis As Long, misSample As String
    Dim vKey As Variant, wt As String
    t0 = Timer
    For Each vKey In vm.Keys
        wt = SafeWTextB(CDbl(vKey), useFmt)
        ' ValueMap 값은 vbNullChar 연결 - 균일서식이면 통상 1키
        If InStr(1, vbNullChar & vm(vKey) & vbNullChar, vbNullChar & wt & vbNullChar) = 0 Then
            nMis = nMis + 1
            If nMis <= 3 Then misSample = misSample & "[" & CStr(vKey) & "->" & _
                                          wt & " vs " & vm(vKey) & "]"
        End If
    Next vKey
    Dim wtMs As Double
    wtMs = (Timer - t0) * 1000!

    ' 현행 .Text 경로 대비용: 같은 고유값 수만큼 셀당 .Text (첫 등장행 무관, 단가 비교 목적)
    Dim n As Long, i As Long, s As String
    n = vm.Count
    If n > lastRow - 1 Then n = lastRow - 1
    t0 = Timer
    For i = 1 To n
        s = body.Cells(i, 1).Text
    Next i
    Dim txMs As Double
    txMs = (Timer - t0) * 1000!

    LogShim.LogEvent "oracle", _
        Array("tag", "uniq", "fmt", "mismatch", "sample", "wtextMs", "dotTextMs"), _
        Array(tag, vm.Count, useFmt, nMis, misSample, _
              Format$(wtMs, "0.0"), Format$(txMs, "0.0"))
End Sub

' WorksheetFunction.Text 안전 시도(실패 시 빈 문자열) - modColumnScan.SafeWText 복제.
Private Function SafeWTextB(ByVal v As Variant, ByVal sFmt As String) As String
    On Error GoTo Fail
    SafeWTextB = Application.WorksheetFunction.Text(v, sFmt)
    Exit Function
Fail:
    SafeWTextB = ""
End Function

' @desc DW_P 고카디널 열에서 RenderBlock 내부 단계별 소요를 복제 계측해 기록함
Public Sub MeasureRenderStages()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    Dim ws As Worksheet
    Set ws = wb.Sheets("DW")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim body As Range
    Set body = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
    Dim nRows As Long
    nRows = lastRow - 1

    Dim t0 As Single

    ' 0) Value2 벌크
    t0 = Timer
    Dim v2 As Variant
    v2 = body.Value2
    StageLog "value2", nRows, (Timer - t0) * 1000!

    ' 1) pass1: 타입태그 dedup (RenderBlock 134-157 복제)
    t0 = Timer
    Dim uniqFirst As Object, uniqCnt As Object
    Set uniqFirst = CreateObject("Scripting.Dictionary")
    Set uniqCnt = CreateObject("Scripting.Dictionary")
    Dim blankCount As Long, i As Long, tag As String
    For i = 1 To nRows
        If IsEmpty(v2(i, 1)) Then
            blankCount = blankCount + 1
        ElseIf IsError(v2(i, 1)) Then
            ' (해당 열 없음)
        Else
            tag = CStr(VarType(v2(i, 1))) & Chr$(1) & CStr(v2(i, 1))
            If uniqCnt.exists(tag) Then
                uniqCnt(tag) = uniqCnt(tag) + 1
            Else
                uniqFirst(tag) = i
                uniqCnt(tag) = 1
            End If
        End If
    Next i
    StageLog "pass1_dedup", uniqFirst.Count, (Timer - t0) * 1000!

    ' 2) pass2a: 고유값당 .Text 렌더 (현행 경로)
    t0 = Timer
    Dim keyByTag As Object
    Set keyByTag = CreateObject("Scripting.Dictionary")
    Dim vt As Variant, key As String
    For Each vt In uniqFirst.Keys
        key = body.Cells(uniqFirst(vt), 1).Text
        keyByTag(vt) = key
    Next vt
    StageLog "pass2_dotText", keyByTag.Count, (Timer - t0) * 1000!

    ' 2') pass2b: 고유값당 WText 렌더 (제안 경로)
    t0 = Timer
    Dim keyByTag2 As Object
    Set keyByTag2 = CreateObject("Scripting.Dictionary")
    For Each vt In uniqFirst.Keys
        keyByTag2(vt) = SafeWTextB(v2(uniqFirst(vt), 1), "#,##0")
    Next vt
    StageLog "pass2_WText", keyByTag2.Count, (Timer - t0) * 1000!

    ' 3) 누적: rowByKey(vbTextCompare) + AppendVm (RenderBlock 200-206 복제)
    t0 = Timer
    Dim rowByKey As Object
    Set rowByKey = CreateObject("Scripting.Dictionary")
    rowByKey.CompareMode = vbTextCompare
    Dim vm As Object
    Set vm = CreateObject("Scripting.Dictionary")
    Dim val As Variant
    For Each vt In uniqFirst.Keys
        key = keyByTag(vt)
        rowByKey(key) = rowByKey(key) + uniqCnt(vt)
        val = v2(uniqFirst(vt), 1)
        If VarType(val) = vbDouble Then AppendVmB vm, CStr(val), key
    Next vt
    StageLog "accumulate", rowByKey.Count, (Timer - t0) * 1000!

    ' 4) displays 구축 (BuildAssetFromSig 115 복제)
    t0 = Timer
    Dim displays As New clsSet
    displays.LoadFromArray rowByKey.Keys
    StageLog "loadfromarray", displays.Count, (Timer - t0) * 1000!

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    LogShim.LogEvent "stages_done", Array("ok"), Array(True)
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "stages_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' @desc Evaluate(TEXT) 벌크 렌더 1회호출의 소요/정합을 현행 자산과 대조해 기록함
Public Sub MeasureBulkEval()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    BulkEvalCol wb.Sheets("DW"), 1, "#,##0", "DW_P_amount"
    BulkEvalCol wb.Sheets("DW"), 5, "General", "DW_C_datenum"
    BulkEvalCol wb.Sheets("DW"), 4, "General", "DW_M_longcode_blanks"

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    LogShim.LogEvent "bulkeval_done", Array("ok"), Array(True)
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "bulkeval_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

Private Sub BulkEvalCol(ByVal ws As Worksheet, ByVal col As Long, _
                        ByVal engFmt As String, ByVal tag As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    Dim body As Range
    Set body = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))
    Dim nRows As Long
    nRows = lastRow - 1

    Dim v2 As Variant
    v2 = body.Value2

    ' 1) Evaluate(TEXT) 벌크 렌더 - COM 1회
    Dim t0 As Single
    t0 = Timer
    Dim rendered As Variant
    rendered = ws.Evaluate("TEXT(" & body.Address & ",""" & engFmt & """)")
    Dim evalMs As Double
    evalMs = (Timer - t0) * 1000!

    Dim evalOk As Boolean
    evalOk = IsArray(rendered)
    If Not evalOk Then
        LogShim.LogEvent "bulkeval", Array("tag", "note"), Array(tag, "Evaluate 비배열 반환")
        Exit Sub
    End If

    ' 2) 빈칸/에러 거동 관찰
    Dim blankRender As String, nErrCells As Long, i As Long
    blankRender = "(none)"
    For i = 1 To nRows
        If IsEmpty(v2(i, 1)) Then
            If blankRender = "(none)" Then
                If IsError(rendered(i, 1)) Then
                    blankRender = "(error)"
                Else
                    blankRender = "[" & CStr(rendered(i, 1)) & "]"
                End If
            End If
        ElseIf IsError(v2(i, 1)) Then
            nErrCells = nErrCells + 1
        End If
    Next i

    ' 3) 신경로 시뮬레이션: 단일 루프 dict 누적 (IsEmpty 가드, 에러셀 .Text 폴백)
    t0 = Timer
    Dim rowByKey As Object
    Set rowByKey = CreateObject("Scripting.Dictionary")
    rowByKey.CompareMode = vbTextCompare
    Dim vm As Object
    Set vm = CreateObject("Scripting.Dictionary")
    Dim k As String
    For i = 1 To nRows
        If Not IsEmpty(v2(i, 1)) Then
            If IsError(v2(i, 1)) Then
                k = body.Cells(i, 1).Text          ' 에러셀(드묾): 셀별 폴백
            ElseIf IsError(rendered(i, 1)) Then
                k = body.Cells(i, 1).Text
            Else
                k = CStr(rendered(i, 1))
            End If
            If Len(k) > 0 Then
                rowByKey(k) = rowByKey(k) + 1
                If VarType(v2(i, 1)) = vbDouble Then AppendVmB vm, CStr(v2(i, 1)), k
            End If
        End If
    Next i
    Dim loopMs As Double
    loopMs = (Timer - t0) * 1000!

    ' 4) displays 구축
    t0 = Timer
    Dim displays As New clsSet
    displays.LoadFromArray rowByKey.Keys
    Dim setMs As Double
    setMs = (Timer - t0) * 1000!

    ' 5) 정합 대조: 현행 엔진 자산과 우주/행수지도 완전 동치?
    modCache.ClearCache
    Dim asset As clsColAsset
    Set asset = modCache.GetColumnAsset(body)
    Dim nMis As Long, misSample As String
    Dim vKey As Variant
    If asset.Displays.Count <> displays.Count Then
        nMis = -1    ' 우주 크기부터 불일치
        misSample = "count " & displays.Count & " vs " & asset.Displays.Count
    Else
        For Each vKey In asset.RowCountByKey.Keys
            If Not rowByKey.exists(vKey) Then
                nMis = nMis + 1
                If nMis <= 3 Then misSample = misSample & "[miss:" & CStr(vKey) & "]"
            ElseIf rowByKey(vKey) <> asset.RowCountByKey(vKey) Then
                nMis = nMis + 1
                If nMis <= 3 Then misSample = misSample & "[cnt:" & CStr(vKey) & "]"
            End If
        Next vKey
    End If

    LogShim.LogEvent "bulkeval", _
        Array("tag", "rows", "uniq", "evalMs", "loopMs", "setMs", "newTotalMs", _
              "blankRender", "errCells", "mismatch", "sample"), _
        Array(tag, nRows, displays.Count, Format$(evalMs, "0.0"), Format$(loopMs, "0.0"), _
              Format$(setMs, "0.0"), Format$(evalMs + loopMs + setMs, "0.0"), _
              blankRender, nErrCells, nMis, misSample)
End Sub

' @desc 렌더배열 직접 멤버십 스캔(무dict 홧경로)의 단가를 토큰수별로 기록함
Public Sub MeasureMembershipScan()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    Dim ws As Worksheet
    Set ws = wb.Sheets("DW")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim body As Range
    Set body = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
    Dim nRows As Long
    nRows = lastRow - 1

    Dim v2 As Variant
    v2 = body.Value2
    Dim rendered As Variant
    rendered = ws.Evaluate("TEXT(" & body.Address & ",""#,##0"")")

    ' 토큰 1개 (StrComp 직비교) + 빈칸 계수 동시
    Dim tok1 As String
    tok1 = CStr(rendered(nRows \ 2, 1))
    Dim t0 As Single, i As Long
    Dim hit1 As Long, blanks As Long
    t0 = Timer
    For i = 1 To nRows
        If IsEmpty(v2(i, 1)) Then
            blanks = blanks + 1
        ElseIf StrComp(CStr(rendered(i, 1)), tok1, vbTextCompare) = 0 Then
            hit1 = hit1 + 1
        End If
    Next i
    LogShim.LogEvent "memscan", Array("mode", "rows", "hits", "ms"), _
        Array("token1_strcomp", nRows, hit1, Format$((Timer - t0) * 1000!, "0.0"))

    ' 토큰 20개 (소형 dict 조회) - 다건 붙여넣기 모사
    Dim toks As Object
    Set toks = CreateObject("Scripting.Dictionary")
    toks.CompareMode = vbTextCompare
    For i = 1 To 20
        toks(CStr(rendered(i * 137, 1))) = 0&
    Next i
    Dim hitN As Long
    t0 = Timer
    For i = 1 To nRows
        If Not IsEmpty(v2(i, 1)) Then
            If toks.exists(CStr(rendered(i, 1))) Then
                toks(CStr(rendered(i, 1))) = toks(CStr(rendered(i, 1))) + 1
                hitN = hitN + 1
            End If
        End If
    Next i
    LogShim.LogEvent "memscan", Array("mode", "rows", "hits", "ms"), _
        Array("token20_dict", nRows, hitN, Format$((Timer - t0) * 1000!, "0.0"))

    ' 값 브리지 대체: 파싱값 -> WText 1회 -> 멤버십 (토큰1 경로와 동일 단가 확인용)
    Dim cand As String
    cand = SafeWTextB(v2(nRows \ 3, 1), "#,##0")
    Dim hitB As Long
    t0 = Timer
    For i = 1 To nRows
        If Not IsEmpty(v2(i, 1)) Then
            If StrComp(CStr(rendered(i, 1)), cand, vbTextCompare) = 0 Then hitB = hitB + 1
        End If
    Next i
    LogShim.LogEvent "memscan", Array("mode", "rows", "hits", "ms"), _
        Array("valuebridge_wtext1", nRows, hitB, Format$((Timer - t0) * 1000!, "0.0"))

    ' 대비: 지연 dict 구축(제외/부분일치 최초 요구 시 1회 지불할 비용)
    t0 = Timer
    Dim uniq As Object
    Set uniq = CreateObject("Scripting.Dictionary")
    uniq.CompareMode = vbTextCompare
    Dim k As String
    For i = 1 To nRows
        If Not IsEmpty(v2(i, 1)) Then
            k = CStr(rendered(i, 1))
            uniq(k) = uniq(k) + 1
        End If
    Next i
    LogShim.LogEvent "memscan", Array("mode", "rows", "hits", "ms"), _
        Array("lazy_dict_build", nRows, uniq.Count, Format$((Timer - t0) * 1000!, "0.0"))

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    LogShim.LogEvent "memscan_done", Array("ok"), Array(True)
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "memscan_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' @desc 제외/부분일치 모드의 현행 e2e와 재설계 시뮬 비용을 열별로 기록함
Public Sub MeasureExcludePartial()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    ExclPartCol wb.Sheets("DW"), 6, 1, "#,##0", "DW_P_hicard"
    ExclPartCol wb.Sheets("DW"), 6, 3, "General", "DW_D_text_locard"

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    LogShim.LogEvent "exclpart_done", Array("ok"), Array(True)
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "exclpart_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

Private Sub ExclPartCol(ByVal ws As Worksheet, ByVal nCols As Long, _
                        ByVal field As Long, ByVal engFmt As String, ByVal tag As String)
    On Error Resume Next
    ws.ShowAllData
    ws.AutoFilterMode = False
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, field).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    Dim area As Range, body As Range
    Set area = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, nCols))
    Set body = ws.Range(ws.Cells(2, field), ws.Cells(lastRow, field))
    Dim nRows As Long
    nRows = lastRow - 1
    area.AutoFilter

    Dim tok As String
    tok = FirstDisplay(body)
    ' 부분일치 토큰: 첫 표시값의 중간 2자(다건 매칭 유도)
    Dim ptok As String
    If Len(tok) >= 3 Then ptok = Mid$(tok, 2, 2) Else ptok = tok

    Dim t0 As Single, r As Object

    ' ===== 제외 모드 =====
    modApply.ResetFilters area
    modCache.ClearCache
    t0 = Timer
    Set r = modFilterRun.RunFilter(body, area, field, tok, True, False, False)
    Dim exclCold As Double
    exclCold = (Timer - t0) * 1000!

    modApply.ResetFilters area
    t0 = Timer
    Set r = modFilterRun.RunFilter(body, area, field, tok, True, False, False)
    Dim exclWarm As Double
    exclWarm = (Timer - t0) * 1000!

    ' 웜 분해: Difference / ExpectedRows / ApplyDisplays(여집합 배열)
    Dim asset As clsColAsset
    Set asset = modCache.GetColumnAsset(body)
    Dim matched As New clsSet
    matched.Add tok
    t0 = Timer
    Dim comp As clsSet
    Set comp = asset.Displays.Difference(matched)
    Dim diffMs As Double
    diffMs = (Timer - t0) * 1000!

    t0 = Timer
    Dim expRows As Long
    expRows = asset.ExpectedRows(comp, asset.HasBlanks)
    Dim expMs As Double
    expMs = (Timer - t0) * 1000!

    modApply.ResetFilters area
    t0 = Timer
    Dim ai As Object
    Set ai = modApply.ApplyDisplays(area, field, comp, asset.HasBlanks, False)
    Dim applyMs As Double
    applyMs = (Timer - t0) * 1000!

    LogShim.LogEvent "excl", _
        Array("tag", "uniq", "compKeys", "coldMs", "warmMs", "diffMs", "expectedMs", "applyMs"), _
        Array(tag, asset.Displays.Count, comp.Count, Format$(exclCold, "0.0"), _
              Format$(exclWarm, "0.0"), Format$(diffMs, "0.00"), _
              Format$(expMs, "0.00"), Format$(applyMs, "0.0"))

    ' ===== 부분일치 모드 =====
    modApply.ResetFilters area
    modCache.ClearCache
    t0 = Timer
    Set r = modFilterRun.RunFilter(body, area, field, ptok, False, False, True)
    Dim partCold As Double
    partCold = (Timer - t0) * 1000!

    modApply.ResetFilters area
    t0 = Timer
    Set r = modFilterRun.RunFilter(body, area, field, ptok, False, False, True)
    Dim partWarm As Double
    partWarm = (Timer - t0) * 1000!

    ' 웜 분해: modMatch 전수 InStr (고유키 KeysArray)
    Set asset = modCache.GetColumnAsset(body)
    t0 = Timer
    Dim outc As Object
    Set outc = modMatch.Execute(Array(ptok), asset, True)
    Dim matchMs As Double
    matchMs = (Timer - t0) * 1000!

    LogShim.LogEvent "part", _
        Array("tag", "uniq", "ptok", "matchedKeys", "coldMs", "warmMs", "matchMs"), _
        Array(tag, asset.Displays.Count, ptok, outc("Matched").Count, _
              Format$(partCold, "0.0"), Format$(partWarm, "0.0"), Format$(matchMs, "0.0"))

    ' ===== 재설계 시뮬: 렌더배열 기반 =====
    Dim v2 As Variant
    v2 = body.Value2
    Dim rendered As Variant
    rendered = ws.Evaluate("TEXT(" & body.Address & ",""" & engFmt & """)")

    ' (a) 지연 uniq dict 구축 (제외/부분일치 최초 1회 지불)
    Dim i As Long, k As String
    t0 = Timer
    Dim uniq As Object
    Set uniq = CreateObject("Scripting.Dictionary")
    uniq.CompareMode = vbTextCompare
    For i = 1 To nRows
        If Not IsEmpty(v2(i, 1)) Then
            k = CStr(rendered(i, 1))
            uniq(k) = uniq(k) + 1
        End If
    Next i
    Dim lazyMs As Double
    lazyMs = (Timer - t0) * 1000!

    ' (b) 여집합 배열 구축 (matched 제외)
    t0 = Timer
    Dim compArr() As String, nComp As Long
    ReDim compArr(0 To uniq.Count - 1)
    Dim vKey As Variant
    For Each vKey In uniq.Keys
        If StrComp(CStr(vKey), tok, vbTextCompare) <> 0 Then
            compArr(nComp) = CStr(vKey)
            nComp = nComp + 1
        End If
    Next vKey
    Dim compBuildMs As Double
    compBuildMs = (Timer - t0) * 1000!

    ' (c) 부분일치: 렌더배열 91k행 직스캔(중복 히트 dict 흡수) vs uniq키 InStr
    t0 = Timer
    Dim hits As Object
    Set hits = CreateObject("Scripting.Dictionary")
    hits.CompareMode = vbTextCompare
    For i = 1 To nRows
        If Not IsEmpty(v2(i, 1)) Then
            k = CStr(rendered(i, 1))
            If InStr(1, k, ptok, vbTextCompare) > 0 Then hits(k) = hits(k) + 1
        End If
    Next i
    Dim partRowScanMs As Double
    partRowScanMs = (Timer - t0) * 1000!

    t0 = Timer
    Dim hits2 As Long
    For Each vKey In uniq.Keys
        If InStr(1, CStr(vKey), ptok, vbTextCompare) > 0 Then hits2 = hits2 + 1
    Next vKey
    Dim partUniqScanMs As Double
    partUniqScanMs = (Timer - t0) * 1000!

    LogShim.LogEvent "sim", _
        Array("tag", "lazyUniqMs", "compBuildMs", "partRowScanMs", "partUniqScanMs", _
              "partHits", "compKeys"), _
        Array(tag, Format$(lazyMs, "0.0"), Format$(compBuildMs, "0.0"), _
              Format$(partRowScanMs, "0.0"), Format$(partUniqScanMs, "0.0"), _
              hits.Count, nComp)
End Sub

' @desc 부분일치 콜드를 오염 없는 신선한 상태에서 단독 계측함(순서효과 격리)
Public Sub MeasurePartialColdOnly()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    Dim ws As Worksheet
    Set ws = wb.Sheets("DW")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim area As Range, body As Range
    Set area = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 6))
    Set body = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
    area.AutoFilter

    Dim ptok As String
    ptok = ",2"

    Dim t0 As Single, r As Object
    modCache.ClearCache
    t0 = Timer
    Set r = modFilterRun.RunFilter(body, area, 1, ptok, False, False, True)
    Dim coldMs As Double
    coldMs = (Timer - t0) * 1000!

    ' 같은 상태에서 가시행 계수만 단독 재측정 (조각난 필터 결과의 SpecialCells 비용)
    t0 = Timer
    Dim vis As Long
    vis = modApply.CountVisibleRows(area)
    Dim cvMs As Double
    cvMs = (Timer - t0) * 1000!

    LogShim.LogEvent "partcold", _
        Array("matched", "coldMs", "countVisMs", "visible"), _
        Array(r("MatchedCount"), Format$(coldMs, "0.0"), Format$(cvMs, "0.0"), vis)

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "partcold_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' @desc 에러 셀 1개가 지문 대조를 영구 불일치로 만드는지(캐시 무력화) 실측함
Public Sub MeasureErrorCellStale()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    Dim ws As Worksheet
    Set ws = wb.Sheets("DW")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Dim body As Range
    Set body = ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 2))   ' Z열(저카디널, 미스 싸다)

    Dim t0 As Single
    Dim a1 As clsColAsset, a2 As clsColAsset, a3 As clsColAsset, a4 As clsColAsset

    ' 1) 정상 열: 미스 -> 적중 (기준선)
    modCache.ClearCache
    t0 = Timer
    Set a1 = modCache.GetColumnAsset(body)
    Dim missMs As Double
    missMs = (Timer - t0) * 1000!

    t0 = Timer
    Set a2 = modCache.GetColumnAsset(body)
    Dim hitMs As Double
    hitMs = (Timer - t0) * 1000!
    Dim cleanHit As Boolean
    cleanHit = (a2 Is a1)

    ' 2) 에러 셀 1개 주입 (=NA()) 후 재접근 2회: 매번 재스캔되는가?
    ws.Cells(100, 2).Formula = "=NA()"
    t0 = Timer
    Set a3 = modCache.GetColumnAsset(body)
    Dim afterErr1 As Double
    afterErr1 = (Timer - t0) * 1000!
    Dim stale1 As Boolean
    stale1 = Not (a3 Is a2)     ' 값이 바뀌었으니 1회 재스캔은 정당

    t0 = Timer
    Set a4 = modCache.GetColumnAsset(body)
    Dim afterErr2 As Double
    afterErr2 = (Timer - t0) * 1000!
    Dim stale2 As Boolean
    stale2 = Not (a4 Is a3)     ' 무변경 재접근인데 또 재스캔이면 = 캐시 영구 무력화

    LogShim.LogEvent "errstale", _
        Array("missMs", "hitMs", "cleanHitIs", "afterErr1Ms", "afterErr2Ms", _
              "rescan1", "rescan2_smoking_gun"), _
        Array(Format$(missMs, "0.0"), Format$(hitMs, "0.0"), cleanHit, _
              Format$(afterErr1, "0.0"), Format$(afterErr2, "0.0"), stale1, stale2)

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "errstale_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' @desc dist .xlam 애드인 환경에서 퀵필터 2회 타이밍으로 캐시 적중을 실측함
Public Sub BenchAddinCacheHit()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    Const ADDIN_NAME As String = "UnivFilter_cachetest.xlam"
    Dim addinPath As String
    addinPath = ThisWorkbook.Path & "\dist\" & ADDIN_NAME

    On Error GoTo Fail
    ' 1) 애드인 장착 (별도 VBProject - 자기 modCache/지문을 가짐)
    Dim wbAddin As Workbook
    Set wbAddin = Workbooks.Open(addinPath)

    ' 2) fixture 열기 + 에러 셀 주입 (실무 시트 모사: 에러 상존 열)
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    Dim ws As Worksheet
    Set ws = wb.Sheets("DW")
    ws.Cells(100, 2).Formula = "=NA()"

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Dim area As Range, body As Range
    Set area = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 6))
    Set body = ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 2))
    area.AutoFilter

    Dim tok As String
    tok = FirstDisplay(body)

    ' 3) 애드인의 파이프라인을 직접 3회 실행 (1회차 콜드, 2·3회차 적중 기대)
    Dim runName As String
    runName = "'" & ADDIN_NAME & "'!modFilterRun.RunQuickFilter"

    Dim t0 As Single, r As Object
    Dim ms1 As Double, ms2 As Double, ms3 As Double, ok1 As Boolean

    t0 = Timer
    Set r = Application.Run(runName, body, area, 2&, tok)
    ms1 = (Timer - t0) * 1000!
    ok1 = r("Ok")

    modApply.ResetFilters area
    t0 = Timer
    Set r = Application.Run(runName, body, area, 2&, tok)
    ms2 = (Timer - t0) * 1000!

    modApply.ResetFilters area
    t0 = Timer
    Set r = Application.Run(runName, body, area, 2&, tok)
    ms3 = (Timer - t0) * 1000!

    LogShim.LogEvent "addinhit", _
        Array("ok", "coldMs", "warm1Ms", "warm2Ms", "hitVerdict"), _
        Array(ok1, Format$(ms1, "0.0"), Format$(ms2, "0.0"), Format$(ms3, "0.0"), _
              IIf(ms2 < ms1 * 0.7 And ms3 < ms1 * 0.7, "HIT", "NO_HIT"))

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    wbAddin.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "addinhit_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    On Error Resume Next
    Workbooks(ADDIN_NAME).Close SaveChanges:=False
    Application.EnableEvents = prevEv
End Sub

Private Sub StageLog(ByVal op As String, ByVal n As Long, ByVal ms As Double)
    LogShim.LogEvent "renderstage", Array("op", "n", "ms"), _
        Array(op, n, Format$(ms, "0.0"))
End Sub

' modColumnScan.AppendVm 복제 (Private라 직접 호출 불가).
Private Sub AppendVmB(ByVal vm As Object, ByVal vk As String, ByVal key As String)
    If vm.exists(vk) Then
        If InStr(1, vbNullChar & vm(vk) & vbNullChar, vbNullChar & key & vbNullChar) = 0 Then
            vm(vk) = vm(vk) & vbNullChar & key
        End If
    Else
        vm(vk) = key
    End If
End Sub

' @desc 셀1개 퀵필터의 실제 체감 소요(RunQuickFilter end-to-end + 단계분해)를 기록함
Public Sub BenchQuickPath()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    ' 폼 인스턴스화(퀵모드도 폼 New - 숨은 고정비) 1회 표본
    On Error Resume Next
    Dim t0 As Single, fi As Double
    t0 = Timer
    Dim frm As frmFilter01
    Set frm = New frmFilter01
    fi = (Timer - t0) * 1000!
    Unload frm
    Set frm = Nothing
    On Error GoTo Fail
    LogShim.LogEvent "forminit", Array("ms"), Array(Format$(fi, "0.0"))

    QuickCol wb.Sheets("DW"), 6, 1, "DW_P_amount_hicard"
    QuickCol wb.Sheets("DW"), 6, 2, "DW_Z_count_locard"
    QuickCol wb.Sheets("DW"), 6, 3, "DW_D_text"
    QuickCol wb.Sheets("DW"), 6, 5, "DW_C_datenum"
    QuickCol wb.Sheets("T"), 2, 1, "SL_V_amtText"
    QuickCol wb.Sheets("T"), 2, 2, "SL_AH_result"

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    LogShim.LogEvent "quickbench_done", Array("ok"), Array(True)
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "quick_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' 셀1개(첫 데이터셀 표시값을 클립보드 삼아) 퀵필터를 실제 파이프라인으로 재고 단계분해한다.
Private Sub QuickCol(ByVal ws As Worksheet, ByVal nCols As Long, _
                     ByVal field As Long, ByVal tag As String)
    ' 직전 열의 필터가 남아 End(xlUp)를 오염시키지 않도록 먼저 전량 표시
    On Error Resume Next
    ws.ShowAllData
    ws.AutoFilterMode = False
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, field).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim area As Range, body As Range
    Set area = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, nCols))
    Set body = ws.Range(ws.Cells(2, field), ws.Cells(lastRow, field))

    Dim inputText As String
    inputText = FirstDisplay(body)

    ' AutoFilter 새로 켜기
    area.AutoFilter

    Dim t0 As Single, r As Object

    ' 1) 총 소요 - 콜드(스캔 미스 포함): 실제 RunQuickFilter
    modApply.ResetFilters area
    modCache.ClearCache
    t0 = Timer
    Set r = modFilterRun.RunQuickFilter(body, area, field, inputText)
    Dim totalCold As Double
    totalCold = (Timer - t0) * 1000!

    ' 2) 총 소요 - 웜(스캔 적중): 필터 리셋 후 재실행
    modApply.ResetFilters area
    t0 = Timer
    Set r = modFilterRun.RunQuickFilter(body, area, field, inputText)
    Dim totalWarm As Double
    totalWarm = (Timer - t0) * 1000!

    ' 3) 단계분해 (캐시 웜)
    Dim asset As clsColAsset
    t0 = Timer
    Set asset = modCache.GetColumnAsset(body)
    Dim scanHit As Double
    scanHit = (Timer - t0) * 1000!

    Dim tokens As Variant
    t0 = Timer
    tokens = modInputParse.Tokenize(inputText)
    Dim tokMs As Double
    tokMs = (Timer - t0) * 1000!

    Dim outcome As Object
    t0 = Timer
    Set outcome = modMatch.Execute(tokens, asset, False)
    Dim matchMs As Double
    matchMs = (Timer - t0) * 1000!
    Dim matched As clsSet
    Set matched = outcome("Matched")

    modApply.ResetFilters area
    Dim ai As Object
    t0 = Timer
    Set ai = modApply.ApplyDisplays(area, field, matched, False, False)
    Dim applyMs As Double
    applyMs = (Timer - t0) * 1000!

    t0 = Timer
    Dim vis As Long
    vis = modApply.CountVisibleRows(area)
    Dim countMs As Double
    countMs = (Timer - t0) * 1000!

    LogShim.LogEvent "quickbench", _
        Array("tag", "rows", "matched", "visible", "input", _
              "totalColdMs", "totalWarmMs", "scanHitMs", "tokMs", "matchMs", "applyMs", "countVisMs"), _
        Array(tag, lastRow - 1, matched.Count, vis, inputText, _
              Format$(totalCold, "0.0"), Format$(totalWarm, "0.0"), Format$(scanHit, "0.0"), _
              Format$(tokMs, "0.00"), Format$(matchMs, "0.0"), _
              Format$(applyMs, "0.0"), Format$(countMs, "0.0"))
End Sub

' 본문 첫 비어있지 않은 셀의 표시 텍스트(사용자가 복사한 셀을 모사).
Private Function FirstDisplay(ByVal body As Range) As String
    Dim i As Long, n As Long
    n = body.Rows.Count
    If n > 5000 Then n = 5000    ' 앞쪽에서 탐색 상한
    For i = 1 To n
        Dim s As String
        s = body.Cells(i, 1).Text
        If Len(s) > 0 Then
            FirstDisplay = s
            Exit Function
        End If
    Next i
End Function

' @desc 셀1개 퀵필터 콜드 경로의 단계별 소요(지문/렌더빌드/매칭/적용/계수/리포트)를 기록함
Public Sub BenchColdBreakdown()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    ColdCol wb.Sheets("DW"), 6, 1, "DW_P_amount_hicard"
    ColdCol wb.Sheets("DW"), 6, 2, "DW_Z_count_locard"
    ColdCol wb.Sheets("DW"), 6, 3, "DW_D_text"
    ColdCol wb.Sheets("DW"), 6, 5, "DW_C_datenum"
    ColdCol wb.Sheets("DW"), 6, 6, "DW_AC_amount_err"
    ColdCol wb.Sheets("T"), 2, 1, "SL_V_amtText"
    ColdCol wb.Sheets("T"), 2, 2, "SL_AH_result"

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    LogShim.LogEvent "coldbench_done", Array("ok"), Array(True)
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "cold_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' 콜드(캐시 미스) 퀵필터를 단계별로 분해 계측한다: e2e 총소요를 먼저 재고,
' 캐시를 다시 비워 동일 콜드 상태에서 파이프라인 각 단계를 개별 타이밍한다.
Private Sub ColdCol(ByVal ws As Worksheet, ByVal nCols As Long, _
                    ByVal field As Long, ByVal tag As String)
    On Error Resume Next
    ws.ShowAllData
    ws.AutoFilterMode = False
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, field).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim area As Range, body As Range
    Set area = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, nCols))
    Set body = ws.Range(ws.Cells(2, field), ws.Cells(lastRow, field))

    Dim inputText As String
    inputText = FirstDisplay(body)

    area.AutoFilter

    Dim t0 As Single

    ' 1) e2e 콜드 총소요 (실제 RunQuickFilter)
    modApply.ResetFilters area
    modCache.ClearCache
    t0 = Timer
    Dim r As Object
    Set r = modFilterRun.RunQuickFilter(body, area, field, inputText)
    Dim totalCold As Double
    totalCold = (Timer - t0) * 1000!
    Dim e2eOk As Boolean, e2eErr As String
    e2eOk = r("Ok")
    e2eErr = r("ErrMsg")

    ' 1') 공식 관문 콜드: GetColumnAsset 전체 (e2e와 동일 진입로의 자산 취득만)
    modApply.ResetFilters area
    modCache.ClearCache
    Utilities.OptimizedMode True    ' RunFilter 내부 상태 모사 (Cleanup이 껐음)
    t0 = Timer
    Dim assetGate As clsColAsset
    Set assetGate = modCache.GetColumnAsset(body)
    Dim assetColdMs As Double
    assetColdMs = (Timer - t0) * 1000!

    ' 2) 단계분해: 캐시를 다시 비워 동일 콜드 상태 재현
    modCache.ClearCache

    ' 2a) 지문 계산 (Value2 벌크 + NumberFormatLocal)
    t0 = Timer
    Dim sig As clsColSig
    Set sig = modColumnScan.ComputeSignature(body)
    Dim sigMs As Double
    sigMs = (Timer - t0) * 1000!

    ' 2b) 자산 빌드 (Evaluate(TEXT) 벌크 렌더 + 행별 키 채움)
    t0 = Timer
    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAssetFromSig(body, body.Address(External:=True), sig)
    Dim buildMs As Double
    buildMs = (Timer - t0) * 1000!

    ' 2b') 벌크 Evaluate 단독 복제 (빌드 내부의 COM 렌더 몫 분리, 균일서식만)
    Dim evalMs As Double
    evalMs = -1
    If (Not sig.IsMixed) And (Not sig.IsPathological) And sig.FmtUniform <> "@" Then
        ' 실경로(RenderBlockBulk)와 동일하게 숫자 셀이 있을 때만 렌더 (텍스트 열 왜곡 방지)
        Dim vChk As Variant, hasNum As Boolean, iChk As Long
        vChk = sig.ContentVec
        For iChk = 1 To lastRow - 1
            If VarType(vChk(iChk, 1)) = vbDouble Then
                hasNum = True
                Exit For
            End If
        Next iChk
        If Not hasNum Then GoTo SkipEval
        Dim fmt As String, useFmt As String
        fmt = sig.FmtUniform
        If fmt = "General" Or fmt = "G/" & ChrW(&HD45C) & ChrW(&HC900) Then
            useFmt = "General"
        Else
            useFmt = Replace(fmt, "!", "")
        End If
        t0 = Timer
        Dim rendered As Variant
        rendered = ws.Evaluate("TEXT(" & body.Address & "," & Chr$(34) & _
                               Replace(useFmt, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34) & ")")
        evalMs = (Timer - t0) * 1000!
        If Not IsArray(rendered) Then evalMs = -2    ' 비배열 = 복제 실패 표시
    End If
SkipEval:

    ' 2c) 토큰화
    t0 = Timer
    Dim tokens As Variant
    tokens = modInputParse.Tokenize(inputText)
    Dim tokMs As Double
    tokMs = (Timer - t0) * 1000!

    ' 2d) 매칭 (렌더키 멤버십 스캔)
    t0 = Timer
    Dim outcome As Object
    Set outcome = modMatch.Execute(tokens, asset, False)
    Dim matchMs As Double
    matchMs = (Timer - t0) * 1000!
    Dim matched As clsSet
    Set matched = outcome("Matched")

    ' 2e) 적용 (내부 ResetFilters 포함 - RunFilter와 동일 경로)
    t0 = Timer
    Dim ai As Object
    Set ai = modApply.ApplyDisplays(area, field, matched, False, False)
    Dim applyMs As Double
    applyMs = (Timer - t0) * 1000!

    ' 2f) 자가검증 가시행 계수
    t0 = Timer
    Dim vis As Long
    vis = modApply.CountVisibleRows(area)
    Dim countMs As Double
    countMs = (Timer - t0) * 1000!

    ' 2g) 리포트 구성
    t0 = Timer
    Dim msgs As Object
    Set msgs = modReport.BuildMessages(outcome, False, ai)
    Dim reportMs As Double
    reportMs = (Timer - t0) * 1000!

    Utilities.OptimizedMode False
    Utilities.OptimizedMode True    ' 다음 열도 동일 조건

    Dim sumMs As Double
    sumMs = sigMs + buildMs + tokMs + matchMs + applyMs + countMs + reportMs

    LogShim.LogEvent "coldbench", _
        Array("tag", "rows", "input", "e2eOk", "e2eErr", "totalColdMs", "assetColdMs", _
              "sigMs", "buildMs", "evalOnlyMs", _
              "tokMs", "matchMs", "applyMs", "countVisMs", "reportMs", "sumMs", "otherMs"), _
        Array(tag, lastRow - 1, inputText, e2eOk, e2eErr, Format$(totalCold, "0.0"), _
              Format$(assetColdMs, "0.0"), _
              Format$(sigMs, "0.0"), Format$(buildMs, "0.0"), Format$(evalMs, "0.0"), _
              Format$(tokMs, "0.00"), Format$(matchMs, "0.0"), Format$(applyMs, "0.0"), _
              Format$(countMs, "0.0"), Format$(reportMs, "0.00"), _
              Format$(sumMs, "0.0"), Format$(totalCold - sumMs, "0.0"))
End Sub

' @desc SL_V 콜드 이상치 판별: 타입 센서스와 e2e 전/후 빌드 소요 대조를 기록함
Public Sub ProbeSLVCold()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Open(FixturePath())
    Utilities.OptimizedMode True

    Dim ws As Worksheet
    Set ws = wb.Sheets("T")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim area As Range, body As Range
    Set area = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 2))
    Set body = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
    Dim nRows As Long
    nRows = lastRow - 1

    ' 1) 타입 센서스 + 서식
    Dim v2 As Variant
    v2 = body.Value2
    Dim nStr As Long, nDbl As Long, nEmp As Long, nOth As Long, i As Long
    For i = 1 To nRows
        If IsEmpty(v2(i, 1)) Then
            nEmp = nEmp + 1
        ElseIf VarType(v2(i, 1)) = vbString Then
            nStr = nStr + 1
        ElseIf VarType(v2(i, 1)) = vbDouble Then
            nDbl = nDbl + 1
        Else
            nOth = nOth + 1
        End If
    Next i
    Dim vFmt As Variant
    vFmt = body.NumberFormatLocal
    If IsNull(vFmt) Then vFmt = "(mixed)"

    ' 2) e2e 이전(처녀 상태) 콜드 빌드
    modCache.ClearCache
    Dim t0 As Single
    t0 = Timer
    Dim a1 As clsColAsset
    Set a1 = modCache.GetColumnAsset(body)
    Dim preMs As Double
    preMs = (Timer - t0) * 1000!

    ' 3) e2e 콜드 (필터 적용까지 - 시트에 필터 이력 생성)
    area.AutoFilter
    modApply.ResetFilters area
    modCache.ClearCache
    Dim inputText As String
    inputText = FirstDisplay(body)
    t0 = Timer
    Dim r As Object
    Set r = modFilterRun.RunQuickFilter(body, area, 1, inputText)
    Dim e2eMs As Double
    e2eMs = (Timer - t0) * 1000!

    ' 4) e2e 이후 콜드 빌드 (필터 이력 있는 시트)
    modApply.ResetFilters area
    modCache.ClearCache
    Utilities.OptimizedMode True
    t0 = Timer
    Dim a2 As clsColAsset
    Set a2 = modCache.GetColumnAsset(body)
    Dim postMs As Double
    postMs = (Timer - t0) * 1000!

    ' 5) .Text 단가 (표본 2000셀)
    Dim nSmp As Long
    nSmp = 2000
    If nSmp > nRows Then nSmp = nRows
    Dim s As String
    t0 = Timer
    For i = 1 To nSmp
        s = body.Cells(i, 1).Text
    Next i
    Dim textMs As Double
    textMs = (Timer - t0) * 1000!

    LogShim.LogEvent "slvprobe", _
        Array("rows", "fmt", "nStr", "nDbl", "nEmp", "nOth", "input", _
              "preE2eColdMs", "e2eMs", "postE2eColdMs", "text2000Ms", _
              "hasRendered1", "hasRendered2"), _
        Array(nRows, CStr(vFmt), nStr, nDbl, nEmp, nOth, inputText, _
              Format$(preMs, "0.0"), Format$(e2eMs, "0.0"), Format$(postMs, "0.0"), _
              Format$(textMs, "0.0"), a1.HasRendered, a2.HasRendered)

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "slvprobe_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

Private Sub BenchColumn(ByVal ws As Worksheet, ByVal col As Long, ByVal tag As String)
    Const REP As Long = 15
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim body As Range
    Set body = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))
    Dim nRows As Long
    nRows = lastRow - 1

    Dim t0 As Single, j As Long
    Dim asset As clsColAsset

    ' 1) 미스(콜드): ComputeSignature + BuildAssetFromSig 전체 재빌드
    modCache.ClearCache
    t0 = Timer
    Set asset = modCache.GetColumnAsset(body)
    Dim missMs As Double
    missMs = (Timer - t0) * 1000!

    ' 2) 적중(웜): 접근당 = ComputeSignature + Matches (렌더 없음)
    t0 = Timer
    For j = 1 To REP
        Set asset = modCache.GetColumnAsset(body)
    Next j
    Dim hitMs As Double
    hitMs = (Timer - t0) * 1000! / REP

    ' 3) ComputeSignature 단독 (Value2 벌크 + 서식판정)
    Dim sig As clsColSig
    t0 = Timer
    For j = 1 To REP
        Set sig = modColumnScan.ComputeSignature(body)
    Next j
    Dim csMs As Double
    csMs = (Timer - t0) * 1000! / REP

    ' 4) Matches 단독(동일 지문): ContentVec 원소별 대조 전량 통과
    Dim s1 As clsColSig, s2 As clsColSig, b As Boolean
    Set s1 = modColumnScan.ComputeSignature(body)
    Set s2 = modColumnScan.ComputeSignature(body)
    t0 = Timer
    For j = 1 To REP
        b = s1.Matches(s2)
    Next j
    Dim mMs As Double
    mMs = (Timer - t0) * 1000! / REP

    LogShim.LogEvent "realbench", _
        Array("tag", "rows", "displays", "blanks", "missMs", "hitMs", "computeSigMs", "matchesMs"), _
        Array(tag, nRows, asset.Displays.Count, asset.BlankCount, _
              Format$(missMs, "0.0"), Format$(hitMs, "0.00"), _
              Format$(csMs, "0.00"), Format$(mMs, "0.00"))
End Sub
