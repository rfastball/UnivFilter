Attribute VB_Name = "modBenchReal"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc 실데이터 열을 격리 fixture로 복제하고 스캔경로 접근당 비용 계측 담당
' @depends modCache, modColumnScan, modFilterRun, modInputParse, modMatch, modApply, modReport, clsColAsset, clsColSig, clsSet, frmFilter01, Utilities, LogShim
Option Explicit

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long

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

' @desc 셀별 .Text 단가의 필터 이력 의존성을 상태별 전행 패스로 판별해 기록함
' 방아쇠 후보 분리: 처녀 -> 반복 -> 필터중(숨김) -> 해제 직후 -> 반복 -> 수동 숨김/해제 -> 반복
Public Sub ProbeTextDecay()
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

    Dim msA As Double, msB As Double, msC As Double, msD As Double
    Dim msE As Double, msF As Double, msG As Double

    ' A: 처녀 상태(방금 연 워크북, AutoFilter 없음) 전행 .Text
    msA = TextPass(body, nRows)
    ' B: 반복 (웜 확인)
    msB = TextPass(body, nRows)

    ' C: AutoFilter 조건 적용(전행 숨김) 상태에서 전행 .Text
    area.AutoFilter
    area.AutoFilter Field:=1, Criteria1:="=###UF_PROBE_NONE###"
    msC = TextPass(body, nRows)

    ' D: ShowAllData(해제) 직후 전행 .Text
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo Fail
    msD = TextPass(body, nRows)
    ' E: 반복
    msE = TextPass(body, nRows)

    ' F: AutoFilter 끄고 수동 행 숨김/해제 직후 전행 .Text
    ws.AutoFilterMode = False
    body.EntireRow.Hidden = True
    body.EntireRow.Hidden = False
    msF = TextPass(body, nRows)
    ' G: 반복
    msG = TextPass(body, nRows)

    ' H/I: 저속 상태에서 위치 의존성 - 앞 2000셀 vs 뒤 2000셀
    Dim msH As Double, msI As Double
    Dim t0 As Single, i As Long, s As String
    t0 = Timer
    For i = 1 To 2000
        s = body.Cells(i, 1).Text
    Next i
    msH = (Timer - t0) * 1000!
    t0 = Timer
    For i = nRows - 1999 To nRows
        s = body.Cells(i, 1).Text
    Next i
    msI = (Timer - t0) * 1000!

    ' K: 저장 -> 재오픈 후 전행 .Text (저속 상태가 파일에 지속되는지)
    Dim sTmp As String
    sTmp = Environ$("TEMP") & "\uf_textdecay_tmp.xlsx"
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.SaveAs sTmp, FileFormat:=51
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = prevAlerts
    Set wb = Workbooks.Open(sTmp)
    Set ws = wb.Sheets("T")
    Set body = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
    Dim msK As Double
    msK = TextPass(body, nRows)

    LogShim.LogEvent "textdecay", _
        Array("rows", "A_freshMs", "B_repeatMs", "C_filteredMs", "D_unhideMs", _
              "E_repeatMs", "F_manualUnhideMs", "G_repeatMs", _
              "H_first2000Ms", "I_last2000Ms", "K_reopenMs"), _
        Array(nRows, Format$(msA, "0.0"), Format$(msB, "0.0"), Format$(msC, "0.0"), _
              Format$(msD, "0.0"), Format$(msE, "0.0"), Format$(msF, "0.0"), _
              Format$(msG, "0.0"), Format$(msH, "0.0"), Format$(msI, "0.0"), _
              Format$(msK, "0.0"))

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    On Error Resume Next
    Kill sTmp
    On Error GoTo 0
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    LogShim.LogEvent "textdecay_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' @desc .Text 우회 후보 검증: 에러/불리언 매핑 동치, @숫자 폭 의존, fmtLocal 이력 저하 기록함
Public Sub ProbeDisplayMap()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Utilities.OptimizedMode True
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)

    ' 1) 에러 표시 매핑 동치: CVErr 코드 -> 고정 문자열이 .Text와 일치하는가 (넓은/좁은 열)
    Dim codes As Variant
    codes = Array(2000, 2007, 2015, 2023, 2029, 2036, 2042, 2043, 2045, 2046, 2047, 2048, 2049, 2050)
    Dim assigned As New Collection      ' Array(row, code)
    Dim i As Long
    ws.Columns(1).ColumnWidth = 40
    For i = 0 To UBound(codes)
        On Error Resume Next
        ws.Cells(i + 1, 1).Value = CVErr(codes(i))
        If Err.Number = 0 Then assigned.Add Array(i + 1, CLng(codes(i)))
        Err.Clear
        On Error GoTo Fail
    Next i
    Dim nOkWide As Long, nMisWide As Long, sMis As String
    Dim itm As Variant, sMap As String, sTxt As String
    For Each itm In assigned
        sMap = ErrDisplay(CLng(itm(1)))
        sTxt = ws.Cells(itm(0), 1).Text
        If sMap = sTxt Then
            nOkWide = nOkWide + 1
        Else
            nMisWide = nMisWide + 1
            sMis = sMis & "[" & itm(1) & ":txt=" & sTxt & " map=" & sMap & "]"
        End If
    Next itm
    ws.Columns(1).ColumnWidth = 3
    Dim nOkNarrow As Long, nMisNarrow As Long
    For Each itm In assigned
        If ErrDisplay(CLng(itm(1))) = ws.Cells(itm(0), 1).Text Then
            nOkNarrow = nOkNarrow + 1
        Else
            nMisNarrow = nMisNarrow + 1
        End If
    Next itm
    LogShim.LogEvent "dispmap_err", _
        Array("tried", "assigned", "okWide", "misWide", "okNarrow", "misNarrow", "misDetail"), _
        Array(UBound(codes) + 1, assigned.Count, nOkWide, nMisWide, nOkNarrow, nMisNarrow, sMis)

    ' 2) 불리언 표시: TRUE/FALSE 고정 매핑 동치 (넓은/좁은 열)
    ws.Columns(2).ColumnWidth = 20
    ws.Cells(1, 2).Value = True
    ws.Cells(2, 2).Value = False
    Dim bWideT As String, bWideF As String, bNarrowT As String
    bWideT = ws.Cells(1, 2).Text
    bWideF = ws.Cells(2, 2).Text
    ws.Columns(2).ColumnWidth = 3
    bNarrowT = ws.Cells(1, 2).Text
    LogShim.LogEvent "dispmap_bool", _
        Array("wideTrue", "wideFalse", "narrowTrue"), _
        Array(bWideT, bWideF, bNarrowT)

    ' 3) @서식 숫자: 폭별 .Text vs WText(General) vs CStr (P11 프로브 설계용 관측)
    ws.Cells(1, 3).Value = 212970000
    ws.Cells(1, 3).NumberFormat = "@"      ' 값 먼저, 서식 나중 (fixture 재현)
    ws.Columns(3).ColumnWidth = 30
    Dim atWide As String
    atWide = ws.Cells(1, 3).Text
    ws.Columns(3).ColumnWidth = 8
    Dim atNarrow As String
    atNarrow = ws.Cells(1, 3).Text
    LogShim.LogEvent "dispmap_at", _
        Array("wideText", "narrowText", "wtextGeneral", "cstr"), _
        Array(atWide, atNarrow, _
              Application.WorksheetFunction.Text(212970000#, "General"), CStr(212970000#))

    ' 4) fmtLocal/.Text 단가의 이력 저하 대조 (20k행, 숨김 전/후 x 앞/뒤 2000셀)
    Dim ws2 As Worksheet
    Set ws2 = wb.Sheets.Add
    Const N As Long = 20000
    Dim arr() As Variant
    ReDim arr(1 To N, 1 To 1)
    For i = 1 To N
        arr(i, 1) = i * 7
    Next i
    Dim rng As Range
    Set rng = ws2.Range(ws2.Cells(1, 1), ws2.Cells(N, 1))
    rng.NumberFormat = "#,##0"
    rng.Value2 = arr

    Dim preTxtF As Double, preTxtB As Double, preFmtF As Double, preFmtB As Double
    preTxtF = OpPass(ws2, 1, 2000, True)
    preTxtB = OpPass(ws2, N - 1999, N, True)
    preFmtF = OpPass(ws2, 1, 2000, False)
    preFmtB = OpPass(ws2, N - 1999, N, False)

    rng.EntireRow.Hidden = True
    rng.EntireRow.Hidden = False

    Dim postTxtF As Double, postTxtB As Double, postFmtF As Double, postFmtB As Double
    postTxtF = OpPass(ws2, 1, 2000, True)
    postTxtB = OpPass(ws2, N - 1999, N, True)
    postFmtF = OpPass(ws2, 1, 2000, False)
    postFmtB = OpPass(ws2, N - 1999, N, False)

    LogShim.LogEvent "dispmap_decay", _
        Array("preTxtFrontMs", "preTxtBackMs", "preFmtFrontMs", "preFmtBackMs", _
              "postTxtFrontMs", "postTxtBackMs", "postFmtFrontMs", "postFmtBackMs"), _
        Array(Format$(preTxtF, "0.0"), Format$(preTxtB, "0.0"), _
              Format$(preFmtF, "0.0"), Format$(preFmtB, "0.0"), _
              Format$(postTxtF, "0.0"), Format$(postTxtB, "0.0"), _
              Format$(postFmtF, "0.0"), Format$(postFmtB, "0.0"))

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    LogShim.LogEvent "dispmap_probe_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' 에러코드 -> 표시 문자열 고정 매핑 (우회 후보 - 프로브가 .Text와 동치 검증)
Private Function ErrDisplay(ByVal code As Long) As String
    Select Case code
        Case 2000: ErrDisplay = "#NULL!"
        Case 2007: ErrDisplay = "#DIV/0!"
        Case 2015: ErrDisplay = "#VALUE!"
        Case 2023: ErrDisplay = "#REF!"
        Case 2029: ErrDisplay = "#NAME?"
        Case 2036: ErrDisplay = "#NUM!"
        Case 2042: ErrDisplay = "#N/A"
        Case 2043: ErrDisplay = "#GETTING_DATA"
        Case 2045: ErrDisplay = "#SPILL!"
        Case 2046: ErrDisplay = "#CONNECT!"
        Case 2047: ErrDisplay = "#BLOCKED!"
        Case 2048: ErrDisplay = "#UNKNOWN!"
        Case 2049: ErrDisplay = "#FIELD!"
        Case 2050: ErrDisplay = "#CALC!"
        Case Else: ErrDisplay = "Error " & code
    End Select
End Function

' [rowFrom,rowTo] 셀별 연산 1패스 소요(ms): bText=True면 .Text, False면 NumberFormatLocal.
Private Function OpPass(ByVal ws As Worksheet, ByVal rowFrom As Long, _
                        ByVal rowTo As Long, ByVal bText As Boolean) As Double
    Dim t0 As Single, i As Long, s As String
    t0 = Timer
    If bText Then
        For i = rowFrom To rowTo
            s = ws.Cells(i, 1).Text
        Next i
    Else
        For i = rowFrom To rowTo
            s = ws.Cells(i, 1).NumberFormatLocal
        Next i
    End If
    OpPass = (Timer - t0) * 1000!
End Function

' @desc 인프로세스 렌더 단가 실측: VBA Format$/CStr vs 벌크 Evaluate(TEXT) 대조 기록함
' (Evaluate 경로 자체의 우회 실익 판정용 - 91k 규모 인메모리 루프)
Public Sub ProbeInprocRender()
    LogShim.Reset
    On Error GoTo Fail

    Const N As Long = 91235
    Dim vals() As Double
    ReDim vals(1 To N)
    Dim i As Long
    For i = 1 To N
        vals(i) = (i * 7919#) + (i Mod 100) / 100#   ' 고카디널 금액 모사
    Next i

    Dim t0 As Single, s As String

    ' 1) VBA Format$ "#,##0" (인프로세스 - 우회 후보)
    t0 = Timer
    For i = 1 To N
        s = Format$(vals(i), "#,##0")
    Next i
    Dim fmtMs As Double
    fmtMs = (Timer - t0) * 1000!

    ' 2) VBA CStr (General 근사 하한)
    t0 = Timer
    For i = 1 To N
        s = CStr(vals(i))
    Next i
    Dim cstrMs As Double
    cstrMs = (Timer - t0) * 1000!

    ' 3) WorksheetFunction.Text 단건 루프 (COM 단건 대조군)
    t0 = Timer
    For i = 1 To 5000
        s = Application.WorksheetFunction.Text(vals(i), "#,##0")
    Next i
    Dim wtMs As Double
    wtMs = (Timer - t0) * 1000! * (CDbl(N) / 5000#)   ' N행 환산

    ' 4) Format$ vs Excel TEXT 표본 대조 (동일성 아닌 참고 - 1000건)
    Dim nMis As Long, sample As String
    Dim wt As String
    For i = 1 To 1000
        s = Format$(vals(i * 91), "#,##0")
        wt = Application.WorksheetFunction.Text(vals(i * 91), "#,##0")
        If s <> wt Then
            nMis = nMis + 1
            If nMis <= 3 Then sample = sample & "[" & s & "/" & wt & "]"
        End If
    Next i

    LogShim.LogEvent "inproc", _
        Array("n", "vbaFormatMs", "cstrMs", "wtextEquivMs", "mis1000", "misSample"), _
        Array(N, Format$(fmtMs, "0.0"), Format$(cstrMs, "0.0"), _
              Format$(wtMs, "0.0"), nMis, sample)
    Exit Sub
Fail:
    LogShim.LogEvent "inproc_err", Array("num", "desc"), Array(Err.Number, Err.Description)
End Sub

' @desc Evaluate(TEXT) 대안 실측: Application.Text 배열 리프팅 소요/정합/빈칸 거동 기록함
Public Sub ProbeEvalAlternatives()
    LogShim.Reset
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Fail
    Const N As Long = 91235
    Dim arr() As Variant
    ReDim arr(1 To N, 1 To 1)
    Dim i As Long
    For i = 1 To N
        arr(i, 1) = (i * 7919#) + (i Mod 100) / 100#
    Next i
    arr(5, 1) = Empty                      ' 빈칸 거동 관찰용
    arr(7, 1) = CVErr(2042)                ' 에러 거동 관찰용

    Dim t0 As Single

    ' 1) Application.Text(배열, 서식) - 리프팅 여부/소요
    t0 = Timer
    Dim res As Variant
    res = Application.Text(arr, "#,##0")
    Dim appMs As Double
    appMs = (Timer - t0) * 1000!
    Dim liftOk As Boolean
    liftOk = IsArray(res)

    Dim sBlank As String, sErr As String, sVal As String, nMis As Long
    If liftOk Then
        If IsError(res(5, 1)) Then sBlank = "(error)" Else sBlank = "[" & CStr(res(5, 1)) & "]"
        If IsError(res(7, 1)) Then sErr = "(error)" Else sErr = "[" & CStr(res(7, 1)) & "]"
        sVal = CStr(res(1, 1))
        ' 표본 정합: WText와 1000건 대조 (빈칸/에러 행 제외)
        Dim wt As String, j As Long
        For i = 1 To 1000
            j = i * 91
            If j <> 5 And j <> 7 Then
                wt = Application.WorksheetFunction.Text(arr(j, 1), "#,##0")
                If CStr(res(j, 1)) <> wt Then nMis = nMis + 1
            End If
        Next i
    End If

    ' 2) 동일 데이터를 임시 시트에 써서 Evaluate(TEXT) 기준선 재측정 (동일 런 대조)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Utilities.OptimizedMode True
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(N, 1))
    rng.NumberFormat = "#,##0"
    rng.Value2 = arr
    t0 = Timer
    Dim res2 As Variant
    res2 = ws.Evaluate("TEXT(" & rng.Address & ",""#,##0"")")
    Dim evalMs As Double
    evalMs = (Timer - t0) * 1000!

    ' 3) WorksheetFunction.Text(배열) 리프팅 시도 (참고)
    Dim wfLift As String
    On Error Resume Next
    Dim res3 As Variant
    res3 = Application.WorksheetFunction.Text(arr, "#,##0")
    If Err.Number <> 0 Then
        wfLift = "ERR:" & Err.Number
    ElseIf IsArray(res3) Then
        wfLift = "array"
    Else
        wfLift = "scalar"
    End If
    Err.Clear
    On Error GoTo Fail

    LogShim.LogEvent "evalalt", _
        Array("n", "appTextMs", "lifted", "evalBaseMs", "blankRender", "errRender", _
              "firstVal", "mis1000", "wfLift"), _
        Array(N, Format$(appMs, "0.0"), liftOk, Format$(evalMs, "0.0"), _
              sBlank, sErr, sVal, nMis, wfLift)

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    LogShim.LogEvent "evalalt_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' @desc 파이프라인 고정비 인벤토리: 폼/토글/리셋/dict/클립보드/상태바 항목별 단가 기록함
' (Evaluate 밖 오버헤드의 감량 여지 판정용 - quickbench의 other/forminit 분해)
Public Sub ProbeFixedOverheads()
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

    Dim t0 As Single, i As Long

    ' 1) 폼 인스턴스화 1회차/2회차 (퀵 핸들러의 숨은 고정비)
    t0 = Timer
    Dim frm1 As frmFilter01
    Set frm1 = New frmFilter01
    Dim formFirst As Double
    formFirst = (Timer - t0) * 1000!
    Unload frm1
    Set frm1 = Nothing
    t0 = Timer
    Dim frm2 As frmFilter01
    Set frm2 = New frmFilter01
    Dim formSecond As Double
    formSecond = (Timer - t0) * 1000!
    Unload frm2
    Set frm2 = Nothing

    ' 2) OptimizedMode True/False 왕복 (RunFilter가 매 실행 지불)
    Utilities.OptimizedMode False
    t0 = Timer
    For i = 1 To 10
        Utilities.OptimizedMode True
        Utilities.OptimizedMode False
    Next i
    Dim optMs As Double
    optMs = (Timer - t0) * 1000! / 10#
    Utilities.OptimizedMode True

    ' 3) ResetFilters - 조건 없는 시트(no-op)와 조건 있는 시트(실해제)
    modApply.ResetFilters area
    t0 = Timer
    For i = 1 To 10
        modApply.ResetFilters area
    Next i
    Dim resetNoop As Double
    resetNoop = (Timer - t0) * 1000! / 10#

    Dim firstKey As String
    firstKey = FirstDisplay(body)
    t0 = Timer
    For i = 1 To 5
        area.AutoFilter Field:=1, Criteria1:=Array(firstKey), Operator:=xlFilterValues
        modApply.ResetFilters area
    Next i
    Dim applyResetPair As Double
    applyResetPair = (Timer - t0) * 1000! / 5#

    ' 4) Scripting.Dictionary 생성 x 실행당 개수(약 8개) 환산
    t0 = Timer
    Dim d As Object
    For i = 1 To 400
        Set d = CreateObject("Scripting.Dictionary")
    Next i
    Dim dictPer8 As Double
    dictPer8 = (Timer - t0) * 1000! / 400# * 8#

    ' 5) 클립보드 읽기 (퀵 핸들러 매 실행)
    t0 = Timer
    Dim sClip As String
    For i = 1 To 10
        sClip = ClipboardAPI.GetClipboardText()
    Next i
    Dim clipMs As Double
    clipMs = (Timer - t0) * 1000! / 10#

    ' 6) 상태바 설정/해제 (리포트 경로)
    t0 = Timer
    For i = 1 To 10
        Application.StatusBar = "UnivFilter: probe"
        Application.StatusBar = False
    Next i
    Dim sbarMs As Double
    sbarMs = (Timer - t0) * 1000! / 10#

    ' 7) 캐시 키 Address(External) (GetColumnAsset 매 호출)
    t0 = Timer
    Dim sKey As String
    For i = 1 To 100
        sKey = body.Address(External:=True)
    Next i
    Dim addrMs As Double
    addrMs = (Timer - t0) * 1000! / 100#

    ' 8) 지문 검증 웜 단가 재확인 (ComputeSignature + Matches)
    modCache.ClearCache
    Dim asset As clsColAsset
    Set asset = modCache.GetColumnAsset(body)
    t0 = Timer
    For i = 1 To 10
        Set asset = modCache.GetColumnAsset(body)
    Next i
    Dim sigVerify As Double
    sigVerify = (Timer - t0) * 1000! / 10#

    LogShim.LogEvent "fixedcost", _
        Array("formFirstMs", "formSecondMs", "optToggleMs", "resetNoopMs", _
              "applyResetPairMs", "dictPer8Ms", "clipReadMs", "statusBarMs", _
              "addrExtMs", "sigVerifyHitMs"), _
        Array(Format$(formFirst, "0.0"), Format$(formSecond, "0.0"), _
              Format$(optMs, "0.00"), Format$(resetNoop, "0.00"), _
              Format$(applyResetPair, "0.0"), Format$(dictPer8, "0.00"), _
              Format$(clipMs, "0.00"), Format$(sbarMs, "0.00"), _
              Format$(addrMs, "0.00"), Format$(sigVerify, "0.0"))

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    LogShim.LogEvent "fixedcost_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' @desc 증분 패치 재스캔 실측: 편집 1셀/3셀 후 GetColumnAsset 소요와 패치 여부 기록함
Public Sub ProbePatchRescan()
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

    Dim t0 As Single
    modCache.ClearCache
    t0 = Timer
    Dim a1 As clsColAsset
    Set a1 = modCache.GetColumnAsset(body)
    Dim coldMs As Double
    coldMs = (Timer - t0) * 1000!

    ' 웜 (무변경)
    t0 = Timer
    Dim a2 As clsColAsset
    Set a2 = modCache.GetColumnAsset(body)
    Dim warmMs As Double
    warmMs = (Timer - t0) * 1000!

    ' 1셀 편집 -> 재취득 (증분 패치 기대)
    ws.Cells(500, 1).Value = 12345678
    t0 = Timer
    Dim a3 As clsColAsset
    Set a3 = modCache.GetColumnAsset(body)
    Dim patch1Ms As Double
    patch1Ms = (Timer - t0) * 1000!
    Dim samePatch1 As Boolean
    samePatch1 = (a3 Is a1)

    ' 3셀 편집 (값->값/값->빈칸/빈칸 유지 혼합)
    ws.Cells(600, 1).Value = 777
    ws.Cells(700, 1).ClearContents
    ws.Cells(800, 1).Value = 888
    t0 = Timer
    Dim a4 As clsColAsset
    Set a4 = modCache.GetColumnAsset(body)
    Dim patch3Ms As Double
    patch3Ms = (Timer - t0) * 1000!

    LogShim.LogEvent "patchbench", _
        Array("rows", "coldMs", "warmMs", "patch1Ms", "patch1Same", "patch3Ms", _
              "hasKey"), _
        Array(lastRow - 1, Format$(coldMs, "0.0"), Format$(warmMs, "0.0"), _
              Format$(patch1Ms, "0.0"), samePatch1, Format$(patch3Ms, "0.0"), _
              a4.Displays.Contains("12,345,678"))

    Utilities.OptimizedMode False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Utilities.OptimizedMode False
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    LogShim.LogEvent "patchbench_err", Array("num", "desc"), Array(Err.Number, Err.Description)
    Application.EnableEvents = prevEv
End Sub

' 전행 셀별 .Text 1패스 소요(ms)를 잰다.
Private Function TextPass(ByVal body As Range, ByVal nRows As Long) As Double
    Dim t0 As Single, i As Long, s As String
    t0 = Timer
    For i = 1 To nRows
        s = body.Cells(i, 1).Text
    Next i
    TextPass = (Timer - t0) * 1000!
End Function

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

' @desc 조건 수별 분류/배열/해제/필터 시간을 격리 워크북에서 측정함
Public Sub MeasureApplyStages()
    Const ROWS As Long = 91235
    Const UNIQ As Long = 47397
    Dim wb As Workbook, ws As Worksheet, area As Range
    Dim values() As Variant, keys() As Variant, i As Long, n As Variant, rep As Long
    Dim displaySet As clsSet, failure As String
    On Error GoTo Failed
    Utilities.OptimizedMode True
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ReDim values(1 To ROWS + 1, 1 To 1)
    values(1, 1) = "금액"
    For i = 1 To ROWS
        values(i + 1, 1) = ((i - 1) Mod UNIQ + 1) * 7
    Next i
    Set area = ws.Range("A1").Resize(ROWS + 1, 1)
    area.NumberFormat = "#,##0"
    area.Value2 = values
    area.AutoFilter
    LogShim.LogEvent "apply_stages_env", Array("rows", "uniq", "excel", "build"), _
        Array(ROWS, UNIQ, Application.Version, Application.Build)
    ' 행수/고유값은 고정하고 남길 조건 수만 바꾼다. 매번 같은 이전 필터 상태.
    For Each n In Array(1000, 5000, 10000, 20000, 40000, 47396)
        ReDim keys(0 To CLng(n) - 1)
        For i = 1 To CLng(n)
            keys(i - 1) = Format$(i * 7, "#,##0")
        Next i
        Set displaySet = New clsSet
        displaySet.LoadFromArray keys
        For rep = 1 To 3
            MeasureApplyStageCase area, displaySet, CLng(n), rep, ROWS, UNIQ
        Next rep
    Next n
    GoTo Cleanup
Failed:
    failure = CStr(Err.Number) & ": " & Err.Description
Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Utilities.OptimizedMode False
    On Error GoTo 0
    LogShim.LogEvent "apply_stages_done", Array("ok", "error"), Array(failure = "", failure)
End Sub

Private Sub MeasureApplyStageCase(ByVal area As Range, ByVal displaySet As clsSet, _
                                  ByVal n As Long, ByVal rep As Long, _
                                  ByVal rows As Long, ByVal uniq As Long)
    Dim shortItems As New Collection, longItems As New Collection
    Dim vKey As Variant, arr() As Variant, direct() As Variant
    Dim t As Double, classifyMs As Double, indexMs As Double, enumMs As Double
    Dim resetMs As Double, filterMs As Double, i As Long, same As Boolean
    Dim expected As Long, actual As Long
    modApply.ResetFilters area
    area.AutoFilter Field:=1, Criteria1:=Array("7"), Operator:=xlFilterValues

    t = ApplyProbeClock()
    For Each vKey In displaySet.KeysArray
        If Len(CStr(vKey)) > 255 Then
            longItems.Add CStr(vKey)
        Else
            shortItems.Add CStr(vKey)
        End If
    Next vKey
    classifyMs = (ApplyProbeClock() - t) * 1000#

    t = ApplyProbeClock()
    ReDim arr(0 To shortItems.Count - 1)
    For i = 1 To shortItems.Count
        arr(i - 1) = shortItems(i)
    Next i
    indexMs = (ApplyProbeClock() - t) * 1000#

    ' 대조군: 같은 Collection의 순차 열거로 동일 배열을 만든다.
    t = ApplyProbeClock()
    ReDim direct(0 To shortItems.Count - 1)
    i = 0
    For Each vKey In shortItems
        direct(i) = vKey
        i = i + 1
    Next vKey
    enumMs = (ApplyProbeClock() - t) * 1000#
    same = True
    For i = 0 To UBound(arr)
        If arr(i) <> direct(i) Then same = False
    Next i

    t = ApplyProbeClock()
    modApply.ResetFilters area
    resetMs = (ApplyProbeClock() - t) * 1000#
    t = ApplyProbeClock()
    area.AutoFilter Field:=1, Criteria1:=arr, Operator:=xlFilterValues
    filterMs = (ApplyProbeClock() - t) * 1000#

    expected = (rows \ uniq) * n
    If rows Mod uniq < n Then
        expected = expected + rows Mod uniq
    Else
        expected = expected + n
    End If
    ' 분리 측정한 원래 배열 적용 결과와 실제 제품 함수의 가시 셀을 대조한다.
    Dim referenceVisible As String, productVisible As String
    Dim info As Object, productMs As Double
    referenceVisible = area.SpecialCells(xlCellTypeVisible).Address
    modApply.ResetFilters area
    area.AutoFilter Field:=1, Criteria1:=Array("7"), Operator:=xlFilterValues
    t = ApplyProbeClock()
    Set info = modApply.ApplyDisplays(area, 1, displaySet, False, False)
    productMs = (ApplyProbeClock() - t) * 1000#
    productVisible = area.SpecialCells(xlCellTypeVisible).Address
    LogShim.AssertEq "apply_stages_product_rows", referenceVisible, productVisible
    LogShim.AssertEq "apply_stages_product_count", n, CLng(info("Applied"))
    LogShim.AssertEq "apply_stages_product_mode", "array", CStr(info("Mode"))
    actual = modApply.CountVisibleRows(area)
    LogShim.AssertEq "apply_stages_array_equal", True, same
    LogShim.AssertEq "apply_stages_visible", expected, actual
    LogShim.LogEvent "apply_stages", _
        Array("conditions", "rep", "classifyMs", "indexMs", "enumMs", "resetMs", "filterMs", "visible", "equal", "productMs"), _
        Array(n, rep, classifyMs, indexMs, enumMs, resetMs, filterMs, actual, same, productMs)
End Sub

Private Function ApplyProbeClock() As Double
    Dim ticks As Currency, frequency As Currency
    QueryPerformanceCounter ticks
    QueryPerformanceFrequency frequency
    ApplyProbeClock = ticks / frequency
End Function

' @desc 부분일치의 자산/인덱스/매칭/적용 비용과 배열 재사용 후보를 측정함
Public Sub MeasurePartialStages()
    AhoProbeParity
    Dim wb As Workbook, ws As Worksheet, body As Range, area As Range
    Dim vals() As Variant, u As Variant, rep As Long, i As Long, n As Variant, mode As Variant
    Dim a As clsColAsset, displays As clsSet, t As Double, buildMs As Double, indexMs As Double
    Dim failure As String
    On Error GoTo Failed
    Utilities.OptimizedMode True
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set area = ws.Range("A1:A91236")
    Set body = ws.Range("A2:A91236")
    ReDim vals(1 To 91235, 1 To 1)
    ws.Range("A1").Value2 = "검색"
    For Each u In Array(225, 47397)
        For i = 1 To 91235
            vals(i, 1) = "ITEM-" & Format$((i - 1) Mod CLng(u) + 1, "000000") & "-END"
        Next i
        body.NumberFormat = "@"
        body.Value2 = vals
        For rep = 1 To 3
            modApply.ResetFilters area
            modCache.ClearCache
            t = ApplyProbeClock()
            Set a = modCache.GetColumnAsset(body)
            buildMs = (ApplyProbeClock() - t) * 1000#
            t = ApplyProbeClock()
            Set displays = a.Displays
            indexMs = (ApplyProbeClock() - t) * 1000#
            LogShim.LogEvent "partial_setup", Array("uniq", "rep", "buildMs", "indexMs"), _
                Array(u, rep, buildMs, indexMs)
            For Each mode In Array("selective", "miss", "broad", "mixed")
                For Each n In Array(1, 10, 50)
                    If mode <> "broad" Or n <= 10 Then
                        PartialStageCase a, area, CLng(u), rep, CStr(mode), CLng(n)
                    End If
                Next n
            Next mode
        Next rep
    Next u
    GoTo Cleanup
Failed:
    failure = CStr(Err.Number) & ": " & Err.Description
Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    modCache.ClearCache
    Utilities.OptimizedMode False
    On Error GoTo 0
    LogShim.LogEvent "partial_done", Array("ok", "error"), Array(failure = "", failure)
End Sub

Private Sub PartialStageCase(ByVal a As clsColAsset, ByVal area As Range, ByVal uniq As Long, _
                             ByVal rep As Long, ByVal mode As String, ByVal n As Long)
    Dim tokens() As Variant, i As Long, k As Long, keys As Variant, token As String
    ReDim tokens(0 To n - 1)
    For i = 0 To n - 1
        Select Case mode
            Case "selective": tokens(i) = Format$(i + 1, "000000")
            Case "miss": tokens(i) = "MISS" & CStr(i)
            Case "broad": tokens(i) = CStr(i)
            Case "mixed"
                Select Case i Mod 4
                    Case 0: tokens(i) = "item"
                    Case 1: tokens(i) = "MISS" & CStr(i)
                    Case 2: tokens(i) = Format$(i + 1, "000000")
                    Case 3: tokens(i) = "end"
                End Select
        End Select
    Next i
    Dim t As Double, matchMs As Double, copyMs As Double, scanMs As Double, candidateMs As Double
    Dim applyMs As Double, expectedMs As Double, countMs As Double
    Dim r As Object, hitSet As clsSet, ai As Object, hits As Long, found As Boolean, missing As Long
    t = ApplyProbeClock()
    Set r = modMatch.Execute(tokens, a, True)
    matchMs = (ApplyProbeClock() - t) * 1000#

    t = ApplyProbeClock()
    For i = 0 To n - 1
        keys = a.Displays.KeysArray
    Next i
    copyMs = (ApplyProbeClock() - t) * 1000#
    t = ApplyProbeClock()
    For i = 0 To n - 1
        token = CStr(tokens(i))
        For k = LBound(keys) To UBound(keys)
            If InStr(1, CStr(keys(k)), token, vbTextCompare) > 0 Then hits = hits + 1
        Next k
    Next i
    scanMs = (ApplyProbeClock() - t) * 1000#

    ' 키와 검색어가 모두 적중했을 때만 비교를 생략한다.
    Dim keyHit() As Boolean, tokenHit() As Boolean, comparisons As Long
    Dim oracleMissing As Variant
    t = ApplyProbeClock()
    Set hitSet = New clsSet
    keys = a.Displays.KeysArray
    ReDim keyHit(LBound(keys) To UBound(keys))
    ReDim tokenHit(0 To n - 1)
    For i = 0 To n - 1
        found = False
        token = CStr(tokens(i))
        For k = LBound(keys) To UBound(keys)
            If Not (keyHit(k) And found) Then
                comparisons = comparisons + 1
                If InStr(1, CStr(keys(k)), token, vbTextCompare) > 0 Then
                    If Not keyHit(k) Then
                        hitSet.Add keys(k)
                        keyHit(k) = True
                    End If
                    found = True
                End If
            End If
        Next k
        tokenHit(i) = found
        If Not found Then missing = missing + 1
    Next i
    candidateMs = (ApplyProbeClock() - t) * 1000#
    LogShim.AssertEq "partial_candidate_count", r("Matched").Count, hitSet.Count
    LogShim.AssertEq "partial_candidate_keys", CLng(0), r("Matched").Difference(hitSet).Count
    oracleMissing = r("Missing")
    LogShim.AssertEq "partial_candidate_missing", UBound(oracleMissing) - LBound(oracleMissing) + 1, missing
    missing = 0
    For i = 0 To n - 1
        If Not tokenHit(i) Then
            If missing <= UBound(oracleMissing) Then
                LogShim.AssertEq "partial_candidate_missing_token", CStr(oracleMissing(missing)), CStr(tokens(i))
            End If
            missing = missing + 1
        End If
    Next i
    LogShim.LogEvent "partial_skip", Array("uniq", "rep", "mode", "tokens", "baselineMs", "skipMs", "comparisons", "fullComparisons"), _
        Array(uniq, rep, mode, n, matchMs, candidateMs, comparisons, CLng(uniq) * n)

    Dim ahoSet As clsSet, ahoHit() As Boolean, ahoBuild As Double, ahoScan As Double, ahoFallback As Long
    Dim ahoTokens As Variant
    ahoTokens = tokens
    t = ApplyProbeClock()
    Set ahoSet = AhoProbe(keys, ahoTokens, ahoHit, ahoBuild, ahoScan, ahoFallback)
    Dim ahoTotal As Double
    ahoTotal = (ApplyProbeClock() - t) * 1000#
    LogShim.AssertEq "aho_count", r("Matched").Count, ahoSet.Count
    LogShim.AssertEq "aho_keys", CLng(0), r("Matched").Difference(ahoSet).Count
    For i = 0 To n - 1
        LogShim.AssertEq "aho_token", tokenHit(i), ahoHit(i)
    Next i
    LogShim.LogEvent "partial_aho", Array("uniq", "rep", "mode", "tokens", "baselineMs", "ahoMs", "buildMs", "scanMs", "fallbackKeys"), _
        Array(uniq, rep, mode, n, matchMs, ahoTotal, ahoBuild, ahoScan, ahoFallback)

    modApply.ResetFilters area
    t = ApplyProbeClock()
    If hitSet.Count > 0 Then
        Set ai = modApply.ApplyDisplays(area, 1, r("Matched"), False, False)
    Else
        modApply.ApplyHideAll area, 1, False
    End If
    applyMs = (ApplyProbeClock() - t) * 1000#
    Dim expected As Long, visible As Long
    t = ApplyProbeClock()
    expected = a.ExpectedRows(r("Matched"), False)
    expectedMs = (ApplyProbeClock() - t) * 1000#
    t = ApplyProbeClock()
    visible = modApply.CountVisibleRows(area)
    countMs = (ApplyProbeClock() - t) * 1000#
    LogShim.AssertEq "partial_visible", expected, visible
    LogShim.LogEvent "partial_stages", _
        Array("uniq", "rep", "mode", "tokens", "matched", "hitPairs", "matchMs", "copyMs", _
              "scanMs", "candidateMs", "applyMs", "expectedMs", "countMs"), _
        Array(uniq, rep, mode, n, hitSet.Count, hits, matchMs, copyMs, scanMs, candidateMs, _
              applyMs, expectedMs, countMs)
End Sub

' ASCII 전용 실험. 비ASCII/빈 토큰은 기존 비교로 폴백한다.
Private Function AhoProbe(ByRef keys As Variant, ByRef tokens As Variant, _
                          ByRef tokenHit() As Boolean, ByRef buildMs As Double, _
                          ByRef scanMs As Double, ByRef fallbackKeys As Long) As clsSet
    Dim t As Double, cap As Long, i As Long, j As Long, c As Long, s As Long, node As Long
    Dim n As Long, text As String, asciiTokens As Boolean, asciiKey As Boolean
    Dim goToState() As Long, fail() As Long, output() As Long, suffix() As Long
    Dim nextToken() As Long, queue() As Long, terminal As Long, v As Long, f As Long
    Dim head As Long, tail As Long, matched As Boolean, item As Long
    Dim result As New clsSet
    t = ApplyProbeClock()
    n = UBound(tokens) - LBound(tokens) + 1
    ReDim tokenHit(LBound(tokens) To UBound(tokens))
    asciiTokens = True
    For i = LBound(tokens) To UBound(tokens)
        text = CStr(tokens(i))
        If Len(text) = 0 Then asciiTokens = False
        If Not AhoAscii(text) Then asciiTokens = False
        cap = cap + Len(text)
    Next i
    If asciiTokens Then
        ReDim goToState(0 To cap, 0 To 127)
        ReDim fail(0 To cap)
        ReDim output(0 To cap)
        ReDim suffix(0 To cap)
        ReDim nextToken(1 To n)
        ReDim queue(0 To cap)
        For i = LBound(tokens) To UBound(tokens)
            text = CStr(tokens(i))
            s = 0
            For j = 1 To Len(text)
                c = AscW(Mid$(text, j, 1))
                If c >= 65 And c <= 90 Then c = c + 32
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
    End If
    buildMs = (ApplyProbeClock() - t) * 1000#
    t = ApplyProbeClock()
    For i = LBound(keys) To UBound(keys)
        text = CStr(keys(i))
        matched = False
        asciiKey = False
        If asciiTokens Then asciiKey = AhoAscii(text)
        If asciiKey Then
            s = 0
            For j = 1 To Len(text)
                c = AscW(Mid$(text, j, 1))
                If c >= 65 And c <= 90 Then c = c + 32
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
            fallbackKeys = fallbackKeys + 1
            For j = LBound(tokens) To UBound(tokens)
                If InStr(1, text, CStr(tokens(j)), vbTextCompare) > 0 Then
                    tokenHit(j) = True
                    matched = True
                End If
            Next j
        End If
        If matched Then result.Add keys(i)
    Next i
    scanMs = (ApplyProbeClock() - t) * 1000#
    Set AhoProbe = result
End Function

Private Function AhoAscii(ByVal text As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1))
        If c < 0 Or c > 127 Then Exit Function
    Next i
    AhoAscii = True
End Function

Private Sub AhoProbeParity()
    Dim keys As Variant, tokens As Variant, cases As Variant, tc As Variant
    Dim hit() As Boolean, expectedHit() As Boolean, actual As clsSet, expected As clsSet
    Dim i As Long, j As Long, buildMs As Double, scanMs As Double, fallback As Long
    keys = Array("ushers", "SHE", "aaaa", "ababa", "ABC", "", "한글ABC", ChrW(&HFF21) & "BC")
    cases = Array(Array("he", "she", "hers", "his", "a", "aa", "aba", "ABC", "abc", "MISS"), _
                  Array("한글", "ABC", "없는값"), Array("", "ABC"), Array())
    For Each tc In cases
        tokens = tc
        Set expected = New clsSet
        If UBound(tokens) >= LBound(tokens) Then
            ReDim expectedHit(LBound(tokens) To UBound(tokens))
            For i = LBound(keys) To UBound(keys)
                For j = LBound(tokens) To UBound(tokens)
                    If InStr(1, CStr(keys(i)), CStr(tokens(j)), vbTextCompare) > 0 Then
                        expected.Add keys(i)
                        expectedHit(j) = True
                    End If
                Next j
            Next i
            fallback = 0
            Set actual = AhoProbe(keys, tokens, hit, buildMs, scanMs, fallback)
            LogShim.AssertEq "aho_parity_count", expected.Count, actual.Count
            LogShim.AssertEq "aho_parity_keys", CLng(0), expected.Difference(actual).Count
            For j = LBound(tokens) To UBound(tokens)
                LogShim.AssertEq "aho_parity_token", expectedHit(j), hit(j)
            Next j
        End If
    Next tc
End Sub

' @desc 검색어 수와 표시문자열 길이에 따른 Aho 전환점을 측정함
Public Sub MeasureAhoCrossover()
    AhoCrossoverGrid Array(225, 47397), Array(15, 64, 128), Array(1, 2, 3, 5, 8, 10)
End Sub

' @desc 긴 표시문자열에서 다검색어 Aho 전환점을 추가 측정함
Public Sub MeasureAhoLongCrossover()
    AhoCrossoverGrid Array(47397), Array(64, 128), Array(20, 50)
End Sub

Private Sub AhoCrossoverGrid(ByVal sizes As Variant, ByVal widths As Variant, ByVal counts As Variant)
    Dim u As Variant, width As Variant, count As Variant, mode As Variant
    Dim values() As Variant, tokens As Variant, i As Long, rep As Long
    Dim displays As clsSet, asset As clsColAsset, emptyMap As Object
    Dim failure As String
    On Error GoTo Failed
    AhoProbeParity
    For Each u In sizes
        For Each width In widths
            ReDim values(0 To CLng(u) - 1)
            For i = 0 To CLng(u) - 1
                values(i) = "ITEM-" & String$(CLng(width) - 15, "x") & Format$(i + 1, "000000") & "-END"
            Next i
            Set displays = New clsSet
            displays.LoadFromArray values
            Set emptyMap = CreateObject("Scripting.Dictionary")
            Set asset = New clsColAsset
            asset.Init displays, emptyMap, emptyMap, 0, CLng(u), "실험", Nothing
            For Each mode In Array("selective", "miss", "broad")
                For Each count In counts
                    ReDim tokens(0 To CLng(count) - 1)
                    For i = 0 To CLng(count) - 1
                        Select Case CStr(mode)
                            Case "selective": tokens(i) = Format$(i + 1, "000000")
                            Case "miss": tokens(i) = "MISS" & Format$(i, "00")
                            Case "broad": tokens(i) = CStr(i)
                        End Select
                    Next i
                    For rep = 1 To 5
                        AhoCrossoverCase asset, tokens, CLng(u), CLng(width), CStr(mode), rep
                    Next rep
                Next count
            Next mode
        Next width
    Next u
    GoTo Done
Failed:
    failure = CStr(Err.Number) & ": " & Err.Description
Done:
    LogShim.LogEvent "aho_crossover_done", Array("ok", "error"), Array(failure = "", failure)
End Sub

Private Sub AhoCrossoverCase(ByVal asset As clsColAsset, ByRef tokens As Variant, _
                             ByVal uniq As Long, ByVal width As Long, ByVal mode As String, ByVal rep As Long)
    Dim r As Object, actual As clsSet, keys As Variant, hit() As Boolean
    Dim t As Double, baselineMs As Double, ahoMs As Double, buildMs As Double, scanMs As Double
    Dim fallback As Long, phase As Long, baselineFirst As Boolean
    baselineFirst = (rep Mod 2 = 1)
    For phase = 0 To 1
        If (phase = 0) = baselineFirst Then
            t = ApplyProbeClock()
            Set r = modMatch.Execute(tokens, asset, True)
            baselineMs = (ApplyProbeClock() - t) * 1000#
        Else
            t = ApplyProbeClock()
            keys = asset.Displays.KeysArray
            Set actual = AhoProbe(keys, tokens, hit, buildMs, scanMs, fallback)
            ahoMs = (ApplyProbeClock() - t) * 1000#
        End If
    Next phase
    LogShim.AssertEq "crossover_count", r("Matched").Count, actual.Count
    LogShim.AssertEq "crossover_keys", CLng(0), r("Matched").Difference(actual).Count
    LogShim.AssertEq "crossover_ascii", CLng(0), fallback
    Dim missing As Object, item As Variant, i As Long
    Set missing = CreateObject("Scripting.Dictionary")
    missing.CompareMode = vbTextCompare
    For Each item In r("Missing")
        missing(CStr(item)) = True
    Next item
    For i = LBound(tokens) To UBound(tokens)
        LogShim.AssertEq "crossover_token", Not missing.exists(CStr(tokens(i))), hit(i)
    Next i
    LogShim.LogEvent "aho_crossover", _
        Array("uniq", "length", "mode", "tokens", "rep", "baselineFirst", "baselineMs", "ahoMs", "buildMs", "scanMs"), _
        Array(uniq, width, mode, UBound(tokens) + 1, rep, baselineFirst, baselineMs, ahoMs, buildMs, scanMs)
End Sub

' @desc 표본 계측으로 검색 경로를 고르는 혼합 후보를 검증함
Public Sub MeasureHybridSearch()
    HybridSearchGrid "sample"
End Sub

' @desc 제품 혼합 검색의 결과와 전체 실행 시간을 검증함
Public Sub MeasureHybridProductSearch()
    HybridSearchGrid "product"
End Sub

' @desc 제품의 소수 검색어와 유니코드 폴백 비용을 반복 검증함
Public Sub MeasureHybridProductFallback()
    HybridSearchGrid "product-fallback"
End Sub

' @desc 길이와 검색어 수를 쓰는 저비용 혼합 규칙을 검증함
Public Sub MeasureHybridStaticSearch()
    HybridSearchGrid "length"
End Sub

Private Sub HybridSearchGrid(ByVal policy As String)
    Dim cases As Variant, spec As Variant, parts As Variant, rep As Long, i As Long, width As Long
    Dim uniq As Long, n As Long, mode As String, tokens As Variant, values() As Variant, repeats As Long
    Dim displays As clsSet, asset As clsColAsset, emptyMap As Object, failure As String
    On Error GoTo Failed
    AhoProbeParity
    cases = Array("225|15|1|selective", "225|128|50|miss", _
        "47397|15|1|selective", "47397|15|2|miss", "47397|15|5|selective", _
        "47397|15|10|selective", "47397|15|50|miss", "47397|15|3|broad", _
        "47397|64|10|miss", "47397|64|20|selective", "47397|128|20|selective", _
        "47397|128|50|miss", "47397|128|8|broad", _
        "20000|32|6|selective", "20000|32|12|miss", "20000|96|12|broad", _
        "20000|96|30|selective", "20000|128|20|cluster", "20000|64|20|unicode")
    repeats = 3
    If policy = "product-fallback" Then
        cases = Array("47397|15|5|selective", "20000|32|6|selective", "20000|64|20|unicode")
        repeats = 9
    End If
    For Each spec In cases
        parts = Split(CStr(spec), "|")
        uniq = CLng(parts(0)): width = CLng(parts(1)): n = CLng(parts(2)): mode = CStr(parts(3))
        ReDim values(0 To uniq - 1)
        For i = 0 To uniq - 1
            If mode = "cluster" And i < uniq \ 2 Then
                values(i) = "ITEM-" & Format$(i + 1, "000000") & "-END"
            Else
                values(i) = "ITEM-" & String$(width - 15, "x") & Format$(i + 1, "000000") & "-END"
            End If
            If mode = "unicode" Then values(i) = "한글" & values(i)
        Next i
        Set displays = New clsSet
        displays.LoadFromArray values
        Set emptyMap = CreateObject("Scripting.Dictionary")
        Set asset = New clsColAsset
        asset.Init displays, emptyMap, emptyMap, 0, uniq, "혼합 실험", Nothing
        ReDim tokens(0 To n - 1)
        For i = 0 To n - 1
            Select Case mode
                Case "miss": tokens(i) = "MISS" & Format$(i, "00")
                Case "broad": tokens(i) = CStr(i)
                Case "unicode": tokens(i) = "한글" & CStr(i)
                Case Else: tokens(i) = Format$(i + 1, "000000")
            End Select
        Next i
        For rep = 1 To repeats
            HybridSearchCase asset, tokens, CStr(spec), rep, policy
        Next rep
    Next spec
    GoTo Done
Failed:
    failure = CStr(Err.Number) & ": " & Err.Description
Done:
    LogShim.LogEvent "hybrid_done", Array("ok", "error"), Array(failure = "", failure)
End Sub

' 후보도 현행과 같은 결과 구조를 만들어 반환 비용을 포함한다.
Private Function AhoOutcome(ByRef keys As Variant, ByRef tokens As Variant) As Object
    Dim hits() As Boolean, buildMs As Double, scanMs As Double, fallback As Long
    Dim matched As clsSet, result As Object, corrections As Object, missing As Collection
    Dim arr As Variant, i As Long
    Set result = CreateObject("Scripting.Dictionary")
    Set corrections = CreateObject("Scripting.Dictionary")
    Set missing = New Collection
    Set matched = AhoProbe(keys, tokens, hits, buildMs, scanMs, fallback)
    For i = LBound(tokens) To UBound(tokens)
        If Not hits(i) Then missing.Add CStr(tokens(i))
    Next i
    arr = Array()
    If missing.Count > 0 Then
        ReDim arr(0 To missing.Count - 1)
        For i = 1 To missing.Count
            arr(i - 1) = missing(i)
        Next i
    End If
    Set result("Matched") = matched
    Set result("Corrections") = corrections
    result("Missing") = arr
    result("InputCount") = UBound(tokens) - LBound(tokens) + 1
    result("MatchedRows") = -1&
    Set AhoOutcome = result
End Function

Private Function HybridOutcome(ByVal asset As clsColAsset, ByRef tokens As Variant, _
                               ByRef choice As String, ByRef chooseMs As Double, _
                               ByRef nativeSampleMs As Double, ByRef ahoSampleMs As Double, _
                               ByVal reverseOrder As Boolean) As Object
    Const SAMPLE_SIZE As Long = 1024
    Dim t As Double, begin As Double, keys As Variant, sample As Variant, i As Long, phase As Long
    Dim displays As clsSet, sampleAsset As clsColAsset, emptyMap As Object, ignored As Object
    begin = ApplyProbeClock()
    choice = "instr"
    If UBound(tokens) - LBound(tokens) < 1 Then GoTo Selected
    If asset.Displays.Count <= SAMPLE_SIZE Then GoTo Selected
    For i = LBound(tokens) To UBound(tokens)
        If Len(CStr(tokens(i))) = 0 Then GoTo Selected
        If Not AhoAscii(CStr(tokens(i))) Then GoTo Selected
    Next i
    keys = asset.Displays.KeysArray
    ReDim sample(0 To SAMPLE_SIZE - 1)
    For i = 0 To SAMPLE_SIZE - 1
        sample(i) = keys(LBound(keys) + CLng(Fix(CDbl(i) * (UBound(keys) - LBound(keys)) / (SAMPLE_SIZE - 1))))
    Next i
    Set displays = New clsSet
    displays.LoadFromArray sample
    Set emptyMap = CreateObject("Scripting.Dictionary")
    Set sampleAsset = New clsColAsset
    sampleAsset.Init displays, emptyMap, emptyMap, 0, SAMPLE_SIZE, "표본", Nothing
    For phase = 0 To 1
        t = ApplyProbeClock()
        If (phase = 0) = reverseOrder Then
            Set ignored = modMatch.Execute(tokens, sampleAsset, True)
            nativeSampleMs = (ApplyProbeClock() - t) * 1000#
        Else
            Set ignored = AhoOutcome(sample, tokens)
            ahoSampleMs = (ApplyProbeClock() - t) * 1000#
        End If
    Next phase
    ' 표본 차이가 작으면 현행을 유지한다. 이 여유율도 실험 대상이다.
    If ahoSampleMs < nativeSampleMs * 0.8 Then choice = "aho"
Selected:
    chooseMs = (ApplyProbeClock() - begin) * 1000#
    If choice = "aho" Then
        Set HybridOutcome = AhoOutcome(keys, tokens)
    Else
        Set HybridOutcome = modMatch.Execute(tokens, asset, True)
    End If
End Function

Private Sub HybridSearchCase(ByVal asset As clsColAsset, ByRef tokens As Variant, ByVal tag As String, ByVal rep As Long, ByVal policy As String)
    Dim baseline As Object, aho As Object, hybrid As Object, keys As Variant, t As Double
    Dim baselineMs As Double, ahoMs As Double, hybridMs As Double, chooseMs As Double
    Dim nativeSampleMs As Double, ahoSampleMs As Double, choice As String, phase As Long
    For phase = 0 To 2
        t = ApplyProbeClock()
        Select Case (rep + phase - 1) Mod 3
            Case 0
                Set baseline = NativePartialOutcome(asset, tokens)
                baselineMs = (ApplyProbeClock() - t) * 1000#
            Case 1
                keys = asset.Displays.KeysArray
                Set aho = AhoOutcome(keys, tokens)
                ahoMs = (ApplyProbeClock() - t) * 1000#
            Case 2
                If policy = "product" Or policy = "product-fallback" Then
                    Set hybrid = modMatch.Execute(tokens, asset, True)
                    choice = "product"
                ElseIf policy = "length" Then
                    Set hybrid = LengthHybridOutcome(asset, tokens, choice, chooseMs)
                Else
                    Set hybrid = HybridOutcome(asset, tokens, choice, chooseMs, nativeSampleMs, ahoSampleMs, rep Mod 2 = 0)
                End If
                hybridMs = (ApplyProbeClock() - t) * 1000#
        End Select
    Next phase
    AssertHybridOutcome baseline, aho
    AssertHybridOutcome baseline, hybrid
    LogShim.LogEvent "hybrid_case", _
        Array("case", "rep", "baselineMs", "ahoMs", "hybridMs", "choice", "chooseMs", "nativeSampleMs", "ahoSampleMs", "policy"), _
        Array(tag, rep, baselineMs, ahoMs, hybridMs, choice, chooseMs, nativeSampleMs, ahoSampleMs, policy)
End Sub

Private Sub AssertHybridOutcome(ByVal expected As Object, ByVal actual As Object)
    LogShim.AssertEq "hybrid_count", expected("Matched").Count, actual("Matched").Count
    LogShim.AssertEq "hybrid_keys", CLng(0), expected("Matched").Difference(actual("Matched")).Count
    LogShim.AssertEq "hybrid_missing", Join(expected("Missing"), vbNullChar), Join(actual("Missing"), vbNullChar)
    LogShim.AssertEq "hybrid_input", expected("InputCount"), actual("InputCount")
    LogShim.AssertEq "hybrid_rows", expected("MatchedRows"), actual("MatchedRows")
    LogShim.AssertEq "hybrid_corrections", expected("Corrections").Count, actual("Corrections").Count
End Sub

' 기존 측정으로 정한 보수적 실험 규칙. 제품용 확정 기준이 아니다.
Private Function LengthHybridOutcome(ByVal asset As clsColAsset, ByRef tokens As Variant, _
                                     ByRef choice As String, ByRef chooseMs As Double) As Object
    Dim begin As Double, n As Long, keys As Variant, i As Long, totalLength As Double
    Dim avgLength As Double, threshold As Long
    begin = ApplyProbeClock()
    choice = "instr"
    n = UBound(tokens) - LBound(tokens) + 1
    If n <= 5 Then GoTo Selected
    If asset.Displays.Count <= 1024 Then GoTo Selected
    For i = LBound(tokens) To UBound(tokens)
        If Len(CStr(tokens(i))) = 0 Then GoTo Selected
        If Not AhoAscii(CStr(tokens(i))) Then GoTo Selected
    Next i
    keys = asset.Displays.KeysArray
    For i = 0 To 31
        totalLength = totalLength + Len(CStr(keys(LBound(keys) + CLng(Fix(CDbl(i) * (UBound(keys) - LBound(keys)) / 31)))))
    Next i
    avgLength = totalLength / 32
    If avgLength <= 16 Then
        threshold = 10
    ElseIf avgLength <= 64 Then
        threshold = 20
    Else
        threshold = 50
    End If
    If n >= threshold Then choice = "aho"
Selected:
    chooseMs = (ApplyProbeClock() - begin) * 1000#
    If choice = "aho" Then
        Set LengthHybridOutcome = AhoOutcome(keys, tokens)
    Else
        Set LengthHybridOutcome = modMatch.Execute(tokens, asset, True)
    End If
End Function

' 변경 전 부분일치 기준. 제품 선택 로직을 호출하지 않는다.
Private Function NativePartialOutcome(ByVal asset As clsColAsset, ByRef tokens As Variant) As Object
    Dim result As Object, matched As New clsSet, corrections As Object, missing As New Collection
    Dim keys As Variant, i As Long, t As String, arr As Variant
    Set result = CreateObject("Scripting.Dictionary")
    Set corrections = CreateObject("Scripting.Dictionary")
    keys = asset.Displays.KeysArray
    For i = LBound(tokens) To UBound(tokens)
        t = CStr(tokens(i))
        If Not NativePartialToken(t, asset, True, matched, corrections, keys) Then missing.Add t
    Next i
    arr = Array()
    If missing.Count > 0 Then
        ReDim arr(0 To missing.Count - 1)
        For i = 1 To missing.Count
            arr(i - 1) = missing(i)
        Next i
    End If
    Set result("Matched") = matched
    Set result("Corrections") = corrections
    result("Missing") = arr
    result("InputCount") = UBound(tokens) - LBound(tokens) + 1
    result("MatchedRows") = -1&
    Set NativePartialOutcome = result
End Function

' 기존 제품과 같은 인자 전달 및 배열 순회 비용을 유지한다.
Private Function NativePartialToken(ByVal t As String, ByVal asset As clsColAsset, _
                                    ByVal bPartial As Boolean, ByVal matched As clsSet, _
                                    ByVal corrections As Object, ByRef keys As Variant) As Boolean
    Dim hit As Boolean, k As Long
    If bPartial Then
        For k = LBound(keys) To UBound(keys)
            If InStr(1, CStr(keys(k)), t, vbTextCompare) > 0 Then
                matched.Add keys(k)
                hit = True
            End If
        Next k
    End If
    NativePartialToken = hit
End Function
' @desc 실제 중단 키 검증용 합성 열과 필터 폼을 준비함
Public Sub PrepareInterruptProbe()
    Dim ws As Worksheet, data() As Variant, i As Long, tokens As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("UF_InterruptProbe")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "UF_InterruptProbe"
    End If
    ws.Cells.Clear
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ws.Range("B1").Value2 = CLng(Application.Cursor)
    ws.Range("B2").Value2 = Application.ScreenUpdating
    ws.Range("B3").Value2 = CLng(Application.Calculation)
    ws.Range("B4").Value2 = Application.EnableEvents
    ws.Range("B5").Value2 = CLng(Application.EnableCancelKey)
    ws.Range("A1").Value2 = "중단 검증"
    ReDim data(1 To 50000, 1 To 1)
    For i = 1 To 50000
        data(i, 1) = String$(240, "x") & "한글" & Format$(i, "000000")
    Next i
    ws.Range("A2:A50001").Value2 = data
    For i = 1 To 2000
        If i > 1 Then tokens = tokens & vbLf
        tokens = tokens & "미일치" & CStr(i)
    Next i
    ws.Activate
    modCache.ClearCache
    Dim frm As frmFilter01
    Set frm = New frmFilter01
    Set frm.InjectBodyRange = ws.Range("A2:A50001")
    Set frm.InjectFilterArea = ws.Range("A1:A50001")
    frm.InjectRelativeCol = 1
    frm.InjectHeaderName = "중단 검증"
    frm.txtFilterValues.text = tokens
    frm.Show vbModeless
    LogShim.LogEvent "interrupt_before", _
        Array("cursor", "screen", "calculation", "events", "cancelKey"), _
        Array(Application.Cursor, Application.ScreenUpdating, Application.Calculation, _
              Application.EnableEvents, Application.EnableCancelKey)
End Sub

' @desc 실제 키 중단 뒤 전역 설정을 관측하고 준비 시점과 대조함
Public Sub InspectInterruptProbe()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("UF_InterruptProbe")
    LogShim.LogEvent "interrupt_after", _
        Array("cursor", "screen", "calculation", "events", "cancelKey"), _
        Array(Application.Cursor, Application.ScreenUpdating, Application.Calculation, _
              Application.EnableEvents, Application.EnableCancelKey)
    LogShim.AssertEq "interrupt_cursor", CLng(ws.Range("B1").Value2), CLng(Application.Cursor)
    LogShim.AssertEq "interrupt_screen", CBool(ws.Range("B2").Value2), Application.ScreenUpdating
    LogShim.AssertEq "interrupt_calculation", CLng(ws.Range("B3").Value2), CLng(Application.Calculation)
    LogShim.AssertEq "interrupt_events", CBool(ws.Range("B4").Value2), Application.EnableEvents
End Sub

' @desc 중단 프로브의 설정과 스크래치 시트를 정리함
Public Sub CleanupInterruptProbe()
    Dim ws As Worksheet, frm As Object, alerts As Boolean
    Set ws = ThisWorkbook.Worksheets("UF_InterruptProbe")
    For Each frm In VBA.UserForms
        Unload frm
    Next frm
    Utilities.OptimizedMode False
    Application.Cursor = CLng(ws.Range("B1").Value2)
    Application.Calculation = CLng(ws.Range("B3").Value2)
    Application.EnableEvents = CBool(ws.Range("B4").Value2)
    Application.ScreenUpdating = CBool(ws.Range("B2").Value2)
    Application.EnableCancelKey = CLng(ws.Range("B5").Value2)
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = alerts
    modCache.ClearCache
End Sub
