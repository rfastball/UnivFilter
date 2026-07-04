Attribute VB_Name = "modTester"
Option Explicit

' ==============================================================================
' [서식 검증기] UnivFilter Validator (Complete Log & No-Symbol Accounting)
' 설명: 모든 테스트 결과(성공/실패)를 로그에 남기고, 실패 건만 실험용 필터를 생성함.
'       통화 기호가 없는 회계 서식 테스트가 추가됨.
'       [v2.3] GetColumnIdentity 강화에 맞춰 테스트 데이터 서식 통일 로직 적용
' ==============================================================================

Private wsSpec As Worksheet     ' 테스트 수행용 임시 시트
Private loSpec As ListObject
Private wsReport As Worksheet   ' 리포트 시트
Private NextLogRow As Long      ' 로그 기록용 행 위치 (좌측)
Private NextExpCol As Long      ' 실험 데이터 기록용 열 위치 (우측)
Private sReport As String
Private iPass As Long, iFail As Long

Public Sub Run_SpecValidation()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    ' 1. 초기화
    SetupSpecSheet
    SetupReportSheet
    
    sReport = "==== [UnivFilter 검증 리포트 (통합 로그)] ====" & vbCrLf & vbCrLf
    iPass = 0: iFail = 0
    
' --------------------------------------------------------------------------
    ' [A. 숫자]
    ' --------------------------------------------------------------------------
    Test_Spec "숫자_콤마_완전", 3000, "#,##0", "3,000", "Exact"
    Test_Spec "숫자_값_완전", 3000, "#,##0", "3000", "Exact"
    Test_Spec "숫자_콤마_부분", 3000, "#,##0", "3,0", "Partial"
    Test_Spec "숫자_00_부분", 3000, "#,##0", "00", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [B. 소수점]
    ' --------------------------------------------------------------------------
    Test_Spec "소수_0채움_완전", 1234.5, "#,##0.00", "1,234.50", "Exact"
    Test_Spec "소수_이하_부분", 1234.5, "#,##0.00", ".50", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [C. 통화]
    ' --------------------------------------------------------------------------
    Test_Spec "통화_기호_완전", 1000, "$#,##0", "###VIEW###", "Exact"
    Test_Spec "통화_기호_부분", 1000, "$#,##0", "###SYMBOL###", "Partial", True
    Test_Spec "통화_숫자_부분", 1000, "$#,##0", "1,0", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [D. 회계]
    ' --------------------------------------------------------------------------
    Dim sAccFmt As String
    sAccFmt = "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""_-;_-@_-"
    Test_Spec "회계_기호_완전", 50000, sAccFmt, "###VIEW###", "Exact"
    Test_Spec "회계_기호_부분", 50000, sAccFmt, "###SYMBOL###", "Partial", True
    
    Dim sAccNoSym As String
    sAccNoSym = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
    Test_Spec "회계_무기호_완전", 50000, sAccNoSym, "###VIEW###", "Exact"
    Test_Spec "회계_무기호_부분", 50000, sAccNoSym, "50", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [E. 날짜]
    ' --------------------------------------------------------------------------
    Dim dDate As Date: dDate = DateSerial(2025, 12, 25)
    'Test_Spec "날짜_표준_완전", dDate, "yyyy-mm-dd", "2025-12-25", "Exact"
    Test_Spec "날짜_월일_부분", dDate, "yyyy-mm-dd", "12-25", "Partial"
    Test_Spec "날짜_연도_부분", dDate, "yyyy-mm-dd", "2025", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [F. 한글 날짜]
    ' --------------------------------------------------------------------------
    'Test_Spec "날짜_한글_완전", dDate, "yyyy""년"" mm""월"" dd""일""", "###VIEW###", "Exact"
    Test_Spec "날짜_한글_부분", dDate, "yyyy""년"" mm""월"" dd""일""", "12월", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [G. 백분율]
    ' --------------------------------------------------------------------------
    Test_Spec "백분율_기호_완전", 0.5, "0%", "50%", "Exact"
    Test_Spec "백분율_기호_부분", 0.5, "0%", "%", "Partial", True
    Test_Spec "백분율_숫자_부분", 0.5, "0%", "5", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [H. 접미사]
    ' --------------------------------------------------------------------------
    Test_Spec "접미사_전체_완전", 10, "0"" kg""", "10 kg", "Exact"
    Test_Spec "접미사_단위_부분", 10, "0"" kg""", "kg", "Partial", True
    
    ' --------------------------------------------------------------------------
    ' [I. 이중 검색 (부분일치 심화)]
    ' --------------------------------------------------------------------------
    Test_Spec "이중_겉_검색", 12345, "#,##0", "2,3", "Partial"
    Test_Spec "이중_속_검색", 12345, "#,##0", "234", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [J. 논리값]
    ' --------------------------------------------------------------------------
    Test_Spec "논리_TRUE_완전", True, "General", "TRUE", "Exact"
    Test_Spec "논리_FALSE_완전", False, "General", "false", "Exact"
    Test_Spec "논리_TRUE_부분", True, "General", "TR", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [K. 회계/통화 음수]
    ' --------------------------------------------------------------------------
    Test_Spec "회계_음수_기호", -5000, "_-$* #,##0_-;-$* #,##0_-", "###VIEW###", "Exact"
    Test_Spec "통화_음수_기호", -5000, "$#,##0;[Red]-$#,##0", "###VIEW###", "Exact"
    
    ' --------------------------------------------------------------------------
    ' [L. 회계 제로]
    ' --------------------------------------------------------------------------
    Test_Spec "회계_제로_표시", 0, sAccFmt, "-", "Exact"
    Test_Spec "회계_제로_숫자검색", 0, sAccFmt, "0", "Exact"
    
    ' --------------------------------------------------------------------------
    ' [M. 과학적 표기법 & 조단위]
    ' --------------------------------------------------------------------------
    Test_Spec "과학_지수_완전", 123000, "0.00E+00", "1.23E+05", "Exact"
    Test_Spec "조단위_대금_콤마", 1234567890000#, "$#,##0", "###VIEW###", "Exact"
    
    ' --------------------------------------------------------------------------
    ' [N. 날짜/시간 심화]
    ' --------------------------------------------------------------------------
    Dim dDateOnly As Date: dDateOnly = DateSerial(2026, 1, 25)
    'Test_Spec "날짜_요일_완전", dDateOnly, "yyyy-mm-dd (ddd)", "###VIEW###", "Exact"
    Dim dDateTime As Date: dDateTime = DateSerial(2026, 1, 25) + TimeSerial(16, 30, 0)
    'Test_Spec "날짜_시간_완전", dDateTime, "yyyy-mm-dd hh:mm", "###VIEW###", "Exact"
    
    ' --------------------------------------------------------------------------
    ' [O. 텍스트 공백]
    ' --------------------------------------------------------------------------
    Test_Spec "텍스트_숫자_완전", "\1,000", "@", "\1,000", "Exact"
    Test_Spec "텍스트_공백_부분", "  \1,000  ", "@", "1,000", "Partial"
    
    ' --------------------------------------------------------------------------
    ' [P. Value Matching (값 기반 검색)]
    ' --------------------------------------------------------------------------
    Test_Spec "통화_값_완전매칭", 1000, "$#,##0", "1000", "Exact"
    Test_Spec "회계_값_완전매칭", 50000, sAccFmt, "50000", "Exact"
    Test_Spec "음수괄호_값_검색", -1234, "$#,##0;($#,##0)", "-1234", "Exact"
    Test_Spec "소수점_값_검색", 1234.5, "#,##0.00", "1234.5", "Exact"
    Test_Spec "백분율_값_검색", 0.5, "0%", "0.5", "Exact"
    'Test_Spec "날짜_연속숫자_검색", dDate, "yyyy-mm-dd", "20251225", "Exact"
    
    ' ... (종료 처리 로직 동일) ...
    ' --------------------------------------------------------------------------
    Application.ScreenUpdating = True
    
    wsReport.Activate
    wsReport.Columns("A:F").AutoFit
    
    If iFail > 0 Then
        If NextExpCol > 9 Then
            Dim rngExpHeader As Range
            Set rngExpHeader = wsReport.Range(wsReport.Cells(1, 9), wsReport.Cells(1, NextExpCol - 1))
            wsReport.Rows(1).AutoFit
            rngExpHeader.Resize(3).AutoFilter
            wsReport.Range("I1").Select
        End If
        MsgBox "검증 완료!" & vbCrLf & "성공: " & iPass & " / 실패: " & iFail, vbExclamation
    Else
        MsgBox "완벽합니다! (All Pass)", vbInformation
    End If
End Sub

' ==============================================================================
' [Worker] Test_Spec (파라미터 추가 및 검증 로직 유연화)
' ==============================================================================
Private Sub Test_Spec( _
    sTitle As String, vValue As Variant, sFmt As String, _
    sKeyword As String, sType As String, _
    Optional bExpectNoise As Boolean = False) ' <--- [New] 노이즈 허용 여부 플래그
    
    ' ... (데이터 세팅 1~2단계 동일) ...
    If Not loSpec.DataBodyRange Is Nothing Then
        Application.DisplayAlerts = False
        loSpec.DataBodyRange.Delete
        Application.DisplayAlerts = True
    End If
    
    With wsSpec.Range("A2")
        .Value = vValue
        .NumberFormat = sFmt
    End With
    wsSpec.Columns("A").AutoFit
    
    Dim sActualView As String: sActualView = wsSpec.Range("A2").text
    If sKeyword = "###VIEW###" Then sKeyword = sActualView
    If sKeyword = "###SYMBOL###" Then sKeyword = Left(Trim(sActualView), 1)
    
    ' [중요] 노이즈 데이터도 타겟과 동일한 서식 적용 (GetColumnIdentity 통과용)
    Dim sNoiseFmt As String: sNoiseFmt = sFmt
    Dim vNoise As Variant: vNoise = GetSafeNoise(vValue)
    
    With wsSpec.Range("A3")
        .Value = vNoise
        .NumberFormat = sNoiseFmt
    End With
    
    ' ... (엔진 호출 및 가동 3~4단계 동일) ...
    modCache.InvalidateCache
    Dim colAsset As clsColumnAsset
    On Error Resume Next
    Set colAsset = modCache.GetColumnAsset(wsSpec.Range("A2:A3"))
    On Error GoTo 0
    If colAsset Is Nothing Then: iFail = iFail + 1: Exit Sub
    
    Dim bExact As Boolean: bExact = (sType = "Exact")
    Dim dictResult As Object
    Set dictResult = modMatchingEngine.CalculateCriteria(Array(sKeyword), colAsset, bExact, False)
    If dictResult Is Nothing Then: iFail = iFail + 1: Exit Sub
    
    modResultHandler.ApplyFilterResult loSpec.Range, 1, dictResult, False
    
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
    
    ' 5. [수정됨] 정밀 판정 로직
    Dim isTargetVisible As Boolean, isNoiseHidden As Boolean
    
    isTargetVisible = (wsSpec.Rows(2).Hidden = False)
    
    ' bExpectNoise가 True면 노이즈가 보이는 것(Hidden=False)이 정상
    If bExpectNoise Then
        isNoiseHidden = (wsSpec.Rows(3).Hidden = False)
    Else
        isNoiseHidden = (wsSpec.Rows(3).Hidden = True)
    End If
    
    Dim sResultStatus As String, sReason As String
    
    If isTargetVisible And isNoiseHidden Then
        iPass = iPass + 1
        sResultStatus = "PASS"
        sReason = "Success"
        sReport = sReport & "[PASS] " & sTitle & vbCrLf
    Else
        iFail = iFail + 1
        sResultStatus = "FAIL"
        If Not isTargetVisible Then sReason = "Target Missing"
        
        ' 실패 원인 메시지도 상황에 맞게 출력
        If bExpectNoise Then
            If wsSpec.Rows(3).Hidden = True Then sReason = sReason & ", Noise Should Be Visible"
        Else
            If wsSpec.Rows(3).Hidden = False Then sReason = sReason & ", Noise Visible"
        End If
        
        sReport = sReport & "[FAIL] " & sTitle & " (" & sReason & ")" & vbCrLf
    End If
    
    Record_Result sTitle, vValue, sFmt, vNoise, sNoiseFmt, sKeyword, sType, sActualView, sReason, sResultStatus
End Sub

' ------------------------------------------------------------------------------
' [Helper] 결과 기록기 (로그는 전체, 실험표는 실패 시에만)
' ------------------------------------------------------------------------------
Private Sub Record_Result( _
    sTitle As String, vTarget As Variant, sTargetFmt As String, _
    vNoise As Variant, sNoiseFmt As String, _
    sKey As String, sType As String, sView As String, _
    sReason As String, sStatus As String)
    
    ' 1. 좌측 [로그 영역] 기록 (성공/실패 모두 기록)
    With wsReport
        .Cells(NextLogRow, 1).Value = sTitle
        .Cells(NextLogRow, 2).Value = sStatus ' PASS or FAIL
        .Cells(NextLogRow, 3).Value = sReason
        .Cells(NextLogRow, 4).Value = "'" & sKey
        .Cells(NextLogRow, 5).Value = sType
        .Cells(NextLogRow, 6).Value = sView
        
        ' 상태에 따른 색상 강조
        If sStatus = "FAIL" Then
            .Range(.Cells(NextLogRow, 1), .Cells(NextLogRow, 6)).Font.Color = RGB(200, 0, 0)
        Else
            .Range(.Cells(NextLogRow, 1), .Cells(NextLogRow, 6)).Font.Color = RGB(0, 100, 0)
        End If
        
        NextLogRow = NextLogRow + 1
    End With
    
    ' 2. 우측 [실험 영역] 기록 (실패한 경우에만 생성)
    If sStatus = "FAIL" Then
        With wsReport
            Dim sHeader As String
            sHeader = sTitle & vbCrLf & _
                      "[Key: " & sKey & "]" & vbCrLf & _
                      "[Mode: " & sType & "]" & vbCrLf & _
                      "[Why: " & sReason & "]"
            
            With .Cells(1, NextExpCol)
                .Value = sHeader
                .Interior.Color = RGB(255, 230, 230)
                .Font.Bold = True
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
            
            With .Cells(2, NextExpCol)
                .Value = vTarget
                .NumberFormat = sTargetFmt
                .Interior.Color = RGB(240, 255, 240)
            End With
            
            With .Cells(3, NextExpCol)
                .Value = vNoise
                .NumberFormat = sNoiseFmt
                .Interior.Color = RGB(245, 245, 245)
                .Font.Color = RGB(150, 150, 150)
            End With
            
            .Columns(NextExpCol).ColumnWidth = 20
            NextExpCol = NextExpCol + 1
        End With
    End If
End Sub

Private Sub SetupReportSheet()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Test_Fail_Report").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsReport = ThisWorkbook.Worksheets.Add
    wsReport.Name = "Test_Fail_Report"
    
    ' 구역 1: 로그 영역 (A~F)
    Dim logHeaders As Variant
    logHeaders = Array("Test Case", "Result", "Reason", "Keyword", "Mode", "View Text")
    
    With wsReport
        .Range("A1").Resize(1, UBound(logHeaders) + 1).Value = logHeaders
        .Range("A1").Resize(1, UBound(logHeaders) + 1).Font.Bold = True
        .Range("A1").Resize(1, UBound(logHeaders) + 1).Interior.Color = RGB(220, 220, 220)
        .Columns("G").ColumnWidth = 1: .Columns("H").ColumnWidth = 1
        .Columns("G").Interior.Color = RGB(100, 100, 100)
        .Range("I1").Value = "▶ EXPERIMENT ZONE (FAIL ONLY)"
        .Range("I1").Font.Bold = True: .Range("I1").Font.Color = vbWhite: .Range("I1").Interior.Color = RGB(50, 50, 150)
    End With
    
    NextLogRow = 2: NextExpCol = 9
End Sub

Private Sub SetupSpecSheet()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Valid_Spec_Sheet").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsSpec = ThisWorkbook.Worksheets.Add
    wsSpec.Name = "Valid_Spec_Sheet"
    wsSpec.Range("A1").Value = "TargetCol"
    Set loSpec = wsSpec.ListObjects.Add(xlSrcRange, wsSpec.Range("A1"), , xlYes)
End Sub

Private Function GetSafeNoise(vTarget As Variant) As Variant
    Select Case VarType(vTarget)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal: GetSafeNoise = -999999
        Case vbDate: GetSafeNoise = DateSerial(1900, 3, 1) '윈도 1900년 윤년 버그로 인한 날짜 설정 변경
        Case vbBoolean: GetSafeNoise = Not vTarget
        Case Else: GetSafeNoise = "###_NOISE_DATA_###"
    End Select
End Function


