Attribute VB_Name = "modUnivFilter"
Option Explicit

' ==============================================================================
' [API] Windows API 및 상수 선언 (좌표/DPI 계산용)
' ==============================================================================
Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

' ==============================================================================
' [Global] 전역 변수 및 상태 관리
' ==============================================================================
Public G_LastExcludeMode As Boolean     ' 마지막 제외 모드 상태
Public G_LastStackMode As Boolean       ' 마지막 누적 모드 상태
Public G_LastMatchMode As Boolean       ' 마지막 일치 모드 상태
Public G_OpenForms As Collection        ' 폼 중복 실행 방지용 컬렉션

' ==============================================================================
' [Main] 유니버설 필터 진입점
' ==============================================================================
Sub ShowUniversalFilter()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim targetCell As Range, btn As Shape
    Dim btnCol As Long
    
    ' 1. [Target] 호출원(버튼/셀) 식별
    On Error Resume Next
    If TypeName(Application.Caller) = "String" Then
        Set btn = ws.Shapes(Application.Caller)
        If Not btn Is Nothing Then Set targetCell = btn.TopLeftCell
    Else
        If Not ActiveCell Is Nothing Then Set targetCell = ActiveCell
    End If
    On Error GoTo 0
    
    If targetCell Is Nothing Then MsgBox "대상 셀 식별 불가", vbExclamation: Exit Sub
    btnCol = targetCell.Column
    
    ' 2. [Singleton] 폼 중복 실행 방지 (이미 열려있으면 활성화)
    Dim formKey As String: formKey = ws.Name & "!" & btnCol
    Dim openFrm As Object
    
    If G_OpenForms Is Nothing Then Set G_OpenForms = New Collection
    
    On Error Resume Next
    Set openFrm = G_OpenForms(formKey)
    On Error GoTo 0
    
    If Not openFrm Is Nothing Then
        openFrm.Show vbModeless
        On Error Resume Next
        openFrm.Controls("txtFilterValues").SetFocus
        On Error GoTo 0
        Exit Sub
    End If

    ' 3. [Data Context] 표(ListObject) 또는 자동필터 범위 탐색
    Dim lo As ListObject
    Dim targetBodyRng As Range, filterAreaRng As Range
    Dim targetHeader As String
    Dim relativeCol As Long
    Dim found As Boolean: found = False
    
    ' A) 표(Table) 확인
    For Each lo In ws.ListObjects
        If Not Intersect(ws.Columns(btnCol), lo.Range) Is Nothing Then
            relativeCol = btnCol - lo.Range.Column + 1
            If lo.ListRows.Count = 0 Then MsgBox "데이터 없음", vbExclamation: Exit Sub
            
            Set targetBodyRng = lo.ListColumns(relativeCol).DataBodyRange
            Set filterAreaRng = lo.Range
            targetHeader = lo.HeaderRowRange.Cells(1, relativeCol).Value
            found = True: Exit For
        End If
    Next lo
    
    ' B) 자동필터(AutoFilter) 확인
    If Not found And ws.AutoFilterMode Then
        If Not Intersect(ws.Columns(btnCol), ws.AutoFilter.Range) Is Nothing Then
            Dim afRng As Range: Set afRng = ws.AutoFilter.Range
            relativeCol = btnCol - afRng.Column + 1
            
            If afRng.Rows.Count > 1 Then
                Set targetBodyRng = afRng.Offset(1, 0).Resize(afRng.Rows.Count - 1, 1).Columns(relativeCol)
                Set filterAreaRng = afRng
                targetHeader = afRng.Cells(1, relativeCol).Value
                found = True
            Else
                MsgBox "데이터 없음", vbExclamation: Exit Sub
            End If
        End If
    End If
    
    If Not found Then MsgBox "필터 대상(표/자동필터)을 찾을 수 없습니다.", vbCritical: Exit Sub
    
    ' 4. [Init & Show] 폼 생성 및 표시
    Dim frmNew As frmFilter01
    Set frmNew = New frmFilter01
    
    With frmNew
        .MyFormKey = formKey                    ' 닫을 때 컬렉션 제거용 Key
        Set .InjectBodyRange = targetBodyRng    ' 데이터 범위 주입
        Set .InjectFilterArea = filterAreaRng   ' 필터 영역 주입
        .InjectHeaderName = targetHeader        ' 헤더명 주입
        .InjectRelativeCol = relativeCol        ' 상대 열 번호 주입
        
        .StartInitialization                  ' 모든 준비가 끝난 후 초기화 실행! (중요)
        ' [Stage 7] StartInitialization이 정책(modTypePolicy) 기반으로 잠금을 이미
        ' 수행하므로, 직후의 중복 EnforceHybridMode 호출은 제거됨 (동작 동일)
        G_OpenForms.Add frmNew, formKey         ' 관리 대장 등록
        
        MoveFormToCursor_API frmNew             ' 마우스 위치로 이동
        .Show vbModeless
    End With
End Sub

' ------------------------------------------------------------------------------
' [Utility] API 기반 폼 위치 이동 (표준 화면 이탈 방지 로직 적용)
' ------------------------------------------------------------------------------
Public Sub MoveFormToCursor_API(frm As Object)
    Dim curPos As POINTAPI
    Dim dc As LongPtr
    Dim pixelsPerPointX As Double, pixelsPerPointY As Double
    Dim ZoomRatio As Double
    
    ' 1. 현재 마우스 커서 위치(Pixel) 가져오기
    GetCursorPos curPos
    
    ' 2. 화면 DPI 및 엑셀 Zoom 배율 계산
    ' (DPI와 엑셀의 Zoom 상태에 따라 좌표가 어긋나는 것을 방지하기 위함)
    dc = GetDC(0)
    pixelsPerPointX = GetDeviceCaps(dc, LOGPIXELSX) / 72
    pixelsPerPointY = GetDeviceCaps(dc, LOGPIXELSY) / 72
    ReleaseDC 0, dc
    
    ' 3. 폼 초기화 (수동 위치 모드)
    frm.StartUpPosition = 0 ' Manual
    
    ' --------------------------------------------------------------------------
    ' [표준 보정 알고리즘] Boundary Clamping & Flipping
    ' --------------------------------------------------------------------------
    Dim targetLeft As Double, targetTop As Double
    Dim appLeft As Double, appTop As Double, appRight As Double, appBottom As Double
    Dim margin As Double
    
    margin = 5 ' 창 가장자리와의 최소 여백 (Point 단위)
    
    ' A. 엑셀 애플리케이션의 표시 영역(ViewPort) 좌표 계산
    '    (폼이 엑셀 창 밖으로 나가지 않도록 가두는 범위)
    appLeft = Application.Left
    appTop = Application.Top
    appRight = Application.Left + Application.Width
    appBottom = Application.Top + Application.Height
    
    ' B. 목표 위치 계산 (Pixel -> Point 변환)
    '    기본 전략: 마우스 커서의 오른쪽 아래에 표시
    targetLeft = (curPos.x / pixelsPerPointX)
    targetTop = (curPos.Y / pixelsPerPointY) + 12 ' 커서보다 살짝 아래
    
    ' C. 가로(Horizontal) 보정: 오른쪽이 잘리면 왼쪽으로 이동
    If (targetLeft + frm.Width) > (appRight - margin) Then
        targetLeft = (appRight - margin) - frm.Width
    End If
    '    그래도 왼쪽이 잘리면(창이 너무 좁음), 왼쪽 벽에 붙임
    If targetLeft < (appLeft + margin) Then
        targetLeft = appLeft + margin
    End If
    
    ' D. 세로(Vertical) 보정: 아래가 잘리면 위로 뒤집기 (Flip)
    If (targetTop + frm.Height) > (appBottom - margin) Then
        ' 커서 높이만큼 위로 올림 (커서 위쪽에 배치)
        ' 24는 일반적인 커서 크기 + 여백 값
        targetTop = (curPos.Y / pixelsPerPointY) - frm.Height - 12
    End If
    
    ' E. 상단 보정: 위로 뒤집었는데 제목 표시줄을 뚫고 나가면, 상단 벽에 붙임
    If targetTop < (appTop + margin) Then
        targetTop = appTop + margin
    End If
    
    ' 4. 최종 좌표 적용
    frm.Left = targetLeft
    frm.Top = targetTop
End Sub

' ------------------------------------------------------------------------------
' [Utility] 셀 위치 추적 이동 (정밀 보정 버전)
' 목적: Zoom 배율 및 틀 고정 상태에서도 정확하게 셀 아래에 폼을 부착
' ------------------------------------------------------------------------------
Public Sub MoveFormToCell_API(frm As Object, targetCell As Range)
    Dim dc As LongPtr
    Dim pxPerPointX As Double, pxPerPointY As Double
    Dim cellLeftPx As Long, cellBottomPx As Long
    Dim currentPane As Pane
    
    ' 1. 시스템 DPI 배율 계산 (폼의 좌표계인 'Point' 단위 변환용)
    dc = GetDC(0)
    pxPerPointX = GetDeviceCaps(dc, LOGPIXELSX) / 72
    pxPerPointY = GetDeviceCaps(dc, LOGPIXELSY) / 72
    ReleaseDC 0, dc
    
    ' 2. [핵심] 현재 활성화된 '패널(Pane)' 기준 좌표 계산
    '    (ActiveWindow 대신 ActivePane을 써야 틀 고정/스크롤 상태가 정확히 반영됨)
    Set currentPane = ActiveWindow.ActivePane
    
    ' PointsToScreenPixels: 엑셀 내부 좌표(Point) -> 모니터 절대 좌표(Pixel) 변환
    ' 이 함수는 이미 엑셀의 Zoom 배율이 반영된 픽셀 값을 반환합니다.
    cellLeftPx = currentPane.PointsToScreenPixelsX(targetCell.Left)
    cellBottomPx = currentPane.PointsToScreenPixelsY(targetCell.Top + targetCell.Height)
    
    ' 3. 폼 위치 설정 (Pixel -> VBA Form Point 역변환)
    '    엑셀이 준 Pixel 좌표를 시스템 DPI 비율로 나누어 폼이 이해하는 Point로 변환
    frm.StartUpPosition = 0 ' Manual
    
    Dim targetLeft As Double, targetTop As Double
    
    ' 셀 바로 아래 배치 (+1 pixel 정도 미세 조정)
    targetLeft = (cellLeftPx / pxPerPointX)
    targetTop = (cellBottomPx / pxPerPointY) + 1
    
    ' --------------------------------------------------------------------------
    ' [화면 이탈 방지 보정]
    ' --------------------------------------------------------------------------
    Dim appRight As Double, appBottom As Double
    ' Application.Left/Width는 폼 좌표계와 동일하므로 그대로 사용 가능
    appRight = Application.Left + Application.Width
    appBottom = Application.Top + Application.Height
    
    ' A. 오른쪽 벽에 막히면 -> 셀 왼쪽 끝에 맞춰서 폼을 왼쪽으로 당김
    If (targetLeft + frm.Width) > appRight Then
        ' 셀의 오른쪽 끝이 아니라, 폼 전체가 화면에 들어오도록 조정
        targetLeft = appRight - frm.Width - 10
    End If
    
    ' B. 아래쪽 벽에 막히면 -> 셀 위쪽으로 뒤집기 (Flip)
    If (targetTop + frm.Height) > appBottom Then
        ' 셀 위쪽 좌표 다시 계산
        Dim cellTopPx As Long
        cellTopPx = currentPane.PointsToScreenPixelsY(targetCell.Top)
        targetTop = (cellTopPx / pxPerPointY) - frm.Height - 1
    End If
    
    ' 최종 적용
    frm.Left = targetLeft
    frm.Top = targetTop
End Sub

' ==============================================================================
' [Context Menu] 우클릭 메뉴 관리
' ==============================================================================
Sub SetupContextMenu()
    RemoveContextMenu
    Dim cBar As CommandBar, menuName As Variant
    Dim btnRun As CommandBarButton, subMenu As CommandBarPopup
    
    For Each menuName In Array("Cell", "List Range Popup")
        On Error Resume Next
        Set cBar = Application.CommandBars(menuName)
        On Error GoTo 0
        
        If Not cBar Is Nothing Then
            ' 실행 버튼
            Set btnRun = cBar.Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
            With btnRun
                .Caption = "UnivFilter 실행"
                .FaceId = 25
                .OnAction = "CallFilterFromMenu"
                .tag = "UniversalFilterContext"
                .BeginGroup = True
            End With
            
            ' 관리 서브메뉴
            Set subMenu = cBar.Controls.Add(Type:=msoControlPopup, Before:=2, Temporary:=True)
            With subMenu
                .Caption = "UnivFilter 버튼 관리"
                .tag = "UniversalFilterContext"
                .Controls.Add(msoControlButton, Temporary:=True).Caption = "버튼 생성 (해당 열)": .Controls(1).OnAction = "ToggleSearchButtonsInSelection": .Controls(1).FaceId = 486
                .Controls.Add(msoControlButton, Temporary:=True).Caption = "버튼 삭제 (해당 열)": .Controls(2).OnAction = "DeleteButtonsInSelection": .Controls(2).FaceId = 358
                .Controls.Add(msoControlButton, Temporary:=True).Caption = "버튼 초기화 (전체)": .Controls(3).OnAction = "DeleteAllFilterButtons": .Controls(3).FaceId = 463: .Controls(3).BeginGroup = True
            End With
        End If
    Next menuName
End Sub

Sub RemoveContextMenu()
    Dim ctrl As CommandBarControl, menuName As Variant
    On Error Resume Next
    For Each menuName In Array("Cell", "List Range Popup")
        For Each ctrl In Application.CommandBars(menuName).Controls
            If ctrl.tag = "UniversalFilterContext" Then ctrl.Delete
        Next ctrl
    Next menuName
    On Error GoTo 0
End Sub

' ==============================================================================
' [Buttons] 시트 내 버튼 생성/삭제 로직
' ==============================================================================
Sub ToggleSearchButtonsInSelection()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim area As Range
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim colRng As Range, headerCell As Range
    Dim btn As Shape, existBtn As Shape
    Dim btnName As String, btnSize As Double
    Dim processedCols As Object: Set processedCols = CreateObject("Scripting.Dictionary")
    
    
    On Error GoTo ErrHandler
    Call OptimizedMode(True)
    
    For Each area In Selection.Areas
        For Each colRng In area.Columns
            If Not processedCols.exists(colRng.Column) Then
                processedCols.Add colRng.Column, True
                
                Set headerCell = Nothing
                
                ' 1. 헤더 셀 찾기
                On Error Resume Next
                If Not colRng.Cells(1).ListObject Is Nothing Then
                    Set headerCell = Intersect(colRng.Cells(1).ListObject.HeaderRowRange, ws.Columns(colRng.Column))
                ElseIf ws.AutoFilterMode Then
                    If Not Intersect(ws.Columns(colRng.Column), ws.AutoFilter.Range) Is Nothing Then
                        Set headerCell = Intersect(ws.AutoFilter.Range.Rows(1), ws.Columns(colRng.Column))
                    End If
                End If
                On Error GoTo 0
                
                ' 2. 버튼 처리 (삭제 후 생성)
                If Not headerCell Is Nothing Then
                    btnName = "SmartFilterBtn_" & headerCell.Address(False, False)
                    
                    ' 일단 존재 여부 묻지 않고 삭제 (없으면 에러 무시하고 넘어감)
                    On Error Resume Next
                    ws.Shapes(btnName).Delete
                    On Error GoTo 0
                    
                    ' 무조건 새로 생성
                    btnSize = Application.Min(Application.Max(headerCell.Height * 0.45, 10), 16)
                    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                              headerCell.Left + 0.5, _
                              headerCell.Top + headerCell.Height - btnSize - 0.5, _
                              btnSize, btnSize)
                    
                    With btn
                        .Placement = xlMove
                        .Adjustments.item(1) = 0.15
                        .Fill.ForeColor.RGB = RGB(240, 240, 240)
                        .Line.ForeColor.RGB = RGB(170, 170, 170)
                        .Line.Weight = 0.75
                        .Shadow.Visible = msoFalse
                        .Name = btnName
                        .OnAction = "ShowUniversalFilter"
                        .ControlFormat.PrintObject = False
                        
                        With .TextFrame2
                            .TextRange.text = "L"
                            .TextRange.Font.Name = "Webdings"
                            .TextRange.Font.Fill.ForeColor.RGB = RGB(80, 80, 80)
                            .VerticalAnchor = msoAnchorMiddle
                            .HorizontalAnchor = msoAnchorCenter
                            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                            .TextRange.Font.Size = btnSize * 0.7
                        End With
                    End With
                End If
            End If
        Next colRng
    Next area
    
    Set processedCols = Nothing
    Call OptimizedMode(False)
    Exit Sub
    
ErrHandler:
    ' 에러 발생 시에도 설정 복구
    Call OptimizedMode(False)
    MsgBox "버튼 생성 중 오류 발생: " & Err.Description
    
End Sub

Sub DeleteButtonsInSelection()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim shp As Shape
    
    On Error GoTo ErrHandler
    Call OptimizedMode(True)
    
    For Each shp In ActiveSheet.Shapes
        If shp.Name Like "SmartFilterBtn_*" Then
            If Not Intersect(ActiveSheet.Columns(shp.TopLeftCell.Column), Selection.EntireColumn) Is Nothing Then shp.Delete
        End If
    Next shp
    Call OptimizedMode(False)
    Exit Sub
    
ErrHandler:
    ' 에러 발생 시에도 설정 복구
    Call OptimizedMode(False)
    MsgBox "버튼 삭제 중 오류 발생: " & Err.Description
End Sub

Sub DeleteAllFilterButtons()
    If MsgBox("모든 필터 버튼을 삭제하시겠습니까?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Dim shp As Shape, delCount As Long
    
    ' [Fix v2.7] 최적화 모드 진입/복구 패턴 통일
    On Error GoTo ErrHandler
    Call OptimizedMode(True)

    For Each shp In ActiveSheet.Shapes
        If shp.Name Like "SmartFilterBtn_*" Then shp.Delete: delCount = delCount + 1
    Next shp
    Call OptimizedMode(False)
    MsgBox "초기화 완료 (" & delCount & "개 삭제)", vbInformation
    Exit Sub
    
ErrHandler:
    ' 에러 발생 시에도 설정 복구
    Call OptimizedMode(False)
    MsgBox "버튼 초기화 중 오류 발생: " & Err.Description

End Sub

Sub CallFilterFromMenu()
    Call ShowUniversalFilter
End Sub
