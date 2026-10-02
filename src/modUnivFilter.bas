Attribute VB_Name = "modUnivFilter"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
' @desc 진입 UI(버튼/메뉴)와 폼 기동 담당
' @depends frmFilter01, modCache, modLogger, AutoCloseMsgAPI, Utilities
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
Public G_OpenForms As Collection        ' 폼 중복 실행 방지용 컬렉션

' ==============================================================================
' [Main] 유니버설 필터 진입점
' ==============================================================================
' @desc 호출원 셀을 식별해 필터 폼을 엶
Public Sub ShowUniversalFilter()
    Dim targetCell As Range, btn As Shape
    On Error Resume Next
    If TypeName(Application.Caller) = "String" Then
        Set btn = ActiveSheet.Shapes(Application.Caller)
        If Not btn Is Nothing Then Set targetCell = btn.TopLeftCell
    Else
        Set targetCell = ActiveCell
    End If
    On Error GoTo 0
    OpenFilterForm targetCell, True
End Sub

' @desc 대상 셀 받아 공용 폼을 표시함, 대상 없으면 알림
Public Sub OpenFilterForm(ByVal targetCell As Range, Optional ByVal atCursor As Boolean = False)
    Dim frm As frmFilter01
    Set frm = GetFilterForm(targetCell)
    If frm Is Nothing Then
        AutoCloseMsgAPI.AutoCloseMsgBox_Popup "표(Table) 또는 자동 필터 범위 내에서 실행해 주세요.", 1, "UnivFilter 알림"
        Exit Sub
    End If
    If atCursor Then
        MoveFormToCursor_API frm
    Else
        MoveFormToCell_API frm, targetCell
    End If
    frm.Show vbModeless
    On Error Resume Next
    frm.Controls("txtFilterValues").SetFocus
    ' 실행 시점에는 파이프라인이 지문을 다시 검증한다.
    modCache.GetColumnAsset frm.InjectBodyRange
    On Error GoTo 0
End Sub

' @desc 대상 셀 받아 주입된 공용 폼 반환, 필터 대상 없으면 Nothing
Public Function GetFilterForm(ByVal targetCell As Range) As frmFilter01
    Dim body As Range, area As Range, field As Long, header As String
    If Not ResolveFilterTarget(targetCell, body, area, field, header) Then Exit Function

    ' 통합문서/시트/표의 헤더 위치/필드를 구분한다. 본문 행수 변경에도 키는 유지.
    Dim formKey As String
    formKey = area.Rows(1).Address(External:=True) & "!" & CStr(field)
    If G_OpenForms Is Nothing Then Set G_OpenForms = New Collection
    Dim frm As frmFilter01
    On Error Resume Next
    Set frm = G_OpenForms(formKey)
    On Error GoTo 0
    If frm Is Nothing Then
        Set frm = New frmFilter01
        frm.MyFormKey = formKey
        G_OpenForms.Add frm, formKey
    End If
    ' 재사용 시에도 현재 범위를 주입해 표 크기 변경을 반영한다.
    Set frm.InjectBodyRange = body
    Set frm.InjectFilterArea = area
    frm.InjectRelativeCol = field
    frm.InjectHeaderName = header
    Set GetFilterForm = frm
End Function

' @desc 대상 셀 받아 본문/필터영역/필드/헤더 반환, 대상 없으면 False
Public Function ResolveFilterTarget(ByVal targetCell As Range, ByRef rngBody As Range, _
                                    ByRef rngArea As Range, ByRef lField As Long, _
                                    ByRef sHeader As String) As Boolean
    Set rngBody = Nothing
    Set rngArea = Nothing
    lField = 0
    sHeader = ""
    If targetCell Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = targetCell.Parent
    Dim lo As ListObject: Set lo = targetCell.ListObject
    If Not lo Is Nothing Then
        If lo.DataBodyRange Is Nothing Then Exit Function
        Set rngArea = lo.Range
        lField = targetCell.Column - rngArea.Column + 1
        Set rngBody = lo.DataBodyRange.Columns(lField)
    Else
        If Not ws.AutoFilterMode Then Exit Function
        Set rngArea = ws.AutoFilter.Range
        If Intersect(targetCell, rngArea) Is Nothing Then
            Set rngArea = Nothing
            Exit Function
        End If
        If rngArea.Rows.Count < 2 Then Exit Function
        lField = targetCell.Column - rngArea.Column + 1
        Set rngBody = rngArea.Columns(lField).Offset(1, 0).Resize(rngArea.Rows.Count - 1, 1)
    End If
    sHeader = CStr(rngArea.Cells(1, lField).Value)
    ResolveFilterTarget = True
End Function

' ------------------------------------------------------------------------------
' [Utility] API 기반 폼 위치 이동 (표준 화면 이탈 방지 로직 적용)
' ------------------------------------------------------------------------------
Public Sub MoveFormToCursor_API(frm As Object)
    Dim curPos As POINTAPI
    Dim dc As LongPtr
    Dim pixelsPerPointX As Double, pixelsPerPointY As Double
    
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
' @desc 선택 범위의 필터 버튼을 생성함
Public Sub ToggleSearchButtonsInSelection()
    If TypeName(Selection) <> "Range" Then Exit Sub
    On Error GoTo Fail
    CreateSearchButtons Selection
    Exit Sub
Fail:
    MsgBox "버튼 생성 중 오류 발생: " & Err.Description
End Sub

' @desc 범위의 필터 버튼을 생성함, 실패 시 설정 복원 후 오류 전달
Public Sub CreateSearchButtons(ByVal target As Range)
    Dim area As Range
    Dim ws As Worksheet: Set ws = target.Parent
    Dim colRng As Range, headerCell As Range
    Dim btn As Shape
    Dim btnName As String, btnSize As Double
    Dim processedCols As Object: Set processedCols = CreateObject("Scripting.Dictionary")
    
    
    On Error GoTo ErrHandler
    Call OptimizedMode(True)
    
    For Each area In target.Areas
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
                On Error GoTo ErrHandler
                
                ' 2. 버튼 처리 (삭제 후 생성)
                If Not headerCell Is Nothing Then
                    btnName = "SmartFilterBtn_" & headerCell.Address(False, False)
                    
                    ' 일단 존재 여부 묻지 않고 삭제 (없으면 에러 무시하고 넘어감)
                    On Error Resume Next
                    ws.Shapes(btnName).Delete
                    On Error GoTo ErrHandler
                    
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
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    Utilities.OptimizedMode False
    Err.Raise errNum, "modUnivFilter.CreateSearchButtons", errDesc
    
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
