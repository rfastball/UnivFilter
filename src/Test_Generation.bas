Attribute VB_Name = "Test_Generation"
' @desc 세대 계약 stale 자산 fail-fast 단위테스트 담당
' @depends modCache, modProcessManager, clsColumnAsset, LogShim
Option Explicit

Private Const TEST_SHEET As String = "Gen_Test_Sheet"

Private Function SetupSheet() As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(TEST_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = TEST_SHEET
    ws.Range("A1").Value = "GenCol"
    ws.Range("A2").Value = "alpha"
    ws.Range("A3").Value = "beta"
    ws.ListObjects.Add xlSrcRange, ws.Range("A1:A3"), , xlYes
    Set SetupSheet = ws
End Function

Private Sub Teardown(ByVal ws As Worksheet)
    Application.DisplayAlerts = False
    On Error Resume Next
    ws.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

' @desc 편집 무효화 후 보유 자산 사용이 9998로 실패하는지 검증함
Public Sub Test_Gen_StaleAsset_FailFast()
    Dim ws As Worksheet
    Set ws = SetupSheet()

    modCache.InvalidateCache
    Dim asset As clsColumnAsset
    Set asset = modCache.GetColumnAsset(ws.Range("A2:A3"))
    LogShim.AssertEq "gen_fresh_not_stale", False, modCache.IsAssetStale(asset)

    ' 편집 이벤트 시뮬레이션: AppEvents_SheetChange가 호출하는 것과 동일 경로
    modCache.InvalidateCacheBySheet ws
    LogShim.AssertEq "gen_stale_after_edit", True, modCache.IsAssetStale(asset)

    Dim dictRes As Object
    Set dictRes = modProcessManager.ExecuteFilterProcess( _
                      "alpha", asset, ws.ListObjects(1).Range, 1, False, False, False)
    LogShim.AssertEq "gen_stale_failfast_success", False, dictRes("IsSuccess")
    LogShim.AssertEq "gen_stale_failfast_errno", 9998, dictRes("ErrorNumber")

    Teardown ws
End Sub

' @desc 보호 시트 필터 에러가 dict 에러 채널로 전달되는지 검증함 (Stage 8 재던지기 경로)
Public Sub Test_Gen_ErrorChannel_ProtectedSheet()
    Dim ws As Worksheet
    Set ws = SetupSheet()

    modCache.InvalidateCache
    Dim asset As clsColumnAsset
    Set asset = modCache.GetColumnAsset(ws.Range("A2:A3"))

    ' 시트 보호로 AutoFilter 적용 에러를 강제 -> ApplyFilterResult 재던지기 ->
    ' ExecuteFilterProcess 에러 채널(IsSuccess=False + ErrorNumber)로 전달돼야 한다
    ws.Protect "test"
    Dim dictRes As Object
    Set dictRes = modProcessManager.ExecuteFilterProcess( _
                      "alpha", asset, ws.ListObjects(1).Range, 1, False, False, False)
    ws.Unprotect "test"

    LogShim.AssertEq "errchan_success_false", False, dictRes("IsSuccess")
    LogShim.AssertEq "errchan_has_errnumber", True, dictRes.exists("ErrorNumber")

    Teardown ws
End Sub

' @desc 무효화 후 재취득 자산의 정상 필터 경로가 green인지 검증함
Public Sub Test_Gen_ReacquiredAsset_Green()
    Dim ws As Worksheet
    Set ws = SetupSheet()

    ' 편집 무효화 후 재취득: 자산은 현재 세대를 스탬프받아 정상 동작해야 함
    modCache.InvalidateCacheBySheet ws
    Dim asset As clsColumnAsset
    Set asset = modCache.GetColumnAsset(ws.Range("A2:A3"))
    LogShim.AssertEq "gen_reacquire_not_stale", False, modCache.IsAssetStale(asset)

    Dim dictRes As Object
    Set dictRes = modProcessManager.ExecuteFilterProcess( _
                      "alpha", asset, ws.ListObjects(1).Range, 1, False, False, False)
    LogShim.AssertEq "gen_reacquire_success", True, dictRes("IsSuccess")
    LogShim.AssertEq "gen_reacquire_target_visible", False, ws.Rows(2).Hidden
    LogShim.AssertEq "gen_reacquire_noise_hidden", True, ws.Rows(3).Hidden

    Teardown ws
End Sub
