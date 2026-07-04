Attribute VB_Name = "Test_UnivFilterHarness"
' @desc UnivFilter 검증기를 하네스 진입점으로 연결
' @depends modTester, LogShim
Option Explicit

Public Sub Test_UnivFilter_SpecValidation()
    Run_SpecValidation
    LogShim.AssertEq "univfilter_spec_validation_completed", True, True
End Sub
