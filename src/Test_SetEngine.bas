Attribute VB_Name = "Test_SetEngine"
' @desc dict 셋 엔진 계약 특성화 단위테스트 담당
' @depends clsSet, LogShim
Option Explicit

' 리팩토링(엔진 벤더링 교체) 전후 동작 불변을 보증하는 특성화 테스트.
' 여기 어서션은 "현행 엔진의 계약"을 기록한 것 - 개선 목적의 수정 금지.

' @desc Compare srAllSplit의 3분할 결과를 검증함
Public Sub Test_Set_Compare_AllSplit()
    Dim setA As clsSet, setB As clsSet
    Dim vRes As Variant

    Set setA = New clsSet
    Set setB = New clsSet
    setA.LoadFromArray Array(CDbl(1), CDbl(2), CDbl(3))
    setB.LoadFromArray Array(CDbl(2), CDbl(3), CDbl(4))

    vRes = setA.Compare(setB, srAllSplit)

    LogShim.AssertEq "set_allsplit_common_count", 2, vRes(0).Count
    LogShim.AssertEq "set_allsplit_diffme_count", 1, vRes(1).Count
    LogShim.AssertEq "set_allsplit_diffother_count", 1, vRes(2).Count
    LogShim.AssertEq "set_allsplit_common_has_2", True, vRes(0).Contains(CDbl(2))
    LogShim.AssertEq "set_allsplit_diffme_has_1", True, vRes(1).Contains(CDbl(1))
    LogShim.AssertEq "set_allsplit_diffother_has_4", True, vRes(2).Contains(CDbl(4))
End Sub

' @desc Union/Difference 래퍼의 결과 집합을 검증함
Public Sub Test_Set_Union_Difference()
    Dim setA As clsSet, setB As clsSet
    Dim setU As clsSet, setD As clsSet

    Set setA = New clsSet
    Set setB = New clsSet
    setA.LoadFromArray Array("a", "b")
    setB.LoadFromArray Array("b", "c")

    Set setU = setA.Union(setB)
    LogShim.AssertEq "set_union_count", 3, setU.Count
    LogShim.AssertEq "set_union_has_c", True, setU.Contains("c")

    Set setD = setA.Difference(setB)
    LogShim.AssertEq "set_diff_count", 1, setD.Count
    LogShim.AssertEq "set_diff_has_a", True, setD.Contains("a")
End Sub

' @desc LoadFromArray가 Empty/Error 항목을 건너뛰는지 검증함
Public Sub Test_Set_LoadFromArray_SkipsInvalid()
    Dim setA As clsSet

    Set setA = New clsSet
    setA.LoadFromArray Array("x", Empty, CVErr(2042), "y")
    LogShim.AssertEq "set_load_skips_empty_error", 2, setA.Count

    ' 배열이 아닌 단일 값은 Add 경로로 수용
    Set setA = New clsSet
    setA.LoadFromArray "solo"
    LogShim.AssertEq "set_load_single_value", 1, setA.Count

    ' Null/Error 단일 값은 가드로 무시
    setA.Add Null
    setA.Add CVErr(2042)
    LogShim.AssertEq "set_add_null_error_guarded", 1, setA.Count
End Sub

' @desc 문자열 키의 대소문자 무시 계약을 검증함
Public Sub Test_Set_CaseInsensitive_Keys()
    Dim setA As clsSet

    Set setA = New clsSet
    setA.Add "Apple"
    setA.Add "APPLE"

    LogShim.AssertEq "set_case_dedup", 1, setA.Count
    LogShim.AssertEq "set_case_contains_lower", True, setA.Contains("apple")
End Sub

' @desc 빈 집합과의 연산이 안전한지 검증함
Public Sub Test_Set_EmptySet_Operations()
    Dim setA As clsSet, setE As clsSet
    Dim vRes As Variant

    Set setA = New clsSet
    Set setE = New clsSet
    setA.LoadFromArray Array("a", "b")

    vRes = setA.Compare(setE, srAllSplit)
    LogShim.AssertEq "set_vs_empty_common", 0, vRes(0).Count
    LogShim.AssertEq "set_vs_empty_diffme", 2, vRes(1).Count
    LogShim.AssertEq "set_vs_empty_diffother", 0, vRes(2).Count

    vRes = setE.Compare(setA, srAllSplit)
    LogShim.LogD "empty_vs_set_diag", "isarr=" & IsArray(vRes) & " lb=" & LBound(vRes) & " ub=" & UBound(vRes)
    LogShim.AssertEq "empty_vs_set_common", 0, vRes(0).Count
    LogShim.LogD "empty_vs_set_diag", "t1=" & TypeName(vRes(1))
    LogShim.LogD "empty_vs_set_diag", "t2=" & TypeName(vRes(2))
    LogShim.AssertEq "empty_vs_set_diffother", 2, vRes(2).Count

    LogShim.AssertEq "empty_toarray_ubound", -1, UBound(setE.ToArray)
End Sub

' @desc Variant 키 의미론(숫자 1과 문자 "1" 구분)을 검증함
Public Sub Test_Set_NumericStringKey_Distinct()
    Dim setA As clsSet

    Set setA = New clsSet
    setA.Add CDbl(1)
    LogShim.AssertEq "set_numkey_not_strkey", False, setA.Contains("1")

    setA.Add "1"
    LogShim.AssertEq "set_num_str_both_kept", 2, setA.Count
End Sub
