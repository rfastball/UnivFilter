Attribute VB_Name = "Core_Enums"
' @desc UnivCore 공통 Enum/상수 선언
Option Explicit

Public Const CORE_USE_UNSAFE_MEM As Long = 0

Public Enum CoreSetRetMode
    srIntersect = 1
    srDiffMe = 2
    srDiffOther = 3
    srAllSplit = 4
End Enum
