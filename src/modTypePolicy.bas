Attribute VB_Name = "modTypePolicy"
' @desc 컬럼 타입별 UI/엔진 모드 정책 제공 담당
Option Explicit

' ==============================================================================
' [L3 정책 단일 소스] (Stage 7, D5)
' 기존 4중 중복(frmFilter01.StartInitialization / modShortcuts / modMatchingEngine /
' modUnivFilter)의 타입 -> 모드 분기를 이 진리표 하나로 통일한다.
'
' 주의: ROUNDED/PERCENT_ROUNDED의 비대칭(UI는 잠그되 엔진은 문자열 강제 안 함,
' 퀵 모드도 즉시 실행)은 검증된 현행 표면 동작이다 - 정합성 목적의 "개선" 금지.
' ==============================================================================

Public Type ModePolicy
    ForceTextStrategy As Boolean       ' 엔진: String 전략 강제 여부
    LockTag As String                  ' 폼 UI 잠금 종류: "" / "HYBRID" / "ROUNDING" / "DATE"
    RequiresFormInQuickMode As Boolean ' 퀵 모드(Ctrl+Shift+Q)에서 폼 표시로 우회 여부
End Type

' @desc 컬럼 타입 문자열 받아 ModePolicy 반환, 미확인 타입은 잠금 없는 기본값
Public Function GetModePolicy(ByVal sColumnType As String) As ModePolicy
    Dim p As ModePolicy

    Select Case UCase(sColumnType)
        Case "HYBRID", "ERROR"
            p.ForceTextStrategy = True
            p.LockTag = "HYBRID"
            p.RequiresFormInQuickMode = True

        Case "DATE"
            p.ForceTextStrategy = True
            p.LockTag = "DATE"
            p.RequiresFormInQuickMode = True

        Case "ROUNDED", "PERCENT_ROUNDED"
            ' 비대칭 주의: UI 잠금만, 엔진 강제/퀵 우회 없음 (현행 동작)
            p.ForceTextStrategy = False
            p.LockTag = "ROUNDING"
            p.RequiresFormInQuickMode = False

        Case Else ' NUMBER, PERCENT, STRING 및 미확인
            p.ForceTextStrategy = False
            p.LockTag = ""
            p.RequiresFormInQuickMode = False
    End Select

    GetModePolicy = p
End Function
