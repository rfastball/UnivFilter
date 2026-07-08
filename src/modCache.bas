Attribute VB_Name = "modCache"
' @desc 열 자산 캐시와 시트 세대 관리 담당
' @depends modColumnScan, clsColAsset, modLogger
Option Explicit

' 캐시 항목: 키 = rng.Address(External:=True), 값 = clsColAsset (세대 스탬프 포함)
Private pColCache As Object

' 시트별 세대 카운터: 무효화(편집/재계산)마다 증가.
' GetColumnAsset이 취득 시 검증해 불일치면 투명하게 재스캔한다.
' (구 계약의 fail-fast 9998을 대체 - "낡은 스냅샷으로 필터 금지" 불변식은 동일)
Private pGenBySheet As Object

Private Function EnsureGenStore() As Object
    If pGenBySheet Is Nothing Then
        Set pGenBySheet = CreateObject("Scripting.Dictionary")
        pGenBySheet.CompareMode = vbTextCompare
    End If
    Set EnsureGenStore = pGenBySheet
End Function

' 캐시 키('[통합문서]시트'!$A$2:$A$9 형태)에서 시트 식별 접두부만 추출
Private Function PatternFromKey(ByVal sKey As String) As String
    Dim p As Long
    p = InStr(1, sKey, "!")
    If p > 0 Then PatternFromKey = Left$(sKey, p - 1) Else PatternFromKey = sKey
End Function

' @desc 시트 식별 접두부 받아 현재 세대 반환, 기록 없으면 0
Public Function GetSheetGeneration(ByVal sPattern As String) As Long
    Dim d As Object
    Set d = EnsureGenStore()
    If d.exists(sPattern) Then GetSheetGeneration = d(sPattern)
End Function

' @desc 열 범위 받아 자산 반환, 캐시 미스나 세대 불일치 시 재스캔 후 반환
Public Function GetColumnAsset(ByVal rngCol As Range) As clsColAsset
    If pColCache Is Nothing Then
        Set pColCache = CreateObject("Scripting.Dictionary")
    End If

    Dim sKey As String
    sKey = rngCol.Address(External:=True)

    Dim lGenNow As Long
    lGenNow = GetSheetGeneration(PatternFromKey(sKey))

    If pColCache.exists(sKey) Then
        Dim cached As clsColAsset
        Set cached = pColCache(sKey)
        If cached.Generation = lGenNow Then
            Log "GetColumnAsset: [CACHE HIT] " & sKey, lvlInfo
            Set GetColumnAsset = cached
            Exit Function
        End If
        ' 낡은 자산: 조용히 폐기하고 재스캔
        pColCache.Remove sKey
        Log "GetColumnAsset: [STALE] 재스캔 " & sKey, lvlInfo
    End If

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAsset(rngCol, sKey, lGenNow)
    Set pColCache(sKey) = asset
    Log "GetColumnAsset: [MISS] 신규 스캔 등록 " & sKey, lvlInfo
    Set GetColumnAsset = asset
End Function

' @desc 전역 캐시를 비우고 모든 시트 세대를 증가시킴
Public Sub InvalidateCache()
    If Not pColCache Is Nothing Then
        Log "InvalidateCache: 전역 캐시 삭제 (" & pColCache.Count & " items)", lvlInfo
        pColCache.RemoveAll
        Set pColCache = Nothing
    End If

    If Not pGenBySheet Is Nothing Then
        Dim vPattern As Variant
        For Each vPattern In pGenBySheet.Keys
            pGenBySheet(vPattern) = pGenBySheet(vPattern) + 1
        Next vPattern
    End If
End Sub

' @desc 변경된 시트의 세대를 올리고 해당 시트의 캐시 항목을 제거함
Public Sub InvalidateCacheBySheet(ByVal Sh As Object)
    ' 삭제 대상 패턴: 해당 시트 A1의 External 주소에서 ! 앞부분
    ' (작은따옴표/공백 처리가 저장 키와 100% 일치하도록 주소 API로 생성)
    Dim sTargetPattern As String
    On Error Resume Next
    sTargetPattern = Sh.Cells(1, 1).Address(External:=True)
    On Error GoTo 0

    If InStr(1, sTargetPattern, "!") > 0 Then
        sTargetPattern = Left$(sTargetPattern, InStr(1, sTargetPattern, "!") - 1)
    End If
    If LenB(sTargetPattern) = 0 Then Exit Sub

    ' 세대 증가 (캐시 보유 여부와 무관)
    Dim dGen As Object
    Set dGen = EnsureGenStore()
    dGen(sTargetPattern) = GetSheetGeneration(sTargetPattern) + 1

    ' 캐시 항목 정리
    If pColCache Is Nothing Then Exit Sub
    If pColCache.Count = 0 Then Exit Sub

    Dim colToDelete As New Collection
    Dim vKey As Variant, i As Long
    For Each vKey In pColCache.Keys
        If InStr(1, vKey, sTargetPattern, vbTextCompare) > 0 Then
            colToDelete.Add vKey
        End If
    Next vKey

    If colToDelete.Count > 0 Then
        For i = 1 To colToDelete.Count
            pColCache.Remove colToDelete(i)
        Next i
        Log "InvalidateCacheBySheet: " & colToDelete.Count & "건 삭제 (" & sTargetPattern & ")", lvlInfo
    End If
End Sub
