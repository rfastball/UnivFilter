Attribute VB_Name = "modCache"
' @desc 열 자산 캐시와 지문 신선도 관측 담당
' @depends modColumnScan, clsColAsset, clsColSig, modLogger
Option Explicit

' 캐시 항목: 키 = rng.Address(External:=True), 값 = clsColAsset (지문 스탬프 포함).
' 신선도는 세대 예측이 아니라 GetColumnAsset이 취득 시 지문을 관측·대조해 확정한다.
' (이벤트/세대 카운터 없음 - 무효화가 테스트 사각인 미검증 경로에 의존하지 않는다.)
Private pColCache As Object

' @desc 열 범위 받아 자산 반환, 캐시 미스나 지문 불일치 시 재스캔 후 반환
Public Function GetColumnAsset(ByVal rngCol As Range) As clsColAsset
    If pColCache Is Nothing Then
        Set pColCache = CreateObject("Scripting.Dictionary")
    End If

    Dim sKey As String
    sKey = rngCol.Address(External:=True)

    ' 신선도 = 관측: 현재 열의 지문을 계산해 캐시된 지문과 대조 (렌더 없는 값싼 prefix)
    Dim sig As clsColSig
    Set sig = modColumnScan.ComputeSignature(rngCol)

    If pColCache.exists(sKey) Then
        Dim cached As clsColAsset
        Set cached = pColCache(sKey)
        Dim isFresh As Boolean
        isFresh = False
        If Not cached.Signature Is Nothing Then
            isFresh = cached.Signature.Matches(sig)
        End If
        If isFresh Then
            Log "GetColumnAsset: [CACHE HIT] " & sKey, lvlInfo
            Set GetColumnAsset = cached
            Exit Function
        End If
        ' 지문 불일치: 값만 바뀐 경우 변경 행만 캐논 재렌더(증분 패치).
        ' 키는 (값, 서식)의 순수 함수(P11/P12)라 무변경 행 재사용이 관측의 재사용이다.
        If modColumnScan.TryPatchAsset(cached, rngCol, sig) Then
            Log "GetColumnAsset: [PATCH] 증분 재렌더 " & sKey, lvlInfo
            Set GetColumnAsset = cached
            Exit Function
        End If
        ' 패치 불가(행수/서식 변경, 병리적, 변경 상한 초과): 조용히 폐기하고 재스캔
        pColCache.Remove sKey
        Log "GetColumnAsset: [STALE] 재스캔 " & sKey, lvlInfo
    End If

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAssetFromSig(rngCol, sKey, sig)
    Set pColCache(sKey) = asset
    Log "GetColumnAsset: [MISS] 신규 스캔 등록 " & sKey, lvlInfo
    Set GetColumnAsset = asset
End Function

' @desc 전역 캐시를 비움(종료/프로브/테스트의 강제 청소용, 신선도와 무관)
Public Sub ClearCache()
    If Not pColCache Is Nothing Then
        Log "ClearCache: 전역 캐시 삭제 (" & pColCache.Count & " items)", lvlInfo
        pColCache.RemoveAll
        Set pColCache = Nothing
    End If
End Sub
