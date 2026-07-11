Attribute VB_Name = "modCache"
' @desc 열 자산 캐시와 지문 신선도 관측 담당
' @depends modColumnScan, clsColAsset, clsColSig, modLogger
Option Explicit

' 캐시 항목: 키 = rng.Address(External:=True), 값 = clsColAsset (지문 스탬프 포함).
' 신선도는 세대 예측이 아니라 GetColumnAsset이 취득 시 지문을 관측·대조해 확정한다.
' (이벤트/세대 카운터 없음 - 무효화가 테스트 사각인 미검증 경로에 의존하지 않는다.)
Private pColCache As Object

' 상한(LRU): 항목당 ContentVec(Value2 전체)+RenderedVec를 잡아 대형 열은 수십 MB다.
' 오래 사는 세션이 열 수백 개를 필터하면 무한 성장하므로 최근접근순(LRU)으로 상한을 둔다.
' 32는 "여러 열 반복 필터" 상용 흐름이 활성 열을 축출하지 않을 여유와 무한성장 차단의
' 절충값(개수 상한). 열이 매우 크면 행수 예산 상한이 더 정밀하나 여기선 과설계.
Private Const MAX_CACHE_ITEMS As Long = 32
' 접근순 대장: item(1)=LRU(가장 오래됨), item(Count)=MRU. sKey로 키잉해 직접 제거.
Private pOrder As Collection

' @desc 열 범위 받아 자산 반환, 캐시 미스나 지문 불일치 시 재스캔 후 반환
Public Function GetColumnAsset(ByVal rngCol As Range) As clsColAsset
    EnsureCache

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
            Touch sKey
            Set GetColumnAsset = cached
            Exit Function
        End If
        ' 지문 불일치: 값만 바뀐 경우 변경 행만 캐논 재렌더(증분 패치).
        ' 키는 (값, 서식)의 순수 함수(P11/P12)라 무변경 행 재사용이 관측의 재사용이다.
        If modColumnScan.TryPatchAsset(cached, rngCol, sig) Then
            Log "GetColumnAsset: [PATCH] 증분 재렌더 " & sKey, lvlInfo
            Touch sKey
            Set GetColumnAsset = cached
            Exit Function
        End If
        ' 패치 불가(행수/서식 변경, 병리적, 변경 상한 초과): 재스캔.
        ' 같은 키 재빌드라 StoreAsset이 덮어쓴다(Count 불변 -> 축출 없음).
        Log "GetColumnAsset: [STALE] 재스캔 " & sKey, lvlInfo
    End If

    Dim asset As clsColAsset
    Set asset = modColumnScan.BuildAssetFromSig(rngCol, sKey, sig)
    StoreAsset sKey, asset
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
    Set pOrder = Nothing
End Sub

' 캐시/접근순 대장 지연 초기화.
Private Sub EnsureCache()
    If pColCache Is Nothing Then Set pColCache = CreateObject("Scripting.Dictionary")
    If pOrder Is Nothing Then Set pOrder = New Collection
End Sub

' 자산을 저장하고 접근순 대장을 갱신한 뒤, 상한 초과 시 LRU 항목을 축출한다.
Private Sub StoreAsset(ByVal sKey As String, ByVal asset As clsColAsset)
    Set pColCache(sKey) = asset
    Touch sKey
    Do While pColCache.Count > MAX_CACHE_ITEMS
        Dim lruKey As String
        lruKey = pOrder(1)
        pOrder.Remove 1
        If pColCache.exists(lruKey) Then pColCache.Remove lruKey
        Log "GetColumnAsset: [EVICT] LRU 축출 " & lruKey, lvlInfo
    Loop
End Sub

' 키를 접근순 대장 끝(MRU)으로 이동한다(기존 위치 제거 후 재추가).
Private Sub Touch(ByVal sKey As String)
    On Error Resume Next
    pOrder.Remove sKey            ' 대장에 없으면 무해(에러 무시)
    On Error GoTo 0
    pOrder.Add sKey, sKey
End Sub
