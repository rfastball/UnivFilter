# 캐시 지문 구현 계획 (2026-07-09)

상태: **구현 완료 (Stage A~D, 55/55 green).** clsColSig + ComputeSignature/BuildAssetFromSig
분할 + GetColumnAsset 지문 재배선 + 이벤트/세대 철거 + 문서 갱신 완료. 아래 계획대로 착수됨.

설계·실측 근거: `2026-07-09_cache_fingerprint_design.md`(관측/결정) + AutoFilterTest
`gap-probe-findings.md` P7/P8. 이 문서는 **구현 절차**만 담는다.

확정 방침(재확인):
- 신선도 = 세대 예측 -> **저장 관측 대조**(지문). `Generation ==` 비교를 `Signature.Matches` 로 대체.
- 정합성 최대: contentVec는 Value2 배열 저장-비교(충돌 0). fmtRep는 균일=문자열/혼합=블록목록 구조대조.
- 이벤트 완전 제거(옵션 2): AppEvents 배선/세대 카운터/시트별 무효화 철거. cfGuard 없음(P8: 우주 CF-불변).
- (후속·완료) modColumnScan의 `.Text` CF추종 렌더버그(P8 결론2)는 지문과 독립으로 수정 완료(아래).

## 아키텍처 변화 (before -> after)

```
before: GetColumnAsset -> GetSheetGeneration(pattern) -> cached.Generation == genNow ?
        무효화: AppEvents_SheetChange/Calculate -> InvalidateCacheBySheet(세대++ + 항목삭제)

after:  GetColumnAsset -> sig = ComputeSignature(rng) -> cached.Signature.Matches(sig) ?
        무효화 없음. 매 접근이 관측으로 신선도 확정. 이벤트/세대 소멸.
```

핵심 재사용: `ComputeSignature`가 읽는 Value2를 `BuildAsset`이 재사용해 접근당 Value2 1회만 읽는다.

## 신설 / 변경 / 철거 인벤토리

### 신설
- **`clsColSig.cls`** (신규 클래스): 지문 1건. 필드/메서드:
  - `RowCount As Long`
  - `ContentVec As Variant` (Value2 2D 배열; 빈 열이면 Empty)
  - `FmtKind As Long` (0=균일, 1=혼합, 2=병리적)
  - `FmtUniform As String` (균일일 때 NumberFormatLocal)
  - `Blocks As Collection` (혼합일 때 `Array(start,end,fmt)` 목록)
  - `Init(...)` 주입, `Matches(other As clsColSig) As Boolean` 구조 대조.
- **`modColumnScan.ComputeSignature(rng As Range) As clsColSig`** (Public): 스캔의 값싼 prefix.
  대상 해석 -> Value2 벌크 읽기 -> NumberFormatLocal 1회. 비Null이면 균일, Null이면 TryPartition
  (예산 내 성공 = 혼합/blocks, 초과 = 병리적). **렌더/AutoFit 없음.**

### 변경
- **`clsColAsset.cls`**: `mGeneration As Long` -> `mSignature As clsColSig`. `Generation` 프로퍼티 ->
  `Signature`. `Init` 마지막 인자 `lGeneration As Long` -> `sig As clsColSig`.
- **`modColumnScan.BuildAsset`**: 시그니처 `(rngCol, sID, lGeneration)` ->
  `(rngCol, sID)`. 내부에서 `sig = ComputeSignature(rngCol)` 후 `BuildAssetFromSig` 위임.
  - 신설 private `BuildAssetFromSig(rngCol, sID, sig)`: sig.ContentVec를 v2로 재사용,
    sig.FmtKind로 렌더 분기(균일=RenderBlock 1회 / 혼합=sig.Blocks 순회 RenderBlock /
    병리적=BuildPerCell). 현 `BuildMixed`의 TryPartition 호출은 ComputeSignature로 이관되어
    여기선 sig.Blocks를 그냥 순회. 렌더 로직(RenderBlock/BuildPerCell/RenderBlock 헬퍼들)은 불변.
  - asset.Init 호출 시 마지막 인자에 sig 전달.
- **`modCache.GetColumnAsset`**: 세대 조회/비교 제거. 
  ```
  sig = modColumnScan.ComputeSignature(rngCol)
  if cache.exists(key):
      if cache(key).Signature.Matches(sig): return cache(key)   ' HIT (렌더 없음)
      cache.Remove key                                          ' STALE
  asset = modColumnScan.BuildAssetFromSig(rngCol, key, sig)     ' MISS (sig 재사용)
  cache(key) = asset; return asset
  ```
  (BuildAssetFromSig를 Public으로 노출하거나, GetColumnAsset이 sig를 넘기는 BuildAsset 오버로드 사용.)
- **`modCache.InvalidateCache`** -> **`ClearCache`** 로 개명·단순화: 세대 증가 로직 삭제,
  `pColCache.RemoveAll`만. (종료/프로브/테스트의 강제 미스용. 신선도 정합성과 무관한 순수 청소.)
- **`현재_통합_문서.cls`**: `WithEvents AppEvents` 선언(18) 삭제, `Set AppEvents = Application`(22)·
  `Set AppEvents = Nothing`(38) 삭제, `AppEvents_SheetChange`(63-69)·`AppEvents_SheetCalculate`(72-78)
  삭제. BeforeClose의 `modCache.InvalidateCache` -> `modCache.ClearCache`. Workbook_Open의
  OnKey/메뉴 배선은 유지. `@depends` 유지(modCache, modUnivFilter).
- **`modPerfProbe.bas`**: `InvalidateCacheBySheet ws` 6곳(55/82/112/144/166/192) -> `modCache.ClearCache`.

### 철거
- `modCache`: `pGenBySheet`, `EnsureGenStore`, `GetSheetGeneration`, `PatternFromKey`,
  `InvalidateCacheBySheet` 전부 삭제. 헤더 `@depends`/주석에서 세대 문구 제거.

## 테스트 변경
- **`Test_Generation2.bas`** -> 지문 계약으로 재작성(이름 유지 가능):
  - `g2_generation_advanced`(48, `a2.Generation > a1.Generation`) 폐기.
  - 신규 단언: (a) 무변경 재취득 시 **동일 인스턴스**(`Is`) 반환 = HIT. (b) 셀 값 편집 후
    재취득 시 **새 자산** = 관측된 STALE. (c) 편집 없이 `InvalidateCacheBySheet` 수동호출 의존 제거.
  - `InvalidateCacheBySheet`(40,83)/`InvalidateCache`(62) 호출 제거 or `ClearCache`로 대체.
- **`Test_Pipeline.bas`**(147): `InvalidateCacheBySheet ws` 제거 — 스크래치 재생성 시 내용이 바뀌면
  지문이 자동 감지. 동일 주소·동일 내용 재사용이면 HIT가 정답이므로 제거가 맞다(확인 후).
- **`Test_Match.bas`**(113): `asset.Init ..., "TEST!A1", 1` 의 마지막 `1`(Generation) -> 테스트용
  `clsColSig`(빈 지문) 또는 `Nothing` 허용하도록 조정. Init 시그니처 변경에 맞춰 갱신.
- **신규 `Test_Signature.bas`**: 
  - `Matches` 단위: 동일 vec/fmt -> True; 값 1개 변경 -> False; RowCount 변경 -> False;
    균일<->혼합 전이 -> False; 타입 변경(숫자5 vs 텍스트"5") -> False; 병리적 -> 항상 False.
  - 열별 격리(스래싱 소멸): A열 편집 후 B열 재취득 -> HIT.
  - 서식편집 감지(결함#3): NumberFormat 변경(이벤트 없이) 후 재취득 -> STALE.
  - 혼합 블록목록 변경 감지: 한 블록 서식 변경 -> STALE.

## clsColSig.Matches / VarEq 계약 (정합성 핵심)
- `Matches`: 아래 순서, 하나라도 실패 시 False.
  1. `FmtKind = 병리적`(자신 또는 other) -> False (항상 재스캔).
  2. `RowCount` 불일치 -> False.
  3. `FmtKind` 불일치 -> False.
  4. 균일: `FmtUniform` 문자열 동일? / 혼합: `Blocks` 개수+각 `(start,end,fmt)` 동일?
  5. `ContentVec` 원소별 `VarEq` 전부 통과?
- `VarEq(a, b)`:
  - 둘 다 IsEmpty -> True; 한쪽만 Empty -> False.
  - 한쪽이라도 IsError -> **False(강제 재스캔)**. (에러 Variant `=` 비교 예외 회피. 에러셀은
    드물고 재스캔은 정합성 안전 — 문서화된 소소한 속도 한계.)
  - `VarType(a) <> VarType(b)` -> False (숫자500 != 텍스트"500", 스캔 타입태그와 동일 규율).
  - 그 외 `a = b`.
- 순서민감(행 스왑 시 불필요 재스캔 1회, 정합성 무해) — 단순성 위해 채택.

## 스테이징 (각 단계 red->green + 커밋; green 전 커밋 금지)

- **Stage A (가산, 무철거)**: `clsColSig.cls` + `ComputeSignature` 신설. `Test_Signature`의
  `Matches` 단위 테스트만 먼저(red->green). 기존 스위트 green 유지(아무것도 안 뺐으므로).
- **Stage B (재배선)**: `clsColAsset.Init` 시그니처 변경(Generation->Signature) +
  `BuildAsset`/`BuildAssetFromSig` 분할 + `GetColumnAsset` Matches 기반 재배선.
  같은 스테이지에서 Init 호출부(modColumnScan 2곳, Test_Match) 갱신. 스위트 green.
- **Stage C (이벤트/세대 철거)**: `현재_통합_문서` AppEvents 제거, modCache 세대 코드 삭제 +
  `InvalidateCache`->`ClearCache`, `InvalidateCacheBySheet` 삭제, modPerfProbe/테스트 호출부 정리,
  `Test_Generation2` 재작성. 스위트 green.
- **Stage D (문서)**: `src/CLAUDE.md`의 "세대(Generation) 계약" 절 -> "지문 계약"으로 갱신,
  "수동 무효화 금지"/"헤드리스 AppEvents 미배선" 주의 폐기(이벤트 자체 소멸). 설계 문서 상태 갱신.

각 스테이지: `vbaflow push` -> `vbaflow test` green -> commit.

## 리스크 / 엣지
- **Init 시그니처 변경 파급**: modColumnScan 2 + Test_Match 1. Stage B에서 원자적으로.
- **메모리**: 열당 ContentVec 보관(50k ~800KB). 캐시는 필터한 열 수만큼(무제한, 이벤트 없어도
  STALE 시 교체). 세션 내 유계. 폭주 시 LRU는 후속(현재 불요).
- **혼합 열 적중 비용**: ComputeSignature가 매 접근 TryPartition 실행(유계, 예산 2000). 렌더보다 훨씬 쌈.
- **stale p-code**: 클래스 신설/모듈 prune 후 크래시 시 재push + `vbaflow session stop/start`(기지 절차).
- **빈 열**: `ResolveTarget=Nothing` -> RowCount 0, ContentVec Empty, 균일 "" -> Matches 자기일관.

## 성능 벤치마크 실측 (`modPerfProbe.RunFingerprintBench`, 200회 평균)

저카디널리티 필터 열, `OptimizedMode`, ms 단위:

| 열 크기 | 미스(재빌드) | **적중(접근당)** | ComputeSig | Matches(동일) | Matches(조기탈출) |
|---|---|---|---|---|---|
| 10k comma | 19.5 | **1.93** | 0.57 | 1.35 | 0.08 |
| 50k comma | 74.2 | **10.2** | 3.18 | 6.86 | 0.63 |
| 100k comma | 136.7 | **20.1** | 6.21 | 13.7 | 1.23 |
| 50k 혼합3블록 | 156.3 | **12.7** | 5.66 | 7.05 | 0.65 |

해석:
- **적중 = ComputeSig + Matches.** 접근당 ~2ms(10k)/~10ms(50k)/~20ms(100k). 구 O(1) Long 비교
  대비 이것이 **"정합성의 대가"**. 필터는 사용자 페이스라 100k에서도 20ms = 무감각 → 설계
  "수십 ms 무감각" 주장 실증.
- **적중은 미스의 ~1/7** (100k: 20 vs 137). prefix/suffix 분할이 재빌드의 ~85%를 회피 → 분할 설계 검증.
- **지배 항은 Value2 읽기(ComputeSig)가 아니라 ContentVec 원소별 대조(Matches).** 배열 저장-비교
  (충돌0)의 대가. 선형이고 100k에서도 ~14ms. (정합성 최대화의 값이 여기 있음 - 수용.)
- **조기탈출**: 변경된 열은 Matches가 첫 원소에서 즉시 이탈(0.08~1.23ms) → stale 감지 실비용은
  사실상 ComputeSig(Value2 읽기)뿐. 스래싱 소멸과 합쳐 편집 잦은 워크북에서 유리.
- **혼합 열**: ComputeSig에 TryPartition ~2.5ms 추가(50k 3블록) - 유계, 무감각.
- 메모리: ContentVec = 행수 x ~16B(Variant). 100k ≈ 1.6MB/열, 필터하는 열 수만큼(유계).

향후 속도 회복 여지(원칙: 측정으로): Matches 지배 → 매우 큰 열에서 문제되면 값싼 사전거부
해시를 배열 옆에 병치(불일치 조기컷)하는 설계 고도화 가능. 현 수치로는 불요.

## 후속 태스크 (지문과 독립) — 완료

- **modColumnScan CF 렌더버그**(P8 결론2): **수정 완료.** P9 프로브로 오라클 확정(`WText(기본서식)`),
  `HasConditionalFormat` + CF 셀 WText 렌더로 수정. `Test_CS_CondFormatRender` green.
  상세는 설계 문서 "P8이 발견한 별개 결함" 절 + `gap-probe-findings.md` P9.
