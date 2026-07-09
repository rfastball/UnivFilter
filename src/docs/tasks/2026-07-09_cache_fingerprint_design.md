# 캐시 지문(Signature) 재설계 — 관측 기록 + 미결 결정 (2026-07-09)

상태: **정합성 전면 고정 완료 + P8 실측 완료, 착수 가능.** 저장-비교(contentVec) + 이벤트 제거(옵션 2).
P8이 "매칭 우주는 CF-불변"을 실측해 **cfGuard 센티넬을 폐기**(속도 회복) — CF 갭은 애초에 없었다.
단 P8이 별개의 스캔 잠복버그(`.Text`가 CF 추종)를 발견 → modColumnScan 후속 수정 필요(아래).
이 문서는 냉시작 재개용 — 배경/실측/설계/결정을 자족적으로 담는다.

## 설계 원칙 (고정, 상위 규칙)

**속도는 설계의 고도화(측정·알고리즘)로 얻는다. 정합성과의 트레이드오프로 얻지 않는다.
정합성은 가능한 최대치로 고정한다.** 이 원칙이 아래 모든 성능/정합성 판단을 지배한다:
- 정합성 갭은 "드물다/현행과 동등하다"를 이유로 수용하지 않는다 → 발견 즉시 보수적으로 닫는다.
- 닫음으로 잃은 속도는 P8 등에서 **측정으로 되찾는다**(예: P7이 폭-불변을 실측해 `.Text` 회피를
  정당화했듯, 갭을 여는 근거는 언제나 실측). 측정 전에는 안전측(재스캔)이 기본값.

## 배경 (왜 지문인가)

현재 `modCache.GetColumnAsset`은 신선도를 `cached.Generation == 현재시트세대`(O(1))로
판정한다. 이건 **예측**(이벤트가 세대를 올려줬는지 믿음)이며 세 결함을 낳는다:

| 결함 | 원인 |
|---|---|
| 스래싱 | 시트 어디든 편집 → 시트 전체 세대 증가 → 무관한 열까지 전부 재스캔 |
| 미검증 이벤트 경로 | 무효화가 AppEvents에 load-bearing인데 헤드리스 미배선(테스트 사각) |
| 서식편집 무음누락 | `NumberFormat` 변경은 SheetChange를 **안 쏨** → 서식만 바꾸면 낡은 우주로 매칭(현행 잠복결함) |

지문 = 신선도를 **관측**으로 전환(예측 폐기). 열별 관측이라 세 결함이 동시에 해소된다.

## P7 실측 (지문 설계의 입력, 완료)

프로브: `AutoFilterTest/src/modProbeGaps.P7_WidthCoupling`(`RunP7Only`).
기록: `AutoFilterTest/src/docs/reference/gap-probe-findings.md` P7.

**결과: AutoFilter 매칭 우주는 폭-불변.** 값·서식 고정, 열폭만 변경 시 xlFilterValues는
폭에 무관하게 캐논형(무제약폭 렌더)으로만 매칭하고, 폭에 잘린 표시(`1E+11`/`0.1235`/`#####`)는
0건. → **열 너비는 지문에 넣지 않는다.** 지문은 `Range.Text`(폭 종속·고비용)를 안 읽어도 됨.

### P7이 덤으로 잡은 버그 (수정 완료, green)

`modColumnScan`의 좁은열 복구가 `IsHashRun`(전부 `#`)에만 걸려 있어, General 열이 중간
좁힘에서 `1E+11`/`0.1235`(# 없음)로 잘리면 통과 못 하고 잘린 키 저장 → xlFilterValues 0건
무음 실패. **트리거를 `General OR #-런`으로 확장**(RenderBlock + BuildPerCell/FixHashKeys).
`Test_CS_GeneralNarrow` 추가(red→green), 전체 41/41 통과.
남은 트레이드오프: General은 흔한 기본서식이라 이제 General 숫자열 스캔마다 AutoFit 1회(캐시됨,
실측 미측정), 폴백 경로는 전용 테스트 없음(RenderBlock과 동일 논리).

## 지문 설계 (제안, 미착수)

우주 = 셀별 `f(Value2, NumberFormat)`. 폭·폰트·줌 무관(P7).

**결정(2026-07-09 세션2): 정합성 최대화 — 해시 대신 저장-비교.** 지문의 각 성분을
해시로 압축하지 않고 원값을 자산에 보관해 신선도검사 때 원소별 비교한다. 확률적 충돌 갭을
0으로 없앤다(아래 "남는 한계" 참조). 명칭은 "지문"이나 실체는 "저장 관측의 읽기시 대조".

```
Signature = combine( RowCount,       ' 저장 Long 비교
                     contentVec,     ' Value2 2D 배열 저장 → 원소별 대조
                                     '   대조 규율: VarType 먼저, 그다음 값 (숫자500 != 텍스트"500")
                                     '   IsEmpty/IsError 가드 (에러 Variant는 = 비교 시 예외)
                     fmtRep )        ' 균일: NumberFormatLocal 문자열 (이미 정확 비교)
                                     ' 혼합: TryPartition 블록목록 [(start,end,fmt)] 구조 대조
                                     '   (블록수 크기라 저장·대조 저렴, 해시 불요 — 충돌 0)
                                     ' 병리적(분할예산 초과): 항상-재스캔 센티넬
```
(cfGuard 없음: P8이 매칭 우주 CF-불변을 실측 → CF는 지문 신호가 아니다. "정합성 갭 처리" 참조.)

핵심: 지문은 **스캔의 값싼 prefix**다. 스캔 = [Value2] → [서식분할] → [고유값렌더] → [셋구축].
지문 = 앞 두 단계의 원값 보관·대조. `BuildAsset`을 `ComputeSignature`(prefix) + `RenderUniverse`(suffix)로
쪼개면 신선도검사·재스캔이 prefix 공유(미스 시 낭비 없음). prefix가 읽는 `v2`(현 BuildAsset:47)는
어차피 스캔이 읽으므로, 자산에 참조를 버리지 않고 보관하기만 하면 된다(추가 COM 0).

### GetColumnAsset 재배선
- `sig = ComputeSignature(rng)`
- 캐시 적중 AND `cached.Signature == sig` → 반환
- 아니면 재스캔 + Signature 스탬프 + 캐시 + 반환
- (즉 `Generation ==` 비교를 `Signature ==` 비교로 대체. clsColAsset에 Signature 필드 추가.)

### 비용
- 적중 경로: Value2 벌크 1회(P4: 10k=16ms, 50k=47ms) + O(셀) 원소대조(COM무접촉, 첫 불일치 조기탈출)
  + 서식 확인 1회(균일) 또는 TryPartition 프로브(혼합, 예산 유계). 렌더·AutoFit 없음.
  필터는 사용자 페이스라 수십 ms 무감각.
- 현재 O(1) Long 비교 대비 적중당 Value2 읽기를 새로 물지만, 그 대가로 3결함 소멸 + 스래싱 소멸.
- 메모리: 열당 Value2 배열 보관(50k행 ≈ 800KB). 필터하는 열 수만큼만 캐시되므로 유계 — 정합성의 대가로 수용.

### 정합성 갭 처리 (전부 닫힘)
- **조건부서식(CF) — P8 실측으로 해소, cfGuard 폐기.** 우려는 "CF가 표시 서식을 바꾸면 지문이
  놓친다"였다. P8(`gap-probe-findings.md` P8)이 **매칭 우주는 CF-불변**임을 실측 — 셀이 CF 형
  (`5.00`)을 표시 중이어도 xlFilterValues는 **기본서식 캐논형(`5`)만 매칭**, CF 표시형은 0건.
  P7의 폭 축 붕괴와 동형으로 **CF 축도 붕괴.** → 지문(Value2+기본서식)이 우주를 완전히 포착하며
  **CF는 지문 신호가 아니다.** cfGuard 센티넬 불필요 → 폐기(속도 회복). 세션2에서 보수적으로
  넣었던 센티넬을 측정이 제거 = 원칙("속도는 측정으로 되찾는다") 그대로.
- ~~해시 충돌~~: **닫힘.** 배열 저장-비교로 확률적 충돌 갭 제거.
- ~~약한 백스톱 의존~~: **무관.** modFilterRun 자가검증은 이제 net이 아니라 순수 안전망 — 지문이
  이미 정합성을 닫으므로 백스톱이 약한 것이 문제되지 않는다.

### P8이 발견한 별개 결함 — modColumnScan CF 렌더 (지문과 독립, 후속 수정)
P8 결론 2: **`.Text`가 CF를 추종한다**(`5`->`5.00`). 스캔(`modColumnScan.RenderBlock`)은 표시키로
`Range.Text`를 읽으므로, 숫자서식 CF가 걸린 열에서 스캔이 `5.00`을 저장하는데 AutoFilter는 `5`를
매칭 → **xlFilterValues 0건 무음 필터 실패.** 이건 신선도(지문)가 아니라 **렌더 정합성** 문제라
센티넬로 못 고친다(재스캔해도 같은 오키). 원칙상 닫아야 하는 실재 갭이나 **지문 작업과 독립**:
- 범위: 스캔이 CF를 무시하고 기본서식으로 렌더해야 함. `.Text`(CF 추종)도 `DisplayFormat`(P8에서
  CF 미반영, `G/표준` 반환)도 오라클 불가 → `WText(Value2, 기본fmtLocal)` 경로 필요하나 P4의
  WText 13/61 불일치 계열(G/표준·`*`패딩·`[색]`·`@`·`!`) 주의. 별도 프로브/설계 필요.
- 우선순위: 지문 착수와 병행 아님. 숫자서식 CF는 드물고(대다수 CF=채움/글꼴, `.Text` 불변) 현
  엔진에도 이미 있는 잠복버그 → 지문 완료 후 독립 태스크로. (여기 기록만, 미착수.)

## 이벤트/세대 메커니즘 — 확정: 옵션 2 (이벤트 완전 제거)

원칙("속도를 정합성과 트레이드오프하지 않는다")이 이 결정을 강제한다. 검토했던 세 안:
1. ~~지문 authority + 이벤트 유지(강등)~~ — 이벤트를 '미리 지우기'/선반영 최적화로. **기각:**
   선반영 이득은 현 lazy 캐시(이벤트=Remove만, 재빌드=접근시)에 **실재하지 않고**, 설령 eager warming을
   배선해 속도를 얻더라도 그 속도는 *미검증 이벤트 경로에 기댄* 것 → 원칙 위반.
2. **지문 authority + 이벤트 완전 제거** — AppEvents/InvalidateCacheBySheet/세대 철거, 매 접근 지문 계산.
   **확정.** 결함 #2(미검증 이벤트 경로)를 코드에서 소거 = 정합성이 테스트 사각에 의존하지 않음.
   지문이 유일 authority라 정합성이 관측으로 완결. 잃는 속도는 없음(선반영은 애초에 허상).
3. ~~보수: 이벤트 유지 + 지문 보조~~ — **기각:** 스래싱/서식누락 절반만 해결, 세대-authority가 잔존.

~~세부 미결: contentSig 해시 vs 배열 저장-비교~~ → **확정: 배열 저장-비교**(위 "지문 설계" 결정).

### 철거/신설 범위 (옵션 2 귀결)
- 철거: `pGenBySheet`/`GetSheetGeneration`/`InvalidateCache(BySheet)`/세대 스탬프, `현재_통합_문서`의
  AppEvents_SheetChange/SheetCalculate 배선, clsColAsset.Generation.
- 신설: clsColAsset에 `ContentVec`(Value2 배열)/`FmtRep` 보관 + `Matches(rng)` 대조 메서드.
  `BuildAsset` → `ComputeSignature`(prefix) + `RenderUniverse`(suffix) 분할. (cfGuard 없음 — P8)
- GetColumnAsset: `캐시적중 AND cached.Matches(rng)` → 반환, 아니면 재스캔. (세대 비교 완전 대체.)
- 테스트: 헤드리스 AppEvents 미배선 이슈가 **소멸**(이벤트 자체가 없어짐) — `InvalidateCacheBySheet`
  수동호출 요구(현 src/CLAUDE.md)도 폐기. `Test_Generation2`는 지문 기반으로 재작성.

## 커밋 안 된 작업 (워크트리 상태)

둘 다 push됨, green, **미커밋** (main 브랜치 — 커밋 시 브랜치 분기 후):
- AutoFilterTest: `modProbeGaps.bas`(P7 추가), `docs/reference/gap-probe-findings.md`(P7 기록)
- UnivFilter: `modColumnScan.bas`(버그수정+헤더주석), `Test_ColumnScan.bas`(Test_CS_GeneralNarrow)
