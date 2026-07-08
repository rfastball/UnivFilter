# UnivFilter 내부 전면 개편 계획 (2026-07-07 승인)

> 상태: 설계 승인 완료, 실행 전. 새 세션에서 Stage 0부터 수행.
> 플랜 원본: `C:\Users\rfast\.claude\plans\federated-soaring-sonnet.md` (이 문서가 정본 사본)

## 실행 전 체크 (2026-07-07 세션에서 확인된 운영 이슈)

`vh open`으로 에이전트가 띄운 Excel 인스턴스는 **ROT 미등록/애드인 미로드** 상태가 되어
`vh run`이 `gate:cached` 이후 무한 대기하고, `vh close`도 `not_open`으로 스킵되는 불일치가 발생했다.
- COM 자동화로 시작된 Excel은 레지스트리 OPEN 항목(`_UnivDbg.xlam`, `CodeManager.xlam`,
  `%APPDATA%\Microsoft\AddIns`)을 자동 로드하지 않는다.
- 에이전트 PowerShell에서 `GetActiveObject` attach 시 Workbooks.Count=0인 엉뚱한 프록시가 잡혔다.
- **결론: 다음 세션에서는 사람이 Excel을 통상 방식으로 실행 + `UnivFilter(v2.6).xlsb` 열기 +
  애드인 자동 로드 확인 후 시작한다** (하네스 전제의 원래 계약대로). 스트레이 EXCEL 프로세스가
  남아 있으면 먼저 정리할 것.

## Context

UnivFilter는 오래된 프로젝트라 내부 로직이 ad hoc 코드로 난잡한 상태. **표면 UX(버튼·컨텍스트메뉴·Ctrl+Shift+Q/U·폼 동작·자동필터 결과)는 불변** 전제로 내부를 갈아엎는다. UnivFilter는 "시각적 쿼리빌더"로서 UnivDB와 구조적 친족 — 안정적인 **dict 기반 경로를 공유 기반**으로 삼는다 (FastHash는 UnivDB 벤치마크상 이득 ~30%로 미미하여 배제; 사용자 확정).

사용자 확정 사항 (binding):
1. 공유 엔진 = dict 기반. **편집 본가 = UnivDB** (dictHash 계보). 라이브 싱크 없음, 드리프트 허용, 출처(git hash)만 기록.
2. 채용 수준 = **계약 2종 + dict 셋**: ① 키 정규화 계약(`modIndexKey`) 그대로 벤더링 ② 세대(Generation) 계약 패턴을 clsColumnAsset/modCache에 채용 ③ 집합 대수는 dict 셋 클래스(본가 = UnivDB `clsSet.cls`)로. rowID 인덱스 계열(DictHashIndex/DbRowList)은 이번엔 미도입 — UX 불변 전제에서 소비처 없음(출력은 AutoFilter criteria = 값 목록).
3. 범위 = 전면 개편: 엔진 교체, 타입 로직 분류·정리·단위테스트·이식, 타입→모드 정책 중복 단일화, 폼 내 비즈니스 로직 추출, MsgBox/클립보드 관심사 분리.
4. 테스트 = 이중화: 순수 로직 단위테스트 신설 + modTester(~40 e2e 시나리오)를 LogShim 보고로 개조. 리팩토링 전 green 기준선 확보.

탐색·검증으로 확인된 사실:
- 두 dict 셋(UnivDB `clsSet` / UnivFilter `Core_clsSet`)의 public API는 1:1 일치 → **어댑터 불필요, 타입명 치환만**. 차이: UnivFilter 쪽이 Null/Error 가드 강함, UnivDB 쪽 @desc 헤더 없음.
- 타입 로직은 `modFilterUtils.bas`(803줄) 집중 + modMatchingEngine/modResultHandler 위성 중복. `Core_modNorm`/`Core_modCompare`는 호출자 0개 휴면.
- 타입→모드 정책 4중 중복: `frmFilter01.StartInitialization`, `modShortcuts:25`, `modMatchingEngine:33`, `modUnivFilter:118`(StartInitialization 직후 중복 EnforceHybridMode — 검증됨).
- 미커밋 `.gitignore`에 `src/Core_*` 추가 대기 중 → Core_ 이름 벤더링과 충돌하므로 `clsSet` 이름 채택 + 해당 줄 제거.
- `G_LastMatchMode` 선언만 있고 미사용(dead).
- UnivDB `DictHashIndex`는 `clsTable` 세대계약에 결속된 key→rowID 멀티맵이라 그대로 이식 불가(이번 범위 밖; 결합도 낮추는 선행 리팩토링은 나중에 필요해지면).

## 확정 설계 결정

**D1. 벤더링 이름 = `clsSet` 채택, `Core_clsSet` 폐기.** `Core_` 접두사는 "UnivCore 싱크로 덮어씀" 계약인데 그 계약을 버리므로. `clsSet` 안에 `Public Enum SetRetMode` 내장 → `Core_Enums.bas`는 같은 커밋에서 삭제(enum 멤버명 충돌 회피). 매직넘버 `Compare(x, 4)` → `Compare(x, srAllSplit)`.

**D2. 업스트림 우선 + 벤더 파일 2종.** UnivDB `src/clsSet.cls`에 Null/Error 가드 + @desc 헤더 보강 커밋(UnivDB 쿼리 경로 무변경 — clsSet은 dictHash 계보의 "집합연산 표면"으로 공식화). **키 의미론은 현행 유지**(Variant 키: 숫자 1 ≠ 문자 "1" — CStr 계약으로 바꾸면 매칭 동작 변경 위험). 그 뒤 UnivFilter로 `clsSet.cls` + `modIndexKey.bas`(살아있는 키 계약, 무수정 이식) 바이트 동일 스냅샷. UnivDB CLAUDE.md에 "clsSet.cls·modIndexKey.bas는 UnivFilter가 스냅샷 벤더링하는 공유 요소, clsSet의 쿼리 경로 사용 금지는 유지" 명기.

**D3. 출처 기록 = `src/vendor.manifest`** (커밋됨, 도구 소비 없음):
```
format: 1
vendor: UnivDB-DictEngine
source_project: ../UnivDB
source_commit: <스냅샷 시점 UnivDB HEAD 전체 해시>
snapshot_date: <날짜>
files: clsSet.cls, modIndexKey.bas
local_delta: none
```

**D4. 타입 로직 4분해** (modFilterUtils 803줄 해체, 검증된 로직은 이동-불-재작성):

| 새 모듈 | 계층 | 이동 내용 |
|---|---|---|
| `modNormalize.bas` | L1 순수 값 정규화 | GetUniqueCleanData, ConvertToFlatArray/StringToArray/RangeToArray(Private), CleanValue→Public `NormalizeValue`, ExtractNumberAggressive→Public, + `SafeText`(modIndexKey.NormalizeIndexKey를 래핑, Error/Null 처리 추가 — Core_modNorm 아이디어 흡수) |
| `modColumnType.bas` | L2a 서식 인지 타입 판별 | FilterDataType enum, GetColumnIdentity, IsDateFmt, GetRepresentativeFormat, GetTypeName, + modMatchingEngine의 MapIdentityToEnum→Public `IdentityToEnum` |
| `modDisplayText.bas` | L2b 표시 텍스트 재구성 | GetSafeFormattedText, GetCleanFormat, GetAccountingFilterString, GetPercentageFilterString |
| `modTypePolicy.bas` | L3 정책 단일 소스 | 신규 `Public Type ModePolicy` {ForceTextStrategy, LockTag, RequiresFormInQuickMode, StrategyType} + `GetModePolicy(sColumnType)` |

**D5. 정책 통일**: 4개 소비처를 GetModePolicy로 재배선. 진리표는 **현행 동작 그대로 인코딩**(ROUNDED가 UI는 잠그되 엔진은 문자열 강제 안 하는 비대칭 포함 — 고치지 말 것, 표면 동작임).

**D6. 세대 계약 채용 (UnivDB 패턴, 코드는 로컬 구현)**: `modCache`에 시트별 Generation 카운터(AppEvents_SheetChange/Calculate에서 기존 캐시 삭제와 함께 증가), `clsColumnAsset`에 Generation 스탬프, 파이프라인 진입점(`ExecuteFilterProcess`)에서 stale 자산 사용 시 fail-fast(Err.Raise 9998, UnivDB 계약 번호 준용). `modDbContract`는 clsTable 타입 결속이라 벤더링하지 않고 패턴만 채용. stale 검출이 기존 정상 경로에서 발화하면 그 경로가 잠재 버그였던 것 — 수정하되 커밋 분리로 되돌림 용이하게.

**D7. 폼 분리**: `frmFilter01.RunQuickFilter` 본문 → `modProcessManager.ExecuteQuickFilter`(폼은 2줄 위임 유지, modShortcuts 표면 불변). `StartInitialization` 내부를 LoadColumnAsset + ApplyTypeLock으로 분할(공개 시그니처 불변). `modResultHandler.ApplyFilterResult:144` MsgBox → `Err.Raise` 재던지기, 팝업은 호출자(폼/단축키)에서 — UX 동일 유지.

**D8. modTester 개조**: `Run_SpecValidation(Optional bSilent As Boolean = False)`, MsgBox는 `If Not pSilent`. 시나리오별 `LogShim.AssertEq "spec_" & sTitle, "PASS", sResultStatus` + 조기 종료 경로 2곳도 assert. 수동 실행 동작 불변.

## 실행 스테이지 (매 스테이지: `vh run` fail:0 → `vh check` → git commit)

**전제(사람)**: Excel 실행 + UnivFilter(v2.6).xlsb 열림 + _UnivDbg/CodeManager 애드인 로드.

- **Stage 0 — 정리 기준선**: 대기 중인 working-tree 변경 6건(.gitignore, AGENTS.md, CLAUDE.md, src/AGENTS.md, src/CLAUDE.md, src/LogShim.bas) 현상태로 커밋. `vh run`/`vh check` green 확인.
- **Stage 1 — 테스트 개조(리팩토링 전)**: `modTester.bas`(D8), `Test_UnivFilterHarness.bas`(True 전달). ~40개 시나리오 assert가 하네스에 잡히는지 확인. 이미 빨간 시나리오 발견 시 현행 상태를 assert하는 known-red 고정(주석 명기) — assert 약화 금지.
- **Stage 2 — 특성화 단위테스트**: `Test_SetEngine.bas`(Compare srAllSplit 3분할/Union/Difference/LoadFromArray 에러스킵/대소문자 무시/빈 집합/숫자-문자 키 구분), `Test_FilterUtilsChar.bas`(GetCleanFormat 회계 `*`/`_`·`;`분할·General/@, GetSafeFormattedText 4단 폴백, 회계 0→`\-`·음수 괄호, 퍼센트, GetUniqueCleanData 콤마수·퍼센트/100·실패 바스켓). Stage 3~5의 안전망.
- **Stage 3 — 엔진 교체 + 출처 기록**: ① UnivDB에서 clsSet 보강 커밋(가드+@desc, em dash 금지, 키 의미론 불변) ② 스냅샷 복사 → `src/clsSet.cls`, `src/modIndexKey.bas`; `Core_clsSet.cls`/`Core_Enums.bas`/`core.manifest` 삭제; `vendor.manifest` 생성; `.gitignore`의 `src/Core_*` 줄 제거 ③ 타입명 치환: `clsColumnAsset.cls`, `modCache.bas`, `modMatchingEngine.bas`(~15개 선언), Test_SetEngine; 매직 4→srAllSplit(modMatchingEngine:148,:302). reload 모듈 수 변화 확인.
- **Stage 4 — 세대 계약(D6)**: modCache 시트별 Generation + clsColumnAsset 스탬프 + ExecuteFilterProcess 진입 검증. `Test_Generation.bas` 신설(편집 이벤트 후 stale 자산 fail-fast, 정상 재취득 경로 green).
- **Stage 5 — L1 `modNormalize`**: 프로시저 이동 + `Core_modNorm.bas`/`Core_modCompare.bas` 삭제. 호출자 갱신: modMatchingEngine:48, modFilterUtils 내 ExtractNumberAggressive 참조 2곳. `Test_Normalize.bas` 신설(타입별 NormalizeValue: 회계 대시→0, ₩/$/NBSP 제거, 콤마, 퍼센트, 공격적 추출, SafeText/NormalizeIndexKey 계약).
- **Stage 6 — L2 `modColumnType` + `modDisplayText`, modFilterUtils 폐기**: 이동 + 호출자 갱신(modCache:51, modResultHandler :38,:39,:67,:70,:74,:79, modMatchingEngine :170,:181,:390,:461, Test_FilterUtilsChar 한정자만 — assert 이름은 유지해 동작 diff 가능하게). 선택 서브스텝(6b): ApplyFilterResult가 GetColumnIdentity 재판별 대신 엔진이 담아준 `dictResult("ColumnType")` 사용(스펙 스위트로 가드, 빨간불 시 되돌림; GetRepresentativeFormat 재유도는 의도적 차이라 유지).
- **Stage 7 — L3 `modTypePolicy`**: ModePolicy UDT + GetModePolicy, 4개 소비처 재배선(modUnivFilter:118 중복 EnforceHybridMode 삭제 — 동작 동일). `Test_TypePolicy.bas`(8 타입 × 3 필드).
- **Stage 8 — UI/비즈니스 분리(D7)**: 스펙 스위트 + 수동 스모크(버튼 경로·Ctrl+Shift+Q/U, 에러 팝업 동등성 — 보호 시트 등으로 1회 강제).
- **Stage 9 — 청소 + 문서**: dead `G_LastMatchMode` 제거, @desc/한국어 주석/U+2014 점검. `src/CLAUDE.md` 재작성: §Core_* → §Vendoring(clsSet·modIndexKey + vendor.manifest + 재스냅샷 절차), 3계층 경고 → "정책 단일 소스 = modTypePolicy", 세대 계약 설명, 의존성 지도 갱신. core.manifest 삭제로 `Pkg_ExportCorePackage`가 이 프로젝트에서 무용해짐 명기. UnivDB CLAUDE.md 갱신 커밋(별도).

## 위험도 상위 지점과 완화

1. **Stage 6 호출자 대량 갱신** — 컴파일 게이트(Option Explicit) + 서식군별 스펙 시나리오 + assert 이름 안정 유지.
2. **Stage 3 모듈 삭제/추가의 하네스 리로드** — 한 vh 사이클 내 원자적 교체, reload:N 확인.
3. **Stage 4 세대 검증이 숨은 stale 사용을 표면화할 수 있음** — 그 경로가 잠재 버그; 수정하되 커밋 분리.
4. **Stage 8 quick 경로 에러 표출 동등성** — 기존 MsgBox가 ApplyFilterResult 내부에 있었음; 호출자 측 재현 + 수동 확인.
5. **Stage 1에서 기존 빨간 시나리오 발견 가능** — known-red 고정으로 정직한 기준선.
6. **Stage 7 진리표 "개선" 유혹** — 금지. ROUNDED 비대칭 포함 현행 그대로.

## 범위 밖 (가드레일)

UnivCore 무변경. UnivDB는 clsSet 보강 + CLAUDE.md 갱신만(쿼리 경로 무변경, FastHash 무관). rowID 인덱스 계열(DictHashIndex/DbRowList/DbIndexCursor) 미도입 — 필요해지면 UnivDB 쪽 clsTable 결합도 낮추는 선행 작업과 함께 후속으로. 폼 레이아웃/frx 무변경. last-row 방어(modCache)·NumberFormatLocal 비대칭 무변경(문서화만). 표면 UX 완전 동일.

## 검증 방법

- 각 스테이지: `vh run`(fail:0) → `vh log -FailedOnly`(실패 시) → `vh check` → 커밋.
- 최종: 전체 스펙 스위트(~40 시나리오) + 신설 단위테스트 전부 green, 수동 스모크(버튼/컨텍스트메뉴/단축키 2종/HYBRID·DATE 잠금 폼/missing-data 클립보드 복사).
