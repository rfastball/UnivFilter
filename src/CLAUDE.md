# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

`UnivFilter` — Excel VBA add-in that adds a custom Universal Filter (in-cell button + context menu + global shortcuts) to any ListObject/AutoFilter range. The shipping artifact is `../UnivFilter(v2.6).xlsm`; the files under `src/` are the source of truth for that workbook's modules.

## Build / Test Workflow

개발 루프는 `vbaflow` (루트 `../CLAUDE.md`의 관리 블록 참조): 편집 → `vbaflow push` → `vbaflow test` → green이면 commit.

- 테스트는 `Public Sub Test_*()` (인자 없음), 단언은 `LogShim.AssertEq`. `Test_UnivFilterHarness.Test_UnivFilter_SpecValidation`이 `modTester.Run_SpecValidation True`(bSilent)를 감싸 ~39개 e2e 시나리오를 돌린다. 나머지 `Test_*` 모듈은 순수 로직 특성화 단위테스트다 — **어서션은 현행 동작의 기록이므로 "개선" 목적의 수정 금지**.
- 스펙 스위트는 실행 시 `Valid_Spec_Sheet`/`Test_Fail_Report` 시트를 생성한다. `--prune` push 전에는 `vbaflow run modTester.Cleanup_TestSheets`로 정리할 것 (떠돌이 문서 모듈이 prune을 막는다).
- 모듈 삭제(prune) 직후 스위트가 이유 없이 특정 테스트에서 크래시하면 stale p-code 의심: 해당 모듈 재push + `vbaflow session stop/start` 후 재실행 (2026-07-08 Test_SetEngine에서 관찰된 일과성 현상).
- `CodeManager.bas`(ExportAllVBAElements 등)는 수동 라운드트립 보조 도구로 남아 있다. 구 UnivCore 배포 경로(`core.manifest` + `Pkg_ExportCorePackage`)는 **폐기됨** — core.manifest는 삭제되었고 해당 프로시저는 이 프로젝트에서 무용하다.

## Vendoring — clsSet · modIndexKey (편집 본가 = UnivDB)

`clsSet.cls`(dict 기반 집합 엔진, `SetRetMode` enum 내장)와 `modIndexKey.bas`(키 정규화 계약)는 **UnivDB에서 스냅샷 벤더링**한 사본이다. 출처는 `src/vendor.manifest`(UnivDB HEAD 해시) — 도구가 소비하지 않는 순수 기록.

- 편집 본가는 `../../UnivDB/src/` — 이 사본을 직접 고치지 말고 업스트림 수정 후 재스냅샷.
- 라이브 싱크 없음, 드리프트 허용. 재스냅샷 절차: ① UnivDB에서 수정·테스트·커밋 ② 두 파일을 바이트 동일 복사 ③ vendor.manifest의 source_commit/snapshot_date 갱신 ④ push + test.
- 키 의미론 계약: Variant 키(숫자 1 ≠ 문자 "1"), `vbTextCompare`(대소문자 무시), Null/Error 키 가드. `Test_SetEngine`이 고정한다.

## Architecture

### Two entry paths share one pipeline

```
[Button click / context menu]        [Ctrl+Shift+Q / Ctrl+Shift+U]
   modUnivFilter.ShowUniversalFilter      modShortcuts.UnivFilter_*Handler
              │                                       │
              └──────────────► frmFilter01 ◄──────────┘
                   (UI만: 주입/잠금/위임)
                                    │
                                    ▼
                  modProcessManager.ExecuteFilterProcess
                   (진입점에서 세대 검증 fail-fast 9998)
                      │              │              │
                      ▼              ▼              ▼
              MatchingEngine   ResultHandler   (Missing → Clipboard)
```

- `modUnivFilter.ShowUniversalFilter`: per-column shape buttons (`SmartFilterBtn_<addr>`) 또는 컨텍스트 메뉴가 호출. `G_OpenForms`(sheetName!columnIndex 키)로 폼 중복 방지.
- 퀵 모드(Ctrl+Shift+Q): `modShortcuts` → `frmFilter01.RunQuickFilter`(2줄 위임) → `modProcessManager.ExecuteQuickFilter`. 타입 정책상 폼이 필요한 컬럼(HYBRID/ERROR/DATE)은 폼 표시로 우회.
- 단축키/메뉴 바인딩: `현재_통합_문서.Workbook_Open`(OnKey + SetupContextMenu), 해제는 Workbook_BeforeClose.

### 계층 구조 (구 modFilterUtils 803줄의 4분해)

| 계층 | 모듈 | 역할 |
|---|---|---|
| L1 | `modNormalize` | 순수 값 정규화 (GetUniqueCleanData, NormalizeValue, ExtractNumberAggressive, SafeText) |
| L2a | `modColumnType` | 서식 인지 타입 판별 (FilterDataType enum, GetColumnIdentity, IdentityToEnum, GetRepresentativeFormat) |
| L2b | `modDisplayText` | 표시 텍스트 재구성 (GetSafeFormattedText, GetCleanFormat, GetAccountingFilterString, GetPercentageFilterString) |
| L3 | `modTypePolicy` | **타입 → 모드 정책의 단일 소스** (ModePolicy UDT + GetModePolicy) |

**정책 단일 소스 = modTypePolicy.** 타입에 따른 UI 잠금(LockTag) / 엔진 문자열 강제(ForceTextStrategy) / 퀵 모드 폼 우회(RequiresFormInQuickMode)는 전부 `GetModePolicy` 진리표 하나에서 나온다. 필터 동작을 바꿀 때 이 진리표만 수정하면 폼·단축키·엔진이 함께 따라온다 (`Test_TypePolicy`가 8타입×3필드를 고정). **주의: ROUNDED/PERCENT_ROUNDED의 비대칭(UI는 잠그되 엔진 강제/퀵 우회 없음)은 검증된 표면 동작 — 정합성 목적의 "개선" 금지.**

### Column data lifecycle + 세대(Generation) 계약

`modCache.GetColumnAsset(rng)`이 열 읽기의 단일 관문. `clsColumnAsset`(데이터 `clsSet` + 서식 + 타입 + **세대 스탬프**)을 반환하며 캐시 키는 `rng.Address(External:=True)`.

- 캐시 무효화는 이벤트 구동: `현재_통합_문서.AppEvents_SheetChange/SheetCalculate` → `InvalidateCacheBySheet`. 데이터를 바꾸는 코드 경로에서 수동 무효화하지 말 것.
- **세대 계약 (UnivDB 패턴 채용)**: 무효화는 시트별 세대 카운터를 증가시키고, 자산은 취득 시점 세대를 스탬프한다. `ExecuteFilterProcess` 진입점이 `modCache.IsAssetStale`로 검증해 불일치 시 `Err.Raise 9998` — "모드리스 폼 열어둔 채 시트 편집 후 Apply" 같은 낡은 자산 사용을 fail-fast로 차단한다 (`Test_Generation`이 고정).
- "true" last row는 `max(End(xlUp), UsedRange.LastRow)` (modCache.GetColumnAsset) — 필터된 범위에서 End(xlUp)이 덜 읽는 것에 대한 의도적 방어. 단일 End(xlUp)로 "단순화" 금지.
- `NumberFormatLocal`(modCache) vs `NumberFormat`(modColumnType)의 비대칭은 현행 유지 (문서화만).

### Global state lives in `modUnivFilter`

`G_LastExcludeMode` / `G_LastStackMode`는 세션 내 폼 인스턴스 간 UI 토글 기억. `G_OpenForms`는 폼 레지스트리. 폼 로컬 상태와 중복이 아니라 세션 메모리다. (dead였던 `G_LastMatchMode`는 제거됨.)

### Module dependency map

```
현재_통합_문서 (ThisWorkbook) ─► modCache, modUnivFilter (단축키/메뉴 바인딩)
modUnivFilter ─► frmFilter01
modShortcuts ─► frmFilter01, modUnivFilter, modTypePolicy, ClipboardAPI, AutoCloseMsgAPI
frmFilter01 ─► modCache, modProcessManager, modTypePolicy, FocusAPI, AutoCloseMsgAPI, clsColumnAsset
modProcessManager ─► modCache(세대 검증), modMatchingEngine, modResultHandler, ClipboardAPI, AutoCloseMsgAPI, clsColumnAsset
modMatchingEngine ─► modNormalize, modColumnType, modDisplayText, modTypePolicy, clsSet, clsColumnAsset
modResultHandler ─► modColumnType, modDisplayText, Utilities
modCache ─► modColumnType, clsColumnAsset, clsSet
modNormalize ─► modIndexKey, modColumnType
modDisplayText ─► modNormalize
clsColumnAsset ─► clsSet
clsSet, modIndexKey                (UnivDB 벤더링 — vendor.manifest)
Test_* ─► LogShim (+ 각 대상 모듈)
```

`modLogger.Log`는 거의 모든 모듈에서 호출되지만 기본 OFF (`ENABLE_ALL_LOGGING = False`). 디버깅 시 True로, 내보내기 전 복원.

## Conventions specific to this codebase

- 주석·로그·사용자 문자열은 전부 **한국어**. 기존 코드와의 톤 일관성을 위해 유지.
- 모듈 헤더/프로시저 `@desc`는 상위 `VBA_Projects/CLAUDE.md` §2 스키마를 따른다.
- em dash(U+2014) 소스 사용 금지 — CP949 밖 문자는 `vbaflow push`가 거부한다.
- 폼은 **모드리스**(`.Show vbModeless`). `G_OpenForms` 재활성화 경로를 이해하기 전에 모달 전환 금지.
- `OnAction` 문자열 간접 호출: shape 버튼 → `ShowUniversalFilter`, 컨텍스트 메뉴 → `CallFilterFromMenu` 등. 해당 Public Sub 이름 변경 시 `modUnivFilter`의 `.OnAction =` 할당도 함께 갱신할 것.
- VBA `Public` const/enum은 프로젝트 전역: `FilterDataType`(modColumnType), `SetRetMode`(clsSet 내장), `ModePolicy`(modTypePolicy), `LogLevel`(modLogger), `DbIndexBackend`(modIndexKey, 벤더링 부산물).

## 설계 문서

- `docs/tasks/2026-07-07_internal_overhaul_plan.md` — 2026-07 내부 전면 개편의 승인 플랜 (D1~D8 결정 기록)
- 자동 생성 전체 스냅샷(`UnivFilter(v2.6).md`)은 삭제됨 — API 조회는 src를 직접 grep.
