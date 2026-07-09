# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

`UnivFilter` — Excel VBA add-in that adds a custom Universal Filter (in-cell button + context menu + global shortcuts) to any ListObject/AutoFilter range. The shipping artifact is `../UnivFilter(v2.6).xlsm`; the files under `src/` are the source of truth for that workbook's modules.

2026-07-09 엔진 백지 재설계 완료: 서식 예측/타입 판별 기반의 구 엔진을 걷어내고
**표시 문자열 우주** 기반으로 재구축했다. 설계 근거는 전부 자매 프로젝트 실측이다.

## Build / Test Workflow

개발 루프는 `vbaflow` (루트 `../CLAUDE.md`의 관리 블록 참조): 편집 → `vbaflow push` → `vbaflow test` → green이면 commit.

- 테스트는 `Public Sub Test_*()` (인자 없음), 단언은 `LogShim.AssertEq`.
- 시트를 만드는 테스트는 각자 스크래치 시트를 만들고 **반드시 삭제**한다 (떠돌이 문서 모듈이 `--prune`을 막는다). 스크래치 시트 재생성 시 `modCache.InvalidateCacheBySheet`를 수동 호출할 것 — 헤드리스에선 AppEvents가 미배선이라 실환경의 SheetChange 무효화가 오지 않는다.
- 모듈 삭제(prune) 직후 스위트가 이유 없이 크래시하면 stale p-code 의심: 재push + `vbaflow session stop/start`.
- `CodeManager.bas`는 수동 라운드트립 보조 도구로 남아 있다.

## 실측 원장 (설계의 유일한 근거 - 수정 전 필독)

AutoFilterTest 프로젝트가 측정한 사실이 이 엔진의 존재 이유다:
- 61-포맷 전수: `../../AutoFilterTest/src/docs/reference/winner_table.csv`
- 갭 프로브 P1~P6: `../../AutoFilterTest/src/docs/reference/gap-probe-findings.md`

핵심 사실:
1. **AutoFilter 등호/xlFilterValues는 렌더링된 표시 문자열로 매칭한다** (값 아님).
   표시가 같으면 타입 불문 충돌(숫자 1,000 = 텍스트 "1,000"), 표시가 다르면 같은 값도 분리(0.50 vs 50%).
2. 표시 키 규칙: `*` 패딩 서식(회계)은 `TEXT(Value2, fmtLocal에서 ! 제거)` 접힌형이 키,
   그 외는 `Range.Text`가 키. #-런(좁은 열)은 AutoFit 재독. 접힌형은 배열 안에서도 매칭된다.
3. 날짜 전용 경로 불필요: Criteria2 날짜 페어는 이 환경에서 1004, 표시 문자열 배열이 날짜를 직접 매칭.
4. 배열 용량 사실상 무제한(20k/39ms). **255자 초과 항목만 조용히 0건** → 등호 스칼라(<=2건) 특례.
5. 빈칸은 배열 항목 `"="`로 포함.

**이 사실들에 반하는 "개선"(서식 파싱 재도입, 타입 잠금 부활, 날짜 특례 추가)을 하지 말 것.**
새 동작 의문이 생기면 추측하지 말고 AutoFilterTest에 프로브를 추가해 실측한다.

## Architecture

### 지배 원리

```
universe  = 열의 표시 키 집합 (modColumnScan이 읽는다 - 재구성이 아니라 관측)
matched   = 표시레인(토큰 그대로) UNION 값레인(파싱값 -> 값지도 -> 표시 키)
적용      = xlFilterValues(matched)            [포함]
          = xlFilterValues(universe - matched) [제외, clsSet.Difference + 빈칸 "="]
missing   = 양 레인 0건 토큰 (팝업 + 클립보드)
corrected = 값 레인으로만 매칭된 토큰 ("1000" -> "1,000" 자동보정 리포트)
```

### Two entry paths share one pipeline

```
[Button click / context menu]        [Ctrl+Shift+Q / Ctrl+Shift+U]
   modUnivFilter.ShowUniversalFilter      modShortcuts.UnivFilter_*Handler
              │                                       │
              └──────────────► frmFilter01 ◄──────────┘
                  (UI만: Range 주입/토글/위임 - 자산 미보유)
                                    │
                                    ▼
                     modFilterRun.RunFilter / RunQuickFilter
                      │           │           │          │
                      ▼           ▼           ▼          ▼
                 modCache    modInputParse  modMatch   modApply -> modReport
              (세대검증+재빌드)  (토큰/값해석)   (2레인)    (xlFilterValues)
```

- 퀵 모드(Ctrl+Shift+Q): 항상 즉시 실행(완전일치/포함/초기화 고정). **매칭 0건이면 아무것도 적용하지 않고 알림만** (필터/폼/클립보드 원상태).
- 타입 잠금(HYBRID/ROUNDED/DATE)은 **제거됨** — 표시 키 매칭에서는 어떤 열이든 완전일치가 안전하다(실측 P5/P6). 부활시키지 말 것.
- 단축키/메뉴 바인딩: `현재_통합_문서.Workbook_Open`(OnKey + SetupContextMenu), 해제는 Workbook_BeforeClose.

### 모듈 지도

| 모듈 | 역할 |
|---|---|
| `modFilterRun` | 파이프라인 오케스트레이션 (폼/퀵 공용 진입점, 자가검증 포함) |
| `modInputParse` | 토큰화(개행/탭, 콤마는 구분자 아님) + 값 후보 해석(통화/괄호음수/%/날짜) |
| `modMatch` | 2레인 매칭: 표시 exact/부분(InStr 스캔), 값 브리지, 누락/보정 분류 |
| `modColumnScan` | 열 1회 스캔: 표시 키 규칙 적용, 값지도/행수지도/빈칸 집계 |
| `modApply` | xlFilterValues 단일 Tier 적용, 255자 특례, 전체숨김/필드해제, 가시행 계수 |
| `modReport` | 상태/누락/보정/경고 메시지 구성 + 누락 클립보드 복사 |
| `modCache` | 자산 캐시 + 시트 세대. stale이면 **투명 재스캔** (구 fail-fast 9998 대체) |
| `clsColAsset` | Displays(clsSet) + ValueMap + RowCountByKey + BlankCount + 세대 스탬프 |
| `frmFilter01` | UI 전용: Range 주입받아 RunFilter 위임, 결과 팝업 |
| `modUnivFilter` | 버튼/메뉴/폼 위치(DPI), 전역 상태 G_LastExcludeMode/G_LastStackMode/G_OpenForms |
| `modShortcuts` | Q/U 핸들러 + SetupFormData (본문 = ListObject DataBodyRange 또는 AutoFilter.Range 경계) |

인프라(불변): `clsSet`+`modIndexKey`(UnivDB 벤더링 - vendor.manifest, 직접 수정 금지),
`LogShim`(vbaflow 관리), `modLogger`(기본 OFF), Win32 shim 4종, `Utilities.OptimizedMode`.

### Column data lifecycle + 세대(Generation) 계약

`modCache.GetColumnAsset(rng)`이 열 읽기의 단일 관문. 캐시 키는 `rng.Address(External:=True)`.

- 무효화는 이벤트 구동: `현재_통합_문서.AppEvents_SheetChange/SheetCalculate` → `InvalidateCacheBySheet`(세대 증가 + 항목 제거). 수동 무효화하지 말 것.
- **세대 계약(재설계판)**: 자산은 취득 시점 세대를 스탬프. `GetColumnAsset`이 반환 전에 현재 세대와 대조해 불일치면 **조용히 재스캔**한다. 파이프라인이 실행 시점마다 캐시를 통과하므로 "낡은 스냅샷으로 필터"는 구조적으로 불가능 (`Test_Generation2`가 고정). 구 계약의 Err 9998 팝업은 폐지됨.
- 스캔은 **주입 범위를 신뢰**한다(끝쪽 빈칸도 필터 대상 행). 범위 경계의 진실은 호출부 소유: ListObject는 DataBodyRange, AutoFilter는 AutoFilter.Range에서 헤더 제외.

### 값 지도(ValueMap) 계약

키 = `CStr(Value2)` (숫자/날짜 셀만 - Value2가 Double인 셀). 값 = 그 값을 렌더링하는 표시 키들의 clsSet.
입력 파싱(`modInputParse.TryParseValues`)도 같은 `CStr(Double)` 키로 조회 → 같은 값의 표시 쌍둥이 전부가 매칭에 합류(FACT 3 의미론을 비교연산자 없이 구현). 마지막 비트가 다른 계산값은 범위 밖(문서화된 한계).

## Conventions specific to this codebase

- 주석·로그·사용자 문자열은 전부 **한국어**.
- 모듈 헤더/프로시저 `@desc`는 상위 `VBA_Projects/CLAUDE.md` §2 스키마를 따른다.
- em dash(U+2014)/이모지 등 CP949 밖 문자 소스 사용 금지 — `vbaflow push`가 거부. 원화 기호는 `ChrW(8361)`.
- VBA `And`/`Or`는 단락평가가 없다 — Null/Nothing 가드는 중첩 If로.
- 폼은 **모드리스**(`.Show vbModeless`). `G_OpenForms` 재활성화 경로를 이해하기 전에 모달 전환 금지.
- `OnAction` 문자열 간접 호출: shape 버튼 → `ShowUniversalFilter`, 컨텍스트 메뉴 → `CallFilterFromMenu`. 해당 Public Sub 이름 변경 시 `modUnivFilter`의 `.OnAction =` 할당도 함께 갱신.
- `Public` const/enum 전역: `SetRetMode`(clsSet), `LogLevel`(modLogger), `DbIndexBackend`(modIndexKey, 벤더링 부산물). 구 `FilterDataType`/`ModePolicy`는 삭제됨.

## 설계 문서

- `docs/tasks/2026-07-09_cache_fingerprint_design.md` — 캐시 지문 재설계 관측기록 + 결정 (P8 실측 완료)
- `docs/tasks/2026-07-09_cache_fingerprint_impl_plan.md` — 캐시 지문 구현 계획 (승인, 스테이지 A~D)
- `docs/tasks/2026-07-09_engine_rebuild_plan.md` — 백지 재설계 승인 플랜 + 스테이지 기록
- `docs/tasks/2026-07-07_internal_overhaul_plan.md` — (역사) 직전 세대 내부 개편 기록
