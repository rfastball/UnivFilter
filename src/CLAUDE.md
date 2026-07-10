# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

`UnivFilter` — Excel VBA add-in that adds a custom Universal Filter (in-cell button + context menu + global shortcuts) to any ListObject/AutoFilter range. The shipping artifact is `../UnivFilter.xlsm`; the files under `src/` are the source of truth for that workbook's modules.

2026-07-09 엔진 백지 재설계 완료: 서식 예측/타입 판별 기반의 구 엔진을 걷어내고
**표시 문자열 우주** 기반으로 재구축했다. 설계 근거는 전부 자매 프로젝트 실측이다.

## Build / Test Workflow

개발 루프는 `vbaflow` (루트 `../CLAUDE.md`의 관리 블록 참조): 편집 → `vbaflow push` → `vbaflow test` → green이면 commit.

- 테스트는 `Public Sub Test_*()` (인자 없음), 단언은 `LogShim.AssertEq`.
- 시트를 만드는 테스트는 각자 스크래치 시트를 만들고 **반드시 삭제**한다 (떠돌이 문서 모듈이 `--prune`을 막는다). 캐시 무효화는 수동 호출이 불필요하다 — 신선도를 지문 관측이 담당하므로 값·서식을 바꾸면 다음 `GetColumnAsset`이 자동 감지한다 (시트 재생성 간 결정성만 필요하면 `modCache.ClearCache`).
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

### 지배 원리 (지연 우주 - docs/tasks/2026-07-10_lazy_universe_design.md)

```
universe  = 열의 행별 렌더키 배열 RenderedVec (modColumnScan이 블록별 Evaluate(TEXT)
            1회로 관측 - 91k행 113ms, 폭/CF 무관 캐논형)
집합/지도 = 우주가 아니라 우주에 대한 '지연 인덱스' - 요구하는 소비자(제외/부분일치)가
            최초 접근할 때 1회 구축(고카디널 ~180ms), 자산에 캐시
matched   = 표시레인(토큰 멤버십 스캔, 행계수 부산물) UNION
            값레인(파싱값 -> 블록서식 RenderValueKey 후보키 -> 멤버십)
적용      = xlFilterValues(matched)            [포함]
          = xlFilterValues(universe - matched) [제외, clsSet.Difference + 빈칸 "="]
missing   = 양 레인 0건 토큰 (팝업 + 클립보드)
corrected = 값 레인으로만 매칭된 토큰 ("1000" -> "1,000" 자동보정 리포트)
```

실측(지출 fixture 91k행/고유 47k): 퀵 콜드 1,006 -> **195ms**, 웜 98ms(매칭 7.8ms),
부분일치 콜드 1,186 -> 801ms. 제외는 여집합 배열 유지 - **P10 실측으로 `"<>"` 스칼라
기각**(부등호는 값 강제 의미론: 표시쌍둥이 과소/값쌍둥이 과잉 제외, 회계 접힌형 매칭
불가). 고카디널 제외 ~4.4s(apply 지배)는 문서화된 한계.

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
              (지문검증+재빌드)  (토큰/값해석)   (2레인)    (xlFilterValues)
```

- 퀵 모드(Ctrl+Shift+Q): 항상 즉시 실행(완전일치/포함/초기화 고정). **매칭 0건이면 아무것도 적용하지 않고 알림만** (필터/폼/클립보드 원상태).
- 타입 잠금(HYBRID/ROUNDED/DATE)은 **제거됨** — 표시 키 매칭에서는 어떤 열이든 완전일치가 안전하다(실측 P5/P6). 부활시키지 말 것.
- 단축키/메뉴 바인딩: `현재_통합_문서.Workbook_Open`(OnKey + SetupContextMenu), 해제는 Workbook_BeforeClose.

### 모듈 지도

| 모듈 | 역할 |
|---|---|
| `modFilterRun` | 파이프라인 오케스트레이션 (폼/퀵 공용 진입점, 자가검증 포함) |
| `modInputParse` | 토큰화(개행/탭, 콤마는 구분자 아님) + 값 후보 해석(통화/괄호음수/%/날짜) |
| `modMatch` | 2레인 매칭: 완전일치는 렌더키 멤버십 스캔(행계수 부산물), 값 브리지는 RenderValueKey 후보키, 부분일치는 지연 인덱스 InStr |
| `modColumnScan` | 열 1회 스캔: 블록별 벌크 Evaluate(TEXT)로 렌더키 배열 생산. `ComputeSignature`(값싼 prefix)/`BuildAssetFromSig`(suffix) 분할, 병리적/벌크실패는 셀별 폴백 |
| `modApply` | xlFilterValues 단일 Tier 적용, 255자 특례, 전체숨김/필드해제, 가시행 계수 |
| `modReport` | 상태/누락/보정/경고 메시지 구성 + 누락 클립보드 복사 |
| `modCache` | 자산 캐시 + 지문 신선도 관측. 불일치면 **투명 재스캔** (이벤트/세대 없음) |
| `clsColAsset` | RenderedVec(행별 렌더키) + 지연 인덱스(Displays/ValueMap/RowCountByKey) + BlankCount + 지문 스탬프 + 멤버십 스캔 프리미티브(ScanOne/ScanCount) |
| `clsColSig` | 열 지문: RowCount + ContentVec(Value2 배열) + 서식rep. `Matches` 저장-비교(충돌0) |
| `frmFilter01` | UI 전용: Range 주입받아 RunFilter 위임, 결과 팝업 |
| `modUnivFilter` | 버튼/메뉴/폼 위치(DPI), 전역 상태 G_LastExcludeMode/G_LastStackMode/G_OpenForms |
| `modShortcuts` | Q/U 핸들러 + SetupFormData (본문 = ListObject DataBodyRange 또는 AutoFilter.Range 경계) |

인프라(불변): `clsSet`+`modIndexKey`(UnivDB 벤더링 - vendor.manifest, 직접 수정 금지),
`LogShim`(vbaflow 관리), `modLogger`(기본 OFF), Win32 shim 4종, `Utilities.OptimizedMode`.

### Column data lifecycle + 지문(Signature) 계약

`modCache.GetColumnAsset(rng)`이 열 읽기의 단일 관문. 캐시 키는 `rng.Address(External:=True)`.

- **지문 계약**: 신선도는 예측이 아니라 **관측**이다. `GetColumnAsset`이 매 접근에 `ComputeSignature`(값싼 prefix: Value2 벌크 + NumberFormatLocal, 렌더 없음)로 현재 지문을 계산해 캐시된 자산의 지문과 `Matches`로 대조하고, 불일치면 **조용히 재스캔**(지문 재사용해 Value2 재독 없음)한다. 파이프라인이 실행 시점마다 캐시를 통과하므로 "낡은 스냅샷으로 필터"는 구조적으로 불가능(`Test_Generation2`가 고정).
- **정합성 최대 고정**(설계 근거: `docs/tasks/2026-07-09_cache_fingerprint_design.md`): 지문 = RowCount + ContentVec(Value2 배열 저장-비교, 충돌0) + 서식rep(균일=문자열/혼합=블록목록). 값·서식 변경을 전부 관측으로 잡는다. **이벤트/세대 없음** — 무효화를 미검증 이벤트 경로에 맡기지 않는다. CF는 우주 불변(P8 실측)이라 지문 대상 아님. CF 열 렌더는 `.Text`(CF 추종) 대신 `WText(기본서식)`(P8 결론2/P9 수정 완료).
- 스캔은 **주입 범위를 신뢰**한다(끝쪽 빈칸도 필터 대상 행). 범위 경계의 진실은 호출부 소유: ListObject는 DataBodyRange, AutoFilter는 AutoFilter.Range에서 헤더 제외.

### 값 레인 계약

- **스캔 경로**(RenderedVec 자산, 완전일치): 입력 파싱값 v(`modInputParse.TryParseValues`)에
  대해 각 블록 서식 f로 `RenderValueKey(v, f)` 후보키를 만들어 우주 멤버십으로 판정 —
  "그 값이 이 서식 블록에 있었다면 어떤 표시키인가"의 재현이라 우주 구성과 대칭이다.
- **인덱스 경로**(병리적/레거시 자산): ValueMap 조회. 키 = `CStr(Value2)` (Value2가 Double인
  셀만), 값 = 그 값을 렌더링하는 표시 키들의 vbNullChar 연결 문자열.

두 경로 모두 같은 값의 표시 쌍둥이 전부가 매칭에 합류한다(FACT 3 의미론을 비교연산자
없이 구현). 마지막 비트가 다른 계산값은 범위 밖(문서화된 한계).

## Conventions specific to this codebase

- 주석·로그·사용자 문자열은 전부 **한국어**.
- 모듈 헤더/프로시저 `@desc`는 상위 `VBA_Projects/CLAUDE.md` §2 스키마를 따른다.
- em dash(U+2014)/이모지 등 CP949 밖 문자 소스 사용 금지 — `vbaflow push`가 거부. 원화 기호는 `ChrW(8361)`.
- VBA `And`/`Or`는 단락평가가 없다 — Null/Nothing 가드는 중첩 If로.
- 폼은 **모드리스**(`.Show vbModeless`). `G_OpenForms` 재활성화 경로를 이해하기 전에 모달 전환 금지.
- `OnAction` 문자열 간접 호출: shape 버튼 → `ShowUniversalFilter`, 컨텍스트 메뉴 → `CallFilterFromMenu`. 해당 Public Sub 이름 변경 시 `modUnivFilter`의 `.OnAction =` 할당도 함께 갱신.
- `Public` const/enum 전역: `SetRetMode`(clsSet), `LogLevel`(modLogger), `DbIndexBackend`(modIndexKey, 벤더링 부산물). 구 `FilterDataType`/`ModePolicy`는 삭제됨.

## 설계 문서

- `docs/tasks/2026-07-10_lazy_universe_design.md` — 지연 우주 재설계: 실측 기록 + 스테이지 구현 (완료)
- `docs/tasks/2026-07-09_cache_fingerprint_design.md` — 캐시 지문 재설계 관측기록 + 결정 (P8 실측 완료)
- `docs/tasks/2026-07-09_cache_fingerprint_impl_plan.md` — 캐시 지문 구현 계획 (승인, 스테이지 A~D)
- `docs/tasks/2026-07-09_engine_rebuild_plan.md` — 백지 재설계 승인 플랜 + 스테이지 기록
- `docs/tasks/2026-07-07_internal_overhaul_plan.md` — (역사) 직전 세대 내부 개편 기록
