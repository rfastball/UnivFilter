# 하네스 차단 지점 로그 (2026-07-07, Stage 0 진입 실패)

> 상태: 내부 개편 플랜(2026-07-07_internal_overhaul_plan.md) Stage 0 착수 중 하네스 사용 불가로 중단.
> src 편집 없음, 커밋 없음. 태스크 목록의 Stage 0은 in_progress 상태로 남김.

## 타임라인

1. `vh status` / `vh lease acquire` : 정상 (lease 획득, leaseId f74f9f42).
2. `vh run` 1차 시도 : `err:no_wb` (워크북 미오픈).
3. Excel 확인 : PID 23428, 13:12:44 시작 (사람이 띄운 것으로 추정, 창 표시 중).
   `vh open` : `open:ok excelPid=[23428]` (기존 인스턴스에 워크북 오픈 성공).
4. **`vh run` 2차 (13:14) : `gate:cached` 출력 후 5분 이상 무응답.**
   jsonl에는 13:14:49 X(reset) 마커만 찍히고 이후 레코드 전무. 프로세스 강제 종료.
   (플랜 문서의 "gate:cached 이후 무한 대기" 증상과 동일하나, 원인은 달랐음. 아래 5 참조)
5. 환경 점검 결과 전제는 전부 충족 상태였음:
   - `_UnivDbg.xlam`, `CodeManager.xlam`, `UnivFilter(v2.6).xlam` 모두 IsOpen=True
   - 워크북 열림, VBE 프로젝트 4개 정상 표시
   - VBAWarnings=1(모든 매크로 허용), AccessVBOM=1, MOTW 없음, AutomationSecurity=1
6. `vh run` 3차 (13:27) : `reload:26 pass:0 fail:1` + Test_UnivFilter_SpecValidation
   **error 1004 "매크로를 실행할 수 없습니다"** (elapsed 4ms, 즉시 실패).
7. COM으로 VBA 프로젝트 직접 점검 : **컴포넌트 40개, 12쌍 중복 발견.**
   - 중복 쌍: clsColumnAsset/clsColumnAsset1, Core_clsSet/1, Core_Enums/1, LogShim/1,
     modCache/1, modFilterUtils/1, modLogger/1, modMatchingEngine/1, modResultHandler/1,
     modTester/1, Test_UnivFilterHarness/1, Utilities/1
   - 원인 추정: 4번의 크래시 사이클이 임포트 도중 죽으면서 구모듈 12개를 제거 불가
     좀비 상태로 남김. 이후 사이클마다 임포트 충돌로 "1" 접미사 사본 생성.
   - 결과: Test_UnivFilter_SpecValidation이 두 모듈에 존재하여 비한정
     Application.Run이 모호성으로 1004. (모듈 한정 호출 `modCache.InvalidateCache`는 성공 확인)
8. `vh run` 4차 (13:30) : `removed:12 imported:26` 출력했으나 실행 후 컴포넌트는 여전히
   동일한 40개 + 같은 12쌍. **mirror sync는 "1" 사본만 제거하고 좀비 구모듈 12개는 제거하지
   못하며, 재임포트 시 충돌이 매번 재발.** 메모리 내 프로젝트는 하네스로 자가 치유 불가 판정.
9. 복구 시도: 디스크의 xlsb 원본은 깨끗하므로 "저장 없이 닫고 재오픈"이 유일한 복구 경로.
   - COM 직접 `Workbook.Close(false)` : 권한 분류기가 거부.
   - **`vh close -NoSave` + `vh open` : 사용자 거부로 중단. <- 최종 차단 지점**

## 현재 상태

- 워크북 `UnivFilter(v2.6).xlsb`는 Excel(PID 23428)에 열려 있고 VBA 프로젝트는 12쌍 중복으로
  오염된 상태. **저장하면 오염이 디스크에 굳으므로 저장 금지.**
- 디스크 파일(345,752 bytes)은 미저장 상태로 깨끗함.
- 프로젝트 lease는 이 세션 소유로 남음. 소유 프로세스가 종료되므로 다음 세션에서
  `vh lease acquire`가 자동 재획득(reclaim) 가능.

## 재개 절차 (다음 세션)

1. 사람이 Excel에서 워크북을 **저장하지 않고** 닫기 (또는 Excel 전체 종료 후 통상 재실행).
2. 워크북을 통상 방식으로 다시 열기 (애드인 자동 로드 확인).
3. `vh lease acquire` 후 `vh run` : reload:26에 중복 없이 green(fail:0) 확인.
4. 이후 플랜 Stage 0(기준선 커밋)부터 재개.

## 근본 원인 추적 (2026-07-07 후속, 확정)

### 증상 재확인 (읽기 전용 COM 관찰)
12개 "1" 사본은 각각 원본과 **바이트 단위로 동일**(line 수·헤더 2줄 일치, 아래 표).
즉 `Import`가 교체가 아니라 **복제**를 수행했다.

| 모듈 | type | lines | 사본 동일 |
|---|---|---|---|
| clsColumnAsset/1 | 2 | 58 | O |
| Core_clsSet/1 | 2 | 175 | O |
| Core_Enums/1 | 1 | 11 | O |
| LogShim/1 | 1 | 120 | O |
| modCache/1 | 1 | 276 | O |
| modFilterUtils/1 | 1 | 802 | O |
| modLogger/1 | 1 | 119 | O |
| modMatchingEngine/1 | 1 | 493 | O |
| modResultHandler/1 | 1 | 221 | O |
| modTester/1 | 1 | 371 | O |
| Test_UnivFilterHarness/1 | 1 | 8 | O |
| Utilities/1 | 1 | 73 | O |

### 결정적 단서: 중복 집합 = 테스트 콜그래프
중복된 12개는 `Test_UnivFilter_SpecValidation`(=Run_SpecValidation)이 **실행 시 컴파일하는
콜그래프와 정확히 1:1 일치**한다. src 참조 추적으로 확인:

```
Test_UnivFilter_SpecValidation
 -> LogShim.AssertEq                         (LogShim)
 -> Run_SpecValidation                        (modTester)
      -> modCache.GetColumnAsset              (modCache -> Core_clsSet, clsColumnAsset, modFilterUtils)
      -> modMatchingEngine.CalculateCriteria  (modMatchingEngine -> 동일 3종)
      -> modResultHandler.ApplyFilterResult   (modResultHandler -> modFilterUtils, Utilities.OptimizedMode)
      -> clsColumnAsset -> Core_clsSet -> Core_Enums
      -> modLogger.Log                         (거의 전 모듈에서 compile-time 참조)
```
= {Test_UnivFilterHarness, LogShim, modTester, modCache, modMatchingEngine, modResultHandler,
   clsColumnAsset, Core_clsSet, Core_Enums, modFilterUtils, modLogger, Utilities} = **정확히 그 12개.**

중복 안 된 14개(관리 대상)는 테스트가 **진입하지 않는** 모듈: UI/진입 경로(modUnivFilter,
modShortcuts, modProcessManager, frmFilter01), API 심(FocusAPI, KeyDownAPI, ClipboardAPI,
AutoCloseMsgAPI), 개발툴(CodeManager), 휴면(Core_modNorm, Core_modCompare - 호출자 0), 시트/문서.

### 메커니즘 (ModuleIO.Mio_ReimportStandard, line 187~228)
표준/클래스 모듈 재주입은 아래처럼 **한 매크로 실행 안에서 Remove 직후 즉시 Import**한다:
```vba
On Error Resume Next
wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(modName)   ' line 219 - 에러 무시됨
On Error GoTo 0
wb.VBProject.VBComponents.Import tempPath                             ' line 222
```
VBIDE의 `VBComponents.Remove`는 **현재 컴파일/실행(running) 상태에 물려 있는 모듈에 대해서는
동기 제거가 안 되고 지연(defer)** 된다. line 218~220의 `On Error Resume Next`가 그 실패를
삼킨다. 그 직후 `.Import`가 즉시 실행되면 같은 이름이 아직 남아 있어 **VBIDE가 임포트 컴포넌트에
"1"을 자동 부착**한다. 반대로 러닝 상태에 없는 모듈(위 14개)은 Remove가 동기 성공 → 교체됨 → 무중복.
→ 12 vs 14 분할이 이 규칙으로 정확히 설명된다.

### 왜 평소엔 되고 이번엔 고착됐나 (flip)
평소 사이클 시작 시 대상 프로젝트는 **design(디컴파일) 상태**라 모든 모듈 Remove가 동기 성공한다.
이번엔 프로젝트가 **running/suspended 상태에 고정**되어 in-use 12개의 Remove가 매번 지연됐다.
고정의 방아쇠(정황상):
- 13:14 2차 `vh run`이 X 마커 직후 멈췄고 5분 후 강제 종료됨 → Dev_Cycle 매크로가 **VBA reset
  없이 중단**되어 런타임이 suspended로 남음.
- 추가로 modTester의 모듈 수준 객체 참조(`Private wsSpec/loSpec/wsReport`)가 Run_SpecValidation
  이후 **해제되지 않고 살아 있어** 해당 모듈군을 in-use로 붙잡는다(End 없이는 안 풀림).
- 매 `vh run`이 prune으로 지난 사이클의 "1" 12개를 지우고(removed:12), 곧바로 import가 12개를
  재생성 → **12쌍 고착 상태의 fixed point**.

### 이번 세션에서 안전 여부
- 관찰은 전부 읽기 전용 COM. src 미편집, 커밋 없음, 워크북 미저장.
- 오염은 메모리 내 VBProject에만 있음. **저장하면 12쌍이 xlsb에 굳으므로 저장 금지.**

### 권고 (요약; 세부는 아래 "재개 절차")
1. 즉시 복구: 워크북을 **저장 없이 닫고** 통상 방식으로 재오픈 → design 상태로 리셋되어 다음
   `vh run`이 무중복 reload:26이 됨. (에이전트 `vh close -NoSave`는 사용자 승인 필요 - 앞서 거부됨)
2. 하네스 항구 수정(별건, CodeManager.xlam 소유자 판단): `Mio_ReimportStandard`가 표준/클래스도
   문서 모듈처럼 **Remove+Import 대신 CodeModule.DeleteLines + AddFromString(=in-place 주입)** 로
   가면 새 컴포넌트를 만들지 않아 "1" 충돌이 원천 차단됨(이미 `Mio_InjectDocument`가 문서 모듈에
   쓰는 패턴). 트레이드오프: Import가 재파싱하던 VB_Name/클래스 Attribute를 다시 안 읽음(이 프로젝트엔 무해).

## 미해결 의문 (참고)

- 13:14 정지의 1차 방아쇠가 "워크북 오픈 직후 첫 사이클 타이밍"인지 다른 요인인지는 미확정.
  재발 시 정지 시점의 VBE 상태(break/running)를 캡처해 확증할 것.
