# 엔진 백지 재설계 계획 (2026-07-09, 승인됨)

자매 프로젝트 AutoFilterTest의 실측(61-포맷 전수 + 갭 프로브 P1~P6)을 근거로,
서식 예측/타입 판별 기반의 현행 엔진을 걷어내고 표시 문자열 우주 기반으로 재구축한다.
기존 구현 무참조, 실측과 최종 목표만이 설계 입력.

실측 원장: `../../../AutoFilterTest/src/docs/reference/gap-probe-findings.md`,
`winner_table.csv` (61/61 승자 전수).

## 확정된 실측 사실 (설계 근거)

1. AutoFilter 등호/xlFilterValues는 렌더링된 표시 문자열로 매칭 (값 아님).
2. 표시가 같으면 타입 불문 충돌, 표시가 다르면 같은 값도 분리.
3. 표시 키 규칙: `*` 패딩 서식은 TEXT(Value2, fmtLocal에서 `!` 제거) 접힌형,
   그 외 Range.Text (#-런은 AutoFit 재독). 접힌형은 배열 안에서도 매칭됨(P6).
4. 날짜 arm 불필요: Criteria2 페어는 이 환경에서 1004, 표시 문자열 배열이 날짜 직접 매칭(P2).
5. 배열 용량 사실상 무제한(20k/39ms), 255자 초과 항목만 조용히 0건(P1).
6. 빈칸은 배열 내 `"="` 항목으로 포함 가능(P3).
7. 혼합서식/회계/분수 열 전부 표시 키 배열로 완전 동작(P5/P6) -> 타입 잠금의 근거 소멸.

## 지배 원리

```
universe  = 열의 표시 키 집합 (스캔으로 취득, 재구성 아님)
matched   = 문자열매칭(토큰, universe) UNION 값브리지(파싱토큰, 값->표시 맵)
적용      = xlFilterValues(matched) [포함] / xlFilterValues(universe - matched) [제외]
missing   = 양 레인 0건 토큰 (리포트+클립보드)
corrected = 값 레인으로만 매칭된 토큰 (자동보정 리포트)
```

## 모듈 구성

철거: modMatchingEngine, modColumnType, modDisplayText, modTypePolicy,
modResultHandler, modNormalize, modProcessManager, clsColumnAsset
(+ modTester, Test_Normalize/TypePolicy/FilterUtilsChar/Generation/UnivFilterHarness).

신규:
| 모듈 | 역할 |
|---|---|
| modInputParse | 입력 토큰화와 값 해석 (순수) |
| modMatch | 2레인 매칭과 누락/보정 판정 (순수) |
| clsColAsset | 열 표시집합/값지도/세대 보관 |
| modColumnScan | 열 1회 스캔해 표시/값 자산 구축 |
| modApply | 매칭 결과 AutoFilter 적용 |
| modFilterRun | 파이프라인 오케스트레이션 (폼/퀵 공용 진입점) |
| modReport | 누락/보정 리포트와 클립보드 |
| modCache (재작성) | 자산 캐시 + 시트 세대, stale이면 투명 재빌드 |

유지: frmFilter01(UI만), modUnivFilter, modShortcuts(수정), 현재_통합_문서,
clsSet+modIndexKey(벤더링), LogShim, modLogger, API shim 4종, Utilities.

## 사용자 확정 결정

1. 테스트 재베이스라인: 기존 ~39 e2e는 참고 신호, 새 스펙 스위트 작성.
2. 철거 최대 범위: 오케스트레이터 시그니처/결과 계약/자산 구조 포함.
3. 폼 잠금(HYBRID/ROUNDED/DATE) 정책은 실측 게이트: P5/P6 결과 전부 통과 ->
   잠금 제거가 실측 결론. Stage 3 착수 전 사용자 재확인.
4. 퀵 모드 0건: 아무것도 적용하지 않고(필터/폼 불변) 알림만.

## 스테이지 (각각 vbaflow test green + 커밋)

- P: AutoFilterTest 갭 프로브 + 승자 CSV (완료, 3878721)
- 0: 베이스라인 + 본 문서 기록
- 1: 순수 계층 병존 (modInputParse, modMatch, clsColAsset + 테스트)
- 2: 시트 계층 병존 (modColumnScan, modApply, modReport + 테스트)
- 3: 스왑 (modFilterRun + modCache 재작성 + 호출부 재배선, 단일 행동변경 커밋)
- 4: 구 엔진 철거 (--prune)
- 5: 재베이스라인 + src/CLAUDE.md 재작성 + 수동 스모크
