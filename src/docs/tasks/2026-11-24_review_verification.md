# 2026-11-24 로직 리뷰 가설 실측

리뷰에서 나온 가설 4건을 실기(vbaloop run)로 판정한 기록. 재현: UnivFilter
`modPerfProbe.RunReviewE2E` / `modPerfProbe.RunFingerprintBench`, AutoFilterTest
`modProbeGaps.RunP13Only`(패치: `../reference/2026-11-24_autofiltertest_p13_probe.patch`).

## P13 (AutoFilterTest) - 텍스트 섹션 서식 / 빈 표시 셀

| 셀 (값 / 서식) | 화면 | 매칭되는 criteria | 매칭 안 됨 |
|---|---|---|---|
| "abc" / `"ID-"@` | ID-abc | `arr["ID-abc"]`, `="ID-abc"` (+ 텍스트 쌍둥이 "ID-abc" 합류) | `arr["abc"]`, `="abc"` |
| "xyz" / `0;-0;0;"["@"]"` | [xyz] | `arr["[xyz]"]` | `arr["xyz"]` |
| 5 / `;;;` | (빈) | `arr[""]`, `arr["="]` | `arr["5"]` |
| "hid" / `;;;` | (빈) | `arr[""]`, `arr["="]` | `arr["hid"]` |

- 텍스트 셀도 원장 1번(표시 문자열 매칭)을 따른다. **텍스트 키 = CStr(값) 규칙은
  텍스트 섹션이 있는 서식에서 틀린다** (`modColumnScan` 벌크/셀별/패치 3곳 공통).
- `arr[""]`와 `arr["="]`는 동치: 진짜 빈칸 + 표시가 빈 비빈칸 셀을 함께 표시한다.

## e2e (UnivFilter 실제 RunFilter)

| 케이스 | 결과 | 판정 |
|---|---|---|
| H1 `"ID-"@` 열에 "ID-abc" 입력 | 누락 1건, 0행 | **결함 확정** (보이는 대로 입력하면 못 찾음) |
| H1 같은 열에 "abc" 입력 | 매칭 1건인데 0행, 자가검증 경고 | **결함 확정** |
| H2 제외 모드, 300자 값 + a + b 에서 a 제외 | 1행(기대 2행), 경고 "필터에 반영되지 못했습니다" | **결함 확정** (긴 값 행이 숨겨짐, 문구는 반대 의미) |
| H3 제외 모드, `;;;` 숫자 + a + b 에서 a 제외 | 2행, 경고 없음 | 정상 (키 "" = AutoFilter 의미와 일치) |
| H3 `;;;` 숫자 5에 "5" 입력 | 누락 | 정상 (화면이 비어 있으므로 표시 의미론과 일치) |

## 성능 가설 - 적중 경로 ContentVec 사본

`RunFingerprintBench` perOp(ms), 200회 평균. text_100k는 이번에 추가한 텍스트 열 케이스.

| 단계 | comma_100k hit / matches | text_100k hit / matches / computesig |
|---|---|---|
| 기준 | 39.4 / 32.9 | 59.3 / 37.4 / 20.4 |
| VecEqual 인라인 | 14.9 / 8.5 | 37.6 / 15.4 / 20.6 |
| + Matches 사본 제거 (ContentEquals ByRef) | 13.6 / 7.2 | 32.5 / 10.4 / 20.6 |
| + ComputeSignature To2D 사본 제거 | **11.4** / 7.2 | **22.0** / 10.4 / **10.8** |

- 원 가설(사본이 비용)은 **부분 적중**: 지배 항은 사본이 아니라 원소당 `VarEq` 함수
  호출이었다(숫자 100k 33ms 중 사본 ~1ms). 세 변경 합계로 적중 경로 숫자 3.5배,
  텍스트 2.7배 단축. 전체 테스트 67/67 green.
- 남은 사본(BuildAssetFromSig/EnsureIndexes/TryPatchAsset의 ContentVec 취득)은
  미스/패치 경로 전용이라 손대지 않았다.

## 미결 (사용자 결정 필요)

1. H1 수정 방식: 텍스트 섹션이 있는 서식의 텍스트 셀 키를 어떻게 얻을지.
   서식 파싱 재도입은 원장 금지 사항이므로 `.Text` 셀별 폴백 또는 `TEXT()` 렌더
   (숫자처럼 보이는 텍스트의 강제 변환 위험 - 추가 프로브 필요) 중 선택.
2. H2 수정 방식: 제외 모드에서 255자 초과 값이 여집합에 있으면 적용 거부(기존 필터 유지)
   또는 경고 문구를 "숨겨졌습니다"로 교정.
