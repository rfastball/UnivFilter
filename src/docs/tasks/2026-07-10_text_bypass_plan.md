# .Text 우회 + 고정비 감량 로드맵 (2026-07-10, 승인됨)

상태: **전 스테이지 구현 완료 (2026-07-10).** 최종 실측은 §구현 후 실측.
근거 실측은 전부 커밋된 프로브(modBenchReal)와 AutoFilterTest P11/P12.

## 배경 실측 (요약)

- 행 숨김/해제 이력이 생긴 시트는 셀별 `.Text` 단가가 행 인덱스 비례로 폭증
  (처녀 4.7us -> 앞행 7.8/뒷행 92us, 전행 패스 133ms -> 1.4s). 세션 내 비가역,
  재오픈만 원복. 벌크 Evaluate(TEXT)는 면역. (ProbeTextDecay, c825318)
- P11(AutoFilterTest): @서식 숫자 셀 매칭 키 = General 캐논형, 폭 무관.
  잘린 표시("2.13E+08")는 키 아님. **현행 .Text 키는 좁은 열 무매칭 버그.**
- P12: 에러/불리언 매칭 키 = 캐논 문자열("#N/A"/"TRUE"), 폭 무관. "##" 잘림형은 키 아님.
- 에러 표시: 고전 7종(2000~2042) 영문 불변, 신형(#SPILL!)은 로캘화("#분산!").
  NumberFormatLocal은 이력 저하 면역. (ProbeDisplayMap, bc14d6d)
- Application.Text(배열) 리프팅 = Evaluate와 동속(97.7 vs 101.6ms/91k),
  빈칸 "0" 함정/에러 전파 동일. 수식 문자열 불요. (ProbeEvalAlternatives, 8d41f7b)
- OptimizedMode 왕복 47.66ms (PrintCommunication + 전시트 DisplayPageBreaks가 주범).
  폼 New 7.8ms. 지문검증 19.9ms/적용쌍 20.3ms는 계약·Excel 바닥. (ProbeFixedOverheads, 43e266c)
- 인프로세스 Format$ 렌더(58.6ms)는 서식 파싱 재도입이라 기각. Evaluate ~100ms는
  렌더 계산 + 91k BSTR 마샬 바닥(범위 접근 무관 - 인메모리 동속으로 증명).

## 스테이지 (각각 push -> test green -> commit)

- **Stage 1 (고정비)**: Utilities.OptimizedMode를 3속성(ScreenUpdating/EnableEvents/
  Calculation) 슬림판으로. **완료(898080e)** - 왕복 47.66 -> 11.72ms (잔여는
  Calculation 복원 재계산 스케줄링 + ScreenUpdating 복원 리페인트, 의미론상 불가피).
- **Stage 2 (.Text 우회)**: modColumnScan -
  (a) 에러 셀: 고전 7종 CVErr 코드 -> 캐논 문자열 매핑(신형/미지 코드는 현행 .Text 폴백 유지),
  (b) 불리언: "TRUE"/"FALSE" 매핑,
  (c) @서식 블록의 숫자 셀: needRender에 편입해 General로 벌크 렌더 (P11).
      RenderValueKey의 @ 분기도 CStr -> WText(v,"General")로 대칭화,
  (d) BuildPerCell: .Text -> fmtLocal + WText 캐논 (CF 분기 흡수, FixHashKeys/AutoFit 소멸).
  **완료(66fb37e)** - 구현 노트: 에러 매핑 조회키는 CStr(CVErr(코드)) 런타임 파생
  (CStr 에러 문자열이 하드코딩 "Error NNNN"과 불일치 - 로캘). 회귀 테스트
  Test_CS_AtNumericCanon/Test_CS_ErrBoolCanon 추가. SL_V 필터이력 콜드 1,418 -> 27.3ms.
- **Stage 3 (Application.Text)**: RenderBlockBulk의 Evaluate 수식 조립 ->
  Application.Text(ContentVec 조각, 서식). 255자 수식 한계 폴백 클래스 소멸.
  **완료(744fe5f)** - 함정: 리프팅 반환이 크기에 따라 1차원/2차원(Nx1)으로 갈림
  (소형에서 첨자오류 9) -> 차원 불문 인덱싱 정규화. 성능 동일(DW_P 콜드 195ms 유지).
- **Stage 4 (껍데기)**: 퀵 핸들러 폼리스화(ResolveFilterTarget 공용 리졸버 분리) +
  폼 오픈 직후 GetColumnAsset 프리웜(모드리스 Show 뒤, On Error 무해 가드).
  **완료(785e212)** - 버튼/메뉴(ShowUniversalFilter)와 Ctrl+Shift+U 양 경로 프리웜.
- **Stage 5 (증분 재렌더)**: 지문 불일치 + 행수/서식rep 동일 시 ContentVec 디프로
  변경 행만 캐논 재렌더, RenderedVec 제자리 패치 + BlankCount 델타 + 지연 인덱스 드롭.
  변경 행 상한(1,000행) 초과/행수·서식 변경/병리적은 현행 전량 재스캔.
  **완료(235cd0a)** - clsColSig.FormatMatches/DiffRows 분리, clsColAsset.PatchRendered,
  modColumnScan.TryPatchAsset, modCache [PATCH] 경로. 기존 테스트의 "재빌드 = 새 객체"
  단언은 내용 신선도 단언으로 교체(불변식은 객체 신원이 아니라 낡은 스냅샷 금지).

## 구현 후 실측 (2026-07-10, 지출 fixture, ProbePatchRescan/BenchQuickPath)

| 시나리오 | 이전(실사용) | 목표 | 실측 |
|---|--:|--:|--:|
| 퀵 웜 (DW_P 91k) | ~120ms | ~60 | **58.6** |
| 퀵 콜드 (열당 세션 1회) | ~220ms | ~165 | **164.1** |
| 편집 1셀 후 재스캔 | ~120ms (@열 1.4s) | ~25 | **19.5 (패치, 동일 객체)** |
| SL_V @숫자 28k 콜드 | 183.6ms (필터이력 후 1.4s) | - | **46.9 (이력 후 27.3)** |
| 폼 모드 콜드 체감 | ~220ms | ~0 | 프리웜 은폐 (구조 반영) |
| 좁은 열 @숫자/에러/불리언 매칭 | 무매칭 버그 | 정상 | **캐논 키 (회귀 테스트 고정)** |

스위트 62/62 green. 지문 계약/Execute 반환 계약 무변경.
