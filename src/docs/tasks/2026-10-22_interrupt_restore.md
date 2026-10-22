# Esc/Ctrl+Break 중단 복원 검증 (2026-10-22)

RunFilter는 실행 전에 원래 EnableCancelKey를 보관하고 xlErrorHandler로 전환한다.
사용자 중단(오류 18)은 실패 결과와 "필터 실행을 중단했습니다." 메시지로 반환한다.
오류 처리 진입부터 공통 Cleanup까지 xlDisabled를 유지하여 복원 중 재중단을 막고,
계산/화면 갱신/이벤트 복원 후 원래 EnableCancelKey를 돌려준다.
Utilities는 세 설정을 모두 저장한 직후 복원 가능 상태를 표시하므로 설정 변경 중 중단도 복구 대상이다.
커서는 기존 수정대로 필터가 변경하지 않는다.

## 실제 키 입력

검증 전용 Excel, 50,000행(248자 문자열), 누락 토큰 2,000개, 폼의 부분일치 검색을 사용했다.
오류 주입 대신 실제 Windows Escape/Control_L+Break 키를 보냈다.
재중단 실험에서는 ArmInterruptRestoreProbe가 임시 시트 C1에 volatile UDF를 넣는다.
이 함수는 xlDisabled인 재계산 구간에서만 15초 대기하며, 중단 정책 자체는 변경하지 않는다.

| 실행 | 최초 중단 | 복원 중 두 번째 키 | 복원 구간(KST) | 결과 |
|---|---|---|---|---|
| Esc 단독 | Esc | 없음 | 별도 지연 없음 | 4개 설정 단언 통과 |
| Ctrl+Break 후 Esc | 23:26:28.577 | Esc 23:26:35.364 | 23:26:28.641 ~ 23:26:43.641 | 4개 설정 단언 통과 |
| Esc 후 Ctrl+Break | 23:28:12.625 | Ctrl+Break 23:28:19.037 | 23:28:12.664 ~ 23:28:27.664 | 4개 설정 단언 통과 |

키 시각은 자동화 도구의 입력 완료 시각이다. 복원 구간은 VBA 로그다.
두 재중단 입력 모두 로그로 확인한 복원 구간 내부에 있으며 구간의 EnableCancelKey는 0(xlDisabled)이다.
단독 Esc에서 사용자 중단 안내를 관측했다. 수정 전 발생했던 VBE 종료 대화상자 처리 없이 검증이 완료됐다.

| 설정 | 실행 전 | 세 실행 모두 종료 후 |
|---|---|---|
| Cursor | xlDefault (-4143) | 동일 |
| ScreenUpdating | True | 동일 |
| Calculation | Automatic (-4105) | 동일 |
| EnableEvents | True | 동일 |
| EnableCancelKey | xlInterrupt (1) | 동일 |

재현 순서: vbaloop session start, run modBenchReal.PrepareInterruptProbe,
재중단 검증 시 run modBenchReal.ArmInterruptRestoreProbe, UI에서 텍스트 포함 검색 선택/확인,
처리 중 첫 중단 키, interrupt_restore_enter 로그 후 15초 안에 두 번째 키,
run modBenchReal.InspectInterruptProbe, run modBenchReal.CleanupInterruptProbe.
전용 세션에서만 실행한다. Prepare가 기준 설정을 명시적으로 바꾸며 Cleanup이 임시 시트를 삭제한다.
검증 후 세션은 저장 없이 정상 종료했다. 원시 로그: [interrupt_restored.jsonl](../reference/2026-10-22_interrupt_restored.jsonl).

## 회귀 게이트와 비용

기존 Test_Pipeline에 Test_PL_RestoresSettings 한 개를 추가했다. 정상/오류 반환 각각에 대해
자동 계산+화면/이벤트 활성과 수동 계산+비활성 설정을 대조하고, 사용자 커서 및 xlDisabled 보존도 확인한다.
기존 실제 Excel 파이프라인 fixture를 재사용한다. 새 테스트 파일/CI lane/coverage 기준 변경/중복 테스트는 없다.
자동 수집 사례 +1, 단독 실행 669ms(1회 관측), 지속적 Excel 기동 증분 0.
15초 대기 UDF와 키 입력은 수동 프로브 전용으로 자동 스위트에 포함하지 않는다.
push VBE compile 성공, 전체 vbaloop test: 68 pass / 0 fail / 0 skipped.
전체 스위트 시간의 전후 차이는 별도로 측정하지 않았다.

보장 범위는 공통 RunFilter 실행 중 사용자 중단과 설정 복원이다.
이미 적용된 필터 결과를 트랜잭션처럼 되돌리지는 않으며, VBE Reset/End 또는 프로세스 강제 종료는 이 경로를 실행하지 않는다.
