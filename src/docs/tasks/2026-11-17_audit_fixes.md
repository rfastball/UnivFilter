# 감사 후속 수정

## 범위와 순서

1. LogShim: vbaloop 0.6.0 관리본 일치 및 워크북 반영 확인. 기준 전체 테스트 68개 통과.
2. 클립보드: Excel 창을 소유자로 지정하고, 데이터 준비 후 비우기. 실패 시 메모리 해제 및 복사 결과에 맞는 안내.
3. 지연 인덱스: 모든 인덱스 할당이 끝난 뒤 완료 표시. 중단/오류 이후 재시도 허용.
4. 긴 문자열: 등호 스칼라 조건에서 와일드카드 문자를 이스케이프.
5. 열린 폼: 적용 직전에 원래 대상 헤더에서 범위/필드 재해석. 선택 셀 변경과 무관하게 원래 대상 유지.
6. F1~F4: 클릭 처리기를 재사용해 마지막 모드 저장도 일치.
7. 배포: README의 xlsm 일반 사용 안내 제거. 사용하지 않는 벤더 modIndexKey는 소스 유지, 배포에서 제외.

감사 5번 센티넬은 수용된 휴리스틱으로 유지한다. 6번 xlsm 실행 방식은 지원 경로가 아니므로 문서만 수정한다.
7번 합계 행은 이전 커밋 3a7c1cb에서 검증용 계수만 수정했고 필터 범위는 유지했다.

## Test delta brief

- Unique risk: 복사 실패를 성공으로 안내, 인덱스 실패 후 재시도 불가, 긴 리터럴의 과다 매칭, 열린 폼에서 추가 행 누락.
- Existing owner: Test_PL_CorrectionAndMissing, Test_PL_QuickZeroNoop, Test_AP_LongItem, Test_PL_TargetForms, Test_Generation2.
- Selected layer: 기존 Excel/Win32 통합 테스트 확장. 인덱스 재시도는 시트 없는 객체 상태 테스트.
- Expected cost delta: 테스트 1개 추가, 기존 4개 확장. 새 테스트 시트/워크북/프로세스/파일 없음. 기존 fixture 안에서 경계 사례 검증.
- Evidence to retire: 중복 또는 대체되는 독립 테스트 없음. 기존의 단순 누락 메시지 존재 단언은 실제 복사/안내 단언으로 대체.

## 검증 및 결과

위 범위의 수정을 완료했다. LogShim은 이미 설치된 0.6.0 관리본과 일치해 직접 편집하지 않았으며, push와 test로 워크북 반영을 검증했다.

- 수정 전 회귀: 69개 중 4개 실패. 긴 별표 문자열 1행 기대에 5행 선택, 물음표 1행 기대에 3행 선택. 인덱스 오류 후 재시도에서 Nothing 반환. 클립보드 소유자 누락 및 퀵 무적용 결과의 복사 성공 문구를 검출했다.
- 클립보드 쓰기 자체는 이 환경의 수정 전 코드에서도 성공했다. 소유자 계약과 실패 정리가 문제이며, 항상 복사 실패라고 단정하지 않는다.
- 폼 회귀: ListObject 행 추가/열 삽입, 다른 워크북 활성화, 빈 표, 일반 AutoFilter 범위 확장/제거. 실제 적용 결과와 원래 대상 유지를 확인했다. 개별 실행 249ms, 실패 0.
- 최종 lint finding 0, push VBE 컴파일 성공, 전체 테스트 69개 통과/실패 0/생략 0.
- 검증용 `_dbg/audit-fixes-dist/UnivFilter.xlam` 빌드 및 VBE 컴파일 성공. modIndexKey 제외 확인. 기존 배포 파일은 교체하지 않았다.
- 실패 강제 주입(메모리 할당/SetClipboardData/다른 프로세스 점유), 실제 Esc/F1~F4 키 입력 및 폼 시각 검사는 수행하지 않았다. F키는 기존 클릭 처리기에 위임한다.

근거: `_dbg/audit-fixes-{baseline,red,green,target-forms,dist}.json`, `_dbg/UnivFilter.jsonl`.

## Test delta

- Unique risk / owner and layer: 위 brief의 경계와 소유자를 유지했다. 인덱스 오류 재시도는 실제 Esc 대신 결정적인 지문 배열 오류로 검사한다.
- Added / extended / removed: 1 / 4 / 0. 단순 메시지 존재 단언 1개를 실제 기록과 성공 문구 단언으로 대체했다.
- Collected-case delta: 68 -> 69 (+1).
- Runtime delta: 같은 환경의 전체 테스트 첫 시작~마지막 종료 로그 기준 4,774ms -> 4,532ms (-242ms). 각 1회 측정이므로 성능 향상으로 해석하지 않는다. Excel 기동 시간 제외.
- Heavy-resource launch delta: 스위트당 Excel 0개 증가, 새 워크북/시트 생성 0개 증가. 배포 제외 검증은 스위트 밖에서 별도 빌드 1회.
- Coverage-floor impact: 설정된 수치 기준 없음. 기존 테스트 제거나 기준 완화 없음.
