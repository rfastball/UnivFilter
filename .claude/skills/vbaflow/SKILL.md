---
name: vbaflow
description: Excel VBA 프로젝트의 자율 개발 루프 절차. VBA 소스(.bas/.cls/.frm) 편집·주입·테스트·디버깅, .xlsm 워크북 매크로 작업, vbaflow 명령(session/push/pull/test/run/diff/lint/doctor) 사용, 실패 복구(poisoned/stale 세션, lease 충돌, 타임아웃, 컴파일 오류 모달)가 필요할 때 사용한다.
---

# vbaflow 워크플로우

vbaflow는 src 텍스트 파일과 Excel 워크북 VBA를 안전하게 왕복시키는 원자 명령 집합이다.
이 스킬은 원자 명령을 엮는 절차와 실패 플레이북을 담는다. 하드 불변식(CP949, 관리 사본
수정 금지 등)은 상위 CLAUDE.md 규약을 따른다. 명령 레퍼런스: `_VbaFlow/README.md`.

## 판독 규칙 (모든 명령 공통)

- 항상 `--json`을 붙인다. 봉투: `{command, ok, exitCode, hint, errors[], payload}`
- 분기는 두 필드로만: `ok`(성공 여부)와 `hint`(다음 행동)
- exit code: 성공 0 / 환경 오류 3 (세션 부재·stale·오염, Excel/COM 치명 - `session start` 후 재시도) / 그 외 2
- 실패 시 `hint`가 다음 행동을 지시한다. hint를 따르는 것이 항상 최우선이다
- 상태를 기억하거나 추론하지 말 것 - "재검증 필요 여부"는 도구가 판단한다(해시 캐시, 멱등 start)

## 표준 루프

1. `vbaflow init --check --json` - 온보딩/드리프트 검사. 비정상(exit 2)이면 `vbaflow init`으로 수렴 후 진행
2. src\*.bas/.cls/.frm 편집
3. `vbaflow push --json` - 미변경이면 no-op. 실패 시 errors에 파일:라인 인라인(예: CP949 불가 문자)
4. `vbaflow test --json` - payload.tests[].detail에 실패 어서션(expected/actual)이 인라인됨
5. fail이면 detail로 원인 파악 후 2로. green(payload.summary.fail == 0 그리고
   payload.summary.skipped == 0)이면 git commit
6. `vbaflow session stop --json` - 저장 후 종료. 변경을 버릴 땐 `--no-save`

세션은 자동 확보된다(auto-ensure): push/pull/test/run이 세션 없음/stale/오염을 만나면
스스로 정리 후 기동하고, 새로 띄운 경우 봉투 `session: {action:"started", pid}`로 알린다.
`vbaflow session start`를 따로 부를 필요 없음(멱등이라 불러도 무해 - 진단/수동용).

테스트는 `Public Sub Test_*()`(인자 없음)만 수집된다. 단언은 `LogShim.AssertEq`.

## 실패 플레이북

### 세션 poisoned / stale

- 평시엔 만날 일이 줄었다 - 세션 필요 명령이 진입 시 스스로 정리+재기동한다(auto-ensure)
- exit code 3이면 자동 기동 자체가 실패한 것(Excel 기동/COM 치명) - JSON 파싱 없이도
  `session start` 후 재시도로 분기 가능. 반복되면 `vbaflow doctor --json`으로 환경 점검
- 치명적 COM 오류나 타임아웃 kill 후 세션은 poisoned로 마킹된다 - 다음 명령의
  auto-ensure가 정리하므로 별도 복구 절차 불필요
- 진단이 필요하면 `vbaflow status --json`으로 4상태(none/active/poisoned/stale) 확인

### lease 충돌 (다중 에이전트)

- 변경성 명령(session/pull/push/test)이 자동 획득한다. 평시엔 신경 쓸 일 없음
- 다른 에이전트가 보유 중이면 즉시 exit 2 + 소유자/유휴시간이 표시된다
- 유휴가 짧으면 잠시 후 재시도. 소유자가 죽었거나 1시간 유휴면 자동 회수된다
- 확실히 버려진 lease만 `vbaflow lease break --force`로 강제 해제

### test 타임아웃 / 모달 / 세션 사망

- 컴파일 오류/런타임 오류 모달은 워치독이 자동 해제하고 해당 테스트만 fail 처리한다
  (detail에 모달 본문이 담김 - 이것이 실제 오류 메시지)
- 컴파일 오류 해제는 Excel 종료(세션 사망)를 유발할 수 있다. 이때 test는 즉시 중단하고
  잔여 테스트를 `result: "skipped"`로, `payload.sessionDied: true`로 보고한다
- **fail 1건 + skipped 다수 패턴이면 첫 fail의 detail(모달 본문)이 유일한 진짜 신호다** -
  skipped는 미실행일 뿐이니 원인 분석에 쓰지 말 것. 코드 수정 후 push + test로 재검증
  (세션은 자동 복구 안내를 따른다 - hint가 session start를 지시하면 그대로)
- MsgBox/무한루프는 타임아웃 kill로 회수된다(기본 120초, `--timeout <초>`) - 세션은 poisoned
- 테스트 코드에 MsgBox/InputBox를 넣지 말 것 - 헤드리스에서 행 유발

### push 거부

- CP949 불가 문자: errors의 파일:라인을 보고 해당 문자를 제거하거나 ChrW로 대체
- src에 없는 워크북 모듈: 의도한 삭제면 `--prune`, 아니면 pull로 먼저 회수
- 문서 모듈(ThisWorkbook/Sheet류)은 in-place 교체만 가능 - 신규 생성 불가
- 워크북과 src가 어긋난 의심(고쳤는데 같은 오류 반복, push가 계속 `cached:true`):
  `vbaflow push --force --json`으로 지문 캐시를 우회해 전량 재주입한다
  (`_dbg\sync.json` 수동 삭제는 불필요). 실물 확인은 `vbaflow diff --json`

## pull 주의

- 동기화되지 않은 src 편집(지문이 마지막 push/pull 어느 쪽과도 불일치)이 있으면 pull이 차단된다.
  기본 대응은 `vbaflow push`로 먼저 반영. 워크북 쪽이 진실이면 `vbaflow pull --force`로 덮어쓴다
- 사람이 세션 Excel에서 직접 편집한 뒤에만 pull이 정당하다
- 확인만 필요하면 pull이 아니라 `vbaflow diff`(읽기 전용)를 쓴다 - pull은 src를 다시 쓴다
- 최초 pull(동기화 이력 없음)은 가드 없이 통과하며 일회성 포맷 정규화(BOM/개행) diff를 만든다 - 정상

## 폼 (.frm/.frx)

- 코드만 변경: 일반 push로 충분(in-place)
- 사람이 세션에서 폼을 수동 편집(특히 그림 교체): `vbaflow push --force-form-import --json`
- src에서 .frx를 직접 교체한 경우는 push가 자동 반영한다
- .frm 텍스트와 .frx는 짝 - lint가 짝 누락을 잡는다

## 왕복 한계 (설계상 불가 - 우회로 대응)

- 프로시저 수준 Attribute(VB_ProcData 매크로 단축키)는 왕복 불가, 최초 pull에서 소실.
  단축키는 Application.OnKey 런타임 등록으로 이전할 것
- 식별자 대소문자는 push 후 재pull에서 1회성 diff가 날 수 있다(수렴 후 안정 - 정상)
- 폼 그림은 frm 텍스트에 안 나타난다 - 수동 교체 감지 불가(위 --force-form-import로 대응)

## VBA pitfall 참조 (조회/기록)

- **조회**: 신규 VBA 작성 전, 원인불명 컴파일/런타임 오류 디버깅 전에
  `references\vba-pitfalls.md`를 확인한다. 실측으로 축적된 함정 목록이라 짧다 - 훑는 비용이 싸다
- **기록**: 목록에 없는 pitfall로 디버깅 사이클을 태웠다면 정본에 추가한다:
  1. 정본은 `..\_VbaFlow\assets\vbaflow-skill\references\vba-pitfalls.md`
     (프로젝트-로컬 사본은 관리 파일 - 직접 수정 금지, init이 덮어씀)
  2. 기존 항목과 중복 확인 후, 문서 머리의 수록 기준·항목 포맷을 따라 추가
  3. `_VbaFlow`에서 `.\rollout.ps1` 실행 - test -> publish -> 전 프로젝트 init 일괄 수렴 원샷
     (publish만 따로 하면 각 프로젝트는 다음 루프의 init --check가 드리프트를 감지해 수렴)
  4. `..\_VbaFlow` 레포가 없는 환경이면 프로젝트 문서(관리 블록 밖)에 후보로 메모만 남긴다

## vbaflow 자체 결함 리포트

- vbaflow 명령의 버그/설계 위반(불변식 깨짐, 잘못된 판정, 데이터 손상)을 만나면
  우회로 넘기지 말고 리포트를 남긴다: `vbaflow bugreport "<제목>" --json`
- payload.path의 생성 파일을 열어 TODO/괄호 자리(심각도/요약/증거/재현)를 채운다.
  증거는 실행한 명령과 --json 출력 원문을 코드블록으로. exe 빌드 정보는 자동 기입된다
- 중복 확인: `vbaflow bugreport --list --json` (resolved=false가 열린 리포트 - 같은 증상이면
  새 리포트 대신 기존 리포트에 증거를 추기)
- `_VbaFlow` 레포가 없는 환경이면 명령이 exit 2로 안내한다 - 프로젝트 문서(관리 블록 밖)에
  메모로 남길 것

## 보조 명령

- `vbaflow doctor --json` - Excel/COM/AccessVBOM 환경 점검(세션 무관). 환경 의심 시 최우선.
  창 없는 고아 EXCEL 프로세스도 리포트한다 - 정리는 사람이 확인 후 `--kill-orphans`로만
- `vbaflow diff --json` - 워크북 VBProject vs src 읽기 전용 비교. 모듈별 verdict
  (same/changed/onlyInWorkbook/onlyInSrc/skipped)와 첫 차이 행(firstDiffLine)을 payload로 보고.
  워크북에 실제로 뭐가 들어있는지 확인할 때는 항상 이것 - 세션 Excel에 PowerShell 등으로
  직접 COM 접근은 금지(하드 불변식). `--module <이름>`이면 해당 모듈의 워크북 쪽 본문 전체를
  payload.module.workbookBody로 덤프. 활성 세션 필요(attach 전용, 자동 기동 없음 - 없으면
  exit 3). 폼은 코드비하인드만 비교(레이아웃/.frx 제외). 차이가 있어도 exit 0 -
  payload.clean으로 판독
- `vbaflow lint --json` - 규약/인코딩/헤더 전수 검사(Excel 불필요). push 전 일괄 사전 점검
- `vbaflow run <Module.Proc> --json` - 임의 매크로 단건 실행. 반환값(원시 타입)/구간 로그 노출.
  테스트가 아닌 진입점 디버깅용 - 어서션 실패가 성패를 가르지 않는다
