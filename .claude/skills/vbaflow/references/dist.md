# vbaflow dist - 배포 아티팩트 빌드

`vbaflow dist --out <경로> --json`은 배포용 클린 사본(`.xlsm` 또는 `.xlam`)을 만든다. on-disk
워크북을 base로 복사 -> 자체 임시 Excel(창 없음)에서 (모듈 제외 -> 버전 스탬프 -> LogShim 무해화
-> 선택 주석제거 -> 선택 부트스트랩)을 적용 -> 원자 이동으로 산출. 세션 워크북은 건드리지 않는다
(lease 불필요).

## 계약

- **push 선행 필수**: dist는 "마지막 push == src"인 on-disk 워크북을 base로 삼는다. 미동기면
  거부한다(`ok:false`, hint="vbaflow push로 반영 후 재시도"). **순서: (편집 ->) push -> dist.**
- 형식은 `--out` 확장자로 판정: `.xlsm`(매크로 워크북) / `.xlam`(애드인 - 자동 IsAddin + xlam
  저장). `--out` 생략 시 `<root>\dist\<워크북명>.xlsm`. `dist\`는 init-관리 .gitignore로 무시된다
  (빌드 아티팩트 = 비커밋).
- 제외 지정: 1차는 모듈 **헤더**의 `' @dist exclude`(본문에 적으면 무시된다 - 헤더 한정),
  보강은 toml `[dist] exclude` 글롭(`Test_*` 등). 제외는 모듈 단위만.
- 버전: toml `[dist] version`(semver). git rev는 자동 부가 - **git repo가 아니면 rev=`nogit`**
  (도구 자신의 rev를 스탬프하지 않는다). BUILD_* 상수 + CustomDocumentProperties 양쪽에 기록.
- 옵션: `--strip-comments`(주석 제거, 문자열 내 `'` 보호), `--bootstrap <Module.Proc>`(주입 후
  워크시트 초기설정 실행 - kill 워치독으로 모달 격리).

## 실패 플레이북

- **sync 미동기(가장 흔함)**: hint대로 `vbaflow push` 후 재시도. **rollout 직후 특히 주의** -
  `rollout.ps1`의 `init --all`이 각 프로젝트 `src\LogShim.bas`를 재벤더링해 src 지문이 어긋난다.
  **rollout -> (프로젝트별) push -> dist** 순서를 지킨다(init은 순수 파일시스템이라 push 못 함).
- **dangling refs(컴파일 안전 실패)**: 제외 모듈의 심볼을 잔존 모듈이 참조. payload.danglingRefs의
  `모듈:라인`을 보고 제외 범위를 좁히거나 참조를 제거(Test 모듈은 서로만 참조해야 통째 제거 안전).
- **LogShim IS_RELEASE 부재(CliError)**: 배포 무해화 계약 위반 - 조용히 로깅 켜진 배포본을 막기
  위해 크게 실패한다. `assets\LogShim.bas`에 상수 추가 후 `rollout.ps1` 재벤더링.
- **부트스트랩 타임아웃(120s)**: 부트스트랩 프로시저가 MsgBox/모달로 행을 건 것. 임시 Excel은
  kill로 회수된다(세션 무해). 해당 프로시저를 비대화형·컴파일 클린으로 고친다.
- **문서/시트 모듈 제외 불가**: warning으로 무시된다(Remove 불가). 진짜 제외가 필요하면 설계 재고.
- **`doctor --kill-orphans` 금지(dist 실행 중)**: dist의 임시 Excel은 창이 없어 doctor 고아 스캔에
  잡히고, doctor는 mutex를 안 잡으므로 dist를 죽일 수 있다. dist 진행 중엔 `--kill-orphans`를 쓰지 말 것.
