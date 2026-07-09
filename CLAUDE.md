# AGENTS.md / CLAUDE.md - UnivFIlter

(시그니처 스키마·src 편집 주의는 상위 VBA_Projects/CLAUDE.md 참조. 아키텍처·테스트 상세는 src/CLAUDE.md)

대상 워크북: UnivFilter(v2.6).xlsm (xlsb에서 xlsm으로 마이그레이션됨)

<!-- vbaflow:managed:start -->
## 개발 루프 (_VbaFlow)

빌드/테스트 루프는 `vbaflow`(C# .NET 8 단일 exe)로 수행한다. PATH 미등록 시 exe 경로:
`..\_VbaFlow\src\VbaFlow\bin\Release\net8.0-windows\win-x64\publish\vbaflow.exe`
온보딩/드리프트 복구: `vbaflow init` (검사만: `vbaflow init --check`).
프로젝트 고유 내용은 이 관리 블록 밖에 작성한다 (init이 블록 내부를 덮어씀).

1. 편집: src\*.bas/.cls/.frm 수정
2. 세션: vbaflow session start (도구 소유 Excel, 멱등)
3. 주입: vbaflow push (src -> 워크북 in-place 주입, --prune/--force 지원)
4. 실행: vbaflow test (Test_* 전체 실행 + LogShim jsonl 판독, --timeout <초>)
5. 수정: fail이면 원인 고치고 1로, green이면 git commit

- 모든 명령 --json 지원: {command, ok, exitCode, hint, errors[], payload}
- exit code: 성공 0 / 환경 오류 3 (세션/Excel/COM - session start 후 재시도) / 그 외 2
- 테스트는 Public Sub Test_*() (인자 없음), 단언은 LogShim.AssertEq (Debug.Assert 금지)
- src\LogShim.bas는 벤더링 사본 - 직접 수정 금지, 갱신은 vbaflow init
- 배포 아티팩트: vbaflow dist --out <경로> (클린 사본 빌드 - 필터/스탬프/무해화/선택 부트스트랩).
  push 선행 필수(마지막 push == src 미동기 시 거부). 절차/실패 플레이북은 SKILL.md
- pull은 동기화 안 된 src 편집이 있으면 차단된다 - 기본은 push로 먼저 반영,
  워크북 쪽이 진실일 때만 pull --force로 덮어쓰기
- 세션 Excel에 직접 COM 접근 금지(GetActiveObject, New-Object -ComObject 등 - 조회/진단 포함).
  워크북 실제 내용 확인은 vbaflow diff --json (읽기 전용, --module <이름>은 단일 모듈 덤프)
- src 규약: UTF-8 BOM 없음 + CRLF + 끝 개행 1개. CP949 밖 문자(em dash U+2014, 이모지)는 push가 거부
- vbaflow 자체 결함 발견 시: vbaflow bugreport "<제목>" 으로 표준 리포트 생성 후 내용 채움
  (절차: .claude\skills\vbaflow\SKILL.md)
- 심화 절차/실패 플레이북: .claude\skills\vbaflow\SKILL.md (vbaflow init이 관리 - 직접 수정 금지)
<!-- vbaflow:managed:end -->
