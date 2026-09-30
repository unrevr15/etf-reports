# 아침 정시 트리거 설정 (Google Apps Script) — 약 5분

GitHub 예약 실행이 2026-08-26부터 전역 장애로 수 시간 늦어져(미해결), 외부에서 07:05 KST에 직접 깨운다.
새 계정 없이 기존 구글 계정으로 한다. 워크플로 쪽 수신부(`morning` 입력)는 이미 배포돼 있다.

## 1. GitHub 토큰 발급 (이 저장소, Actions 권한만)
1. github.com → 오른쪽 위 프로필 → **Settings** → 맨 아래 **Developer settings**
2. **Personal access tokens → Fine-grained tokens → Generate new token**
3. Token name: `etf-morning-trigger`, Expiration: 1년(가장 긴 것)
4. Repository access: **Only select repositories → unrevr15/etf-reports**
5. Permissions → Repository permissions → **Actions: Read and write** (나머지는 건드리지 않음)
6. Generate → 토큰 문자열 복사 (한 번만 보임)

## 2. Apps Script 만들기
1. https://script.google.com → **새 프로젝트**
2. 편집기의 기본 코드를 지우고 `cloud/morning_trigger.gs` 내용을 통째로 붙여넣기 → 저장(디스크 아이콘)
3. 왼쪽 톱니 **프로젝트 설정** → 아래 **스크립트 속성 → 속성 추가**
   - 속성: `GITHUB_TOKEN`  값: (1에서 복사한 토큰) → 저장

## 3. 1분 트리거 걸기
1. 왼쪽 시계 아이콘 **트리거** → **트리거 추가**
2. 실행할 함수: `tick` / 이벤트 소스: **시간 기반** / 유형: **분 단위 타이머** / 간격: **1분마다** → 저장
3. 권한 요청 창이 뜨면 허용 (외부 서비스 연결 = GitHub API 호출)

## 4. 확인
- 편집기에서 함수 `fireNow` 를 선택해 **실행** → 로그에 `HTTP 204` 가 찍히면 연결 성공.
  (07:10 이전이면 잡이 발송창까지 잡 안에서 기다리고, 이미 발송된 날이면 렌더만 하고 끝난다.)
- 다음 영업일 07:05~07:20 사이 GitHub Actions 탭에 `workflow_dispatch` 회차가 뜨고, 07:10~08:40 안에 방으로 발송된다.

## 동작 요약
- 매 1분 `tick` 실행 → 07:05~07:20 창에서 하루 1회 dispatch → 잡은 발송창(07:10)까지 대기 후 완결되는 즉시 발송, 늦어도 08:40.
- 주말은 스크립트가 건너뛰고, 공휴일은 워크플로가 스스로 종료.
- 기존 GitHub cron 은 그대로 둔다(늦게라도 뜨면 마커 확인 후 조용히 끝남 = 이중 안전망).
- 토큰 만료 1년 후 갱신 필요.
