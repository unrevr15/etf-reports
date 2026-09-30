// ── ETF 리포트 아침 트리거 (Google Apps Script) ──────────────────────────────
// 왜 필요한가: GitHub Actions 예약 실행(cron)이 2026-08-26부터 GitHub 전역 장애로
//   수 시간 지연되거나 아예 안 뜬다(미해결). 확인된 유일한 우회책은 외부에서
//   workflow_dispatch API로 직접 깨우는 것이며, 이 경우 수 초 안에 시작된다.
// 동작: 매 1분 트리거로 돌면서 KST 07:05~07:20 창 안에서 하루 1회
//   unrevr15/etf-reports 의 daily.yml 을 morning=true 로 실행한다.
//   주말은 건너뛰고, 공휴일은 워크플로 자체가 판단해 바로 종료한다.
// 준비물: 스크립트 속성 GITHUB_TOKEN (fine-grained PAT, 이 저장소만, Actions: Read and write)
//
const REPO     = 'unrevr15/etf-reports';
const WORKFLOW = 'daily.yml';
const WIN_FROM = '07:05';   // KST
const WIN_TO   = '07:20';   // 이 창 안에서 첫 tick 이 발사(분 트리거는 가끔 한 분을 건너뜀)

function tick() {
  const tz   = 'Asia/Seoul';
  const now  = new Date();
  const hhmm = Utilities.formatDate(now, tz, 'HH:mm');
  const ymd  = Utilities.formatDate(now, tz, 'yyyyMMdd');
  const dow  = Utilities.formatDate(now, tz, 'u');          // 1=월 … 7=일
  if (hhmm < WIN_FROM || hhmm > WIN_TO) return;
  if (dow === '6' || dow === '7') return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAST_FIRED') === ymd) return;      // 하루 1회
  const token = props.getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('스크립트 속성 GITHUB_TOKEN 이 없습니다');
  const res = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + REPO + '/actions/workflows/' + WORKFLOW + '/dispatches', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
      payload: JSON.stringify({ ref: 'main', inputs: { morning: 'true' } }),
      muteHttpExceptions: true,
    });
  const code = res.getResponseCode();
  if (code === 204) { props.setProperty('LAST_FIRED', ymd); console.log('dispatch OK ' + ymd + ' ' + hhmm); }
  else throw new Error('dispatch 실패 HTTP ' + code + ' ' + res.getContentText().slice(0, 200));
}

// 수동 점검용: 지금 즉시 한 번 깨운다(창·요일·1회 제한 무시). 방 발송은 07:10 이후에만 일어나므로
// 07:10 이전에 누르면 잡이 발송창까지 잡 안에서 기다린다.
function fireNow() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_TOKEN');
  const res = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + REPO + '/actions/workflows/' + WORKFLOW + '/dispatches', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
      payload: JSON.stringify({ ref: 'main', inputs: { morning: 'true' } }),
      muteHttpExceptions: true,
    });
  console.log('HTTP ' + res.getResponseCode() + ' ' + res.getContentText());
}
