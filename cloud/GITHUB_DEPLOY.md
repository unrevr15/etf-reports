# GitHub Actions 배포 (평생 무료, 서버 없음)

매일 아침 GitHub가 자동으로 스크립트를 돌려 7개 코스닥 액티브 ETF 리포트를 만들고
**Google Drive(etf_reports/)** 에 올린다. 내 PC·서버 불필요. 비용 0원.

> 해외 러너에서도 운용사 endpoint가 정상 동작함을 확인함(지역차단 없음).
> 수집에 비밀키 없음. **유일한 비밀 = Drive 토큰(rclone)** → GitHub Secret에만 저장.

---

## 1. GitHub 계정·저장소
1. https://github.com 가입(무료)
2. 우상단 **+ → New repository**
   - 이름: `etf-reports` (아무거나)
   - **Private** 선택(중요)
   - Create

## 2. 코드 올리기 (드래그&드롭, git 몰라도 됨)
1. 저장소 페이지 → **Add file → Upload files**
2. 로컬 폴더 `C:\Users\정호준\Desktop\etf 총량` 에서 아래를 끌어다 놓기:
   - `pdf_change.py`, `portfolio_chart.py`
   - `cloud/` 폴더 통째
   - `.github/` 폴더 통째 (← 워크플로. 숨김폴더라 안 보이면 주소창에 경로 직접 입력하거나 탐색기에서 숨김표시 켜기)
3. **Commit changes**
> `.github/workflows/daily.yml` 이 올라가야 자동 실행이 켜진다. 업로드 후 저장소 **Actions** 탭에 워크플로가 보이면 성공.

## 3. Google Drive 토큰 만들기 (로컬 PC에서 1회)
내 PC(윈도우)에서 rclone로 Drive를 1회 인증해 설정파일을 얻는다.
```powershell
# rclone 없으면: winget install Rclone.Rclone   (또는 rclone.org 다운로드)
rclone config
#  n) New remote → name: gdrive → storage: drive
#  client_id/secret 엔터(비움) → scope: 1(전체) 또는 drive
#  "Edit advanced config? n" → "Use auto config? y" → 브라우저로 내 구글 로그인·허용
#  team drive: n → 저장 y
rclone mkdir gdrive:etf_reports         # 업로드 폴더 생성
type %USERPROFILE%\.config\rclone\rclone.conf   # ← 이 내용 전체 복사
```
복사한 내용(예: `[gdrive]` 로 시작하는 블록 전체)을 다음 단계에 붙여넣는다.

## 4. GitHub Secret 등록
저장소 → **Settings → Secrets and variables → Actions → New repository secret**
- Name: `RCLONE_CONF`
- Secret: 3번에서 복사한 rclone.conf 전체 붙여넣기 → Add secret

(선택) 업로드 폴더명 바꾸려면 **Variables** 탭에 `RCLONE_REMOTE = gdrive:폴더명` 추가.

## 5. 동작 확인
- 저장소 **Actions** 탭 → "ETF PDF 변화 일일 자동" → **Run workflow**(수동 실행) 눌러 1회 테스트
- 초록 체크 뜨고 Drive `etf_reports/` 에 오늘자 파일(일일·롤링·차트, 금요일이면 주간) 올라오면 완료
- 이후 매 영업일 아침 자동 실행(KST 06·07·08·09·10·16시 시도, 멱등)

---

## 기존 파일 한 번에 Drive로 이전 (로컬 PC, 1회)
```powershell
rclone copy "C:\Users\정호준\Desktop\etf 총량" gdrive:etf_reports --include "pdf_*.xlsx" --include "portfolio_*.png"
```

## 참고
- GitHub Actions 무료 한도(비공개 2,000분/월)에 한참 못 미침(1회 ~2분).
- cron은 UTC 기준이며 부하 시 수 분 지연될 수 있음 — 여러 번 시도 + 멱등이라 괜찮음.
- 상태파일 불필요: 매 실행이 당일·전일 PDF를 새로 받아 비교(dateapi).
- 보안: 저장소 Private 유지, RCLONE_CONF는 Secret에만(코드/커밋에 절대 X).
