# 클라우드 배포 가이드 — 코스닥 액티브 ETF PDF 변화 (서울 리전 + Google Drive)

매일 아침 자동으로 7개 코스닥 액티브 ETF의 PDF 변화/주간/롤링/차트를 만들어 **Google Drive**에 올린다. 내 PC를 켜둘 필요 없음.

## 구조 한눈에
```
서울 리전 VM (cron)  ──▶  pdf_change.py / portfolio_chart.py 실행
                          └─▶ rclone 으로 Google Drive(etf_reports/) 업로드
```
- **수집 인증 불필요**(운용사 공개 endpoint) → 코드에 비밀키 없음
- **유일한 비밀** = Drive 접근 토큰(rclone) → VM에만, 권한 600

---

## 0. 왜 서울 리전인가 (중요)
RISE·현대 등 일부 운용사 사이트는 **해외 IP를 차단**할 수 있다. 반드시 **한국(서울) 리전** 호스트를 쓴다.
- AWS Lightsail **ap-northeast-2(서울)** 최소 인스턴스($3.5~5/월) — 추천(쉬움)
- 또는 Naver Cloud / Vultr·Linode Seoul / 가비아·카페24 VPS

## 1. VM 만들기 (예: AWS Lightsail 서울)
1. Lightsail → Create instance → **Region: Seoul(ap-northeast-2)**
2. OS: **Ubuntu 22.04**, 플랜 최소($5)
3. 생성 후 SSH 접속.

## 2. 코드 올리기
로컬 PC(현재 폴더)에서 VM으로 복사. 필요한 파일: `pdf_change.py`, `portfolio_chart.py`, `cloud/` 폴더.
```bash
# 로컬에서 (scp 예)
scp -r "pdf_change.py" "portfolio_chart.py" cloud/  ubuntu@<VM_IP>:/home/ubuntu/etf/
```

## 3. 의존성 설치 (VM에서)
```bash
sudo apt update && sudo apt install -y python3-pip fonts-nanum rclone tzdata
sudo timedatectl set-timezone Asia/Seoul
cd /home/ubuntu/etf && pip3 install -r cloud/requirements.txt
```

## 4. ★ 지역차단 먼저 테스트 (배포 성패 판가름)
```bash
cd /home/ubuntu/etf && python3 cloud/test_connectivity.py
```
- **응답 7/7** 이면 진행. 일부 ❌면 그 운용사가 이 IP를 막는 것 → 다른 서울 호스트로 교체.

## 5. Google Drive 연결 (rclone, 1회)
```bash
rclone config
#  n) 새 리모트 → 이름: gdrive → 종류: drive → client_id/secret 비워도 됨
#  → "Use auto config? n" (헤드리스라 n) → 표시되는 URL을 로컬 브라우저로 열어 인증
#  → 토큰 붙여넣기 → team drive n → y(저장)
rclone mkdir gdrive:etf_reports          # 업로드 폴더 생성
rclone copy "pdf_change_20260625.xlsx" gdrive:etf_reports/   # 테스트 1개 업로드 확인
```
> 토큰은 `~/.config/rclone/rclone.conf` 에 저장됨. `chmod 600 ~/.config/rclone/rclone.conf`.

## 6. 수동 1회 실행 확인
```bash
mkdir -p state
python3 cloud/run_daily.py
#  → Drive etf_reports/ 에 오늘자 4종(일일·주간(금)·롤링·차트) 올라오면 성공
```

## 7. cron 등록 (매일 자동)
```bash
crontab cloud/crontab.txt    # 경로/사용자 맞게 수정 후
crontab -l                   # 확인
```
이제 매 영업일 07:07·08·09·16시에 자동 실행(멱등). 늦게 뜨는 TIGER·금요일 주간 자동 포함.

---

## 기존 파일 한 번에 Drive로 이전 (로컬 PC에서 1회)
로컬에도 rclone 설치 후 같은 gdrive 리모트 설정하고:
```powershell
rclone copy "C:\Users\정호준\Desktop\etf 총량" gdrive:etf_reports --include "pdf_*.xlsx" --include "portfolio_*.png"
```

## Docker로 돌리고 싶다면
```bash
docker build -t etf-daily -f cloud/Dockerfile .
docker run --rm -v $PWD/state:/app/state -v $PWD/rclone:/root/.config/rclone etf-daily
# rclone.conf 를 ./rclone/ 에 두고 마운트. cron 은 crontab.txt 의 docker 줄 참고.
```

## 보안 체크리스트
- [ ] 코드/깃에 비밀키 없음(수집은 공개 endpoint, 키 불필요)
- [ ] rclone.conf 권한 600, VM 외부 노출 금지
- [ ] Drive 폴더는 **본인만** 접근(공유 링크 만들지 말 것)
- [ ] VM 방화벽: 인바운드 22(SSH, 내 IP만) 외 전부 차단, 아웃바운드만 사용
- [ ] 매매·계좌 접근 코드 없음(읽기 전용) → 사고 시 피해 한정
