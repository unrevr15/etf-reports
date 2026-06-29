#!/usr/bin/env bash
# 우분투 VM에서 1회 실행 — 의존성·한글폰트·rclone·타임존 설치 + 연결 테스트.
# 사용:  bash bootstrap.sh
set -e
echo "== 1) 패키지 설치 =="
sudo apt-get update -y
sudo DEBIAN_FRONTEND=noninteractive apt-get install -y \
    python3-pip fonts-nanum rclone tzdata ca-certificates
sudo timedatectl set-timezone Asia/Seoul || true

echo "== 2) 파이썬 라이브러리 =="
pip3 install --user -r cloud/requirements.txt

echo "== 3) 상태 폴더 =="
mkdir -p state

echo "== 4) 지역차단 연결 테스트 (7/7 이어야 정상) =="
python3 cloud/test_connectivity.py

echo
echo "다음 단계:"
echo "  1) rclone config  → Google Drive(gdrive) 연결, 그다음 'rclone mkdir gdrive:etf_reports'"
echo "  2) python3 cloud/run_daily.py  로 1회 수동 실행 확인"
echo "  3) crontab cloud/crontab.txt  로 매일 자동 등록"
