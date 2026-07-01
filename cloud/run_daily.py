# -*- coding: utf-8 -*-
"""
클라우드 일일 오케스트레이터 (서울 리전 VM에서 cron 실행).
1) 일일 PDF 변화 리포트 (+ 금요일/주말 자동 주간 리포트)
2) 최근 5거래일 롤링 리포트
3) 포트폴리오 차트(PNG)
4) 당일 산출물을 Google Drive(rclone)로 업로드
모두 멱등(증분 스냅샷) — 하루에 여러 번 돌려도 안전(늦게 뜨는 TIGER를 채움).
"""
import os, sys, glob, subprocess, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.dirname(HERE)  # pdf_change.py / portfolio_chart.py 위치
sys.path.insert(0, APP)
os.chdir(APP)

import pdf_change as P
import portfolio_chart as PC

# 업로드 대상 rclone 리모트:폴더 (환경변수로 오버라이드 가능)
RCLONE_REMOTE = os.getenv("RCLONE_REMOTE", "gdrive:etf_reports")

def log(m): print(f"[{datetime.datetime.now():%H:%M:%S}] {m}", flush=True)

def main():
    today = datetime.date.today()
    if not P.is_trading_day(today):
        log(f"{today} 휴장일 — 종료"); return
    ymd = today.strftime("%Y%m%d")

    log("1) 일일 리포트")
    ctx = P.run(today)                  # pdf_change_YYYYMMDD.xlsx (+ 금/주말이면 주간 자동)
    log("1-b) 표 이미지 렌더 (+ 07:30/수동 실행이면 채널 전송)")
    try:
        png = P.render_report_image(ctx["groups"], ctx["today"], ctx["prev"], ctx["status_line"])
        send = os.getenv("TELEGRAM_SEND", "").lower() in ("1", "true", "yes")
        if png and send:
            cap = f"코스닥 액티브 ETF PDF 변화  {ctx['today']:%Y-%m-%d}\n{ctx['status_line']}"
            P.send_telegram(png, cap)
        elif not png:
            log("  변화 없음 — 이미지/전송 없음")
        else:
            log("  전송 시각 아님 — 이미지 렌더/저장만")
    except Exception as e:
        log(f"  이미지/텔레그램 실패: {type(e).__name__}: {e}")
    log("2) 최근 5거래일 롤링")
    try: P.rolling_report(today, 5)    # pdf_rolling5_YYYYMMDD.xlsx
    except Exception as e: log(f"  롤링 실패: {type(e).__name__}: {e}")
    log("3) 포트폴리오 차트")
    try: PC.generate(today)            # portfolio_YYYYMMDD.png
    except Exception as e: log(f"  차트 실패: {type(e).__name__}: {e}")

    # 4) 오늘 산출물 업로드
    targets = []
    for pat in (f"pdf_change_{ymd}.xlsx", f"pdf_change_{ymd}.png", f"pdf_weekly_{ymd}.xlsx",
                f"pdf_rolling5_{ymd}.xlsx", f"portfolio_{ymd}.png"):
        targets += glob.glob(os.path.join(APP, pat))
    if not targets:
        log("업로드할 파일 없음"); return
    log(f"4) Drive 업로드 {len(targets)}개 → {RCLONE_REMOTE}")
    for f in targets:
        try:
            subprocess.run(["rclone", "copy", f, RCLONE_REMOTE + "/", "--quiet"], check=True, timeout=120)
            log(f"  ✅ {os.path.basename(f)}")
        except Exception as e:
            log(f"  ❌ {os.path.basename(f)}: {e}")
    log("완료")

if __name__ == "__main__":
    main()
