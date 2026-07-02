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
    log("1-b) 이미지 렌더 + 채널 전송 (07:10/수동, 각 채널 하루 1회)")
    send = os.getenv("TELEGRAM_SEND", "").lower() in ("1", "true", "yes")
    cap = f"코스닥 액티브 ETF PDF 변화  {ctx['today']:%Y-%m-%d}\n{ctx['status_line']}"
    def deliver(label, png, tok, chat, marker_name):
        try:
            if not png: log(f"  [{label}] 변화 없음 — 스킵"); return
            if not (tok and chat): log(f"  [{label}] 대상 미설정 — 스킵"); return
            if not send: log(f"  [{label}] 전송 시각 아님 — 렌더만"); return
            m = os.path.join(APP, "reports", marker_name)
            if os.path.exists(m): log(f"  [{label}] 오늘 이미 발송 — 스킵"); return
            if P.send_telegram(png, cap, token=tok, chat=chat):
                os.makedirs(os.path.dirname(m), exist_ok=True); open(m, "w").close()
                log(f"  [{label}] 전송 완료 — 잠금 생성")
        except Exception as e:
            log(f"  [{label}] 실패: {type(e).__name__}: {e}")
    # 채널1: 기존(색상·가로) @pefscreener
    deliver("채널1", P.render_report_image(ctx["groups"], ctx["today"], ctx["prev"], ctx["status_line"]),
            os.getenv("TELEGRAM_TOKEN"), os.getenv("TELEGRAM_CHAT_ID"), f".tg_sent_{ymd}")
    # 채널2: 신규(무채색·모바일 세로)
    deliver("채널2", P.render_report_image_mobile(ctx["groups"], ctx["today"], ctx["prev"], ctx["status_line"]),
            os.getenv("TELEGRAM_TOKEN2"), os.getenv("TELEGRAM_CHAT_ID2"), f".tg_sent2_{ymd}")
    log("2) 최근 5거래일 롤링")
    try: P.rolling_report(today, 5)    # pdf_rolling5_YYYYMMDD.xlsx
    except Exception as e: log(f"  롤링 실패: {type(e).__name__}: {e}")
    log("3) 포트폴리오 차트")
    try: PC.generate(today)            # portfolio_YYYYMMDD.png
    except Exception as e: log(f"  차트 실패: {type(e).__name__}: {e}")

    # 4) 오늘 산출물 업로드
    targets = []
    for pat in (f"pdf_change_{ymd}.xlsx", f"pdf_change_{ymd}.png", f"pdf_change_m_{ymd}.png",
                f"pdf_weekly_{ymd}.xlsx", f"pdf_rolling5_{ymd}.xlsx", f"portfolio_{ymd}.png"):
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
