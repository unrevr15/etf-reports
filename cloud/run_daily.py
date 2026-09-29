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

DEGRADE_AFTER = 3   # 연속 이 거래일 수만큼 결측이면 '장기 파손'으로 보고 완결 판정에서 뺀다
MAX_DEGRADED = 2    # 이보다 많이 깨졌으면 자동 제외하지 않는다(껍데기 리포트가 나가는 것 방지)

def degraded_set(ctx):
    """직전 DEGRADE_AFTER 거래일 내내 유효 캡처가 없는 ETF 이름 집합.
    RISE처럼 한 곳이 오래 파손되면 매일 마감(8:40)까지 기다리게 되어 발송이 9시 반으로 밀린다.
    스냅샷만 보고 판정하므로 상태파일이 필요 없고, 복구되면 자동으로 집합에서 빠진다.
    오늘은 세지 않는다(아직 게시 전일 수 있으므로)."""
    try:
        snap = P.load_snap()
        days = []
        d = ctx["today"]
        for _ in range(DEGRADE_AFTER):
            d = P.prev_trading_day(d)
            days.append(d.strftime("%Y%m%d"))
        out = set()
        for e in P.ETFS:
            kk = snap.get(f"{e['am']}:{e['id']}", {})
            if all(not kk.get(x) for x in days):
                out.add(e["name"])
        return out if len(out) <= MAX_DEGRADED else set()
    except Exception as ex:
        log(f"  [열화] 판정 실패(무시): {type(ex).__name__}")
        return set()

def main():
    today = datetime.date.today()
    if not P.is_trading_day(today):
        log(f"{today} 휴장일 — 종료"); return
    ymd = today.strftime("%Y%m%d")

    log("1) 일일 리포트")
    ctx = P.run(today)                  # pdf_change_YYYYMMDD.xlsx (+ 금/주말이면 주간 자동)
    # ── 대시보드 업로드 ──────────────────────────────────
    # 리포트가 만들어진 직후, 발송 여부와 무관하게 올린다.
    # (발송은 완결 판정을 기다리지만 대시보드는 부분치라도 보이는 편이 낫다 —
    #  같은 날 다시 실행되면 덮어쓴다)
    try:
        sys.path.insert(0, HERE)
        import dashboard_push as DP
        # 경로를 짐작하지 않는다 — pdf_change 가 방금 쓴 파일을 ctx 로 받아서 쓴다.
        # (예전엔 없는 파일 이름을 넣어 두어 조용히 0건이었다)
        DP.push_from_csv(ctx.get("csv_path"), today.strftime("%Y-%m-%d"))
    except Exception as e:
        log(f"  대시보드 업로드 건너뜀: {e}")

    log("1-b) 이미지 렌더 + 채널 전송 (아침 7:10~8:40, 완결되면 즉시·마감시 되는대로, 하루 1회)")
    send_win = os.getenv("TELEGRAM_SEND", "").lower() in ("1", "true", "yes")   # 발송 후보 실행인가
    now_t = datetime.datetime.now().time()            # TZ=Asia/Seoul (워크플로 env)
    earliest = now_t >= datetime.time(7, 10)          # 7:10 이전엔 발송 안 함
    deadline = now_t >= datetime.time(8, 40)          # 8:40 넘으면 더 안 기다리고 되는대로 발송(운용사 늦게 게시 대비)
    hard_stop = now_t >= datetime.time(11, 0)         # 11시 넘은 회차는 '아침 리포트'가 아니다 → 발송 금지
    # 오후 게시가 정상인 ETF. 부분문자열이 아니라 이름 전체로 맞춘다
    # (부분매칭이면 이름에 TIGER가 든 ETF가 나중에 추가될 때 조용히 빠진다).
    LATE = ("TIGER 기술이전바이오액티브",)
    deg = degraded_set(ctx)                           # 장기 파손 ETF → 완결 판정에서 자동 제외
    pend_raw = [g["etf"] for g in ctx["groups"] if g["state"] == "pending" and g["etf"] not in LATE]
    pend_now = [n for n in pend_raw if n not in deg]
    deg_hit = [n for n in pend_raw if n in deg]       # 실제로 제외된 것만 로그에 남긴다
    krx_ok = ctx.get("krx_ok", True)                    # KRX 실패면 금액·순자산·현금 통째 빈칸 → 발송 보류
    complete = (len(pend_now) == 0) and krx_ok
    send = send_win and earliest and (not hard_stop) and (complete or deadline)
    if deg_hit:
        log(f"  [열화] 연속 {DEGRADE_AFTER}거래일 결측 → 완결 판정 제외: {', '.join(sorted(deg_hit))}")
    if send_win and earliest and hard_stop:
        log("  발송 금지 — 11시 넘음(아침 리포트 아님). 파일만 갱신.")
    elif send_win and earliest and not send:
        why = list(dict.fromkeys([p.split()[0] for p in pend_now] + ([] if krx_ok else ["KRX(금액)"])))
        log(f"  발송 보류 — 대기: {', '.join(why)} (마감 전 → 다음 회차 재시도)")
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
