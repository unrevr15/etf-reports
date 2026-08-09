# -*- coding: utf-8 -*-
"""주간 리포트 러너 (클라우드에서 토요일 실행).
  1) 새 영업일만 증분 수집(캐시가 레포에 커밋돼 있어 매주 5일치만 받음 → rate limit 회피)
  2) ①자금유출입·AUM + ②테마 보유비중 변화(그주 5영업일) + ③종목별 순매수 생성
  3) 텔레그램 전송 (WEEKLY_CHAT_ID / 없으면 DM)
  4) 캐시 슬림화(오래된 날짜·불필요 종목 제거)
"""
import os, sys, json, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.dirname(HERE)
sys.path.insert(0, APP); sys.path.insert(0, HERE)
import pdf_change as P

KEEP_DAYS = 30          # 캐시에 보관할 최대 영업일 수
CMP_DAYS = 5            # 덤벨 비교 구간(그 주 5영업일)

def log(m): print(f"[{datetime.datetime.now():%H:%M:%S}] {m}", flush=True)

def prune():
    """캐시 슬림화: 오래된 날짜 제거 + 종가는 보유종목만 남김."""
    from collect import trading_days, KRX_FILE, PDF_FILE, load, save
    keep = {d.strftime("%Y%m%d") for d in trading_days(KEEP_DAYS)}
    for f in (KRX_FILE, PDF_FILE):
        c = load(f)
        if not c: continue
        if f == KRX_FILE:
            c = {k: v for k, v in c.items() if k in keep}
        else:
            c = {eid: {y: v for y, v in d.items() if y in keep} for eid, d in c.items()}
        save(f, c)
    # 종가: 보유 종목만 (전체 1821종목 → 실제 필요분)
    pxf = os.path.join(HERE, "krx_px.json")
    try:
        with open(pxf, encoding="utf-8") as fp: px = json.load(fp)
    except (OSError, ValueError):
        return
    pdf = load(PDF_FILE)
    need = {nm for eid in pdf for y in pdf[eid] for nm in (pdf[eid][y] or {})}
    px = {y: {k: v for k, v in m.items() if k in need} for y, m in px.items() if y in keep}
    with open(pxf, "w", encoding="utf-8") as fp: json.dump(px, fp, ensure_ascii=False)
    log(f"캐시 정리 — 보관 {len(keep)}영업일, 종가 {len(need)}종목")

def main():
    # 1) 증분 수집
    log("1) 데이터 증분 수집")
    from collect import trading_days, collect_krx, collect_pdf
    days = trading_days()
    collect_krx(days); collect_pdf(days)
    import collect_px; collect_px.main()
    log("2) 섹터 매핑 갱신")
    import subprocess
    subprocess.run([sys.executable, os.path.join(HERE, "sectors.py")], cwd=HERE)

    # 2) 차트 생성
    log("3) 차트 생성")
    import importlib
    import analyze; importlib.reload(analyze)
    import charts; importlib.reload(charts)
    rows, skipped = analyze.net_buys()
    agg, detail = analyze.flows_aum()
    note = charts.coverage_note(rows, skipped)
    from themes import ORDER
    tag = analyze.DAYS[-1]
    p1 = charts.chart1(agg, os.path.join(HERE, f"주간_자금유출입_{tag}.png"))
    p2 = charts.chart2(rows, os.path.join(HERE, f"주간_테마비중_{tag}.png"), note,
                       order=ORDER, cmp_days=CMP_DAYS)
    p3 = charts.chart3(rows, os.path.join(HERE, f"주간_종목별_{tag}.png"), note=note)

    # 3) 전송
    tok = os.getenv("TELEGRAM_TOKEN2") or os.getenv("TELEGRAM_TOKEN")
    chat = os.getenv("WEEKLY_CHAT_ID") or os.getenv("TELEGRAM_DM_ID")
    if not (tok and chat):
        log("전송 대상 미설정 — 파일만 생성"); return
    d0 = analyze.DAYS[max(0, len(analyze.DAYS) - 1 - CMP_DAYS)]
    cap = (f"[주간] 코스닥 액티브 ETF 테마 비중변화\n"
           f"{d0[:4]}-{d0[4:6]}/{d0[6:]} → {tag[4:6]}/{tag[6:]} ({CMP_DAYS}영업일)")
    for i, p in enumerate([p for p in (p1, p2, p3) if p]):
        P.send_telegram(p, cap if i == 0 else "", token=tok, chat=chat)

    # 4) 캐시 슬림화
    prune()
    log("주간 리포트 완료")

if __name__ == "__main__":
    main()
