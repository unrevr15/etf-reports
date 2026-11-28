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
    import analyze   # 그룹 스냅샷(일일 리포트) 종목도 포함
    need |= {nm for k, d in analyze.SNAP.items() for y in d if y in keep for nm in (d[y] or {})}
    px = {y: {k: v for k, v in m.items() if k in need} for y, m in px.items() if y in keep}
    with open(pxf, "w", encoding="utf-8") as fp: json.dump(px, fp, ensure_ascii=False)
    log(f"캐시 정리, 보관 {len(keep)}영업일, 종가 {len(need)}종목")

def _top(rows, window):
    """window(일자 집합) 안의 종목별 순매수(억) 1위, 순매도 1위."""
    from collections import defaultdict
    t = defaultdict(float)
    for r in rows:
        if r["일자"] in window: t[r["종목"]] += r["순매수금액(원)"] / 1e8
    if not t: return ""
    b = max(t, key=t.get); s_ = min(t, key=t.get)
    import weekly_style as S
    parts = []
    if t[b] > 0: parts.append(f"순매수 {b} {S.eok(t[b])}억")     # 11-28: 순매수가 없는 주에 같은 종목을 양쪽에 쓰던 것 고침
    if t[s_] < 0: parts.append(f"순매도 {s_} {S.eok(t[s_])}억")
    return " / ".join(parts)

def make_images(out_dir, tag_override=None, log_fn=None):
    """그림과 캡션 만들기. 반환: [([그림 경로...], 캡션), ...]. 전송은 하지 않는다."""
    import importlib
    import analyze; importlib.reload(analyze)
    import weekly_style as S
    from themes import ORDER
    from universe import active_groups
    lg = log_fn or log
    tag = tag_override or analyze.DAYS[-1]
    F, fdays = analyze.flows_by_category()
    i0f = max(0, len(fdays) - CMP_DAYS)
    wk = {c: sum(F[c]["flow"][i0f:]) for c in F}
    # 캡션은 첫 줄 "제목 기간" 뿐(11-27 2판). 순유입, 순유출 값은 그림 표에 전부 있어 따로 쓰지 않는다.
    cap1 = f"주간 자금 유출입 {S.span(fdays[i0f], fdays[-1])}"
    msgs = [([S.chart_flows(F, fdays, os.path.join(out_dir, f"주간_자금유출입_{tag}.png"), cmp_days=CMP_DAYS)], cap1)]
    for gk, lab, es in active_groups():
        rows, skipped = analyze.group_net_buys(es)
        i0 = max(0, len(analyze.DAYS) - 1 - CMP_DAYS)
        # 보유비중 비교는 주초와 현재 PDF가 '둘 다' 있는 ETF만, 한쪽만 있으면 비중이 0에서 출발하는 착시가 생긴다
        es_cmp = [e for e in es if analyze._snap_of(e).get(analyze.DAYS[i0]) and analyze._snap_of(e).get(analyze.DAYS[-1])]
        miss = sorted({s_[0] for s_ in skipped if s_[2] == "PDF없음"})
        lg(f"  [{lab}] 대상 {len(es)}종, 비중 비교 {len(es_cmp)}종" + (f", PDF 결측 {len(miss)}종 {miss}" if miss else ""))
        # 앨범 장 수: 보유비중 0~1장 + 순매수 표 n장 + 추이 1장 (둘째 장부터 제목에 (k/N))
        n_w = 1 if es_cmp else 0
        n_nb = S.netbuy_pages(rows) if rows else 0
        total = n_w + n_nb + (1 if rows else 0)
        pics = []; k = 1
        if es_cmp:
            kind = "theme" if gk == "kosdaq" else "stock"
            fn = f"주간_{gk}_{'테마비중' if kind == 'theme' else '종목비중'}_{tag}.png"
            pics.append(S.chart_weights(analyze.group_holdings(es_cmp, kind), rows, os.path.join(out_dir, fn), lab,
                                        cmp_days=CMP_DAYS, key=kind, order=ORDER if kind == "theme" else None, part=(1, total)))
            k += 1
        else:
            lg(f"  [{lab}] 주초 PDF 없음, 보유비중 변화 생략(다음 주부터)")
        if rows:
            pics += S.chart_netbuy(rows, os.path.join(out_dir, f"주간_{gk}_종목별_{tag}.png"), lab, part_start=k, part_total=total)
            pics.append(S.chart_trend(rows, os.path.join(out_dir, f"주간_{gk}_종목별추이_{tag}.png"), lab, part=(total, total)))
        win = set(analyze.DAYS[i0 + 1:])
        cap = f"주간 {S.clean(lab)} ETF {S.span(analyze.DAYS[i0], analyze.DAYS[-1])}"
        t = _top(rows, win)
        if t: cap += "\n\n" + t      # 그림에 없는 사실(5영업일 합 1위)만 셋째 줄에 남김
        msgs.append((pics, cap))
    return msgs

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

    # 2) 차트 생성, 11-27 채널 공통 틀 A안: weekly_style.py (폭 1080, 제목 줄, 팔레트 고정)
    log("3) 차트 생성")
    msgs = make_images(HERE)

    # 3) 전송, 대상마다 봇 토큰이 다름(채널1=봇1, 채널2·DM=봇2). 메시지마다 사진 묶음(앨범) 1개.
    T1, T2 = os.getenv("TELEGRAM_TOKEN"), os.getenv("TELEGRAM_TOKEN2")
    targets = []                                   # [(라벨, 토큰, chat_id)]
    if os.getenv("WEEKLY_TO_CH1", "1") == "1" and T1 and os.getenv("TELEGRAM_CHAT_ID"):
        targets.append(("채널1", T1, os.getenv("TELEGRAM_CHAT_ID")))          # @pefscreener
    if os.getenv("WEEKLY_TO_CH2", "0") == "1" and T2 and os.getenv("TELEGRAM_CHAT_ID2"):
        targets.append(("채널2", T2, os.getenv("TELEGRAM_CHAT_ID2")))         # 수급방
    if os.getenv("WEEKLY_TO_DM", "1") == "1" and T2 and os.getenv("TELEGRAM_DM_ID"):
        targets.append(("DM", T2, os.getenv("TELEGRAM_DM_ID")))
    if not targets:
        log("전송 대상 미설정, 파일만 생성"); return
    for lab, tk, ch in targets:
        log(f"  → {lab} 전송 (메시지 {len(msgs)}개)")
        for pics, cap in msgs:
            if pics: P.send_telegram(pics, cap, token=tk, chat=ch)

    # 4) 캐시 슬림화
    prune()
    log("주간 리포트 완료")

if __name__ == "__main__":
    main()
