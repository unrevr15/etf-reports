# -*- coding: utf-8 -*-
"""최근 20영업일 분석용 데이터 수집 (일회성 프로젝트).
1) KRX 코스닥 ETF 일별(좌수·NAV·순자산) — 자금유출입/AUM용
2) 액티브 ETF PDF 일별(1CU 수량·종가) — 섹터/종목 순매수용
결과는 JSON 캐시로 저장(재실행 시 이미 받은 건 스킵).
"""
import os, sys, json, datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

HERE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.dirname(HERE)
sys.path.insert(0, APP)
import pdf_change as P

NDAYS = 16          # 스냅샷 16개 = 변동 15영업일(3주) — 주간리포트는 이 중 최근 5일로 비교
KRX_FILE = os.path.join(HERE, "krx_etf.json")
PDF_FILE = os.path.join(HERE, "pdf_raw.json")

# 코스닥 ETF 분류 (파생·혼합 제외)
EXC = ("레버리지", "인버스", "선물", "채권", "커버드콜", "혼합", "숏")

def trading_days(n=NDAYS, asof=None):
    d = asof or datetime.date.today()
    while not P.is_trading_day(d):
        d = P.prev_trading_day(d)
    out = [d]
    while len(out) < n:
        d = P.prev_trading_day(d)
        out.append(d)
    return sorted(out)          # 과거→현재

def load(path):
    try:
        with open(path, encoding="utf-8") as f: return json.load(f)
    except (OSError, ValueError): return {}

def save(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)

def collect_krx(days):
    """{ymd: {code: {...}}} — 코스닥 ETF만."""
    cache = load(KRX_FILE)
    key = P._krx_key()
    todo = [d for d in days if d.strftime("%Y%m%d") not in cache]
    print(f"[KRX] 총 {len(days)}일 중 신규 {len(todo)}일 수집")
    for d in todo:
        ymd = d.strftime("%Y%m%d")
        items = P._krx_get("http://data-dbg.krx.co.kr/svc/apis/etp/etf_bydd_trd",
                           {"AUTH_KEY": key, "basDd": ymd}).get("OutBlock_1", [])
        day = {}
        from universe import KOSDAQ_ACTIVE_CODES
        TRACKED = {e.get("krx") for e in P.ETFS} | set(KOSDAQ_ACTIVE_CODES)
        for it in items:
            nm = it.get("ISU_NM", "")
            code = it.get("ISU_CD", "")
            keep = (code in TRACKED) or ("코스닥" in nm and not any(x in nm for x in EXC))
            if not keep: continue
            def num(k):
                try: return float(str(it.get(k, "0")).replace(",", "") or 0)
                except ValueError: return 0.0
            day[it["ISU_CD"]] = {"nm": nm, "shrs": num("LIST_SHRS"), "nav": num("NAV"),
                                 "close": num("TDD_CLSPRC"), "netasset": num("INVSTASST_NETASST_TOTAMT"),
                                 "type": "액티브" if "액티브" in nm else "패시브"}
        if day:
            cache[ymd] = day
            print(f"  {ymd}: {len(day)}종목")
    save(KRX_FILE, cache)
    return cache

def collect_pdf(days):
    """{etf_id: {ymd: {종목: [수량, 종가]}}} — 추적 중인 액티브 7종."""
    cache = load(PDF_FILE)
    from universe import ACTIVE_PDF
    jobs = []
    for e in ACTIVE_PDF:
        cache.setdefault(e["id"], {})
        for d in days:
            ymd = d.strftime("%Y%m%d")
            if ymd not in cache[e["id"]]:
                jobs.append((e, ymd))
    print(f"[PDF] 신규 {len(jobs)}건 수집 ({len(ACTIVE_PDF)}종 x {len(days)}일)")
    if not jobs: return cache

    def one(job):
        e, ymd = job
        try:
            m, act = P._fetch(e, ymd)
            return (e["id"], ymd, m if (m and str(act) == ymd) else {})
        except Exception:
            return (e["id"], ymd, {})

    done = 0
    with ThreadPoolExecutor(max_workers=6) as ex:
        futs = [ex.submit(one, j) for j in jobs]
        for f in as_completed(futs):
            eid, ymd, m = f.result()
            cache[eid][ymd] = m
            done += 1
            if done % 20 == 0:
                print(f"  {done}/{len(jobs)}...", flush=True)
                save(PDF_FILE, cache)
    save(PDF_FILE, cache)
    ok = sum(1 for eid in cache for y in cache[eid] if cache[eid][y])
    print(f"[PDF] 완료 — 유효 스냅샷 {ok}건")
    return cache

if __name__ == "__main__":
    days = trading_days()
    print(f"대상 영업일 {len(days)}일: {days[0]} ~ {days[-1]}")
    collect_krx(days)
    collect_pdf(days)
    print("수집 완료")
