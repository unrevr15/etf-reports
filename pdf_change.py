# -*- coding: utf-8 -*-
"""
운용사 직접 당일 PDF → 전일대비 구성종목 수량 변화(PDF 변화) 계산. (액티브 ETF 전용)
두 가지 fetcher 모드:
  - dateapi : 날짜 파라미터 API. 임의 영업일 조회 가능 → 첫 실행부터 diff (예: KoAct/삼성액티브)
  - snapshot: 당일치만 주는 소스. 매 영업일 스냅샷 저장 → 2영업일째부터 self-diff (예: TIME)
스냅샷은 snapshots.json 에 {etf_key:{YYYYMMDD:{종목:수량}}} 누적.
curl 차단 사이트(PLUS/UNICORN/RISE/DB)는 헤드드 Selenium fetcher로 추가 예정(PC 켜둠).
"""
import sys, os, csv, json, requests
from datetime import datetime, date, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
SNAP_FILE = os.path.join(BASE, "snapshots.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36"

HOLIDAYS_2026 = {
    "2026-01-01","2026-02-16","2026-02-17","2026-02-18","2026-03-02","2026-05-01","2026-05-05",
    "2026-05-25","2026-06-03","2026-07-17","2026-08-17","2026-09-24","2026-09-25","2026-10-05",
    "2026-10-09","2026-12-25","2026-12-31",
    # 07-17 제헌절: 전 운용사 PDF 부재로 휴장 확인(2026-07-20 진단). 05-01 근로자의날도 KRX 휴장.
}
def is_trading_day(d): return d.weekday() < 5 and d.strftime("%Y-%m-%d") not in HOLIDAYS_2026
def prev_trading_day(d):
    p = d - timedelta(days=1)
    while not is_trading_day(p): p -= timedelta(days=1)
    return p
def next_trading_day(d):
    n = d + timedelta(days=1)
    while not is_trading_day(n): n += timedelta(days=1)
    return n

SKIP = ("현금","원화예금","원화 예금","예수금","설정현금","CASH","원화","달러","USD","선물","원화현금")
MIN_CHANGE = 3  # 이 미만(±1·±2주)의 수량확대/축소는 노이즈로 숨김. 신규편입·편출은 크기 무관 항상 표시.

# ── 현금(dry powder) 캐시: 페처가 스킵하는 현금 라인 평가금액을 (id,ymd)별 1CU 현금으로 누적 ──
_CASH_CACHE = {}
_CASHWORDS = ("현금", "예금", "예수금")   # 선물/달러는 현금 아님 → 제외
def _note_cash(key, ymd, nm, ev):
    if ev and any(w in nm for w in _CASHWORDS):
        _CASH_CACHE[(key, ymd)] = _CASH_CACHE.get((key, ymd), 0.0) + float(ev)

# ── 스냅샷 저장소 ──────────────────────────────────────────────
def load_snap():
    if os.path.exists(SNAP_FILE):
        try: return json.load(open(SNAP_FILE, encoding="utf-8"))
        except Exception: pass
    return {}
def save_snap(s):
    tmp = SNAP_FILE + ".tmp"
    json.dump(s, open(tmp, "w", encoding="utf-8"), ensure_ascii=False)
    os.replace(tmp, SNAP_FILE)

# ── fetcher: 삼성액티브(KoAct) 날짜 API → {종목:수량}, 기준일 ──
def fetch_samsungactive(prod_id, ymd):
    g = f"{ymd[:4]}.{ymd[4:6]}.{ymd[6:]}"
    r = requests.get(f"https://www.samsungactive.co.kr/api/v1/product/etf-pdf/{prod_id}.do",
                     params={"gijunYMD": g},
                     headers={"User-Agent": UA, "Referer": f"https://www.samsungactive.co.kr/etf/view.do?id={prod_id}"},
                     timeout=20)
    d = r.json().get("pdf", {})
    actual = str(d.get("gijunYMD", "")).replace(".", "")
    _CASH_CACHE.pop((prod_id, actual), None)   # 현금은 요청일이 아니라 실제 기준일로(재조회 시 누적 방지)
    out = {}
    for it in d.get("list", []):
        nm = (it.get("secNm") or "").strip()
        if not nm: continue
        try: ev = float(str(it.get("evalA", "0")).replace(",", ""))
        except ValueError: ev = 0.0
        if any(s in nm for s in SKIP):
            _note_cash(prod_id, actual, nm, ev); continue
        try: q = float(str(it.get("applyQ", "0")).replace(",", ""))
        except ValueError: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, (ev / q if q else None)]
    return out, actual

# ── fetcher: 타임폴리오(TIME) pdf_excel.php (pdfDate 날짜 파라미터) → {종목:수량}, 기준일 ──
def fetch_time(idx, ymd):
    import openpyxl, io
    pdfdate = f"{ymd[:4]}-{ymd[4:6]}-{ymd[6:]}"  # YYYY-MM-DD (대시 포함이어야 함)
    r = requests.get("https://timeetf.co.kr/pdf_excel.php",
                     params={"idx": idx, "pdfDate": pdfdate},
                     headers={"User-Agent": UA, "Referer": f"https://timeetf.co.kr/m11_view.php?idx={idx}"},
                     timeout=20)
    wb = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    out = {}
    if rows:
        hdr = [str(c).strip() if c else "" for c in rows[0]]
        try:
            i_nm, i_q = hdr.index("종목명"), hdr.index("수량")
        except ValueError:
            return {}, ""
        i_ev = next((j for j, c in enumerate(hdr) if "평가금액" in c), None)
        for row in rows[1:]:
            if not row or len(row) <= max(i_nm, i_q): continue
            nm = str(row[i_nm]).strip() if row[i_nm] else ""
            if not nm: continue
            try: ev = float(str(row[i_ev]).replace(",", "")) if i_ev is not None and row[i_ev] is not None else 0.0
            except (ValueError, TypeError): ev = 0.0
            if any(s in nm for s in SKIP):
                _note_cash(idx, ymd, nm, ev); continue
            try: q = float(str(row[i_q]).replace(",", ""))
            except (ValueError, TypeError): continue
            if nm in out: out[nm][0] += q
            else: out[nm] = [q, (ev / q if (q and ev) else None)]
    return out, (ymd if out else "")  # 파일에 기준일 셀 없음 → 요청한 날짜로

# ── fetcher: 한화(PLUS) — 사이트 API pdf/list (엑셀보다 안정적; 엑셀은 '날짜만 있고 빈' 상태로 옴) ──
def fetch_plus(n, ymd):
    # 사이트가 화면표시에 쓰는 API. content: {jmNm,amount(수량),ratio(비중%),wkdate(발행일)}. page 0-base.
    r = requests.get("https://www.plusetf.co.kr/api/v1/product/pdf/list",
                     params={"n": n, "d": ymd, "page": 0, "pageSize": 300},
                     headers={"User-Agent": UA, "Referer": f"https://www.plusetf.co.kr/product/detail?n={n}"},
                     timeout=20)
    out = {}; actual = ""
    try:
        content = r.json().get("content") or []
    except Exception:
        return {}, ""
    cash = []; ratio = {}
    for it in content:
        nm = str(it.get("jmNm") or "").strip()
        if not nm: continue
        actual = str(it.get("wkdate") or actual)
        try: q = float(it.get("amount") or 0)
        except (ValueError, TypeError): continue
        if any(s in nm for s in SKIP):
            cash.append((nm, q)); continue   # 현금라인 amount = 현금액(원)/CU — 실제 기준일(wkdate)로 아래에서 기록
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, None]   # PLUS는 평가금액 없음 → 종가 미상(KRX fallback)
        try: ratio[nm] = ratio.get(nm, 0.0) + float(it.get("ratio") or 0)
        except (ValueError, TypeError): pass
    # 코스닥이 아닌 종목(코스피)은 KRX 종가가 없어 금액·비중이 빈다(2026-10-29 PLUS AI반도체소부장 4/22종목).
    # PLUS는 종목별 비중(ratio, 소수 4자리)을 준다 → 종가를 아는 코스닥 종목으로 바스켓 총액을 구하고
    # 단가 = 비중 × 총액 ÷ 수량 으로 역산한다(코스닥 종목 검증 오차 0.002%). 네이버 종가는 NXT 시간외가 섞여 0.2~1.3% 어긋나 못 쓴다.
    if out and actual:
        kq = krx_kosdaq_prices()
        known = [(nm, v[0] * kq[nm]) for nm, v in out.items() if kq.get(nm) and ratio.get(nm)]
        rk = sum(ratio[nm] for nm, _ in known)
        if known and rk > 0:
            total = sum(ev for _, ev in known) / (rk / 100)
            for nm, v in out.items():
                if v[1] is None and nm not in kq and ratio.get(nm) and v[0]:
                    v[1] = round(ratio[nm] / 100 * total / v[0])
    _CASH_CACHE.pop((n, actual), None)
    for nm, q in cash: _note_cash(n, actual, nm, q)
    return out, actual

# ── fetcher: KB(RISE) 1CU 구성종목 ──────────────────────────────
# 2026-09 사이트 이전 대응.
#   구: riseetf.co.kr/prod/finder/productViewTabExcel3 (HTML-xls) → 사망.
#       지금은 Next.js SPA 껍데기(145KB)만 200으로 돌려준다(table 0개) → 2026-09-19부터 매일 빈값.
#   신: RISE가 kbam.co.kr(KB자산운용 통합)로 이전. 구성종목은 JSON API.
#       GET https://kbam.co.kr/api/products/etfs/{stid}/holdings?base_dt=YYYYMMDD
#       → {base_dt, holding_count, total_count, available_dates[],
#          items:[{rank,item_cd,item_nm,shares,unit_price,vamt,ratio}...]}
#       쿠키·세션 불필요. stid는 기존 그대로('44I0').
#   검증(2026-09-29): 구 페처 결과가 남아있는 20260916 스냅샷과 종목집합·수량 완전 일치(22종목).
#
# 주의 1) 미게시일·휴장일을 요청하면 404가 아니라 '최신 영업일'로 조용히 폴백한다.
#         (실측: 20260925·20260101·20261231 요청 → 전부 base_dt 20260929 반환)
#         → 요청한 ymd가 아니라 반드시 응답의 base_dt를 기준일로 반환할 것.
#           그래야 run()의 `if tm and ta == t_ymd` 판정이 '대기'로 잡아준다.
# 주의 2) JSON items는 상위 30종목 상한. total_count가 더 크면 같은 URL의 xlsx(전체)로 교체한다.
#         보강 실패 시 잘린 바스켓을 쓰면 나머지가 전부 '편출'로 오탐되므로 빈값을 반환해 '대기'로 넘긴다.
# 주의 3) 2026-10-06부터 kbam.co.kr이 GitHub 러너에서 접속 시간초과(한국에선 정상) → RISE가 리포트에서 통째로 빠졌다.
#         첫 접속 실패에서 이번 실행 동안은 건너뛰고(같은 결과에 회차마다 4분씩 쓰던 것) FunETF → WiseReport로 받는다.
#         둘 다 최신분만 주므로 아침엔 전일분 → ETFS에서 "lag":True.
_KBAM_DOWN = False
_RISE_KRX = {"44I0": "0000Z0", "44B9": "388420", "44L2": "0254H0", "44C7": "422420"}   # stid → KRX 단축코드(FunETF·WiseReport 폴백용)

def fetch_rise(stid, ymd):
    global _KBAM_DOWN
    out, actual = {}, ""
    if not _KBAM_DOWN:
        try:
            out, actual = _rise_direct(stid, ymd)
        except requests.exceptions.ConnectionError as ex:      # ConnectTimeout 포함
            _KBAM_DOWN = True
            print(f"  [RISE] kbam.co.kr 접속 실패({type(ex).__name__}) → 이번 실행에선 FunETF로", flush=True)
    if out:
        return out, actual          # 운용사 원본이 답하면 날짜가 늦어도 그대로(정직한 기준일)
    code = _RISE_KRX.get(stid, "")
    if not code:
        return {}, ""
    fm, fd = fetch_funetf(krx_isin(code), key=stid)
    if fm:
        print(f"  [RISE] FunETF 폴백 {len(fm)}종목 (기준일 {fd})", flush=True)
        return fm, fd
    return fetch_wisereport(code, ymd)

def _rise_direct(stid, ymd):
    hdr = {"User-Agent": UA, "Accept-Language": "ko-KR",
           "Referer": f"https://kbam.co.kr/products/{stid}"}
    url = f"https://kbam.co.kr/api/products/etfs/{stid}/holdings"
    r = requests.get(url, params={"base_dt": ymd},
                     headers=dict(hdr, Accept="application/json"), timeout=(5, 20))
    try:
        j = r.json()
    except Exception:
        return {}, ""
    actual = str(j.get("base_dt") or "").strip()   # 실제 기준일(폴백 감지용)
    _CASH_CACHE.pop((stid, actual), None)           # 같은 날짜 재조회(대기 폴링) 시 현금 누적 방지
    items = j.get("items") or []
    if not actual or not items:
        return {}, ""
    rows = [(str(it.get("item_nm") or "").strip(), it.get("shares"), it.get("vamt"))
            for it in items]
    try: total = int(j.get("total_count") or 0)
    except (ValueError, TypeError): total = 0
    if total > len(rows):          # 30종목 상한에 걸림 → 전체 엑셀로 교체
        try:
            import pandas as pd, io as _io
            x = requests.get(url, params={"download": "xlsx", "base_dt": actual},
                             headers=hdr, timeout=(5, 30))
            df = pd.read_excel(_io.BytesIO(x.content))   # [종목코드,종목명,수량(주),보유비중(%),평가금액(원)]
            c_nm = next(c for c in df.columns if "종목명" in str(c))
            c_q = next(c for c in df.columns if "수량" in str(c))
            c_ev = next((c for c in df.columns if "평가금액" in str(c)), None)
            rows = [(str(v[c_nm]).strip(), v[c_q], (v[c_ev] if c_ev is not None else None))
                    for _, v in df.iterrows()]
            if len(rows) < total:
                return {}, ""
        except Exception:
            return {}, ""
    out = {}
    for nm, q, ev in rows:
        if not nm or nm == "nan": continue
        try: ev = float(str(ev).replace(",", "")) if ev is not None else 0.0
        except (ValueError, TypeError): ev = 0.0
        if any(s in nm for s in SKIP):
            _note_cash(stid, actual, nm, ev); continue  # 원화예금 등 → 현금 캐시(요청일이 아니라 실제 기준일로)
        try: q = float(str(q).replace(",", ""))
        except (ValueError, TypeError): continue
        if not q: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, (ev / q if ev else None)]
    return out, (actual if out else "")

# ── fetcher: 미래에셋(TIGER) — GET 단발 + clamp 탐지 + 진단로그 + WiseReport 폴백 ──
#   2026-09-29 진단: 이 페처만 클라우드(GitHub Actions·미국 러너)에서 2026-08-18 배포 이후
#   단 한 번도 성공한 적이 없다(레포 snapshots.json 의 미래에셋 항목 0일). 로컬(한국)은 33종목 정상.
#   러너에서 두 요청이 1.4초에 끝난다(같은 러너의 RISE 145KB는 4.5초, 한국은 같은 요청 2.08초)
#   → 본문을 못 받고 CloudFront 엣지에서 잘린 것. 국가 차단은 아니고(해외 egress는 정상 응답)
#   러너 IP대역(Azure/GitHub ASN) 대상 WAF 거부가 유력. 기존 코드는 read_html 실패를 통째로
#   삼켜 '빈값'만 남겨 원인 추적이 불가능했다 → 실패 시 HTTP코드·본문 앞부분을 로그로 남긴다.
#
#   ★ clamp 주의: 이 endpoint는 아직 안 나온 날짜(미래 포함)를 요청하면 404가 아니라
#     '최신 바스켓'을 주면서 요청 날짜를 그대로 echo한다. 그대로 믿으면 묵은 바스켓이
#     오늘치로 저장된다(켐트로닉스 오보와 같은 유형). → 전영업일 응답과 본문이 완전히 같으면
#     clamp로 보고 빈값을 반환해 '대기'로 넘긴다. 종가가 매일 다르므로 본문 동일 = 같은 파일이다.
_TIGER_KRX = {"KR70168K0008": "0168K0", "KR70204S0006": "0204S0", "KR7471780007": "471780"}   # ISIN → KRX 단축코드(WiseReport 폴백용)

def _tiger_rows(text):
    """pdfListAjax 응답(<tr> 조각) → [[셀,...], ...]. 표가 아니면 [] (pandas·lxml 불필요)."""
    import re
    rows = []
    for tr in re.findall(r"<tr[^>]*>.*?</tr>", text, re.S | re.I):
        cells = [" ".join(re.sub(r"<[^>]+>", " ", c).split())
                 for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)]
        if len(cells) >= 4:
            rows.append(cells)
    return rows

_MIRAE_DOWN = False   # 러너가 403을 받으면 이번 실행 동안은 미래에셋 직접 호출을 건너뛴다(매번 같은 403, 시간만 듦)

def _tiger_raw(isin, ymd, tries=3):
    """운용사 직접 호출 → 응답 text. POST·세션워밍업 불필요(GET 한 번으로 동일 본문, 실측 확인)."""
    global _MIRAE_DOWN
    import time
    if _MIRAE_DOWN: return ""
    base = "https://investments.miraeasset.com/tigeretf/ko/product/search/detail"
    hdr = {"User-Agent": UA, "Accept": "text/html, */*; q=0.01",
           "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
           "X-Requested-With": "XMLHttpRequest",
           "Referer": f"{base}/index.do?ksdFund={isin}"}
    params = {"ksdFund": isin, "pageIndex": 1, "firstIndex": 0, "listCnt": 300,
              "fixDate": f"{ymd[:4]}.{ymd[4:6]}.{ymd[6:]}", "prfPrd": "Week01", "order": "SRD"}
    last = ""
    for i in range(tries):
        try:
            r = requests.get(f"{base}/pdfListAjax.ajax", params=params, headers=hdr, timeout=25)
            if r.status_code == 200 and _tiger_rows(r.text):
                return r.text
            if r.status_code == 403:                    # 러너 IP 차단(WAF) — 재시도해도 같다
                _MIRAE_DOWN = True
                print(f"  [TIGER] 미래에셋 HTTP 403(러너 차단) → 이번 실행에선 FunETF로", flush=True)
                return ""
            last = f"HTTP {r.status_code} len={len(r.text)} head={r.text[:140]!r}"
        except Exception as ex:
            last = f"{type(ex).__name__}: {str(ex)[:140]}"
        if i < tries - 1: time.sleep(2 * (i + 1))
    print(f"  [TIGER] 미래에셋 응답에 표 없음 → {last}", flush=True)   # 조용히 넘기지 않는다
    return ""

def _tiger_wisereport(krx):
    """폴백: WiseReport(네이버) CU 구성 → (행들, 실제기준일). 형제 프로젝트 import 없이 단독 동작."""
    import re, json
    if not krx: return [], ""
    try:
        r = requests.get("https://navercomp.wisereport.co.kr/v2/ETF/index.aspx",
                         params={"cmp_cd": krx},
                         headers={"User-Agent": UA, "Referer": "https://finance.naver.com/"},
                         timeout=25)
        r.encoding = "utf-8"
        m = re.search(r"var\s+CU_data\s*=\s*(\{.*?\});", r.text, re.S)
        if not m:
            print(f"  [WiseReport {krx}] 실패 → HTTP {r.status_code} len={len(r.text)} CU_data 없음", flush=True)
            return [], ""
        g = (json.loads(m.group(1)) or {}).get("grid_data") or []
        d = str(g[0].get("TRD_DT", ""))[:10].replace("-", "") if g else ""
        return g, d
    except Exception as ex:
        print(f"  [WiseReport {krx}] 오류 → {type(ex).__name__}: {str(ex)[:120]}", flush=True)
        return [], ""

def fetch_tiger(isin, ymd):
    # 휴장일 요청은 거부: 미래에셋은 주말 fixDate에 다음 영업일 파일을 주는데 clamp 검사(전영업일과 비교)가 못 잡아
    # 요청한 휴장일을 그대로 기준일로 반환했다(2026-10-29 검증에서 발견, 정규 실행은 영업일만 요청해 영향 없음).
    if not is_trading_day(datetime.strptime(ymd, "%Y%m%d").date()): return {}, ""
    txt = _tiger_raw(isin, ymd)
    if txt:
        # clamp 탐지: 전영업일 응답과 본문이 똑같으면 아직 당일분이 안 나온 것
        try:
            pv = prev_trading_day(datetime.strptime(ymd, "%Y%m%d").date()).strftime("%Y%m%d")
            if _tiger_raw(isin, pv, tries=1) == txt:
                print(f"  [TIGER] {ymd} 미게시(전영업일과 동일 본문) → 대기", flush=True)
                txt = ""
        except Exception:
            pass
    if txt: _CASH_CACHE.pop((isin, ymd), None)   # 재조회(대기 폴링) 시 현금 누적 방지
    out = {}
    for c in _tiger_rows(txt) if txt else []:
        nm = c[1]                               # 0=종목코드 1=종목명 2=수량 3=평가금액 4=비중
        if not nm or nm == "nan": continue
        try: ev = float(c[3].replace(",", ""))
        except (ValueError, TypeError): ev = 0.0
        if any(sk in nm for sk in SKIP):
            _note_cash(isin, ymd, nm, ev); continue   # 직접 경로의 ymd = 실제 기준일(위 clamp·휴장일 검사 통과)
        try: q = float(c[2].replace(",", ""))
        except (ValueError, TypeError): continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, (ev / q if (q and ev) else None)]
    if out: return out, ymd
    # 운용사 직접이 비면 폴백: FunETF(평가금액·현금 포함, 러너 접속 안정) → WiseReport(단가 없음 → KRX 종가).
    # ★ 둘 다 최신분만 준다. 요청일이 아니라 소스가 보고한 실제 기준일을 반환한다(묵은 바스켓의 당일 위장 방지).
    fm, fd = fetch_funetf(isin, key=isin)
    if fm:
        print(f"  [TIGER] FunETF 폴백 {len(fm)}종목 (기준일 {fd})", flush=True)
        return fm, fd
    rows, src = _tiger_wisereport(_TIGER_KRX.get(isin, ""))
    for it in rows:
        nm = str(it.get("STK_NM_KOR", "")).strip()
        if not nm or nm == "nan" or any(sk in nm for sk in SKIP): continue
        try: q = float(str(it.get("AGMT_STK_CNT", 0) or 0).replace(",", ""))
        except (ValueError, TypeError): continue
        if q <= 0: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, None]
    if out:
        print(f"  [TIGER] WiseReport 폴백 {len(out)}종목 (기준일 {src or '미상'})", flush=True)
        return out, src
    return {}, ""

def fetch_unicorn(combo, ymd):
    fund, etf = combo.split(":")
    r = requests.get("https://www.hyundaiam.com/api/etfPdf",
                     params={"fundCode": fund, "etfCode": etf, "ymd": ymd},
                     headers={"User-Agent": UA, "Accept-Language": "ko-KR",
                              "Referer": "https://www.hyundaiam.com/kor/HD-KP-FG/HD-KP-FG-07-D.html"},
                     timeout=20)
    out = {}; actual = ""
    for it in r.json():
        actual = str(it.get("date", "")) or actual
        nm = (it.get("구성종목명") or "").strip()
        if not nm or any(s in nm for s in SKIP): continue
        try: out[nm] = out.get(nm, 0.0) + float(it.get("구성종목수", 0) or 0)
        except (ValueError, TypeError): pass
    return out, (actual or (ymd if out else ""))

# ── fetcher: DB(마이티) — 운용사 사이트는 TOP5만 → WiseReport(전체 수량). snapshot ──
def fetch_db(code):
    import sys as _sys
    SIB = r"C:/Users/정호준/Desktop/코스닥 etf 분류"
    if SIB not in _sys.path: _sys.path.insert(0, SIB)
    from etf_discovery import get_etf_pdf_from_wisereport
    df, d = get_etf_pdf_from_wisereport(code, return_date=True)
    out = {}
    if df is not None and not df.empty:
        for _, row in df.iterrows():
            nm = str(row.get("종목명", "")).strip()
            if not nm or nm == "nan" or any(s in nm for s in SKIP): continue
            try: out[nm] = out.get(nm, 0.0) + float(row.get("수량(주)", 0) or 0)
            except (ValueError, TypeError): pass
    return out, (d or "").replace("-", "")

# ── fetcher: WiseReport(네이버) CU 구성 — FunETF까지 실패했을 때의 최후 폴백(TIGER·RISE·MIDAS) ──
# 실측(2026-10-16 15시): TRD_DT는 PDF 게시일 기준이고 당일분이 아니라 '전일 PDF'다
# (KoAct 10-15 PDF와 74/74 일치, 10-16과 불일치). → 이 소스를 쓰는 ETF는 "lag":True 로 등록해 전일분으로 표시.
def fetch_wisereport(krx, ymd):
    """WiseReport 최신 CU 구성 → ({종목: [수량, None]}, 실제기준일). ymd는 무시(최신분만 제공). 단가 없음 → KRX 종가 폴백."""
    rows, src = _tiger_wisereport(krx)
    out = {}
    for it in rows:
        nm = str(it.get("STK_NM_KOR", "")).strip()
        if not nm or nm == "nan" or any(sk in nm for sk in SKIP): continue
        try: q = float(str(it.get("AGMT_STK_CNT", 0) or 0).replace(",", ""))
        except (ValueError, TypeError): continue
        if q <= 0: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, None]
    return (out, src) if out else ({}, "")

def krx_isin(code):
    """KRX 단축코드 → ISIN (KR7+코드+00+체크디지트: 문자를 숫자로 바꾼 뒤 Luhn). 예: 0191B0 → KR70191B0000."""
    s = "KR7" + code + "00"; digits = "".join(str(int(c, 36)) for c in s); tot = 0
    for i, ch in enumerate(reversed(digits)):
        n = int(ch) * (2 if i % 2 == 0 else 1); tot += n // 10 + n % 10
    return s + str((10 - tot % 10) % 10)

# ── fetcher: FunETF(funetf.co.kr) — 운용사 원본과 같은 PDF를 상품 페이지에 서버렌더로 싣는다 ──
# 2026-10-21~23 러너 로그 132회 실측: GitHub 러너에서 접속 132/132, 당일분 내용은 운용사 원본과 9/9 지문 일치.
# 단 당일분 갱신이 07:20~08:37 사이라 아침 발송 시각엔 전날 PDF다 → lag ETF의 전일분 소스로 쓴다(WiseReport보다 안정,
# 종목별 평가금액·원화현금 포함). 최신분만 주므로 ymd는 받지 않고 실제 기준일(etfPdfYmd)을 반환한다.
_FUNETF_DOWN = False
_FUNETF_ALIAS = {"네이버": "NAVER", "LG씨앤에스": "LG씨엔에스"}   # FunETF만 한글로 쓰는 종목명 → KRX·운용사 표기(2026-10-29 DAISHIN 대조에서 발견)

def fetch_funetf(isin, key=None):
    """FunETF 상품 페이지 → ({종목: [1CU수량, 단가]}, PDF기준일). 원화현금 행은 key(기본 isin)의 1CU 현금으로 기록."""
    global _FUNETF_DOWN
    import re
    if _FUNETF_DOWN: return {}, ""
    try:
        for i in range(2):   # 429(요청 과다)면 잠깐 쉬고 한 번 더 — 회차당 FunETF 호출이 10건 안팎으로 늘었다
            r = requests.get(f"https://www.funetf.co.kr/product/etf/view/{isin}",
                             headers={"User-Agent": UA, "Accept-Language": "ko-KR"}, timeout=(5, 25))
            if r.status_code != 429 or i: break
            import time
            ra = str(r.headers.get("Retry-After") or "5")
            time.sleep(min(15, int(ra) if ra.isdigit() else 5))
    except requests.exceptions.RequestException as ex:     # 접속·응답 시간초과 등 — 이번 실행 동안은 건너뛰고 WiseReport로
        _FUNETF_DOWN = True
        print(f"  [FunETF] 접속 실패({type(ex).__name__}) → 이번 실행에선 건너뜀", flush=True)
        return {}, ""
    d = re.findall(r"etfPdfYmd\s*=\s*['\"]?(\d{8})", r.text)
    m = re.search(r"setPdfChart\(\s*(\[.*?\])\s*\)", r.text, re.S)
    if r.status_code != 200 or not d or not m:
        print(f"  [FunETF {isin}] 구성종목 없음 → HTTP {r.status_code} len={len(r.text)}", flush=True)
        return {}, ""
    try: rows = json.loads(m.group(1))
    except ValueError: return {}, ""
    ymd, key = d[0], (key or isin)
    _CASH_CACHE.pop((key, ymd), None)        # 같은 날짜를 여러 번 받아도(대기 폴링·주간) 현금이 누적되지 않게
    out = {}
    for it in rows:
        nm = str(it.get("citmNm") or "").strip()
        if "/" in nm: nm = nm.split("/")[0].strip()   # 'POSCO홀딩스/포스코홀딩스' → 운용사 원본 표기(소스 전환 시 가짜 편출·편입 방지)
        nm = _FUNETF_ALIAS.get(nm, nm)
        if not nm: continue
        try: q = float(it.get("icuStkc") or 0); ev = float(it.get("evAmt") or 0)
        except (ValueError, TypeError): continue
        if any(sk in nm for sk in SKIP) or not it.get("ticker"):
            _note_cash(key, ymd, nm, ev); continue
        if q <= 0: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, (ev / q if ev else None)]
    return (out, ymd) if out else ({}, "")

def fetch_mirror(code, ymd):
    """자체 사이트에 PDF가 없는 운용사(MIDAS)용: FunETF → WiseReport. 둘 다 최신분만(아침엔 전일분)."""
    fm, fd = fetch_funetf(krx_isin(code), key=code)
    if fm: return fm, fd
    return fetch_wisereport(code, ymd)

# ── fetcher: DS자산운용(dsetf.com) — Next.js 서버렌더 페이지에 최신·직전일 바스켓(수량·평가금액·현금)이 통째로 들어있다 ──
# 날짜 표기 주의: 사이트의 latestHoldingsDate 는 '보유일(종가 시점)'이라 다른 운용사 PDF 기준일보다 하루 앞서 보인다.
# 실측(2026-10-16): DS '10-14' 바스켓 = WiseReport 10-15 PDF와 50/50 일치, DS '10-15'는 10-15 종가로 평가
# (KoAct 10-16 PDF와 같은 평가일) → PDF 기준일 = next_trading_day(표기일). 즉 당일분을 제공한다(전일분 아님).
def _ds_payload(code):
    import re
    r = requests.get(f"https://dsetf.com/etf/{code}", headers={"User-Agent": UA, "Accept": "text/html"}, timeout=25)
    if r.status_code != 200:
        print(f"  [DS] HTTP {r.status_code} len={len(r.text)}", flush=True)
        return ""
    parts = []
    for c in re.findall(r'self\.__next_f\.push\(\[1,"(.*?)"\]\)', r.text, re.S):
        try: parts.append(json.loads('"' + c + '"'))      # JS 문자열 리터럴 = JSON 이스케이프(한글 보존)
        except ValueError: parts.append(c)
    return "".join(parts)

def _ds_array(pl, anchor):
    """pl 안에서 anchor(예: '"latestHoldings":') 바로 뒤의 JSON 배열을 괄호 균형으로 잘라 파싱."""
    i = pl.find(anchor)
    if i < 0: return []
    s = pl.find("[", i + len(anchor))
    if s < 0: return []
    depth = 0; instr = False; esc = False
    for j in range(s, len(pl)):
        ch = pl[j]
        if instr:
            if esc: esc = False
            elif ch == "\\": esc = True
            elif ch == '"': instr = False
            continue
        if ch == '"': instr = True
        elif ch == "[": depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                try: return json.loads(pl[s:j + 1])
                except ValueError: return []
    return []

def _ds_map(items, key, pdf_ymd):
    out = {}
    _CASH_CACHE.pop((key, pdf_ymd), None)   # 같은 날짜를 여러 번 파싱해도(대기 폴링·주간) 현금이 누적되지 않게
    for it in items:
        nm = str(it.get("name", "")).strip()
        if not nm: continue
        try: q = float(it.get("quantity", 0) or 0); mv = float(it.get("marketValue", 0) or 0)
        except (ValueError, TypeError): continue
        if any(sk in nm for sk in SKIP) or not it.get("ticker"):   # 원화현금 행(ticker 빈값) → 1CU 현금
            _note_cash(key, pdf_ymd, nm, mv); continue
        if q <= 0: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, (mv / q if mv else None)]
    return out

def fetch_ds(code, ymd):
    """DS 상품 페이지 → ({종목: [수량, 단가]}, PDF기준일). 요청일이 최신/직전일 바스켓과 맞으면 그것, 아니면 최신을 실제 날짜와 함께."""
    import re
    pl = _ds_payload(code)
    if not pl: return {}, ""
    m = re.search(r'"latestHoldingsDate":"(\d{4})-(\d{2})-(\d{2})"', pl)
    if not m:
        print(f"  [DS] latestHoldingsDate 없음 (len={len(pl)})", flush=True)
        return {}, ""
    cur = next_trading_day(date(int(m.group(1)), int(m.group(2)), int(m.group(3)))).strftime("%Y%m%d")
    cands = [(cur, _ds_array(pl, '"latestHoldings":'))]
    m1 = re.search(r'"1d":\{"date":"(\d{4})-(\d{2})-(\d{2})","items":', pl)
    if m1:
        d1 = next_trading_day(date(int(m1.group(1)), int(m1.group(2)), int(m1.group(3)))).strftime("%Y%m%d")
        cands.append((d1, _ds_array(pl, '"1d":{"date":"%s-%s-%s","items":' % m1.groups())))
    for d, items in cands:
        if d == ymd and items:
            return _ds_map(items, code, d), d
    d, items = cands[0]
    return (_ds_map(items, code, d), d) if items else ({}, "")


# ════ 2026-10-29 추가: 반도체·ESS 그룹 운용사 (출처 탐색·독립 검증: 날짜 정직성·CU·FunETF 대조) ════
# ── fetcher: 우리자산운용(WON) wooriam.kr — 날짜 지정 PDF ──
#   GetPDFListAJX(POST, CSRF 필요)는 상위 10종목만 주지만 행마다 LAST_DATE(실제 기준일)가 있다 → 기준일 확인용.
#   ExcelDownload/PDFofETF(GET)는 전체 바스켓+원화예금(KRD010010001) 행. 날짜 컬럼이 없고 파일명은 요청일 에코.
#   없는 날(미래·휴장일)은 둘 다 빈값(clamp 아님, 2026-10-29 실측: 10-30/10-31/10-09/10-05 → 0행).
#   id = 사이트 fundCode (474590 → 'ETF9'), 상세페이지 uid는 _WON_UID에.
_WON_UID = {"ETF9": "ts8ym1V3HHeDnN9S"}

def fetch_won(fund_code, ymd):
    import re, io
    import pandas as pd
    g = f"{ymd[:4]}.{ymd[4:6]}.{ymd[6:]}"
    s = requests.Session(); s.headers.update({"User-Agent": UA, "Accept-Language": "ko-KR"})
    view = f"https://www.wooriam.kr/investment/etf-view/{_WON_UID.get(fund_code, '')}"
    r = s.get(view, timeout=(5, 20))
    m = re.search(r'name="csrf-token"\s+content="([^"]+)"', r.text)
    if not m: return {}, ""
    a = s.post("https://www.wooriam.kr/investment/GetPDFListAJX",
               files={"fundCode": (None, fund_code), "pdfSearchDate": (None, g)},
               headers={"X-CSRF-TOKEN": m.group(1), "Referer": view, "X-Requested-With": "XMLHttpRequest"},
               timeout=(5, 20))
    try: top = (a.json().get("data") or {}).get("PDFList") or []
    except ValueError: return {}, ""
    dates = {str(it.get("LAST_DATE") or "").replace(".", "") for it in top}
    if len(dates) != 1: return {}, ""
    actual = dates.pop()
    if not re.fullmatch(r"\d{8}", actual): return {}, ""
    ag = f"{actual[:4]}.{actual[4:6]}.{actual[6:]}"
    x = s.get("https://www.wooriam.kr/ExcelDownload/PDFofETF",
              params={"fundCode": fund_code, "pdfSearchDate": ag}, headers={"Referer": view}, timeout=(5, 30))
    try: df = pd.read_excel(io.BytesIO(x.content), dtype=str)   # [번호,종목코드,종목명,수량(주)/액면(원),평가금액(원),비중(%)]
    except Exception: return {}, ""
    try:
        c_cd = next(c for c in df.columns if "코드" in str(c)); c_nm = next(c for c in df.columns if "종목명" in str(c))
        c_q = next(c for c in df.columns if "수량" in str(c)); c_ev = next(c for c in df.columns if "평가" in str(c))
    except StopIteration: return {}, ""
    _CASH_CACHE.pop((fund_code, actual), None)
    out = {}
    for _, v in df.iterrows():
        nm = str(v[c_nm]).strip(); cd = str(v[c_cd]).strip()
        if not nm or nm == "nan": continue
        try: ev = float(str(v[c_ev]).replace(",", ""))
        except (ValueError, TypeError): ev = 0.0
        if any(sk in nm for sk in SKIP) or cd.startswith("KRD") or cd.startswith("CASH"):
            _note_cash(fund_code, actual, nm, ev); continue
        try: q = float(str(v[c_q]).replace(",", ""))
        except (ValueError, TypeError): continue
        if not q: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, (ev / q if ev else None)]
    # 교차검증: AJAX 상위10 수량이 엑셀과 다르면(다른 날 파일) 버린다
    for it in top:
        nm = str(it.get("GUSEONG_NAME") or "").strip()
        if nm in out:
            try:
                if abs(float(str(it.get("JEUNGGWON")).replace(",", "")) - out[nm][0]) > 1e-6: return {}, ""
            except (ValueError, TypeError): pass
    return (out, actual) if out else ({}, "")

# ── fetcher: 대신자산운용(DAISHIN, asset.daishin.com) — 공시/PDF 게시판에 매 영업일 'ETF명_YYYYMMDD' 엑셀이 올라온다 ──
# 흐름: 게시판 검색(POST q=YYYYMMDD) → 제목에 키워드가 들고 '_YYYYMMDD'로 끝나는 글 idx → 상세페이지 vDown.php?idx=파일번호 → xlsx.
# 엑셀 A2 = PDF 기준일(YYYY-MM-DD), 헤더행 [번호,펀드구분,ETF명,자산구분,종목코드,종목명,수량,가격,평가액,비중]. 가격=전영업일 종가.
# 원화현금(KRD010010001)은 1CU 현금으로 기록. 기준일은 요청일이 아니라 엑셀에 적힌 날짜를 반환(미게시·미래일은 글이 없어 빈값).
# 주의: 2026-10-23 게시분은 5개 ETF 전부 첨부가 0바이트(서버 결함) → 그런 날은 FunETF 폴백(최신분·실제 기준일 반환).
_DS_BASE = "https://asset.daishin.com"
_DAISHIN_KRX = {"AI반도체": "486240"}   # 키워드 → KRX 단축코드(FunETF 폴백용)

def fetch_daishin(kw, ymd):
    try:
        out, actual = _daishin_direct(kw, ymd)
    except Exception as ex:      # 상세·첨부 요청 시간초과, 깨진 엑셀, 헤더 변경 등 → FunETF로 넘어간다
        print(f"  [DAISHIN] 직접 경로 실패({type(ex).__name__}) → FunETF", flush=True)
        out, actual = {}, ""
    if out: return out, actual
    code = _DAISHIN_KRX.get(kw, "")
    if not code: return {}, ""
    fm, fd = fetch_funetf(krx_isin(code), key=kw)   # 최신분만 → 실제 기준일 그대로 반환
    if fm: print(f"  [DAISHIN] FunETF 폴백 {len(fm)}종목 (기준일 {fd})", flush=True)
    return (fm, fd) if fm else ({}, "")

def _daishin_direct(kw, ymd):
    import re, io
    import pandas as pd
    hdr = {"User-Agent": UA, "Accept-Language": "ko-KR"}
    lst = f"{_DS_BASE}/ko/?pages=etf&sub=etf5011&tab=tab02"
    try:
        r = requests.post(lst, data={"q": ymd, "y": "", "no": "0"}, headers=hdr, timeout=(5, 20))
    except requests.exceptions.RequestException as ex:
        print(f"  [DAISHIN] 게시판 접속 실패({type(ex).__name__})", flush=True)
        return {}, ""
    hits = re.findall(r"m=view&idx=(\d+)'\);\">([^<]*)</a>", r.text)
    idx = next((i for i, t in hits if kw in t and t.strip().endswith("_" + ymd)), None)
    if not idx:
        return {}, ""                      # 아직 미게시(미래일 포함) → 대기
    v = requests.get(f"{lst}&catecode=19&m=view&idx={idx}", headers=hdr, timeout=(5, 20))
    f = re.search(r"vDown\.php\?idx=(\d+)", v.text)
    if not f:
        print(f"  [DAISHIN] {ymd} 첨부 없음 (글 {idx})", flush=True)
        return {}, ""
    x = requests.get(f"{_DS_BASE}/inc/vDown.php?idx={f.group(1)}",
                     headers=dict(hdr, Referer=f"{lst}&catecode=19&m=view&idx={idx}"), timeout=(5, 30))
    if x.status_code != 200 or not x.content.startswith(b"PK"):
        print(f"  [DAISHIN] 파일 HTTP {x.status_code} len={len(x.content)}", flush=True)
        return {}, ""
    df = pd.read_excel(io.BytesIO(x.content), header=None)
    m = re.search(r"(\d{4})-?(\d{2})-?(\d{2})", str(df.iat[1, 0]))
    actual = "".join(m.groups()) if m else ""
    if not actual:
        return {}, ""
    hrow = next(i for i in range(len(df)) if any("종목명" in str(c) for c in df.iloc[i]))
    cols = [str(c).strip() for c in df.iloc[hrow]]
    c_nm, c_q = cols.index("종목명"), cols.index("수량")
    c_ev = next((k for k, c in enumerate(cols) if "평가" in c), None)
    _CASH_CACHE.pop((kw, actual), None)   # 재조회 시 현금 누적 방지
    out = {}
    for _, row in df.iloc[hrow + 1:].iterrows():
        nm = str(row.iloc[c_nm]).strip()
        if not nm or nm == "nan": continue
        try: ev = float(str(row.iloc[c_ev]).replace(",", "")) if c_ev is not None else 0.0
        except (ValueError, TypeError): ev = 0.0
        if any(s in nm for s in SKIP):
            _note_cash(kw, actual, nm, ev); continue
        try: q = float(str(row.iloc[c_q]).replace(",", ""))
        except (ValueError, TypeError): continue
        if not q: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, (ev / q if ev else None)]
    return (out, actual) if out else ({}, "")

# ── fetcher: 한국투자신탁운용(ACE, aceetf.co.kr) — Next.js SPA가 부르는 papi JSON API ──
#   GET https://papi.aceetf.co.kr/api/funds/{fundCd}/pdf?page=1&size=1000&std_dt=YYYYMMDD
#   → {std_DT:'YYYY-MM-DD', last_STD_DT, pdfList:[{sec_NM, jm_KSC_CD, cu_ITEM_CNT(1CU수량,문자열), val_AM(평가금액), wg}]}
#   std_dt는 대시 없이(YYYYMMDD) — 대시 넣으면 빈값. 미래·휴장일은 std_DT=None·빈 리스트(echo/clamp 없음).
#   pdfList에 원화현금 행이 없다 → 1CU 현금은 FunETF(최신분만 주므로 actual==last_STD_DT이고 FunETF 기준일도 같을 때만)에서 보충.
#   fundCd 찾기: GET papi.aceetf.co.kr/api/funds?searchValue=<이름> → data[].fundCd (385600 = K55101DH7878)
_ACE_KRX = {"K55101DH7878": "385600"}
def fetch_ace(fund_cd, ymd):
    import time
    j = None
    for i in range(3):          # papi는 간헐적으로 응답이 멈춘다(ReadTimeout 실측) → 재시도
        try:
            r = requests.get(f"https://papi.aceetf.co.kr/api/funds/{fund_cd}/pdf",
                             params={"page": 1, "size": 1000, "std_dt": ymd},
                             headers={"User-Agent": UA, "Accept": "application/json",
                                      "Referer": f"https://www.aceetf.co.kr/fund/{fund_cd}",
                                      "Origin": "https://www.aceetf.co.kr"}, timeout=(5, 15))
            j = r.json(); break
        except (requests.exceptions.RequestException, ValueError) as ex:
            print(f"  [ACE {fund_cd}] {ymd} 시도{i+1} 실패: {type(ex).__name__}", flush=True)
            time.sleep(2 * (i + 1))
    if not isinstance(j, dict): return {}, ""
    actual = str(j.get("std_DT") or "").replace("-", "")
    rows = j.get("pdfList") or []
    if not actual or not rows: return {}, ""
    out = {}
    for it in rows:
        nm = str(it.get("sec_NM") or "").strip()
        if not nm: continue
        d_row = str(it.get("std_DT") or "").replace("-", "")
        if d_row and d_row != actual: continue
        try: ev = float(it.get("val_AM") or 0)
        except (ValueError, TypeError): ev = 0.0
        if any(s in nm for s in SKIP):
            _note_cash(fund_cd, actual, nm, ev); continue
        try: q = float(str(it.get("cu_ITEM_CNT") or 0).replace(",", ""))
        except (ValueError, TypeError): continue
        if not q: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, (ev / q if ev else None)]
    latest = str(j.get("last_STD_DT") or "").replace("-", "")
    if out and actual == latest and (fund_cd, actual) not in _CASH_CACHE and fund_cd in _ACE_KRX:
        try:   # 현금행 미제공 → FunETF가 같은 기준일이면 그 원화현금만 차용
            fm, fd = fetch_funetf(krx_isin(_ACE_KRX[fund_cd]), key="_ace_tmp")
            c = _CASH_CACHE.pop(("_ace_tmp", fd), None)
            if fd == actual and c: _CASH_CACHE[(fund_cd, actual)] = c
        except Exception: pass
    return (out, actual) if out else ({}, "")

# ── fetcher: 신한(SOL) soletf.co.kr /api/fund/pdfList (work_dt 날짜 파라미터) → {종목:[1CU수량, 단가]}, 기준일 ──
#   fund_cd = 신한 내부 펀드코드(444200 → 210940; /api/etf/pds/search?keyword= 로 조회). PRICE = 1CU 평가금액(원).
#   미게시·휴일·미래 날짜는 [] 를 준다(echo 없음). 행마다 WORK_DT 가 실려 있어 그것을 실제 기준일로 쓴다.
#   현금성자산(STOCK_CODE KRD010010001) 행 → 1CU 현금.
def fetch_sol(fund_cd, ymd):
    r = requests.get("https://www.soletf.co.kr/api/fund/pdfList",
                     params={"fund_cd": fund_cd, "work_dt": ymd},
                     headers={"User-Agent": UA, "Accept": "application/json",
                              "Referer": f"https://www.soletf.co.kr/ko/fund/etf/{fund_cd}?tabIndex=3"},
                     timeout=(5, 20))
    try: rows = r.json()
    except ValueError: return {}, ""
    if not isinstance(rows, list) or not rows: return {}, ""
    actual = str(rows[0].get("WORK_DT") or "").strip()
    if not actual or any(str(it.get("WORK_DT") or "").strip() != actual for it in rows):
        return {}, ""
    _CASH_CACHE.pop((fund_cd, actual), None)   # 재조회 시 현금 누적 방지(실제 기준일 키)
    out = {}
    for it in rows:
        nm = str(it.get("SEC_NM") or "").strip()
        if not nm: continue
        try: ev = float(it.get("PRICE") or 0)
        except (ValueError, TypeError): ev = 0.0
        cd = str(it.get("STOCK_CODE") or "")
        if any(s in nm for s in SKIP) or cd.startswith("KRD"):
            _note_cash(fund_cd, actual, nm, ev); continue
        try: q = float(it.get("QTY") or 0)
        except (ValueError, TypeError): continue
        if q <= 0: continue
        if nm in out: out[nm][0] += q
        else: out[nm] = [q, (ev / q if ev else None)]
    return (out, actual) if out else ({}, "")

# ── 발송 그룹 ─────────────────────────────────────────────────
# 2026-10-29 사용자 요청: 아침 리포트를 4개로 나눠 각각 보낸다(채널마다 그룹별 사진 1장씩, 길면 그 그룹만 앨범).
# ETFS 의 "grp" 로 소속을 정한다. 그룹마다 따로 완결 판정·발송하므로 한 그룹이 늦어도 다른 그룹은 먼저 나간다.
REPORT_GROUPS = [("kosdaq", "코스닥 액티브"), ("semi", "반도체 액티브"), ("bio", "바이오 액티브"), ("ess", "ESS·2차전지 액티브")]
REPORT_LABEL = dict(REPORT_GROUPS)

def short_name(name):
    """상태줄용 약칭: 첫 단어(운용사 브랜드). 등록 ETF 사이에서 겹치면 둘째 단어('액티브' 뗀 앞 4자)를 붙인다."""
    w = name.split()
    first = [x["name"].split()[0] for x in ETFS]
    return w[0] if first.count(w[0]) == 1 or len(w) < 2 else f"{w[0]} {w[1].replace('액티브', '')[:4]}"

def group_status_line(gs):
    """한 발송 그룹의 상태줄: '캡처 5/6 · 전일분: … · 대기: … · 검증필요 N건'."""
    done = [g for g in gs if g["state"] == "captured"]
    lag = [short_name(g["etf"]) for g in done if g.get("asof")]
    pend = [short_name(g["etf"]) for g in gs if g["state"] != "captured"]
    blocked = sum(1 for g in done for r in g["buy"] + g["sell"] if r.get("검증") == "불가")
    line = (f"캡처 {len(done)}/{len(gs)}" + (f"  ·  전일분: {', '.join(lag)}" if lag else "")
            + (f"  ·  대기: {', '.join(pend)}" if pend else "  ·  전부 반영 완료"))
    return line + (f"  ·  검증필요 {blocked}건" if blocked else "")

# ── 등록된 액티브 ETF ──────────────────────────────────────────
# krx: KRX 단축코드(상장좌수 조회용), cu: 설정단위(CU) 크기 → 전체CU수 = 좌수/cu
ETFS = [
    {"name":"KoAct 코스닥액티브","grp":"kosdaq","am":"삼성액티브","mode":"dateapi","fetch":fetch_samsungactive,"id":"2ETFU6","krx":"0163Y0","cu":50000},
    {"name":"KoAct 바이오헬스케어액티브","grp":"bio","am":"삼성액티브","mode":"dateapi","fetch":fetch_samsungactive,"id":"2ETFJ9","krx":"462900","cu":50000},
    {"name":"TIME 코스닥액티브","grp":"kosdaq","am":"타임폴리오","mode":"dateapi","fetch":fetch_time,"id":"24","krx":"0162Y0","cu":50000},
    {"name":"TIME K바이오액티브","grp":"bio","am":"타임폴리오","mode":"dateapi","fetch":fetch_time,"id":"13","krx":"463050","cu":50000},
    {"name":"PLUS 코스닥150액티브","grp":"kosdaq","am":"한화","mode":"dateapi","fetch":fetch_plus,"id":"006399","krx":"0166N0","cu":50000},
    {"name":"RISE 바이오TOP10액티브","grp":"bio","am":"KB","mode":"dateapi","fetch":fetch_rise,"id":"44I0","krx":"0000Z0","cu":50000,"lag":True},  # 클라우드는 kbam 접속 불가 → FunETF 전일분(2026-10-23)
    {"name":"TIGER 기술이전바이오액티브","grp":"bio","am":"미래에셋","mode":"dateapi","fetch":fetch_tiger,"id":"KR70168K0008","krx":"0168K0","cu":10000,"lag":True},  # CU=10000 (좌수/CU 정수검증으로 확정; 기존 config 20000은 오류)
    # 2026-10-16 추가: 코스닥액티브 나머지 3종. "lag":True = 당일분이 늦거나 전일 PDF까지만 오는 ETF
    #   (클라우드의 TIGER=403·RISE=접속불가 → FunETF/WiseReport, MIDAS=원본 없음 → FunETF/WiseReport. 둘 다 아침엔 전날 PDF)
    #   → 당일분이 없으면 '전일분'으로 표시(run()의 전일분 모드)하고, 아침 발송 완결 판정에서 빠진다(run_daily.LATE).
    #   DS는 자체 사이트가 당일분을 주지만(전일 장 마감 후 갱신) 게시 시각이 검증 전이라 안전망으로 lag.
    {"name":"DS 코스닥액티브","grp":"kosdaq","am":"DS","mode":"dateapi","fetch":fetch_ds,"id":"0220B0","krx":"0220B0","cu":50000,"lag":True},  # CU=50000 (좌수 정수검증 8일 + 1CU평가액÷NAV=50,000)
    {"name":"TIGER 코스닥액티브","grp":"kosdaq","am":"미래에셋","mode":"dateapi","fetch":fetch_tiger,"id":"KR70204S0006","krx":"0204S0","cu":10000,"lag":True},  # CU=10000 (좌수 정수검증)
    # 2026-10-30: SOL·WON·ACE는 아침 07:20에도 당일 PDF 미게시(첫 실전) → 그룹 발송을 붙잡지 않게 lag(없으면 전일분).
    # 2026-10-29 추가 — 반도체(국내 6종, KoAct 반도체&2차전지 혼합형 제외)·ESS·2차전지(국내 3종). 순서 = 그룹 안 표시 순서(순자산 큰 순).
    #   CU는 KRX 실좌수(순자산/NAV) 정수배 + (1CU 평가합+현금)/NAV 로 확인. 운용사 직접 소스는 당일분 → lag 아님(늦으면 그 그룹만 08:40까지 대기).
    #   RISE(kbam)는 클라우드에서 접속 불가 → FunETF 전일분이라 lag. DS는 기존 DS 코스닥과 같은 이유로 lag.
    {"name":"SOL 코리아메가테크액티브","grp":"semi","am":"신한","mode":"dateapi","fetch":fetch_sol,"id":"210940","krx":"444200","cu":50000,"lag":True},   # 2026-10-29 추가(이름에 반도체 없지만 삼성·하이닉스 중심 국내 100%)
    {"name":"WON 반도체밸류체인액티브","grp":"semi","am":"우리","mode":"dateapi","fetch":fetch_won,"id":"ETF9","krx":"474590","cu":50000,"lag":True},
    {"name":"TIGER 코리아테크액티브","grp":"semi","am":"미래에셋","mode":"dateapi","fetch":fetch_tiger,"id":"KR7471780007","krx":"471780","cu":10000,"lag":True},
    {"name":"RISE 비메모리반도체액티브","grp":"semi","am":"KB","mode":"dateapi","fetch":fetch_rise,"id":"44B9","krx":"388420","cu":20000,"lag":True},
    {"name":"DAISHIN AI반도체&인프라액티브","grp":"semi","am":"대신","mode":"dateapi","fetch":fetch_daishin,"id":"AI반도체","krx":"486240","cu":100000},
    {"name":"KoAct AI인프라액티브","grp":"semi","am":"삼성액티브","mode":"dateapi","fetch":fetch_samsungactive,"id":"2ETFN8","krx":"487130","cu":50000},
    {"name":"PLUS AI반도체소부장액티브","grp":"semi","am":"한화","mode":"dateapi","fetch":fetch_plus,"id":"006410","krx":"0239Z0","cu":50000},
    {"name":"RISE 반도체소부장액티브","grp":"semi","am":"KB","mode":"dateapi","fetch":fetch_rise,"id":"44L2","krx":"0254H0","cu":50000,"lag":True},
    {"name":"DS AI반도체액티브","grp":"semi","am":"DS","mode":"dateapi","fetch":fetch_ds,"id":"0243M0","krx":"0243M0","cu":50000,"lag":True},
    {"name":"RISE 2차전지액티브","grp":"ess","am":"KB","mode":"dateapi","fetch":fetch_rise,"id":"44C7","krx":"422420","cu":20000,"lag":True},
    {"name":"KoAct 수소전력ESS인프라액티브","grp":"ess","am":"삼성액티브","mode":"dateapi","fetch":fetch_samsungactive,"id":"2ETFT9","krx":"0150K0","cu":50000},
    {"name":"ACE 2차전지&친환경차액티브","grp":"ess","am":"한국투자","mode":"dateapi","fetch":fetch_ace,"id":"K55101DH7878","krx":"385600","cu":100000,"lag":True},
    {"name":"KoAct 반도체&2차전지핵심소재액티브","grp":"ess","am":"삼성액티브","mode":"dateapi","fetch":fetch_samsungactive,"id":"2ETFM8","krx":"482030","cu":50000},
    {"name":"MIDAS 코스닥액티브","grp":"kosdaq","am":"마이다스","mode":"dateapi","fetch":fetch_mirror,"id":"0191B0","krx":"0191B0","cu":25000,"lag":True},   # 자체 사이트에 PDF 없음 / CU=25000 (1CU평가액÷NAV≈24,993; 좌수 20만좌 고정이라 정수검증 불가)
    # 제외(2026-06-01, 사용자 요청): UNICORN 포스트IPO(현대, 111188:476000) / 마이티(DB, 0001P0, WiseReport).
    #   fetch_unicorn·fetch_db 함수는 코드에 유지 → 재추가 가능.
]

# ── 상장좌수(KRX) → 전체 CU수 계산 ─────────────────────────────
_SHARES_CACHE = {}   # {기준ymd: {KRX단축코드: 좌수}}
_NAV_CACHE = {}      # {기준ymd: {KRX단축코드: 순자산총액(원)}} — 비중 분모(전체 ETF 액수)
_SHARES_ASOF = {}    # {기준ymd: 실제로 읽어온 영업일} — 진단용
_KRX_VOL_CACHE = {}  # {ymd: {종목명: 거래량(주)}} — 타당성 가드용
_KRX_PX_CACHE = {}   # {종목명: 종가} — 운용사가 평가금액 미제공(예: PLUS)일 때 종가 fallback
def _krx_key():
    k = os.getenv("KRX_API_KEY", "")
    if k: return k
    # 로컬: sibling .env 에서 로드 (클라우드는 환경변수/시크릿 사용)
    envp = r"C:/Users/정호준/Desktop/코스닥 etf 분류/.env"
    try:
        for line in open(envp, encoding="utf-8", errors="replace"):
            if line.strip().startswith("KRX_API_KEY"):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    except OSError:
        pass
    return ""

def _krx_get(url, params, tries=3):
    """KRX OpenAPI GET + 재시도(백오프). 일시적 장애로 금액·순자산이 통째로 빈칸되는 것 방지."""
    import time
    for i in range(tries):
        try:
            r = requests.get(url, params=params, timeout=20)
            if r.status_code == 200:
                return r.json()
        except Exception:
            pass
        if i < tries - 1: time.sleep(1.5 * (i + 1))
    return {}

def _asof_date(asof):
    """asof(None / date / 'YYYYMMDD') → date. None이면 오늘."""
    if asof is None: return datetime.now().date()
    if isinstance(asof, str): return datetime.strptime(asof, "%Y%m%d").date()
    return asof

def listed_shares(asof=None):
    """{KRX단축코드: 좌수} — PDF 기준일 asof의 바스켓을 환산할 때 쓸 좌수.
    PDF(T) 바스켓은 T-1 종가 시점 보유이므로 좌수도 T-1 것을 쓴다.

    ★ LIST_SHRS(상장좌수)를 쓰지 않는다 ★
      같은 행 안에 있지만 그 필드는 설정·환매 결제 지연 때문에 실좌수와 어긋난다.
      실측(2026-09-08~15, 4개 ETF): 차이 -2 ~ +7 CU, 최대 1.1%, 부호도 뒤집힘.
      실좌수 = 순자산총액 ÷ NAV 로 역산해야 하며, 이때만
      '1CU 바스켓 평가액 × 좌수 = 순자산' 항등식이 성립한다.
      순자산이 실린 영업일만 채택하고(KRX가 최신일엔 0으로 주는 날이 있다),
      NAV·순자산이 모두 없을 때만 LIST_SHRS로 폴백한다.
    """
    start = prev_trading_day(_asof_date(asof))
    ck = start.strftime("%Y%m%d")
    if ck in _SHARES_CACHE: return _SHARES_CACHE[ck]
    _SHARES_CACHE[ck] = {}; _NAV_CACHE[ck] = {}
    key = _krx_key()
    if not key: return _SHARES_CACHE[ck]
    def _n(it, k):
        try: return float(str(it.get(k, "0")).replace(",", "") or 0)
        except (ValueError, TypeError): return 0.0
    d = start
    for _ in range(8):
        ymd = d.strftime("%Y%m%d")
        items = _krx_get("http://data-dbg.krx.co.kr/svc/apis/etp/etf_bydd_trd",
                         {"AUTH_KEY": key, "basDd": ymd}).get("OutBlock_1", [])
        if items and any(_n(it, "INVSTASST_NETASST_TOTAMT") for it in items):
            sh = {}; nv = {}
            for it in items:
                code = str(it.get("ISU_SRT_CD") or it.get("ISU_CD") or "")
                if not code: continue
                nav = _n(it, "NAV"); na = _n(it, "INVSTASST_NETASST_TOTAMT")
                if na: nv[code] = int(na)
                sh[code] = int(round(na / nav)) if (nav and na) else int(_n(it, "LIST_SHRS"))
            _SHARES_CACHE[ck] = sh; _NAV_CACHE[ck] = nv; _SHARES_ASOF[ck] = ymd
            break
        d = prev_trading_day(d)
    return _SHARES_CACHE[ck]

def num_cu(e, asof=None):
    """ETF의 전체 CU수 = 좌수/CU. 좌수 미상이면 0."""
    sh = listed_shares(asof).get(e.get("krx", ""), 0)
    cu = e.get("cu", 0)
    return (sh / cu) if (sh and cu) else 0

def etf_navtotal(e, asof=None):
    """ETF 순자산총액(원, KRX 공식) = 비중 분모. 미상이면 0."""
    listed_shares(asof)  # 캐시 채움(좌수와 동일 API)
    ck = prev_trading_day(_asof_date(asof)).strftime("%Y%m%d")
    return _NAV_CACHE.get(ck, {}).get(e.get("krx", ""), 0)

def etf_cash_percu(e, ymd):
    """1CU 현금(원) — PDF 현금라인 평가금액. 페처가 채운 캐시에서. 미상이면 None."""
    return _CASH_CACHE.get((e.get("id"), ymd))

def _lagtag(g, sep="  "):
    """전일분 그룹 표시 '※전일분 PDF 10/16' — 당일분이면 빈 문자열."""
    a = g.get("asof")
    return f"{sep}※전일분 PDF {a[4:6]}/{a[6:]}" if a else ""

def _cashlabel(g):
    """배너용 현금 문구: '현금 61억(1.2%)'. 미상이면 빈 문자열."""
    c = g.get("cash")
    if c is None: return ""
    r = g.get("cashr")
    return f"현금 {c/1e8:,.0f}억" + (f"({r}%)" if r is not None else "")

VOL_BLOCK = 1.00   # 구간 누적 거래량의 100% 초과 → 금액 발행 중단(장내 체결 불가)
VOL_WARN  = 0.50   # 50% 초과 → 경고만(금액은 유지)

def krx_kosdaq_volumes(ymd):
    """{종목명: 그날 거래량(주)} — 코스닥. 미공시·실패 시 빈 dict.
    코스피 종목은 KRX 구독 밖이라 조회되지 않는다(= 가드 사각지대, 그 경우 검사 생략)."""
    if ymd in _KRX_VOL_CACHE: return _KRX_VOL_CACHE[ymd]
    _KRX_VOL_CACHE[ymd] = {}
    key = _krx_key()
    if not key: return _KRX_VOL_CACHE[ymd]
    items = _krx_get("http://data-dbg.krx.co.kr/svc/apis/sto/ksq_bydd_trd",
                     {"AUTH_KEY": key, "basDd": ymd}).get("OutBlock_1", [])
    m = {}
    for it in items:
        try: m[it["ISU_NM"].strip()] = float(str(it["ACC_TRDVOL"]).replace(",", ""))
        except (ValueError, KeyError, TypeError): pass
    _KRX_VOL_CACHE[ymd] = m
    return m

def flow_window(pk, t_ymd, cap=15):
    """변화가 실제로 일어난 거래일 구간 = [pk, t의 직전 영업일].
    PDF(T)는 T-1 종가 시점 보유이므로 PDF(pk)→PDF(T) 변화는 pk일부터 T-1일까지의 매매다.
    스냅샷이 며칠 비면 구간이 자동으로 넓어져 '하루 거래량 대비' 오탐을 막는다."""
    a = datetime.strptime(pk, "%Y%m%d").date()
    b = prev_trading_day(datetime.strptime(t_ymd, "%Y%m%d").date())
    out = []; d = a
    while d <= b and len(out) < cap:
        if is_trading_day(d): out.append(d.strftime("%Y%m%d"))
        d += timedelta(days=1)
    return out

def volume_ratio(nm, shares, days):
    """(필요 주식수 / 구간 누적 거래량, 누적 거래량). 거래량을 못 구하면 (None, 0)."""
    tot = 0.0
    for ymd in days:
        tot += krx_kosdaq_volumes(ymd).get(nm, 0.0)
    if tot <= 0: return None, 0.0
    return shares / tot, tot

def last_held(kk, before_ymd, nm, max_back=30, max_gap=30):
    """before_ymd 이전 스냅샷에서 nm을 마지막으로 보유한 (기준일, 수량). 없으면 None.
    '신규편입'이 사실은 며칠 공백 뒤의 재등장인지 가린다.
    max_gap(일)보다 오래된 보유는 무시한다 — 그 정도면 진짜 신규로 본다."""
    try: lim = datetime.strptime(before_ymd, "%Y%m%d").date() - timedelta(days=max_gap)
    except ValueError: return None
    for ymd in sorted((d for d in kk if d < before_ymd), reverse=True)[:max_back]:
        try:
            if datetime.strptime(ymd, "%Y%m%d").date() < lim: break
        except ValueError: continue
        v = (kk.get(ymd) or {}).get(nm)
        if v is None: continue
        q = _qp(v)[0]
        if q > 0: return ymd, q
    return None

def krx_kosdaq_prices():
    """{종목명: 종가} — KRX 코스닥 최근 영업일 종가. 운용사 평가금액 미제공 종목의 종가 fallback."""
    if _KRX_PX_CACHE: return _KRX_PX_CACHE
    key = _krx_key()
    if not key: return {}
    for off in range(1, 8):
        d = (datetime.now().date() - timedelta(days=off)).strftime("%Y%m%d")
        items = _krx_get("http://data-dbg.krx.co.kr/svc/apis/sto/ksq_bydd_trd",
                         {"AUTH_KEY": key, "basDd": d}).get("OutBlock_1", [])
        if items:
            for it in items:
                try: _KRX_PX_CACHE[it["ISU_NM"].strip()] = float(str(it["TDD_CLSPRC"]).replace(",", ""))
                except (ValueError, KeyError): pass
            break
    return _KRX_PX_CACHE

def _qp(v):
    """맵 값 → (수량, 종가). 값은 [수량,종가](신형) 또는 숫자(구형) 모두 허용."""
    if isinstance(v, (list, tuple)):
        return float(v[0]), (float(v[1]) if v[1] not in (None, "") else None)
    return float(v), None

def _px(nm, primary):
    """종가: 운용사 평가금액 우선, 없으면 KRX 코스닥 종가 fallback."""
    return primary if primary else krx_kosdaq_prices().get(nm)

def _basket_total(m):
    """맵의 1CU 총평가액(종가 fallback 포함) — 보유비중 분모."""
    return sum(q * pr for q, pr in ((qp[0], _px(nm, qp[1])) for nm, qp in ((k, _qp(v)) for k, v in m.items())) if pr)

def diff(tmap, pmap):
    Tt = _basket_total(tmap); Tp = _basket_total(pmap)   # 당일·전일 총평가액(보유비중 분모)
    rows = []
    for nm in set(tmap) | set(pmap):
        tq, tp = _qp(tmap.get(nm, (0.0, None)))
        pq, pp = _qp(pmap.get(nm, (0.0, None)))
        ch = tq - pq
        if pq == 0 and tq > 0:   typ = "신규편입"
        elif tq == 0 and pq > 0: typ = "편출"
        elif ch > 0:             typ = "수량확대"
        elif ch < 0:             typ = "수량축소"
        else:                    typ = "유지"
        price = _px(nm, tp if tp else pp)   # 당일 종가 우선, 편출이면 전일 종가(둘 다 없으면 KRX)
        amt = round(ch * price) if price else None   # 변화금액(원) = 변화 × 종가
        tpe = _px(nm, tp); ppe = _px(nm, pp)
        w_t = (tq * tpe / Tt * 100) if (tpe and Tt) else None   # 당일 보유비중(%)
        w_p = (pq * ppe / Tp * 100) if (ppe and Tp) else None   # 전일 보유비중(%)
        rows.append({"종목명": nm, "당일수량": tq, "전일수량": pq, "변화": ch, "구분": typ, "변화금액": amt,
                     "보유비중당일": w_t, "보유비중전일": w_p})
    rows.sort(key=lambda r: -abs(r["변화"]))
    return rows

def _prev_key(kk, t_ymd, max_gap=10):
    """당일 이전 중 실제 보유한 가장 최근 날짜 = 사실상의 전영업일.
    휴장일이 달력에서 누락돼도(예: 제헌절) 자동 보정. 너무 오래된 건 제외."""
    cands = sorted(d for d in kk if d < t_ymd)
    if not cands: return None
    pk = cands[-1]
    try:
        gap = (datetime.strptime(t_ymd, "%Y%m%d") - datetime.strptime(pk, "%Y%m%d")).days
    except ValueError:
        return None
    return pk if gap <= max_gap else None

def _fetch(e, ymd, tries=3):
    """모드 무관 통일 호출 → (map, 기준일YYYYMMDD). 일시적 네트워크 오류는 재시도."""
    import time
    last = None
    for i in range(tries):
        try:
            if e["mode"] == "dateapi":
                return e["fetch"](e["id"], ymd)
            return e["fetch"](e["id"])  # snapshot: ymd 무시, 최신 반환
        except Exception as ex:
            last = ex
            if i < tries - 1: time.sleep(1.5 * (i + 1))
    raise last


def run(today=None, weekly=True):
    """증분 구조: 이미 당일분 캡처된 ETF는 재요청 스킵, 미게시/지연만 매번 시도해 채움.
    매 실행 시 당일 파일(xlsx/csv)을 '현재까지 캡처된 전체 상태'로 덮어씀."""
    today = today or datetime.now().date()
    if not is_trading_day(today):
        print(f"[{today}] 휴장일 — 스킵"); return
    t_ymd = today.strftime("%Y%m%d")
    prev = prev_trading_day(today); p_ymd = prev.strftime("%Y%m%d")
    snap = load_snap()
    print(f"=== PDF 변화(증분): 당일 {today} vs 전영업일 {prev} ===")
    for e in ETFS:
        key = f"{e['am']}:{e['id']}"; snap.setdefault(key, {})
        # 1) 전일 베이스라인 (없을 때만 1회)
        if p_ymd not in snap[key]:
            try:
                pm, pa = _fetch(e, p_ymd)
                if pm and pa: snap[key][pa] = pm
            except Exception as ex:
                print(f"  [{e['name'][:16]}] 전일 조회 실패: {type(ex).__name__}")
        # 2) 당일 — 이미 있으면 스킵(증분), 없으면 시도
        if t_ymd in snap[key]:
            status = "이미 캡처(스킵)"
        else:
            try:
                tm, ta = _fetch(e, t_ymd)
                if tm and ta == t_ymd:
                    snap[key][t_ymd] = tm; status = f"신규 캡처 {len(tm)}종목"
                else:
                    status = f"대기 (현재 {ta or '빈값'})"
            except Exception as ex:
                status = f"대기 (오류 {type(ex).__name__})"
        ready = (t_ymd in snap[key]) and bool(_prev_key(snap[key], t_ymd))
        print(f"  {'✅' if ready else '⏳'} {e['name'][:18]:18} {status}")
    save_snap(snap)

    # 출력 빌드 (스냅샷에서, 8개 전부 상태 포함)
    groups = []; csv_rows = []; done = 0; pend = []; blocked = 0
    lagged = []
    def _short(e): return short_name(e["name"])
    for e in ETFS:
        kk = snap.get(f"{e['am']}:{e['id']}", {})
        cur = t_ymd if t_ymd in kk else None
        pk = _prev_key(kk, t_ymd) if cur else None   # 실제 전영업일(휴장 누락 자동 보정)
        # 전일분 모드: 소스가 전일 PDF까지만 주는 ETF("lag")는 당일분이 없으면 '전일 PDF vs 그 전일'을 대신 보여준다
        # (매일 하루 늦게, 빠짐없이). 당일분이 들어오면 자동으로 정상 경로. 2거래일 이상 묵은 건 대기로 둔다.
        if cur is None and e.get("lag"):
            lp = _prev_key(kk, t_ymd)
            if lp and lp >= p_ymd and _prev_key(kk, lp):
                cur, pk = lp, _prev_key(kk, lp)
        if cur and pk:
            done += 1
            if cur != t_ymd: lagged.append(_short(e))
            rows = [r for r in diff(kk[cur], kk[pk])
                    if r["구분"] != "유지"
                    and (r["구분"] in ("신규편입", "편출") or abs(r["변화"]) >= MIN_CHANGE)]
            nc = num_cu(e, cur)
            stockval = _basket_total(kk[cur]) * nc  # 보유주식총액(KRX 종가 fallback 포함 → PLUS 현금비율 정확)
            aum = etf_navtotal(e, cur) or (round(stockval) if stockval else None)  # 비중 분모 = KRX 순자산총액(우선)
            win = flow_window(pk, cur)   # 이 변화가 실제로 일어난 거래일 구간
            for r in rows:
                # (1) '신규편입'이 사실은 공백 뒤 재등장인지 확인 → 재편입으로 정정
                if r["구분"] == "신규편입":
                    lh = last_held(kk, pk, r["종목명"])
                    if lh: r["구분"] = "재편입"; r["직전보유"] = lh
                # (2) 금액
                r["전체금액"] = round(r["변화금액"] * nc) if (r["변화금액"] is not None and nc) else None
                # (3) 물리적 타당성: 필요 주식수가 구간 시장 거래량을 넘으면 금액을 발행하지 않는다
                need = abs(r["변화"]) * nc if nc else 0
                ratio, wvol = volume_ratio(r["종목명"], need, win)
                r["거래량비"] = ratio; r["구간거래량"] = wvol; r["구간일수"] = len(win)
                if ratio is not None and ratio > VOL_BLOCK:
                    r["검증"] = "불가"; r["보류금액"] = r["전체금액"]; r["전체금액"] = None
                elif ratio is not None and ratio > VOL_WARN:
                    r["검증"] = "의심"
                r["비중"] = round(r["전체금액"] / aum * 100, 3) if (r.get("전체금액") is not None and aum) else None
            buy = sorted([r for r in rows if r["구분"] in ("신규편입", "재편입", "수량확대")], key=lambda r: -r["변화"])
            sell = sorted([r for r in rows if r["구분"] in ("수량축소", "편출")], key=lambda r: r["변화"])
            cpc = etf_cash_percu(e, cur); spc = stockval / nc if nc else 0   # 1CU 현금·주식
            cash = round(cpc * nc) if (cpc is not None and nc) else None
            cashr = round(cpc / (cpc + spc) * 100, 1) if (cpc is not None and (cpc + spc) > 0) else None
            groups.append({"etf": e["name"], "am": e["am"], "state": "captured", "buy": buy, "sell": sell,
                           "aum": aum, "cash": cash, "cashr": cashr, "pk": pk, "win": len(win),
                           "asof": (cur if cur != t_ymd else None),    # asof = 전일분일 때의 PDF 기준일
                           "grp": e.get("grp", "kosdaq")})
            blocked += sum(1 for r in rows if r.get("검증") == "불가")
            for r in rows:
                csv_rows.append([e["name"], e["am"], cur, pk, r["구분"], r["종목명"],
                                 int(r["당일수량"]), int(r["전일수량"]), int(r["변화"]),
                                 (r["변화금액"] if r["변화금액"] is not None else ""),
                                 (r["전체금액"] if r["전체금액"] is not None else ""),
                                 (r["비중"] if r.get("비중") is not None else ""),
                                 (round(r["보유비중전일"], 2) if r.get("보유비중전일") is not None else ""),
                                 (round(r["보유비중당일"], 2) if r.get("보유비중당일") is not None else ""),
                                 r.get("검증", ""),
                                 (round(r["거래량비"] * 100, 1) if r.get("거래량비") is not None else ""),
                                 (int(r["구간거래량"]) if r.get("구간거래량") else ""),
                                 r.get("구간일수", ""),
                                 (r["직전보유"][0] if r.get("직전보유") else ""),
                                 (int(r["직전보유"][1]) if r.get("직전보유") else ""),
                                 (r.get("보류금액") if r.get("보류금액") is not None else "")])
        else:
            groups.append({"etf": e["name"], "am": e["am"], "state": "pending", "buy": [], "sell": [],
                           "grp": e.get("grp", "kosdaq")})
            pend.append(_short(e))
    status_line = (f"캡처 {done}/{len(ETFS)}" + (f"  ·  전일분: {', '.join(lagged)}" if lagged else "")
                   + (f"  ·  대기: {', '.join(pend)}" if pend else "  ·  전부 반영 완료"))
    if blocked: status_line += f"  ·  검증필요 {blocked}건"
    csv_path = os.path.join(BASE, f"pdf_change_{t_ymd}.csv")
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["ETF","운용사","당일기준일","전일기준일","구분","종목명","당일수량","전일수량","수량변화(1CU)","1CU변화금액(원)","전체매매금액(원)","전체대비비중(%)","보유비중전일(%)","보유비중당일(%)","검증","거래량비(%)","구간거래량(주)","구간일수","직전보유일","직전보유수량","보류금액(원)"])
        for r in csv_rows: w.writerow(r)
    xlsx_path = write_excel(groups, today, prev, status_line)
    print(f"\n[{status_line}]  변화 {len(csv_rows)}행 → {os.path.basename(xlsx_path)}")

    # 그 주 마지막 거래일(보통 금요일, 휴장이면 목요일)이면 주간 리포트도 생성.
    # weekly=False 는 호출자가 나중에 한 번만 만들겠다는 뜻(대기 루프가 폴링마다 재생성하는 것 방지).
    if weekly and is_last_trading_day_of_week(today):
        try:
            weekly_report(today)
        except Exception as ex:
            print(f"[주간] 생성 실패: {type(ex).__name__}: {ex}")

    return {"groups": groups, "today": today, "prev": prev, "status_line": status_line,
            "csv_path": csv_path, "xlsx_path": xlsx_path, "done": done, "pending": pend,
            "krx_ok": bool(listed_shares())}   # KRX 좌수 확보 여부(금액·순자산·현금 계산 가능한지)


def week_trading_days(any_day):
    """any_day가 속한 주의 거래일 리스트(과거→현재). 월요일부터 any_day까지 중 거래일만.
    (금요일 실행=그 주 전체, 주중 실행=그 주 진행분까지 = '이번주 현재까지')."""
    monday = any_day - timedelta(days=any_day.weekday())
    return [monday + timedelta(days=i) for i in range(any_day.weekday() + 1)
            if is_trading_day(monday + timedelta(days=i))]

def is_last_trading_day_of_week(d):
    """d가 그 주(월~금)의 마지막 거래일인가 (이후 금요일까지 거래일 없음)."""
    if not is_trading_day(d):
        return False
    fri = d - timedelta(days=d.weekday()) + timedelta(days=4)
    nxt = d + timedelta(days=1)
    while nxt <= fri:
        if is_trading_day(nxt):
            return False
        nxt += timedelta(days=1)
    return True

def nth_trading_day_before(d, n):
    x = d
    for _ in range(n):
        x = prev_trading_day(x)
    return x

def _fetch_exact(e, ymd):
    """요청일과 실제 기준일이 같을 때만 맵을 준다 — 최신분만 주는 소스(DS·WiseReport)가 다른 날짜로 위장하는 것 방지."""
    try:
        m, a = _fetch(e, ymd)
        return m if (m and a == ymd) else {}
    except Exception:
        return {}

def _build_period_groups(first, last):
    """first↔last PDF 비교 그룹 빌드. return (groups, done, pend_names).
    일일 캡처 스냅샷을 먼저 쓰고, 없을 때만 소스에서 받되 날짜가 정확히 맞는 것만 인정한다."""
    f_ymd, l_ymd = first.strftime("%Y%m%d"), last.strftime("%Y%m%d")
    groups = []; done = 0; pend = []
    snap = load_snap()
    for e in ETFS:
        kk = snap.get(f"{e['am']}:{e['id']}", {})
        fm = kk.get(f_ymd) or _fetch_exact(e, f_ymd)
        lm = kk.get(l_ymd) or _fetch_exact(e, l_ymd)
        if fm and lm:
            done += 1
            rows = [r for r in diff(lm, fm)
                    if r["구분"] != "유지"
                    and (r["구분"] in ("신규편입", "편출") or abs(r["변화"]) >= MIN_CHANGE)]
            nc = num_cu(e, l_ymd)
            stockval = _basket_total(lm) * nc  # 보유주식총액(마지막일, KRX 종가 fallback 포함)
            aum = etf_navtotal(e, l_ymd) or (round(stockval) if stockval else None)  # 비중 분모 = KRX 순자산총액(우선)
            win = flow_window(f_ymd, l_ymd)
            for r in rows:
                r["전체금액"] = round(r["변화금액"] * nc) if (r["변화금액"] is not None and nc) else None
                need = abs(r["변화"]) * nc if nc else 0
                ratio, wvol = volume_ratio(r["종목명"], need, win)
                r["거래량비"] = ratio; r["구간거래량"] = wvol; r["구간일수"] = len(win)
                if ratio is not None and ratio > VOL_BLOCK:
                    r["검증"] = "불가"; r["보류금액"] = r["전체금액"]; r["전체금액"] = None
                elif ratio is not None and ratio > VOL_WARN:
                    r["검증"] = "의심"
                r["비중"] = round(r["전체금액"] / aum * 100, 3) if (r.get("전체금액") is not None and aum) else None
            buy = sorted([r for r in rows if r["구분"] in ("신규편입", "재편입", "수량확대")], key=lambda r: -r["변화"])
            sell = sorted([r for r in rows if r["구분"] in ("수량축소", "편출")], key=lambda r: r["변화"])
            cpc = etf_cash_percu(e, l_ymd); spc = stockval / nc if nc else 0
            cash = round(cpc * nc) if (cpc is not None and nc) else None
            cashr = round(cpc / (cpc + spc) * 100, 1) if (cpc is not None and (cpc + spc) > 0) else None
            groups.append({"etf": e["name"], "am": e["am"], "state": "captured", "buy": buy, "sell": sell,
                           "aum": aum, "cash": cash, "cashr": cashr})
        else:
            groups.append({"etf": e["name"], "am": e["am"], "state": "pending", "buy": [], "sell": []})
            pend.append(e["name"].split()[0])
    return groups, done, pend

def weekly_report(asof=None):
    """그 주 첫 거래일 PDF ↔ 마지막 거래일 PDF 비교 = 주간 순변화."""
    asof = asof or datetime.now().date()
    tdays = week_trading_days(asof)
    if len(tdays) < 2:
        print(f"[주간] {asof} 주의 거래일 {len(tdays)}일 — 비교 불가"); return None
    first, last = tdays[0], tdays[-1]
    print(f"=== 주간 변화: {first}(주초) ↔ {last}(주말)  거래일 {len(tdays)}일 ===")
    groups, done, pend = _build_period_groups(first, last)
    span = f"{first.strftime('%Y-%m-%d')} ~ {last.strftime('%Y-%m-%d')}  ({len(tdays)}거래일)"
    status = (f"주간 변화  ·  {span}  ·  캡처 {done}/{len(ETFS)}" + (f"  ·  대기: {', '.join(pend)}" if pend else ""))
    path = write_excel(groups, last, first, status, title=f"코스닥 액티브 ETF 주간 PDF 변화   ·   {span}",
                       lbl_cur="주말", lbl_prev="주초", fname=f"pdf_weekly_{last.strftime('%Y%m%d')}.xlsx")
    print(f"[주간 저장] {os.path.basename(path)}")
    return path

def rolling_report(asof=None, ndays=5):
    """오늘 기준 최근 N거래일 순변화 = PDF(오늘) ↔ PDF(N거래일 전)."""
    asof = asof or datetime.now().date()
    while not is_trading_day(asof):
        asof = prev_trading_day(asof)
    last = asof
    first = nth_trading_day_before(last, ndays)
    print(f"=== 최근 {ndays}거래일 변화: {first} ↔ {last} ===")
    groups, done, pend = _build_period_groups(first, last)
    span = f"{first.strftime('%Y-%m-%d')} ~ {last.strftime('%Y-%m-%d')}  (최근 {ndays}거래일)"
    status = (f"최근 {ndays}거래일 변화  ·  {span}  ·  캡처 {done}/{len(ETFS)}" + (f"  ·  대기: {', '.join(pend)}" if pend else ""))
    path = write_excel(groups, last, first, status, title=f"코스닥 액티브 ETF 최근 {ndays}거래일 PDF 변화   ·   {span}",
                       lbl_cur=f"오늘", lbl_prev=f"{ndays}일전", fname=f"pdf_rolling{ndays}_{last.strftime('%Y%m%d')}.xlsx")
    print(f"[롤링 저장] {os.path.basename(path)}")
    return path


def _rowtag(r, short=False):
    """종목명 뒤 표시. '재편입'은 공백 전 보유를 같이 알려 '신규'로 오독되는 것을 막는다."""
    t = r.get("구분")
    if t == "재편입":
        lh = r.get("직전보유")
        s0 = "  (재편입)" if (short or not lh) else f"  (재편입·{lh[0][4:6]}/{lh[0][6:]}까지 {int(lh[1]):,}주)"
    elif t == "신규편입": s0 = "  (신규)"
    elif t == "편출":     s0 = "  (편출)"
    else:                 s0 = ""
    v = r.get("검증")
    # 모바일(short)에서는 금액 칸이 이미 '검증필요'를 표시하므로 이름 쪽 중복을 뺀다(글자 축소 방지).
    if v == "불가" and not short: s0 += " ※검증필요"   # ⚠(U+26A0)은 맑은고딕에 없어 네모로 깨진다
    elif v == "의심":             s0 += " ※"
    return s0

def _namecell(r, short=False):
    """종목명 셀. 표시(재편입·편출·검증)는 둘째 줄로 내려 글자 자동축소를 막는다."""
    t = _rowtag(r, short).strip()
    return r["종목명"] + (chr(10) + t if t else "")

def _amtcell(r, short=False):
    """금액 셀. 타당성 가드에 걸린 행은 금액 대신 사유를 쓴다(단정형 오보 방지)."""
    if r.get("검증") == "불가":
        rt = r.get("거래량비")
        if short or rt is None: return "검증필요"
        return "검증필요" + chr(10) + f"(거래량 {rt*100:,.0f}%)"
    a = r.get("전체금액")
    return "-" if a is None else f"{a/1e8:+,.1f}억"

def write_excel(groups, today, prev, status_line="", title=None, lbl_cur="당일", lbl_prev="전영업일", fname=None):
    """8개 ETF 전부 표시(변화있음=좌매수/우매도, 변화없음, 대기). 날짜 YYYY-MM-DD.
    groups: [{etf, am, state('captured'/'pending'), buy:[diff행dict], sell:[...]}]
    title/lbl_cur/lbl_prev: 주간 등 다른 기간 리포트용 라벨 오버라이드."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "PDF변화"
    RED = Font(color="C00000", bold=True); BLUE = Font(color="0070C0", bold=True)
    BANNERF = Font(bold=True, size=11, color="1F3864"); BANNERBG = PatternFill("solid", fgColor="DCE6F2")
    PENDF = Font(bold=True, size=11, color="7F7F7F"); PENDBG = PatternFill("solid", fgColor="EDEDED")
    BUYHEAD = Font(bold=True, color="FFFFFF"); BUYHBG = PatternFill("solid", fgColor="C00000")
    SELLHEAD = Font(bold=True, color="FFFFFF"); SELLHBG = PatternFill("solid", fgColor="0070C0")
    BUYBG = PatternFill("solid", fgColor="FCE4E4"); SELLBG = PatternFill("solid", fgColor="E4ECF6")
    thin = Side(style="thin", color="D9D9D9"); border = Border(left=thin, right=thin, top=thin, bottom=thin)
    NUMFMT_UP = "+#,##0;-#,##0;0"   # 수량(주) 부호표시
    AMTFMT = "+#,##0;-#,##0;0"      # 변화금액(원) 부호표시
    PCTFMT = '+0.00"%";-0.00"%";0'  # 전체대비 비중 부호표시
    td = today.strftime("%Y-%m-%d"); pd_ = prev.strftime("%Y-%m-%d")
    NC = 11  # A~K (좌5 + gap + 우5)
    def mergerow(r): ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)

    mergerow(1)
    ws["A1"] = title or f"액티브 ETF PDF 변화   ·   {lbl_cur} {td}  vs  {lbl_prev} {pd_}"
    ws["A1"].font = Font(bold=True, size=14); ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 26
    if status_line:
        mergerow(2); ws["A2"] = status_line
        ws["A2"].font = Font(size=10, color="595959", italic=True)
    row = 3
    for g in groups:
        etf, am, state = g["etf"], g["am"], g["state"]
        buy = sorted(g["buy"], key=lambda r: -r["변화"]); sell = sorted(g["sell"], key=lambda r: r["변화"])
        mergerow(row)
        if state == "pending":
            bc = ws.cell(row, 1, f"■ {etf}  ({am})      ⏳ 대기 — 당일 PDF 미게시/지연 (게시되면 자동 반영)")
            bc.font = PENDF; bc.alignment = Alignment(vertical="center")
            for c in range(1, NC + 1): ws.cell(row, c).fill = PENDBG
            ws.row_dimensions[row].height = 19; row += 2; continue
        aum = g.get("aum")
        aum_txt = f"      순자산 ≈ {aum/1e8:,.0f}억" if aum else ""
        _cl = _cashlabel(g); cash_txt = f"   ·   {_cl}" if _cl else ""
        if g.get("asof"):   # 전일분: 이 그룹만의 실제 기준일을 쓴다
            a, p = g["asof"], (g.get("pk") or "")
            dates_txt = f"      ※전일분  PDF {a[:4]}-{a[4:6]}-{a[6:]}  vs  {p[:4]}-{p[4:6]}-{p[6:]}"
        else:
            dates_txt = f"      {lbl_cur} {td}  vs  {lbl_prev} {pd_}"
        bc = ws.cell(row, 1, f"■ {etf}  ({am}){aum_txt}{cash_txt}{dates_txt}      매수 {len(buy)}종목  /  매도 {len(sell)}종목")
        bc.font = BANNERF; bc.alignment = Alignment(vertical="center")
        for c in range(1, NC + 1): ws.cell(row, c).fill = BANNERBG
        ws.row_dimensions[row].height = 19; row += 1
        if not buy and not sell:
            mergerow(row)
            nc = ws.cell(row, 1, "변화 없음 (구성종목 동일)"); nc.font = Font(color="808080", italic=True)
            row += 2; continue
        # 좌(매수 1~5) / 우(매도 7~11) 헤더
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=11)
        hl = ws.cell(row, 1, "🔴 매수 (확대·신규편입)"); hl.font = BUYHEAD; hl.alignment = Alignment(horizontal="center")
        hr = ws.cell(row, 7, "🔵 매도 (축소·편출)"); hr.font = SELLHEAD; hr.alignment = Alignment(horizontal="center")
        for c in (1, 2, 3, 4, 5): ws.cell(row, c).fill = BUYHBG
        for c in (7, 8, 9, 10, 11): ws.cell(row, c).fill = SELLHBG
        row += 1
        for c, t in zip((1, 2, 3, 4, 5, 7, 8, 9, 10, 11),
                        ("종목명", "변화(1CU주)", "전체매매금액(원)", "보유비중(전→당)", "전일→당일",
                         "종목명", "변화(1CU주)", "전체매매금액(원)", "보유비중(전→당)", "전일→당일")):
            cell = ws.cell(row, c, t); cell.font = Font(bold=True, size=9, color="595959")
            cell.alignment = Alignment(horizontal="center"); cell.border = border
        row += 1
        def put(r0, base, fnt, bg):
            nm = r0["종목명"]; pq = int(r0["전일수량"]); tq = int(r0["당일수량"])
            amt = r0.get("전체금액")
            wp = r0.get("보유비중전일"); wt = r0.get("보유비중당일")
            hold = "-" if (wp is None and wt is None) else f"{(wp or 0):.2f}%→{(wt or 0):.2f}%"
            tag = _rowtag(r0)
            ws.cell(row, base, nm + tag).font = fnt
            cc = ws.cell(row, base + 1, int(r0["변화"])); cc.number_format = NUMFMT_UP; cc.font = fnt; cc.fill = bg; cc.alignment = Alignment(horizontal="right")
            ca = ws.cell(row, base + 2, amt if amt is not None else ("검증필요" if r0.get("검증") == "불가" else "-")); ca.font = fnt; ca.fill = bg; ca.alignment = Alignment(horizontal="right")
            if amt is not None: ca.number_format = AMTFMT
            cp = ws.cell(row, base + 3, hold); cp.font = fnt; cp.fill = bg; cp.alignment = Alignment(horizontal="right")
            ws.cell(row, base + 4, f"{pq:,}→{tq:,}").alignment = Alignment(horizontal="right")
        for i in range(max(len(buy), len(sell))):
            if i < len(buy): put(buy[i], 1, RED, BUYBG)
            if i < len(sell): put(sell[i], 7, BLUE, SELLBG)
            for c in (1, 2, 3, 4, 5, 7, 8, 9, 10, 11): ws.cell(row, c).border = border
            row += 1
        row += 1

    for col, w in zip("ABCDEFGHIJK", [18, 9, 14, 11, 13, 3, 18, 9, 14, 11, 13]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"
    base_name = fname or f"pdf_change_{today.strftime('%Y%m%d')}.xlsx"
    path = os.path.join(BASE, base_name)
    # 과거에 남았을 수 있는 _HHMM 대체파일 청소(하루 1개 유지)
    import glob as _glob
    stem = base_name.rsplit(".", 1)[0]
    for old in _glob.glob(os.path.join(BASE, f"{stem}_*.xlsx")):
        try: os.remove(old)
        except OSError: pass
    try:
        wb.save(path)
        return path
    except PermissionError:
        # 파일이 Excel에 열려 잠김 → 이번 회차 저장 스킵(대체파일 만들지 않음).
        # 다음 회차(10분 뒤)에 파일이 닫혀 있으면 동일 이름으로 덮어써짐 → 하루 1개 유지.
        print(f"  [스킵] {base_name} 열려있어 저장 보류 — 파일 닫으면 다음 회차에 갱신됨")
        return path

# ── 텔레그램용 이미지 렌더 + 전송 ─────────────────────────────
def _setup_mpl_font():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.font_manager as fm
    avail = {f.name for f in fm.fontManager.ttflist}
    for f in ("Malgun Gothic", "NanumGothic", "NanumBarunGothic", "AppleGothic", "Noto Sans CJK KR"):
        if f in avail:
            matplotlib.rcParams["font.family"] = f; break
    matplotlib.rcParams["axes.unicode_minus"] = False

def _eok(x):
    """원 → '±NN.N억' 문자열(없으면 '-')."""
    return "-" if x is None else f"{x/1e8:+,.1f}억"

def render_report_image(groups, today, prev, status_line="", title=None,
                        lbl_cur="당일", lbl_prev="전영업일", out_png=None, part=None, label="코스닥 액티브"):
    """엑셀 표를 한 장의 PNG로 렌더(ETF별 매수/매도). 반환: 경로, 변화 없으면 None.
    part=(i, n) 이면 여러 장 중 i번째 — 제목에 (i/n) 표시(render_report_parts 가 나눠 부른다)."""
    _setup_mpl_font()
    import matplotlib.pyplot as plt
    from matplotlib.gridspec import GridSpec
    caps = [g for g in groups if g["state"] == "captured" and (g["buy"] or g["sell"])]
    if not caps:
        return None
    td = today.strftime("%Y-%m-%d"); pdd = prev.strftime("%Y-%m-%d")
    nrows = [max(len(g["buy"]), len(g["sell"])) for g in caps]
    ratios = [n + 1.6 for n in nrows]
    fig_h = max(3.6, 1.2 + sum(r * 0.6 for r in ratios))  # 2줄 셀 여유 + 짧은 리포트 제목겹침 방지 최소높이
    fig = plt.figure(figsize=(12.5, fig_h), dpi=170)
    gs = GridSpec(len(caps), 1, height_ratios=ratios, hspace=0.6,
                  top=1 - 1.0 / fig_h, bottom=0.2 / fig_h, left=0.012, right=0.988)  # 상단 1인치=제목 전용(겹침 방지)
    RED = "#C00000"; BLUE = "#0070C0"; RBG = "#FCE4E4"; BBG = "#E4ECF6"
    for gi, g in enumerate(caps):
        ax = fig.add_subplot(gs[gi]); ax.axis("off")
        buy = sorted(g["buy"], key=lambda r: -r["변화"]); sell = sorted(g["sell"], key=lambda r: r["변화"])
        aum = g.get("aum")
        _cl = _cashlabel(g)
        ttl = f"■ {_chunk_name(g)}  ({g['am']})" + _lagtag(g, "   ") + (f"   순자산 약 {aum/1e8:,.0f}억" if aum else "") + \
              (f"   ·   {_cl}" if _cl else "") + f"      매수 {g.get('n_buy', len(buy))}  /  매도 {g.get('n_sell', len(sell))}"
        ax.set_title(ttl, loc="left", fontsize=11, fontweight="bold", color="#1F3864", pad=6)
        n = max(len(buy), len(sell))
        cols = ["매수 종목", "변화(전→당)", "금액", "보유비중", "", "매도 종목", "변화(전→당)", "금액", "보유비중"]
        def _chcell(r):   # 주수변화 + 전→당 한 칸에 (2줄)
            return f"{int(r['변화']):+,}\n{int(r['전일수량']):,}→{int(r['당일수량']):,}"
        def _holdcell(r):  # 보유비중 전→당 (2줄)
            wp = r.get("보유비중전일"); wt = r.get("보유비중당일")
            return "-" if (wp is None and wt is None) else f"{(wp or 0):.2f}%\n→{(wt or 0):.2f}%"
        cell = []; ccol = []
        for i in range(n):
            row = [""] * 9; col = ["white"] * 9
            if i < len(buy):
                b = buy[i]
                row[0] = _namecell(b); row[1] = _chcell(b); row[2] = _amtcell(b)
                row[3] = _holdcell(b)
                for c in (0, 1, 2, 3): col[c] = RBG
            if i < len(sell):
                s = sell[i]
                row[5] = _namecell(s); row[6] = _chcell(s); row[7] = _amtcell(s)
                row[8] = _holdcell(s)
                for c in (5, 6, 7, 8): col[c] = BBG
            cell.append(row); ccol.append(col)
        tbl = ax.table(cellText=cell, colLabels=cols, cellColours=ccol,
                       colColours=[RED, RED, RED, RED, "white", BLUE, BLUE, BLUE, BLUE],
                       cellLoc="center", loc="upper center",
                       colWidths=[0.19, 0.12, 0.095, 0.085, 0.02, 0.19, 0.12, 0.095, 0.085])
        tbl.auto_set_font_size(False); tbl.set_fontsize(10); tbl.scale(1, 2.2)
        for (r0, c0), cobj in tbl.get_celld().items():
            t = cobj.get_text()
            if r0 == 0:
                if c0 != 4: t.set_color("white"); t.set_fontweight("bold"); t.set_fontsize(10)
            elif c0 in (0, 2, 5, 7):   # 종목명·금액 강조(굵게)
                t.set_fontweight("bold")
            if r0 > 0 and c0 in (0, 5):   # 종목명 길면 자동 축소(칸 넘침 방지)
                s = t.get_text()
                if s: t.set_fontsize(max(6.5, min(10, 10 * 13 / len(s))))
    sub = f"\n{status_line}" if status_line else ""
    sup = title or f"{label} ETF PDF 변화   ·   {lbl_cur} {td}  vs  {lbl_prev} {pdd}"
    if part: sup += f"   ({part[0]}/{part[1]})"
    fig.suptitle(sup + sub, fontsize=13, fontweight="bold", y=1 - 0.12 / fig_h, va="top")
    out_png = out_png or os.path.join(BASE, f"pdf_change_{today.strftime('%Y%m%d')}.png")
    fig.savefig(out_png, bbox_inches="tight", facecolor="white"); plt.close(fig)
    return out_png

def render_report_image_mobile(groups, today, prev, status_line="", title=None,
                               lbl_cur="당일", lbl_prev="전영업일", out_png=None, part=None, totals=None,
                               label="코스닥 액티브"):
    """모바일 세로 표형(무채색). ETF명 → ▼매수 → ▼매도, 정보는 칸(셀)으로 구분.
    글자 크게 + 양옆 여백 최소. part=(i, n)·totals=(ETF수, 매수, 매도) 는 여러 장으로 나눌 때 —
    제목에 (i/n), 요약줄은 이 장이 아니라 전체 합계."""
    _setup_mpl_font()
    import matplotlib.pyplot as plt
    from matplotlib.gridspec import GridSpec
    caps = [g for g in groups if g["state"] == "captured" and (g["buy"] or g["sell"])]
    if not caps:
        return None
    td = today.strftime("%Y-%m-%d"); pdd = prev.strftime("%Y-%m-%d")
    def _chc(r): return f"{int(r['변화']):+,}\n{int(r['전일수량']):,}→{int(r['당일수량']):,}"  # 2줄(변화/전→당)
    def _eok(r):
        a = r.get("전체금액"); return "-" if a is None else f"{a/1e8:+,.1f}억"
    def _hold(r):   # 보유비중 전→당 (2줄)
        wp = r.get("보유비중전일"); wt = r.get("보유비중당일")
        if wp is None and wt is None: return "-"
        return f"{(wp or 0):.2f}%\n→{(wt or 0):.2f}%"
    # 연속 표(단일) 구성 → ETF 사이 빈 공간 제거
    rows = []; styles = []; banners = {}
    rows.append(["종목", "변화(전→당)", "금액", "보유비중"]); styles.append("hdr")
    tot_b = tot_s = 0
    for gi, g in enumerate(caps):
        buy = sorted(g["buy"], key=lambda r: -r["변화"]); sell = sorted(g["sell"], key=lambda r: r["변화"])
        tot_b += len(buy); tot_s += len(sell)
        if gi > 0: rows.append(["", "", "", ""]); styles.append("gap")
        aum = g.get("aum")
        _cl = _cashlabel(g)
        banners[len(rows)] = f"■ {_chunk_name(g)} ({g['am']})" + _lagtag(g, "  ") + (f"  ·  {aum/1e8:,.0f}억" if aum else "") + \
                             (f"  ·  {_cl}" if _cl else "")
        rows.append(["", "", "", ""]); styles.append("etf")   # 배너는 표 위에 겹쳐 그림(잘림 방지)
        show_b, show_s = _mobile_sections(g)                 # 여러 장에 걸친 ETF 조각은 행이 있는 섹션만
        if show_b:
            rows.append([f"▼ 매수 {g.get('n_buy', len(buy))}" + (" (계속)" if g.get("cont_buy") else ""), "", "", ""])
            styles.append("sect")
            for b in (buy or [None]):
                rows.append(["(없음)", "", "", ""] if b is None else
                            [_namecell(b, short=True), _chc(b), _amtcell(b, short=True), _hold(b)])
                styles.append("data")
        if show_s:
            rows.append([f"▼ 매도 {g.get('n_sell', len(sell))}" + (" (계속)" if g.get("cont_sell") else ""), "", "", ""])
            styles.append("sect")
            for s in (sell or [None]):
                rows.append(["(없음)", "", "", ""] if s is None else
                            [_namecell(s, short=True), _chc(s), _amtcell(s, short=True), _hold(s)])
                styles.append("data")
    HU = {"hdr": 1.0, "etf": 1.25, "sect": 1.0, "data": 1.75, "gap": 1.4}   # 행 높이(상대)
    units = sum(HU[s] for s in styles)
    title_in = 1.05
    ax_in = units * 0.235
    fig_h = title_in + ax_in
    fig = plt.figure(figsize=(5.2, fig_h), dpi=205)     # 폭 좁게
    ax = fig.add_axes([0.008, 0.006, 0.984, ax_in / fig_h]); ax.axis("off")
    GRID = "#8A8A8A"; SECT = "#DCDCDC"; HDR = "#C2C2C2"; ETFBG = "#B9B9B9"
    tbl = ax.table(cellText=rows, cellLoc="left", colWidths=[0.36, 0.25, 0.21, 0.18], bbox=[0, 0, 1, 1])
    tbl.auto_set_font_size(False); tbl.set_fontsize(12)
    for (r0, c0), cell in tbl.get_celld().items():
        st = styles[r0]; cell.set_height(HU[st]); t = cell.get_text()
        cell.set_edgecolor(GRID); cell.set_linewidth(0.6)
        if st == "hdr":
            cell.set_facecolor(HDR); t.set_fontweight("bold"); t.set_fontsize(10)
        elif st == "etf":
            cell.set_facecolor(ETFBG); cell.set_edgecolor(ETFBG)
        elif st == "sect":
            cell.set_facecolor(SECT); t.set_fontweight("bold")
        elif st == "gap":
            cell.set_facecolor("white"); cell.set_linewidth(0)  # 선 자체를 안 그림(위 표 밑줄 보존)
        else:
            cell.set_facecolor("white")
            if c0 == 0:
                s0 = t.get_text()
                if len(s0) > 9: t.set_fontsize(max(8, min(12, 12 * 9 / len(s0))))
                else: t.set_fontweight("bold")
            elif c0 in (2, 3):   # 금액·비중 굵게
                t.set_fontweight("bold")
    # ETF 배너를 표 위에 겹쳐 그림(칸 폭 무시, 잘림 방지)
    cum = 0.0
    for i, st in enumerate(styles):
        if st == "etf":
            cy = 1 - (cum + 0.5 * HU[st]) / units
            ax.text(0.012, cy, banners[i], transform=ax.transAxes, ha="left", va="center",
                    fontweight="bold", fontsize=11.5, clip_on=False, zorder=5)
        cum += HU[st]
    n = len(caps)
    if totals: n, tot_b, tot_s = totals
    summary = f"변화 {n}개 ETF  ·  매수 {tot_b} / 매도 {tot_s} 종목"
    head = f"{label} ETF PDF 변화" + (f"  ({part[0]}/{part[1]})" if part else "")
    sup = title or f"{head}\n{lbl_cur} {td}  vs  {lbl_prev} {pdd}"
    fig.suptitle(f"{sup}\n{summary}", fontsize=11.5, fontweight="bold", y=1 - 0.12 / fig_h, va="top")
    out_png = out_png or os.path.join(BASE, f"pdf_change_m_{today.strftime('%Y%m%d')}.png")
    fig.savefig(out_png, bbox_inches="tight", facecolor="white"); plt.close(fig)
    return out_png

# ── 길면 여러 장으로 ─────────────────────────────────────────────
# 텔레그램 사진은 가로+세로 ≤ 10000px(넘으면 PHOTO_INVALID_DIMENSIONS → 파일로 감)이고, 길수록 축소돼 글자가 작아진다.
# 2026-10-26(10종·94행) 모바일 1089×9618px가 파일로, 가로형 2108×7303px는 작게 나갔다.
# → 한 장 세로가 MAX_PART_PX를 넘으면 나눠 앨범으로 보낸다. 작은 ETF끼리는 한 장에 모으고,
#   한 장보다 큰 ETF는 행 단위로 고르게 쪼갠다(이름 뒤 (1/2), 매수·매도 건수는 전체 그대로).
# 상한은 PC에서도 편하게(2026-10-26 사용자 요청): 텔레그램 PC 뷰어는 사진을 창 높이(1080p 기준 약 950px)에 맞춰
# 줄이므로 세로로 긴 장은 글자가 작아진다(가로형 2107×4447px → 약 1/4). 가로형은 세로 ≤1600(화면에 약 0.6배 이상),
# 모바일형은 세로 ≤2100(글자가 커서 0.45배여도 읽힘, 폰에선 화면 폭을 꽉 채움). 긴 변 ≤ 2560이라 해상도도 안 깎인다.
MAX_PART_PX = {"wide": 1600, "mobile": 2100}
# 2026-10-29 사용자: 4개 그룹으로 나누면서 '아침에 사진 4장 하나씩' → 그룹당 1장이 기본. PC용 상한(위)은 끄고,
# 텔레그램 사진 한계(가로+세로 ≤ 10000px — 넘으면 파일로 감)에 걸릴 때만 나눈다(여유 두고 가로형 7000·모바일형 8000).
MAX_PART_PX = {"wide": 7000, "mobile": 8000}

def _chunk_name(g):
    """여러 장에 걸친 ETF 조각이면 이름 뒤에 (i/n)."""
    c = g.get("chunk")
    return f"{g['etf']} ({c[0]}/{c[1]})" if c else g["etf"]

def _mobile_sections(g):
    """모바일 표에서 그릴 섹션 (매수, 매도). 조각이 아니면 둘 다(빈 쪽은 '(없음)').
    조각이면 행이 있는 섹션만 — 단 매수가 없는 ETF는 첫 조각에, 매도가 없는 ETF는 마지막 조각에 '(없음)'."""
    c = g.get("chunk")
    if not c:
        return True, True
    nb, ns = g.get("n_buy", len(g["buy"])), g.get("n_sell", len(g["sell"]))
    return bool(g["buy"]) or (nb == 0 and c[0] == 1), bool(g["sell"]) or (ns == 0 and c[0] == c[1])

def _est_px_wide(gs):
    """render_report_image 의 ETF 표 높이(px) 추정: (행수+1.6)×0.6in × 170dpi."""
    return sum((max(len(g["buy"]), len(g["sell"])) + 1.6) * 0.6 * 170 for g in gs)

def _est_px_mobile(gs):
    """render_report_image_mobile 의 표 높이(px) 추정: 행 단위(HU) × 0.235in × 205dpi."""
    u = 0.0
    for i, g in enumerate(gs):
        sb, ss = _mobile_sections(g)
        u += (1.4 if i else 0) + 1.25 + ((1.0 + 1.75 * max(1, len(g["buy"]))) if sb else 0) \
             + ((1.0 + 1.75 * max(1, len(g["sell"]))) if ss else 0)
    return u * 0.235 * 205

def _chunk_oversized(caps, kind, est, base_px, max_px):
    """한 장 상한보다 큰 ETF를 행 단위로 고르게 쪼갠 조각들로 바꾼다(나머지는 그대로)."""
    out = []
    for g in caps:
        if base_px + est([g]) <= max_px:
            out.append(g); continue
        buy, sell = g["buy"], g["sell"]
        if kind == "wide":      # 매수·매도가 좌우로 나란히 → 같은 행 범위로 자른다
            n = max(len(buy), len(sell))
            per = max(1, int((max_px - base_px) / (0.6 * 170) - 1.6))
            c = -(-n // per); size = -(-n // c)
            pieces = [(buy[i * size:(i + 1) * size], sell[i * size:(i + 1) * size]) for i in range(c)]
        else:                   # 매수 뒤에 매도가 이어지는 세로 표 → 이어 붙인 행 목록을 자른다
            seq = [("b", r) for r in buy] + [("s", r) for r in sell]
            per = max(1, int(((max_px - base_px) / (0.235 * 205) - 1.25 - 2.0) / 1.75))
            c = -(-len(seq) // per); size = -(-len(seq) // c)
            pieces = []
            for i in range(c):
                seg = seq[i * size:(i + 1) * size]
                pieces.append(([r for t, r in seg if t == "b"], [r for t, r in seg if t == "s"]))
        seen_b = seen_s = 0
        for i, (b, s) in enumerate(pieces, 1):
            out.append(dict(g, buy=b, sell=s, n_buy=len(buy), n_sell=len(sell), chunk=(i, len(pieces)),
                            cont_buy=bool(b) and seen_b > 0, cont_sell=bool(s) and seen_s > 0))
            seen_b += len(b); seen_s += len(s)
    return out

def _split_caps(caps, est, base_px, max_px):
    """장 수는 최소로, 장마다 높이는 고르게. 다음 ETF를 넣으면 목표 높이를 절반 넘게 넘길 때 새 장을 시작한다.
    상한을 넘는 장은 ETF 하나짜리일 때만 허용(한 ETF는 쪼개지 않으므로)."""
    if base_px + est(caps) <= max_px or len(caps) < 2:
        return [caps]
    k = 2
    while True:
        target = est(caps) / k
        parts, cur = [], []
        for g in caps:
            if cur and len(parts) < k - 1 and est(cur + [g]) - est([g]) / 2 > target:
                parts.append(cur); cur = []
            cur.append(g)
        parts.append(cur)
        if all(base_px + est(p) <= max_px or len(p) == 1 for p in parts) or k >= len(caps):
            return parts
        k += 1

def render_report_parts(kind, groups, today, prev, status_line="", label="코스닥 액티브", tag=""):
    """kind='wide'(채널1) | 'mobile'(채널2) → [PNG 경로...]. 짧으면 1장, 길면 _1·_2… 여러 장. 변화 없으면 [].
    label 은 제목(예: '반도체 액티브'), tag 는 파일명 구분(그룹 키 — 그룹별 이미지가 서로 덮어쓰지 않게)."""
    caps = [g for g in groups if g["state"] == "captured" and (g["buy"] or g["sell"])]
    if not caps:
        return []
    if kind == "wide":
        render, est, base, stem = render_report_image, _est_px_wide, 330, "pdf_change"
    else:
        render, est, base, stem = render_report_image_mobile, _est_px_mobile, 330, "pdf_change_m"
    ymd = today.strftime("%Y%m%d") + (f"_{tag}" if tag else "")
    pieces = _chunk_oversized(caps, kind, est, base, MAX_PART_PX[kind])
    parts = _split_caps(pieces, est, base, MAX_PART_PX[kind])
    if len(parts) == 1:
        return [render(groups, today, prev, status_line, label=label, out_png=os.path.join(BASE, f"{stem}_{ymd}.png"))]
    totals = (len(caps), sum(len(g["buy"]) for g in caps), sum(len(g["sell"]) for g in caps))
    out = []
    for i, p in enumerate(parts, 1):
        kw = {"part": (i, len(parts)), "label": label, "out_png": os.path.join(BASE, f"{stem}_{ymd}_{i}.png")}
        if kind != "wide": kw["totals"] = totals
        out.append(render(p, today, prev, status_line, **kw))
    return out

def send_text(text, token=None, chat=None):
    """짧은 텍스트 메시지(변화 없는 그룹 알림 등)."""
    tok = token or os.environ.get("TELEGRAM_TOKEN"); chat = chat or os.environ.get("TELEGRAM_CHAT_ID")
    if not (tok and chat):
        print("  [텔레그램] 토큰/챗ID 미설정 — 전송 스킵"); return False
    try:
        r = requests.post(f"https://api.telegram.org/bot{tok}/sendMessage",
                          data={"chat_id": chat, "text": text[:4000]}, timeout=60)
        if r.ok and r.json().get("ok"):
            print("  [텔레그램] 텍스트 전송 완료"); return True
        print(f"  [텔레그램] 텍스트 실패: {r.text[:150]}")
    except Exception as e:
        print(f"  [텔레그램] 텍스트 오류: {type(e).__name__}: {e}")
    return False

def _send_album(paths, caption="", token=None, chat=None):
    """여러 장을 앨범(sendMediaGroup, 최대 10장씩)으로. 캡션은 첫 장. 앨범이 거부되면 한 장씩(사진→파일).
    한 장이라도 나갔으면 True — 발송 마커를 남겨 다음 회차가 같은 날 다시 보내지 않게 한다."""
    tok = token or os.environ.get("TELEGRAM_TOKEN"); chat = chat or os.environ.get("TELEGRAM_CHAT_ID")
    if not (tok and chat):
        print("  [텔레그램] 토큰/챗ID 미설정 — 전송 스킵"); return False
    sent = False
    for s in range(0, len(paths), 10):
        chunk = paths[s:s + 10]
        media = [dict({"type": "photo", "media": f"attach://p{i}"},
                      **({"caption": caption[:1024]} if (s == 0 and i == 0 and caption) else {}))
                 for i in range(len(chunk))]
        files = {}; why = ""
        try:
            for i, p in enumerate(chunk): files[f"p{i}"] = open(p, "rb")
            r = requests.post(f"https://api.telegram.org/bot{tok}/sendMediaGroup",
                              data={"chat_id": chat, "media": json.dumps(media)}, files=files, timeout=120)
            if r.ok and r.json().get("ok"):
                print(f"  [텔레그램] 앨범 {len(chunk)}장 전송 완료"); sent = True; continue
            why = r.text[:150]
        except Exception as e:
            why = f"{type(e).__name__}: {e}"
        finally:
            for f in files.values(): f.close()
        print(f"  [텔레그램] 앨범 실패: {why} → 한 장씩")
        for i, p in enumerate(chunk):
            if send_telegram(p, caption if (s == 0 and i == 0) else "", token=tok, chat=chat): sent = True
    return sent

def send_telegram(png_path, caption="", token=None, chat=None):
    """이미지 전송(sendPhoto→실패 시 sendDocument). token/chat 미지정 시 환경변수 사용.
    경로 목록을 주면 2장 이상은 앨범으로(_send_album)."""
    if isinstance(png_path, (list, tuple)):
        paths = [p for p in png_path if p and os.path.exists(p)]
        if len(paths) > 1: return _send_album(paths, caption, token, chat)
        png_path = paths[0] if paths else None
    tok = token or os.environ.get("TELEGRAM_TOKEN"); chat = chat or os.environ.get("TELEGRAM_CHAT_ID")
    if not (tok and chat):
        print("  [텔레그램] 토큰/챗ID 미설정 — 전송 스킵"); return False
    if not (png_path and os.path.exists(png_path)):
        print("  [텔레그램] 보낼 이미지 없음"); return False
    for method, field in (("sendPhoto", "photo"), ("sendDocument", "document")):
        try:
            with open(png_path, "rb") as f:
                r = requests.post(f"https://api.telegram.org/bot{tok}/{method}",
                                  data={"chat_id": chat, "caption": caption[:1024]},
                                  files={field: f}, timeout=60)
            if r.ok and r.json().get("ok"):
                print(f"  [텔레그램] {method} 전송 완료"); return True
            print(f"  [텔레그램] {method} 실패: {r.text[:150]}")
        except Exception as e:
            print(f"  [텔레그램] {method} 오류: {type(e).__name__}: {e}")
    return False


if __name__ == "__main__":
    args = sys.argv[1:]
    if args and args[0] == "--weekly":
        d = datetime.strptime(args[1], "%Y-%m-%d").date() if len(args) > 1 else None
        weekly_report(d)
    elif args and args[0] == "--days":   # 최근 N거래일 롤링: --days 5 [YYYY-MM-DD]
        n = int(args[1]) if len(args) > 1 else 5
        d = datetime.strptime(args[2], "%Y-%m-%d").date() if len(args) > 2 else None
        rolling_report(d, n)
    else:
        d = datetime.strptime(args[0], "%Y-%m-%d").date() if args else None
        run(d)
