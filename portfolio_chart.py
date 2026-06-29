# -*- coding: utf-8 -*-
"""
ETF 구성종목 포트폴리오 차트: 비중 원그래프 + 평가액(종가 기준) 막대그래프.
PDF의 평가금액(=수량×종가, 1CU 기준)·비중을 그대로 사용. 7종을 한 장(PNG)에.
실행: python portfolio_chart.py [YYYY-MM-DD]
"""
import sys, io, json, requests
from datetime import datetime
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
# 한글 폰트 적응: Windows(Malgun Gothic) / Linux 클라우드(NanumGothic) 자동 선택
_avail = {f.name for f in fm.fontManager.ttflist}
for _f in ("Malgun Gothic", "NanumGothic", "NanumBarunGothic", "AppleGothic", "Noto Sans CJK KR"):
    if _f in _avail:
        matplotlib.rcParams["font.family"] = _f; break
matplotlib.rcParams["axes.unicode_minus"] = False

import pdf_change as P  # is_trading_day, prev_trading_day, UA, SKIP, BASE

UA = P.UA; SKIP = P.SKIP; BASE = P.BASE


# ── 풀데이터 fetcher: [(종목명, 평가금액, 비중)] ───────────────
def full_koact(pid, ymd):
    g = f"{ymd[:4]}.{ymd[4:6]}.{ymd[6:]}"
    d = requests.get(f"https://www.samsungactive.co.kr/api/v1/product/etf-pdf/{pid}.do",
                     params={"gijunYMD": g}, headers={"User-Agent": UA,
                     "Referer": f"https://www.samsungactive.co.kr/etf/view.do?id={pid}"}, timeout=20).json().get("pdf", {})
    out = []
    for it in d.get("list", []):
        nm = (it.get("secNm") or "").strip()
        if not nm or any(s in nm for s in SKIP): continue
        out.append((nm, float(it.get("evalA", 0) or 0), float(it.get("ratio", 0) or 0)))
    return out

def _xlsx_rows(content):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    return list(wb.active.iter_rows(values_only=True))

def full_time(idx, ymd):
    d = f"{ymd[:4]}-{ymd[4:6]}-{ymd[6:]}"
    r = requests.get("https://timeetf.co.kr/pdf_excel.php", params={"idx": idx, "pdfDate": d},
                     headers={"User-Agent": UA, "Referer": f"https://timeetf.co.kr/m11_view.php?idx={idx}"}, timeout=20)
    rows = _xlsx_rows(r.content); out = []
    if not rows: return out
    h = [str(c).strip() if c else "" for c in rows[0]]
    try: i_nm, i_ev, i_wt = h.index("종목명"), next(j for j,c in enumerate(h) if "평가금액" in c), next(j for j,c in enumerate(h) if "비중" in c)
    except (ValueError, StopIteration): return out
    for row in rows[1:]:
        if not row or len(row) <= max(i_nm,i_ev,i_wt): continue
        nm = str(row[i_nm]).strip() if row[i_nm] else ""
        if not nm or any(s in nm for s in SKIP): continue
        try: out.append((nm, float(str(row[i_ev]).replace(",","")), float(str(row[i_wt]).replace(",",""))))
        except (ValueError, TypeError): pass
    return out

def full_plus(n, ymd):
    # PLUS xlsx엔 평가금액 없음 → 비중만, 평가액은 비중으로 대체(상대크기 동일)
    r = requests.get("https://www.plusetf.co.kr/excel/product/pdf", params={"n": n, "d": ymd, "title": "x"},
                     headers={"User-Agent": UA, "Referer": f"https://www.plusetf.co.kr/product/detail?n={n}"}, timeout=20)
    rows = _xlsx_rows(r.content); out = []; hdr = None
    for i, row in enumerate(rows):
        cells = [str(c).strip() if c is not None else "" for c in row]
        if "종목명" in cells and any("비중" in c for c in cells):
            i_nm = cells.index("종목명"); i_wt = next(j for j,c in enumerate(cells) if "비중" in c); hdr = i; break
    if hdr is None: return out
    for row in rows[hdr+1:]:
        if not row or len(row) <= max(i_nm,i_wt): continue
        nm = str(row[i_nm]).strip() if row[i_nm] else ""
        if not nm or nm=="nan" or any(s in nm for s in SKIP): continue
        try:
            wt = float(str(row[i_wt]).replace(",",""))
            if wt < 1: wt *= 100  # 0.06 → 6%
            out.append((nm, wt, wt))  # 평가액 자리에 비중(상대) 사용
        except (ValueError, TypeError): pass
    return out

def full_rise(stid, ymd):
    import pandas as pd
    d = f"{ymd[:4]}-{ymd[4:6]}-{ymd[6:]}"
    r = requests.get("https://www.riseetf.co.kr/prod/finder/productViewTabExcel3",
                     params={"searchTargetId": stid, "searchDate": d}, headers={"User-Agent": UA, "Accept-Language":"ko-KR",
                     "Referer": f"https://www.riseetf.co.kr/prod/finderDetail/{stid}?searchFlag=viewtab2"}, timeout=20)
    out = []
    try: t = pd.read_html(io.StringIO(r.text))[0]
    except Exception: return out
    hdr = None
    for i in range(len(t)):
        v = [str(x).strip() for x in t.iloc[i].tolist()]
        if "종목명" in v and any("수량" in x for x in v):
            i_nm=v.index("종목명"); i_wt=next(j for j,x in enumerate(v) if "비중" in x); i_ev=next((j for j,x in enumerate(v) if "평가" in x), None); hdr=i; break
    if hdr is None: return out
    for i in range(hdr+1, len(t)):
        v = [str(x).strip() for x in t.iloc[i].tolist()]
        if len(v) <= i_wt: continue
        nm = v[i_nm]
        if not nm or nm=="nan" or any(s in nm for s in SKIP): continue
        try:
            wt = float(v[i_wt].replace(",",""))
            ev = float(v[i_ev].replace(",","")) if i_ev is not None else wt
            out.append((nm, ev, wt))
        except ValueError: pass
    return out

def full_tiger(isin, ymd):
    import pandas as pd
    d = f"{ymd[:4]}.{ymd[4:6]}.{ymd[6:]}"
    base = "https://investments.miraeasset.com/tigeretf/ko/product/search/detail"
    s = requests.Session(); s.headers.update({"User-Agent": UA, "Accept-Language": "ko-KR"})
    s.get(f"{base}/index.do", params={"ksdFund": isin}, timeout=20)
    r = s.post(f"{base}/pdfListAjax.ajax", data={"ksdFund": isin,"pageIndex":1,"firstIndex":0,"listCnt":300,"fixDate":d,"prfPrd":"Week01","order":"SRD"},
               headers={"X-Requested-With":"XMLHttpRequest","Referer": f"{base}/index.do?ksdFund={isin}"}, timeout=20)
    out = []
    try: df = pd.read_html(io.StringIO("<table>"+r.text+"</table>"))[0]
    except Exception: return out
    for _, row in df.iterrows():  # col1=종목명 col3=평가금액 col4=비중
        nm = str(row[1]).strip()
        if not nm or nm=="nan" or any(s in nm for s in SKIP): continue
        try: out.append((nm, float(str(row[3]).replace(",","")), float(str(row[4]).replace(",",""))))
        except (ValueError, TypeError): pass
    return out

CHARTS = [
    ("KoAct 코스닥액티브", full_koact, "2ETFU6", True),
    ("KoAct 바이오헬스케어", full_koact, "2ETFJ9", True),
    ("TIME 코스닥액티브", full_time, "24", True),
    ("TIME K바이오액티브", full_time, "13", True),
    ("PLUS 코스닥150액티브", full_plus, "006399", False),  # 평가액 미제공 → 비중축
    ("RISE 바이오TOP10액티브", full_rise, "44I0", True),
    ("TIGER 기술이전바이오", full_tiger, "KR70168K0008", True),
]

N_DAYS = 12   # 추이 거래일 수
TOP_LINES = 8  # 선으로 그릴 상위 종목 수

def trading_days(today, n):
    days = []; d = today
    while len(days) < n:
        if P.is_trading_day(d): days.append(d)
        d = P.prev_trading_day(d)
    return sorted(days)  # 과거→현재

def build_series(fn, eid, dates):
    """{date: {종목명: 비중}} — 데이터 있는 날짜만."""
    series = {}
    for d in dates:
        try:
            rows = fn(eid, d.strftime("%Y%m%d"))
        except Exception:
            rows = []
        if rows:
            series[d] = {nm: wt for nm, ev, wt in rows}
    return series

def generate(today=None):
    today = today or datetime.now().date()
    while not P.is_trading_day(today):
        today = P.prev_trading_day(today)
    dates = trading_days(today, N_DAYS)

    # ETF별 시계열 병렬 수집
    from concurrent.futures import ThreadPoolExecutor
    def work(item):
        name, fn, eid, _ = item
        return name, build_series(fn, eid, dates)
    with ThreadPoolExecutor(max_workers=7) as ex:
        results = dict(ex.map(work, CHARTS))

    n = len(CHARTS)
    fig, axes = plt.subplots(n, 2, figsize=(16, 3.6*n), gridspec_kw={"width_ratios":[1, 1.5]})
    fig.suptitle(f"코스닥 액티브 ETF 포트폴리오   ·   {today.strftime('%Y-%m-%d')}   "
                 f"(좌: 비중 원그래프 / 우: 상위 {TOP_LINES}종목 비중 추이 {len(dates)}거래일)",
                 fontsize=15, fontweight="bold", y=0.998)
    cmap = plt.get_cmap("tab20"); lcmap = plt.get_cmap("tab10")
    for r, (name, fn, eid, _) in enumerate(CHARTS):
        series = results.get(name, {})
        ax_pie, ax_line = axes[r]
        sdates = sorted(series.keys())
        if not sdates:
            ax_pie.text(0.5,0.5,f"{name}\n데이터 없음", ha="center", va="center"); ax_pie.axis("off"); ax_line.axis("off"); continue
        latest = sdates[-1]; wmap = series[latest]
        datelbl = "" if latest==today else f"  ({latest.strftime('%m-%d')})"
        ranked = sorted(wmap.items(), key=lambda x: -x[1])
        # 원그래프: 최신일 비중 top12 + 기타
        top = ranked[:12]; etc = sum(w for _, w in ranked[12:])
        labels = [k for k,_ in top] + (["기타"] if etc>0.01 else [])
        sizes  = [w for _,w in top] + ([etc] if etc>0.01 else [])
        colors = [cmap(i%20) for i in range(len(top))] + (["#cccccc"] if etc>0.01 else [])
        ax_pie.pie(sizes, labels=labels, colors=colors, autopct=lambda p: f"{p:.0f}%" if p>=4 else "",
                   startangle=90, counterclock=False, textprops={"fontsize":8}, pctdistance=0.78)
        ax_pie.set_title(f"{name}{datelbl}  비중", fontsize=11, fontweight="bold")
        # 선그래프: 상위 TOP_LINES 종목의 비중 추이
        names = [k for k,_ in ranked[:TOP_LINES]]
        xs = list(range(len(sdates)))
        for i, nm in enumerate(names):
            ys = [series[d].get(nm, 0.0) for d in sdates]  # 미보유=0(편출시 선이 0으로 내려감)
            ax_line.plot(xs, ys, marker="o", ms=3, lw=1.6, color=lcmap(i%10), label=nm)
        ax_line.set_xticks(xs); ax_line.set_xticklabels([d.strftime("%m-%d") for d in sdates], fontsize=7, rotation=45)
        ax_line.set_ylabel("비중(%)", fontsize=8); ax_line.grid(True, alpha=0.3, lw=0.5)
        ax_line.set_title(f"{name}{datelbl}  상위{TOP_LINES} 비중 추이", fontsize=11, fontweight="bold")
        ax_line.legend(fontsize=6.5, ncol=2, loc="upper left", framealpha=0.85)
    plt.tight_layout(rect=[0,0,1,0.996])
    path = f"{BASE}/portfolio_{today.strftime('%Y%m%d')}.png"
    fig.savefig(path, dpi=110, bbox_inches="tight")
    plt.close(fig)
    print(f"[저장] {path}")
    return path

if __name__ == "__main__":
    generate(datetime.strptime(sys.argv[1], "%Y-%m-%d").date() if len(sys.argv) > 1 else None)
