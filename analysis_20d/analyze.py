# -*- coding: utf-8 -*-
"""최근 20영업일 분석 — 그래프 3종 + 엑셀.
 ① 코스닥 ETF 자금유출입(설정/환매) + AUM  (패시브 vs 액티브)
 ② 액티브 섹터별 일간 순매수금액
 ③ 액티브 종목별 일간 순매수금액
순매수금액 = 1CU수량변화 × 종가 × 전체CU수 (설정/환매 효과 제외한 순수 리밸런싱)
"""
import os, sys, json, datetime
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import pdf_change as P
from collect import trading_days

EOK = 1e8   # 억

def load(name):
    with open(os.path.join(HERE, name), encoding="utf-8") as f: return json.load(f)

# ── 데이터 로드 ────────────────────────────────────────────────
krx = load("krx_etf.json")        # {ymd: {code: {nm,shrs,nav,close,netasset,type}}}
pdf = load("pdf_raw.json")        # {etf_id: {ymd: {종목:[수량,종가]}}}
sec = load("sectors.json")        # {종목: {code,sector,market}}
px  = load("krx_px.json")         # {ymd: {종목: 종가}}
DAYS = [d.strftime("%Y%m%d") for d in trading_days() if d.strftime("%Y%m%d") in krx]
from universe import ACTIVE_PDF, is_kosdaq_passive, is_kosdaq_active
from themes import theme_of

def classify(code, nm):
    """① 대상 여부와 그룹. 코스닥 패시브 / 코스닥액티브(바이오·커버드콜 제외)."""
    if is_kosdaq_active(code, nm): return "액티브"
    if is_kosdaq_passive(nm): return "패시브"
    return None

# ── ① 자금유출입 + AUM ────────────────────────────────────────
def flows_aum():
    """returns rows[(ymd, type, flow원, aum원)] 집계 + ETF별 상세"""
    agg = []           # 일자별 그룹 집계
    detail = []        # ETF별 상세(엑셀용)
    for i in range(1, len(DAYS)):
        t, p = DAYS[i], DAYS[i - 1]
        g = defaultdict(lambda: [0.0, 0.0])   # type -> [flow, aum]
        for code, it in krx[t].items():
            typ = classify(code, it["nm"])
            if not typ: continue
            # KRX는 최신일 순자산총액을 0으로 주는 경우가 있음 → 좌수×NAV로 보정
            aum = it["netasset"] or (it["shrs"] * it["nav"])
            prev = krx[p].get(code)
            flow = (it["shrs"] - prev["shrs"]) * it["nav"] if (prev and it["nav"]) else 0.0
            g[typ][0] += flow; g[typ][1] += aum
            detail.append({"일자": t, "구분": typ, "ETF": it["nm"], "코드": code,
                           "좌수": it["shrs"], "좌수변화": (it["shrs"] - prev["shrs"]) if prev else 0,
                           "NAV": it["nav"], "순유출입(원)": flow, "순자산(원)": aum,
                           "순자산출처": "KRX" if it["netasset"] else "좌수xNAV(보정)"})
        for typ, (f, a) in g.items():
            agg.append({"일자": t, "구분": typ, "순유출입(원)": f, "순자산(원)": a})
    return agg, detail

# ── ②③ 순매수금액 ────────────────────────────────────────────
def net_buys():
    """returns rows[{일자, ETF, 종목, 섹터, 순매수금액}] + 결측 리포트"""
    rows = []; skipped = []
    byid = {e["id"]: e for e in ACTIVE_PDF}
    for eid, e in byid.items():
        snaps = pdf.get(eid, {})
        for i in range(1, len(DAYS)):
            t, p = DAYS[i], DAYS[i - 1]
            mt, mp = snaps.get(t), snaps.get(p)
            if not mt or not mp:
                skipped.append((e["name"], t, "PDF없음")); continue
            kinfo = krx[p].get(e["krx"])          # PDF(t)는 p일 종가/좌수 기준
            if not kinfo or not kinfo["shrs"]:
                skipped.append((e["name"], t, "좌수없음")); continue
            ncu = kinfo["shrs"] / e["cu"]
            names = set(mt) | set(mp)
            for nm in names:
                qt, pxt = (mt.get(nm) or [0.0, None])[:2]
                qp, pxp = (mp.get(nm) or [0.0, None])[:2]
                dq = float(qt) - float(qp)
                if dq == 0: continue
                price = pxt or pxp or px.get(p, {}).get(nm)
                if not price:
                    skipped.append((e["name"], t, f"종가없음:{nm}")); continue
                sc = (sec.get(nm, {}) or {}).get("sector") or "미분류"
                rows.append({"일자": t, "ETF": e["name"], "종목": nm,
                             "섹터": sc, "테마": theme_of(nm, sc),
                             "수량변화(1CU)": dq, "종가": price, "전체CU": ncu,
                             "순매수금액(원)": dq * price * ncu})
    return rows, skipped

def theme_weights():
    """일자별 테마 보유비중(%). returns (DAYS, [{테마: 비중%}, ...], [{테마: 보유액원}, ...])"""
    ws = []; hs = []
    for i, day in enumerate(DAYS):
        prev = DAYS[i - 1] if i > 0 else day       # PDF(t)는 t-1 종가·좌수 기준
        h = defaultdict(float)
        for e in ACTIVE_PDF:
            m = pdf.get(e["id"], {}).get(day); k = krx[prev].get(e["krx"])
            if not m or not k or not k["shrs"]: continue
            ncu = k["shrs"] / e["cu"]
            for nm, v in m.items():
                q, pr = P._qp(v)
                pr = pr or px.get(prev, {}).get(nm)
                if not pr: continue
                sc = (sec.get(nm, {}) or {}).get("sector") or "미분류"
                h[theme_of(nm, sc)] += q * pr * ncu
        t = sum(h.values())
        ws.append({k: v / t * 100 for k, v in h.items()} if t else {})
        hs.append(dict(h))
    return DAYS, ws, hs

# ══ 2026-11-02 주간 리포트 4그룹화 ═══════════════════════════════
# 액티브 PDF 이력은 일일 리포트 스냅샷(snapshots.json, 매일 당일분)에서 읽는다 — 클라우드에서 운용사가 막힌
# TIGER·RISE·MIDAS 도 날짜별 당일분이 쌓여 있어서, 운용사에서 과거분을 다시 받을 필요가 없다.
from universe import passive_kind, active_groups, PASSIVE_MIN_AUM

def _load_snap():
    try:
        with open(os.path.join(os.path.dirname(HERE), "snapshots.json"), encoding="utf-8") as f: return json.load(f)
    except (OSError, ValueError):
        return {}
SNAP = _load_snap()

def _snap_of(e):
    return SNAP.get(f"{e['am']}:{e['id']}", {})

def _ncu(e, ymd):
    k = krx.get(ymd, {}).get(e["krx"])
    return (k["shrs"] / e["cu"]) if (k and k["shrs"] and e.get("cu")) else None

def _price(nm, v, p):
    q, pr = P._qp(v)
    return q, (pr or px.get(p, {}).get(nm))

def flow_categories():
    """통합 자금흐름의 범주: [(라벨, 판별함수(code, nm) -> bool)]. 패시브 3 + 액티브 4그룹."""
    last = krx[DAYS[-1]] if DAYS else {}
    big = {c for c, it in last.items() if (it["netasset"] or it["shrs"] * it["nav"]) >= PASSIVE_MIN_AUM}
    cats = [(f"패시브·{k}", (lambda k: lambda c, nm: c in big and passive_kind(nm) == k)(k)) for k in ("코스닥", "반도체", "2차전지")]
    for gk, lab, es in active_groups():
        codes = {e["krx"] for e in es}
        cats.append((f"액티브·{lab.replace(' 액티브', '')}", (lambda codes: lambda c, nm: c in codes)(codes)))
    return cats

def flows_by_category():
    """{범주: {"flow": [일별 순유출입 억], "aum": [일별 순자산 억], "n": ETF수}}, 날짜(변동일) 목록."""
    cats = flow_categories()
    out = {lab: {"flow": [], "aum": [], "codes": set()} for lab, _ in cats}
    days = DAYS[1:]
    for i in range(1, len(DAYS)):
        t, p = DAYS[i], DAYS[i - 1]
        acc = {lab: [0.0, 0.0] for lab, _ in cats}
        for code, it in krx[t].items():
            for lab, fn in cats:
                if not fn(code, it["nm"]): continue
                prev = krx[p].get(code)
                acc[lab][0] += (it["shrs"] - prev["shrs"]) * it["nav"] if (prev and it["nav"]) else 0.0
                acc[lab][1] += it["netasset"] or (it["shrs"] * it["nav"])
                out[lab]["codes"].add(code)
                break
        for lab in acc:
            out[lab]["flow"].append(acc[lab][0] / EOK); out[lab]["aum"].append(acc[lab][1] / EOK)
    return out, days

def group_net_buys(es):
    """그룹 ETF들의 일별 종목 순매수(스냅샷 기반). rows 형식은 net_buys() 와 같다."""
    rows = []; skipped = []
    for e in es:
        snaps = _snap_of(e)
        for i in range(1, len(DAYS)):
            t, p = DAYS[i], DAYS[i - 1]
            mt, mp = snaps.get(t), snaps.get(p)
            if not mt or not mp:
                skipped.append((e["name"], t, "PDF없음")); continue
            ncu = _ncu(e, p)                           # PDF(t)는 p일 종가·좌수 기준
            if not ncu:
                skipped.append((e["name"], t, "좌수없음")); continue
            for nm in set(mt) | set(mp):
                qt, pt = _price(nm, mt.get(nm, [0.0, None]), p)
                qp, pp = _price(nm, mp.get(nm, [0.0, None]), p)
                dq = qt - qp
                if dq == 0: continue
                price = pt or pp
                if not price:
                    skipped.append((e["name"], t, f"종가없음:{nm}")); continue
                sc = (sec.get(nm, {}) or {}).get("sector") or "미분류"
                rows.append({"일자": t, "ETF": e["name"], "종목": nm, "섹터": sc, "테마": theme_of(nm, sc),
                             "수량변화(1CU)": dq, "종가": price, "전체CU": ncu, "순매수금액(원)": dq * price * ncu})
    return rows, skipped

def group_holdings(es, key="stock"):
    """일자별 보유액(원) — key='stock' 종목별, 'theme' 테마별. returns (DAYS, [{k: 비중%}], [{k: 보유액}])."""
    ws = []; hs = []
    for i, day in enumerate(DAYS):
        prev = DAYS[i - 1] if i > 0 else day
        h = defaultdict(float)
        for e in es:
            m = _snap_of(e).get(day); ncu = _ncu(e, prev)
            if not m or not ncu: continue
            for nm, v in m.items():
                q, pr = _price(nm, v, prev)
                if not pr: continue
                if key == "theme":
                    sc = (sec.get(nm, {}) or {}).get("sector") or "미분류"
                    h[theme_of(nm, sc)] += q * pr * ncu
                else:
                    h[nm] += q * pr * ncu
        t = sum(h.values())
        ws.append({k: v / t * 100 for k, v in h.items()} if t else {})
        hs.append(dict(h))
    return DAYS, ws, hs

if __name__ == "__main__":
    print(f"분석 대상: {DAYS[0]} ~ {DAYS[-1]} ({len(DAYS)}일 → 변동 {len(DAYS)-1}일)")
    agg, detail = flows_aum()
    rows, skipped = net_buys()
    print(f"① 자금흐름 집계 {len(agg)}행 / ETF상세 {len(detail)}행")
    print(f"②③ 순매수 {len(rows)}행, 스킵 {len(skipped)}건")
    from collections import Counter
    print("   스킵사유:", Counter(s[2].split(':')[0] for s in skipped).most_common())
    # 요약 출력
    tot = defaultdict(float)
    for r in agg: tot[r["구분"]] += r["순유출입(원)"]
    print("\n[20일 누적 순유출입]")
    for k, v in tot.items(): print(f"  {k}: {v/EOK:+,.0f}억")
    s = defaultdict(float)
    for r in rows: s[r["섹터"]] += r["순매수금액(원)"]
    print("\n[섹터 20일 순매수 상위/하위]")
    ss = sorted(s.items(), key=lambda x: -x[1])
    for k, v in ss[:5]: print(f"  {k}: {v/EOK:+,.0f}억")
    for k, v in ss[-3:]: print(f"  {k}: {v/EOK:+,.0f}억")
