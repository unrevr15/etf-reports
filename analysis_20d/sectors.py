# -*- coding: utf-8 -*-
"""종목명 → 업종(섹터) 매핑 구축. 네이버 금융 기준.
   ① 자동완성 API로 종목명→코드  ② 종목 페이지에서 업종명 추출
   결과는 sectors.json 캐시(재실행 시 신규 종목만 조회).
"""
import os, sys, json, re, requests
from concurrent.futures import ThreadPoolExecutor, as_completed

HERE = os.path.dirname(os.path.abspath(__file__))
PDF_FILE = os.path.join(HERE, "pdf_raw.json")
SEC_FILE = os.path.join(HERE, "sectors.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36"

def load(p, d=None):
    try:
        with open(p, encoding="utf-8") as f: return json.load(f)
    except (OSError, ValueError): return d if d is not None else {}

def stock_names():
    """PDF 캐시에 등장한 모든 종목명."""
    raw = load(PDF_FILE)
    names = set()
    for eid in raw:
        for ymd in raw[eid]:
            names.update(raw[eid][ymd].keys())
    return sorted(names)

def naver_code(nm):
    r = requests.get("https://ac.stock.naver.com/ac", params={"q": nm, "target": "stock"},
                     headers={"User-Agent": UA}, timeout=12)
    for it in r.json().get("items", []):
        if it.get("nationCode") == "KOR" and it.get("code"):
            # 이름이 정확히 일치하는 걸 우선
            if it.get("name", "").strip() == nm: return it["code"], it.get("typeCode", "")
    for it in r.json().get("items", []):
        if it.get("nationCode") == "KOR" and it.get("code"):
            return it["code"], it.get("typeCode", "")
    return None, None

def naver_sector(code):
    r = requests.get(f"https://finance.naver.com/item/main.naver?code={code}",
                     headers={"User-Agent": UA}, timeout=15)
    r.encoding = "utf-8"
    m = re.search(r'href="[^"]*sise_group_detail[^"]*"[^>]*>([^<]+)</a>', r.text)
    return m.group(1).strip() if m else None

def resolve(nm):
    try:
        code, mkt = naver_code(nm)
        if not code: return nm, {"code": None, "sector": None, "market": None}
        sec = naver_sector(code)
        return nm, {"code": code, "sector": sec, "market": mkt}
    except Exception:
        return nm, {"code": None, "sector": None, "market": None}

if __name__ == "__main__":
    cache = load(SEC_FILE)
    names = stock_names()
    todo = [n for n in names if n not in cache or not cache[n].get("sector")]
    print(f"종목 {len(names)}개 중 신규/미해결 {len(todo)}개 조회")
    done = 0
    with ThreadPoolExecutor(max_workers=5) as ex:
        futs = [ex.submit(resolve, n) for n in todo]
        for f in as_completed(futs):
            nm, info = f.result()
            cache[nm] = info
            done += 1
            if done % 20 == 0:
                print(f"  {done}/{len(todo)}...", flush=True)
                with open(SEC_FILE, "w", encoding="utf-8") as fp:
                    json.dump(cache, fp, ensure_ascii=False, indent=0)
    with open(SEC_FILE, "w", encoding="utf-8") as fp:
        json.dump(cache, fp, ensure_ascii=False, indent=0)
    ok = sum(1 for n in cache if cache[n].get("sector"))
    print(f"완료 — 섹터 확보 {ok}/{len(cache)}")
    from collections import Counter
    c = Counter(cache[n]["sector"] for n in cache if cache[n].get("sector"))
    print("상위 섹터:", c.most_common(12))
    miss = [n for n in cache if not cache[n].get("sector")]
    if miss: print("미해결:", miss[:20])
