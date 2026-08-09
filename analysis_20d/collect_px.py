# -*- coding: utf-8 -*-
"""일자별 코스닥 종가 수집 (PLUS 등 평가금액 미제공 ETF의 종가 보완용)."""
import os, sys, json
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import pdf_change as P
from collect import trading_days

PX_FILE = os.path.join(HERE, "krx_px.json")

def main():
    try:
        with open(PX_FILE, encoding="utf-8") as f: cache = json.load(f)
    except (OSError, ValueError): cache = {}
    key = P._krx_key()
    days = trading_days()
    todo = [d for d in days if d.strftime("%Y%m%d") not in cache]
    print(f"[종가] {len(days)}일 중 신규 {len(todo)}일")
    for d in todo:
        ymd = d.strftime("%Y%m%d")
        items = P._krx_get("http://data-dbg.krx.co.kr/svc/apis/sto/ksq_bydd_trd",
                           {"AUTH_KEY": key, "basDd": ymd}).get("OutBlock_1", [])
        m = {}
        for it in items:
            try: m[it["ISU_NM"].strip()] = float(str(it["TDD_CLSPRC"]).replace(",", ""))
            except (ValueError, KeyError): pass
        if m:
            cache[ymd] = m
            print(f"  {ymd}: {len(m)}종목")
    with open(PX_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)
    print("완료")

if __name__ == "__main__":
    main()
