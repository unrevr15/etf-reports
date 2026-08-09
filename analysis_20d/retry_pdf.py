# -*- coding: utf-8 -*-
"""누락된 PDF 스냅샷만 천천히 재수집 (rate limit 회피: 순차 + 지연)."""
import os, sys, json, time
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import pdf_change as P
from collect import trading_days, PDF_FILE

DELAY = 2.5   # 호출 간 지연(초)

def main():
    with open(PDF_FILE, encoding="utf-8") as f: cache = json.load(f)
    days = [d.strftime("%Y%m%d") for d in trading_days()]
    todo = [(e, y) for e in P.ETFS for y in days if not cache.get(e["id"], {}).get(y)]
    print(f"누락 {len(todo)}건 재수집 (지연 {DELAY}s, 예상 {len(todo)*DELAY/60:.1f}분)")
    ok = 0
    for i, (e, ymd) in enumerate(todo, 1):
        try:
            m, act = P._fetch(e, ymd)
            good = bool(m) and str(act) == ymd
            cache.setdefault(e["id"], {})[ymd] = m if good else {}
            ok += good
            print(f"  [{i}/{len(todo)}] {e['name'][:14]:14} {ymd} {'OK ' + str(len(m)) + '종목' if good else '실패(' + str(act or 'empty') + ')'}", flush=True)
        except Exception as ex:
            print(f"  [{i}/{len(todo)}] {e['name'][:14]:14} {ymd} ERR {type(ex).__name__}", flush=True)
        if i % 10 == 0:
            with open(PDF_FILE, "w", encoding="utf-8") as f: json.dump(cache, f, ensure_ascii=False)
        time.sleep(DELAY)
    with open(PDF_FILE, "w", encoding="utf-8") as f: json.dump(cache, f, ensure_ascii=False)
    print(f"재수집 완료 — 성공 {ok}/{len(todo)}")

if __name__ == "__main__":
    main()
