# -*- coding: utf-8 -*-
"""배포 직후 1회 실행 — 서울 VM IP에서 7개 운용사 endpoint가 응답하는지 확인.
하나라도 빈값/오류면 그 운용사가 이 IP를 차단하는 것 → 호스트/리전 재고 필요."""
import os, sys, datetime
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pdf_change as P

d = datetime.date.today()
while not P.is_trading_day(d):
    d = P.prev_trading_day(d)
ymd = d.strftime("%Y%m%d")
print(f"=== 연결 테스트 (기준일 {d}, 현재 IP에서) ===")
ok = 0
for e in P.ETFS:
    try:
        m, a = (e["fetch"](e["id"], ymd) if e["mode"] == "dateapi" else e["fetch"](e["id"]))
        # 당일 비면 직전 영업일로 1회 더(늦게 뜨는 곳 대비)
        if not m:
            p = P.prev_trading_day(d).strftime("%Y%m%d")
            m, a = (e["fetch"](e["id"], p) if e["mode"] == "dateapi" else e["fetch"](e["id"]))
        status = f"✅ {len(m):3}종목 (기준일 {a})" if m else "❌ 빈값/차단 의심"
        ok += 1 if m else 0
    except Exception as ex:
        status = f"❌ {type(ex).__name__}: {str(ex)[:50]}"
    print(f"  {e['name'][:18]:18} {status}")
print(f"\n응답 {ok}/{len(P.ETFS)}  —  " + ("전부 정상, 배포 진행 OK" if ok == len(P.ETFS)
      else "일부 차단 의심: 해당 운용사가 이 IP를 막는지 확인(다른 서울 호스트 시도)"))
