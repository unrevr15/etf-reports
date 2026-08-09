# -*- coding: utf-8 -*-
"""분석 대상 유니버스 정의 (이 폴더 전용 — 운영 pdf_change.py는 건드리지 않음).

사용자 지정: **코스닥액티브만** (바이오 전문 액티브 제외).
 - ① 자금유출입/AUM: KRX만 쓰므로 코스닥액티브 6종 전부 + 패시브
 - ②③ 섹터/종목 순매수: 20영업일 PDF 히스토리가 되는 4종만
   (DS·MIDAS는 자체 사이트에 날짜별 PDF가 없고 WiseReport는 최신일만 제공 → 제외)
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pdf_change as P

# 코스닥액티브 KRX 단축코드 (ACE K바이오코스닥액티브·커버드콜·채권혼합 제외)
KOSDAQ_ACTIVE_CODES = {
    "0163Y0": "KoAct 코스닥액티브",
    "0162Y0": "TIME 코스닥액티브",
    "0204S0": "TIGER 코스닥액티브",
    "0166N0": "PLUS 코스닥150액티브",
    "0220B0": "DS 코스닥액티브",        # ① 전용(PDF 히스토리 없음)
    "0191B0": "MIDAS 코스닥액티브",     # ① 전용(PDF 히스토리 없음)
}
# ① 액티브 그룹에서 뺄 것(코스닥이지만 순수 코스닥액티브가 아님)
ACTIVE_EXCLUDE_KEYWORDS = ("바이오", "커버드콜", "채권", "혼합")

def _find(pid):
    for e in P.ETFS:
        if e["id"] == pid: return dict(e)
    raise KeyError(pid)

# ②③용: 20영업일 PDF 히스토리가 확보되는 코스닥액티브 4종
ACTIVE_PDF = [
    _find("2ETFU6"),                       # KoAct 코스닥액티브
    _find("24"),                           # TIME 코스닥액티브
    {"name": "TIGER 코스닥액티브", "am": "미래에셋", "mode": "dateapi",
     "fetch": P.fetch_tiger, "id": "KR70204S0006", "krx": "0204S0", "cu": 10000},  # CU=10000(좌수 정수검증)
    _find("006399"),                       # PLUS 코스닥150액티브
]

def is_kosdaq_passive(nm):
    EXC = ("레버리지", "인버스", "선물", "채권", "커버드콜", "혼합", "숏", "액티브")
    return "코스닥" in nm and not any(x in nm for x in EXC)

def is_kosdaq_active(code, nm):
    return code in KOSDAQ_ACTIVE_CODES and not any(k in nm for k in ACTIVE_EXCLUDE_KEYWORDS)

if __name__ == "__main__":
    print("②③ PDF 대상:", [e["name"] for e in ACTIVE_PDF])
    print("① 액티브 코드:", KOSDAQ_ACTIVE_CODES)
