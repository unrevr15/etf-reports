# -*- coding: utf-8 -*-
"""전체 파이프라인 한 번에: 수집 → (누락 재시도) → 섹터 → 그래프·엑셀.
   실행:  python run_all.py
   캐시(json)가 있으면 신규분만 받으므로 재실행이 빠름.
"""
import os, sys, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

def step(name, script):
    print(f"\n{'='*60}\n▶ {name}\n{'='*60}", flush=True)
    r = subprocess.run([PY, os.path.join(HERE, script)], cwd=HERE)
    if r.returncode != 0:
        print(f"  [경고] {script} 종료코드 {r.returncode}")

if __name__ == "__main__":
    step("1) KRX ETF 일별 + 액티브 PDF 수집", "collect.py")
    step("2) 누락 PDF 재시도 (rate limit 회피 지연)", "retry_pdf.py")
    step("3) 코스닥 종가 수집", "collect_px.py")
    step("4) 종목→섹터 매핑", "sectors.py")
    step("5) 그래프 3종 + 엑셀 생성", "charts.py")
    print("\n완료 — analysis_20d 폴더의 PNG 3개와 20영업일_분석.xlsx 확인")
