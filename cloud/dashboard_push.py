# -*- coding: utf-8 -*-
"""ETF PDF 구성변화를 블로그 대시보드로 올린다.

이미 계산이 끝난 결과(pdf_change 의 행 목록)를 그대로 넘기기만 한다. 여기서 다시
판정하지 않는다 — 규칙이 두 곳에 있으면 텔레그램 리포트와 대시보드가 갈라진다.

'코스닥 etf 분류' 가 올리는 것과 성격이 다르다:
  · 코스닥 etf 분류 → 보유주수의 증감 (얼마나 더 담았나)
  · 여기(ETF 총량)  → 구성종목의 변화  (새로 넣었나 뺐나)
둘을 합쳐야 "신규 편입 + 대량 매집" 을 구분할 수 있다.

설정이 없으면 조용히 건너뛴다(선택 기능). 실패해도 예외를 올리지 않는다 —
업로드 때문에 리포트 생성·발송이 막히면 본말전도다.
"""
from __future__ import annotations

import json
import os
import urllib.error
import urllib.request

TIMEOUT = 30
CHUNK = 300


def _num(v):
    try:
        return int(float(str(v).replace(",", "")))
    except (TypeError, ValueError):
        return None


def push(rows: list[dict], base_date: str) -> int:
    """rows: [{etf, house, kind, name, today_qty, prev_qty, delta}, ...]"""
    url = (os.getenv("DASHBOARD_URL") or "").rstrip("/")
    token = os.getenv("DASHBOARD_TOKEN") or ""
    if not url or not token or not rows:
        return 0

    clean = []
    for r in rows:
        nm = str(r.get("name") or "").strip()
        if not nm:
            continue
        clean.append({
            "etf": str(r.get("etf") or "")[:40],
            "house": str(r.get("house") or "")[:20],
            "kind": str(r.get("kind") or "")[:12],     # 신규편입 / 전량제외 / 수량확대 / 수량축소
            "name": nm[:30],
            "today_qty": _num(r.get("today_qty")),
            "prev_qty": _num(r.get("prev_qty")),
            "delta": _num(r.get("delta")),
        })
    if not clean:
        return 0

    sent = 0
    for i in range(0, len(clean), CHUNK):
        body = json.dumps({"base_date": base_date, "rows": clean[i:i + CHUNK]},
                          ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(
            url + "/api/ingest/etfchange",
            data=body,
            headers={
                "Content-Type": "application/json",
                "X-Ingest-Token": token,
                # Cloudflare 가 기본 urllib UA 를 403 으로 막는다
                "User-Agent": "etf-total/1.0 (+dashboard-ingest)",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
                sent += json.loads(resp.read().decode("utf-8")).get("saved", 0)
        except urllib.error.HTTPError as e:
            print(f"[DASHBOARD] 업로드 실패 HTTP {e.code}: "
                  f"{e.read().decode('utf-8', 'replace')[:160]}", flush=True)
            return sent
        except Exception as e:
            print(f"[DASHBOARD] 업로드 실패: {e}", flush=True)
            return sent
    print(f"[DASHBOARD] ETF 구성변화 {sent}건 업로드", flush=True)
    return sent


def push_from_csv(csv_path: str, base_date: str) -> int:
    """pdf_change 가 남긴 CSV 를 그대로 읽어 올린다.

    열: 실행시각 · ETF · 운용사 · 당일기준일 · 전일기준일 · 구분 · 종목명 ·
        당일수량 · 전일수량 · 수량변화
    같은 날 여러 번 실행되므로 '당일기준일' 이 base_date 인 행만 고른다.
    """
    import csv as _csv
    import io as _io

    if not os.path.isfile(csv_path):
        return 0
    ymd = base_date.replace("-", "")
    rows = []
    with _io.open(csv_path, encoding="utf-8-sig", errors="replace", newline="") as f:
        for r in _csv.DictReader(f):
            if (r.get("당일기준일") or "").strip() != ymd:
                continue
            rows.append({
                "etf": r.get("ETF"), "house": r.get("운용사"), "kind": r.get("구분"),
                "name": r.get("종목명"), "today_qty": r.get("당일수량"),
                "prev_qty": r.get("전일수량"), "delta": r.get("수량변화"),
            })
    return push(rows, base_date)
