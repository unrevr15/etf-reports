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


# ── 종목명 → 종목코드 ────────────────────────────────────────
# 대시보드는 종목을 코드로 잇는다. 지금까지 이름만 보내서, 운용사 PDF 표기가
# 대시보드의 이름표(DART 발음 이름)와 다르면 코드가 안 붙었다(51종목 실측).
# 한국투자 종목마스터는 공개 파일이라 인증이 필요 없고 코스피·코스닥을 다 준다.
# 실패하면 코드를 안 붙일 뿐, 업로드는 예전과 똑같이 돈다.
_CODE_MAP: dict[str, str] = {}
_CODE_TRIED = False
_MASTER = "https://new.real.download.dws.co.kr/common/master/{}_code.mst.zip"


def _name_to_code() -> dict:
    global _CODE_TRIED
    if _CODE_MAP or _CODE_TRIED:
        return _CODE_MAP
    _CODE_TRIED = True
    import io as _io
    import zipfile
    dup = set()
    for mk, tail in (("kospi", 228), ("kosdaq", 222)):
        try:
            with urllib.request.urlopen(_MASTER.format(mk), timeout=TIMEOUT) as r:
                z = zipfile.ZipFile(_io.BytesIO(r.read()))
            raw = z.read(z.namelist()[0]).decode("cp949", "replace")
        except Exception as e:
            print(f"[DASHBOARD] 종목마스터({mk}) 실패 — 코드 없이 올립니다: {e}", flush=True)
            continue
        for line in raw.splitlines():
            if not line.strip():
                continue
            head = line[: len(line) - tail]
            code, name = head[:9].strip(), head[21:].strip()
            if len(code) != 6 or not code.isdigit() or not name:
                continue
            # 같은 이름이 두 종목에 걸리면 어느 쪽인지 단정할 수 없다 → 둘 다 버린다
            if name in _CODE_MAP and _CODE_MAP[name] != code:
                dup.add(name)
            else:
                _CODE_MAP[name] = code
    for n in dup:
        _CODE_MAP.pop(n, None)
    print(f"[DASHBOARD] 종목명→코드 {len(_CODE_MAP)}개"
          + (f" (이름 중복 {len(dup)}개 제외)" if dup else ""), flush=True)
    return _CODE_MAP


def push(rows: list[dict], base_date: str) -> int:
    """rows: [{etf, house, kind, name, today_qty, prev_qty, delta}, ...]"""
    url = (os.getenv("DASHBOARD_URL") or "").rstrip("/")
    token = os.getenv("DASHBOARD_TOKEN") or ""
    if not url or not token or not rows:
        return 0

    cmap = _name_to_code()
    clean = []
    hit = 0
    for r in rows:
        nm = str(r.get("name") or "").strip()
        if not nm:
            continue
        # 행이 이미 코드를 들고 있으면 그것을 쓴다(TIGER 는 KSD 표에서 코드를 받아온다)
        code = str(r.get("code") or "").strip() or cmap.get(nm, "")
        if len(code) != 6 or not code.isdigit():
            code = ""
        if code:
            hit += 1
        clean.append({
            "etf": str(r.get("etf") or "")[:40],
            "house": str(r.get("house") or "")[:20],
            "kind": str(r.get("kind") or "")[:12],     # 신규편입 / 전량제외 / 수량확대 / 수량축소
            "name": nm[:30],
            "code": code or None,
            "today_qty": _num(r.get("today_qty")),
            "prev_qty": _num(r.get("prev_qty")),
            "delta": _num(r.get("delta")),
        })
    if not clean:
        return 0
    print(f"[DASHBOARD] 코드 붙은 행 {hit}/{len(clean)}", flush=True)

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

    if not csv_path or not os.path.isfile(csv_path):
        _log(f"CSV 를 찾지 못했습니다: {csv_path}")   # 조용히 0건이 되던 자리
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
                "prev_qty": r.get("전일수량"),
                # 열 이름이 '수량변화(1CU)' 다. '수량변화' 로 찾으면 늘 None 이 된다.
                "delta": r.get("수량변화(1CU)") or r.get("수량변화"),
            })
    return push(rows, base_date)
