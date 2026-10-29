# -*- coding: utf-8 -*-
"""
클라우드 일일 오케스트레이터 (서울 리전 VM에서 cron 실행).
1) 일일 PDF 변화 리포트 (+ 금요일/주말 자동 주간 리포트)
2) 최근 5거래일 롤링 리포트
3) 포트폴리오 차트(PNG)
4) 당일 산출물을 Google Drive(rclone)로 업로드
모두 멱등(증분 스냅샷) — 하루에 여러 번 돌려도 안전(늦게 뜨는 TIGER를 채움).
"""
import os, sys, glob, subprocess, datetime, json, base64, faulthandler, urllib.request, urllib.error
from contextlib import contextmanager

HERE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.dirname(HERE)  # pdf_change.py / portfolio_chart.py 위치
sys.path.insert(0, APP)
os.chdir(APP)

import pdf_change as P
import portfolio_chart as PC

# 업로드 대상 rclone 리모트:폴더 (환경변수로 오버라이드 가능)
RCLONE_REMOTE = os.getenv("RCLONE_REMOTE", "gdrive:etf_reports")

def log(m): print(f"[{datetime.datetime.now():%H:%M:%S}] {m}", flush=True)

SEND_OPEN  = datetime.time(7, 10)   # 이 전에는 절대 안 보낸다
SEND_CLOSE = datetime.time(8, 40)   # 여기까지만 '완결'을 기다린다. 넘으면 되는대로 발송
SEND_HARD  = datetime.time(15, 0)   # 이 뒤 회차는 발송 금지. 11시였으나 GitHub 지연으로 첫 잡이 11시 넘어 뜨는 날(08-28 14:04)은 리포트가 통째로 빠져 15시로 완화. 중복은 날짜 마커가 막는다.
WAIT_FLOOR = datetime.time(6, 30)   # 이보다 이른 회차는 잡 안에서 기다리지 않는다
POLL_SEC   = 240                    # 잡 내부 재수집 간격(초)
WAIT_MAX_MIN = 150                  # 잡 하나가 기다릴 수 있는 절대 상한(분). 러너 점유 안전장치.
                                    # ※ 마감(SEND_CLOSE)까지는 이 상한 안에서 반드시 기다린다 —
                                    #   마감 직전에 예산이 끊기면 보낼 수 있었는데 못 보내고 끝난다.
                                    #   WAIT_FLOOR(06:30) 시작 시 마감까지 131분이므로 150이면 충분하다.
# 완결 판정에서 빼는 ETF = pdf_change.ETFS 에서 "lag":True 인 것(당일분이 늦거나 전일분까지만 오는 소스).
# 이들은 당일분이 없으면 '전일분'으로 표시되므로 기다릴 이유가 없다. 이름 전체로 맞춘다
# (부분매칭이면 이름에 TIGER가 든 ETF가 나중에 추가될 때 조용히 빠진다).
LATE = tuple(e["name"] for e in P.ETFS if e.get("lag"))

# ── 무한 정지 안전장치 ─────────────────────────────────────────
# 2026-10-21 09:32 회차(#705)가 일일 리포트 계산 중 아무 로그 없이 멈췄고, 180분 잡 제한에 걸릴 때까지
# 같은 concurrency 그룹의 다음 회차들이 전부 대기열에 묶였다. 원인은 기록이 없어 모른다.
# → 구간마다 감시 타이머: 넘기면 모든 스레드의 스택을 로그에 남기고(원인 추적) 프로세스를 끝낸다.
#   커밋 단계는 if: always() 라 그때까지의 캡처는 레포에 남고, 발송 마커는 보내는 즉시 따로 올린다.
@contextmanager
def guard(minutes):
    faulthandler.dump_traceback_later(minutes * 60, exit=True)
    try:
        yield
    finally:
        faulthandler.cancel_dump_traceback_later()

# ── 발송 마커: 레포(원격)에서 직접 확인·기록 ─────────────────────
# 대기열에 섰던 회차는 '생성 시점' 커밋을 checkout 한다(2026-10-21 #707). 그 사이 앞 회차가 보내고 마커를
# 올렸어도 이 회차의 작업트리엔 마커가 없다 → 그대로면 같은 날 방에 한 번 더 보낸다.
# 그래서 ① 발송 판단 직전에 원격 main의 마커를 GitHub API로 확인하고 ② 보내자마자 마커를 API로 바로 올린다
# (잡 끝 커밋 단계가 실패·중단돼도 마커는 남게). 토큰이 없으면(로컬 실행) 로컬 파일만 본다.
MARK = "1\n"

def _gh(method, path, body=None):
    tok, repo = os.getenv("GITHUB_TOKEN"), os.getenv("GITHUB_REPOSITORY")
    if not (tok and repo):
        return None
    req = urllib.request.Request(f"https://api.github.com/repos/{repo}/{path}", method=method,
                                 data=(json.dumps(body).encode("utf-8") if body is not None else None),
                                 headers={"Authorization": f"Bearer {tok}", "Accept": "application/vnd.github+json",
                                          "Content-Type": "application/json", "User-Agent": "etf-daily"})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.status
    except urllib.error.HTTPError as e:
        return e.code
    except Exception:
        return None

def sync_marker(path):
    """원격 main에 이 마커가 있으면 로컬에도 만들고 True. 로컬에 이미 있어도 True."""
    if os.path.exists(path):
        return True
    if _gh("GET", f"contents/reports/{os.path.basename(path)}?ref=main") == 200:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w") as f: f.write(MARK)
        log(f"  [마커] 원격에 {os.path.basename(path)} 있음 — 오늘 이미 발송된 것으로 처리")
        return True
    return False

def publish_marker(name):
    st = _gh("PUT", f"contents/reports/{name}",
             {"message": f"발송 마커 {name}", "branch": "main", "content": base64.b64encode(MARK.encode()).decode()})
    if st is None: return
    log(f"  [마커] 레포에 바로 기록 {'완료' if st in (200, 201) else f'실패(HTTP {st}) — 잡 끝 커밋에서 다시 올림'}")

DEGRADE_AFTER = 3   # 연속 이 거래일 수만큼 결측이면 '장기 파손'으로 보고 완결 판정에서 뺀다
MAX_DEGRADED = 3    # 이보다 많이 깨졌으면 자동 제외하지 않는다(껍데기 리포트가 나가는 것 방지)

def degraded_set(ctx):
    """직전 DEGRADE_AFTER 거래일 내내 유효 캡처가 없는 ETF 이름 집합.
    RISE처럼 한 곳이 오래 파손되면 매일 마감(8:40)까지 기다리게 되어 발송이 9시 반으로 밀린다.
    스냅샷만 보고 판정하므로 상태파일이 필요 없고, 복구되면 자동으로 집합에서 빠진다.
    오늘은 세지 않는다(아직 게시 전일 수 있으므로)."""
    try:
        snap = P.load_snap()
        days = []
        d = ctx["today"]
        for _ in range(DEGRADE_AFTER):
            d = P.prev_trading_day(d)
            days.append(d.strftime("%Y%m%d"))
        out = set()
        for e in P.ETFS:
            kk = snap.get(f"{e['am']}:{e['id']}", {})
            if all(not kk.get(x) for x in days):
                out.add(e["name"])
        return out if len(out) <= MAX_DEGRADED else set()
    except Exception as ex:
        log(f"  [열화] 판정 실패(무시): {type(ex).__name__}")
        return set()

def _push_dash(ctx, today):
    """대시보드 업로드. 발송 여부와 무관하게, 부분치라도 올린다(같은 날 재실행 시 덮어씀).
    경로를 짐작하지 않고 pdf_change 가 방금 쓴 파일을 ctx 로 받아 쓴다."""
    try:
        sys.path.insert(0, HERE)
        import dashboard_push as DP
        DP.push_from_csv(ctx.get("csv_path"), today.strftime("%Y-%m-%d"))
    except Exception as e:
        log(f"  대시보드 업로드 건너뜀: {e}")

def _send_state(ctx, grp=None):
    """지금 시각 기준 발송 판정 상태. grp 를 주면 그 발송 그룹 ETF만 완결을 본다."""
    now_t = datetime.datetime.now().time()            # TZ=Asia/Seoul (워크플로 env)
    deg = degraded_set(ctx)                           # 장기 파손 ETF → 완결 판정에서 자동 제외
    gs = [g for g in ctx["groups"] if grp is None or g.get("grp") == grp]
    pend_raw = [g["etf"] for g in gs if g["state"] == "pending" and g["etf"] not in LATE]
    pend = [n for n in pend_raw if n not in deg]
    krx_ok = ctx.get("krx_ok", True)                  # KRX 실패면 금액·순자산·현금 통째 빈칸 → 보류
    return {"earliest": now_t >= SEND_OPEN,
            "deadline": now_t >= SEND_CLOSE,
            "hard_stop": now_t >= SEND_HARD,
            "complete": (len(pend) == 0) and krx_ok,
            "pend": pend,
            "deg_hit": [n for n in pend_raw if n in deg],
            "krx_ok": krx_ok}

def _can_send(st):
    return st["earliest"] and (not st["hard_stop"]) and (st["complete"] or st["deadline"])

# ── 그룹별 발송 ────────────────────────────────────────────────
# 2026-10-29부터 리포트를 발송 그룹(P.REPORT_GROUPS: 코스닥·반도체·바이오·ESS)마다 따로 보낸다.
# 마커도 채널×그룹: reports/.tg_sent_{ymd}_{grp}(채널1), .tg_sent2_{ymd}_{grp}(채널2).
# 그룹 도입 전 마커(.tg_sent_{ymd})가 있으면 그 채널은 그날 이미 전부 보낸 것으로 본다(배포 당일 재발송 방지).
CHANNELS = (("채널1", "wide", "TELEGRAM_TOKEN", "TELEGRAM_CHAT_ID", ".tg_sent"),     # @pefscreener 색상·가로
            ("채널2", "mobile", "TELEGRAM_TOKEN2", "TELEGRAM_CHAT_ID2", ".tg_sent2"))  # 수급방 무채색·모바일

def _mk(prefix, ymd, grp=None):
    return os.path.join(APP, "reports", f"{prefix}_{ymd}" + (f"_{grp}" if grp else ""))

def _sent(prefix, ymd, grp):
    return sync_marker(_mk(prefix, ymd)) or sync_marker(_mk(prefix, ymd, grp))

_GAVE_UP = set()   # 이번 잡에서 MAX_SEND_FAIL번 실패한 (채널, 그룹) — 대기 루프가 그 방 때문에 마감까지 돌지 않게(다음 회차가 재시도)
MAX_SEND_FAIL = 2
_FAILS = {}

def _configured(ch):
    c = next(x for x in CHANNELS if x[0] == ch)
    return bool(os.getenv(c[2]) and os.getenv(c[3]))

def _group_done(ymd, grp):
    """그룹이 이번 잡에서 더 할 일 없음 = 설정된 채널마다 보냈거나(마커) 이번 잡에서 포기."""
    return all(_sent(c[4], ymd, grp) or not _configured(c[0]) or (c[0], grp) in _GAVE_UP for c in CHANNELS)

def _group_payload(ctx, grp):
    label = P.REPORT_LABEL[grp]
    gs = [g for g in ctx["groups"] if g.get("grp") == grp]
    status = P.group_status_line(gs)
    cap = f"{label} ETF PDF 변화  {ctx['today']:%Y-%m-%d}\n{status}"
    return label, gs, status, cap

def send_group(ctx, grp, dm=None):
    """한 그룹을 두 채널로 보낸다. 변화가 하나도 없으면 사진 대신 짧은 텍스트.
    dm=(토큰, chat) 이면 점검용: 두 형식을 그 DM으로만, 마커 없음."""
    label, gs, status, cap = _group_payload(ctx, grp)
    ymd = ctx["today"].strftime("%Y%m%d")
    for ch, kind, tok_env, chat_env, prefix in CHANNELS:
        try:
            tok, chat = (dm if dm else (os.getenv(tok_env), os.getenv(chat_env)))
            if not (tok and chat): log(f"  [{ch}·{label}] 대상 미설정 — 스킵"); continue
            if not dm and _sent(prefix, ymd, grp): log(f"  [{ch}·{label}] 오늘 이미 발송 — 스킵"); continue
            pngs = P.render_report_parts(kind, gs, ctx["today"], ctx["prev"], status, label=label, tag=grp)
            if not pngs and not dm and not _send_state(ctx, grp)["complete"]:
                # 그림이 없는 이유가 '변화 없음'이 아니라 '아직 안 들어옴'(마감 시점 대기 ETF)이면 확정 불가 →
                # '변화 없음'을 보내거나 마커를 남기지 않는다. 데이터가 오면 다음 회차(15:00 전)가 제대로 보낸다.
                log(f"  [{ch}·{label}] 그릴 변화 없음이나 대기 ETF 있음({', '.join(_send_state(ctx, grp)['pend'])}) — 마커 없이 스킵")
                continue
            c = (f"[DM 점검 · {ch} 형식]\n" if dm else "") + cap
            ok = (P.send_telegram(pngs, c, token=tok, chat=chat) if pngs
                  else P.send_text(c + "\n(오늘 구성 변화 없음)", token=tok, chat=chat))
            log(f"  [{ch}·{label}] {'사진 %d장' % len(pngs) if pngs else '텍스트(변화 없음)'} → {'완료' if ok else '실패'}")
            if not ok and not dm:
                _FAILS[(ch, grp)] = _FAILS.get((ch, grp), 0) + 1
                if _FAILS[(ch, grp)] >= MAX_SEND_FAIL:
                    _GAVE_UP.add((ch, grp)); log(f"  [{ch}·{label}] {MAX_SEND_FAIL}회 실패 — 이번 잡에선 포기(다음 회차 재시도)")
            if ok and not dm:
                m = _mk(prefix, ymd, grp)
                os.makedirs(os.path.dirname(m), exist_ok=True)
                with open(m, "w") as f: f.write(MARK)
                publish_marker(os.path.basename(m))
        except Exception as e:
            log(f"  [{ch}·{label}] 실패: {type(e).__name__}: {e}")

def wait_and_send(ctx, today):
    """크론 회차를 기다리지 말고 '잡 안에서' 기다리며, 그룹마다 완결되는 즉시 보낸다.

    2026-08-27부터 GitHub 스케줄러가 열화됐다 — 예정 회차가 수 시간 늦거나 안 뜬다.
    그래서 같은 잡 안에서 재수집하며 조건이 서는 그룹부터 바로 보낸다(08:40 마감이면 남은 그룹을 되는대로).
    재수집은 증분이라 이미 캡처된 ETF는 다시 받지 않는다(대기 중인 것만 재시도)."""
    import time
    ymd = today.strftime("%Y%m%d")
    keys = [k for k, _ in P.REPORT_GROUPS]
    t0 = datetime.datetime.now()
    while True:
        todo = [k for k in keys if not _group_done(ymd, k)]
        for k in todo:
            st = _send_state(ctx, k)
            if _can_send(st):
                if st["deg_hit"]:
                    log(f"  [열화] 연속 {DEGRADE_AFTER}거래일 결측 → 완결 판정 제외: {', '.join(sorted(st['deg_hit']))}")
                send_group(ctx, k)
        todo = [k for k in todo if not _group_done(ymd, k)]
        if not todo:
            return ctx
        now = datetime.datetime.now()
        st_all = _send_state(ctx)
        if st_all["hard_stop"]:
            log(f"  발송 금지 — {SEND_HARD:%H:%M} 넘음(아침 리포트 아님). 미발송: {', '.join(todo)}"); return ctx
        if now.time() < WAIT_FLOOR:
            log(f"  [대기] {WAIT_FLOOR:%H:%M} 이전 회차 — 기다리지 않음"); return ctx
        spent = (now - t0).total_seconds() / 60.0
        if spent >= WAIT_MAX_MIN:
            log(f"  [대기] 상한 {WAIT_MAX_MIN}분 도달 — 미발송 그룹: {', '.join(todo)}"); return ctx
        if st_all["deadline"] and all(_can_send(_send_state(ctx, k)) for k in todo):
            log(f"  [대기] 마감 후에도 미발송(대상 미설정·전송 실패): {', '.join(todo)}"); return ctx
        if not st_all["earliest"]:                    # 아직 발송창 전 → 07:10까지 한 번에 잔다
            nap = (datetime.datetime.combine(now.date(), SEND_OPEN) - now).total_seconds() + 20
            why = f"발송창 {SEND_OPEN:%H:%M} 대기"
        else:
            nap = POLL_SEC
            why = " / ".join(f"{P.REPORT_LABEL[k]}: " + (", ".join(n.split()[0] for n in _send_state(ctx, k)["pend"]) or "KRX(금액)")
                             for k in todo)
        nap = max(30.0, min(nap, (WAIT_MAX_MIN - spent) * 60.0))
        log(f"  [대기] {why} — {nap/60:.0f}분 후 재수집 (누적 {spent:.0f}분, 상한 {WAIT_MAX_MIN})")
        time.sleep(nap)
        with guard(30):
            nc = P.run(today, weekly=False)           # 증분 재수집(주간은 나중에 한 번만)
        if nc: ctx = nc

def main():
    today = datetime.date.today()
    if not P.is_trading_day(today):
        log(f"{today} 휴장일 — 종료"); return
    ymd = today.strftime("%Y%m%d")

    log("1) 일일 리포트")
    # weekly=False: 주간 리포트는 대기가 끝난 뒤(데이터가 가장 많이 찬 시점) 한 번만 만든다.
    # 대기 루프가 폴링마다 주간을 재생성하면 운용사 서버를 수십 번 더 두드린다.
    with guard(30):
        ctx = P.run(today, weekly=False)    # pdf_change_YYYYMMDD.xlsx
    # ── 대시보드 업로드 ──────────────────────────────────
    # 리포트가 만들어진 직후, 발송 여부와 무관하게 올린다.
    # (발송은 완결 판정을 기다리지만 대시보드는 부분치라도 보이는 편이 낫다 —
    #  같은 날 다시 실행되면 덮어쓴다)
    _push_dash(ctx, today)

    log("1-b) 그룹별 이미지 렌더 + 채널 전송 (아침 7:10~8:40, 그룹마다 완결되면 즉시·마감시 되는대로, 하루 1회)")
    send_win = os.getenv("TELEGRAM_SEND", "").lower() in ("1", "true", "yes")   # 발송 후보 실행인가
    if os.getenv("DM_TEST", "").lower() == "true":
        # 수동 점검 전용(workflow_dispatch dm_test): 4개 그룹 × 두 형식을 사용자 DM(봇2)으로만.
        # 방 발송·발송 마커 없음 — force_send 를 같이 켜도 방으로는 안 나간다.
        dm = (os.getenv("TELEGRAM_TOKEN2"), os.getenv("TELEGRAM_DM_ID"))
        if all(dm):
            for k, _ in P.REPORT_GROUPS: send_group(ctx, k, dm=dm)
        else:
            log("  [DM 점검] DM 미설정 — 스킵")
    elif send_win:
        if all(_group_done(ymd, k) for k, _ in P.REPORT_GROUPS):
            log("  오늘 4개 그룹 모두 이미 발송")
        else:
            ctx = wait_and_send(ctx, today)               # ★ 크론 사이가 아니라 잡 안에서 대기, 그룹마다 즉시 발송
            _push_dash(ctx, today)                        # 대기 동안 채워진 것을 대시보드에 반영
    else:
        for k, label in P.REPORT_GROUPS:                  # 무발송 회차: 렌더만(산출물 갱신)
            _, gs, status, _ = _group_payload(ctx, k)
            for _, kind, *_ in CHANNELS:
                P.render_report_parts(kind, gs, ctx["today"], ctx["prev"], status, label=label, tag=k)
        log("  전송 회차 아님 — 렌더만")
    # 주간 리포트 — 대기까지 끝난 뒤 한 번만(그 주 마지막 거래일에만)
    if P.is_last_trading_day_of_week(today):
        with guard(20):
            try: P.weekly_report(today)
            except Exception as ex: log(f"  [주간] 생성 실패: {type(ex).__name__}: {ex}")

    log("2) 최근 5거래일 롤링")
    with guard(20):
        try: P.rolling_report(today, 5)    # pdf_rolling5_YYYYMMDD.xlsx
        except Exception as e: log(f"  롤링 실패: {type(e).__name__}: {e}")
    log("3) 포트폴리오 차트")
    with guard(20):
        try: PC.generate(today)            # portfolio_YYYYMMDD.png
        except Exception as e: log(f"  차트 실패: {type(e).__name__}: {e}")

    # 4) 오늘 산출물 업로드
    targets = []
    for pat in (f"pdf_change_{ymd}.xlsx", f"pdf_change_{ymd}*.png", f"pdf_change_m_{ymd}*.png",   # *: 여러 장(_1·_2…)
                f"pdf_weekly_{ymd}.xlsx", f"pdf_rolling5_{ymd}.xlsx", f"portfolio_{ymd}.png"):
        targets += glob.glob(os.path.join(APP, pat))
    if not targets:
        log("업로드할 파일 없음"); return
    log(f"4) Drive 업로드 {len(targets)}개 → {RCLONE_REMOTE}")
    for f in targets:
        try:
            subprocess.run(["rclone", "copy", f, RCLONE_REMOTE + "/", "--quiet"], check=True, timeout=120)
            log(f"  ✅ {os.path.basename(f)}")
        except Exception as e:
            log(f"  ❌ {os.path.basename(f)}: {e}")
    log("완료")

if __name__ == "__main__":
    main()
