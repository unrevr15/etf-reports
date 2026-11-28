# -*- coding: utf-8 -*-
"""주간 ETF 리포트 그림 (채널 공통 틀 A안, 11-27).
폭 1080px 흰 바탕, 바깥 여백 24px, 제목 줄(왼쪽 제목, 오른쪽 날짜) 나눔고딕 Bold 48px 남색.
글자는 전부 28px 이상, 색은 남색 빨강 파랑 검정 회색과 줄무늬, 연한 파랑, 묶음 회색만 쓴다.
그림 한 장 높이 1920px 이하, 넘으면 행 단위로 나눠 앨범으로 보낸다(둘째 장부터 제목 뒤에 "(2/3)").
모든 좌표는 px(위에서 아래). 캔버스를 1920px 높이로 잡고 쓴 만큼만 잘라 낸다.
"""
import os, io, sys
from collections import defaultdict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.ticker import FuncFormatter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import pdf_change as P

EOK = 1e8
W, MAXH, M, DPI = 1080, 1920, 24, 100
PT = 72.0 / DPI                      # px -> pt
RIGHT = W - M
TOP, TITLE_H = M, 84          # 위 여백 24, 제목 줄 24~108
CONTENT0 = TOP + TITLE_H      # 첫 내용 시작 y=108
GUT = 120                     # 선 그래프 왼쪽 고정 여백(눈금 글자 자리). 그림마다 그림 영역이 x=24+120=144 에서 시작
SMALL = 31                    # 눈금, 범례, 주석 글자 px(글자 상자 높이 28px 이상)

def _c(t): return "#%02x%02x%02x" % t
NAVY = _c((31, 56, 100)); RED = _c((200, 30, 30)); BLUE = _c((0, 90, 190))
INK = _c((20, 20, 20)); GRAY = _c((90, 90, 90)); GRID = _c((120, 130, 150))
STRIPE = _c((244, 247, 252)); GROUPBG = _c((232, 232, 232)); LB = _c((214, 224, 240)); LBTXT = _c((20, 40, 80))
DOT = _c((190, 194, 202))            # 연한 회색 점선 그리드

# ── 글꼴 ──────────────────────────────────────────────────────────
_FONT_OK = []
FONT_FILES = []
CHECKS = []      # save() 가 그림마다 남기는 검사 결과(최소 글자 px, 눈금 글자 높이 px 등)
def setup_font():
    """나눔고딕(보통 NanumGothic.ttf, Bold NanumGothicBold.ttf)만 쓴다. 환경변수 PDF_FONT_REGULAR,PDF_FONT_BOLD > 우분투 fonts-nanum. 없으면 예외."""
    if _FONT_OK: return _FONT_OK[0]
    cands = []
    r, b = os.environ.get("PDF_FONT_REGULAR"), os.environ.get("PDF_FONT_BOLD")
    if r or b: cands.append((r or b, b or r))
    cands += [("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf")]
    for r, b in cands:
        if os.path.exists(r) and os.path.exists(b):
            fm.fontManager.addfont(r)
            if b != r: fm.fontManager.addfont(b)
            FONT_FILES[:] = [r, b]
            name = fm.FontProperties(fname=r).get_name()
            matplotlib.rcParams["font.family"] = name
            matplotlib.rcParams["axes.unicode_minus"] = False
            _FONT_OK.append(name); return name
    raise RuntimeError("한글 글꼴 없음 (fonts-nanum 필요)")

# ── 숫자 표기 ─────────────────────────────────────────────────────
def md(ymd): return f"{int(ymd[4:6])}/{int(ymd[6:])}"
def span(a, b): return f"{md(a)}~{md(b)}"
def clean(s): return s.replace("\u00b7", " ")

def eok(v, sign=True):
    """10억 이상 정수, 10억 미만 소수 한 자리(반올림은 half-up)."""
    r1 = P._half_up(v, 1)
    r = P._half_up(v, 0) if abs(r1) >= 10 else r1
    s = f"{abs(r):,.0f}" if abs(r1) >= 10 else f"{abs(r):.1f}"
    if r == 0: return "0.0" if abs(r1) < 10 else "0"
    return (("+" if r > 0 else "-") if sign else ("-" if r < 0 else "")) + s

def pct2(v): return f"{v:.2f}"
def pp2(v):
    r = P._half_up(v, 2)
    return "0.00" if r == 0 else f"{r:+.2f}"
def sgn_color(v): return RED if v > 0 else (BLUE if v < 0 else GRAY)

# ── 캔버스 ────────────────────────────────────────────────────────
def ink_left(s, px=48):
    """굵은 글꼴로 s 를 x=100 에서 그렸을 때 잉크 왼쪽 끝과 100 의 차이(글자 왼쪽 여백 px)."""
    import numpy as np
    setup_font()
    f = plt.figure(figsize=(12, 1.2), dpi=DPI, facecolor="white")
    f.text(100 / (12 * DPI), 0.5, s, fontsize=px * PT, fontweight="bold", va="center", ha="left")
    buf = io.BytesIO(); f.savefig(buf, format="png", dpi=DPI, facecolor="white"); plt.close(f)
    from PIL import Image
    a = np.asarray(Image.open(io.BytesIO(buf.getvalue())).convert("RGB")).min(axis=2) < 235
    return float(np.where(a.any(axis=0))[0][0]) - 100.0

TITLE_Y = 69.5   # 제목 줄(24~108) 가운데. 글자 잉크 y 45~89

class Canvas:
    def __init__(self):
        setup_font()
        self.fig = plt.figure(figsize=(W / DPI, MAXH / DPI), dpi=DPI, facecolor="white")
        self.ov = self.fig.add_axes([0, 0, 1, 1], zorder=0)
        self.ov.set_xlim(0, W); self.ov.set_ylim(MAXH, 0); self.ov.axis("off")
        self.ov.patch.set_visible(False)
        self.rend = self.fig.canvas.get_renderer()
        self.bottom = 0

    def tw(self, s, px, bold=False):
        t = self.ov.text(0, 0, s, fontsize=px * PT, fontweight="bold" if bold else "normal")
        w = t.get_window_extent(self.rend).width; t.remove(); return w

    def text(self, x, y, s, px, color=INK, ha="left", bold=False, va="center", small=False):
        t = self.ov.text(x, y, s, fontsize=px * PT, color=color, ha=ha, va=va,
                         fontweight="bold" if bold else "normal", zorder=5)
        if small: t.set_gid("small")
        return t

    def rect(self, x0, y0, x1, y1, fc, ec=None, lw=1, z=1):
        self.ov.add_patch(plt.Rectangle((x0, y0), x1 - x0, y1 - y0, facecolor=fc,
                                        edgecolor=ec or "none", linewidth=lw * PT, zorder=z))

    def hline(self, y, x0=M, x1=RIGHT, color=GRID, lw=1, z=3, ls="-"):
        self.ov.plot([x0, x1], [y, y], color=color, lw=lw * PT, zorder=z, ls=ls, solid_capstyle="butt")

    def vline(self, x, y0, y1, color=GRID, lw=1, z=3, ls="-"):
        self.ov.plot([x, x], [y0, y1], color=color, lw=lw * PT, zorder=z, ls=ls, solid_capstyle="butt")

    def axes(self, l, t, w, h):
        ax = self.fig.add_axes([l / W, 1 - (t + h) / MAXH, w / W, h / MAXH], zorder=2)
        ax.patch.set_visible(False)
        return ax

    def title_row(self, candidates, date):
        """제목 줄. 후보 중 날짜와 겹치지 않는 가장 긴 것을 쓴다. 높이 84px(여백 포함)."""
        dw = self.tw(date, 48, True)
        pick = candidates[-1]
        for c in candidates:
            if self.tw(c, 48, True) + dw + 36 <= W - 2 * M: pick = c; break
        self.text(M - ink_left(pick), TITLE_Y, pick, 48, NAVY, bold=True)   # 제목 글자 잉크의 왼쪽 끝이 x=24
        self.text(RIGHT, TITLE_Y, date, 48, NAVY, ha="right", bold=True)
        self.title_used = pick
        return CONTENT0

    def check(self, out, h):
        """글자 크기(28px 이상), ASCII 마이너스, 글자끼리 겹침, 왼쪽 끝 x=24 를 그리기 전에 확인하고 CHECKS 에 남긴다."""
        self.fig.canvas.draw()
        rend = self.fig.canvas.get_renderer()
        items = [(t, "본문") for t in self.ov.texts if t.get_text().strip()]
        for ax in self.fig.axes:
            if ax is self.ov: continue
            for axis in (ax.xaxis, ax.yaxis):
                for tk in axis._update_ticks():
                    if tk.label1.get_visible() and tk.label1.get_text().strip(): items.append((tk.label1, "눈금"))
        fs_min = min(t.get_fontsize() * DPI / 72 for t, _ in items)
        tick_h = [t.get_window_extent(rend).height for t, k in items if k == "눈금"]
        tick_fs = [t.get_fontsize() * DPI / 72 for t, k in items if k == "눈금"]
        if fs_min < 28 - 1e-6: raise ValueError(f"글자 {fs_min:.1f}px < 28px: {out}")
        if tick_h and min(tick_h) < 28: raise ValueError(f"눈금 글자 상자 높이 {min(tick_h):.1f}px < 28px: {out}")
        for t, k in items:
            if t.get_gid() == "small" and t.get_window_extent(rend).height < 28:
                raise ValueError(f"범례, 주석 글자 상자 높이 {t.get_window_extent(rend).height:.1f}px < 28px: {t.get_text()!r}")
        for t, _ in items:
            if "\u2212" in t.get_text(): raise ValueError(f"유니코드 마이너스: {t.get_text()}")
            if "\u00b7" in t.get_text() or "\u2014" in t.get_text(): raise ValueError(f"가운뎃점/줄표: {t.get_text()}")
        boxes = []
        for t, k in items:
            e = t.get_window_extent(rend)
            hh = e.height * 0.15                     # 글줄 여백을 뺀 대략의 잉크 상자
            boxes.append((e.x0, MAXH - e.y1 + hh, e.x1, MAXH - e.y0 - hh, t.get_text(), k))
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a_, b_ = boxes[i], boxes[j]
                if a_[0] < b_[2] - 1 and b_[0] < a_[2] - 1 and a_[1] < b_[3] - 1 and b_[1] < a_[3] - 1:
                    raise ValueError(f"글자 겹침 {a_[4]!r}/{b_[4]!r} in {out}")
        lo = min(b_[0] for b_ in boxes); hi = max(b_[2] for b_ in boxes)
        CHECKS.append(dict(file=os.path.basename(out), fs_min=round(fs_min, 1), tick_fs_min=round(min(tick_fs), 1) if tick_fs else None,
                           tick_h_min=round(min(tick_h), 1) if tick_h else None, left=round(lo, 1), right=round(hi, 1),
                           title=getattr(self, "title_used", None),
                           plot_left=round(min(ax.get_position().x0 * W for ax in self.fig.axes if ax is not self.ov), 1) if len(self.fig.axes) > 1 else None))

    def save(self, out, bottom):
        """bottom(px) 아래 24px 여백까지 자르고 저장. 1920 초과면 예외."""
        h = int(round(bottom + M))
        if h > MAXH: raise ValueError(f"그림 높이 {h}px > {MAXH}px")
        self.check(out, h)
        buf = io.BytesIO()
        self.fig.savefig(buf, format="png", dpi=DPI, facecolor="white")
        plt.close(self.fig)
        from PIL import Image
        im = Image.open(io.BytesIO(buf.getvalue())).convert("RGB")
        assert im.size[0] == W, im.size
        # 아래 여백 24px 를 잉크(채널 하나라도 235 미만인 픽셀) 기준으로 맞춘다
        import numpy as np
        a = np.asarray(im).min(axis=2) < 235
        last = int(np.where(a.any(axis=1))[0].max())
        h = last + 1 + M
        if h > MAXH: raise ValueError(f"그림 높이 {h}px > {MAXH}px")
        im.crop((0, 0, W, h)).save(out)
        return out

GROUP_SHORT = {"ESS 2차전지": "ESS"}

def group_titles(label, tail, part):
    """제목 꼴은 네 그룹 모두 "{그룹} 액티브 {꼬리}{ (k/N)}". 넘치면 꼬리가 아니라 그룹 이름을 줄인다(ESS 2차전지 -> ESS)."""
    g = clean(label.replace(" 액티브", ""))
    g = GROUP_SHORT.get(g, g)       # "ESS 2차전지 액티브 종목별 순매수 (2/3)" 가 날짜와 겹쳐서 ESS 는 앨범 세 장 모두 "ESS" 로 통일
    return [f"{g} 액티브 {tail}" + suffix(part)]

def suffix(part):
    i, n = part
    return f" ({i}/{n})" if n > 1 and i > 1 else ""

def style_axes(ax, ytick_fmt=None):
    for s in ("top", "right"): ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color(GRID); ax.spines[s].set_linewidth(1 * PT)
    ax.tick_params(colors=GRAY, labelsize=SMALL * PT, length=0, pad=6)
    ax.grid(True, axis="y", color=DOT, lw=1.3 * PT, ls=(0, (1.5, 3)))
    ax.set_axisbelow(True)
    if ytick_fmt: ax.yaxis.set_major_formatter(FuncFormatter(ytick_fmt))

def nice_step(span_v, target=4):
    raw = span_v / target
    for s in (0.5, 1, 2, 2.5, 5, 10, 20, 25, 50, 100, 200, 250, 500, 1000, 2000, 5000):
        if s >= raw: return s
    return raw

def xticks_idx(n, step=3):
    return list(range(n - 1, -1, -step))[::-1]

# 분류 구분(패시브, 액티브, 그룹 등): 남색 실선, 회색 실선, 회색 점선, 남색 점선 순서
DASH = (0, (7, 3))
CAT_STYLES = [(NAVY, "-"), (GRAY, "-"), (GRAY, DASH), (NAVY, DASH)]
def cat_style(j): return CAT_STYLES[j % len(CAT_STYLES)]

def ls_for(i, color):
    return [(color, "-"), (color, (0, (7, 3))), (color, (0, (1.2, 2.2))), (color, (0, (7, 2.5, 1.2, 2.5)))][i % 4]

# ══ ① 통합 자금흐름 ═══════════════════════════════════════════════
def chart_flows(F, days, out, cmp_days=5, part=(1, 1)):
    cats = list(F)
    i0 = max(0, len(days) - cmp_days)
    wk = {c: sum(F[c]["flow"][i0:]) for c in cats}
    cv = Canvas()
    y = cv.title_row(["자금 유출입과 순자산"], span(days[i0], days[-1]))
    # 표 머리(글자 36 굵게, 오른쪽 안쪽 여백 24), 본문 40px, 행 64px
    HH, RH, BODY, HEAD, PADR = 72, 64, 40, 36, 24
    groups = []
    for pref in ("패시브", "액티브"):
        items = [c for c in cats if c.startswith(pref)]
        if items: groups.append((pref, items))
    nm = lambda c: clean(c.split("·", 1)[1]) if "·" in c else clean(c)
    namew = max(cv.tw(nm(c), BODY) for c in cats) + 36 + 16
    aum_w = max(max(cv.tw(f"{F[c]['aum'][-1]:,.0f}", BODY) for c in cats), cv.tw("순자산(억)", HEAD, True)) + PADR + 16
    bx0 = M + namew; bx1 = RIGHT - aum_w
    cv.rect(M, y, RIGHT, y + HH, NAVY, z=2)
    cv.text(M + 12, y + HH / 2, "구분", HEAD, "white", bold=True)
    cv.text((bx0 + bx1) / 2, y + HH / 2, "순유출입(억)", HEAD, "white", ha="center", bold=True)
    cv.text(RIGHT - PADR, y + HH / 2, "순자산(억)", HEAD, "white", ha="right", bold=True)
    y += HH
    lo = min(list(wk.values()) + [0]); hi = max(list(wk.values()) + [0])
    lab = max([cv.tw("+000", BODY, True)] + [cv.tw(eok(v), BODY, True) for v in wk.values()]) + 14   # 11-28: '+2,042' 처럼 네 자리 이상이면 칸 경계를 넘던 것 고침
    zx0, zx1 = bx0 + lab + 6, bx1 - lab - 6
    scale = (zx1 - zx0) / ((hi - lo) or 1)
    zero = zx0 + (0 - lo) * scale
    top_rows = y
    k = 0
    for pref, items in groups:
        cv.rect(M, y, RIGHT, y + 56, GROUPBG, z=1)
        cv.text(M + 12, y + 28, pref, HEAD, INK, bold=True)
        cv.hline(y + 56); y += 56
        for c in items:
            if k % 2 == 1: cv.rect(M, y, RIGHT, y + RH, STRIPE, z=1)
            k += 1
            v = wk[c]; cy = y + RH / 2
            cv.text(M + 36, cy, nm(c), BODY, INK)
            x2 = zero + v * scale
            cv.rect(min(zero, x2), cy - 18, max(zero, x2), cy + 18, RED if v >= 0 else BLUE, z=4)
            if v >= 0: cv.text(x2 + 8, cy, eok(v), BODY, RED, bold=True)
            else: cv.text(x2 - 8, cy, eok(v), BODY, BLUE, ha="right", bold=True)
            cv.text(RIGHT - PADR, cy, f"{F[c]['aum'][-1]:,.0f}", BODY, INK, ha="right")
            cv.hline(y + RH); y += RH
    cv.vline(zero, top_rows, y, color=INK, lw=1.5, z=3)
    cv.vline(bx0, top_rows - HH, y); cv.vline(bx1, top_rows - HH, y)
    cv.rect(M, top_rows - HH, RIGHT, y, "none", GRID, 1, z=6)
    y += 34
    # 누적 순유출입 두 패널: 패널 제목은 띠 없이 남색 굵은 30px, 범례 한 줄은 x=24 에서 시작
    n = len(days); xs = list(range(n))
    for pref, items in groups:
        cv.text(M, y + 18, f"{pref} 누적 순유출입(억)", 30, NAVY, bold=True)
        cv.text(RIGHT, y + 18, span(days[0], days[-1]), 30, NAVY, ha="right", bold=True)
        y += 44
        series = []
        for j, c in enumerate(items):
            acc = 0.0; cum = []
            for v in F[c]["flow"]: acc += v; cum.append(acc)
            series.append((nm(c), cum, j))
        # 범례 위 한 줄(이름과 마지막 값). 넘치면 값을 뺀다
        def legend(withval):
            ent = [(f"{a} {eok(cum[-1])}" if withval else a) for a, cum, j in series]
            return ent, sum(54 + cv.tw(e, SMALL) for e in ent) + 24 * (len(ent) - 1)
        ent, tot = legend(True)
        if tot > W - 2 * M: ent, tot = legend(False)
        x = M
        for (a, cum, j), e in zip(series, ent):
            col, ls = cat_style(j)
            cv.ov.plot([x, x + 44], [y + 16, y + 16], color=col, lw=4.5 * PT, ls=ls, zorder=5)
            cv.text(x + 54, y + 16, e, SMALL, INK, small=True)
            x += 54 + cv.tw(e, SMALL) + 24
        y += 40
        ph = 330
        ax = cv.axes(M + GUT, y + 6, RIGHT - M - GUT - 24, ph)
        style_axes(ax, lambda v, p: f"{v:,.0f}")
        ax.axvspan(i0 - 0.5, n - 0.5, color=LB, zorder=0, lw=0)
        ax.axhline(0, color=GRID, lw=1 * PT, zorder=1)
        allv = []
        for a, cum, j in series:
            col, ls = cat_style(j)
            ax.plot(xs, cum, color=col, lw=4.5 * PT, ls=ls, zorder=3, solid_capstyle="butt"); allv += cum
        ax.set_xlim(-0.5, n - 0.5)
        a, b = min(allv + [0]), max(allv + [0]); pad = (b - a) * 0.08 or 1
        ax.set_ylim(a - pad, b + pad)
        ax.yaxis.set_major_locator(matplotlib.ticker.MaxNLocator(4))
        ti = xticks_idx(n, 3)
        ax.set_xticks(ti); ax.set_xticklabels([md(days[i]) for i in ti])
        y += 6 + ph + 48 + 26
    return cv.save(out, y - 26)

# ══ ② 보유비중 변화 (덤벨 표 + 테마는 소형다중) ══════════════════════
RING, DOTD = 30, 20      # 시작일 고리, 현재 점 지름(px). 범례도 같은 값

def _axis_ticks(vmax):
    step = nice_step(vmax, 4)
    top = step * (int(vmax / step) + 1)
    return step, top

def chart_weights(weights, rows, out, label, cmp_days=5, key="stock", order=None, topn=15, part=(1, 1)):
    """key='stock' 종목 상위 topn, 'theme' 테마(+소형다중). 반환: out 경로."""
    days, Wt, H = weights
    i0 = max(0, len(days) - 1 - cmp_days)
    win = set(days[i0 + 1:])
    flow = defaultdict(float)
    fk = "테마" if key == "theme" else "종목"
    for r in rows:
        if r["일자"] in win: flow[r[fk]] += r["순매수금액(원)"] / EOK
    if key == "theme":
        names = [g for g in (order or sorted(Wt[-1], key=lambda k: -Wt[-1].get(k, 0))) if g in Wt[-1]]
    else:
        names = sorted(Wt[-1], key=lambda k: -Wt[-1][k])[:topn]
    seq = sorted(names, key=lambda g: -Wt[-1].get(g, 0))
    short = label.replace(" 액티브", "")
    cv = Canvas()
    y = cv.title_row(group_titles(label, "보유비중", part), span(days[i0], days[-1]))
    # 범례 한 줄: 시작 점, 현재 점
    # 범례도 그림과 같은 모양: 속 빈 회색 고리 = 시작일, 빨강 점 = 현재(증가), 파랑 점 = 현재(감소)
    ly = y + 17
    lx = M + RING / 2
    cv.ov.scatter([lx], [ly], s=(RING * PT) ** 2, facecolors="none", edgecolors=GRAY, linewidths=3 * PT, zorder=6)
    cv.text(lx + 26, ly, md(days[i0]), SMALL, GRAY, small=True)
    lx += 26 + cv.tw(md(days[i0]), SMALL) + 40
    for colr, nm_ in ((RED, "증가"), (BLUE, "감소")):
        cv.ov.scatter([lx], [ly], s=(DOTD * PT) ** 2, color=colr, zorder=6)
        t_ = f"{md(days[-1])} {nm_}"
        cv.text(lx + 18, ly, t_, SMALL, GRAY, small=True)
        lx += 18 + cv.tw(t_, SMALL) + 40
    y += 44
    # 열 배치
    BF = 34
    namew = max(cv.tw(clean(g), BF) for g in seq) + 12 + 16
    c3r = RIGHT - 12; c3l = c3r - max(cv.tw("+000.0", BF), cv.tw("순매수", 30, True)) - 14
    c2r = c3l - 22; c2l = c2r - max(cv.tw("+0.00", BF), cv.tw("변화", 30, True)) - 14
    c1r = c2l - 22; c1l = c1r - max(cv.tw("00.00", BF), cv.tw("현재", 30, True)) - 14
    px0 = M + namew; px1 = c1l - 14
    HH = 92; RH = 64
    cv.rect(M, y, RIGHT, y + HH, NAVY, z=2)
    cv.text(M + 12, y + HH / 2, "테마" if key == "theme" else "종목", 30, "white", bold=True)
    cv.text((px0 + c1l) / 2, y + HH / 2, "보유비중", 30, "white", ha="center", bold=True)
    for r_, a, b in ((c1r, "현재", "(%)"), (c2r, "변화", "(%p)"), (c3r, "순매수", "(억)")):
        cv.text(r_, y + HH / 2 - 17, a, 30, "white", ha="right", bold=True)
        cv.text(r_, y + HH / 2 + 17, b, 30, "white", ha="right", bold=True)
    y0 = y; y += HH
    vmax = max(max(Wt[i0].get(g, 0), Wt[-1].get(g, 0)) for g in seq)
    step, top = _axis_ticks(vmax * 1.02)
    # 눈금 글자(28px)가 서로 겹치지 않게 간격을 넓힌다
    while True:
        nt = int(round(top / step))
        if nt < 2 or (px1 - px0 - 28) / nt >= cv.tw(f"{top:g}%", SMALL) + 22: break
        step *= 2; top = step * (int(vmax * 1.02 / step) + 1)
    # 마커가 끝에 닿지 않게 약간 안쪽
    lo_pad, hi_pad = 14, 14
    mx = lambda v: px0 + lo_pad + v / top * (px1 - px0 - lo_pad - hi_pad)
    rows_top = y
    for i, g in enumerate(seq):
        if i % 2 == 1: cv.rect(M, y, RIGHT, y + RH, STRIPE, z=1)
        a, b = Wt[i0].get(g, 0), Wt[-1].get(g, 0); cy = y + RH / 2
        col = sgn_color(b - a) if abs(b - a) >= 0.005 else GRAY
        cv.text(M + 12, cy, clean(g), BF, INK)
        cv.ov.plot([mx(a), mx(b)], [cy, cy], color=col, lw=6 * PT, zorder=4, solid_capstyle="round")
        cv.ov.scatter([mx(b)], [cy], s=(DOTD * PT) ** 2, color=col, zorder=5)
        cv.ov.scatter([mx(a)], [cy], s=(RING * PT) ** 2, facecolors="none", edgecolors=GRAY, linewidths=3 * PT, zorder=6)
        cv.text(c1r, cy, pct2(b), BF, INK, ha="right")
        cv.text(c2r, cy, pp2(b - a), BF, sgn_color(P._half_up(b - a, 2)), ha="right", bold=True)
        f = flow.get(g, 0.0)
        cv.text(c3r, cy, eok(f) if abs(f) >= 0.05 else "0.0", BF, sgn_color(f) if abs(f) >= 0.05 else GRAY, ha="right")
        cv.hline(y + RH); y += RH
    # 가로 눈금: 점선 세로 그리드와 아래 눈금 글자
    t = 0.0
    while t <= top + 1e-9:
        xx = mx(t)
        cv.vline(xx, rows_top, y, color=DOT, lw=1.3, z=2, ls=(0, (1.5, 3)))
        cv.text(xx, y + 24, f"{t:g}%", SMALL, GRAY, ha="center", small=True)
        t += step
    for xx in (px0, c1l, c2l, c3l): cv.vline(xx, y0, y)
    cv.rect(M, y0, RIGHT, y, "none", GRID, 1, z=6)
    y += 58
    if key == "theme":
        y += 22
        ncol = 2; gap = 24; pw = (W - 2 * M - gap) // ncol; PH = 300; TH = 48
        x_idx = list(range(len(days)))
        for i, g in enumerate(names):
            c = [Wt[j].get(g, 0) for j in x_idx]
            r_, cc = divmod(i, ncol)
            px_ = M + cc * (pw + gap); py_ = y + r_ * (PH + 22)
            d = c[-1] - c[i0]
            col = sgn_color(P._half_up(d, 2)) if abs(d) >= 0.005 else GRAY
            cv.text(px_ + 10, py_ + TH / 2, f"{clean(g)} {pct2(c[-1])}%", 30, NAVY, bold=True)
            cv.text(px_ + pw - 10, py_ + TH / 2, f"{pp2(d)}%p", 30, col, ha="right", bold=True)
            ax = cv.axes(px_ + GUT, py_ + TH + 10, pw - GUT - 18, PH - TH - 10 - 46)
            style_axes(ax, lambda v, p: f"{v:.1f}")
            ax.axvspan(i0, len(days) - 1, color=LB, zorder=0, lw=0)
            ax.plot(x_idx, c, color=col if col != GRAY else NAVY, lw=4.5 * PT, zorder=3)
            ax.set_xlim(0, len(days) - 1)
            ax.yaxis.set_major_locator(matplotlib.ticker.MaxNLocator(3))
            ti = sorted({0, len(days) - 1})      # 비교 구간 시작일 눈금은 끝 눈금과 붙어서 뺌(바탕색 띠로 구간은 보임)
            ax.set_xticks(ti); ax.set_xticklabels([md(days[j]) for j in ti])
            tl = ax.get_xticklabels(); tl[0].set_ha("left"); tl[-1].set_ha("right")
            cv.rect(px_, py_, px_ + pw, py_ + PH, "none", GRID, 1, z=6)
        nrow = (len(names) + ncol - 1) // ncol
        y += nrow * (PH + 22)
        y -= 22
    return cv.save(out, y - 20 if key != "theme" else y)

# ══ ③ 종목별 순매수 (막대 표 + 누적 추이) ═══════════════════════════
def _by_stock(rows):
    d = defaultdict(lambda: defaultdict(float))
    for r in rows: d[r["종목"]][r["일자"]] += r["순매수금액(원)"] / EOK
    return d

NB_ROW = 54; NB_BAND = 52; NB_FS = 36   # 40px/행 64 는 15위x2 표가 1920px 를 넘어 못 씀(README 3판)
def _nb_per_max():
    # 108(제목) + 2*(NB_BAND + n*NB_ROW) + 22(간격) + 24(아래 여백) <= 1920
    return (MAXH - CONTENT0 - 22 - M - 2 * NB_BAND) // (2 * NB_ROW)

def netbuy_pages(rows, topn=15):
    """막대 표가 몇 장이 되는지(앨범 번호 계산용)."""
    bystk = _by_stock(rows)
    n = min(topn, len(bystk))
    return max(1, -(-n // _nb_per_max()))

def chart_netbuy(rows, out_prefix, label, topn=15, part_start=2, part_total=3):
    """종목별 순매수 상위, 하위 표. topn 이 커서 1920px 를 넘으면 순위 단위로 나눠 여러 장. 경로 목록 반환."""
    bystk = _by_stock(rows)
    days = sorted({r["일자"] for r in rows})
    tot = {k: sum(v.values()) for k, v in bystk.items()}
    buys = sorted(tot.items(), key=lambda kv: -kv[1])[:topn]
    sells = sorted(tot.items(), key=lambda kv: kv[1])[:topn]
    per = _nb_per_max()
    pages = max(1, -(-max(len(buys), len(sells)) // per))
    short = label.replace(" 액티브", "")
    outs = []
    for pg in range(pages):
        sl = slice(pg * per, (pg + 1) * per)
        cv = Canvas()
        idx = part_start + pg
        y = cv.title_row(group_titles(label, "종목별 순매수", (idx, part_total)), span(days[0], days[-1]))
        bl, sl_ = buys[sl], sells[sl]
        nw = max([cv.tw(clean(k), NB_FS) for k, v in bl + sl_] + [0]) + 12 + 14
        vw = cv.tw("+000", NB_FS, True) + 14
        bx0 = M + nw; bx1 = RIGHT - 16 - vw
        mx = max([abs(v) for k, v in buys + sells] + [1e-9])      # 장이 나뉘어도 막대 길이 기준은 전체 순위 공통
        for data, name, color in ((bl, "순매수 상위", RED), (sl_, "순매도 상위", BLUE)):
            if not data: continue
            r0 = pg * per + 1
            cv.rect(M, y, RIGHT, y + NB_BAND, LB, z=1)
            cv.text(M + 12, y + NB_BAND / 2, f"{name} {r0}~{r0 + len(data) - 1}위 (억)" if pages > 1 else f"{name} {len(data)} (억)", 34, LBTXT, bold=True)
            y += NB_BAND; ytop = y
            for i, (k, v) in enumerate(data):
                if i % 2 == 1: cv.rect(M, y, RIGHT, y + NB_ROW, STRIPE, z=1)
                cy = y + NB_ROW / 2
                cv.text(M + 12, cy, clean(k), NB_FS, INK)
                L = abs(v) / mx * (bx1 - bx0)
                cv.rect(bx0, cy - 16, bx0 + L, cy + 16, color, z=4)
                cv.text(bx0 + L + 8, cy, eok(v), NB_FS, color, bold=True)
                y += NB_ROW
            cv.hline(y); cv.rect(M, ytop - NB_BAND, RIGHT, y, "none", GRID, 1, z=6)
            cv.vline(bx0, ytop, y)
            y += 22
        o = out_prefix if pages == 1 else out_prefix.replace(".png", f"_p{pg + 1}.png")
        outs.append(cv.save(o, y - 22))
    return outs

def chart_trend(rows, out, label, part=(3, 3), npick=4):
    """누적 순매수 추이: 매수 상위 4 + 매도 상위 4. 선 끝에 이름과 값."""
    bystk = _by_stock(rows)
    days = sorted({r["일자"] for r in rows}); n = len(days); xs = list(range(n))
    tot = {k: sum(v.values()) for k, v in bystk.items()}
    buys = sorted(tot.items(), key=lambda kv: -kv[1])[:npick]
    sells = sorted(tot.items(), key=lambda kv: kv[1])[:npick]
    short = label.replace(" 액티브", "")
    cv = Canvas()
    y = cv.title_row(group_titles(label, "순매수 추이", part), span(days[0], days[-1]))
    cv.text(M, y + 15, "누적 순매수 추이(억)", 30, NAVY, bold=True)
    y += 44
    series = []
    for i, (nm, _) in enumerate(buys):
        acc = 0.0; c = []
        for d in days: acc += bystk[nm].get(d, 0.0); c.append(acc)
        series.append((nm, c, RED, i))
    for i, (nm, _) in enumerate(sells):
        acc = 0.0; c = []
        for d in days: acc += bystk[nm].get(d, 0.0); c.append(acc)
        series.append((nm, c, BLUE, i))
    labs = [f"{clean(nm)} {eok(c[-1])}" for nm, c, col, i in series]
    labw = max(cv.tw(s, SMALL, True) for s in labs) + 22
    ph = 820
    l = M + GUT; r = RIGHT - labw
    ax = cv.axes(l, y, r - l, ph)
    style_axes(ax, lambda v, p: f"{v:,.0f}")
    ax.axhline(0, color=GRID, lw=1 * PT, zorder=1)
    allv = []
    for nm, c, col, i in series:
        colr, ls = ls_for(i, col)
        ax.plot(xs, c, color=colr, lw=4.5 * PT, ls=ls, zorder=3, solid_capstyle="butt"); allv += c
    a, b = min(allv + [0]), max(allv + [0]); pad = (b - a) * 0.06 or 1
    ax.set_ylim(a - pad, b + pad); ax.set_xlim(0, n - 1)
    ax.yaxis.set_major_locator(matplotlib.ticker.MaxNLocator(5))
    ti = xticks_idx(n, 3); ax.set_xticks(ti); ax.set_xticklabels([md(days[i]) for i in ti])
    # 선 끝 이름표: 겹치면 px 단위로 밀어 낸다
    ylo, yhi = a - pad, b + pad
    ypx = lambda v: y + (yhi - v) / (yhi - ylo) * ph
    items = sorted([(ypx(c[-1]), s, col) for (nm, c, col, i), s in zip(series, labs)], key=lambda t: t[0])
    g = 36; ys = [t[0] for t in items]
    for k in range(1, len(ys)): ys[k] = max(ys[k], ys[k - 1] + g)
    over = ys[-1] - (y + ph)
    if over > 0:
        ys = [v - over for v in ys]
        for k in range(len(ys) - 2, -1, -1): ys[k] = min(ys[k], ys[k + 1] - g)
    for (oy, s, col), yy in zip(items, ys):
        if abs(oy - yy) > 6: cv.ov.plot([r + 2, r + 12], [oy, yy], color=GRID, lw=1 * PT, zorder=4)
        cv.text(r + 16, yy, s, SMALL, col, bold=True, small=True)
    y += ph + 52
    return cv.save(out, y)
