# -*- coding: utf-8 -*-
"""잡 끝에 리포트·스냅샷을 레포에 반영한다(.github/workflows/daily.yml 의 커밋 단계).

왜 단순 'git commit && git push'가 아닌가:
  같은 concurrency 그룹에서 대기열에 섰던 회차는 '생성 시점'의 커밋을 checkout 한다. 기다리는 동안
  앞 회차가 push 하면 이 회차의 push 는 non-fast-forward 로 거부되고(2026-10-21 #707), 그 회차의
  캡처·리포트·발송 마커가 통째로 사라진다.
방법: 이번 실행이 바꾼 파일(reports/·snapshots.json)을 따로 떠 둔다 → 최신 origin/main 으로 맞춘다 →
  다시 얹는다. snapshots.json 은 한 줄짜리 JSON 이라 git 병합이 통째로 충돌하므로 직접 합집합
  (ETF·날짜별, 같은 날짜는 이번 실행 우선)으로 합친다. push 가 또 거부되면(그 사이 다른 push) 처음부터 재시도.
"""
import json, os, shutil, subprocess, sys, tempfile, time
from datetime import datetime

PATHS = ("reports", "snapshots.json")
TRIES = 4


def git(*args, check=True):
    r = subprocess.run(("git", "-c", "core.quotepath=false") + args,
                       capture_output=True, text=True, encoding="utf-8")
    if check and r.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} → {r.returncode}: {r.stderr.strip()[:300]}")
    return r


def changed_files():
    """이번 실행이 만들거나 바꾼 파일. .gitignore 가 *.xlsx·*.png·*.json 을 무시하므로 --ignored 로 새 파일도 잡는다."""
    out = git("status", "--porcelain", "--ignored", "--untracked-files=all", "--", *PATHS).stdout
    files = []
    for line in out.splitlines():
        code, path = line[:2], line[3:]
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        path = path.strip().strip('"')
        if "D" in code or not os.path.isfile(path):
            continue                        # 이 잡은 파일을 지우지 않는다
        files.append(path)
    return files


def merge_snapshots(ours_path, base_path):
    """base(원격 최신) ∪ ours(이번 실행). 같은 ETF·같은 날짜는 ours 우선."""
    try:
        with open(base_path, encoding="utf-8") as f: base = json.load(f)
    except (OSError, ValueError):
        base = {}
    with open(ours_path, encoding="utf-8") as f: ours = json.load(f)
    for key, days in ours.items():
        base.setdefault(key, {}).update(days)
    tmp = base_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f: json.dump(base, f, ensure_ascii=False)
    os.replace(tmp, base_path)


def main():
    files = changed_files()
    if not files:
        print("변경 없음(휴장/미게시)")
        return 0
    stash = tempfile.mkdtemp(prefix="etf_out_")
    for p in files:
        dst = os.path.join(stash, p)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(p, dst)
    msg = f"리포트 {datetime.now():%Y-%m-%d_%H%M}"      # 워크플로가 TZ=Asia/Seoul 을 준다
    for attempt in range(1, TRIES + 1):
        git("fetch", "--quiet", "--depth=1", "origin", "main")
        git("reset", "--quiet", "--hard", "origin/main")   # 추적 파일만 원격으로 — 새 파일(무시·미추적)은 그대로 남는다
        for p in files:
            src = os.path.join(stash, p)
            if p == "snapshots.json" and os.path.exists(p):
                merge_snapshots(src, p)
            else:
                os.makedirs(os.path.dirname(p) or ".", exist_ok=True)
                shutil.copy2(src, p)
        git("add", "-f", "--", *files)
        if git("diff", "--cached", "--quiet", check=False).returncode == 0:
            print("변경 없음(원격과 동일)")
            return 0
        git("commit", "--quiet", "-m", msg)
        r = git("push", "--quiet", "origin", "HEAD:main", check=False)
        if r.returncode == 0:
            print(f"push 완료 — 파일 {len(files)}개" + (f", 재시도 {attempt - 1}회" if attempt > 1 else ""))
            return 0
        print(f"push 거부({attempt}/{TRIES}): {r.stderr.strip()[:200]} → 최신 원격 위에 다시 얹어 재시도")
        time.sleep(3 * attempt)
    print("push 실패 — 이번 회차 산출물은 레포에 남기지 못함")
    return 1


if __name__ == "__main__":
    sys.exit(main())
