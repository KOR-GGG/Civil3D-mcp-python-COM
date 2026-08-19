"""trial_guard.py — 시행 중 정답이 실행 환경의 사정거리 안에 있지 않게 한다.

왜 필요한가
-----------
제5장 5.3.4 의 조건 ② — **정답 키가 실행 환경의 사정거리 밖에 있어야 한다.**
2026-08-17 의 첫 시행을 폐기한 이유가 이것이고, 2026-08-19 의 C4 두 시행에서도
대화에 파일·터미널 접근이 열려 있어("14개 파일 읽음, 명령 12개 실행함") 같은
위험 구간에 들어갔다. 그때는 답변이 정답표와 다른 값을 내어 유효로 판정했으나,
조건을 깨끗이 두는 편이 심사에서 방어하기 쉽다.

호스트의 도구 설정을 끄는 방법도 있으나 UI 에 의존하고 시행마다 확인이 어렵다.
이 스크립트는 **읽힐 것을 치우는** 쪽을 택한다. 파일 시스템에서 사라지므로
파일 접근이 열려 있어도 정답에 닿지 못하고, 상태를 명령으로 확인할 수 있다.

무엇을 가리는가
---------------
정답 수치·환경별 미끼 성격·판정 기준을 담은 파일들이다. 특히
`make_test_surfaces.py` 의 기대값 사전은 층별 판정 기준과 환경별 미끼의 성격까지
적고 있어 노출도가 가장 크다.

사용법
------
    .\.venv\Scripts\python.exe experiments\trial_guard.py hide      # 시행 직전
    .\.venv\Scripts\python.exe experiments\trial_guard.py status    # 확인
    .\.venv\Scripts\python.exe experiments\trial_guard.py restore   # 시행 직후

⚠ 가린 동안에는 `make_test_surfaces.py` 를 쓸 수 없다. 도면 생성은 hide 전에
끝내 둘 것. `activate_env.py`·`provenance_marker.py` 는 가리지 않으므로 표식
심기·환경 확인은 그대로 된다.

⚠ 가린 동안 git 은 해당 파일을 삭제로 본다. 그 상태로 커밋하지 말 것.
restore 하면 원상 복구된다.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:                                              # noqa: BLE001
        pass

REPO = Path(__file__).resolve().parents[1]

# 보관 위치는 **저장소 밖**이어야 한다. 저장소 안에 두면 파일 접근이 열린 대화가
# 그대로 읽을 수 있어 가리는 의미가 없다.
STASH = REPO.parent / "_trial_stash"

# 2026-08-19 실측으로 뽑은 목록. 정답 수치·판정 기준·환경별 미끼 성격을 담는다.
GUARDED = [
    "experiments/기록지_D.md",          # 판정 기준표 (노출도 최대)
    "experiments/PROTOCOL_D.md",        # 정답과 실행 절차
    "make_test_surfaces.py",            # 케이스별 기대값 사전 + 미끼 성격
    "experiments/harness_d.py",         # 기대값 참조
    "experiments/exp_c5.py",
    "experiments/exp_sensitivity.py",
    "experiments/c5_result.json",
    "experiments/robustness_result.json",
    "experiments/sensitivity_result.json",
    "experiments/reproducibility_result.json",
    "test_earthwork_core.py",
    "test_hydraulics_core.py",
    ".pytest_cache",                    # 시험 이름에 기대값이 남는다
]


def _pairs() -> list[tuple[Path, Path]]:
    return [(REPO / rel, STASH / rel) for rel in GUARDED]


def hide() -> int:
    STASH.mkdir(parents=True, exist_ok=True)
    moved, missing = [], []
    for src, dst in _pairs():
        if not src.exists():
            missing.append(src.relative_to(REPO))
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src), str(dst))
        moved.append(src.relative_to(REPO))
    print("가림 완료 — %d개 파일을 저장소 밖으로 옮겼다." % len(moved))
    for m in moved:
        print("   -", m)
    if missing:
        print("\n⚠ 원본이 없어 건너뛴 것 %d개 (이미 가려져 있을 수 있다):" % len(missing))
        for m in missing:
            print("   -", m)
    print("\n보관 위치 :", STASH)
    print("⚠ 시행이 끝나면 반드시  trial_guard.py restore  를 실행할 것.")
    return 0


def restore() -> int:
    if not STASH.exists():
        print("보관 폴더가 없다. 가려진 파일이 없는 것으로 본다:", STASH)
        return 0
    back, missing = [], []
    for src, dst in _pairs():
        if not dst.exists():
            missing.append(src.relative_to(REPO))
            continue
        src.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(dst), str(src))
        back.append(src.relative_to(REPO))
    print("복구 완료 — %d개 파일을 되돌렸다." % len(back))
    for b in back:
        print("   -", b)
    if missing:
        print("\n(보관본이 없어 건너뜀 %d개 — 원래 자리에 있을 것이다)" % len(missing))
    # 빈 폴더 정리
    for p in sorted(STASH.rglob("*"), key=lambda x: -len(x.parts)):
        if p.is_dir() and not any(p.iterdir()):
            p.rmdir()
    if STASH.exists() and not any(STASH.iterdir()):
        STASH.rmdir()
        print("\n보관 폴더를 지웠다.")
    return 0


def status() -> int:
    at_repo = [rel for rel, _ in zip(GUARDED, _pairs()) if (REPO / rel).exists()]
    at_stash = [rel for rel in GUARDED if (STASH / rel).exists()]
    print("보관 위치 :", STASH, "(있음)" if STASH.exists() else "(없음)")
    print("저장소에 있는 파일 : %d / %d" % (len(at_repo), len(GUARDED)))
    print("가려진 파일        : %d / %d" % (len(at_stash), len(GUARDED)))
    if at_repo and not at_stash:
        print("\n=> 지금은 **가려지지 않은 상태**다. 시행 전에 hide 를 실행할 것.")
    elif at_stash and not at_repo:
        print("\n=> 지금은 **가려진 상태**다. 시행이 끝나면 restore 를 실행할 것.")
    else:
        print("\n⚠ 절반만 옮겨진 상태다. restore 로 정리한 뒤 다시 hide 할 것.")
        for rel in GUARDED:
            a = (REPO / rel).exists()
            b = (STASH / rel).exists()
            print("   %-42s 저장소 %s / 보관 %s" % (rel, "O" if a else "-", "O" if b else "-"))
    return 0


def main() -> int:
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    fn = {"hide": hide, "restore": restore, "status": status}.get(cmd)
    if fn is None:
        print("사용법: trial_guard.py [hide|restore|status]")
        return 2
    return fn()


if __name__ == "__main__":
    raise SystemExit(main())
