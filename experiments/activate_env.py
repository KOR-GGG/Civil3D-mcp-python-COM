"""activate_env.py — D 실험 한 시행의 준비를 한 명령으로 끝낸다

무엇을 하는가
-------------
    1) 지정한 환경의 도면을 열고 활성화한다 (닫혀 있으면 연다)
    2) 서피스 목록을 출력해 **눈으로 확인**할 수 있게 한다
    3) 무작위 **표식 서피스**를 심는다 (도구 사용의 사후 확인용, §4.6.6 조건 3)
    4) 시행 전 점검 3항목을 출력한다

왜 한 명령인가
--------------
n=5 × 3 환경이면 시행이 열다섯 번이다. 명령을 두세 개 순서대로 치게 하면
언젠가 순서를 틀리고, 그 시행은 조용히 무효가 된다. 준비를 한 번에 끝내고
사람은 **확인만** 하게 한다.

사용
----
    .\\.venv\\Scripts\\python.exe experiments\\activate_env.py 3     # 환경 3 준비
    .\\.venv\\Scripts\\python.exe experiments\\activate_env.py       # 현재 상태만 확인
    .\\.venv\\Scripts\\python.exe experiments\\activate_env.py clean # 시행 후 표식 제거

★ 2026-08-17 실측: **서버는 활성 도면을 추종한다.** 매 호출마다 ActiveDocument
를 다시 읽으므로 도면만 바꾸면 되고 **Claude Desktop 재시작은 불필요하다.**
"""
from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

import win32com.client as w32

# 2026-08-19 추가: 한글 Windows 콘솔은 기본이 cp949 라, 아래 출력에 쓰이는
# em-dash 나 경고 기호 하나 때문에 UnicodeEncodeError 로 스크립트가 즉사한다.
# (setup_check.py 에서 같은 결함을 고쳤으나 실험 스크립트에는 남아 있었다.)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:                                          # noqa: BLE001
        pass

# 하위 프로세스(표식 스크립트)의 출력과 순서가 뒤섞이지 않게 줄 단위로 흘린다.
try:
    sys.stdout.reconfigure(line_buffering=True)
except Exception:                                               # noqa: BLE001
    pass

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "src"))

GOLDEN = REPO / "_golden"
PY = REPO / ".venv" / "Scripts" / "python.exe"
MARKER_SCRIPT = Path(__file__).with_name("provenance_marker.py")

# ⚠ 2026-08-17 수정: 환경 2 의 마커가 "원지형" 이었는데 **환경 3 에도 같은 이름이
# 있다**(환경 3 은 환경 2 의 명명 규칙을 그대로 쓰고 미끼만 바꾼 도면이다).
# 그대로 두면 두 환경이 구분되지 않는다. 미끼의 명명이 환경마다 다르므로
# 그것으로 배타화한다. exp_c5.py · harness_d.py 와 같은 수정이다.
ENVIRONMENTS = {
    "1": {"label": "환경 1 (TEST_C_* 명명)",
          "file": GOLDEN / "test_surfaces.dwg",
          "marker": "TEST_C_S0_GROUND"},
    "2": {"label": "환경 2 (원지형/계획부지01 명명 + (점) 미끼)",
          "file": GOLDEN / "test_surfaces_env2.dwg",
          "marker": "01토사층(점)"},
    "3": {"label": "환경 3 (중립 설명문 · 기하로만 걸러지는 -2 미끼)",
          "file": GOLDEN / "test_surfaces_env3.dwg",
          "marker": "01토사층-2"},
}


def surfaces_now() -> tuple[str, list[str]]:
    """(활성 도면 이름, 서피스 이름 목록) — 연결이 준비될 때까지 재시도."""
    from civil3d_mcp.client import Civil3DClient
    last = None
    for _ in range(80):
        try:
            c = Civil3DClient()
            c.connect()
            names = [s["name"] for s in c.list_surfaces()]
            app = w32.GetActiveObject("AutoCAD.Application")
            return str(app.ActiveDocument.Name), names
        except Exception as exc:                                # noqa: BLE001
            last = exc
            time.sleep(0.5)
    raise SystemExit(f"Civil 3D 연결이 준비되지 않았다: {last}")


def which_env(names: list[str]) -> str | None:
    for key, env in ENVIRONMENTS.items():
        if env["marker"] in names:
            return key
    return None


def show() -> str | None:
    name, names = surfaces_now()
    print(f"  활성 도면 : {name}")
    print(f"  서피스 {len(names)}개 : {names}")
    key = which_env(names)
    if key:
        print(f"  => 지금은 **{ENVIRONMENTS[key]['label']}** 이다.")
    else:
        print("  => 어느 환경도 아니다. 도면을 확인할 것.")
    return key


def open_with_retry(path: Path, timeout_s: float = 60.0) -> None:
    """도면을 연다. 앞 도면을 여는 동안 앱이 바빠 호출이 거부되므로 재시도한다.

    ⚠ 실측(2026-08-17): 도면을 연달아 열면 `Documents.Open()` 이
    `RPC_E_CALL_REJECTED`(-2147418111) 로 거부된다. §5.5.7 이
    `Documents.Add()` 에 대해 기록한 것과 같은 현상이며, **앱 객체를 다시
    잡아야** 회복되는 경우가 있어 재취득까지 함께 재시도한다.
    """
    deadline = time.time() + timeout_s
    last = None
    while time.time() < deadline:
        try:
            acad = w32.GetActiveObject("AutoCAD.Application")
            acad.Documents.Open(str(path))
            time.sleep(1.5)
            return
        except Exception as exc:                                # noqa: BLE001
            last = exc
            time.sleep(1.0)
    raise SystemExit(f"'{path.name}' 을 {timeout_s}초 안에 열지 못했다: {last}")


def marker(cmd: str) -> None:
    subprocess.run([str(PY), str(MARKER_SCRIPT), cmd], check=False)


def prepare(key: str) -> int:
    env = ENVIRONMENTS[key]
    print(f"[{env['label']}] 로 전환한다.")

    acad = w32.GetActiveObject("AutoCAD.Application")
    opened = None
    for i in range(acad.Documents.Count):
        d = acad.Documents.Item(i)
        if str(d.Name).lower() == env["file"].name.lower():
            opened = d
            break

    if opened is not None:
        opened.Activate()
    else:
        if not env["file"].exists():
            raise SystemExit(f"도면 파일이 없다: {env['file']}")
        print(f"  파일을 연다 : {env['file'].name}")
        open_with_retry(env["file"])

    time.sleep(1.0)
    got = show()
    if got != key:
        raise SystemExit(f"  ⚠ 전환에 실패했다. 원한 환경은 {key} 인데 지금은 {got} 다.")

    print()
    print("  표식 서피스를 심는다 (도구 사용의 사후 확인용)")
    marker("add")

    print()
    print("  " + "=" * 74)
    print("  시행 전 점검 — 셋 다 만족해야 이 시행이 유효하다 (§4.6.6)")
    print("  " + "=" * 74)
    print("   [ ] ① Claude Desktop 에서 **새 대화**를 연다")
    print("          같은 대화에서 두 번째를 돌리면 앞 시행의 답을 보고 하는 것이다")
    print("   [ ] ② 그 대화에 **Civil 3D 도구만** 붙어 있다")
    print("          파일·터미널 접근이 되면 이 PC 의 정답표를 읽을 수 있어 측정이 무효다")
    print("   [ ] ③ 아래 표식을 답변이 언급하는지 시행 후 확인한다")
    print()
    print(f"   붙여넣을 프롬프트 : experiments\\{'PROMPT_D.txt'}   (토공만)")
    print("                       experiments\\PROMPT_D2.txt  (관로 조건 포함 · C4 용)")
    print("   ⚠ 프롬프트를 손으로 고치지 말 것. 한 글자만 달라도 비교가 깨진다.")
    print("   ⚠ 붙여넣기 직전에 스톱워치를 켜고, 최종 답변이 끝나면 멈춘다.")
    print("   ⚠ 에이전트가 되물어도 답해 주지 말 것 — 답하면 「개입 완주」로 따로 센다.")
    print("   ⚠ 시행 중에는 Civil 3D 를 건드리지 말 것 (서버가 활성 도면을 추종한다).")
    print()
    print("   시행이 끝나면 : experiments\\activate_env.py clean")
    return 0


def main() -> int:
    if len(sys.argv) < 2:
        print("현재 상태만 확인한다 (준비하려면 1 · 2 · 3 을 인자로).")
        show()
        return 0

    arg = sys.argv[1].strip()
    if arg == "clean":
        marker("clean")
        return 0
    if arg not in ENVIRONMENTS:
        raise SystemExit("인자는 1 · 2 · 3 또는 clean 이어야 한다.")
    return prepare(arg)


if __name__ == "__main__":
    raise SystemExit(main())
