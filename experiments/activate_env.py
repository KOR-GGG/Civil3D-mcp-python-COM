"""activate_env.py — D 실험용 도면 전환기 (안 ⓐ Claude Desktop 실행 보조)

왜 필요한가
-----------
MCP 서버는 **기동 시점의 활성 도면**에 붙는다. 그래서 환경을 바꾸려면
① 대상 도면을 활성화하고 ② Claude Desktop 을 재시작해야 한다.
도면을 손으로 찾아 클릭하면 엉뚱한 도면이 활성화된 채 시행이 진행될 수
있으므로, 이 스크립트로 확실히 바꾸고 **눈으로 확인**한다.

사용
----
    .\.venv\Scripts\python.exe experiments\activate_env.py 1   # 환경 1
    .\.venv\Scripts\python.exe experiments\activate_env.py 2   # 환경 2
    .\.venv\Scripts\python.exe experiments\activate_env.py     # 현재 상태만 확인
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import win32com.client as w32

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "src"))

GOLDEN = REPO / "_golden"

ENVIRONMENTS = {
    "1": {"label": "환경 1 (TEST_C_* 명명)",
          "file": GOLDEN / "test_surfaces.dwg",
          "marker": "TEST_C_S0_GROUND"},
    "2": {"label": "환경 2 (원지형/계획부지01 + 미끼)",
          "file": GOLDEN / "test_surfaces_env2.dwg",
          "marker": "원지형"},
}


def surfaces_now() -> tuple[str, list[str]]:
    """(활성 도면 이름, 서피스 이름 목록) — 연결이 준비될 때까지 재시도."""
    from civil3d_mcp.client import Civil3DClient
    last = None
    for _ in range(60):
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


def show() -> None:
    name, names = surfaces_now()
    print(f"  활성 도면: {name}")
    print(f"  서피스 {len(names)}개: {names}")
    for key, env in ENVIRONMENTS.items():
        if env["marker"] in names:
            print(f"  => 지금은 **{env['label']}** 이다.")
            return
    print("  => 어느 환경도 아니다. 도면을 확인할 것.")


def main() -> int:
    acad = w32.GetActiveObject("AutoCAD.Application")

    if len(sys.argv) < 2:
        print("현재 상태만 확인한다 (전환하려면 1 또는 2 를 인자로).")
        show()
        return 0

    key = sys.argv[1].strip()
    if key not in ENVIRONMENTS:
        raise SystemExit("인자는 1 또는 2 여야 한다.")
    env = ENVIRONMENTS[key]
    print(f"[{env['label']}] 로 전환한다.")

    # 1) 이미 열려 있으면 활성화, 아니면 파일을 연다
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
        print(f"  파일을 연다: {env['file'].name}")
        acad.Documents.Open(str(env["file"]))

    time.sleep(1.5)
    show()

    print()
    print("  ⚠ 다음 순서를 지킬 것:")
    print("     1) 위에서 환경이 맞는지 확인")
    print("     2) Claude Desktop 을 **완전히 종료**(작업 표시줄 아이콘 우클릭 → 종료)")
    print("     3) Claude Desktop 을 다시 시작")
    print("     4) **새 대화**를 열고 PROMPT_D.txt 내용을 붙여넣기")
    print("     서버는 기동 시점의 활성 도면에 붙으므로 재시작을 건너뛰면 안 된다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
