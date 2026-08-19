"""exp_c5.py — 고정 스크립트의 환경 의존성 측정 (제5장 §3, 명제 C5)

무엇을 재는가
-------------
제5장 §3이 정직하게 좁혀 둔 명제는 다음이다.

  ❌ "에이전트가 항상 우월하다" — 방어 불가
  ✅ "서피스 구성과 명명 규칙이 다른 도면에서
      C(고정 스크립트)의 성공률은 급락하는 반면 D(에이전트)는 완주한다"

이 실험은 그 **전반부**를 측정한다. 즉 **C가 환경 1에서는 완주하고 환경 2에서는
실패한다**는 사실을 확인한다.

⚠ **후반부(D가 완주한다)는 여기서 측정하지 않는다.** 에이전트 실행에는 MCP
호스트가 필요하며, 본 스크립트로 그 대역을 만들어 D라고 부르면 그것은 측정이
아니라 연출이다. D 는 별도 환경에서 측정하고, 두 값을 합쳐 보고할 때 출처를
명시해야 한다.

환경
----
  환경 1 : test_surfaces.dwg   — TEST_C_* 명명, 서피스 14개
  환경 2 : (새 도면)            — 원지형 / 01토사층(Tin) / 계획부지01 명명,
                                 **미끼 (점) 서피스 3개 포함**, 서피스 8개
  두 환경의 지층 구성과 정답은 **동일**하다. 다른 것은 이름과 개수뿐이다.

선행 조건
---------
    make_test_surfaces.py C P        (환경 1)
    make_test_surfaces.py E --new    (환경 2)
두 도면이 모두 열려 있어야 한다.

실행
----
    .\.venv\Scripts\python.exe experiments\exp_c5.py
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import win32com.client as w32

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from baseline_c_script import run_guarded                      # noqa: E402
from civil3d_mcp.client import Civil3DClient                   # noqa: E402

# 2026-08-19 추가: 한글 Windows 콘솔은 기본이 cp949 라, 아래 출력에 쓰이는
# em-dash 나 경고 기호 하나 때문에 UnicodeEncodeError 로 스크립트가 즉사한다.
# (setup_check.py 에서 같은 결함을 고쳤으나 실험 스크립트에는 남아 있었다.)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:                                          # noqa: BLE001
        pass

# 환경 식별용 — 그 도면에만 있는 서피스 이름
#
# ⚠ 2026-08-17 수정: 환경 2 의 마커가 "원지형" 이었는데 환경 3 에도 같은 이름이
# 있다(환경 3 은 환경 2 의 명명 규칙을 그대로 쓰고 미끼만 바꾼 도면이다).
# 그대로 두면 먼저 활성화된 쪽이 "환경 2" 를 차지해 두 환경을 구분하지 못한다.
# 서로 배타적인 이름으로 바꾼다 — 미끼의 명명이 환경마다 다르다.
#   환경 2 의 미끼: 01토사층(점) …      환경 3 의 미끼: 01토사층-2 …
ENV_MARKERS = {
    "환경 1 (TEST_C_* 명명)": "TEST_C_S0_GROUND",
    "환경 2 (원지형/계획부지01 명명 + 미끼)": "01토사층(점)",
    "환경 3 (중립 설명문 · 기하로만 걸러지는 미끼)": "01토사층-2",
}

# 환경별 도면 파일 — 열려 있지 않으면 여기서 연다
GOLDEN = Path(__file__).resolve().parents[1] / "_golden"
ENV_FILES = {
    "환경 1 (TEST_C_* 명명)": GOLDEN / "test_surfaces.dwg",
    "환경 2 (원지형/계획부지01 명명 + 미끼)": GOLDEN / "test_surfaces_env2.dwg",
    "환경 3 (중립 설명문 · 기하로만 걸러지는 미끼)": GOLDEN / "test_surfaces_env3.dwg",
}


def list_documents(acad) -> list:
    return [acad.Documents.Item(i) for i in range(acad.Documents.Count)]


def open_missing(acad, found: dict) -> None:
    """식별되지 않은 환경의 도면을 연다.

    exp_c5 는 '열려 있는 도면' 중에서 환경을 찾으므로, 도면이 닫혀 있으면
    측정이 통째로 빠진다. 빠진 것을 실패로 적으면 측정이 아니라 착오다.
    """
    for label, path in ENV_FILES.items():
        if label in found:
            continue
        if not path.exists():
            print(f"  ⚠ 도면 파일이 없다: {path}")
            continue
        print(f"  {label} 의 도면이 열려 있지 않다 -> 연다: {path.name}")
        open_with_retry(acad, path)


def open_with_retry(acad, path: Path, timeout_s: float = 60.0):
    """도면을 연다. 앞 도면을 여는 동안 앱이 바빠 호출이 거부되므로 재시도한다.

    ⚠ 실측(2026-08-17): 도면을 연달아 열면 두 번째·세 번째 `Documents.Open()`
    에서 `RPC_E_CALL_REJECTED`(-2147418111, "피호출자가 호출을 거부했습니다")
    가 난다. §5.5.7 이 `Documents.Add()` 에 대해 기록한 것과 같은 현상이며,
    **앱 객체를 다시 잡아야** 회복되는 경우가 있어 재취득까지 함께 재시도한다.
    """
    deadline = time.time() + timeout_s
    last = None
    while time.time() < deadline:
        try:
            acad = w32.GetActiveObject("AutoCAD.Application")   # 핸들 재취득
            acad.Documents.Open(str(path))
            time.sleep(1.5)
            return
        except Exception as exc:                                # noqa: BLE001
            last = exc
            time.sleep(1.0)
    raise SystemExit(f"'{path.name}' 을 {timeout_s}초 안에 열지 못했다: {last}")


def activate(acad, doc) -> str:
    doc.Activate()
    for _ in range(120):
        try:
            app = w32.GetActiveObject("AutoCAD.Application")
            if str(app.ActiveDocument.Name) == str(doc.Name):
                return str(doc.Name)
        except Exception:                                       # noqa: BLE001
            pass
        time.sleep(0.25)
    raise SystemExit(f"'{doc.Name}' 활성화에 실패했다.")


def client_when_ready(timeout_s: float = 60.0) -> Civil3DClient:
    """도면 전환 직후에는 Civil 3D 인터페이스가 잠시 무효하다.

    ⚠ 실측(2026-08-17): `Documents.Item(i).Activate()` 직후 `GetInterfaceObject`
    로 얻은 AeccApplication 이 `ActiveDocument` 를 내주지 못하고
    `AttributeError: GetInterfaceObject.ActiveDocument` 가 난다. 활성 문서 이름이
    이미 바뀐 뒤에도 그렇다. 따라서 **연결과 첫 조회가 함께 성공할 때까지**
    재시도해야 한다.
    """
    deadline = time.time() + timeout_s
    last = None
    while time.time() < deadline:
        try:
            c = Civil3DClient()
            c.connect()
            c.list_surfaces()          # 실제 조회까지 성공해야 준비된 것
            return c
        except Exception as exc:                                # noqa: BLE001
            last = exc
            time.sleep(0.5)
    raise SystemExit(f"Civil 3D 연결이 {timeout_s}초 안에 준비되지 않았다: {last}")


def surfaces_of(client: Civil3DClient) -> list[str]:
    return [s["name"] for s in client.list_surfaces()]


def main() -> int:
    acad = w32.GetActiveObject("AutoCAD.Application")
    docs = list_documents(acad)
    print("=" * 96)
    print("  C5 전반부 — 고정 스크립트의 환경 의존성 (제5장 §3)")
    print("=" * 96)
    print(f"  열린 도면 {len(docs)}개: {[str(d.Name) for d in docs]}")
    print()

    # 각 도면을 활성화해 어떤 환경인지 식별한다
    found: dict[str, object] = {}
    for d in docs:
        activate(acad, d)
        c = client_when_ready()
        names = surfaces_of(c)
        for label, marker in ENV_MARKERS.items():
            if marker in names and label not in found:
                found[label] = d
                print(f"  {label:38s} <- {d.Name}  (서피스 {len(names)}개)")

    # 열려 있지 않은 환경이 있으면 도면을 열고 한 번 더 식별한다
    if [k for k in ENV_MARKERS if k not in found]:
        print()
        open_missing(acad, found)
        print()
        for d in list_documents(acad):
            if d in found.values():
                continue
            activate(acad, d)
            c = client_when_ready()
            names = surfaces_of(c)
            for label, marker in ENV_MARKERS.items():
                if marker in names and label not in found:
                    found[label] = d
                    print(f"  {label:44s} <- {d.Name}  (서피스 {len(names)}개)")

    missing = [k for k in ENV_MARKERS if k not in found]
    if missing:
        print(f"\n  ⚠ 환경을 찾지 못했다: {missing}")
        print("     make_test_surfaces.py C P  /  make_test_surfaces.py E --new 를 먼저 실행할 것.")
        return 1

    print()
    print("-" * 96)
    print("  비교군 C(고정 스크립트) 실행")
    print("-" * 96)

    rows = []
    for label, doc in found.items():
        activate(acad, doc)
        client = client_when_ready()
        r = run_guarded(client)
        rows.append({"env": label, "drawing": str(doc.Name), **r})

        if r["completed"]:
            print(f"  {label}")
            print(f"     완주 O  ({r['elapsed_s']:.2f} s)  "
                  f"총 절토 {r['total_cut_m3']:,.2f} ㎥ · "
                  f"관경 D{r['selected_diameter_mm']} · "
                  f"가압 {r['pressurization_points']}개소")
            for k, v in r["by_stratum"].items():
                print(f"       {k:20s} {v:12,.2f} ㎥")
        else:
            print(f"  {label}")
            print(f"     완주 X  ({r['elapsed_s']:.2f} s)  실패 지점: {r['failed_at']}")
            print(f"       {r['error']}")
        print()

    print("=" * 96)
    print("  결과")
    print("=" * 96)
    ok = [r for r in rows if r["completed"]]
    ng = [r for r in rows if not r["completed"]]
    print(f"  완주 {len(ok)}/{len(rows)} — 성공 {[r['env'] for r in ok]}")
    print(f"                실패 {[r['env'] for r in ng]}")
    print()
    if len(ok) == 1 and len(ng) >= 1:
        print("  → 명제의 전반부가 확인되었다: 지층 구성과 정답이 동일해도")
        print("    **이름과 개수가 달라지면 고정 스크립트는 완주하지 못한다.**")
        print("    실패는 성능 저하가 아니라 첫 도구 호출에서의 즉시 중단이다.")
    print()
    print("  ⚠ 후반부(D 에이전트가 완주하는가)는 미측정. MCP 호스트 필요.")
    print("    두 값을 합쳐 보고할 때 출처를 반드시 구분할 것.")

    out = Path(__file__).with_name("c5_result.json")
    out.write_text(json.dumps({"arm": "C only", "rows": rows},
                              ensure_ascii=False, indent=2, default=str),
                   encoding="utf-8")
    print(f"\n  결과 저장 -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
