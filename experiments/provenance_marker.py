"""provenance_marker.py — 도구 호출의 출처를 확정하는 표식 서피스

왜 필요한가
-----------
D 실험의 답변이 **도구를 불러서 나온 것인지** 확인해야 한다. 그런데
Claude Desktop 채팅의 도구 호출은 서버별 로그에 남지 않는 것으로 확인되어
(2026-08-17), 로그 대조라는 원래 방법이 이 환경에서 작동하지 않는다.

대안: **시행 직전에 무작위 이름·표고의 서피스를 도면에 하나 만든다.**
그 값은 소스에도 문서에도 대화에도 없으므로 **읽어서 알 방법이 없다.**
에이전트가 그것을 언급하면 살아 있는 도면을 조회한 것이 확정된다.

⚠ 표식은 **산정에 영향을 주지 않아야** 한다. 그래서
  · 이름을 `ZZ_PROBE_...` 로 두어 지층·계획고로 오인될 여지를 줄이고
  · 설명에 "verification marker, not part of the site" 를 적고
  · 표고를 원지반보다 **훨씬 높게**(EL 500) 두어 절토 계산에 끼어들 수 없게 한다

사용
----
    .\.venv\Scripts\python.exe experiments\provenance_marker.py add     # 시행 직전
    .\.venv\Scripts\python.exe experiments\provenance_marker.py check   # 답변 대조용
    .\.venv\Scripts\python.exe experiments\provenance_marker.py clean   # 시행 후 제거
"""
from __future__ import annotations

import json
import random
import string
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client as w32
from win32com.client import VARIANT

REPO = Path(__file__).resolve().parents[1]
STATE = Path(__file__).with_name("provenance_marker_state.json")
PREFIX = "ZZ_PROBE_"
ELEV = 500.0                    # 원지반(EL100)보다 훨씬 위 — 절토 계산에 안 끼어든다
X0, X1, Y0, Y1 = 200.0, 260.0, 200.0, 260.0   # 부지(0~100)와 떨어진 위치
GRID = 3


def _civil():
    acad = w32.GetActiveObject("AutoCAD.Application")
    for ver in ("13.6", "13.7", "13.8"):
        try:
            civil = acad.GetInterfaceObject(f"AeccXUiLand.AeccApplication.{ver}")
            return acad, civil.ActiveDocument, ver
        except Exception:                                       # noqa: BLE001
            continue
    raise SystemExit("Civil 3D 인터페이스를 얻지 못했다.")


def _names(cdoc) -> list[str]:
    return [cdoc.Surfaces.Item(i).Name for i in range(cdoc.Surfaces.Count)]


def add() -> None:
    acad, cdoc, prog = _civil()
    tag = "".join(random.choices(string.ascii_uppercase + string.digits, k=6))
    elev = round(ELEV + random.uniform(0, 99), 2)      # 표고도 무작위
    name = f"{PREFIX}{tag}"

    data = acad.GetInterfaceObject(f"AeccXLand.AeccTinCreationData.{prog}")
    data.Name = name
    data.Description = ("verification marker for provenance check - "
                        "NOT part of the site, ignore in earthwork computation")
    data.Layer = "0"
    data.BaseLayer = "0"
    data.Style = cdoc.SurfaceStyles.Item(0).Name
    surf = cdoc.Surfaces.AddTinSurface(data)

    pts: list[float] = []
    for i in range(GRID):
        x = X0 + (X1 - X0) * i / (GRID - 1)
        for j in range(GRID):
            y = Y0 + (Y1 - Y0) * j / (GRID - 1)
            pts.extend([x, y, elev])
    surf.AddPointMultiple(VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, pts))
    try:
        surf.Rebuild()
    except Exception:                                           # noqa: BLE001
        pass

    STATE.write_text(json.dumps(
        {"name": name, "tag": tag, "elevation": elev,
         "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
         "drawing": str(cdoc.Name)},
        ensure_ascii=False, indent=2), encoding="utf-8")

    print("표식 서피스를 만들었다.")
    print(f"  이름  : {name}")
    print(f"  표고  : EL {elev}")
    print(f"  도면  : {cdoc.Name}  (서피스 {cdoc.Surfaces.Count}개)")
    print()
    print("  => 이제 시행하라. 답변에 위 이름이나 표고가 나오면")
    print("     **살아 있는 도면을 조회한 것이 확정**된다.")
    print("     (이 값은 소스·문서·대화 어디에도 없다)")


def check() -> None:
    if not STATE.exists():
        raise SystemExit("표식 기록이 없다. 먼저 add 를 실행할 것.")
    st = json.loads(STATE.read_text(encoding="utf-8"))
    print("이번 시행의 표식:")
    print(f"  이름  : {st['name']}")
    print(f"  표고  : EL {st['elevation']}")
    print(f"  생성  : {st['created_at']}  ({st['drawing']})")
    print()
    print("  판정 - 답변에서 다음을 찾을 것:")
    print(f"    · '{st['name']}' 또는 '{st['tag']}'")
    print(f"    · 표고 {st['elevation']} 또는 EL 5xx 대 언급")
    print("    · '무시했다/제외했다' 는 취지의 언급도 조회의 증거다")


def clean() -> None:
    """열려 있는 **모든** 도면에서 표식을 제거한다.

    ⚠ 2026-08-17 수정: 이전에는 **활성 도면만** 훑었다. 환경을 바꿔 가며
    시행하면 앞 환경의 표식이 그 도면에 남고, `clean` 은 상태 파일이 아니라
    활성 도면을 보므로 그것을 지우지 못한다. 실제로 환경 3 에 `ZZ_PROBE_OKL413`
    이 고아로 남아 있었다.

    남은 표식은 조용히 다음 시행을 오염시킨다 — 에이전트가 **옛 표식**을
    언급하면 출처 확인이 성립하지 않는데, 표만 보아서는 알 수 없다.
    """
    import win32com.client as w32

    acad = w32.GetActiveObject("AutoCAD.Application")
    original = str(acad.ActiveDocument.Name)
    total = []

    for i in range(acad.Documents.Count):
        doc = acad.Documents.Item(i)
        doc.Activate()
        # ⚠ `_civil()` 은 실패를 SystemExit 로 올린다. SystemExit 는 Exception 이
        #    아니라 BaseException 이므로 `except Exception` 으로는 잡히지 않는다.
        #    도면 전환 직후에는 인터페이스 획득이 잠시 실패하므로(§5.5.7) 재시도가
        #    필요하고, 따라서 SystemExit 까지 함께 잡아야 한다.
        for _ in range(60):
            try:
                _, cdoc, _ = _civil()
                break
            except (Exception, SystemExit):                     # noqa: BLE001
                time.sleep(0.4)
        else:
            print(f"  ⚠ {doc.Name} 에 접근하지 못했다. 직접 확인할 것.")
            continue

        removed = []
        for j in range(cdoc.Surfaces.Count - 1, -1, -1):
            s = cdoc.Surfaces.Item(j)
            if s.Name.startswith(PREFIX):
                removed.append(s.Name)
                s.Erase()
        total += removed
        print(f"  {doc.Name:26s} 제거 {removed or '없음'}  남은 서피스 {cdoc.Surfaces.Count}개")

    # 원래 활성 도면으로 되돌린다
    for i in range(acad.Documents.Count):
        if str(acad.Documents.Item(i).Name) == original:
            acad.Documents.Item(i).Activate()
            break

    print("제거 합계:", total or "없음")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "check"
    {"add": add, "check": check, "clean": clean}.get(cmd, check)()
