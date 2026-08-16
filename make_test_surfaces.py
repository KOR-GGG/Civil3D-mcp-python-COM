"""make_test_surfaces.py — 정답을 해석적으로 아는 검증용 TIN 서피스를 생성한다.

왜 필요한가
-----------
실제 프로젝트 도면으로 하는 기준값 대조는 도면을 구할 수 있을 때만 가능하고,
암층을 관통하지 않는 사례에서는 '암질별로 나누는 능력' 자체가 검증되지 않는다.
여기서 만드는 서피스는 전부 **평면**이므로 TIN 삼각분할 방식과 무관하게
체적이 엄밀하게 결정되고, 정답을 산수로 확인할 수 있다.

또한 모든 서피스가 **같은 외곽 범위**를 갖도록 만들어, 두 서피스의 겹침
영역(overlap)이 산정 영역이 되는 현 구현에서 경계 변수를 제거한다.

시험 케이스
-----------
A. 평면 대 평면 — 부호 규약과 절토 산정
     BASE EL 100 / COMP EL 90  ->  절토 100,000.00 m3 / 성토 0.00 m3

B. 평면 대 경사면 — 절토·성토 동시 산정
     BASE EL 100 / COMP EL 90 + 0.2x (x=0 -> 90, x=100 -> 110)
     두 면은 x=50 에서 교차한다.
     절토 = int_0^50 (10 - 0.2x) dx * 100m = 250 * 100 = 25,000.00 m3
     성토 = 대칭이므로 동일        = 25,000.00 m3
     따라서 net = 0 이어야 한다.

C. 층서 관통 — 암질별 분할 능력 (실제 검증 도면으로는 확인 불가한 항목)
     원지반  S0 EL 100
     토사 하면 S1 EL  95   (토사   두께 5 m)
     풍화암 하면 S2 EL  92 (풍화암 두께 3 m)
     연암 하면 S3 EL  90   (연암   두께 2 m)
     계획고  P  EL  88

     C(S) = int max(S - P, 0) dA 라 두면 (면적 10,000 m2)
       C(S0) = 12 * 10,000 = 120,000
       C(S1) =  7 * 10,000 =  70,000
       C(S2) =  4 * 10,000 =  40,000
       C(S3) =  2 * 10,000 =  20,000
     층별 절토량
       V1 토사   = C(S0) - C(S1) =  50,000.00 m3
       V2 풍화암 = C(S1) - C(S2) =  30,000.00 m3
       V3 연암   = C(S2) - C(S3) =  20,000.00 m3
       최하층 하부(경암)          =  20,000.00 m3
     항등식 검산  50,000 + 30,000 + 20,000 + 20,000 = 120,000 = C(S0)  OK
     성토 = 0 (계획고가 원지반보다 항상 아래)

사용법
------
    .\.venv\Scripts\python.exe make_test_surfaces.py            # A B C 전부
    .\.venv\Scripts\python.exe make_test_surfaces.py C          # 일부만
    .\.venv\Scripts\python.exe make_test_surfaces.py --clean    # 시험 서피스 삭제

주의: 반드시 **비어 있는 새 도면**에서 실행할 것. 실무 도면에 서피스를 추가한다.
"""
from __future__ import annotations

import sys

import pythoncom
import win32com.client as w32
from win32com.client import VARIANT

# 산정 영역 — 모든 서피스가 동일하게 쓴다(경계 변수 제거).
X0, X1 = 0.0, 100.0
Y0, Y1 = 0.0, 100.0
AREA = (X1 - X0) * (Y1 - Y0)          # 10,000 m2
GRID = 5                              # 격자 분할 수 (평면이므로 정밀도와 무관)

PREFIX = "TEST_"

# 이름 -> (표고 함수, 설명)
CASES: dict[str, list[tuple[str, object, str]]] = {
    "A": [
        ("TEST_A_BASE", lambda x, y: 100.0, "flat EL100 (base)"),
        ("TEST_A_COMP", lambda x, y: 90.0, "flat EL90 (comparison)"),
    ],
    "B": [
        ("TEST_B_BASE", lambda x, y: 100.0, "flat EL100 (base)"),
        ("TEST_B_COMP", lambda x, y: 90.0 + 0.2 * x, "sloped EL90->EL110"),
    ],
    "C": [
        ("TEST_C_S0_GROUND",   lambda x, y: 100.0, "original ground EL100"),
        ("TEST_C_S1_SOIL",     lambda x, y: 95.0,  "bottom of soil layer EL95"),
        ("TEST_C_S2_WEATHERED", lambda x, y: 92.0, "bottom of weathered rock EL92"),
        ("TEST_C_S3_SOFT",     lambda x, y: 90.0,  "bottom of soft rock EL90"),
        ("TEST_C_DESIGN",      lambda x, y: 88.0,  "design grade EL88"),
    ],
}

EXPECTED = {
    "A": "절토 100,000.00 / 성토 0.00",
    "B": "절토  25,000.00 / 성토 25,000.00  (net = 0)",
    "C": "토사 50,000.00 / 풍화암 30,000.00 / 연암 20,000.00 / 경암 20,000.00"
         "  합계 120,000.00 / 성토 0.00",
}


def connect(prog: str = "13.6"):
    acad = w32.GetActiveObject("AutoCAD.Application")
    civil = None
    for ver in (prog, "13.8", "13.7", "13.6"):
        try:
            civil = acad.GetInterfaceObject(f"AeccXUiLand.AeccApplication.{ver}")
            prog = ver
            break
        except Exception:
            continue
    if civil is None:
        raise SystemExit("Civil 3D 인터페이스를 얻지 못했다. Civil 3D가 실행 중인지 확인할 것.")
    return acad, civil.ActiveDocument, prog


def grid_points(z_fn) -> list[float]:
    """GRID x GRID 격자를 (x, y, z) 평탄 배열로."""
    pts: list[float] = []
    for i in range(GRID):
        x = X0 + (X1 - X0) * i / (GRID - 1)
        for j in range(GRID):
            y = Y0 + (Y1 - Y0) * j / (GRID - 1)
            pts.extend([x, y, float(z_fn(x, y))])
    return pts


def erase_if_exists(cdoc, name: str) -> bool:
    for i in range(cdoc.Surfaces.Count - 1, -1, -1):
        s = cdoc.Surfaces.Item(i)
        if s.Name == name:
            s.Erase()
            return True
    return False


def make_surface(acad, cdoc, prog: str, style: str,
                 name: str, z_fn, desc: str):
    erase_if_exists(cdoc, name)
    data = acad.GetInterfaceObject(f"AeccXLand.AeccTinCreationData.{prog}")
    data.Name = name
    data.Description = desc
    data.Layer = "0"
    data.BaseLayer = "0"
    data.Style = style
    surf = cdoc.Surfaces.AddTinSurface(data)

    surf.AddPointMultiple(
        VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, grid_points(z_fn))
    )
    try:
        surf.Rebuild()
    except Exception as exc:                                   # noqa: BLE001
        print(f"    Rebuild 경고: {exc}")

    st = surf.Statistics
    print(f"    {name:22s} 점 {st.NumberOfPoints:3d} / 삼각형 {st.NumberOfTriangles:3d}"
          f" / EL {st.MinElevation:7.2f}~{st.MaxElevation:7.2f}"
          f" / 2D면적 {st.Area2d:10,.2f} m2")
    if abs(st.Area2d - AREA) > 1e-6:
        print(f"    ⚠ 2D 면적이 기대값 {AREA:,.2f} 와 다르다 — 경계 변수가 생긴다.")
    return surf


def main() -> int:
    args = [a for a in sys.argv[1:]]
    clean = "--clean" in args
    wanted = [a.upper() for a in args if a.upper() in CASES] or list(CASES)

    acad, cdoc, prog = connect()
    print(f"문서: {cdoc.Name}   ProgID 접미사: {prog}   기존 서피스: {cdoc.Surfaces.Count}개")

    if clean:
        removed = 0
        for i in range(cdoc.Surfaces.Count - 1, -1, -1):
            s = cdoc.Surfaces.Item(i)
            if s.Name.startswith(PREFIX):
                print("  삭제:", s.Name)
                s.Erase()
                removed += 1
        print(f"시험 서피스 {removed}개 삭제 완료.")
        return 0

    style = cdoc.SurfaceStyles.Item(0).Name
    print(f"사용할 서피스 스타일: {style}\n")

    for case in wanted:
        print(f"[시험 {case}]  기대값: {EXPECTED[case]}")
        for name, z_fn, desc in CASES[case]:
            make_surface(acad, cdoc, prog, style, name, z_fn, desc)
        print()

    print("현재 서피스 목록:")
    for i in range(cdoc.Surfaces.Count):
        print("  -", cdoc.Surfaces.Item(i).Name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
