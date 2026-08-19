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
    .\.venv\Scripts\python.exe make_test_surfaces.py            # 전 케이스
    .\.venv\Scripts\python.exe make_test_surfaces.py C          # 일부만
    .\.venv\Scripts\python.exe make_test_surfaces.py E --new    # 새 도면에 환경 2
    .\.venv\Scripts\python.exe make_test_surfaces.py --clean    # 시험 서피스 삭제

    # 만든 도면을 파일로 저장한다(경로가 상대면 _golden/ 아래로 해석).
    .\.venv\Scripts\python.exe make_test_surfaces.py A B C P --save-as test_surfaces.dwg
    .\.venv\Scripts\python.exe make_test_surfaces.py E --new --save-as test_surfaces_env2.dwg
    .\.venv\Scripts\python.exe make_test_surfaces.py N --new --save-as test_surfaces_env3.dwg

⚠ 2026-08-18 추가: 예전에는 저장 기능이 없어 **Civil 3D 에서 손으로 「다른 이름으로
저장」을 해야 했다.** experiments/activate_env.py 가 _golden/ 아래의 위 세 파일명을
그대로 찾으므로, 이름이 하나라도 다르면 "도면 파일이 없다" 로 죽는다. --save-as 는
그 수작업을 없애고 파일명을 코드가 보증하게 한다.

주의: 반드시 **비어 있는 새 도면**에서 실행할 것. 실무 도면에 서피스를 추가한다.
"""
from __future__ import annotations

import glob
import os
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client as w32
from win32com.client import VARIANT

# 2026-08-19 추가: 한글 Windows 콘솔은 기본이 cp949 라, 아래 출력에 쓰이는
# em-dash 나 경고 기호 하나 때문에 UnicodeEncodeError 로 스크립트가 즉사한다.
# (setup_check.py 에서 같은 결함을 고쳤으나 실험 스크립트에는 남아 있었다.)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:                                          # noqa: BLE001
        pass

# 산정 영역 — 모든 서피스가 동일하게 쓴다(경계 변수 제거).
X0, X1 = 0.0, 100.0
Y0, Y1 = 0.0, 100.0
AREA = (X1 - X0) * (Y1 - Y0)          # 10,000 m2
GRID = 5                              # 격자 분할 수 (평면이므로 정밀도와 무관)

PREFIX = "TEST_"

# 검증 도면을 두는 곳. experiments/activate_env.py 의 GOLDEN 과 같은 위치여야 한다.
GOLDEN = Path(__file__).resolve().parent / "_golden"


def _take_save_as(args: list[str]) -> tuple[str | None, list[str]]:
    """--save-as <경로> 를 인자 목록에서 꺼내고 나머지를 돌려준다."""
    for i, a in enumerate(args):
        if a == "--save-as":
            if i + 1 >= len(args):
                raise SystemExit("--save-as 뒤에 파일명이 필요하다.")
            return args[i + 1], args[:i] + args[i + 2:]
        if a.startswith("--save-as="):
            return a.split("=", 1)[1], args[:i] + args[i + 1:]
    return None, args


def _resolve_save_path(raw: str) -> Path:
    """맨 파일명·상대 경로는 _golden/ 아래로 해석하고 확장자를 .dwg 로 맞춘다."""
    p = Path(raw)
    if not p.is_absolute():
        p = GOLDEN / p
    if p.suffix.lower() != ".dwg":
        p = p.with_suffix(".dwg")
    return p


def _save_drawing(acad, target: Path) -> None:
    """활성 도면을 target 으로 저장한다(폴더가 없으면 만든다)."""
    target.parent.mkdir(parents=True, exist_ok=True)
    acad.ActiveDocument.SaveAs(str(target))
    print()
    print(f"저장 완료: {target}")

# 이름 -> (표고 함수, 설명) [, (x0, x1, y0, y1) 로 외곽 범위를 달리 지정]
CASES: dict[str, list[tuple]] = {
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
    # ------------------------------------------------------------------
    # E. 환경 2 — C5(스크립트 대 에이전트) 실험용 두 번째 도면
    #
    # 지층 구성과 정답은 C와 **완전히 동일**하되, 명명 규칙과 서피스 개수를
    # 실제 검증 도면(2023 연구성과)의 관행에 맞추어 다르게 두었다.
    #   · 이름이 TEST_C_* 가 아니라 원지형 / 01토사층(Tin) / 계획부지01 …
    #   · 시추점 기반 중간 산출물 (점) 서피스가 섞여 있다 — 산정에 쓰면 안 되는
    #     서피스가 도면에 함께 존재하는 실제 상황의 재현
    #
    # 서피스 이름을 하드코딩한 스크립트는 이 도면에서 즉시 실패해야 하고,
    # 정답이 C와 같으므로 완주했을 때 값의 정오를 바로 판정할 수 있다.
    # ------------------------------------------------------------------
    "E": [
        ("원지형",           lambda x, y: 100.0, "E: original ground EL100"),
        ("01토사층(Tin)",     lambda x, y: 95.0,  "E: bottom of soil EL95"),
        ("02풍화암층(Tin)",   lambda x, y: 92.0,  "E: bottom of weathered rock EL92"),
        ("03연암층(Tin)",     lambda x, y: 90.0,  "E: bottom of soft rock EL90"),
        ("계획부지01",        lambda x, y: 88.0,  "E: design grade EL88"),
        # 아래 셋은 산정에 쓰면 안 되는 중간 산출물(시추점 기반). 미끼.
        ("01토사층(점)",      lambda x, y: 95.0,  "E: DECOY - borehole point surface"),
        ("02풍화암층(점)",    lambda x, y: 92.0,  "E: DECOY - borehole point surface"),
        ("03연암층(점)",      lambda x, y: 90.0,  "E: DECOY - borehole point surface"),
    ],
    # ------------------------------------------------------------------
    # N. 환경 3 — 미끼가 **자기 정체를 밝히지 않는** 변별 시험
    #
    # 환경 2의 결함을 고친 것이다. 환경 2에서 에이전트는 미끼를 정확히
    # 배제했으나 그 근거가 **"설명에 DECOY 로 표기"** 였다 — 미끼가 스스로
    # 정체를 밝히고 있었으므로 자율 변별의 증거가 되지 못한다.
    #
    # 여기서는
    #   · **모든 서피스의 설명을 동일한 중립 문구**로 둔다. 정답 서피스만
    #     설명이 다르면 그 차이 자체가 신호가 되므로 전부 같게 한다
    #   · 이름에도 힌트를 두지 않는다(`-2` 접미사뿐. BAD/DECOY/INVERTED 없음)
    #   · 미끼는 **기하로만** 걸러진다:
    #       N1 `01토사층-2` EL 91  -> 풍화암(EL92)보다 아래 = 층서 순서 위반
    #       N2 `02풍화암층-2`      -> 외곽이 x 0~60 뿐 = 산정 영역 불일치
    #       N3 `03연암층-2` EL 87  -> 계획고(EL88)보다 아래 = 절토 구간 밖
    #
    # 정답은 환경 1·2와 동일하다(토사 50,000 / 풍화암 30,000 / 연암 20,000 /
    # 경암 20,000 = 120,000). 미끼를 쓰면 값이 달라지므로 결과로 판별된다.
    # ------------------------------------------------------------------
    "N": [
        ("원지형",        lambda x, y: 100.0, "TIN surface"),
        ("01토사층",      lambda x, y: 95.0,  "TIN surface"),
        ("02풍화암층",    lambda x, y: 92.0,  "TIN surface"),
        ("03연암층",      lambda x, y: 90.0,  "TIN surface"),
        ("계획부지01",    lambda x, y: 88.0,  "TIN surface"),
        # --- 미끼 3종. 설명이 위와 완전히 같다 ---
        ("01토사층-2",    lambda x, y: 91.0,  "TIN surface"),
        ("02풍화암층-2",  lambda x, y: 92.0,  "TIN surface",
         (0.0, 60.0, 0.0, 100.0)),
        ("03연암층-2",    lambda x, y: 87.0,  "TIN surface"),
    ],
    # ------------------------------------------------------------------
    # P. 교란 케이스 — 오식별 검출률 측정용 (제5장 §4.6.4)
    #
    # 도구가 잘못된 입력을 「반려」하는지 보기 위한 서피스다. 값을 내면
    # 안 되는 입력을 일부러 만든다.
    # ------------------------------------------------------------------
    "P": [
        # P1 층서 역전 — 토사 하면(92)이 풍화암 하면(95)보다 아래
        ("TEST_P_S0_GROUND",    lambda x, y: 100.0, "P: original ground EL100"),
        ("TEST_P_S1_SOIL_BAD",  lambda x, y: 92.0,
         "P1: bottom of soil EL92 - INVERTED (below weathered)"),
        ("TEST_P_S2_WEATHERED_BAD", lambda x, y: 95.0,
         "P1: bottom of weathered EL95 - INVERTED (above soil)"),
        ("TEST_P_DESIGN",       lambda x, y: 88.0, "P: design grade EL88"),
        # P5 산정 영역 불일치 — 외곽 범위를 x∈[0,85] 로 좁힌 지층면
        ("TEST_P_NARROW_SOIL",  lambda x, y: 95.0,
         "P5: bottom of soil EL95 but NARROWER extent (x 0~85)",
         (0.0, 85.0, 0.0, 100.0)),
    ],
}

EXPECTED = {
    "A": "절토 100,000.00 / 성토 0.00",
    "B": "절토  25,000.00 / 성토 25,000.00  (net = 0)",
    "C": "토사 50,000.00 / 풍화암 30,000.00 / 연암 20,000.00 / 경암 20,000.00"
         "  합계 120,000.00 / 성토 0.00",
    "P": "교란 케이스 — 값이 나오면 안 된다(반려 기대). 제5장 §4.6.4",
    "E": "환경 2 — C와 같은 지층 구성이나 **명명 규칙과 서피스 개수가 다르다**."
         " 정답은 C와 동일(토사 50,000 / 풍화암 30,000 / 연암 20,000 / 경암 20,000)",
    "N": "환경 3 — **미끼가 자기 정체를 밝히지 않는다**(설명이 전부 동일)."
         " 미끼는 기하로만 걸러진다. 정답은 C·E와 동일(계 120,000)",
}


def _civil_template() -> str | None:
    """Civil 3D 미터 템플릿(.dwt) 경로를 찾는다. 없으면 None.

    ⚠ 일반 acad.dwt 로 만든 도면에는 **서피스 스타일이 없어** AddTinSurface 가
    실패한다. 반드시 Civil 3D 템플릿이어야 한다.
    """
    root = os.path.expandvars(r"%LOCALAPPDATA%\Autodesk")
    hits = glob.glob(os.path.join(root, "C3D*", "*", "Template",
                                  "_Autodesk Civil 3D (Metric)*.dwt"))
    return hits[0] if hits else None


RPC_E_CALL_REJECTED = -2147418111


def _wait_ready(timeout_s: float = 120.0):
    """Civil 3D 가 COM 호출을 받을 때까지 기다린다.

    반환: (재취득한 Application 객체, 활성 도면 이름)

    ⚠ 도면을 여는 동안 두 가지 증상이 모두 나타난다.
      · RPC_E_CALL_REJECTED (-2147418111) "피호출자가 호출을 거부했습니다"
      · AttributeError — 앱이 바쁜 사이 win32com 의 동적 디스패치가
        타입 정보를 얻지 못해 ActiveDocument 를 아예 못 찾는 상태
    후자 때문에 **기존 핸들을 재사용하면 안 되고 앱 객체를 다시 잡아야** 한다.
    """
    deadline = time.time() + timeout_s
    last = None
    while time.time() < deadline:
        try:
            app = w32.GetActiveObject("AutoCAD.Application")
            return app, str(app.ActiveDocument.Name)
        except Exception as exc:                                # noqa: BLE001
            last = exc
            time.sleep(0.5)
    raise SystemExit(f"Civil 3D 가 {timeout_s}초 안에 준비되지 않았다: {last}")


def connect(prog: str = "13.6", new_drawing: bool = False):
    acad = w32.GetActiveObject("AutoCAD.Application")

    if new_drawing:
        tpl = _civil_template()
        if not tpl:
            raise SystemExit(
                "Civil 3D 미터 템플릿(.dwt)을 찾지 못했다. 새 도면을 만들 수 없다."
            )
        # Documents.Add 로 만든 도면은 이미 활성 상태다(Activate 메서드 없음).
        acad.Documents.Add(tpl)
        print(f"새 도면 생성 — 템플릿: {os.path.basename(tpl)}")
        # 도면을 여는 동안 COM 호출이 거부되므로 준비될 때까지 기다린다.
        acad, name = _wait_ready()
        print(f"  활성 도면: {name}")

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


def grid_points(z_fn, extent: tuple[float, float, float, float] | None = None
                ) -> list[float]:
    """GRID x GRID 격자를 (x, y, z) 평탄 배열로.

    extent 를 주면 그 범위로 서피스를 만든다. 기본 범위와 다르게 두면
    두 서피스의 겹침 영역이 달라지므로 **산정 영역 불일치**를 재현할 수 있다.
    """
    x0, x1, y0, y1 = extent if extent else (X0, X1, Y0, Y1)
    pts: list[float] = []
    for i in range(GRID):
        x = x0 + (x1 - x0) * i / (GRID - 1)
        for j in range(GRID):
            y = y0 + (y1 - y0) * j / (GRID - 1)
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
                 name: str, z_fn, desc: str, extent=None):
    erase_if_exists(cdoc, name)
    data = acad.GetInterfaceObject(f"AeccXLand.AeccTinCreationData.{prog}")
    data.Name = name
    data.Description = desc
    data.Layer = "0"
    data.BaseLayer = "0"
    data.Style = style
    surf = cdoc.Surfaces.AddTinSurface(data)

    surf.AddPointMultiple(
        VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, grid_points(z_fn, extent))
    )
    try:
        surf.Rebuild()
    except Exception as exc:                                   # noqa: BLE001
        print(f"    Rebuild 경고: {exc}")

    st = surf.Statistics
    print(f"    {name:26s} 점 {st.NumberOfPoints:3d} / 삼각형 {st.NumberOfTriangles:3d}"
          f" / EL {st.MinElevation:7.2f}~{st.MaxElevation:7.2f}"
          f" / 2D면적 {st.Area2d:10,.2f} m2")
    if extent is None and abs(st.Area2d - AREA) > 1e-6:
        print(f"    ⚠ 2D 면적이 기대값 {AREA:,.2f} 와 다르다 — 경계 변수가 생긴다.")
    if extent is not None:
        print(f"    ※ 외곽 범위를 일부러 다르게 둔 서피스(교란용)")
    return surf


def main() -> int:
    args = [a for a in sys.argv[1:]]
    clean = "--clean" in args
    new_drawing = "--new" in args
    save_raw, args = _take_save_as(args)
    save_as = _resolve_save_path(save_raw) if save_raw else None
    wanted = [a.upper() for a in args if a.upper() in CASES] or list(CASES)

    acad, cdoc, prog = connect(new_drawing=new_drawing)
    print(f"문서: {cdoc.Name}   ProgID 접미사: {prog}   기존 서피스: {cdoc.Surfaces.Count}개")

    # 2026-08-18 안전장치. --save-as 는 *활성 도면*을 그 경로로 저장한다(SaveAs).
    # --new 없이 실무 도면이 열린 상태로 실행하면 그 도면에 시험 서피스가 들어가고
    # 저장 경로까지 _golden 쪽으로 바뀐다. 서피스가 이미 있으면 실무 도면으로 보고 멈춘다.
    if save_as is not None and not new_drawing and cdoc.Surfaces.Count > 0:
        print()
        print("중단 — --save-as 는 활성 도면을 그 경로로 저장한다.")
        print(f"  활성 도면 : {cdoc.Name}  (서피스 {cdoc.Surfaces.Count}개)")
        print("  실무 도면일 수 있다. 이대로 진행하면 그 도면에 시험 서피스가 들어가고")
        print(f"  저장 경로까지 {save_as} 로 바뀐다.")
        print("  --new 를 붙여 빈 새 도면에서 만들거나, 빈 도면을 연 뒤 다시 실행할 것.")
        return 2

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
        for item in CASES[case]:
            name, z_fn, desc = item[0], item[1], item[2]
            extent = item[3] if len(item) > 3 else None
            make_surface(acad, cdoc, prog, style, name, z_fn, desc, extent)
        print()

    print("현재 서피스 목록:")
    for i in range(cdoc.Surfaces.Count):
        print("  -", cdoc.Surfaces.Item(i).Name)

    if save_as is not None:
        _save_drawing(acad, save_as)
    else:
        print()
        print("⚠ 도면이 저장되지 않았다. experiments/activate_env.py 는 _golden/ 의")
        print("  test_surfaces.dwg · test_surfaces_env2.dwg · test_surfaces_env3.dwg 를")
        print("  찾는다. --save-as <파일명> 을 주거나 그 이름으로 직접 저장할 것.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
