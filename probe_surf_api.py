"""probe_surf_api.py — Civil 3D 서피스 API 조사 스크립트 (테스트 아님)

⚠⚠ **이 파일은 원래 `test_surf_api.py` 라는 이름이었고, 그것이 위험했다.**

이 스크립트는 pytest 테스트가 아니라 **모듈 최상위에서 곧바로 실행되는 조사
코드**다. 그런데 이름이 `test_` 로 시작해 **pytest 가 수집하여 import 시점에
실행**했다. 그 결과 `pytest` 를 한 번 돌리는 것만으로

  · 실행 중인 Civil 3D 에 접속하고
  · **열려 있는 도면에 TIN 서피스(`__FullWorkflowTest__`)를 생성**하며
  · 작업 디렉터리에 `surf_api_out.txt` 를 덮어썼다.

즉 **테스트 수트를 돌리면 사용자의 열린 도면이 변경된다.** 검증 도면 원본을
열어 둔 채 pytest 를 돌리면 원본이 오염되므로, 원본 보호 수칙과 정면으로
충돌한다. (2026-08-17 확인. 이 랩탑에서는 ProgID 13.7 이 등록되어 있지 않아
예외로 끝나 무해했으나, Civil 3D 2025 가 설치된 환경에서는 실제로 생성된다.)

→ pytest 가 수집하지 않도록 이름을 `probe_` 로 바꾸었다. 실행하려면 직접
   `python probe_surf_api.py` 로 호출할 것. **반드시 사본 도면에서.**
"""
import sys
import win32com.client as w32
import pythoncom
from civil3d_mcp.client import Civil3DClient
import array

# 2026-08-19 추가: 한글 Windows 콘솔은 기본이 cp949 라, 아래 출력에 쓰이는
# em-dash 나 경고 기호 하나 때문에 UnicodeEncodeError 로 스크립트가 즉사한다.
# (setup_check.py 에서 같은 결함을 고쳤으나 실험 스크립트에는 남아 있었다.)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:                                          # noqa: BLE001
        pass

out = []

try:
    c = Civil3DClient()
    c.connect()

    # Get TypeInfo for AddPointMultiple
    cd = c._new_aecc_object("AeccTinCreationData")   # 2026-08-19: "13.7" 하드코딩 제거
                                                  # (2026 에서 UI 는 13.7 이 되지만 LAND 는 13.8 만 된다)
    cd.Name = "__FullWorkflowTest__"
    cd.Layer = "0"
    cd.BaseLayer = "0"
    cd.Style = "Border Only"
    surf = c._doc.Surfaces.AddTinSurface(cd)

    # Get AddPointMultiple TypeInfo
    try:
        ti = surf._oleobj_.GetTypeInfo()
        ta = ti.GetTypeAttr()
        for fi in range(ta.cFuncs):
            fd = ti.GetFuncDesc(fi)
            fname = ti.GetNames(fd.memid)[0]
            if fname in ("AddPointMultiple", "AddPoint"):
                out.append(f"func: {fname} memid={fd.memid} args={fd.args}")
    except Exception as ex:
        out.append("TypeInfo err: " + str(ex)[:100])

    # Probe Breaklines collection
    blines = surf.Breaklines
    out.append("Breaklines dir: " + str([a for a in dir(blines) if not a.startswith("_")]))

    # Get TypeInfo for Breaklines.Add
    try:
        ti2 = blines._oleobj_.GetTypeInfo()
        ta2 = ti2.GetTypeAttr()
        for fi in range(ta2.cFuncs):
            fd2 = ti2.GetFuncDesc(fi)
            fname2 = ti2.GetNames(fd2.memid)[0]
            if fname2 == "Add":
                out.append(f"Breaklines.Add args: {fd2.args}")
                break
    except Exception as ex:
        out.append("Breaklines TypeInfo err: " + str(ex)[:100])

    # Find the closed polyline
    ms = c._acad.ActiveDocument.ModelSpace
    pline_obj = None
    for i in range(ms.Count):
        try:
            raw = ms.Item(i)
            obj = w32.Dispatch(raw)
            oname = getattr(obj, "ObjectName", "")
            if "Polyline" in oname and getattr(obj, "Closed", False) and pline_obj is None:
                pline_obj = obj
        except:
            pass

    if not pline_obj:
        coords = array.array('d', [0,0,0, 100,0,0, 100,100,0, 0,100,0])
        pline = c._acad.ActiveDocument.ModelSpace.AddPolyline(coords)
        pline.Closed = True
        pline_obj = w32.Dispatch(pline)
        out.append("Created polyline: " + pline_obj.Handle)

    out.append("polyline: " + str(pline_obj.Handle))

    # Try AddPointMultiple with different arg formats
    coords_raw = pline_obj.Coordinates
    n = len(coords_raw) // 2
    pts_flat = []
    for i in range(n):
        pts_flat.extend([coords_raw[i*2], coords_raw[i*2+1], 0.0])

    for arr_type in [
        array.array('d', pts_flat),
        tuple(pts_flat),
        pts_flat,
        tuple((coords_raw[i*2], coords_raw[i*2+1], 0.0) for i in range(n)),
    ]:
        try:
            surf.AddPointMultiple(arr_type)
            out.append("AddPointMultiple OK with " + type(arr_type).__name__)
            break
        except Exception as ex:
            out.append("AddPointMultiple " + type(arr_type).__name__ + " FAIL: " + str(ex.args[1] if len(ex.args)>1 else ex)[:60])

    # Try adding breakline
    try:
        blines = surf.Breaklines
        for args in [
            (pline_obj, "Outer", 0, False, 0.0),
            (pline_obj, "Outer", 0),
            (pline_obj, "Outer"),
            (pline_obj,),
        ]:
            try:
                result = blines.Add(*args)
                out.append(f"Breaklines.Add({len(args)} args) SUCCESS!")
                break
            except Exception as ex:
                out.append(f"Breaklines.Add({len(args)}) err: " + str(ex.args[:3] if ex.args else ex)[:120])
    except Exception as ex:
        out.append("breaklines err: " + str(ex)[:100])

    surf.Erase()
    out.append("Erased")
except Exception as ex:
    out.append("TOP ERR: " + str(ex)[:300])

with open("surf_api_out.txt", "w") as f:
    f.write("\n".join(out))
