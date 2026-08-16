"""
client.py  –  Civil3DClient
============================
Wraps COM automation to the running Civil 3D instance.

Connection strategy
-------------------
1.  win32com.client.GetActiveObject(prog_id)
    Grabs the live acad.exe COM server.  We try Civil 3D's own ProgID first
    ("AeccXUiLand.AeccApplication.14.0") so the full Civil 3D object model
    (CogoPoints, Surfaces, Alignments …) is available.  Fall back to the
    plain AutoCAD ProgID when only basic drawing tools are needed.

2.  pythonnet (clr) – optional managed-assembly bridge
    Loads AeccDbMgd.dll / AeccLandMgd.dll so Python can call Civil 3D .NET
    APIs that are NOT exposed through the raw IDispatch COM interface.

Civil 3D must be running and have a drawing open.
Every public method returns plain Python dicts / lists so the MCP tool layer
never touches COM objects directly.
"""

from __future__ import annotations

import logging
import os
from typing import Any

log = logging.getLogger("civil3d_mcp.client")

# ---------------------------------------------------------------------------
# Windows-only import guards
# ---------------------------------------------------------------------------
try:
    import pythoncom
    import win32com.client as w32
    _WIN32 = True
except ImportError:
    _WIN32 = False
    log.warning("pywin32 not found – COM automation unavailable")

try:
    import clr  # pythonnet
    _CLR = True
except ImportError:
    _CLR = False
    log.warning("pythonnet not found – managed assembly access unavailable")


# ---------------------------------------------------------------------------
# Known Autodesk assembly locations (Civil 3D 2023 – 2026 defaults)
# ---------------------------------------------------------------------------
# Required DLL — always present in Civil 3D installations
_CIVIL3D_DLL_NAMES = ["AeccDbMgd.dll"]
# Optional DLL — present in Civil 3D 2023/2024 but removed in 2025+
_CIVIL3D_DLL_NAMES_OPTIONAL = ["AeccLandMgd.dll"]
_ACAD_DLL_NAME = "acdbmgd.dll"
_AUTODESK_ROOTS = [
    r"C:\Program Files\Autodesk\AutoCAD 2026\C3D",
    r"C:\Program Files\Autodesk\AutoCAD 2026",
    r"C:\Program Files\Autodesk\AutoCAD 2025\C3D",
    r"C:\Program Files\Autodesk\AutoCAD 2025",
    r"C:\Program Files\Autodesk\AutoCAD 2024\C3D",
    r"C:\Program Files\Autodesk\AutoCAD 2024",
    r"C:\Program Files\Autodesk\AutoCAD 2023",
]
# Override the entire bin folder via env var
_CIVIL3D_BIN = os.getenv("CIVIL3D_BIN_PATH", "").strip()


def _find_dll(name: str) -> str | None:
    """Return the first existing path for a named Autodesk DLL."""
    roots: list[str] = []
    if _CIVIL3D_BIN:
        roots.append(_CIVIL3D_BIN)
        # Also search the parent folder (e.g. AutoCAD 2025\ when BIN_PATH is AutoCAD 2025\C3D)
        parent = os.path.dirname(_CIVIL3D_BIN)
        if parent and parent != _CIVIL3D_BIN:
            roots.append(parent)
    roots.extend(_AUTODESK_ROOTS)
    for root in roots:
        path = os.path.join(root, name)
        if os.path.exists(path):
            return path
    return None


# ---------------------------------------------------------------------------
# Civil3DClient
# ---------------------------------------------------------------------------

class Civil3DError(RuntimeError):
    """Raised for any Civil 3D COM / API error."""


class Civil3DClient:
    """
    Thin wrapper around the Civil 3D COM object model.

    Usage
    -----
        client = Civil3DClient()
        client.connect()               # once at server startup
        info = client.get_drawing_info()
        client.disconnect()            # on shutdown
    """

    def __init__(self) -> None:
        self._acad: Any = None       # AcadApplication COM object (AutoCAD base)
        self._civil: Any = None      # AeccApplication COM object (Civil 3D layer)
        self._doc: Any = None        # AeccDocument (Civil 3D) or AcadDocument fallback
        self._connected = False
        # ProgID that actually succeeded in connect(); its version suffix is
        # reused to instantiate other AeccXLand coclasses (creation-data objects).
        self._prog_id: str | None = None

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    # Civil 3D ProgIDs in order of preference.
    #
    # 2026-08-06 수정: 원본 저장소는 Civil 3D 2026을 "14.4"로 표기했으나,
    # 실제 등록된 ProgID는 "13.8"이다(레지스트리 HKCR 조회로 확인).
    # 이 오류 때문에 Civil 3D 2026에서 GetInterfaceObject가 전부 실패하고
    # 일반 AutoCAD로 폴백되어 Surfaces 컬렉션에 접근할 수 없었다.
    #
    # 2026-08-16 추가: 두 번째 PC(Civil 3D 2024)에서 "13.6"을 확인했다.
    # 즉 등록 ProgID는 릴리스 연도가 아니라 13.6(2024)·13.7(2025)·13.8(2026)의
    # 연속된 내부 버전을 따르며, 원본이 표기한 14.x 계열은 확인한 세 버전 중
    # 어디에도 대응하지 않는다. 원본의 오표기는 2026 한정 사고가 아니라
    # 버전 표(表) 전체가 실제 등록값과 무관하게 작성된 결과로 보인다.
    _CIVIL3D_PROG_IDS = [
        "AeccXUiLand.AeccApplication.13.8",  # Civil 3D 2026 — HKCR로 검증
        "AeccXUiLand.AeccApplication.13.7",  # Civil 3D 2025 — HKCR로 검증
        "AeccXUiLand.AeccApplication.13.6",  # Civil 3D 2024 — HKCR로 검증
        "AeccXUiLand.AeccApplication.14.4",  # 원본 값(미검증) — 폴백용 유지
        "AeccXUiLand.AeccApplication.14.0",  # 원본 값(미검증) — 폴백용 유지
        "AeccXUiLand.AeccApplication.13.0",  # 원본 값(미검증) — 폴백용 유지
    ]

    def connect(self) -> None:
        """Attach to the running Civil 3D / AutoCAD instance via COM.

        Strategy
        --------
        1. Connect to AutoCAD (always registers as 'AutoCAD.Application').
        2. Promote to Civil 3D by calling acad.GetInterfaceObject() with known
           Civil 3D ProgIDs.  This works even when Civil 3D hasn't registered
           its AeccApplication in the Running Object Table (Civil 3D 2025).
        3. Fall back to the AutoCAD document only — Civil 3D collections won't
           be available but basic drawing info will work.
        """
        if not _WIN32:
            raise Civil3DError(
                "pywin32 is required for COM automation. "
                "Install with: pip install pywin32"
            )

        log.info("Connecting to Civil 3D via COM…")

        # Step 1 — connect to AutoCAD base application.
        try:
            self._acad = w32.GetActiveObject("AutoCAD.Application")
            log.info("AutoCAD.Application connected – drawing: %s",
                     self._acad.ActiveDocument.Name)
        except Exception as exc:
            raise Civil3DError(
                "Could not connect to AutoCAD/Civil 3D. "
                "Make sure Civil 3D is open with a drawing loaded."
            ) from exc

        # Step 2 — promote to Civil 3D via GetInterfaceObject.
        # This asks the already-running AutoCAD/C3D process to return its
        # Civil 3D application interface — works even when the Aecc ProgID isn't
        # in the ROT (Civil 3D 2025 behaviour).
        self._civil = None
        for prog_id in self._CIVIL3D_PROG_IDS:
            try:
                civil = self._acad.GetInterfaceObject(prog_id)
                if civil is not None:
                    self._civil = civil
                    self._doc = civil.ActiveDocument
                    self._prog_id = prog_id
                    log.info("Civil 3D interface acquired via GetInterfaceObject(%s)", prog_id)
                    break
            except Exception:
                continue

        if self._civil is None:
            # Fallback: plain AutoCAD document — Civil 3D collections unavailable.
            self._doc = self._acad.ActiveDocument
            log.warning(
                "Could not acquire Civil 3D interface. "
                "Connected as plain AutoCAD — Civil 3D collections (surfaces, "
                "alignments, COGO points) will not be available. "
                "Ensure Civil 3D is launched (not just AutoCAD)."
            )

        if self._doc is None:
            raise Civil3DError("No active drawing found.")

        self._load_managed_assemblies()
        self._connected = True
        log.info("Civil 3D client ready – drawing: %s", self._doc.Name)

    def _load_managed_assemblies(self) -> None:
        """Load Autodesk .NET DLLs via pythonnet for Civil 3D managed types."""
        if not _CLR:
            log.warning(
                "pythonnet not available – Civil 3D managed types "
                "(surfaces, alignments) will use COM IDispatch only."
            )
            return

        for dll in [_ACAD_DLL_NAME] + _CIVIL3D_DLL_NAMES:
            path = _find_dll(dll)
            if path:
                try:
                    clr.AddReference(path)
                    log.debug("Loaded assembly: %s", path)
                except Exception as exc:
                    log.warning("Could not load %s: %s", dll, exc)
            else:
                log.warning(
                    "Assembly not found: %s. "
                    "Set CIVIL3D_BIN_PATH to the folder containing this DLL.",
                    dll,
                )
        # Optional DLLs — load if present, skip silently if not (removed in Civil 3D 2025+)
        for dll in _CIVIL3D_DLL_NAMES_OPTIONAL:
            path = _find_dll(dll)
            if path:
                try:
                    clr.AddReference(path)
                    log.debug("Loaded optional assembly: %s", path)
                except Exception as exc:
                    log.warning("Could not load optional %s: %s", dll, exc)
            else:
                log.debug("Optional assembly not present (skipping): %s", dll)

    def disconnect(self) -> None:
        self._acad = None
        self._civil = None
        self._doc = None
        self._connected = False
        log.info("Civil 3D client disconnected")

    def _ensure_connected(self) -> None:
        if not self._connected or self._acad is None:
            raise Civil3DError(
                "Not connected to Civil 3D. "
                "Restart the MCP server with Civil 3D open."
            )
        try:
            # Refresh active document from Civil 3D layer if available, else AutoCAD
            source = self._civil if self._civil is not None else self._acad
            self._doc = source.ActiveDocument
        except Exception as exc:
            self._connected = False
            raise Civil3DError(f"Lost connection to Civil 3D: {exc}") from exc

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _pt3d(x: float, y: float, z: float) -> Any:
        """Build a VT_ARRAY|VT_R8 VARIANT for a 3-D point."""
        return w32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y, z])

    @staticmethod
    def _out_double() -> Any:
        """Build a by-reference VT_R8 VARIANT for COM out-parameters."""
        return w32.VARIANT(pythoncom.VT_R8 | pythoncom.VT_BYREF, 0.0)

    # ------------------------------------------------------------------
    # Drawing info
    # ------------------------------------------------------------------

    def get_drawing_info(self) -> dict[str, Any]:
        """Return metadata about the active drawing."""
        self._ensure_connected()
        doc = self._doc
        try:
            return {
                "name": str(doc.Name),
                "full_path": str(doc.FullName),
                "saved": bool(doc.Saved),
                "insertion_units": str(doc.GetVariable("INSUNITS")),
                "linear_units": str(doc.GetVariable("LUNITS")),
                "angular_units": str(doc.GetVariable("AUNITS")),
                "precision": str(doc.GetVariable("LUPREC")),
                "coordinate_system": str(
                    getattr(doc, "CoordinateSystemCode", "not set")
                ),
            }
        except Exception as exc:
            raise Civil3DError(f"get_drawing_info failed: {exc}") from exc

    def list_object_types(self) -> dict[str, Any]:
        """
        Count Civil 3D objects by type.
        Returns two dicts:
          'model_space_counts'  – entity counts from model space iteration
          'civil3d_collections' – counts from named Civil 3D collections
        """
        self._ensure_connected()
        # --- model space entity tally ---
        model_counts: dict[str, int] = {}
        try:
            for obj in self._doc.ModelSpace:
                name = getattr(obj, "ObjectName", "Unknown")
                model_counts[name] = model_counts.get(name, 0) + 1
        except Exception as exc:
            log.warning("Model space iteration partial: %s", exc)

        # --- Civil 3D named collections ---
        civil_counts: dict[str, int] = {}
        collection_attrs = {
            "CogoPoints": "cogo_points",
            "AlignmentsSiteless": "alignments",
            "Surfaces": "surfaces",
            "Corridors": "corridors",
            "Sites": "sites",
            "Profiles": "profiles",
            "PipeNetworks": "pipe_networks",
        }
        for attr, label in collection_attrs.items():
            try:
                col = getattr(self._doc, attr, None)
                if col is not None:
                    civil_counts[label] = int(col.Count)
            except Exception:
                pass

        return {
            "model_space_counts": model_counts,
            "civil3d_collections": civil_counts,
            "total_model_objects": sum(model_counts.values()),
        }

    def get_selected_objects_info(self, max_count: int = 10) -> list[dict[str, Any]]:
        """Return properties of currently selected Civil 3D objects."""
        self._ensure_connected()
        results: list[dict[str, Any]] = []
        try:
            # PickfirstSelectionSet is the live current selection (grips/pickfirst)
            active_ss = getattr(self._doc, "PickfirstSelectionSet", None)
            if active_ss is None or active_ss.Count == 0:
                # Fall back: find the PICKFIRST named set in SelectionSets
                active_ss = None
                ss = self._doc.SelectionSets
                for i in range(ss.Count):
                    s = ss.Item(i)
                    if s.Name.upper() == "PICKFIRST" and s.Count > 0:
                        active_ss = s
                        break
            if active_ss is None or active_ss.Count == 0:
                return []
            for i, raw in enumerate(active_ss):
                if i >= max_count:
                    break
                object_name = getattr(raw, "ObjectName", "Unknown")
                # Re-dispatch to the concrete COM interface so type-specific
                # attributes (StartPoint, Coordinates, Length, etc.) are available.
                obj = w32.Dispatch(raw)
                info: dict[str, Any] = {
                    "index": i,
                    "object_name": object_name,
                    "handle": str(getattr(obj, "Handle", "N/A")),
                    "layer": str(getattr(obj, "Layer", "N/A")),
                    "object_id": str(getattr(obj, "ObjectID", "N/A")),
                }
                # Named Civil 3D / AutoCAD attributes
                for attr in ("Name", "Description", "StyleName"):
                    val = getattr(obj, attr, None)
                    if val is not None:
                        info[attr.lower()] = str(val)
                # Length (lines, polylines, alignments, …)
                try:
                    info["length"] = float(obj.Length)
                except Exception:
                    pass
                # Geometry by entity type
                if object_name == "AcDbLine":
                    try:
                        sp, ep = obj.StartPoint, obj.EndPoint
                        info["start"] = {"x": float(sp[0]), "y": float(sp[1]), "z": float(sp[2])}
                        info["end"]   = {"x": float(ep[0]), "y": float(ep[1]), "z": float(ep[2])}
                    except Exception:
                        pass
                elif object_name in ("AcDb2dPolyline", "AcDb3dPolyline"):
                    try:
                        coords = obj.Coordinates
                        verts = [
                            {"x": float(coords[j]), "y": float(coords[j + 1]), "z": float(coords[j + 2])}
                            for j in range(0, len(coords) - 2, 3)
                        ]
                        info["vertex_count"] = len(verts)
                        if verts:
                            info["start"] = verts[0]
                            info["end"] = verts[-1]
                        info["vertices"] = verts
                    except Exception:
                        pass
                elif object_name == "AcDbPolyline":
                    try:
                        coords = obj.Coordinates
                        elev = 0.0
                        try:
                            elev = float(obj.Elevation)
                        except Exception:
                            pass
                        verts = [
                            {"x": float(coords[j]), "y": float(coords[j + 1]), "z": elev}
                            for j in range(0, len(coords) - 1, 2)
                        ]
                        info["vertex_count"] = len(verts)
                        if verts:
                            info["start"] = verts[0]
                            info["end"] = verts[-1]
                        info["vertices"] = verts
                    except Exception:
                        pass
                elif object_name == "AcDbArc":
                    try:
                        info["center"] = {"x": float(obj.Center[0]), "y": float(obj.Center[1]), "z": float(obj.Center[2])}
                        info["radius"] = float(obj.Radius)
                        info["start_angle_deg"] = float(obj.StartAngle) * 180.0 / 3.141592653589793
                        info["end_angle_deg"]   = float(obj.EndAngle)   * 180.0 / 3.141592653589793
                    except Exception:
                        pass
                elif object_name == "AcDbCircle":
                    try:
                        info["center"] = {"x": float(obj.Center[0]), "y": float(obj.Center[1]), "z": float(obj.Center[2])}
                        info["radius"] = float(obj.Radius)
                    except Exception:
                        pass
                results.append(info)
        except Exception as exc:
            raise Civil3DError(f"get_selected_objects_info failed: {exc}") from exc
        return results

    # ------------------------------------------------------------------
    # COGO Points
    # ------------------------------------------------------------------

    def _get_cogo_collection(self) -> Any:
        # Try self._doc first (same source list_object_types uses), then _civil_doc()
        docs_to_try: list[Any] = [self._doc]
        civil_doc = self._civil_doc()
        if civil_doc is not self._doc:
            docs_to_try.append(civil_doc)

        for doc in docs_to_try:
            col = getattr(doc, "CogoPoints", None)
            if col is not None:
                try:
                    _ = col.Count  # verify interface is live
                    log.debug("_get_cogo_collection: found CogoPoints from doc %r", doc)
                    return col
                except Exception as exc:
                    log.debug("_get_cogo_collection: CogoPoints.Count failed on %r: %s", doc, exc)

        raise Civil3DError(
            "CogoPoints collection not found. "
            "Ensure Civil 3D (not plain AutoCAD) is running."
        )

    def create_cogo_point(
        self,
        northing: float,
        easting: float,
        elevation: float = 0.0,
        description: str = "",
    ) -> dict[str, Any]:
        """Create a COGO point and return its properties."""
        self._ensure_connected()
        try:
            pts = self._get_cogo_collection()
            pt_id = pts.Add(easting, northing, elevation, description)
            pt = pts.Find(pt_id)
            return {
                "point_number": int(pt.PointNumber),
                "northing": float(pt.Northing),
                "easting": float(pt.Easting),
                "elevation": float(pt.Elevation),
                "description": str(pt.RawDescription),
            }
        except Civil3DError:
            raise
        except Exception as exc:
            raise Civil3DError(f"create_cogo_point failed: {exc}") from exc

    def list_cogo_points(self, max_count: int = 50) -> list[dict[str, Any]]:
        """Return up to max_count COGO points."""
        self._ensure_connected()
        results: list[dict[str, Any]] = []
        try:
            pts = self._get_cogo_collection()
            for i, pt in enumerate(pts):
                if i >= max_count:
                    break
                results.append({
                    "point_number": int(pt.PointNumber),
                    "northing": float(pt.Northing),
                    "easting": float(pt.Easting),
                    "elevation": float(pt.Elevation),
                    "description": str(pt.RawDescription),
                })
        except Civil3DError:
            raise
        except Exception as exc:
            raise Civil3DError(f"list_cogo_points failed: {exc}") from exc
        return results

    def delete_cogo_point(self, point_number: int) -> bool:
        """Delete a COGO point by its point number."""
        self._ensure_connected()
        try:
            pts = self._get_cogo_collection()
            pt = pts.FindByPointNumber(point_number)
            if pt is None:
                raise Civil3DError(f"Point {point_number} not found.")
            pts.Delete(pt.PointNumber)
            return True
        except Civil3DError:
            raise
        except Exception as exc:
            raise Civil3DError(f"delete_cogo_point failed: {exc}") from exc

    # ------------------------------------------------------------------
    # Lines & Polylines
    # ------------------------------------------------------------------

    def create_line(
        self,
        x1: float, y1: float, z1: float,
        x2: float, y2: float, z2: float,
        layer: str = "0",
    ) -> dict[str, Any]:
        """Create a 3-D line in model space."""
        self._ensure_connected()
        try:
            # Use the base AutoCAD document for geometry creation — the Civil 3D
            # AeccDocument wrapper does not reliably forward AddLine/Add3DPoly.
            acad_doc = self._acad.ActiveDocument
            mspace = acad_doc.ModelSpace
            line = mspace.AddLine(self._pt3d(x1, y1, z1), self._pt3d(x2, y2, z2))
            line.Layer = layer
            line.Update()
            try:
                acad_doc.Regen(1)
            except Exception:
                pass
            return {
                "object_id": str(line.ObjectID),
                "handle": str(line.Handle),
                "layer": str(line.Layer),
                "start": {"x": x1, "y": y1, "z": z1},
                "end": {"x": x2, "y": y2, "z": z2},
                "length": float(line.Length),
            }
        except Civil3DError:
            raise
        except Exception as exc:
            raise Civil3DError(f"create_line failed: {exc}") from exc

    def create_polyline(
        self,
        vertices: list[tuple[float, float, float]],
        closed: bool = False,
        layer: str = "0",
    ) -> dict[str, Any]:
        """Create a 3-D polyline from an ordered list of (x, y, z) tuples."""
        self._ensure_connected()
        if len(vertices) < 2:
            raise Civil3DError("At least 2 vertices are required.")
        try:
            flat = [c for v in vertices for c in v]
            pt_array = w32.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, flat
            )
            # Use the base AutoCAD document for geometry creation — the Civil 3D
            # AeccDocument wrapper does not reliably forward Add3DPoly.
            acad_doc = self._acad.ActiveDocument
            mspace = acad_doc.ModelSpace
            pline = mspace.Add3DPoly(pt_array)
            pline.Layer = layer
            if closed:
                pline.Closed = True
            pline.Update()
            try:
                acad_doc.Regen(1)
            except Exception:
                pass
            return {
                "object_id": str(pline.ObjectID),
                "handle": str(pline.Handle),
                "layer": str(pline.Layer),
                "vertex_count": len(vertices),
                "closed": closed,
                "length": float(pline.Length),
            }
        except Civil3DError:
            raise
        except Exception as exc:
            raise Civil3DError(f"create_polyline failed: {exc}") from exc

    def list_lines(self, layer_filter: str = "") -> list[dict[str, Any]]:
        """List lines and polylines in model space, optionally filtered by layer."""
        self._ensure_connected()
        results: list[dict[str, Any]] = []
        _line_types = {"AcDbLine", "AcDb3dPolyline", "AcDbPolyline", "AcDb2dPolyline"}
        try:
            for raw in self._doc.ModelSpace:
                object_name = raw.ObjectName
                if object_name not in _line_types:
                    continue
                if layer_filter and raw.Layer.lower() != layer_filter.lower():
                    continue
                # ModelSpace iteration returns base IAcadEntity; re-dispatch to the
                # concrete interface so type-specific attributes (StartPoint etc.) work.
                obj = w32.Dispatch(raw)
                info: dict[str, Any] = {
                    "object_name": object_name,
                    "handle": str(obj.Handle),
                    "layer": str(obj.Layer),
                }
                try:
                    info["length"] = float(obj.Length)
                except Exception:
                    pass
                if object_name == "AcDbLine":
                    sp, ep = obj.StartPoint, obj.EndPoint
                    info["start"] = {"x": float(sp[0]), "y": float(sp[1]), "z": float(sp[2])}
                    info["end"]   = {"x": float(ep[0]), "y": float(ep[1]), "z": float(ep[2])}
                elif object_name in ("AcDb2dPolyline", "AcDb3dPolyline"):
                    # Coordinates are a flat (x, y, z, x, y, z, ...) tuple
                    coords = obj.Coordinates
                    verts = [
                        {"x": float(coords[i]), "y": float(coords[i + 1]), "z": float(coords[i + 2])}
                        for i in range(0, len(coords) - 2, 3)
                    ]
                    info["vertex_count"] = len(verts)
                    if verts:
                        info["start"] = verts[0]
                        info["end"] = verts[-1]
                    info["vertices"] = verts
                    try:
                        info["closed"] = bool(obj.Closed)
                    except Exception:
                        pass
                elif object_name == "AcDbPolyline":
                    # LWPolyline: flat (x, y, x, y, ...) 2-D pairs; elevation separate
                    coords = obj.Coordinates
                    elev = 0.0
                    try:
                        elev = float(obj.Elevation)
                    except Exception:
                        pass
                    verts = [
                        {"x": float(coords[i]), "y": float(coords[i + 1]), "z": elev}
                        for i in range(0, len(coords) - 1, 2)
                    ]
                    info["vertex_count"] = len(verts)
                    if verts:
                        info["start"] = verts[0]
                        info["end"] = verts[-1]
                    info["vertices"] = verts
                    try:
                        info["closed"] = bool(obj.Closed)
                    except Exception:
                        pass
                results.append(info)
        except Exception as exc:
            raise Civil3DError(f"list_lines failed: {exc}") from exc
        return results

    # ------------------------------------------------------------------
    # Surfaces
    # ------------------------------------------------------------------

    def _civil_doc(self) -> Any:
        """Return the Civil 3D document interface.

        Prefers self._civil.ActiveDocument (guaranteed AeccDocument) over
        self._doc which may be a base AcadDocument if the Civil 3D interface
        was not fully acquired.
        """
        if self._civil is not None:
            try:
                return self._civil.ActiveDocument
            except Exception:
                pass
        return self._doc

    def _get_surfaces(self) -> list[Any]:
        """Collect all surfaces; falls back to model-space scan if needed."""
        # Try every possible document source.  list_object_types() successfully
        # reads Surfaces.Count from self._doc directly — mirror that here first.
        docs_to_try: list[Any] = []
        docs_to_try.append(self._doc)  # same source as list_object_types
        civil_doc = self._civil_doc()
        if civil_doc is not self._doc:
            docs_to_try.append(civil_doc)

        # 2026-08-06 추가: 컬렉션 접근 자체가 성공했는지를 별도로 기록한다.
        # (원본은 "서피스 0개"와 "컬렉션 접근 실패"를 구분하지 못해,
        #  빈 도면에서도 "Civil 3D가 아닌 AutoCAD가 실행 중"이라는 잘못된
        #  오류를 냈다.)
        collection_found = False
        for doc in docs_to_try:
            col = getattr(doc, "Surfaces", None)
            if col is None:
                continue
            count = 0
            try:
                count = int(col.Count)
                collection_found = True
            except Exception:
                pass
            log.debug("_get_surfaces: Surfaces.Count=%d from doc %r", count, doc)
            if count == 0:
                continue
            surfaces = self._iter_com_collection(col)
            if surfaces:
                return surfaces
            log.warning(
                "_get_surfaces: Surfaces.Count=%d but iteration returned nothing; "
                "trying next document source or model-space fallback.", count
            )

        # Fallback: scan model space for TIN/Grid surface entities
        #
        # 2026-08-16 수정: 이 경고는 collection_found 와 무관하게 발생하고 있었다.
        # 컬렉션에 정상 접근했고 단지 서피스가 0개인 빈 도면에서도
        # "inaccessible"이라고 보고해, 진단하는 사람을 연결 실패 쪽으로 오도한다.
        # 반환값은 이미 옳았으므로(빈 리스트) 기능 결함은 아니지만,
        # 원본 결함 B("서피스 0개"를 "연결 실패"로 오보)와 같은 성격의 오보가
        # 로그 계층에 남아 있던 것이다. 실제 상태에 따라 문구를 나눈다.
        if collection_found:
            log.info(
                "Surfaces collection accessible but reported 0 surfaces; "
                "running a model-space scan as a cross-check."
            )
        else:
            log.warning(
                "Civil 3D Surfaces collection inaccessible — "
                "falling back to model space scan."
            )
        results: list[Any] = []
        # ObjectName examples: AeccDbTinSurface, AeccDbGridSurface,
        # AeccDbTinVolumeSurface, AeccDbGridVolumeSurface
        _surface_types = (
            "aeccdbtin", "aeccdbgrid", "aeccdbtinvolume", "aeccdbgridvolume",
            "surface",  # catch-all in case naming convention differs
        )
        scan_doc = self._doc
        try:
            for raw in scan_doc.ModelSpace:
                obj_name = getattr(raw, "ObjectName", "").lower()
                if any(t in obj_name for t in _surface_types):
                    obj = w32.Dispatch(raw)
                    if getattr(obj, "Name", None) is not None:
                        results.append(obj)
        except Exception as exc:
            log.warning("Surface model-space scan failed: %s", exc)

        if not results:
            if collection_found:
                # 컬렉션은 정상 접근됐고 단지 서피스가 없는 상태 — 정상이며 오류가 아니다.
                log.info("Surfaces collection accessible but empty (0 surfaces).")
                return []
            raise Civil3DError(
                "Surfaces collection not found. "
                "Ensure Civil 3D (not plain AutoCAD) is running."
            )
        return results

    def list_surfaces(self) -> list[dict[str, Any]]:
        """List all TIN/Grid surfaces with basic elevation statistics."""
        self._ensure_connected()
        results: list[dict[str, Any]] = []
        try:
            for surf in self._get_surfaces():
                info: dict[str, Any] = {
                    "name": str(surf.Name),
                    "description": str(getattr(surf, "Description", "")),
                    "style": str(getattr(surf, "StyleName", "")),
                    "object_id": str(surf.ObjectID),
                }
                try:
                    st = surf.Statistics
                    info["min_elevation"] = float(st.MinElevation)
                    info["max_elevation"] = float(st.MaxElevation)
                    info["mean_elevation"] = float(st.MeanElevation)
                    info["point_count"] = int(st.NumberOfPoints)
                    info["triangle_count"] = int(st.NumberOfTriangles)
                except Exception:
                    pass
                results.append(info)
        except Civil3DError:
            raise
        except Exception as exc:
            raise Civil3DError(f"list_surfaces failed: {exc}") from exc
        return results

    def get_surface_info(self, surface_name: str) -> dict[str, Any]:
        """Return detailed statistics for a named surface."""
        self._ensure_connected()
        surf = self._find_surface(surface_name)
        info: dict[str, Any] = {
            "name": str(surf.Name),
            "description": str(getattr(surf, "Description", "")),
            "style": str(getattr(surf, "StyleName", "")),
        }
        try:
            st = surf.Statistics
            info.update({
                "min_elevation": float(st.MinElevation),
                "max_elevation": float(st.MaxElevation),
                "mean_elevation": float(st.MeanElevation),
                "point_count": int(st.NumberOfPoints),
                "triangle_count": int(st.NumberOfTriangles),
                "area_2d": float(st.Area2d),
                "area_3d": float(st.Area3d),
                "min_x": float(st.MinX),
                "min_y": float(st.MinY),
                "max_x": float(st.MaxX),
                "max_y": float(st.MaxY),
            })
        except Exception as exc:
            log.warning("Could not read surface statistics: %s", exc)
        return info

    def sample_surface_elevation(
        self,
        surface_name: str,
        easting: float,
        northing: float,
    ) -> dict[str, Any]:
        """Sample interpolated elevation of a surface at (easting, northing)."""
        self._ensure_connected()
        surf = self._find_surface(surface_name)
        try:
            elev = surf.FindElevationAtXY(easting, northing)
            return {
                "surface_name": surface_name,
                "easting": easting,
                "northing": northing,
                "elevation": float(elev),
            }
        except Exception as exc:
            raise Civil3DError(
                f"sample_surface_elevation failed "
                f"(point may be outside surface extent): {exc}"
            ) from exc

    def _find_surface(self, name: str) -> Any:
        for surf in self._get_surfaces():
            if getattr(surf, "Name", "").lower() == name.lower():
                return surf
        raise Civil3DError(f"Surface '{name}' not found in the active drawing.")

    # ------------------------------------------------------------------
    # Volume between surfaces  (신규 2026-08-13)
    #
    # 기반 저장소에는 서피스 '읽기' 도구만 있고 두 서피스 사이의 체적을
    # 계산하는 연산이 없었다. 이 메서드가 그 공백을 메우며, 상위 도메인
    # 도구(compute_earthwork_by_rock_quality)가 호출하는 프리미티브다.
    # ------------------------------------------------------------------

    # AeccXLand coclass 버전 접미사. connect()에서 성공한 ProgID의 접미사를
    # 최우선으로 쓰고, 나머지는 폴백.
    _AECC_LAND_VERSIONS = ["13.8", "13.7", "13.6", "14.4", "14.0", "13.0"]

    def _aecc_versions_to_try(self) -> list[str]:
        versions: list[str] = []
        if self._prog_id:
            tail = self._prog_id.rsplit(".", 2)[-2:]
            if len(tail) == 2:
                versions.append(".".join(tail))
        for ver in self._AECC_LAND_VERSIONS:
            if ver not in versions:
                versions.append(ver)
        return versions

    def _new_aecc_object(self, type_name: str) -> Any:
        """Instantiate an AeccXLand coclass such as AeccTinVolumeCreationData.

        These coclasses are served by the running Civil 3D process, so they are
        obtained through acad.GetInterfaceObject() rather than CreateObject().
        The same pattern is already proven for AeccTinCreationData.
        """
        self._ensure_connected()
        errors: list[str] = []
        for ver in self._aecc_versions_to_try():
            prog_id = f"AeccXLand.{type_name}.{ver}"
            try:
                obj = self._acad.GetInterfaceObject(prog_id)
                if obj is not None:
                    log.debug("Instantiated %s", prog_id)
                    return obj
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{prog_id} -> {exc}")
        raise Civil3DError(
            f"Could not instantiate AeccXLand.{type_name}. "
            f"Tried {len(errors)} ProgIDs; first failures: " + " | ".join(errors[:3])
        )

    # 도면 단위 — 2026-08-16 추가.
    #
    # 반환 키가 cut_m3 로 ㎥ 를 단정하는데 도면 단위 검사가 없었다. Civil 3D 는
    # **도면 단위**로 값을 돌려주므로 mm 도면이면 부피가 배율의 세제곱만큼,
    # 즉 10⁹ 배 어긋난 채 '㎥'라는 이름으로 보고된다. 결과만 보아서는 알 수 없다.
    #
    # AutoCAD INSUNITS 코드 → 미터 환산 계수. 실무에서 나올 수 있는 것만 둔다.
    _INSUNITS_TO_M: dict[int, tuple[float, str]] = {
        1: (0.0254, "inch"),
        2: (0.3048, "foot"),
        3: (1609.344, "mile"),
        4: (0.001, "millimeter"),
        5: (0.01, "centimeter"),
        6: (1.0, "meter"),
        7: (1000.0, "kilometer"),
        10: (0.9144, "yard"),
        14: (0.1, "decimeter"),
        15: (10.0, "dekameter"),
        16: (100.0, "hectometer"),
    }

    def _drawing_length_unit(self) -> dict[str, Any]:
        """도면의 길이 단위와 미터 환산 계수를 낸다.

        GetVariable("INSUNITS") 가 COM 열거형/문자열로 오는 경우가 있어
        (2026-08-06 실측 기록) 정수로 명시 정규화한다.
        """
        raw: Any = None
        try:
            raw = self._doc.GetVariable("INSUNITS")
        except Exception as exc:  # noqa: BLE001
            log.warning("INSUNITS could not be read: %s", exc)

        code: int | None = None
        if raw is not None:
            try:
                code = int(str(raw).strip())
            except (TypeError, ValueError):
                log.warning("INSUNITS value %r is not an integer", raw)

        if code in self._INSUNITS_TO_M:
            scale, name = self._INSUNITS_TO_M[code]
            return {
                "insunits": code, "name": name,
                "scale_to_m": scale, "assumed": False,
            }
        return {
            "insunits": code, "name": "unspecified" if code == 0 else "unknown",
            "scale_to_m": 1.0, "assumed": True,
        }

    # 경계 클립 — 2026-08-16 실측으로 확보한 경로.
    #
    # 벤더 문서에는 산정 영역을 한정하는 방법이 없어 미확인으로 남아 있었으나,
    # 타입 라이브러리를 직접 열람한 결과 볼륨 서피스의 Statistics 에
    # BoundedVolumes(varPoints, pCut, pFill, pNet) 가 존재했다.
    #
    # 실측으로 확인한 호출 조건(문서화된 곳이 없다):
    #   · 좌표는 반드시 3D (x, y, z). 2D 쌍만 넘기면 COM 예외.
    #   · 폴리곤을 **명시적으로 닫아야** 한다(첫 점을 끝에 반복).
    #     열린 폴리곤은 서피스 전체 범위를 덮을 때는 우연히 통과하고
    #     부분 영역에서만 실패한다 — 즉 시험을 전체 범위로만 하면
    #     통과한 것처럼 보인다. 가장 위험한 형태의 오류다.
    #   · SAFEARRAY 는 VT_R8 이어야 한다. VT_VARIANT 배열은 전부 실패.
    #   · z 값은 결과에 영향이 없다(z=0 과 z=95 가 동일한 값을 반환).
    #
    # 검증: 100 m x 100 m, 두 평면 EL 100 / EL 90 (전체 절토 100,000 ㎥)에서
    # x∈[0,50] 폴리곤을 넘겨 50,000.00 ㎥ 를 얻었다(해석해와 정확히 일치).

    _BOUNDARY_MIN_POINTS = 3

    @staticmethod
    def _normalize_boundary_points(coords: Any) -> list[float]:
        """경계 좌표를 BoundedVolumes 가 받는 평탄 3D 배열로 정규화한다."""
        pts: list[tuple[float, float]] = []

        flat: list[float] = []
        if isinstance(coords, (list, tuple)) and coords and isinstance(
            coords[0], (list, tuple)
        ):
            for item in coords:
                if len(item) < 2:
                    raise Civil3DError(
                        f"boundary point {item!r} needs at least an x and a y."
                    )
                pts.append((float(item[0]), float(item[1])))
        else:
            flat = [float(v) for v in coords]
            if len(flat) % 3 == 0 and len(flat) % 2 != 0:
                stride = 3
            elif len(flat) % 2 == 0 and len(flat) % 3 != 0:
                stride = 2
            elif len(flat) % 3 == 0:
                # 6, 12, 18 … 은 양쪽으로 읽히므로 3D 로 본다(더 흔한 형식).
                stride = 3
            else:
                raise Civil3DError(
                    f"boundary has {len(flat)} numbers, which is not a whole "
                    f"number of 2D or 3D points."
                )
            pts = [(flat[i], flat[i + 1]) for i in range(0, len(flat), stride)]

        if len(pts) >= 2 and pts[0] == pts[-1]:
            pts = pts[:-1]                      # 중복 폐합점 제거 후 다시 붙인다
        if len(pts) < Civil3DClient._BOUNDARY_MIN_POINTS:
            raise Civil3DError(
                f"boundary needs at least {Civil3DClient._BOUNDARY_MIN_POINTS} "
                f"distinct points to form a polygon; got {len(pts)}."
            )

        closed = [*pts, pts[0]]                 # 명시적 폐합 — 생략하면 조용히 틀린다
        out: list[float] = []
        for x, y in closed:
            out.extend([x, y, 0.0])
        return out

    def _resolve_boundary(self, boundary: Any) -> tuple[list[float], str]:
        """boundary 인자를 좌표 배열로 바꾼다.

        받는 형태는 둘이다.
          · 문자열  — 도면 안 폐합 폴리라인의 **핸들**(AutoCAD 엔티티 핸들)
          · 좌표열  — [[x, y], …] 또는 [x, y, z, x, y, z, …]
        """
        if isinstance(boundary, str):
            handle = boundary.strip()
            try:
                acad_doc = self._acad.ActiveDocument
                ent = acad_doc.HandleToObject(handle)
            except Exception as exc:  # noqa: BLE001
                raise Civil3DError(
                    f"boundary '{handle}' could not be resolved to a drawing "
                    f"object. Pass the entity handle of a closed polyline, or "
                    f"pass the vertex coordinates directly: {exc}"
                ) from exc
            obj_name = str(getattr(ent, "ObjectName", ""))
            if "Polyline" not in obj_name:
                raise Civil3DError(
                    f"boundary handle '{handle}' refers to a {obj_name}, not a "
                    f"polyline. The boundary must be a closed polyline."
                )
            try:
                closed_flag = bool(ent.Closed)
            except Exception:  # noqa: BLE001
                closed_flag = True
            if not closed_flag:
                raise Civil3DError(
                    f"polyline '{handle}' is not closed. An open boundary would "
                    f"be silently completed by the API and the clipped region "
                    f"would not be the one you drew."
                )
            coords = list(ent.Coordinates)
            return self._normalize_boundary_points(coords), f"polyline handle {handle}"

        return self._normalize_boundary_points(boundary), "explicit coordinates"

    def compute_volume_between_surfaces(
        self,
        base_surface: str,
        comparison_surface: str,
        boundary: Any = None,
    ) -> dict[str, Any]:
        """Cut / fill / net volume between two TIN surfaces.

        A TIN volume surface is created temporarily, its statistics are read,
        and it is erased again so the drawing is left unchanged.

        Sign convention (Civil 3D): the volume surface represents
        ``comparison - base``.  With base = existing ground and comparison =
        design (corridor datum), CutVolume is the material to be removed.

        The computation covers the *overlap* of the two surfaces.  When the
        design surface is a corridor datum surface its own extent already
        limits the region to the graded area plus its slopes, which is the
        intended region for earthwork quantities.
        """
        import time

        self._ensure_connected()
        if base_surface.strip().lower() == comparison_surface.strip().lower():
            raise Civil3DError(
                "base_surface and comparison_surface must be different surfaces."
            )
        boundary_points: list[float] | None = None
        boundary_source = ""
        if boundary is not None:
            boundary_points, boundary_source = self._resolve_boundary(boundary)

        units = self._drawing_length_unit()
        vol_scale = units["scale_to_m"] ** 3          # 부피는 배율의 세제곱
        area_scale = units["scale_to_m"] ** 2

        started = time.perf_counter()
        base = self._find_surface(base_surface)
        comp = self._find_surface(comparison_surface)

        # 스타일은 필수 입력이며 이름이 템플릿·로케일에 따라 다르다.
        # 기존 서피스의 스타일을 재사용하는 것이 가장 안전하다.
        style_name = ""
        for cand in (base, comp):
            try:
                style_name = str(cand.StyleName)
                if style_name:
                    break
            except Exception:  # noqa: BLE001
                continue

        layer_name = "0"
        try:
            layer_name = str(base.Layer) or "0"
        except Exception:  # noqa: BLE001
            pass

        temp_name = f"__mcp_vol_{int(time.time() * 1000) % 100_000_000}"
        vol_surface = None
        cut = fill = net = 0.0
        area_2d: float | None = None
        bbox: dict[str, float] | None = None
        erase_ok = True
        erase_error = ""

        try:
            data = self._new_aecc_object("AeccTinVolumeCreationData")
            data.Name = temp_name
            data.Description = "temporary volume surface created by civil3d-mcp"
            data.Layer = layer_name
            data.BaseLayer = layer_name
            if style_name:
                try:
                    data.Style = style_name
                except Exception as exc:  # noqa: BLE001
                    log.warning("Could not set volume surface style '%s': %s",
                                style_name, exc)
            data.BaseSurface = base
            data.ComparisonSurface = comp

            vol_surface = self._doc.Surfaces.AddTinVolumeSurface(data)

            stats = vol_surface.Statistics
            if boundary_points is None:
                cut = float(stats.CutVolume)
                fill = float(stats.FillVolume)
                net = float(stats.NetVolume)
            else:
                arr = w32.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_R8, boundary_points
                )
                bounded = stats.BoundedVolumes(arr)
                cut, fill, net = (float(v) for v in bounded[:3])

            # 볼륨 서피스의 Statistics 에는 Area2d 가 없다(실측 확인 2026-08-16).
            # 일반 TIN 서피스에는 있으나 볼륨 서피스 인터페이스에는 노출되지
            # 않으므로, 산정 영역은 면적 대신 경계 상자로 보고한다.
            try:
                bbox = {
                    "min_x": float(stats.MinX), "min_y": float(stats.MinY),
                    "max_x": float(stats.MaxX), "max_y": float(stats.MaxY),
                }
                area_2d = None
            except Exception:  # noqa: BLE001
                bbox = None
        except Civil3DError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise Civil3DError(
                f"compute_volume_between_surfaces failed for "
                f"('{base_surface}' -> '{comparison_surface}'): {exc}"
            ) from exc
        finally:
            # 도면 오염 방지 — 실패 경로에서도 임시 객체를 남기지 않는다(원칙 P4).
            #
            # 2026-08-16 수정: 삭제 실패를 로그로만 남기고 호출자에게는 정상
            # 응답을 주고 있었다. 도면이 조용히 오염되는데 결과만 보아서는
            # 알 수 없고, 반복 실행의 초기 조건이 같아야 재현성 측정이 성립하므로
            # 이 사실은 반드시 응답에 나와야 한다.
            if vol_surface is not None:
                try:
                    vol_surface.Erase()
                    erase_ok = True
                except Exception as exc:  # noqa: BLE001
                    erase_ok = False
                    erase_error = str(exc)
                    log.warning(
                        "Temporary volume surface '%s' could not be erased "
                        "and remains in the drawing: %s", temp_name, exc,
                    )

        warnings: list[str] = []
        if units["assumed"]:
            warnings.append(
                f"Drawing units are {units['name']} (INSUNITS={units['insunits']}), "
                f"so the values could not be converted and are reported in drawing "
                f"units cubed, not cubic metres. Set INSUNITS on the drawing to get "
                f"a checked conversion. Volume error scales with the cube of the "
                f"length ratio — a millimetre drawing read as metres is off by 10^9."
            )
        if not erase_ok:
            warnings.append(
                f"The temporary volume surface '{temp_name}' could not be erased "
                f"and remains in the drawing: {erase_error}. Delete it before "
                f"measuring repeatability — repeated runs no longer start from the "
                f"same drawing state."
            )

        result: dict[str, Any] = {
            "surface_pair": {
                "base": str(base.Name),
                "comparison": str(comp.Name),
            },
            "cut_m3": cut * vol_scale,
            "fill_m3": fill * vol_scale,
            "net_m3": net * vol_scale,
            "drawing_units": units,
            "temp_surface_erased": erase_ok,
            "warnings": warnings,
            "boundary_applied": boundary_points is not None,
            "region": (
                f"clipped to the supplied boundary ({boundary_source})"
                if boundary_points is not None
                else "overlap of the two surfaces"
            ),
            "elapsed_ms": int((time.perf_counter() - started) * 1000),
        }
        if boundary_points is not None:
            # 폐합점을 제외한 정점 수와 폴리곤 면적(신발끈 공식)을 함께 돌려준다.
            xy = [(boundary_points[i], boundary_points[i + 1])
                  for i in range(0, len(boundary_points), 3)][:-1]
            shoelace = abs(sum(
                xy[i][0] * xy[(i + 1) % len(xy)][1]
                - xy[(i + 1) % len(xy)][0] * xy[i][1]
                for i in range(len(xy))
            )) / 2.0
            result["boundary_vertex_count"] = len(xy)
            result["bounded_area_m2"] = shoelace * area_scale
        elif area_2d is not None:
            result["bounded_area_m2"] = area_2d * area_scale
        if bbox is not None:
            # 볼륨 서피스에는 Area2d 가 없으므로 산정 영역은 경계 상자로 보고한다.
            # 실제 겹침 영역의 상계이지 그 면적 자체가 아니다.
            result["region_bbox"] = {
                k: v * units["scale_to_m"] for k, v in bbox.items()
            }
        return result

    # ------------------------------------------------------------------
    # Earthwork by rock quality  (신규 2026-08-16)
    #
    # 2계층 구조의 상위 도메인 도구. compute_volume_between_surfaces 를
    # 경계면 수(n+1)회 호출하고, 결과를 earthwork_core 의 차분식으로 조합한다.
    # 이 메서드에는 CAD 조작만 있고 공학 판단은 전부 earthwork_core 에 있다.
    # ------------------------------------------------------------------

    def compute_earthwork_by_rock_quality(
        self,
        ground_surface: str,
        design_surface: str,
        strata: list[dict[str, Any]],
        boundary: str | None = None,
        site_id: str | None = None,
        below_lowest_unit_cost: float | None = None,
    ) -> dict[str, Any]:
        """암질별 절토량과 공사비를 산정한다.

        ``strata`` 는 상→하로 정렬된 암층 **경계면** 목록이며 각 항목은
        ``{"name": str, "surface": str, "unit_cost": float | None}`` 이다.
        경계면이므로 예컨대 '토사층'의 surface 는 *토사층의 하면*이다.

        산정 영역은 두 서피스의 겹침이다. ``design_surface`` 가 코리더 데이텀
        서피스이면 그 외곽이 이미 조성 부지와 사면을 포함하므로 이것이
        의도한 영역이다. 부지 경계선으로 클립하면 사면부 물량이 빠져
        계통적 과소 산정이 된다.
        """
        import time

        from . import earthwork_core as core

        self._ensure_connected()
        started = time.perf_counter()

        if not strata:
            raise Civil3DError(
                "strata is empty. Supply the stratum boundary surfaces ordered "
                "from top to bottom, e.g. "
                '[{"name": "토사층", "surface": "01토사층(Tin)", "unit_cost": 4500}]'
            )

        names: list[str] = []
        surfaces: list[str] = []
        unit_costs: list[float | None] = []
        for i, item in enumerate(strata):
            if not isinstance(item, dict):
                raise Civil3DError(f"strata[{i}] must be an object, got {type(item).__name__}.")
            name = item.get("name")
            surf = item.get("surface")
            if not name or not surf:
                raise Civil3DError(
                    f"strata[{i}] requires both 'name' and 'surface'. Got {item!r}."
                )
            names.append(str(name))
            surfaces.append(str(surf))
            raw_cost = item.get("unit_cost")
            unit_costs.append(None if raw_cost is None else float(raw_cost))

        # C(S) 를 얻는다 — 경계면마다 (경계면, 계획고) 조합으로 1회씩.
        # 성토는 원지반 호출에서 함께 나온다.
        ground_result = self.compute_volume_between_surfaces(
            ground_surface, design_surface, boundary
        )
        c_ground = float(ground_result["cut_m3"])
        total_fill = float(ground_result["fill_m3"])
        area_m2 = ground_result.get("bounded_area_m2")
        units = ground_result.get("drawing_units", {})

        # 하위 호출이 낸 경고(단위 미확정·임시 서피스 잔존)를 삼키지 않는다.
        volume_warnings: list[str] = list(ground_result.get("warnings", []))

        # 명세가 밝힌 전제 — "절토(원지반 > 계획고) 구간 대상" — 를 확인한다.
        #
        # 2026-08-17 추가: 이 전제가 명세에는 있으나 코드에는 없었다. 두 서피스를
        # 뒤바꿔 지정하면 모든 층의 절토량이 0으로 나오는데, 수식 자체는 성립하므로
        # 오류가 발생하지 않고 **전부 0인 무의미한 결과가 조용히 반환**된다.
        # 전면 성토 부지가 있을 수 있으므로 반려하지는 않되, 가장 흔한 원인인
        # 인자 뒤바뀜을 지목해 경고한다.
        if c_ground <= core.TOL_M3 and total_fill > core.TOL_M3:
            volume_warnings.append(
                f"The design surface lies entirely above the ground surface: the cut "
                f"volume is zero while the fill volume is {total_fill:,.2f} m3, so every "
                f"stratum will report zero. This tool is specified for the cut region "
                f"(ground above design). The usual cause is that ground_surface and "
                f"design_surface were given in the wrong order — check the argument "
                f"order before using this result."
            )

        # 각 체적이 같은 산정 영역에서 나왔는지 확인하려고 영역을 함께 모은다.
        # 항등식은 Vᵢ 의 정의상 항상 성립하므로 영역 불일치를 잡지 못한다.
        regions: list[tuple[str, dict | None]] = [
            (ground_surface, ground_result.get("region_bbox"))
        ]

        c_strata: list[float] = []
        for surf_name in surfaces:
            r = self.compute_volume_between_surfaces(surf_name, design_surface, boundary)
            c_strata.append(float(r["cut_m3"]))
            regions.append((surf_name, r.get("region_bbox")))
            for w in r.get("warnings", []):
                if w not in volume_warnings:
                    volume_warnings.append(w)

        if boundary is not None:
            # 모든 호출이 같은 폴리곤으로 잘렸으므로 영역은 구조적으로 동일하다.
            area_warnings: list[str] = []
            region_basis = "explicit boundary — identical for every call"
            regions_consistent = True
        else:
            area_warnings = core.check_bbox_consistency(regions)
            region_basis = (
                "bounding box of each volume surface — detects mismatch, "
                "does not prove identity"
            )
            regions_consistent = not area_warnings

        try:
            diff = core.differential_volumes(c_ground, names, c_strata)
        except core.EarthworkError as exc:
            raise Civil3DError(str(exc)) from exc

        # 단가 부착 — 최하층 하부는 별도 인자로만 받는다.
        # 연암 단가를 경암에 물려쓰면 총액이 한 방향(과소)으로 틀어지고,
        # 결과만 보아서는 그것이 '싼 후보지'로 보인다.
        for layer, cost in zip(diff.layers, unit_costs):
            layer.unit_cost = cost
        diff.layers[-1].unit_cost = below_lowest_unit_cost

        by_method = core.aggregate_by_method(diff.layers)
        cost_total, cost_warnings = core.total_cost(by_method)

        design_elev: float | None = None
        connection_point: dict[str, float] | None = None
        length_scale = float(units.get("scale_to_m", 1.0))
        try:
            st = self._find_surface(design_surface).Statistics
            design_elev = float(st.MeanElevation) * length_scale
            connection_point = {
                "easting": (float(st.MinX) + float(st.MaxX)) / 2.0 * length_scale,
                "northing": (float(st.MinY) + float(st.MaxY)) / 2.0 * length_scale,
            }
        except Exception as exc:  # noqa: BLE001
            log.warning("Could not read design surface statistics: %s", exc)

        result: dict[str, Any] = {
            "site_id": site_id or design_surface,
            "ground_surface": ground_surface,
            "design_surface": design_surface,
            "by_method": by_method,
            "by_stratum": [
                {
                    "name": layer.name,
                    "method": layer.method,
                    "cut_m3": layer.cut_m3,
                    "unit_cost": layer.unit_cost,
                    "cost": layer.cost,
                    "assumed": layer.assumed,
                }
                for layer in diff.layers
            ],
            "total_cut_m3": diff.total_cut_m3,
            "total_fill_m3": total_fill,
            "total_earthwork_cost": cost_total,
            "boundary_applied": boundary is not None,
            "region": (
                "clipped to the supplied boundary — every stratum uses the same "
                "polygon, so the differences are taken over an identical region"
                if boundary is not None
                else "overlap of each surface pair with the design surface"
            ),
            "identity_check": {
                "expression": "sum(Vi) + C(Sn) = C(S0)",
                "left_m3": diff.total_cut_m3,
                "right_m3": c_ground,
                "residual_m3": diff.identity_residual_m3,
                "holds": abs(diff.identity_residual_m3) <= core.TOL_M3,
                "note": (
                    "This identity is a structural guarantee, not a data check: "
                    "each Vi is defined as a difference of design-referenced cut "
                    "volumes, so the sum telescopes and always closes. It confirms "
                    "the implementation, not the input. The substantive cross-check "
                    "is 'area_consistency'."
                ),
            },
            "region_consistency": {
                "basis": region_basis,
                "regions": {name: box for name, box in regions},
                "consistent": regions_consistent,
            },
            "assumptions": {
                "method_mapping": {
                    name: layer.method
                    for name, layer in zip(
                        [layer.name for layer in diff.layers], diff.layers
                    )
                },
                "below_lowest_stratum": diff.layers[-1].name,
                "below_lowest_is_extrapolated": True,
            },
            "drawing_units": units,
            "warnings": [
                *volume_warnings, *area_warnings, *diff.warnings, *cost_warnings,
            ],
            "elapsed_ms": int((time.perf_counter() - started) * 1000),
        }
        if design_elev is not None:
            result["design_elevation_representative_m"] = design_elev
        if connection_point is not None:
            result["connection_point"] = connection_point
        if area_m2 is not None:
            result["boundary_area_m2"] = area_m2
        return result

    def list_surface_definition(self, surface_name: str) -> dict[str, Any]:
        """List all definition items for a surface.

        Each collection (boundaries, breaklines, etc.) is queried individually
        so a failure on one category does not abort the whole call.
        Uses _iter_com_collection() to avoid silent IEnumVARIANT failures.
        """
        self._ensure_connected()
        surf = self._find_surface(surface_name)
        result: dict[str, Any] = {"surface_name": surface_name, "definitions": {}}

        # 2026-08-16 수정: 원본은 surf.DataDefinition 에서 정의 컬렉션을 찾았으나
        # Civil 3D 2024(AeccXLand 13.6)의 AeccTinSurface 에는 그런 속성이 없다
        # (타입 라이브러리 직접 열람으로 확인). 정의 컬렉션은 **서피스에 직접**
        # 달려 있다. 그래서 이 도구는 언제나 definitions={} 와 오류 문자열만
        # 돌려주고 있었다 — 응답 형식은 멀쩡해서 호출자는 '정의 항목이 없는
        # 서피스'로 읽게 된다. 두 경로를 모두 시도한다.
        sources: list[Any] = []
        try:
            defn = surf.DataDefinition
            if defn is not None:
                sources.append(defn)
        except Exception:  # noqa: BLE001
            pass
        sources.append(surf)
        result["definition_source"] = (
            "DataDefinition" if len(sources) > 1 else "surface object"
        )

        collection_map = [
            ("Boundaries",     "boundaries"),
            ("Breaklines",     "breaklines"),
            ("Contours",       "contours"),
            ("DEMFiles",       "dem_files"),
            ("DrawingObjects", "drawing_objects"),
            ("PointFiles",     "point_files"),
            ("PointGroups",    "point_groups"),
            ("SurveyPoints",   "survey_points"),
            ("SurveyFigures",  "survey_figures"),
        ]

        for attr, label in collection_map:
            try:
                col = None
                for src in sources:
                    try:
                        col = getattr(src, attr, None)
                    except Exception:  # noqa: BLE001
                        col = None
                    if col is not None:
                        break
                if col is None:
                    # 이 버전의 인터페이스에 없는 컬렉션. '0개'가 아니라
                    # '이 버전에 없음'으로 구분해 보고한다.
                    result["definitions"][label] = {"available": False}
                    continue
                raw_items = self._iter_com_collection(col)
                items: list[dict[str, Any]] = []
                for item in raw_items:
                    item_info: dict[str, Any] = {
                        "name": str(getattr(item, "Name", repr(item)))
                    }
                    for prop in ("Description", "Type", "BreaklineType",
                                 "BoundaryType", "FileName", "StyleName"):
                        try:
                            val = getattr(item, prop, None)
                            if val is not None:
                                item_info[prop.lower()] = str(val)
                        except Exception:
                            pass
                    items.append(item_info)
                result["definitions"][label] = {"count": len(items), "items": items}
            except Exception as exc:
                result["definitions"][label] = {"error": str(exc)}

        return result

    
    # ------------------------------------------------------------------
    # Alignments
    # ------------------------------------------------------------------

    @staticmethod
    def _iter_com_collection(col: Any) -> list[Any]:
        """Iterate a Civil 3D COM collection safely.

        Tries multiple strategies in order:
        1. 0-based Item(i) index access
        2. 1-based Item(i) index access
        3. for-loop on the raw collection
        All exceptions are logged so failures are visible in the server log.
        """
        items: list[Any] = []
        try:
            count = int(col.Count)
        except Exception as exc:
            log.debug("_iter_com_collection: cannot read Count: %s", exc)
            count = -1

        if count > 0:
            # Strategy 1: 0-based Item(i)
            zero_items: list[Any] = []
            for i in range(count):
                try:
                    zero_items.append(col.Item(i))
                except Exception as exc:
                    log.debug("_iter_com_collection Item(%d) 0-based failed: %s", i, exc)
                    break
            if len(zero_items) == count:
                return zero_items

            # Strategy 2: 1-based Item(i)
            one_items: list[Any] = []
            for i in range(1, count + 1):
                try:
                    one_items.append(col.Item(i))
                except Exception as exc:
                    log.debug("_iter_com_collection Item(%d) 1-based failed: %s", i, exc)
                    break
            if len(one_items) == count:
                return one_items

            # Use whichever partial set is longer
            best = zero_items if len(zero_items) >= len(one_items) else one_items
            if best:
                log.debug(
                    "_iter_com_collection: returning partial index results (%d/%d)",
                    len(best), count,
                )
                return best

        # Strategy 3: raw for-loop
        try:
            for item in col:
                items.append(item)
        except Exception as exc:
            log.debug("_iter_com_collection raw for-loop failed: %s", exc)

        if not items and count > 0:
            log.warning(
                "_iter_com_collection: Count=%d but all iteration strategies failed.",
                count,
            )
        return items

    def _get_alignments(self) -> list[Any]:
        """Collect all alignments from siteless and site-based collections."""
        results: list[Any] = []

        # Try self._doc first (same source that list_object_types uses
        # successfully for AlignmentsSiteless.Count), then _civil_doc().
        docs_to_try: list[Any] = [self._doc]
        civil_doc = self._civil_doc()
        if civil_doc is not self._doc:
            docs_to_try.append(civil_doc)

        for doc in docs_to_try:
            # Siteless alignments (Civil 3D 2016+)
            siteless = getattr(doc, "AlignmentsSiteless", None)
            if siteless is not None:
                found = self._iter_com_collection(siteless)
                log.debug(
                    "_get_alignments: AlignmentsSiteless from doc %r -> %d items",
                    doc, len(found),
                )
                results.extend(found)

            # Site-based alignments
            sites = getattr(doc, "Sites", None)
            if sites is not None:
                for site in self._iter_com_collection(sites):
                    al_col = getattr(site, "Alignments", None)
                    if al_col is not None:
                        results.extend(self._iter_com_collection(al_col))

            if results:
                break  # found via this document source; stop

        # Last-resort fallback: scan model space for alignment entities
        if not results:
            log.warning(
                "Civil 3D alignment collections empty or inaccessible — "
                "falling back to model space scan."
            )
            try:
                for raw in self._doc.ModelSpace:
                    obj_name = getattr(raw, "ObjectName", "")
                    if "alignment" in obj_name.lower():
                        obj = w32.Dispatch(raw)
                        if (
                            getattr(obj, "Name", None) is not None
                            and getattr(obj, "StartingStation", None) is not None
                        ):
                            results.append(obj)
            except Exception as exc:
                log.warning("Alignment model-space scan failed: %s", exc)

        if not results:
            raise Civil3DError(
                "Alignments collection not found. "
                "Ensure Civil 3D (not plain AutoCAD) is running."
            )
        return results

    def _find_alignment(self, name: str) -> Any:
        for al in self._get_alignments():
            if al.Name.lower() == name.lower():
                return al
        raise Civil3DError(f"Alignment '{name}' not found in the active drawing.")

    def list_alignments(self) -> list[dict[str, Any]]:
        """List all alignments with station range and length."""
        self._ensure_connected()
        results: list[dict[str, Any]] = []
        try:
            for al in self._get_alignments():
                results.append({
                    "name": str(al.Name),
                    "description": str(getattr(al, "Description", "")),
                    "style": str(getattr(al, "StyleName", "")),
                    "object_id": str(al.ObjectID),
                    "length": float(al.Length),
                    "start_station": float(al.StartingStation),
                    "end_station": float(al.EndingStation),
                    "station_index_increment": float(
                        getattr(al, "StationIndexIncrement", 0)
                    ),
                })
        except Civil3DError:
            raise
        except Exception as exc:
            raise Civil3DError(f"list_alignments failed: {exc}") from exc
        return results

    def get_alignment_info(self, alignment_name: str) -> dict[str, Any]:
        """Return geometry breakdown (tangents, curves, spirals) for an alignment."""
        self._ensure_connected()
        al = self._find_alignment(alignment_name)
        info: dict[str, Any] = {
            "name": str(al.Name),
            "description": str(getattr(al, "Description", "")),
            "length": float(al.Length),
            "start_station": float(al.StartingStation),
            "end_station": float(al.EndingStation),
        }
        entities: list[dict[str, Any]] = []
        try:
            ents = al.Entities
            for i in range(ents.Count):
                ent = ents.EntityAt(i)
                ent_info: dict[str, Any] = {
                    "index": i,
                    "type": str(ent.EntityType),
                    "start_station": float(ent.StartStation),
                    "end_station": float(ent.EndStation),
                    "length": float(ent.Length),
                }
                for attr in ("Radius", "TangentLength", "Delta", "Direction"):
                    val = getattr(ent, attr, None)
                    if val is not None:
                        try:
                            ent_info[attr.lower()] = float(val)
                        except (TypeError, ValueError):
                            ent_info[attr.lower()] = str(val)
                entities.append(ent_info)
        except Exception as exc:
            log.warning("Could not read alignment entities: %s", exc)
        info["entities"] = entities
        return info

    def get_station_offset(
        self,
        alignment_name: str,
        easting: float,
        northing: float,
    ) -> dict[str, Any]:
        """
        Project (easting, northing) onto an alignment.
        Returns station and perpendicular offset.
        """
        self._ensure_connected()
        al = self._find_alignment(alignment_name)
        try:
            # StationOffset is a COM method with two by-reference out-parameters.
            # We pass mutable VARIANT objects; Civil 3D writes the values back.
            station_var = self._out_double()
            offset_var = self._out_double()
            al.StationOffset(easting, northing, station_var, offset_var)
            return {
                "alignment_name": alignment_name,
                "easting": easting,
                "northing": northing,
                "station": float(station_var.value),
                "offset": float(offset_var.value),
            }
        except Exception as exc:
            raise Civil3DError(f"get_station_offset failed: {exc}") from exc

    def list_profiles(self, alignment_name: str) -> list[dict[str, Any]]:
        """List all vertical profiles attached to a named alignment."""
        self._ensure_connected()
        al = self._find_alignment(alignment_name)
        results: list[dict[str, Any]] = []
        try:
            profiles = getattr(al, "Profiles", None)
            if profiles is None:
                raise Civil3DError(
                    f"Alignment '{alignment_name}' has no Profiles collection. "
                    "Ensure Civil 3D (not plain AutoCAD) is running."
                )
            for prof in self._iter_com_collection(profiles):
                info: dict[str, Any] = {
                    "name": str(prof.Name),
                    "description": str(getattr(prof, "Description", "")),
                    "style": str(getattr(prof, "StyleName", "")),
                    "object_id": str(getattr(prof, "ObjectID", "")),
                }
                for attr in ("Length", "StartingStation", "EndingStation",
                              "MinElevation", "MaxElevation"):
                    val = getattr(prof, attr, None)
                    if val is not None:
                        try:
                            info[attr[0].lower() + attr[1:]] = float(val)
                        except (TypeError, ValueError):
                            pass
                results.append(info)
        except Civil3DError:
            raise
        except Exception as exc:
            raise Civil3DError(f"list_profiles failed: {exc}") from exc
        return results

    def get_profile_info(self, alignment_name: str, profile_name: str) -> dict[str, Any]:
        """Return full geometry breakdown of a named vertical profile."""
        self._ensure_connected()
        al = self._find_alignment(alignment_name)
        try:
            profiles = getattr(al, "Profiles", None)
            if profiles is None:
                raise Civil3DError(
                    f"Alignment '{alignment_name}' has no Profiles collection."
                )
            prof = None
            for p in self._iter_com_collection(profiles):
                if p.Name.lower() == profile_name.lower():
                    prof = p
                    break
            if prof is None:
                raise Civil3DError(
                    f"Profile '{profile_name}' not found on alignment '{alignment_name}'."
                )
        except Civil3DError:
            raise
        except Exception as exc:
            raise Civil3DError(f"get_profile_info: lookup failed: {exc}") from exc

        info: dict[str, Any] = {
            "alignment_name": alignment_name,
            "name": str(prof.Name),
            "description": str(getattr(prof, "Description", "")),
            "style": str(getattr(prof, "StyleName", "")),
            "object_id": str(getattr(prof, "ObjectID", "")),
        }
        for attr in ("Length", "StartingStation", "EndingStation",
                     "MinElevation", "MaxElevation"):
            val = getattr(prof, attr, None)
            if val is not None:
                try:
                    info[attr[0].lower() + attr[1:]] = float(val)
                except (TypeError, ValueError):
                    pass

        # Entity breakdown (PVIs, tangents, curves)
        entities: list[dict[str, Any]] = []
        try:
            ents = prof.Entities
            count = int(ents.Count)
            for i in range(count):
                try:
                    ent = ents.EntityAt(i)
                except Exception:
                    try:
                        ent = ents.Item(i)
                    except Exception:
                        continue
                ent_info: dict[str, Any] = {
                    "index": i,
                    "type": str(getattr(ent, "EntityType", "Unknown")),
                }
                for attr in ("StartStation", "EndStation", "Length",
                             "StartElevation", "EndElevation",
                             "Radius", "K", "HighLowPtStation", "HighLowPtElevation"):
                    val = getattr(ent, attr, None)
                    if val is not None:
                        try:
                            ent_info[attr[0].lower() + attr[1:]] = float(val)
                        except (TypeError, ValueError):
                            pass
                entities.append(ent_info)
        except Exception as exc:
            log.warning("Could not read profile entities for '%s': %s", profile_name, exc)
        info["entities"] = entities

        # PVI list
        pvis: list[dict[str, Any]] = []
        try:
            pvi_col = prof.PVIs
            for pvi in self._iter_com_collection(pvi_col):
                pvi_info: dict[str, Any] = {}
                for attr in ("Station", "Elevation", "CurveLength"):
                    val = getattr(pvi, attr, None)
                    if val is not None:
                        try:
                            pvi_info[attr[0].lower() + attr[1:]] = float(val)
                        except (TypeError, ValueError):
                            pass
                pvis.append(pvi_info)
        except Exception as exc:
            log.warning("Could not read PVIs for '%s': %s", profile_name, exc)
        info["pvis"] = pvis

        return info
