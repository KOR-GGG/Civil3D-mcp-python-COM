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

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    # Civil 3D ProgIDs in order of preference.
    #
    # 2026-08-06 수정: 원본 저장소는 Civil 3D 2026을 "14.4"로 표기했으나,
    # 실제 등록된 ProgID는 "13.8"이다(레지스트리 HKCR 조회로 확인).
    # 이 오류 때문에 Civil 3D 2026에서 GetInterfaceObject가 전부 실패하고
    # 일반 AutoCAD로 폴백되어 Surfaces 컬렉션에 접근할 수 없었다.
    _CIVIL3D_PROG_IDS = [
        "AeccXUiLand.AeccApplication.13.8",  # Civil 3D 2026 — HKCR로 검증
        "AeccXUiLand.AeccApplication.13.7",  # Civil 3D 2025 — HKCR로 검증
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

    
    def list_surface_definition(self, surface_name: str) -> dict[str, Any]:
        """List all definition items for a surface.

        Each collection (boundaries, breaklines, etc.) is queried individually
        so a failure on one category does not abort the whole call.
        Uses _iter_com_collection() to avoid silent IEnumVARIANT failures.
        """
        self._ensure_connected()
        surf = self._find_surface(surface_name)
        result: dict[str, Any] = {"surface_name": surface_name, "definitions": {}}

        # Obtain the DataDefinition object — not all surface types expose it.
        try:
            defn = surf.DataDefinition
        except Exception as exc:
            result["definitions_error"] = (
                f"DataDefinition not available for this surface type: {exc}"
            )
            return result

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
                col = getattr(defn, attr, None)
                if col is None:
                    result["definitions"][label] = {"count": 0, "items": []}
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
