"""
tools_drawing.py  –  Drawing info & object type MCP tools
"""
from __future__ import annotations

import logging
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from .client import Civil3DClient, Civil3DError

log = logging.getLogger("civil3d_mcp.tools.drawing")


def register(mcp: FastMCP, client: Civil3DClient, run_com: Callable) -> None:

    @mcp.tool(
        name="get_drawing_info",
        description=(
            "Retrieve metadata about the currently active Civil 3D drawing: "
            "file name, full path, save state, unit settings, precision and "
            "coordinate system code."
        ),
    )
    async def get_drawing_info() -> dict[str, Any]:
        try:
            return await run_com(client.get_drawing_info)
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="list_civil_object_types",
        description=(
            "List and count Civil 3D objects in the active drawing. "
            "Returns two sections: model_space_counts (AutoCAD entity counts by "
            "object name) and civil3d_collections (named collection counts for "
            "surfaces, alignments, COGO points, corridors, pipe networks, etc.)."
        ),
    )
    async def list_civil_object_types() -> dict[str, Any]:
        try:
            return await run_com(client.list_object_types)
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="get_selected_objects_info",
        description=(
            "Return properties of the currently selected Civil 3D objects "
            "(handle, layer, object type, name/description where available). "
            "Use max_count to limit results."
        ),
    )
    async def get_selected_objects_info(max_count: int = 10) -> dict[str, Any]:
        """
        Parameters
        ----------
        max_count : int
            Maximum number of selected objects to return (default 10).
        """
        try:
            objs = await run_com(client.get_selected_objects_info, max_count)
            return {"selected_objects": objs, "count": len(objs)}
        except Civil3DError as exc:
            return {"error": str(exc)}
