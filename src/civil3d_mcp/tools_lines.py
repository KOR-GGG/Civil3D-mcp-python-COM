"""
tools_lines.py  –  Line & Polyline MCP tools
"""
from __future__ import annotations

import logging
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from .client import Civil3DClient, Civil3DError

log = logging.getLogger("civil3d_mcp.tools.lines")


def register(mcp: FastMCP, client: Civil3DClient, run_com: Callable) -> None:

    @mcp.tool(
        name="create_line",
        description=(
            "Create a 3D line segment in the active Civil 3D drawing model space. "
            "Provide start and end coordinates in (X/easting, Y/northing, Z/elevation). "
            "Returns the object handle, layer and computed length."
        ),
    )
    async def create_line(
        x1: float,
        y1: float,
        z1: float,
        x2: float,
        y2: float,
        z2: float,
        layer: str = "0",
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        x1, y1, z1 : float
            Start point (easting, northing, elevation).
        x2, y2, z2 : float
            End point (easting, northing, elevation).
        layer : str
            Target layer name (default '0').
        """
        try:
            return await run_com(
                client.create_line,
                x1, y1, z1, x2, y2, z2, layer,
            )
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="create_polyline",
        description=(
            "Create a 3D polyline in the active Civil 3D drawing model space "
            "from an ordered list of vertices. Each vertex is [x, y, z]. "
            "Optionally close the polyline. Returns handle, vertex count and length."
        ),
    )
    async def create_polyline(
        vertices: list[list[float]],
        closed: bool = False,
        layer: str = "0",
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        vertices : list of [x, y, z]
            Ordered list of 3D vertices, e.g. [[100.0, 200.0, 10.5], ...].
            Minimum 2 vertices required.
        closed : bool
            Whether to close the polyline back to the first vertex.
        layer : str
            Target layer name (default '0').
        """
        if not vertices or len(vertices) < 2:
            return {"error": "At least 2 vertices are required."}

        # Accept [[x,y,z], ...] or [[x,y], ...] (z defaults to 0)
        verts_3d: list[tuple[float, float, float]] = []
        for v in vertices:
            if len(v) == 2:
                verts_3d.append((float(v[0]), float(v[1]), 0.0))
            elif len(v) >= 3:
                verts_3d.append((float(v[0]), float(v[1]), float(v[2])))
            else:
                return {"error": f"Invalid vertex: {v}. Expected [x, y] or [x, y, z]."}

        try:
            return await run_com(client.create_polyline, verts_3d, closed, layer)
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="list_lines",
        description=(
            "List lines and polylines in the active Civil 3D drawing model space. "
            "Optionally filter by layer name. Returns object type, handle, layer "
            "and length for each entity found."
        ),
    )
    async def list_lines(layer_filter: str = "") -> dict[str, Any]:
        """
        Parameters
        ----------
        layer_filter : str
            If provided, only return lines on this layer (case-insensitive).
            Leave empty to return all lines.
        """
        try:
            lines = await run_com(client.list_lines, layer_filter)
            return {"lines": lines, "count": len(lines)}
        except Civil3DError as exc:
            return {"error": str(exc)}
