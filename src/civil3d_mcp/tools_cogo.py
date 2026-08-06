"""
tools_cogo.py  –  COGO Point MCP tools
"""
from __future__ import annotations

import logging
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from .client import Civil3DClient, Civil3DError

log = logging.getLogger("civil3d_mcp.tools.cogo")


def register(mcp: FastMCP, client: Civil3DClient, run_com: Callable) -> None:

    @mcp.tool(
        name="create_cogo_point",
        description=(
            "Create a new COGO (Coordinate Geometry) point in the active Civil 3D "
            "drawing. Provide northing, easting, optional elevation and description. "
            "Returns the assigned point number and coordinates."
        ),
    )
    async def create_cogo_point(
        northing: float,
        easting: float,
        elevation: float = 0.0,
        description: str = "",
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        northing : float
            Y coordinate (northing) in drawing units.
        easting : float
            X coordinate (easting) in drawing units.
        elevation : float
            Z elevation in drawing units (default 0.0).
        description : str
            Point description / raw description text.
        """
        try:
            return await run_com(
                client.create_cogo_point,
                northing, easting, elevation, description,
            )
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="list_cogo_points",
        description=(
            "List COGO points in the active Civil 3D drawing. "
            "Returns point number, northing, easting, elevation and description "
            "for up to max_count points."
        ),
    )
    async def list_cogo_points(max_count: int = 50) -> dict[str, Any]:
        """
        Parameters
        ----------
        max_count : int
            Maximum number of points to return (default 50).
        """
        try:
            pts = await run_com(client.list_cogo_points, max_count)
            return {"cogo_points": pts, "count": len(pts)}
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="delete_cogo_point",
        description=(
            "Delete a COGO point from the active drawing by its point number. "
            "Returns success confirmation or an error if the point is not found."
        ),
    )
    async def delete_cogo_point(point_number: int) -> dict[str, Any]:
        """
        Parameters
        ----------
        point_number : int
            The Civil 3D point number to delete.
        """
        try:
            await run_com(client.delete_cogo_point, point_number)
            return {"success": True, "deleted_point_number": point_number}
        except Civil3DError as exc:
            return {"error": str(exc)}
