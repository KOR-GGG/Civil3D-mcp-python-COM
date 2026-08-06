"""
tools_surfaces.py  –  Surface MCP tools
"""
from __future__ import annotations

import logging
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from .client import Civil3DClient, Civil3DError

log = logging.getLogger("civil3d_mcp.tools.surfaces")


def register(mcp: FastMCP, client: Civil3DClient, run_com: Callable) -> None:

    @mcp.tool(
        name="list_surfaces",
        description=(
            "List all TIN and Grid surfaces in the active Civil 3D drawing. "
            "Returns name, description, style and elevation statistics "
            "(min/max/mean) for each surface."
        ),
    )
    async def list_surfaces() -> dict[str, Any]:
        try:
            surfs = await run_com(client.list_surfaces)
            return {"surfaces": surfs, "count": len(surfs)}
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="get_surface_info",
        description=(
            "Return detailed statistics for a named Civil 3D surface: "
            "elevation range, mean elevation, point count, triangle count, "
            "2D and 3D area."
        ),
    )
    async def get_surface_info(surface_name: str) -> dict[str, Any]:
        """
        Parameters
        ----------
        surface_name : str
            The exact name of the surface as shown in Civil 3D Toolspace.
        """
        try:
            return await run_com(client.get_surface_info, surface_name)
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="sample_surface_elevation",
        description=(
            "Sample the elevation of a named Civil 3D surface at a specific "
            "(easting, northing) coordinate. Returns the interpolated elevation "
            "at that point. Raises an error if the point lies outside the surface."
        ),
    )
    async def sample_surface_elevation(
        surface_name: str,
        easting: float,
        northing: float,
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        surface_name : str
            Name of the surface to query.
        easting : float
            X / easting coordinate in drawing units.
        northing : float
            Y / northing coordinate in drawing units.
        """
        try:
            return await run_com(
                client.sample_surface_elevation,
                surface_name, easting, northing,
            )
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="list_surface_definition",
        description=(
            "List all definition items (boundaries, breaklines, contours, DEM files, "
            "drawing objects, point files, point groups, survey points, survey figures) "
            "that make up a named Civil 3D surface."
        ),
    )
    async def list_surface_definition(surface_name: str) -> dict[str, Any]:
        """
        Parameters
        ----------
        surface_name : str
            The exact name of the surface as shown in Civil 3D Toolspace.
        """
        try:
            return await run_com(client.list_surface_definition, surface_name)
        except Civil3DError as exc:
            return {"error": str(exc)}

    # ------------------------------------------------------------------
    # Future tools to add:
    # - create_surface_from_points
    # - create_surface_from_dem
    # - create_surface_from_tin
    # - add_breakline_to_surface
    # - add_boundary_to_surface
    # - delete_surface
    # ------------------------------------------------------------------

    

