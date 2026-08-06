"""
tools_alignments.py  –  Alignment MCP tools
"""
from __future__ import annotations

import logging
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from .client import Civil3DClient, Civil3DError

log = logging.getLogger("civil3d_mcp.tools.alignments")


def register(mcp: FastMCP, client: Civil3DClient, run_com: Callable) -> None:

    @mcp.tool(
        name="list_alignments",
        description=(
            "List all alignments in the active Civil 3D drawing. "
            "Returns name, description, style, length, starting station "
            "and ending station for each alignment."
        ),
    )
    async def list_alignments() -> dict[str, Any]:
        try:
            aligns = await run_com(client.list_alignments)
            return {"alignments": aligns, "count": len(aligns)}
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="get_alignment_info",
        description=(
            "Return detailed geometry for a named Civil 3D alignment: "
            "length, start/end stations and a breakdown of every entity "
            "(tangent, circular curve, spiral) with type, station range, "
            "length and radius where applicable."
        ),
    )
    async def get_alignment_info(alignment_name: str) -> dict[str, Any]:
        """
        Parameters
        ----------
        alignment_name : str
            Exact alignment name as shown in Civil 3D Toolspace.
        """
        try:
            return await run_com(client.get_alignment_info, alignment_name)
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="get_station_offset",
        description=(
            "Given a world coordinate (easting, northing), compute the "
            "station and perpendicular offset relative to a named alignment. "
            "Useful for checking whether a point is on or near a road centreline."
        ),
    )
    async def get_station_offset(
        alignment_name: str,
        easting: float,
        northing: float,
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        alignment_name : str
            Name of the alignment to project onto.
        easting : float
            X / easting of the point to project.
        northing : float
            Y / northing of the point to project.
        """
        try:
            return await run_com(
                client.get_station_offset,
                alignment_name, easting, northing,
            )
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="list_profiles",
        description=(
            "List all vertical profiles (AeccDbVAlignment) attached to a "
            "named Civil 3D alignment. Returns name, description, style, "
            "length, start/end stations and min/max elevations for each profile."
        ),
    )
    async def list_profiles(alignment_name: str) -> dict[str, Any]:
        """
        Parameters
        ----------
        alignment_name : str
            Exact alignment name as shown in Civil 3D Toolspace.
        """
        try:
            profiles = await run_com(client.list_profiles, alignment_name)
            return {"alignment_name": alignment_name, "profiles": profiles, "count": len(profiles)}
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="get_profile_info",
        description=(
            "Return detailed geometry for a named vertical profile on a Civil 3D "
            "alignment: start/end stations, min/max elevations, full entity breakdown "
            "(tangent grades, sag/crest curves) with station range, elevations, radius "
            "and K-value, plus a complete PVI list with station, elevation and curve length."
        ),
    )
    async def get_profile_info(
        alignment_name: str,
        profile_name: str,
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        alignment_name : str
            Exact alignment name as shown in Civil 3D Toolspace.
        profile_name : str
            Exact profile name as shown under the alignment in Toolspace.
        """
        try:
            return await run_com(client.get_profile_info, alignment_name, profile_name)
        except Civil3DError as exc:
            return {"error": str(exc)}
