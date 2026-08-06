"""
tools_corridors.py  –  Corridor MCP tools
"""
from __future__ import annotations

import logging
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from .client import Civil3DClient, Civil3DError

log = logging.getLogger("civil3d_mcp.tools.corridors")


def register(mcp: FastMCP, client: Civil3DClient, run_com: Callable) -> None:

    @mcp.tool(
        name="get_corridor_info",
        description=(
            "Return detailed information for a named Civil 3D corridor: "
            "description, style, baselines (with their alignment and profile names), "
            "regions per baseline, and the list of assemblies applied to each region."
        ),
    )
    async def get_corridor_info(corridor_name: str) -> dict[str, Any]:
        """
        Parameters
        ----------
        corridor_name : str
            Exact corridor name as shown in Civil 3D Toolspace.
        """
        try:
            return await run_com(client.get_corridor_info, corridor_name)
        except Civil3DError as exc:
            return {"error": str(exc)}
