"""
tools_earthwork.py  –  Earthwork / volume MCP tools

Layer 1 (primitive) of the two-layer tool architecture:
``compute_volume_between_surfaces`` supplies the cut/fill volume that the
base server does not provide.  The domain-level operation
``compute_earthwork_by_rock_quality`` builds on it and is added separately.
"""
from __future__ import annotations

import logging
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from .client import Civil3DClient, Civil3DError

log = logging.getLogger("civil3d_mcp.tools.earthwork")


def register(mcp: FastMCP, client: Civil3DClient, run_com: Callable) -> None:

    @mcp.tool(
        name="compute_volume_between_surfaces",
        description=(
            "Compute cut, fill and net volume between two Civil 3D TIN surfaces. "
            "Creates a temporary TIN volume surface, reads its statistics, then "
            "erases it so the drawing is left unchanged. "
            "The volume surface represents 'comparison - base': pass the existing "
            "ground (or a stratum boundary) as base_surface and the design surface "
            "as comparison_surface, and cut_m3 is the material to be removed. "
            "The computation covers the overlap of the two surfaces; when the "
            "design surface is a corridor datum surface, that overlap is the "
            "graded area plus its slopes."
        ),
    )
    async def compute_volume_between_surfaces(
        base_surface: str,
        comparison_surface: str,
        boundary: str | None = None,
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        base_surface : str
            Name of the reference surface — existing ground or a stratum
            boundary surface. Must exist and be built.
        comparison_surface : str
            Name of the surface compared against the base — normally the
            design surface (corridor datum surface).
        boundary : str, optional
            Name of a closed polyline limiting the computation area. Explicit
            boundary clipping is not implemented yet; supplying this argument
            returns an error instead of a silently unbounded result.
        """
        try:
            return await run_com(
                client.compute_volume_between_surfaces,
                base_surface, comparison_surface, boundary,
            )
        except Civil3DError as exc:
            return {"error": str(exc)}
