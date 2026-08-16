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
            "Without a boundary the computation covers the overlap of the two "
            "surfaces; when the design surface is a corridor datum surface, that "
            "overlap is the graded area plus its slopes. Supply 'boundary' to "
            "clip the computation to a closed polygon instead."
        ),
    )
    async def compute_volume_between_surfaces(
        base_surface: str,
        comparison_surface: str,
        boundary: Any = None,
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
        boundary : str or list, optional
            Closed polygon limiting the computation area. Either the AutoCAD
            entity handle of a closed polyline, or the vertex coordinates as
            [[x, y], ...] or a flat [x, y, z, x, y, z, ...] list. Beware that
            clipping to the site boundary alone omits the earthwork in the
            slopes outside it, which understates the quantities; the boundary
            must cover the construction influence area.
        """
        try:
            return await run_com(
                client.compute_volume_between_surfaces,
                base_surface, comparison_surface, boundary,
            )
        except Civil3DError as exc:
            return {"error": str(exc)}

    @mcp.tool(
        name="compute_earthwork_by_rock_quality",
        description=(
            "Compute cut volume split by rock quality (soil / ripping / blasting) "
            "and the resulting earthwork cost for one candidate site. "
            "Give the existing ground, the design surface (corridor datum surface) "
            "and the stratum BOUNDARY surfaces ordered from top to bottom — the "
            "surface for a soil layer is the BOTTOM of that layer. "
            "Volumes are derived as differences of design-referenced cut volumes, "
            "so the layer volumes always sum to the total; a negative layer volume "
            "means the strata are inverted or a surface was misidentified and the "
            "call is rejected rather than answered. "
            "The mapping from stratum to construction method is decided inside the "
            "tool, not by the caller, and is reported back under 'assumptions'. "
            "Material below the lowest stratum boundary is extrapolated as hard "
            "rock and flagged with assumed=true."
        ),
    )
    async def compute_earthwork_by_rock_quality(
        ground_surface: str,
        design_surface: str,
        strata: list[dict[str, Any]],
        boundary: Any = None,
        site_id: str | None = None,
        below_lowest_unit_cost: float | None = None,
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        ground_surface : str
            Existing ground surface name.
        design_surface : str
            Design surface — normally the corridor datum surface, whose extent
            already covers the graded area plus its slopes.
        strata : list of dict
            Stratum boundary surfaces ordered from top to bottom. Each item is
            ``{"name": str, "surface": str, "unit_cost": float | None}``.
            ``name`` must identify the rock quality (e.g. 토사층 / 풍화암층 /
            연암층) because the tool classifies the construction method from it
            and refuses to guess when the name is unclear or ambiguous.
        boundary : str or list, optional
            Closed polygon limiting every stratum computation. Passing it makes
            all the layer volumes come from the identical region, which removes
            the main way the differencing can go wrong. Either an AutoCAD
            polyline handle or the vertex coordinates.
        site_id : str, optional
            Label echoed back so several candidate sites can be compared.
        below_lowest_unit_cost : float, optional
            Unit cost for the extrapolated material below the lowest stratum
            boundary. Left out on purpose when unknown: the total cost is then
            reported as null instead of being understated.
        """
        try:
            return await run_com(
                client.compute_earthwork_by_rock_quality,
                ground_surface, design_surface, strata, boundary,
                site_id, below_lowest_unit_cost,
            )
        except Civil3DError as exc:
            return {"error": str(exc)}
