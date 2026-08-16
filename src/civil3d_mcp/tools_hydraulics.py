"""
tools_hydraulics.py  –  Pipeline hydraulics / economic diameter MCP tools

★ 이 도구군은 기반 서버에서 **client 계층을 거치지 않는 첫 사례**다.

세 도구 모두 COM 에 접근하지 않으므로 Civil3DClient 에 메서드를 두지 않고,
순수 계산은 hydraulics_core 에 둔다. "client 는 COM 브리지 전용"이라는 기존
계층 규약을 지키기 위한 것이다.

부수 효과로 이 도구들은 run_com 의 단일 STA executor(max_workers=1)를 거치지
않으므로 **동시성 상한 1의 병목에 묶이지 않는다.** 다만 이벤트 루프를 막지
않도록 후보 관경 수와 측점 수에 상한을 둔다.

오류 관례: 기존 도구는 Civil3DError 만 포착하지만 여기서는 도메인 검증 오류가
주를 이룬다. 예외를 FastMCP 까지 올리지 않고 검증 단계에서 {"error": ...} 로
돌려준다(기존 create_polyline 의 사전 검증 패턴).
"""
from __future__ import annotations

import logging
import time
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from . import hydraulics_core as core
from .client import Civil3DClient

log = logging.getLogger("civil3d_mcp.tools.hydraulics")

# 이벤트 루프를 블로킹하지 않기 위한 상한. 넘으면 계산하지 않고 거부한다 —
# 조용히 잘라내면 후보를 빠뜨린 채 '최적'이라고 보고하게 된다.
MAX_CANDIDATES = 200
MAX_PROFILE_POINTS = 20_000


def register(mcp: FastMCP, client: Civil3DClient, run_com: Callable) -> None:
    # client·run_com 은 등록 규약을 맞추기 위해 받되 쓰지 않는다(위 주석 참조).

    @mcp.tool(
        name="compute_head_loss",
        description=(
            "Head loss, hydraulic gradient and velocity for one pipe reach using the "
            "Hazen-Williams formula. Pure calculation — it does not read the drawing. "
            "Diameter is given in millimetres and converted internally; flow is m3/s "
            "and length is metres. When c_value is omitted it is assigned from the "
            "diameter and the basis is reported back in c_value_source. Minor losses "
            "from fittings and valves are NOT included."
        ),
    )
    async def compute_head_loss(
        flow_m3_s: float,
        diameter_mm: float,
        length_m: float,
        c_value: int | None = None,
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        flow_m3_s : float
            Design flow in cubic metres per second.
        diameter_mm : float
            Internal diameter in millimetres (standard pipe notation).
        length_m : float
            Reach length in metres.
        c_value : int, optional
            Hazen-Williams roughness coefficient. Omit to have it assigned from
            the diameter; the applied value and its basis are always returned.
        """
        started = time.perf_counter()
        try:
            result = core.head_loss(flow_m3_s, diameter_mm, length_m, c_value)
        except core.HydraulicsError as exc:
            return {"error": str(exc)}
        result["elapsed_ms"] = int((time.perf_counter() - started) * 1000)
        return result

    @mcp.tool(
        name="select_economic_diameter",
        description=(
            "Choose the pipe diameter with the lowest life-cycle cost — construction "
            "cost plus the present value of pumping energy — among the candidates that "
            "satisfy the allowed velocity range. Returns the full candidate comparison "
            "and the diameters excluded by velocity with their reason, so the choice "
            "can be audited and a near-tie with the runner-up is visible. Pure "
            "calculation; it does not read the drawing."
        ),
    )
    async def select_economic_diameter(
        flow_m3_s: float,
        length_m: float,
        candidate_diameters_mm: list[int],
        unit_construction_cost: list[dict[str, float]],
        power_tariff_won_per_kwh: float,
        pump_efficiency: float,
        annual_operating_hours: float,
        discount_rate: float,
        service_life_years: int,
        velocity_range_m_s: list[float] | None = None,
        static_head_m: float = 0.0,
        pipe_material: str | None = None,
        site_id: str | None = None,
        pumping_required: bool = True,
        c_value: int | None = None,
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        candidate_diameters_mm : list of int
            Standard diameters to evaluate, e.g. [400, 450, 500, 600, 700, 800].
        unit_construction_cost : list of dict
            Cost per metre for every candidate, as
            ``[{"diameter_mm": 600, "cost_per_m": 690000}, ...]``. A candidate
            without a cost is rejected rather than skipped, because skipping it
            would shrink the search space without saying so.
        velocity_range_m_s : list of float, optional
            Allowed velocity range, default [0.3, 3.0] — the lower bound keeps
            sediment moving, the upper bound limits wear and water hammer.
        static_head_m : float
            Static lift. Constant across diameters, so it does not change the
            ranking within one route, but it does matter when comparing routes.
        pumping_required : bool
            Set false for a gravity-fed route; the energy cost then becomes zero
            for every candidate and the choice reduces to construction cost.
        """
        started = time.perf_counter()
        if len(candidate_diameters_mm or []) > MAX_CANDIDATES:
            return {"error": (
                f"candidate_diameters_mm has {len(candidate_diameters_mm)} entries, "
                f"above the limit of {MAX_CANDIDATES}. Narrow the list to the "
                f"standard diameters actually under consideration."
            )}
        try:
            result = core.select_economic_diameter(
                flow_m3_s=flow_m3_s,
                length_m=length_m,
                candidate_diameters_mm=candidate_diameters_mm,
                unit_construction_cost=unit_construction_cost,
                power_tariff_won_per_kwh=power_tariff_won_per_kwh,
                pump_efficiency=pump_efficiency,
                annual_operating_hours=annual_operating_hours,
                discount_rate=discount_rate,
                service_life_years=service_life_years,
                velocity_range_m_s=(
                    velocity_range_m_s or list(core.DEFAULT_VELOCITY_RANGE)
                ),
                static_head_m=static_head_m,
                pipe_material=pipe_material,
                site_id=site_id,
                pumping_required=pumping_required,
                c_value=c_value,
            )
        except core.HydraulicsError as exc:
            return {"error": str(exc)}
        result["elapsed_ms"] = int((time.perf_counter() - started) * 1000)
        return result

    @mcp.tool(
        name="compute_hydraulic_profile",
        description=(
            "Develop the hydraulic grade line along a ground profile, report the "
            "residual head at every station and find where pressurisation is needed. "
            "When the residual head drops below the limit, pump_head_m is added at "
            "that station and the reach downstream is recomputed, repeating until the "
            "whole line clears. The reported points indicate reaches that need a "
            "pumping study — they are NOT pumping station locations, and the tool does "
            "not optimise the head. Pure calculation; it does not read the drawing."
        ),
    )
    async def compute_hydraulic_profile(
        flow_m3_s: float,
        diameter_mm: float,
        start_elevation_m: float,
        start_pump_head_m: float,
        ground_profile: list[dict[str, float]],
        min_residual_head_m: float,
        pump_head_m: float = 0.0,
        station_interval_m: float | None = None,
        c_value: int | None = None,
        site_id: str | None = None,
    ) -> dict[str, Any]:
        """
        Parameters
        ----------
        ground_profile : list of dict
            Ground elevations in ascending station order, as
            ``[{"station": 0, "elevation": 42.0}, ...]``. Reversed or duplicated
            stations are rejected.
        start_elevation_m, start_pump_head_m : float
            The head at the start is the sum of the two.
        min_residual_head_m : float
            Required residual head; stations below it trigger pressurisation.
        pump_head_m : float
            Head added by one intermediate pumping step. Leave at zero only when
            no pressurisation is expected — if the line then breaches the limit
            the call returns an error instead of a profile that ignores it.
        station_interval_m : float, optional
            Resample the profile at this spacing before judging. Judgement
            resolution follows the station spacing, so a coarse profile can miss
            a breach between two given points.
        """
        started = time.perf_counter()
        if len(ground_profile or []) > MAX_PROFILE_POINTS:
            return {"error": (
                f"ground_profile has {len(ground_profile)} points, above the limit of "
                f"{MAX_PROFILE_POINTS}. Thin the profile before calling."
            )}
        try:
            result = core.hydraulic_profile(
                flow_m3_s=flow_m3_s,
                diameter_mm=diameter_mm,
                start_elevation_m=start_elevation_m,
                start_pump_head_m=start_pump_head_m,
                ground_profile=ground_profile,
                min_residual_head_m=min_residual_head_m,
                pump_head_m=pump_head_m,
                station_interval_m=station_interval_m,
                c_value=c_value,
                site_id=site_id,
            )
        except core.HydraulicsError as exc:
            return {"error": str(exc)}
        if len(result["stations"]) > MAX_PROFILE_POINTS:
            return {"error": (
                f"Resampling produced {len(result['stations'])} stations, above the "
                f"limit of {MAX_PROFILE_POINTS}. Increase station_interval_m."
            )}
        result["elapsed_ms"] = int((time.perf_counter() - started) * 1000)
        return result
