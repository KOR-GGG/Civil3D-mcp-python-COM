"""hydraulics_core.py — 관로 수리·경제관경 검토의 순수 계산 계층.

이 모듈은 **도면에 접근하지 않는다.** 기존 도구군은 예외 없이 Civil3DClient 를
거쳐 COM 을 만지지만, 수리 검토는 폐합형 수식에 기반한 수치 연산이라 CAD 상태가
필요 없다. 따라서 client 계층을 거치지 않으며, 그 결과 단일 STA 스레드
(max_workers=1)라는 동시성 병목에도 묶이지 않는다.

계층을 가르는 기준은 **CAD 접근 여부가 아니라 설계 판단의 개입 여부**다.
그래서 프리미티브(L1)는 도면 조작의 최소 단위일 수도 있고, 여기처럼 도메인의
기초 물리식일 수도 있다.

    head_loss                      L1 프리미티브 — Hazen-Williams 한 식
        │
        ├─────────────────┐
        ▼                 ▼
    select_economic_   hydraulic_profile     L2 도메인 복합 연산
    diameter

②③이 ①을 공유하므로 관경 선정과 종단 검토가 **같은 손실수두 값**을 근거로
삼는다. 두 결과 사이의 수치 불일치가 구조적으로 발생하지 않는다.

단위 규약
--------
길이 m · 유량 m³/s · **관경 mm** · 비용 원.
관경만 mm 인 것은 표준 관경 표기 관행을 따르기 위함이며, **계산 진입 시점에
m 로 정규화한다.** 이 환산이 빠지면 손실수두가 10^14 배 어긋나므로
단위 시험으로 고정해 두었다.
"""
from __future__ import annotations

import math
from typing import Any, Sequence

# Hazen-Williams (SI, 원형관 만관)
#   Δh = 10.666 · C^-1.85 · D^-4.87 · Q^1.85 · L      (D 는 m)
#
# ✅ 2026-08-16 계수 확정 — 10.666.
#
# 명세 안에 불일치가 있었다. 「방법」 칸의 공식은 10.666 인데 검산 예시
# (Δh 15.43 / 동수경사 0.003674)는 10.67 로 계산된 값이었다
# (10.666 -> 15.4234 / 0.003672,  10.67 -> 15.4292 / 0.003674).
#
# 저자 확인 결과 기준 산출물(result.xlsx)이 **10.666** 이므로 「방법」 칸이 옳고
# 예시 쪽이 틀렸다. 구현은 처음부터 「방법」 칸을 따랐으므로 변경하지 않는다.
# 명세의 예시 수치는 10.666 기준으로 정정되었다.
#
# 참고: 국제적으로 널리 쓰이는 SI 형태는 계수 10.67 · 지수 1.852 이고,
# 10.666 · 1.85 는 국내 실무의 변형이다. 이 프로젝트는 후자를 쓴다.
# 적용 상수는 formula_constants 로 항상 반환하므로, 다른 관례로 작성된
# 계산서와 맞댈 때 0.04% 차이의 출처를 즉시 식별할 수 있다.
HW_COEFFICIENT = 10.666
HW_C_EXP = -1.85
HW_D_EXP = -4.87
HW_Q_EXP = 1.85

# C값 허용 범위. 이 밖의 값은 관종·경년 조건을 벗어난 것으로 보고 거부한다.
C_VALUE_MIN, C_VALUE_MAX = 80, 150

# 도·송수관 표준 적용 유속(m/s). 하한은 침전 방지, 상한은 마모·수격작용 억제.
DEFAULT_VELOCITY_RANGE = (0.3, 3.0)

GRAVITY_POWER_COEFF = 9.8      # P[kW] = 9.8·Q[m³/s]·H[m]/η

# 가압 반복 상한. 넘으면 수렴 실패로 보고 관경 재선정을 권고한다.
MAX_PRESSURIZATION_PASSES = 50


class HydraulicsError(ValueError):
    """계산을 계속하면 틀린 값이 나오는 입력. 조용히 넘어가지 않는다."""


# ---------------------------------------------------------------------------
# C값 자동 배정 (제3장 표 3-2)
# ---------------------------------------------------------------------------

def auto_c_value(diameter_mm: float) -> tuple[int, str]:
    """관경에서 C값을 배정하고 **그 근거 문자열을 함께** 돌려준다.

    근거를 함께 반환하는 이유: C값은 손실수두를 C^-1.85 로 좌우하므로
    결과만 보아서는 어떤 가정이 쓰였는지 알 수 없다. 감사 가능성(P5).

    표 3-2 는 D700 이하 / D800~900 / D1,000 이상만 규정하므로 701~799 와
    901~999 는 **표에 없는 구간**이다. 임의로 메우되 그 사실을 근거 문자열에
    남긴다 — 표에 있는 값처럼 보이게 하면 안 된다.
    """
    if diameter_mm <= 700:
        return 100, "auto (D <= 700, 표 3-2)"
    if diameter_mm < 1000:
        if 800 <= diameter_mm <= 900:
            return 110, "auto (D800~900, 표 3-2)"
        return 110, f"auto (D{diameter_mm:g} — 표 3-2에 없는 구간, D800~900 값을 적용)"
    return 120, "auto (D >= 1000, 표 3-2)"


# ---------------------------------------------------------------------------
# ① head_loss — L1 프리미티브
# ---------------------------------------------------------------------------

def head_loss(
    flow_m3_s: float,
    diameter_mm: float,
    length_m: float,
    c_value: int | None = None,
) -> dict[str, Any]:
    """Hazen-Williams 로 단일 구간의 손실수두·동수경사·유속을 낸다.

    전제: 원형관 만관 압력류. **국부손실과 밸브·부속류 손실은 포함하지 않는다**
    (안전측으로 별도 가산할 것).
    """
    if diameter_mm <= 0:
        raise HydraulicsError(f"diameter_mm must be positive; got {diameter_mm}.")
    if length_m <= 0:
        raise HydraulicsError(f"length_m must be positive; got {length_m}.")
    if flow_m3_s < 0:
        raise HydraulicsError(f"flow_m3_s must not be negative; got {flow_m3_s}.")

    if c_value is None:
        c_applied, c_source = auto_c_value(diameter_mm)
    else:
        c_applied, c_source = int(c_value), "explicit"
        if not (C_VALUE_MIN <= c_applied <= C_VALUE_MAX):
            raise HydraulicsError(
                f"c_value {c_applied} is outside the accepted range "
                f"{C_VALUE_MIN}~{C_VALUE_MAX}. Values beyond it do not correspond "
                f"to a real pipe material and ageing condition, and the head loss "
                f"scales with C^-1.85, so an out-of-range C quietly distorts every "
                f"downstream cost."
            )

    d_m = diameter_mm / 1000.0            # ★ mm -> m. 빠지면 10^14 배 어긋난다.
    area = math.pi * d_m ** 2 / 4.0
    velocity = flow_m3_s / area

    dh = (
        HW_COEFFICIENT
        * c_applied ** HW_C_EXP
        * d_m ** HW_D_EXP
        * flow_m3_s ** HW_Q_EXP
        * length_m
    )

    return {
        "head_loss_m": dh,
        "hydraulic_gradient": dh / length_m,
        "velocity_m_s": velocity,
        "c_value_applied": c_applied,
        "c_value_source": c_source,
        "formula": "hazen_williams",
        # 어느 관례로 나온 값인지 결과만 보고도 알 수 있어야 한다. 계수와 지수가
        # 조금만 달라도(10.666 vs 10.67) 0.04% 의 계통 오차가 생기고, 기준값과
        # 대조할 때 그 차이의 출처를 못 찾으면 도구가 틀린 것으로 오인된다.
        "formula_constants": {
            "coefficient": HW_COEFFICIENT,
            "flow_exponent": HW_Q_EXP,
            "diameter_exponent": HW_D_EXP,
            "c_exponent": HW_C_EXP,
        },
        "excludes": "minor losses (fittings, valves) not included",
    }


# ---------------------------------------------------------------------------
# 현가계수
# ---------------------------------------------------------------------------

def present_value_factor(discount_rate: float, years: int) -> float:
    """균등연금의 현가계수 P/A = [(1+i)^n − 1] / [i(1+i)^n]."""
    if discount_rate <= 0:
        raise HydraulicsError(
            f"discount_rate must be positive; got {discount_rate}. "
            f"A zero or negative rate makes the present-value factor meaningless "
            f"and would silently inflate the weight of operating cost."
        )
    if years <= 0:
        raise HydraulicsError(f"service_life_years must be positive; got {years}.")
    growth = (1.0 + discount_rate) ** years
    return (growth - 1.0) / (discount_rate * growth)


def shaft_power_kw(flow_m3_s: float, total_head_m: float, efficiency: float) -> float:
    """축동력 P = 9.8·Q·H/η [kW]."""
    if not (0.0 < efficiency <= 1.0):
        raise HydraulicsError(
            f"pump_efficiency must be in (0, 1]; got {efficiency}."
        )
    return GRAVITY_POWER_COEFF * flow_m3_s * total_head_m / efficiency


# ---------------------------------------------------------------------------
# ② select_economic_diameter — L2
# ---------------------------------------------------------------------------

def _unit_cost_lookup(unit_construction_cost: Sequence[dict]) -> dict[int, float]:
    table: dict[int, float] = {}
    for i, row in enumerate(unit_construction_cost or []):
        if not isinstance(row, dict):
            raise HydraulicsError(
                f"unit_construction_cost[{i}] must be an object like "
                f'{{"diameter_mm": 600, "cost_per_m": 690000}}; got {row!r}.'
            )
        try:
            d = int(row["diameter_mm"])
            c = float(row["cost_per_m"])
        except (KeyError, TypeError, ValueError) as exc:
            raise HydraulicsError(
                f"unit_construction_cost[{i}] needs numeric 'diameter_mm' and "
                f"'cost_per_m'; got {row!r} ({exc})."
            ) from exc
        table[d] = c
    return table


def select_economic_diameter(
    flow_m3_s: float,
    length_m: float,
    candidate_diameters_mm: Sequence[int],
    unit_construction_cost: Sequence[dict],
    power_tariff_won_per_kwh: float,
    pump_efficiency: float,
    annual_operating_hours: float,
    discount_rate: float,
    service_life_years: int,
    velocity_range_m_s: Sequence[float] = DEFAULT_VELOCITY_RANGE,
    static_head_m: float = 0.0,
    pipe_material: str | None = None,
    site_id: str | None = None,
    pumping_required: bool = True,
    c_value: int | None = None,
) -> dict[str, Any]:
    """적정 유속을 만족하는 후보 중 **건설비 + 동력비 현가**가 최소인 관경을 고른다.

    후보별 비교표(`candidates`)를 함께 돌려주는 것이 설계의 요점이다. 선정 결과만
    주면 에이전트가 "왜 이 관경인가"를 설명하지 못하고 검토자가 판단을 다시 볼 수
    없다. 차순위와의 차이가 근소하면 그 사실 자체가 보고되어야 한다.

    실양정(`static_head_m`)은 관경에 무관한 상수이므로 **동일 노선 안에서는 후보 간
    비교 결과를 바꾸지 않는다.** 그러나 후보지끼리 비교할 때는 노선마다 다르므로
    입력받아 합산한다.
    """
    if not candidate_diameters_mm:
        raise HydraulicsError(
            "candidate_diameters_mm is empty. Supply the standard diameters to "
            "evaluate, e.g. [400, 450, 500, 600, 700, 800]."
        )
    if length_m <= 0:
        raise HydraulicsError(f"length_m must be positive; got {length_m}.")
    if flow_m3_s < 0:
        raise HydraulicsError(f"flow_m3_s must not be negative; got {flow_m3_s}.")

    try:
        v_min, v_max = float(velocity_range_m_s[0]), float(velocity_range_m_s[1])
    except (TypeError, IndexError, ValueError) as exc:
        raise HydraulicsError(
            f"velocity_range_m_s must be a two-element range like [0.3, 3.0]; "
            f"got {velocity_range_m_s!r}."
        ) from exc
    if v_min >= v_max:
        raise HydraulicsError(
            f"velocity_range_m_s lower bound must be below the upper bound; "
            f"got [{v_min}, {v_max}]."
        )

    cost_table = _unit_cost_lookup(unit_construction_cost)
    missing = [int(d) for d in candidate_diameters_mm if int(d) not in cost_table]
    if missing:
        raise HydraulicsError(
            f"unit_construction_cost is missing these candidate diameters: {missing}. "
            f"Dropping them would quietly shrink the search space and could return a "
            f"diameter that is not actually the cheapest."
        )

    pv = present_value_factor(discount_rate, service_life_years)
    if pumping_required:
        # 효율은 동력이 실제로 계상될 때만 의미가 있다.
        shaft_power_kw(flow_m3_s, 0.0, pump_efficiency)

    candidates: list[dict[str, Any]] = []
    excluded: list[dict[str, Any]] = []

    for raw_d in candidate_diameters_mm:
        d = int(raw_d)
        hl = head_loss(flow_m3_s, d, length_m, c_value)
        v = hl["velocity_m_s"]

        if v < v_min or v > v_max:
            side = "below" if v < v_min else "above"
            bound = v_min if v < v_min else v_max
            excluded.append({
                "diameter_mm": d,
                "velocity_m_s": v,
                "reason": f"velocity {side} allowed range ({bound} m/s)",
            })
            continue

        construction = cost_table[d] * length_m
        if pumping_required:
            power = shaft_power_kw(
                flow_m3_s, hl["head_loss_m"] + static_head_m, pump_efficiency
            )
            energy = power * annual_operating_hours
            energy_pv = energy * power_tariff_won_per_kwh * pv
        else:
            power = energy = energy_pv = 0.0

        candidates.append({
            "diameter_mm": d,
            "velocity_m_s": v,
            "c_value": hl["c_value_applied"],
            "c_value_source": hl["c_value_source"],
            "head_loss_m": hl["head_loss_m"],
            "hydraulic_gradient": hl["hydraulic_gradient"],
            "shaft_power_kw": power,
            "annual_energy_kwh": energy,
            "construction_cost": construction,
            "energy_pv": energy_pv,
            "total_pv": construction + energy_pv,
        })

    if not candidates:
        raise HydraulicsError(
            f"No candidate diameter satisfies the velocity range "
            f"[{v_min}, {v_max}] m/s at Q = {flow_m3_s} m3/s. "
            f"Re-check the design flow, or widen the candidate list. "
            f"Excluded: {excluded}"
        )

    best = min(candidates, key=lambda c: c["total_pv"])
    best["selected"] = True

    ranked = sorted(candidates, key=lambda c: c["total_pv"])
    runner_up_margin = None
    if len(ranked) > 1 and ranked[0]["total_pv"] > 0:
        runner_up_margin = (
            (ranked[1]["total_pv"] - ranked[0]["total_pv"]) / ranked[0]["total_pv"]
        )

    result: dict[str, Any] = {
        "site_id": site_id,
        "design_flow_m3_s": flow_m3_s,
        "length_m": length_m,
        "pipe_material": pipe_material,
        "pumping_required": pumping_required,
        "selected": {k: v for k, v in best.items() if k != "selected"},
        "candidates": candidates,
        "excluded": excluded,
        "assumptions": {
            "velocity_range_m_s": [v_min, v_max],
            "discount_rate": discount_rate,
            "service_life_years": service_life_years,
            "pv_factor": pv,
            "power_tariff_won_per_kwh": power_tariff_won_per_kwh,
            "pump_efficiency": pump_efficiency,
            "annual_operating_hours": annual_operating_hours,
            "static_head_m": static_head_m,
            "energy_basis": (
                "friction head plus static head"
                if static_head_m else "friction head only"
            ),
            "minor_losses": "not included",
        },
        "warnings": [],
    }

    if runner_up_margin is not None:
        result["runner_up_margin"] = runner_up_margin
        if runner_up_margin < 0.05:
            result["warnings"].append(
                f"D{ranked[0]['diameter_mm']} and D{ranked[1]['diameter_mm']} differ by "
                f"only {runner_up_margin * 100:.2f}% in total present value. A small "
                f"change in unit cost or tariff can flip the selection, so treat this "
                f"as a tie and check the sensitivity before fixing the diameter."
            )
    if not pumping_required:
        result["warnings"].append(
            "pumping_required is false, so the energy cost is zero for every "
            "candidate and the selection reduces to the cheapest construction cost. "
            "Confirm the route really is gravity-fed."
        )
    return result


# ---------------------------------------------------------------------------
# ③ hydraulic_profile — L2
# ---------------------------------------------------------------------------

def _normalise_profile(ground_profile: Sequence[dict]) -> list[tuple[float, float]]:
    if not ground_profile:
        raise HydraulicsError(
            "ground_profile is empty. Supply the ground elevations as "
            '[{"station": 0, "elevation": 42.0}, ...] in ascending station order.'
        )
    pts: list[tuple[float, float]] = []
    for i, row in enumerate(ground_profile):
        if not isinstance(row, dict):
            raise HydraulicsError(
                f"ground_profile[{i}] must be an object with 'station' and "
                f"'elevation'; got {row!r}."
            )
        try:
            st = float(row["station"])
            el = float(row["elevation"])
        except (KeyError, TypeError, ValueError) as exc:
            raise HydraulicsError(
                f"ground_profile[{i}] needs numeric 'station' and 'elevation'; "
                f"got {row!r} ({exc})."
            ) from exc
        pts.append((st, el))

    for i in range(1, len(pts)):
        if pts[i][0] < pts[i - 1][0]:
            raise HydraulicsError(
                f"ground_profile stations must ascend; station {pts[i][0]} at index "
                f"{i} comes after {pts[i - 1][0]}. A reversed profile would place the "
                f"hydraulic grade line against the wrong ground and the pressurisation "
                f"check would be meaningless."
            )
        if pts[i][0] == pts[i - 1][0]:
            raise HydraulicsError(
                f"ground_profile has a duplicated station {pts[i][0]} at index {i}."
            )
    return pts


def _interpolate(pts: Sequence[tuple[float, float]], station: float) -> float:
    if station <= pts[0][0]:
        return pts[0][1]
    if station >= pts[-1][0]:
        return pts[-1][1]
    for i in range(1, len(pts)):
        x0, y0 = pts[i - 1]
        x1, y1 = pts[i]
        if station <= x1:
            span = x1 - x0
            return y0 if span == 0 else y0 + (y1 - y0) * (station - x0) / span
    return pts[-1][1]


def _resample(
    pts: Sequence[tuple[float, float]], interval: float
) -> list[tuple[float, float]]:
    """판정 해상도를 높이기 위해 종단을 균등 간격으로 다시 뽑는다.

    측점 간격이 성기면 그 사이에서 잔류수두가 기준을 밑돌아도 판정에서 빠진다.
    원 측점은 전부 보존하고 균등 격자를 덧붙인다.
    """
    if interval <= 0:
        raise HydraulicsError(
            f"station_interval_m must be positive; got {interval}."
        )
    start, end = pts[0][0], pts[-1][0]
    stations = {p[0] for p in pts}
    n = int(math.floor((end - start) / interval))
    for k in range(n + 1):
        stations.add(start + k * interval)
    stations.add(end)
    return [(s, _interpolate(pts, s)) for s in sorted(stations)]


def hydraulic_profile(
    flow_m3_s: float,
    diameter_mm: float,
    start_elevation_m: float,
    start_pump_head_m: float,
    ground_profile: Sequence[dict],
    min_residual_head_m: float,
    pump_head_m: float = 0.0,
    station_interval_m: float | None = None,
    c_value: int | None = None,
    site_id: str | None = None,
) -> dict[str, Any]:
    """동수경사선을 전개해 잔류수두를 내고 가압 필요구간을 판정한다.

    ⚠ **가압 지점은 위치의 확정이 아니라 검토가 필요한 구간의 지시다.**
    실제 가압장은 부압 구간보다 상류에 두어야 하고 그 위치는 용지·전력 인입·
    유지관리 접근성 등 수리 외적 조건으로 정해진다. 또 이 도구는 기준을 순차적으로
    충족시킬 뿐 **양정을 최적화하지 않는다** — 종점 잔류수두가 과다하게 남는 것은
    그 한계의 표현이다.
    """
    if min_residual_head_m < 0:
        raise HydraulicsError(
            f"min_residual_head_m must not be negative; got {min_residual_head_m}."
        )

    pts = _normalise_profile(ground_profile)
    # `if station_interval_m:` 로 두면 0 이 거짓이라 **재샘플링 요청이 조용히
    # 무시된다.** 잘못된 값은 무시가 아니라 거부여야 한다.
    if station_interval_m is not None:
        pts = _resample(pts, station_interval_m)

    total_length = pts[-1][0] - pts[0][0]
    if total_length <= 0:
        raise HydraulicsError(
            f"ground_profile spans zero length (station {pts[0][0]} to {pts[-1][0]}); "
            f"at least two distinct stations are required."
        )

    hl = head_loss(flow_m3_s, diameter_mm, total_length, c_value)
    gradient = hl["hydraulic_gradient"]

    origin = pts[0][0]
    start_head = start_elevation_m + start_pump_head_m

    # 가압을 만나면 그 지점 이후의 절편이 올라간다. 지점별 수두는
    #   head(x) = start_head - gradient * (x - origin) + (그 지점까지의 누적 가압)
    boosts: list[float] = []          # 가압 지점의 측점
    passes = 0
    while True:
        passes += 1
        if passes > MAX_PRESSURIZATION_PASSES:
            raise HydraulicsError(
                f"Pressurisation did not converge within {MAX_PRESSURIZATION_PASSES} "
                f"passes. pump_head_m ({pump_head_m} m) is too small for this profile, "
                f"or the diameter is too small — the friction gradient "
                f"({gradient:.6f} m/m) outruns each boost. Re-select the diameter or "
                f"raise the pump head."
            )

        violation: float | None = None
        for station, ground in pts:
            applied = pump_head_m * sum(1 for b in boosts if b <= station)
            head = start_head - gradient * (station - origin) + applied
            if head - ground < min_residual_head_m - 1e-9:
                violation = station
                break

        if violation is None:
            break
        if pump_head_m <= 0:
            first_ground = _interpolate(pts, violation)
            applied = pump_head_m * sum(1 for b in boosts if b <= violation)
            head = start_head - gradient * (violation - origin) + applied
            raise HydraulicsError(
                f"Residual head falls to {head - first_ground:.2f} m at station "
                f"{violation:g}, below the required {min_residual_head_m} m, but "
                f"pump_head_m is {pump_head_m}. Supply a positive pump_head_m to let "
                f"the tool place intermediate pumping, or raise start_pump_head_m."
            )
        boosts.append(violation)

    stations: list[dict[str, Any]] = []
    pressurization_points: list[dict[str, Any]] = []
    for station, ground in pts:
        before_applied = pump_head_m * sum(1 for b in boosts if b < station)
        after_applied = pump_head_m * sum(1 for b in boosts if b <= station)
        base = start_head - gradient * (station - origin)
        entry: dict[str, Any] = {
            "station": station,
            "ground_elevation": ground,
            "hydraulic_head": base + after_applied,
            "residual_head": base + after_applied - ground,
            "pressurized": station in boosts,
        }
        if station in boosts:
            # 가압 전·후를 함께 보고한다 — 왜 가압했는지가 결과에 남아야 한다.
            entry["hydraulic_head_before"] = base + before_applied
            entry["residual_head_before"] = base + before_applied - ground
            entry["pump_head_applied"] = pump_head_m
            pressurization_points.append({
                "station": station,
                "reason": (
                    f"residual head {entry['residual_head_before']:.2f} m "
                    f"< {min_residual_head_m} m"
                ),
                "pump_head_m": pump_head_m,
            })
        stations.append(entry)

    return {
        "site_id": site_id,
        "diameter_mm": diameter_mm,
        "flow_m3_s": flow_m3_s,
        "c_value_applied": hl["c_value_applied"],
        "c_value_source": hl["c_value_source"],
        "hydraulic_gradient": gradient,
        "total_head_loss_m": hl["head_loss_m"],
        "min_residual_head_m": min_residual_head_m,
        "stations": stations,
        "pressurization_points": pressurization_points,
        "summary": {
            "start_pump_head_m": start_pump_head_m,
            # 시점 펌프는 계획 조건으로 주어지므로 개소 집계에서 제외하되
            # 총 양정에는 포함한다.
            "intermediate_pump_stations": len(boosts),
            "intermediate_pump_head_m": pump_head_m if boosts else 0.0,
            "total_pump_head_m": start_pump_head_m + pump_head_m * len(boosts),
            "terminal_residual_head_m": stations[-1]["residual_head"],
        },
        "interpretation": (
            "Pressurisation points mark where the residual head first breaches the "
            "limit. They indicate reaches that need a pumping study, not the location "
            "of a pumping station: the real site sits upstream of the low-pressure "
            "reach and is fixed by land, power supply and access. The tool meets the "
            "criterion sequentially and does not optimise the head, so a large "
            "terminal residual head means the start head has slack."
        ),
        "resolution_note": (
            f"Judgement resolution follows the station spacing "
            f"({len(pts)} stations over {total_length:g} m). Pass "
            f"station_interval_m to resample and reduce missed violations."
        ),
    }
