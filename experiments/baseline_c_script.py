"""baseline_c_script.py — 비교군 C: 고정 스크립트 (제5장 §3)

이것은 **측정 대상**이지 도구가 아니다.

제5장 §3이 규정한 비교군 C의 정의를 그대로 구현한다.

    "C는 D와 동일한 도구 집합을 동일한 순서로 호출하되
     호출 순서와 서피스 이름을 하드코딩한 파이썬 스크립트로 작성한다."

즉 이 스크립트는 **실무자가 한 도면을 보고 작성한 자동화**를 대표한다.
서피스 이름과 호출 순서가 소스에 박혀 있으므로, 그 도면에서는 빠르고
완전히 재현 가능하지만 **다른 도면에서는 동작하지 않는다.**

⚠ 이 스크립트를 "개선"하지 말 것. 이름을 탐색하거나 실패 시 복구하도록
고치면 비교군의 정의가 무너지고 C5 실험이 성립하지 않는다.
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from civil3d_mcp import hydraulics_core as hc                 # noqa: E402
from civil3d_mcp.client import Civil3DClient, Civil3DError    # noqa: E402

# 2026-08-19 추가: 한글 Windows 콘솔은 기본이 cp949 라, 아래 출력에 쓰이는
# em-dash 나 경고 기호 하나 때문에 UnicodeEncodeError 로 스크립트가 즉사한다.
# (setup_check.py 에서 같은 결함을 고쳤으나 실험 스크립트에는 남아 있었다.)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:                                          # noqa: BLE001
        pass

# ---------------------------------------------------------------------------
# 하드코딩 — 환경 1(합성 시험 도면)을 보고 작성한 것
# ---------------------------------------------------------------------------
GROUND = "TEST_C_S0_GROUND"
DESIGN = "TEST_C_DESIGN"
STRATA = [
    {"name": "01토사층",   "surface": "TEST_C_S1_SOIL",      "unit_cost": 4500},
    {"name": "02풍화암층", "surface": "TEST_C_S2_WEATHERED", "unit_cost": 9000},
    {"name": "03연암층",   "surface": "TEST_C_S3_SOFT",      "unit_cost": 18000},
]
BELOW_LOWEST_COST = 25000

# 관로 조건도 하드코딩
FLOW_M3_S = 0.35
LENGTH_M = 4200.0
CANDIDATES = [300, 400, 450, 500, 600, 700, 800]
UNIT_COST = [
    {"diameter_mm": 300, "cost_per_m": 330_000},
    {"diameter_mm": 400, "cost_per_m": 430_000},
    {"diameter_mm": 450, "cost_per_m": 480_000},
    {"diameter_mm": 500, "cost_per_m": 550_000},
    {"diameter_mm": 600, "cost_per_m": 690_000},
    {"diameter_mm": 700, "cost_per_m": 860_000},
    {"diameter_mm": 800, "cost_per_m": 1_050_000},
]
ECON = dict(power_tariff_won_per_kwh=130, pump_efficiency=0.75,
            annual_operating_hours=8760, discount_rate=0.045,
            service_life_years=30)
GROUND_PROFILE = [
    {"station": 0, "elevation": 42.0},
    {"station": 1000, "elevation": 55.0},
    {"station": 2000, "elevation": 78.0},
    {"station": 2500, "elevation": 86.0},
    {"station": 4200, "elevation": 90.0},
]


def run(client: Civil3DClient) -> dict:
    """고정 절차를 그대로 실행한다. 실패하면 그대로 실패한다."""
    started = time.perf_counter()
    trace: list[str] = []

    # 1) 토공
    trace.append("compute_earthwork_by_rock_quality")
    earth = client.compute_earthwork_by_rock_quality(
        ground_surface=GROUND, design_surface=DESIGN, strata=STRATA,
        site_id="baseline_C", below_lowest_unit_cost=BELOW_LOWEST_COST)

    # 2) 경제적 관경
    trace.append("select_economic_diameter")
    econ = hc.select_economic_diameter(
        flow_m3_s=FLOW_M3_S, length_m=LENGTH_M,
        candidate_diameters_mm=CANDIDATES, unit_construction_cost=UNIT_COST,
        site_id="baseline_C", **ECON)

    # 3) 수리종단 — 앞 두 단계의 값을 그대로 받는다
    trace.append("compute_hydraulic_profile")
    prof = hc.hydraulic_profile(
        flow_m3_s=FLOW_M3_S,
        diameter_mm=econ["selected"]["diameter_mm"],
        start_elevation_m=earth["design_elevation_representative_m"],
        start_pump_head_m=50.0, ground_profile=GROUND_PROFILE,
        min_residual_head_m=5.0, pump_head_m=40.0, site_id="baseline_C")

    return {
        "completed": True,
        "trace": trace,
        "total_cut_m3": earth["total_cut_m3"],
        "total_earthwork_cost": earth["total_earthwork_cost"],
        "by_stratum": {s["name"]: s["cut_m3"] for s in earth["by_stratum"]},
        "selected_diameter_mm": econ["selected"]["diameter_mm"],
        "pressurization_points": len(prof["pressurization_points"]),
        "warnings": earth.get("warnings", []) + econ.get("warnings", []),
        "elapsed_s": time.perf_counter() - started,
    }


def run_guarded(client: Civil3DClient) -> dict:
    """실패를 결과로 돌려주기 위한 얇은 껍데기(측정용). 복구는 하지 않는다."""
    started = time.perf_counter()
    try:
        return run(client)
    except Civil3DError as exc:
        return {"completed": False, "error": str(exc).replace("\n", " ")[:200],
                "failed_at": "tool call", "elapsed_s": time.perf_counter() - started}
    except Exception as exc:                                    # noqa: BLE001
        return {"completed": False, "error": f"{type(exc).__name__}: {exc}"[:200],
                "failed_at": "script", "elapsed_s": time.perf_counter() - started}
