"""test_hydraulics_core.py — 수리·경제관경 검토 논리의 단위 시험.

Civil 3D 가 없어도 돌아간다. 이 도구군은 애초에 COM 에 접근하지 않는다.

기준 사례
--------
도구명세_경제관경수리검토.md 의 검산 예시를 그대로 쓴다.
Q = 0.35 m³/s · L = 4,200 m · D600 · C = 100 일 때
    V  = 4×0.35/(π×0.6²)          = 1.238 m/s
    Δh = 15.43 m                   → 동수경사 0.003674 m/m
    P  = 9.8×0.35×15.43/0.75       = 70.5 kW
    현가계수(i=0.045, n=30)        = 16.289
"""
import math

import pytest

from civil3d_mcp import hydraulics_core as core

Q = 0.35
L = 4200.0

# 명세 예시의 관경별 단가(원/m) — 총건설비를 연장으로 나눈 값
UNIT_COST = [
    {"diameter_mm": 300, "cost_per_m": 330_000},
    {"diameter_mm": 400, "cost_per_m": 430_000},
    {"diameter_mm": 450, "cost_per_m": 480_000},
    {"diameter_mm": 500, "cost_per_m": 550_000},
    {"diameter_mm": 600, "cost_per_m": 690_000},
    {"diameter_mm": 700, "cost_per_m": 860_000},
    {"diameter_mm": 800, "cost_per_m": 1_050_000},
]
CANDIDATES = [300, 400, 450, 500, 600, 700, 800]

ECON = dict(
    power_tariff_won_per_kwh=130,
    pump_efficiency=0.75,
    annual_operating_hours=8760,
    discount_rate=0.045,
    service_life_years=30,
)


# ---------------------------------------------------------------------------
# ① head_loss — 특히 mm -> m 환산
# ---------------------------------------------------------------------------

def test_기준사례가_명세_공식대로_계산된다():
    """반증 대상: Hazen-Williams 구현이 틀렸다.

    기대값은 명세 「방법」 칸의 공식(k=10.666, n=1.85)을 그대로 적용한 값이며,
    2026-08-16 정정 이후의 명세 예시와도 일치한다(Δh 15.42 / 경사 0.003672).
    """
    r = core.head_loss(Q, 600, L)
    assert r["velocity_m_s"] == pytest.approx(1.23787, abs=1e-5)
    assert r["head_loss_m"] == pytest.approx(15.42340, abs=1e-5)
    assert r["hydraulic_gradient"] == pytest.approx(0.00367224, abs=1e-8)
    assert r["c_value_applied"] == 100


def test_계수는_10_666이며_10_67과_구별된다():
    """★ 계수 관례를 코드로 못박아 둔다 (2026-08-16 확정).

    한때 명세의 「방법」 칸(10.666)과 검산 예시(10.67 로 계산됨)가 어긋나 있었다.
    저자 확인 결과 기준 산출물 result.xlsx 가 **10.666** 이므로 「방법」 칸이 옳고
    예시가 틀렸던 것이며, 명세의 예시 수치는 정정되었다.

    이 시험은 나중에 누가 "국제 표준형은 10.67·1.852 인데" 하며 바꾸는 것을
    막는다. 0.04% 는 작아 보이지만 방향이 일정한 계통 오차라, 기준값 대조에서
    도구가 틀린 것으로 오인되게 만든다.
    """
    r = core.head_loss(Q, 600, L)
    ours = r["head_loss_m"]
    other_convention = 10.67 * 100 ** -1.85 * 0.6 ** -4.87 * Q ** 1.85 * L

    assert r["formula_constants"]["coefficient"] == 10.666      # ← 확정값
    assert r["formula_constants"]["flow_exponent"] == 1.85
    assert ours == pytest.approx(15.4234, abs=1e-3)
    assert other_convention == pytest.approx(15.4292, abs=1e-3)
    # 두 관례의 차이는 0.05% 미만이라 눈으로는 구분되지 않는다.
    # 그래서 적용 상수를 결과에 실어 보낸다(감사 가능성).
    assert abs(ours - other_convention) / other_convention < 0.0005


def test_관경은_mm로_받아_m로_환산해_계산한다():
    """반증 대상: mm 를 그대로 식에 넣는다.

    명세가 단위 시험으로 고정하라고 지정한 항목이다. D^-4.87 이므로 1000배를
    놓치면 손실수두가 1000^4.87 ≈ 10^14.6 배로 어긋난다. 그런데 출력 형식은
    멀쩡해서 결과만 보아서는 알 수 없다.
    """
    r = core.head_loss(Q, 600, L)

    # 유속은 면적으로 직접 검산 — mm 로 계산했다면 10^6 배 작아진다
    area_m2 = math.pi * 0.6 ** 2 / 4
    assert r["velocity_m_s"] == pytest.approx(Q / area_m2, rel=1e-12)

    # mm 를 그대로 넣었을 때의 값과 확실히 다르다
    wrong = (
        core.HW_COEFFICIENT * 100 ** core.HW_C_EXP * 600 ** core.HW_D_EXP
        * Q ** core.HW_Q_EXP * L
    )
    assert r["head_loss_m"] / wrong == pytest.approx(1000 ** 4.87, rel=1e-9)


def test_동수경사는_연장에_무관하다():
    """반증 대상: 구간을 나눠 부르면 동수경사가 달라진다.

    Δh ∝ L 이므로 기울기는 L 과 무관해야 한다. ②③이 같은 프리미티브를 공유해도
    수치가 어긋나지 않는 근거다.
    """
    a = core.head_loss(Q, 600, 1.0)["hydraulic_gradient"]
    b = core.head_loss(Q, 600, 99_999.0)["hydraulic_gradient"]
    assert a == pytest.approx(b, rel=1e-12)


@pytest.mark.parametrize("d,expected", [
    (300, 100), (700, 100), (800, 110), (900, 110), (1000, 120), (1500, 120),
])
def test_C값_자동배정이_표_3_2를_따른다(d, expected):
    assert core.auto_c_value(d)[0] == expected


def test_표에_없는_구간은_값을_주되_그_사실을_밝힌다():
    """반증 대상: 표에 없는 관경을 표에 있는 것처럼 조용히 처리한다."""
    value, source = core.auto_c_value(750)
    assert value == 110
    assert "없는 구간" in source


def test_C값이_명시되면_근거도_명시로_바뀐다():
    r = core.head_loss(Q, 600, L, c_value=130)
    assert r["c_value_applied"] == 130
    assert r["c_value_source"] == "explicit"


@pytest.mark.parametrize("c", [79, 151, 0, -100])
def test_허용범위_밖_C값은_거부한다(c):
    """반증 대상: 비현실적 C값을 받아 그럴듯한 수치를 낸다."""
    with pytest.raises(core.HydraulicsError, match="outside the accepted range"):
        core.head_loss(Q, 600, L, c_value=c)


@pytest.mark.parametrize("kwargs,pattern", [
    (dict(diameter_mm=0), "diameter_mm must be positive"),
    (dict(diameter_mm=-600), "diameter_mm must be positive"),
    (dict(length_m=0), "length_m must be positive"),
    (dict(flow_m3_s=-0.1), "flow_m3_s must not be negative"),
])
def test_잘못된_입력은_계산_전에_거부한다(kwargs, pattern):
    args = dict(flow_m3_s=Q, diameter_mm=600, length_m=L)
    args.update(kwargs)
    with pytest.raises(core.HydraulicsError, match=pattern):
        core.head_loss(**args)


# ---------------------------------------------------------------------------
# 현가·동력
# ---------------------------------------------------------------------------

def test_현가계수가_명세값과_일치한다():
    assert core.present_value_factor(0.045, 30) == pytest.approx(16.289, abs=5e-4)


@pytest.mark.parametrize("rate,years", [(0.0, 30), (-0.01, 30), (0.045, 0)])
def test_할인율_내구년도가_0이하면_거부한다(rate, years):
    with pytest.raises(core.HydraulicsError):
        core.present_value_factor(rate, years)


def test_축동력이_명세_검산값과_일치한다():
    assert core.shaft_power_kw(Q, 15.4311, 0.75) == pytest.approx(70.5, abs=0.1)


# ---------------------------------------------------------------------------
# ② select_economic_diameter
# ---------------------------------------------------------------------------

def _select(**over):
    kw = dict(
        flow_m3_s=Q, length_m=L,
        candidate_diameters_mm=CANDIDATES,
        unit_construction_cost=UNIT_COST,
        **ECON,
    )
    kw.update(over)
    return core.select_economic_diameter(**kw)


def test_기준사례에서_D600이_선정된다():
    """반증 대상: 총현가 최소가 아닌 관경을 고른다."""
    r = _select()
    assert r["selected"]["diameter_mm"] == 600
    assert r["selected"]["total_pv"] == min(c["total_pv"] for c in r["candidates"])


def test_유속_상한을_넘는_후보는_사유와_함께_제외된다():
    """반증 대상: 제약 위반 후보를 조용히 버려 왜 빠졌는지 알 수 없게 한다."""
    r = _select()
    excluded = {e["diameter_mm"]: e for e in r["excluded"]}
    assert 300 in excluded
    assert excluded[300]["velocity_m_s"] == pytest.approx(4.951, abs=5e-4)
    assert "above allowed range" in excluded[300]["reason"]
    assert 300 not in [c["diameter_mm"] for c in r["candidates"]]


def test_유속_하한도_적용된다():
    r = _select(velocity_range_m_s=[1.0, 3.0])
    excluded = {e["diameter_mm"]: e for e in r["excluded"]}
    assert 700 in excluded and "below allowed range" in excluded[700]["reason"]
    assert 800 in excluded


def test_단가가_빠진_후보가_있으면_거부한다():
    """반증 대상: 단가 없는 후보를 말없이 건너뛴다.

    건너뛰면 탐색 공간이 조용히 줄어들어, 실제로는 더 싼 관경이 있는데도
    '최소'라고 보고하게 된다.
    """
    with pytest.raises(core.HydraulicsError, match="missing these candidate diameters"):
        _select(unit_construction_cost=[u for u in UNIT_COST
                                        if u["diameter_mm"] != 600])


def test_유속을_만족하는_후보가_없으면_거부한다():
    """반증 대상: 제약을 만족하는 후보가 없는데 아무거나 고른다."""
    with pytest.raises(core.HydraulicsError, match="No candidate diameter satisfies"):
        _select(candidate_diameters_mm=[300], velocity_range_m_s=[0.3, 3.0])


def test_후보_리스트가_비면_거부한다():
    with pytest.raises(core.HydraulicsError, match="candidate_diameters_mm is empty"):
        _select(candidate_diameters_mm=[])


def test_차순위와_근소하면_경고한다():
    """반증 대상: 0.6% 차이를 확정된 결론처럼 보고한다.

    입력 단가가 조금만 변해도 선정이 뒤집히는데, 결과만 보면 알 수 없다.
    """
    r = _select()
    assert r["runner_up_margin"] < 0.05
    assert r["warnings"] and "flip the selection" in r["warnings"][0]


def test_실양정은_같은_노선_안에서_선정을_바꾸지_않는다():
    """반증 대상: 관경에 무관한 상수가 순위를 바꾼다고 오해한다.

    실양정은 모든 후보에 같은 값이 더해지므로 노선 내부의 순위를 바꾸지 못한다.
    노선끼리 비교할 때만 의미가 있다.
    """
    base = _select()
    lifted = _select(static_head_m=80.0)
    assert lifted["selected"]["diameter_mm"] == base["selected"]["diameter_mm"]
    assert lifted["selected"]["total_pv"] > base["selected"]["total_pv"]


def test_자연유하면_동력비가_0이고_그_사실을_경고한다():
    r = _select(pumping_required=False)
    assert all(c["energy_pv"] == 0.0 for c in r["candidates"])
    # 동력비가 0이면 건설비가 가장 싼 후보가 남는다
    assert r["selected"]["diameter_mm"] == min(
        c["diameter_mm"] for c in r["candidates"])
    assert any("gravity-fed" in w for w in r["warnings"])


def test_적용된_가정이_전부_반환된다():
    """반증 대상: 어떤 가정으로 나온 값인지 결과만 보고는 알 수 없다(P5)."""
    a = _select()["assumptions"]
    for key in ("velocity_range_m_s", "discount_rate", "service_life_years",
                "pv_factor", "power_tariff_won_per_kwh", "pump_efficiency",
                "annual_operating_hours", "static_head_m", "minor_losses"):
        assert key in a
    assert a["pv_factor"] == pytest.approx(16.289, abs=5e-4)


# ---------------------------------------------------------------------------
# ③ compute_hydraulic_profile
# ---------------------------------------------------------------------------

PROFILE = [
    {"station": 0, "elevation": 42.0},
    {"station": 1000, "elevation": 55.0},
    {"station": 2000, "elevation": 78.0},
    {"station": 2500, "elevation": 86.0},
    {"station": 4200, "elevation": 90.0},
]


def _profile(**over):
    kw = dict(
        flow_m3_s=Q, diameter_mm=600,
        start_elevation_m=42.0, start_pump_head_m=50.0,
        ground_profile=PROFILE, min_residual_head_m=5.0, pump_head_m=40.0,
    )
    kw.update(over)
    return core.hydraulic_profile(**kw)


def test_명세_예시_종단을_재현한다():
    """반증 대상: 동수경사선 전개나 가압 반영이 틀렸다.

    기대값은 명세 「방법」 칸의 공식(k=10.666)으로 계산한 값이며, 2026-08-16
    정정 이후의 명세 예시 표와 일치한다.
    """
    r = _profile()
    by_station = {s["station"]: s for s in r["stations"]}

    assert by_station[0]["hydraulic_head"] == pytest.approx(92.00, abs=1e-9)
    assert by_station[1000]["residual_head"] == pytest.approx(33.3278, abs=1e-3)
    assert by_station[2000]["residual_head"] == pytest.approx(6.6555, abs=1e-3)
    # 가압 지점 — 전·후를 함께 보고해야 왜 가압했는지가 결과에 남는다
    p = by_station[2500]
    assert p["pressurized"] is True
    assert p["residual_head_before"] == pytest.approx(-3.1806, abs=1e-3)
    assert p["hydraulic_head"] == pytest.approx(122.8194, abs=1e-3)
    assert p["residual_head"] == pytest.approx(36.8194, abs=1e-3)
    assert by_station[4200]["residual_head"] == pytest.approx(26.5766, abs=1e-3)

    # 정정된 명세 예시 표와 소수 둘째 자리까지 일치해야 한다
    spec_table = {0: 50.00, 1000: 33.33, 2000: 6.66, 4200: 26.58}
    for station, expected in spec_table.items():
        assert round(by_station[station]["residual_head"], 2) == expected


def test_가압지점과_요약이_규약대로_집계된다():
    r = _profile()
    assert [p["station"] for p in r["pressurization_points"]] == [2500]
    s = r["summary"]
    # 시점 펌프는 개소 집계에서 빼되 총 양정에는 넣는다
    assert s["intermediate_pump_stations"] == 1
    assert s["total_pump_head_m"] == pytest.approx(90.0)
    assert s["terminal_residual_head_m"] == pytest.approx(26.5766, abs=1e-3)


def test_전_구간이_기준을_만족하면_가압하지_않는다():
    r = _profile(start_pump_head_m=100.0)
    assert r["pressurization_points"] == []
    assert all(s["residual_head"] >= 5.0 for s in r["stations"])


def test_측점이_역순이면_거부한다():
    """반증 대상: 뒤집힌 종단으로 계산해 무의미한 판정을 낸다."""
    bad = [{"station": 2000, "elevation": 78.0}, {"station": 1000, "elevation": 55.0}]
    with pytest.raises(core.HydraulicsError, match="must ascend"):
        _profile(ground_profile=bad)


def test_측점이_중복되면_거부한다():
    bad = [{"station": 0, "elevation": 42.0}, {"station": 0, "elevation": 55.0}]
    with pytest.raises(core.HydraulicsError, match="duplicated station"):
        _profile(ground_profile=bad)


def test_종단이_비면_거부한다():
    with pytest.raises(core.HydraulicsError, match="ground_profile is empty"):
        _profile(ground_profile=[])


def test_가압이_필요한데_양정이_0이면_거부한다():
    """반증 대상: 기준 미달을 무시한 종단을 정상인 것처럼 돌려준다."""
    with pytest.raises(core.HydraulicsError, match="Supply a positive pump_head_m"):
        _profile(pump_head_m=0.0)


def test_양정이_너무_작으면_수렴실패를_보고한다():
    """반증 대상: 무한 반복하거나, 기준 미달인 채로 끝낸다."""
    with pytest.raises(core.HydraulicsError, match="did not converge"):
        _profile(start_pump_head_m=0.0, pump_head_m=0.05)


def test_잔류수두_기준이_음수면_거부한다():
    with pytest.raises(core.HydraulicsError, match="min_residual_head_m"):
        _profile(min_residual_head_m=-1.0)


def test_재샘플링이_판정_해상도를_높인다():
    """반증 대상: 측점이 성기면 그 사이의 기준 미달을 놓친다.

    두 측점 사이에 솟은 지형을 두고, 성긴 종단에서는 위반이 안 잡히지만
    재샘플링하면 잡히는지 본다.
    """
    coarse = [
        {"station": 0, "elevation": 42.0},
        {"station": 1000, "elevation": 44.0},
        {"station": 2000, "elevation": 46.0},
    ]
    r_coarse = core.hydraulic_profile(
        flow_m3_s=Q, diameter_mm=600, start_elevation_m=42.0,
        start_pump_head_m=10.0, ground_profile=coarse,
        min_residual_head_m=5.0, pump_head_m=20.0)
    assert len(r_coarse["stations"]) == 3

    r_fine = core.hydraulic_profile(
        flow_m3_s=Q, diameter_mm=600, start_elevation_m=42.0,
        start_pump_head_m=10.0, ground_profile=coarse,
        min_residual_head_m=5.0, pump_head_m=20.0, station_interval_m=100.0)
    assert len(r_fine["stations"]) == 21          # 0, 100, ... 2000
    # 원 측점은 전부 보존된다
    fine_stations = {s["station"] for s in r_fine["stations"]}
    assert {0.0, 1000.0, 2000.0} <= fine_stations


def test_재샘플링_간격이_0이하면_거부한다():
    with pytest.raises(core.HydraulicsError, match="station_interval_m must be positive"):
        _profile(station_interval_m=0)


def test_한계를_결과에_명시한다():
    """반증 대상: 가압 지점을 가압장 위치로 읽게 둔다."""
    r = _profile()
    assert "not the location" in r["interpretation"]
    assert "does not optimise" in r["interpretation"]
    assert "resolution_note" in r


# ---------------------------------------------------------------------------
# ②③이 ①을 공유한다는 성질
# ---------------------------------------------------------------------------

def test_관경선정과_종단검토가_같은_손실수두를_쓴다():
    """반증 대상: 두 복합 연산이 서로 다른 손실수두로 계산해 결과가 어긋난다.

    프리미티브를 분리해 둔 이유가 이것이다.
    """
    sel = _select()["selected"]
    prof = _profile(diameter_mm=sel["diameter_mm"])
    assert prof["total_head_loss_m"] == pytest.approx(sel["head_loss_m"], rel=1e-12)
    assert prof["hydraulic_gradient"] == pytest.approx(
        sel["hydraulic_gradient"], rel=1e-12)
