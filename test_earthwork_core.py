"""test_earthwork_core.py — 암질별 토공량 산정 논리의 단위 시험.

Civil 3D 가 없어도 돌아간다. 산정 논리의 핵심인 항등식이 실행 환경에
인질로 잡히지 않게 하려는 것이다.

시험 값의 출처
--------------
2026-08-16 Civil 3D 2024 에서 해석해를 아는 합성 서피스로 실측한 값이다
(make_test_surfaces.py 시험 C). 면적 10,000 ㎡ 의 평면들:

    원지반 S0      EL 100    C(S0) = 12 * 10,000 = 120,000
    토사 하면 S1   EL  95    C(S1) =  7 * 10,000 =  70,000
    풍화암 하면 S2 EL  92    C(S2) =  4 * 10,000 =  40,000
    연암 하면 S3   EL  90    C(S3) =  2 * 10,000 =  20,000
    계획고 P       EL  88

    토사   = 120,000 - 70,000 = 50,000
    풍화암 =  70,000 - 40,000 = 30,000
    연암   =  40,000 - 20,000 = 20,000
    경암   =                    20,000   (최하층 하부, 외삽)
"""
import pytest

from civil3d_mcp import earthwork_core as core

C_GROUND = 120_000.0
NAMES = ["01토사층", "02풍화암층", "03연암층"]
C_STRATA = [70_000.0, 40_000.0, 20_000.0]
EXPECTED = {"01토사층": 50_000.0, "02풍화암층": 30_000.0, "03연암층": 20_000.0}


# ---------------------------------------------------------------------------
# 분류 — 시공방법 판단이 도구 안에 있다는 설계의 시험
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("name,expected", [
    ("01토사층", "토사"),
    ("토사", "토사"),
    ("표토층", "토사"),
    ("02풍화암층", "리핑"),
    ("풍화토", "리핑"),
    ("03연암층", "발파"),
    ("경암", "발파"),
    ("보통암층", "발파"),
    ("경암(최하층 하부)", "발파"),
])
def test_classify_known_strata(name, expected):
    """반증 대상: 지층 이름에서 시공방법을 못 뽑는다."""
    assert core.classify_stratum(name) == expected


def test_풍화암이_발파로_분류되지_않는다():
    """반증 대상: '풍화암'이 '암'을 포함한다는 이유로 발파로 새어 들어간다.

    이것이 새면 단가가 2배인 축에 물량이 통째로 잘못 합산되고,
    결과 표만 보아서는 알 수 없다.
    """
    assert core.classify_stratum("풍화암") == "리핑"
    assert core.classify_stratum("02풍화암층") == "리핑"


def test_분류할_수_없는_이름은_추측하지_않는다():
    """반증 대상: 모르는 이름을 조용히 한쪽으로 밀어 넣는다."""
    with pytest.raises(core.EarthworkError, match="분류할 수 없다"):
        core.classify_stratum("층A")


def test_모호한_이름은_거부한다():
    """반증 대상: 두 암질이 섞인 이름을 임의로 하나로 정한다."""
    with pytest.raises(core.EarthworkError, match="모호"):
        core.classify_stratum("토사및풍화암혼재층")


def test_분류_실패는_모아서_보고한다():
    """반증 대상: 첫 실패에서 멈춰 사용자가 반복 실행하게 만든다."""
    with pytest.raises(core.EarthworkError) as exc:
        core.classify_strata(["층A", "01토사층", "층B"])
    assert "층A" in str(exc.value) and "층B" in str(exc.value)


# ---------------------------------------------------------------------------
# 차분식
# ---------------------------------------------------------------------------

def test_층별_절토량이_해석해와_일치한다():
    """반증 대상: 차분식이 층별 물량을 틀리게 낸다."""
    r = core.differential_volumes(C_GROUND, NAMES, C_STRATA)
    got = {layer.name: layer.cut_m3 for layer in r.layers}
    for name, expected in EXPECTED.items():
        assert got[name] == pytest.approx(expected, abs=core.TOL_M3)
    assert got[core.BELOW_LOWEST_DEFAULT_NAME] == pytest.approx(20_000.0, abs=core.TOL_M3)


def test_항등식이_성립한다():
    """반증 대상: 층별 물량의 합이 총 절토량과 어긋난다.

    Σ Vᵢ + C(Sₙ) = C(S₀). 이 성질 때문에 계획고가 어느 층 내부에서
    끝나도 하부 층이 자동으로 0 이 되고, 합이 총량과 맞음이 구조적으로
    보장된다.
    """
    r = core.differential_volumes(C_GROUND, NAMES, C_STRATA)
    assert r.total_cut_m3 == pytest.approx(C_GROUND, abs=core.TOL_M3)
    assert r.identity_residual_m3 == pytest.approx(0.0, abs=core.TOL_M3)
    assert not r.warnings


def test_계획고가_토사층_안에서_끝나면_하부가_전부_0():
    """반증 대상: 절토가 암층에 닿지 않는데도 암 물량이 나온다.

    실제 검증 도면(2023 연구성과)이 이 경우다 — 리핑·발파 기준값이 0이다.
    """
    # 계획고가 토사층 상부에 있어 어느 암층 경계면도 계획고 위로 올라오지 않는다.
    r = core.differential_volumes(1_480.0, NAMES, [0.0, 0.0, 0.0])
    got = {layer.name: layer.cut_m3 for layer in r.layers}
    assert got["01토사층"] == pytest.approx(1_480.0, abs=core.TOL_M3)
    assert got["02풍화암층"] == 0.0
    assert got["03연암층"] == 0.0
    assert got[core.BELOW_LOWEST_DEFAULT_NAME] == 0.0
    assert r.total_cut_m3 == pytest.approx(1_480.0, abs=core.TOL_M3)


def test_층서가_역전되면_값을_내지_않고_반려한다():
    """반증 대상: 층서가 뒤집혀도 그럴듯한 숫자를 돌려준다.

    층서가 정상이면 Vᵢ ≥ 0 이므로 음의 체적은 그 자체가 오식별의 증거다.
    산정 논리가 검증 논리를 겸한다는 4.3.3 의 주장이 여기서 실행된다.
    """
    바뀐 = [40_000.0, 70_000.0, 20_000.0]     # 토사 하면과 풍화암 하면이 뒤바뀜
    with pytest.raises(core.EarthworkError) as exc:
        core.differential_volumes(C_GROUND, NAMES, 바뀐)
    assert "층서" in str(exc.value)
    assert "02풍화암층" in str(exc.value)


def test_이름과_체적의_개수가_다르면_거부한다():
    with pytest.raises(core.EarthworkError, match="수가 다르다"):
        core.differential_volumes(C_GROUND, NAMES, [70_000.0, 40_000.0])


@pytest.mark.parametrize("c0,cs", [
    (120_000.0, [70_000.0, 40_000.0, 20_000.0]),
    (121_000.0, [70_000.0, 40_000.0, 20_000.0]),
    (1_480.0,   [0.0, 0.0, 0.0]),
    (9_999.5,   [9_999.5, 9_999.5, 0.25]),
])
def test_항등식은_입력과_무관하게_성립한다__검사가_아니라_보장이다(c0, cs):
    """★ 항등식의 성격을 코드로 못박아 둔다.

    Vᵢ = C(Sᵢ₋₁) − C(Sᵢ) 로 **정의**되므로 합은 망원경처럼 접혀
    항상 C(S₀) 가 된다. 즉 이 항등식은 **구조적 보장이지 검사가 아니다.**
    입력을 아무리 비틀어도 잔차는 0 이다.

    이 시험이 잡는 것은 알고리즘의 옳음이 아니라 **구현의 인덱스 실수**다
    (예: boundaries[i+1] - boundaries[i] 로 뒤집어 쓰는 경우).

    알고리즘 자체를 검증하는 것은
    test_층별_절토량이_해석해와_일치한다 이고,
    실제로 일어나는 오류를 잡는 것은 check_area_consistency 다.
    """
    r = core.differential_volumes(c0, NAMES, cs)
    assert r.total_cut_m3 == pytest.approx(c0, abs=core.TOL_M3)
    assert r.identity_residual_m3 == pytest.approx(0.0, abs=core.TOL_M3)


# ---------------------------------------------------------------------------
# 산정 영역 일치 — 항등식이 못 잡는 실제 오류
# ---------------------------------------------------------------------------

BOX = {"min_x": 0.0, "min_y": 0.0, "max_x": 100.0, "max_y": 100.0}
BOX_SMALL = {"min_x": 0.0, "min_y": 0.0, "max_x": 85.0, "max_y": 100.0}


def test_영역이_같으면_경고가_없다():
    assert core.check_bbox_consistency([
        ("원지형", BOX), ("01토사층", BOX), ("02풍화암층", BOX),
    ]) == []


def test_영역이_다르면_경고한다():
    """반증 대상: 외곽 범위가 다른 서피스의 체적을 그대로 빼서 층 물량이라 부른다.

    항등식은 이 경우에도 0 을 보고하므로, 이 검사가 없으면 오류가 드러나지 않는다.
    """
    warns = core.check_bbox_consistency([("원지형", BOX), ("01토사층", BOX_SMALL)])
    assert warns and "01토사층" in warns[0] and "max_x" in warns[0]


def test_영역을_못_읽으면_확인_불가를_밝힌다():
    """반증 대상: 확인하지 못한 것을 확인한 것처럼 조용히 넘어간다."""
    warns = core.check_bbox_consistency([("원지형", BOX), ("01토사층", None)])
    assert warns and "확인할 수 없다" in warns[0]


def test_영역_불일치와_항등식이_함께_있을_때_항등식은_침묵한다():
    """★ 두 검사의 역할 분담을 못박는다.

    같은 시나리오에서 항등식은 통과하고 영역 검사만 걸려야 한다. 이것이
    '항등식은 보장이지 검사가 아니다'의 실증이다.
    """
    r = core.differential_volumes(120_000.0, NAMES, [59_500.0, 34_000.0, 17_000.0])
    assert r.identity_residual_m3 == pytest.approx(0.0, abs=core.TOL_M3)   # 침묵
    assert not r.warnings
    warns = core.check_bbox_consistency([
        ("원지형", BOX), ("01토사층", BOX_SMALL),
        ("02풍화암층", BOX_SMALL), ("03연암층", BOX_SMALL),
    ])
    assert warns                                                          # 이쪽이 잡는다


# ---------------------------------------------------------------------------
# 집계와 공사비
# ---------------------------------------------------------------------------

def _layers_with_costs(costs, below_cost=None):
    r = core.differential_volumes(C_GROUND, NAMES, C_STRATA)
    for layer, c in zip(r.layers, costs):
        layer.unit_cost = c
    r.layers[-1].unit_cost = below_cost
    return r.layers


def test_시공방법_3분류로_집계된다():
    by = core.aggregate_by_method(_layers_with_costs([4_500, 9_000, 18_000], 18_000))
    assert [e["method"] for e in by] == ["토사", "리핑", "발파"]
    soil, rip, blast = by
    assert soil["cut_m3"] == pytest.approx(50_000.0)
    assert rip["cut_m3"] == pytest.approx(30_000.0)
    # 연암 20,000 + 최하층 하부 20,000 이 같은 발파 축으로 합쳐진다
    assert blast["cut_m3"] == pytest.approx(40_000.0)
    assert blast["strata"] == ["03연암층", core.BELOW_LOWEST_DEFAULT_NAME]
    assert blast["assumed"] is True
    assert soil["assumed"] is False


def test_같은_분류에_단가가_다르면_대표단가를_지어내지_않는다():
    """반증 대상: 평균 단가를 만들어 내면 그것이 실제 계약 단가로 읽힌다."""
    by = core.aggregate_by_method(_layers_with_costs([4_500, 9_000, 18_000], 25_000))
    blast = [e for e in by if e["method"] == "발파"][0]
    assert blast["unit_cost"] is None
    assert "unit_cost_note" in blast
    # 그래도 공사비는 지층별 단가로 정확히 합산된다
    assert blast["cost"] == pytest.approx(20_000 * 18_000 + 20_000 * 25_000)


def test_단가가_하나라도_없으면_총액을_내지_않는다():
    """반증 대상: 빠진 항목을 0 으로 두고 합해 총액이 과소 산정된다.

    부지 비교가 목적이므로 이 오류는 '싼 후보지'라는 결론을 직접 만든다.
    """
    by = core.aggregate_by_method(_layers_with_costs([4_500, 9_000, 18_000], None))
    total, warns = core.total_cost(by)
    assert total is None
    assert warns and "과소 산정" in warns[0]


def test_단가가_전부_있으면_총액이_나온다():
    by = core.aggregate_by_method(_layers_with_costs([4_500, 9_000, 18_000], 18_000))
    total, warns = core.total_cost(by)
    assert not warns
    assert total == pytest.approx(
        50_000 * 4_500 + 30_000 * 9_000 + 20_000 * 18_000 + 20_000 * 18_000
    )
