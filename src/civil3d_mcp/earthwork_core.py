"""earthwork_core.py — 암질별 토공량 산정의 순수 계산 계층.

COM에 닿지 않는 계산·판정을 전부 이 파일로 분리했다. 이유는 둘이다.

1. **Civil 3D 없이 단위 시험이 돈다.** 차분식의 항등식
   Σ Vᵢ + C(Sₙ) = C(S₀) 은 산정 논리의 핵심인데, 이것을 확인하는 데
   CAD가 떠 있어야 한다면 시험이 환경에 인질로 잡힌다.

2. **'도메인 판단'과 'CAD 조작'의 경계가 파일 경계와 일치한다.**
   지층을 어느 시공방법으로 볼 것인가는 공학 판단이고,
   서피스에서 체적을 읽는 것은 CAD 조작이다. 전자는 여기, 후자는
   client.py 에 둔다.

용어
----
P      계획고(설계 서피스). 코리더 데이텀 서피스를 전제한다.
S₀     원지반
S₁…Sₙ  상→하로 정렬된 암층 경계면
C(S)   ∫_A max(S − P, 0) dA — 계획고 기준 절토체적.
       compute_volume_between_surfaces(base=S, comparison=P) 의 cut_m3 이다.

차분식
------
       Vᵢ = C(Sᵢ₋₁) − C(Sᵢ)        (제i 지층의 절토량)
       최하층 하부 = C(Sₙ)
       Σ Vᵢ + C(Sₙ) = C(S₀)        (항등식 — 구조적으로 성립)

인접 경계면 사이의 체적을 그대로 쓰는 '경계쌍' 방식은 *지층의 전체 체적*이지
*절토된 부분*이 아니어서 크게 과대 산정된다. 검증 도면 수치로 약 14배.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, Sequence

# 체적 비교에 쓰는 허용 오차(㎥). Civil 3D 가 돌려주는 배정밀도 값의
# 반올림 잔차를 흡수하되, 의미 있는 물량 차이는 잡을 수 있는 크기.
TOL_M3 = 1e-6


class EarthworkError(ValueError):
    """산정을 계속하면 틀린 값이 나오는 상황. 조용히 넘어가지 않는다."""


# ---------------------------------------------------------------------------
# 지층 → 시공방법 3분류
#
# 이 표가 '도구 안에' 있다는 것이 설계의 요점이다. 어느 지층이 발파 대상인가는
# 시공방법 판단이며, 이를 호출자(에이전트)의 인자로 받으면 공학 판단이
# LLM 추론에 위임된다. 집계 축을 토사/리핑/발파로 둔 것은 표준품셈과
# 기존 실무 산출물(result.xlsx)이 이 3분류로 집계하기 때문이며, 그래야
# 제5장에서 기준값과 직접 대조된다.
#
# 순서가 의미를 갖는다 — '풍화암'은 '연암·경암'보다 먼저 걸려야 한다.
# ---------------------------------------------------------------------------
METHOD_ORDER: tuple[str, ...] = ("토사", "리핑", "발파")

METHOD_KEYWORDS: dict[str, tuple[str, ...]] = {
    "토사": ("토사", "표토", "매립", "사질토", "점성토", "붕적", "퇴적층",
             "soil", "topsoil", "fill"),
    "리핑": ("풍화암", "풍화토", "리핑", "ripping", "weathered"),
    "발파": ("연암", "보통암", "경암", "발파", "기반암",
             "blast", "soft rock", "hard rock", "bedrock"),
}

# 최하층 경계면 하부의 기본 가정. 시추 자료는 통상 연암 상단까지만
# 제공되므로 그 아래는 경암으로 가정하며, 이는 지반조사 심도를 넘어서는
# 외삽이다. 따라서 assumed 플래그로 반드시 구분해 보고한다.
BELOW_LOWEST_DEFAULT_NAME = "경암(최하층 하부)"


def classify_stratum(name: str) -> str:
    """지층 이름 하나를 시공방법으로 분류한다.

    분류할 수 없거나 둘 이상에 걸리면 **추측하지 않고 예외를 낸다.**
    조용히 한쪽으로 밀어 넣으면 단가가 몇 배 다른 물량이 잘못된 축에
    합산되고, 결과만 보아서는 그 사실을 알 수 없다.
    """
    lowered = name.lower()
    hits = [
        method
        for method in METHOD_ORDER
        if any(kw.lower() in lowered for kw in METHOD_KEYWORDS[method])
    ]
    if not hits:
        raise EarthworkError(
            f"지층 '{name}'을 시공방법(토사/리핑/발파)으로 분류할 수 없다. "
            f"이름에 다음 중 하나가 포함되어야 한다: "
            f"{', '.join(k for m in METHOD_ORDER for k in METHOD_KEYWORDS[m][:4])}. "
            f"분류를 추측하면 단가가 몇 배 다른 물량이 잘못된 축에 합산된다."
        )
    if len(hits) > 1:
        raise EarthworkError(
            f"지층 '{name}'이 시공방법 {hits} 에 동시에 해당해 분류가 모호하다. "
            f"지층 이름을 하나의 암질만 가리키도록 정리할 것."
        )
    return hits[0]


def classify_strata(names: Sequence[str]) -> list[str]:
    """여러 지층을 한 번에 분류하고, 실패는 **모아서** 보고한다.

    첫 실패에서 멈추면 사용자가 이름을 하나 고칠 때마다 다시 돌려야 한다.
    """
    methods: list[str] = []
    problems: list[str] = []
    for name in names:
        try:
            methods.append(classify_stratum(name))
        except EarthworkError as exc:
            problems.append(str(exc))
    if problems:
        raise EarthworkError(
            f"지층 {len(problems)}건을 분류하지 못했다.\n  - "
            + "\n  - ".join(problems)
        )
    return methods


# ---------------------------------------------------------------------------
# 차분식
# ---------------------------------------------------------------------------

@dataclass
class LayerVolume:
    """한 지층(또는 최하층 하부)의 절토량."""
    name: str
    method: str
    cut_m3: float
    assumed: bool = False
    unit_cost: float | None = None

    @property
    def cost(self) -> float | None:
        if self.unit_cost is None:
            return None
        return self.cut_m3 * self.unit_cost


@dataclass
class DifferentialResult:
    layers: list[LayerVolume]
    total_cut_m3: float
    identity_residual_m3: float
    warnings: list[str] = field(default_factory=list)


def differential_volumes(
    c_ground: float,
    strata_names: Sequence[str],
    c_strata: Sequence[float],
    *,
    below_lowest_name: str = BELOW_LOWEST_DEFAULT_NAME,
    tol: float = TOL_M3,
) -> DifferentialResult:
    """차분식으로 층별 절토량을 낸다.

    Parameters
    ----------
    c_ground : C(S₀) — 원지반을 계획고로 자른 절토체적
    strata_names : 상→하로 정렬된 암층 이름
    c_strata : 같은 순서의 C(Sᵢ)

    Raises
    ------
    EarthworkError
        층별 체적이 음수인 경우. 층서가 정상이면 Sᵢ₋₁ ≥ Sᵢ 이므로 Vᵢ ≥ 0 이다.
        즉 **음의 체적은 그 자체가 층서 역전 또는 서피스 오식별의 증거**이며,
        산정 논리가 검증 논리를 겸한다. 이때 값을 내면 틀린 근거로 후보지를
        비교하게 되므로 반려한다.
    """
    if len(strata_names) != len(c_strata):
        raise EarthworkError(
            f"지층 이름 {len(strata_names)}개와 체적 {len(c_strata)}개의 수가 다르다."
        )

    methods = classify_strata(strata_names)
    boundaries = [c_ground, *c_strata]

    layers: list[LayerVolume] = []
    violations: list[str] = []
    for i, (name, method) in enumerate(zip(strata_names, methods)):
        v = boundaries[i] - boundaries[i + 1]
        if v < -tol:
            upper = "원지반" if i == 0 else strata_names[i - 1]
            violations.append(
                f"'{name}' 층의 절토량이 음수({v:,.6f} ㎥)다. "
                f"C({upper})={boundaries[i]:,.6f} < C({name})={boundaries[i + 1]:,.6f} — "
                f"'{upper}'이 '{name}'보다 아래에 있다는 뜻이므로 층서가 역전되었거나 "
                f"서피스가 잘못 지정되었다."
            )
        layers.append(LayerVolume(name=name, method=method, cut_m3=max(v, 0.0)))

    if violations:
        raise EarthworkError(
            "층서 검산에 실패했다. 산정을 중단한다.\n  - " + "\n  - ".join(violations)
        )

    below = boundaries[-1]
    if below < -tol:
        raise EarthworkError(
            f"최하층 하부 물량이 음수({below:,.6f} ㎥)다. "
            f"최하층 경계면이 계획고보다 위에 있는지 확인할 것."
        )
    layers.append(
        LayerVolume(
            name=below_lowest_name,
            method=classify_stratum(below_lowest_name),
            cut_m3=max(below, 0.0),
            assumed=True,
        )
    )

    total = sum(layer.cut_m3 for layer in layers)
    residual = total - c_ground

    warnings: list[str] = []
    if abs(residual) > tol:
        # 항등식은 구조적으로 성립해야 한다. 어긋나면 계산이 아니라
        # 입력(예: 서로 다른 겹침 영역에서 나온 C 값)을 의심해야 한다.
        warnings.append(
            f"항등식 Σ Vᵢ + C(Sₙ) = C(S₀) 이 {residual:+,.6f} ㎥ 만큼 어긋났다. "
            f"각 C 값이 같은 산정 영역에서 나왔는지 확인할 것 — 서피스마다 "
            f"외곽 범위가 다르면 겹침 영역이 달라져 합이 맞지 않는다."
        )

    return DifferentialResult(
        layers=layers,
        total_cut_m3=total,
        identity_residual_m3=residual,
        warnings=warnings,
    )


# ---------------------------------------------------------------------------
# 시공방법별 집계
# ---------------------------------------------------------------------------

def aggregate_by_method(layers: Iterable[LayerVolume]) -> list[dict]:
    """층별 물량을 토사/리핑/발파 3분류로 집계한다.

    같은 분류에 단가가 다른 지층이 섞이면 대표 단가를 만들어 내지 않고
    ``unit_cost`` 를 null 로 두고 ``unit_cost_note`` 를 붙인다. 평균 단가를
    지어내면 결과만 보는 사람이 그것을 실제 계약 단가로 읽는다.
    """
    buckets: dict[str, list[LayerVolume]] = {m: [] for m in METHOD_ORDER}
    for layer in layers:
        buckets[layer.method].append(layer)

    out: list[dict] = []
    for method in METHOD_ORDER:
        members = buckets[method]
        if not members:
            continue
        cut = sum(m.cut_m3 for m in members)
        costs = [m.cost for m in members]
        entry: dict = {
            "method": method,
            "strata": [m.name for m in members],
            "cut_m3": cut,
            "assumed": any(m.assumed and m.cut_m3 > TOL_M3 for m in members),
        }

        unit_costs = {m.unit_cost for m in members}
        if len(unit_costs) == 1 and None not in unit_costs:
            entry["unit_cost"] = unit_costs.pop()
        else:
            entry["unit_cost"] = None
            if None in unit_costs:
                entry["unit_cost_note"] = (
                    "단가가 주어지지 않은 지층이 포함되어 있다: "
                    + ", ".join(m.name for m in members if m.unit_cost is None)
                )
            else:
                entry["unit_cost_note"] = (
                    "이 분류에 단가가 서로 다른 지층이 섞여 있어 대표 단가를 두지 않는다. "
                    "cost 는 지층별 단가로 계산한 합이다."
                )

        entry["cost"] = None if any(c is None for c in costs) else sum(costs)
        out.append(entry)
    return out


# ---------------------------------------------------------------------------
# 산정 영역 일치 검사
#
# 항등식 Σ Vᵢ + C(Sₙ) = C(S₀) 은 Vᵢ 의 정의상 망원경처럼 접히므로 **항상**
# 성립한다. 구조적 보장이지 검사가 아니며, 특히 실제로 가장 일어나기 쉬운
# 오류를 잡지 못한다 — 각 C(Sᵢ) 는 '그 서피스와 계획고의 겹침 영역'에서
# 계산되므로, 서피스마다 외곽 범위가 다르면 **서로 다른 면적의 값을 빼게 된다.**
# 그래도 항등식은 0 을 보고한다. 그래서 면적을 따로 본다.
# ---------------------------------------------------------------------------

def check_bbox_consistency(
    regions: Sequence[tuple[str, dict | None]],
    *,
    abs_tol: float = 1e-6,
) -> list[str]:
    """각 체적이 같은 산정 영역에서 나왔는지 경계 상자로 확인한다.

    볼륨 서피스의 Statistics 에는 2D 면적이 노출되지 않으므로(2026-08-16 실측)
    영역 비교는 경계 상자로 한다. 경계 상자가 같아도 실제 영역이 다를 수 있으니
    **일치의 증거가 아니라 불일치의 탐지**로만 쓴다. 확실히 하려면 boundary 를
    명시해 모든 호출이 같은 폴리곤을 쓰게 하는 편이 낫다.
    """
    known = [(n, b) for n, b in regions if b]
    unknown = [n for n, b in regions if not b]

    problems: list[str] = []
    if unknown:
        problems.append(
            f"산정 영역을 읽지 못한 항목이 있어 영역 일치를 확인할 수 없다: "
            f"{', '.join(unknown)}."
        )
    if len(known) < 2:
        return problems

    ref_name, ref = known[0]
    keys = ("min_x", "min_y", "max_x", "max_y")
    for name, box in known[1:]:
        diffs = [
            f"{k} {box.get(k, float('nan')):,.3f} vs {ref.get(k, float('nan')):,.3f}"
            for k in keys
            if abs(float(box.get(k, 0.0)) - float(ref.get(k, 0.0))) > abs_tol
        ]
        if diffs:
            problems.append(
                f"'{name}'의 산정 영역이 '{ref_name}'과 다르다 ({'; '.join(diffs)}). "
                f"두 값은 서로 다른 영역에서 나온 것이므로 그 차분을 층 물량으로 "
                f"쓸 수 없다. 항등식은 이 경우에도 성립하므로 이 검사가 없으면 "
                f"오류가 드러나지 않는다. 계획고 서피스와 각 경계면의 외곽 범위를 "
                f"맞추거나, boundary 를 명시해 모든 호출이 같은 폴리곤을 쓰게 할 것."
            )
    return problems


def total_cost(by_method: Sequence[dict]) -> tuple[float | None, list[str]]:
    """총 공사비. 하나라도 단가가 없으면 **합계를 내지 않는다.**

    빠진 항목을 0 으로 두고 합하면 총액이 계통적으로 과소 산정되는데,
    결과만 보아서는 그것이 '싼 후보지'로 보인다. 부지 비교가 목적이므로
    이 오류는 결론을 직접 뒤집는다.
    """
    missing = [e["method"] for e in by_method if e.get("cost") is None]
    if missing:
        return None, [
            f"시공방법 {missing} 의 공사비를 산정할 수 없어 총 공사비를 내지 않았다. "
            f"해당 지층의 단가를 입력하면 총액이 산출된다. "
            f"빠진 항목을 0 으로 두고 합하면 총액이 과소 산정되어 후보지 비교가 뒤집힌다."
        ]
    return sum(e["cost"] for e in by_method), []
