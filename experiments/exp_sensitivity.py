"""exp_sensitivity.py — 경제적 관경 선정의 민감도 검토 (제5장 §4.5)

왜 필요한가
-----------
기준 사례에서 D600(총현가 4,206백만)과 차순위 D700(4,230백만)의 차이가
**0.55%** 에 불과하다. 입력이 조금만 달라져도 선정 결과가 뒤집힐 수 있다는
뜻이며, 그렇다면 "도구가 기준값과 같은 관경을 골랐다"는 정확도 판정이
얼마나 견고한지를 함께 보고해야 한다.

이 실험은 **선정 관경이 뒤집히는 임계값**을 입력 변수별로 찾는다.

측정 방법
---------
각 변수를 단조 변화시키며 선정 관경을 관찰하고, 관경이 바뀌는 지점을
이분법으로 좁힌다. 모든 계산은 COM 에 접근하지 않으므로 Civil 3D 없이 돈다.

실행
----
    .\.venv\Scripts\python.exe experiments\exp_sensitivity.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from civil3d_mcp import hydraulics_core as hc   # noqa: E402

# ---------------------------------------------------------------------------
# 기준 사례 — 도구명세_경제관경수리검토.md 의 검산 예시
# ---------------------------------------------------------------------------
BASE = dict(
    flow_m3_s=0.35,
    length_m=4200.0,
    candidate_diameters_mm=[300, 400, 450, 500, 600, 700, 800],
    power_tariff_won_per_kwh=130.0,
    pump_efficiency=0.75,
    annual_operating_hours=8760.0,
    discount_rate=0.045,
    service_life_years=30,
)
UNIT_COST = {
    300: 330_000.0, 400: 430_000.0, 450: 480_000.0, 500: 550_000.0,
    600: 690_000.0, 700: 860_000.0, 800: 1_050_000.0,
}
TOL = 1e-4          # 이분법 수렴 폭(상대비)


def select(cost_table: dict[int, float], **over) -> dict:
    kw = dict(BASE)
    kw.update(over)
    kw["unit_construction_cost"] = [
        {"diameter_mm": d, "cost_per_m": c} for d, c in cost_table.items()
    ]
    return hc.select_economic_diameter(**kw)


def chosen(cost_table: dict[int, float], **over) -> int:
    return select(cost_table, **over)["selected"]["diameter_mm"]


def scaled(factor: float, only: int | None = None) -> dict[int, float]:
    """건설단가 배율. only 가 주어지면 그 관경만 변동시킨다."""
    return {
        d: (c * factor if (only is None or d == only) else c)
        for d, c in UNIT_COST.items()
    }


def find_flip(make_case, lo: float, hi: float, base_choice: int):
    """[lo, hi] 구간에서 선정 관경이 바뀌는 지점을 이분법으로 찾는다.

    lo 에서는 기준과 같고 hi 에서는 달라야 한다. 그렇지 않으면 None.
    """
    if chosen(*make_case(lo)[0], **make_case(lo)[1]) != base_choice:
        return None, None
    hi_choice = chosen(*make_case(hi)[0], **make_case(hi)[1])
    if hi_choice == base_choice:
        return None, None

    a, b = lo, hi
    for _ in range(60):
        m = (a + b) / 2
        if chosen(*make_case(m)[0], **make_case(m)[1]) == base_choice:
            a = m
        else:
            b = m
        if abs(b - a) < TOL:
            break
    # ⚠ 바뀐 관경은 **임계점 직후(b)** 에서 평가해야 한다.
    # 탐색 구간 끝(hi)의 선택을 보고하면 임계값은 맞고 관경만 틀리는,
    # 표만 보아서는 알 수 없는 오류가 된다.
    return b, chosen(*make_case(b)[0], **make_case(b)[1])


def sweep_report(label: str, make_case, base_choice: int,
                 down=(1.0, 0.5), up=(1.0, 2.0), unit="배율"):
    """아래·위 양방향으로 임계값을 찾아 한 줄로 보고한다."""
    lo_thr, lo_new = find_flip(make_case, down[0], down[1], base_choice)
    hi_thr, hi_new = find_flip(make_case, up[0], up[1], base_choice)

    def fmt(thr, new):
        if thr is None:
            return "역전 없음"
        pct = (thr - 1.0) * 100
        return f"{pct:+.2f}% (→ D{new})"

    print(f"  {label:34s} 하향 {fmt(lo_thr, lo_new):>22s}   상향 {fmt(hi_thr, hi_new):>22s}")
    return {
        "variable": label,
        "down_threshold_pct": None if lo_thr is None else (lo_thr - 1.0) * 100,
        "down_new_diameter": lo_new,
        "up_threshold_pct": None if hi_thr is None else (hi_thr - 1.0) * 100,
        "up_new_diameter": hi_new,
    }


def main() -> int:
    base = select(UNIT_COST)
    base_choice = base["selected"]["diameter_mm"]
    ranked = sorted(base["candidates"], key=lambda c: c["total_pv"])

    print("=" * 92)
    print("  경제적 관경 선정 민감도 검토 — 제5장 §4.5")
    print("=" * 92)
    print(f"  기준 사례: Q={BASE['flow_m3_s']} m3/s · L={BASE['length_m']:,.0f} m · "
          f"전력 {BASE['power_tariff_won_per_kwh']:.0f}원/kWh · "
          f"i={BASE['discount_rate']:.3f} · n={BASE['service_life_years']}년")
    print(f"  Hazen-Williams 계수 = {hc.HW_COEFFICIENT} (2026-08-16 확정)")
    print()
    print(f"  {'관경':>6s} {'유속':>8s} {'총현가(백만원)':>16s} {'1위 대비':>10s}")
    best_pv = ranked[0]["total_pv"]
    for c in ranked:
        gap = (c["total_pv"] - best_pv) / best_pv * 100
        mark = "  <= 선정" if c["diameter_mm"] == base_choice else ""
        print(f"  D{c['diameter_mm']:<5d} {c['velocity_m_s']:8.3f} "
              f"{c['total_pv']/1e6:16,.0f} {gap:9.2f}%{mark}")
    print()
    print(f"  차순위 격차 = {base['runner_up_margin']*100:.4f} %  "
          f"(D{ranked[0]['diameter_mm']} vs D{ranked[1]['diameter_mm']})")
    print()

    print("-" * 92)
    print("  선정 관경이 뒤집히는 임계값")
    print("-" * 92)

    rows = []
    # ① 전체 건설단가
    rows.append(sweep_report(
        "전체 건설단가", lambda f: ((scaled(f),), {}), base_choice,
        down=(1.0, 0.3), up=(1.0, 3.0)))
    # ② 선정 관경의 단가만
    rows.append(sweep_report(
        f"D{base_choice} 단가만", lambda f: ((scaled(f, only=base_choice),), {}),
        base_choice, down=(1.0, 0.3), up=(1.0, 3.0)))
    # ③ 차순위 관경의 단가만
    runner = ranked[1]["diameter_mm"]
    rows.append(sweep_report(
        f"D{runner} 단가만 (차순위)", lambda f: ((scaled(f, only=runner),), {}),
        base_choice, down=(1.0, 0.3), up=(1.0, 3.0)))
    # ④ 전력요금
    rows.append(sweep_report(
        "전력요금", lambda f: ((UNIT_COST,), {
            "power_tariff_won_per_kwh": BASE["power_tariff_won_per_kwh"] * f}),
        base_choice, down=(1.0, 0.1), up=(1.0, 5.0)))
    # ⑤ 할인율
    rows.append(sweep_report(
        "할인율", lambda f: ((UNIT_COST,), {
            "discount_rate": BASE["discount_rate"] * f}),
        base_choice, down=(1.0, 0.1), up=(1.0, 5.0)))
    # ⑥ 연간 가동시간
    rows.append(sweep_report(
        "연간 가동시간", lambda f: ((UNIT_COST,), {
            "annual_operating_hours": BASE["annual_operating_hours"] * f}),
        base_choice, down=(1.0, 0.1), up=(1.0, 1.0)))
    # ⑦ 펌프효율
    rows.append(sweep_report(
        "펌프효율", lambda f: ((UNIT_COST,), {
            "pump_efficiency": min(BASE["pump_efficiency"] * f, 1.0)}),
        base_choice, down=(1.0, 0.5), up=(1.0, 1.0 / 0.75)))

    print()
    print("=" * 92)
    print("  해석")
    print("=" * 92)
    tight = [r for r in rows
             if (r["down_threshold_pct"] is not None and abs(r["down_threshold_pct"]) < 10)
             or (r["up_threshold_pct"] is not None and abs(r["up_threshold_pct"]) < 10)]
    if tight:
        print("  ⚠ 10% 이내의 변동으로 선정이 뒤집히는 변수:")
        for r in tight:
            print(f"     - {r['variable']}")
    else:
        print("  10% 이내 변동으로 뒤집히는 변수 없음")

    out = Path(__file__).with_name("sensitivity_result.json")
    out.write_text(json.dumps({
        "base_case": {k: v for k, v in BASE.items()},
        "unit_cost": UNIT_COST,
        "hw_coefficient": hc.HW_COEFFICIENT,
        "selected": base_choice,
        "runner_up_margin_pct": base["runner_up_margin"] * 100,
        "ranking": [{"diameter_mm": c["diameter_mm"],
                     "total_pv": c["total_pv"]} for c in ranked],
        "thresholds": rows,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n  결과 저장 -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
