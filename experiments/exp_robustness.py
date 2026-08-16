"""exp_robustness.py — 오식별 검출률 측정 (제5장 §4.4 · 판정 기준 §4.6.4)

무엇을 재는가
-------------
잘못된 입력을 도구에 **직접** 투입하여, 도구가 값을 내지 않고 사유를 밝히며
중단하는지(=검출) 본다.

⚠ **이것은 도구 계층의 검출률이다.** 에이전트 계층의 검출률(교란 도면을
에이전트에게 제시했을 때의 반려 비율)은 MCP 호스트가 필요하므로 별도로
측정해야 하며, 두 값을 섞어 보고하면 안 된다(§4.6.1).

판정 (§4.6.4)
-------------
  반려  : 수치를 반환하지 않고 사유를 밝히며 중단
  경고  : 값은 내되 warnings 에 사유를 남김
  미검출: 값을 내고 경고도 없음

선행 조건
---------
    .\.venv\Scripts\python.exe make_test_surfaces.py C P
Civil 3D 가 실행 중이고 해당 서피스가 있어야 한다.

실행
----
    .\.venv\Scripts\python.exe experiments\exp_robustness.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from civil3d_mcp.client import Civil3DClient, Civil3DError   # noqa: E402

COST = {"토사": 4500, "풍화암": 9000, "연암": 18000, "경암": 25000}

NORMAL_STRATA = [
    {"name": "01토사층",   "surface": "TEST_C_S1_SOIL",      "unit_cost": COST["토사"]},
    {"name": "02풍화암층", "surface": "TEST_C_S2_WEATHERED", "unit_cost": COST["풍화암"]},
    {"name": "03연암층",   "surface": "TEST_C_S3_SOFT",      "unit_cost": COST["연암"]},
]


def case_p1():
    """층서 역전 — 토사 하면(EL92)이 풍화암 하면(EL95)보다 아래."""
    return dict(
        ground_surface="TEST_P_S0_GROUND",
        design_surface="TEST_P_DESIGN",
        strata=[
            {"name": "01토사층",   "surface": "TEST_P_S1_SOIL_BAD",
             "unit_cost": COST["토사"]},
            {"name": "02풍화암층", "surface": "TEST_P_S2_WEATHERED_BAD",
             "unit_cost": COST["풍화암"]},
        ],
        below_lowest_unit_cost=COST["경암"],
    )


def case_p2():
    """지층명을 분류 불가로 교란."""
    s = [dict(x) for x in NORMAL_STRATA]
    s[0]["name"] = "층A"
    return dict(ground_surface="TEST_C_S0_GROUND", design_surface="TEST_C_DESIGN",
                strata=s, below_lowest_unit_cost=COST["경암"])


def case_p3():
    """지층명이 두 암질에 걸침."""
    s = [dict(x) for x in NORMAL_STRATA]
    s[0]["name"] = "토사및풍화암혼재층"
    return dict(ground_surface="TEST_C_S0_GROUND", design_surface="TEST_C_DESIGN",
                strata=s, below_lowest_unit_cost=COST["경암"])


def case_p4():
    """원지반과 계획고를 뒤바꿔 지정.

    2026-08-17 1차 실행에서 **미검출**로 나왔다 — 모든 층이 0이 되는데도
    아무 경고가 없었다. 명세에 *"절토(원지반 > 계획고) 구간 대상"*이라는
    전제가 있으나 코드에 검사가 없었던 것이다. 전면 성토 부지가 있을 수
    있으므로 반려가 아니라 **경고**로 처리하도록 구현을 보강했다.
    """
    return dict(ground_surface="TEST_C_DESIGN", design_surface="TEST_C_S0_GROUND",
                strata=NORMAL_STRATA, below_lowest_unit_cost=COST["경암"])


def case_p5():
    """산정 영역이 다른 서피스 혼입 — 외곽이 좁은 지층면(면적 8,500 m2)."""
    return dict(
        ground_surface="TEST_C_S0_GROUND", design_surface="TEST_C_DESIGN",
        strata=[
            {"name": "01토사층", "surface": "TEST_P_NARROW_SOIL",
             "unit_cost": COST["토사"]},
        ],
        below_lowest_unit_cost=COST["경암"],
    )


def case_p6():
    """단가 누락."""
    s = [dict(x) for x in NORMAL_STRATA]
    s[1].pop("unit_cost")
    return dict(ground_surface="TEST_C_S0_GROUND", design_surface="TEST_C_DESIGN",
                strata=s, below_lowest_unit_cost=COST["경암"])


CASES = [
    ("P1", "층서 역전",              case_p1, "반려"),
    ("P2", "지층명 분류 불가",        case_p2, "반려"),
    ("P3", "지층명 모호",            case_p3, "반려"),
    ("P4", "원지반/계획고 뒤바꿈",     case_p4, "경고"),
    ("P5", "산정 영역 불일치",        case_p5, "경고"),
    ("P6", "단가 누락",              case_p6, "경고"),
]

# 정상 케이스 — 교란이 아닌 입력을 잘못 반려하지 않는지(위양성) 확인
CONTROL = ("N0", "정상 입력(대조군)",
           lambda: dict(ground_surface="TEST_C_S0_GROUND",
                        design_surface="TEST_C_DESIGN",
                        strata=NORMAL_STRATA,
                        below_lowest_unit_cost=COST["경암"]),
           "정상")


def classify(client, kwargs) -> tuple[str, str]:
    """(판정, 근거 한 줄)"""
    try:
        r = client.compute_earthwork_by_rock_quality(**kwargs)
    except Civil3DError as exc:
        return "반려", str(exc).replace("\n", " ")[:150]
    warns = r.get("warnings") or []
    if warns:
        return "경고", warns[0][:150]
    return "미검출", (f"total_cut={r['total_cut_m3']:,.2f} "
                     f"total_cost={r['total_earthwork_cost']}")


def main() -> int:
    c = Civil3DClient()
    c.connect()

    print("=" * 96)
    print("  오식별 검출률 — 도구 계층 (제5장 §4.4 · 판정 기준 §4.6.4)")
    print("=" * 96)
    print("  ⚠ 에이전트 계층의 검출률은 MCP 호스트가 필요하므로 여기 포함되지 않는다.")
    print()
    print(f"  {'#':4s} {'교란 유형':22s} {'기대':6s} {'실제':6s}  판정")
    print("  " + "-" * 92)

    rows = []
    for cid, label, builder, expected in [CONTROL] + CASES:
        got, why = classify(c, builder())
        if expected == "정상":
            ok = (got == "미검출")          # 정상 입력은 값이 나와야 정상
            verdict = "OK (위양성 없음)" if ok else "!! 정상 입력을 반려/경고함"
        elif expected == "미정":
            ok = None
            verdict = "기대 동작 미정 — 아래 해석 참조"
        else:
            ok = (got == expected)
            verdict = "OK" if ok else f"!! 기대 {expected} != 실제 {got}"
        print(f"  {cid:4s} {label:22s} {expected:6s} {got:6s}  {verdict}")
        print(f"       근거: {why}")
        rows.append({"id": cid, "label": label, "expected": expected,
                     "observed": got, "ok": ok, "evidence": why})

    print()
    print("=" * 96)
    print("  집계")
    print("=" * 96)
    graded = [r for r in rows if r["ok"] is not None and r["id"] != "N0"]
    hit = sum(1 for r in graded if r["ok"])
    print(f"  판정 가능한 교란 {len(graded)}종 중 기대대로 동작 {hit}종 "
          f"= {hit/len(graded)*100:.1f}%")
    undecided = [r for r in rows if r["ok"] is None]
    if undecided:
        print(f"  ⚠ 기대 동작 미정 {len(undecided)}종: "
              f"{', '.join(r['id'] for r in undecided)} — §4.6.5에서 결정 필요")
    ctrl = rows[0]
    print(f"  대조군(정상 입력): {ctrl['observed']} "
          f"→ {'위양성 없음' if ctrl['observed'] == '미검출' else '⚠ 위양성 발생'}")

    out = Path(__file__).with_name("robustness_result.json")
    out.write_text(json.dumps({"layer": "tool", "rows": rows},
                              ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n  결과 저장 -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
