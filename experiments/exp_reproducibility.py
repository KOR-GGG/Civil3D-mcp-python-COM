"""exp_reproducibility.py — 재현성 측정 (제5장 C2 · 판정 기준 §4.6.2 ⓐ)

무엇을 재는가
-------------
**도구 계층의 결정성.** 동일 입력을 n회 반복해 반환값이 완전히 일치하는지 본다.

§4.6.2 ⓐ가 고정한 조건
  · 도구 인자      완전히 동일한 JSON
  · 도면 상태      동일 도면, **직전 호출의 부작용이 제거된 상태**
  · Civil 3D 세션  동일 프로세스, 동일 도면 단위(INSUNITS)
  · 도구 버전      동일 커밋 해시
  · 반복 횟수      n = 10
  · 판정          `elapsed_ms` 를 제외한 **모든 수치 필드가 완전 일치**

⚠ **이것은 도구 계층의 값이다.** 워크플로 계층의 재현성(동일 요청 n회에 대한
최종 수치의 편차)은 MCP 호스트가 필요하므로 별도로 측정해야 하며, 두 값을
섞어 보고하면 안 된다(§4.6.1).

부수 측정
---------
`elapsed_ms` 는 결정성 판정에서 제외하지만 **효율 지표(§4.2 도구 연산 시간)**
의 재료이므로 분포를 함께 보고한다.

또한 매 회 **도면 상태가 보존되는지**(서피스 수 불변, `__mcp_vol_*` 잔여 0)를
확인한다. 이것이 깨지면 "동일 입력"의 전제가 무너져 결정성 판정 자체가
성립하지 않는다.

실행
----
    .\.venv\Scripts\python.exe experiments\exp_reproducibility.py
"""
from __future__ import annotations

import json
import statistics
import subprocess
import sys
import time
from pathlib import Path

import win32com.client as w32

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "src"))

from civil3d_mcp import hydraulics_core as hc                  # noqa: E402
from civil3d_mcp.client import Civil3DClient                   # noqa: E402

N = 10
TARGET_DRAWING_MARKER = "TEST_C_S0_GROUND"      # 환경 1 식별용

STRATA = [
    {"name": "01토사층",   "surface": "TEST_C_S1_SOIL",      "unit_cost": 4500},
    {"name": "02풍화암층", "surface": "TEST_C_S2_WEATHERED", "unit_cost": 9000},
    {"name": "03연암층",   "surface": "TEST_C_S3_SOFT",      "unit_cost": 18000},
]

# 제외 대상 — 실행 시간이므로 당연히 변한다
VOLATILE_KEYS = {"elapsed_ms"}


def strip_volatile(obj):
    """비교에서 제외할 키를 재귀적으로 걷어낸다."""
    if isinstance(obj, dict):
        return {k: strip_volatile(v) for k, v in obj.items()
                if k not in VOLATILE_KEYS}
    if isinstance(obj, list):
        return [strip_volatile(v) for v in obj]
    return obj


def canon(obj) -> str:
    return json.dumps(strip_volatile(obj), ensure_ascii=False,
                      sort_keys=True, default=str)


def first_diff(a: dict, b: dict, path: str = "") -> str | None:
    """두 결과가 처음 갈라지는 지점을 찾아 사람이 읽을 문자열로."""
    if type(a) is not type(b):
        return f"{path}: 타입이 다름 {type(a).__name__} vs {type(b).__name__}"
    if isinstance(a, dict):
        for k in sorted(set(a) | set(b)):
            if k in VOLATILE_KEYS:
                continue
            if k not in a or k not in b:
                return f"{path}.{k}: 한쪽에만 존재"
            d = first_diff(a[k], b[k], f"{path}.{k}")
            if d:
                return d
        return None
    if isinstance(a, list):
        if len(a) != len(b):
            return f"{path}: 길이 {len(a)} vs {len(b)}"
        for i, (x, y) in enumerate(zip(a, b)):
            d = first_diff(x, y, f"{path}[{i}]")
            if d:
                return d
        return None
    if a != b:
        return f"{path}: {a!r} vs {b!r}"
    return None


def activate_env1(acad) -> str:
    """TEST_C_* 서피스가 있는 도면을 활성화한다."""
    for i in range(acad.Documents.Count):
        doc = acad.Documents.Item(i)
        doc.Activate()
        for _ in range(80):
            try:
                app = w32.GetActiveObject("AutoCAD.Application")
                if str(app.ActiveDocument.Name) != str(doc.Name):
                    time.sleep(0.25)
                    continue
                c = Civil3DClient()
                c.connect()
                if TARGET_DRAWING_MARKER in [s["name"] for s in c.list_surfaces()]:
                    return str(doc.Name)
                break
            except Exception:                                   # noqa: BLE001
                time.sleep(0.25)
    raise SystemExit(f"'{TARGET_DRAWING_MARKER}' 이 있는 도면을 찾지 못했다.")


def client_when_ready(timeout_s: float = 60.0) -> Civil3DClient:
    deadline = time.time() + timeout_s
    last = None
    while time.time() < deadline:
        try:
            c = Civil3DClient()
            c.connect()
            c.list_surfaces()
            return c
        except Exception as exc:                                # noqa: BLE001
            last = exc
            time.sleep(0.5)
    raise SystemExit(f"연결이 준비되지 않았다: {last}")


def _tool_elapsed(r: dict, wall_ms: float) -> float:
    """도구가 스스로 잰 elapsed_ms 를 쓴다.

    ⚠ 벽시계로 재면 상태 확인용 조회까지 포함되어 값이 부풀려진다.
    §4.2 가 정의한 「도구 연산 시간」은 **각 도구의 elapsed_ms** 이므로
    그것을 쓰고, 없을 때만 벽시계로 대체한다.
    """
    if "elapsed_ms" in r:
        return float(r["elapsed_ms"])
    inner = [v["elapsed_ms"] for v in r.values()
             if isinstance(v, dict) and "elapsed_ms" in v]
    return float(sum(inner)) if inner else wall_ms


def repeat(label: str, call, n: int = N) -> dict:
    """n회 반복하고 결정성·소요시간·도면 상태 보존을 함께 본다."""
    results, elapsed, states = [], [], []
    for _ in range(n):
        t0 = time.perf_counter()
        r = call()
        wall = (time.perf_counter() - t0) * 1000
        elapsed.append(_tool_elapsed(r, wall))
        results.append(r)
        states.append(r.pop("__state__", None))

    base = results[0]
    mismatches = []
    for i, r in enumerate(results[1:], start=2):
        if canon(r) != canon(base):
            mismatches.append((i, first_diff(base, r)))

    return {
        "label": label,
        "n": n,
        "identical": not mismatches,
        "mismatches": mismatches,
        "elapsed_ms": {
            "min": min(elapsed), "max": max(elapsed),
            "mean": statistics.fmean(elapsed),
            "stdev": statistics.stdev(elapsed) if n > 1 else 0.0,
        },
        "state_preserved": len(set(map(str, states))) <= 1,
        "state_sample": states[0],
    }


def main() -> int:
    acad = w32.GetActiveObject("AutoCAD.Application")
    drawing = activate_env1(acad)
    client = client_when_ready()

    commit = subprocess.run(["git", "rev-parse", "--short", "HEAD"],
                            cwd=REPO, capture_output=True, text=True).stdout.strip()
    units = client._drawing_length_unit()
    n_surfaces = len(client.list_surfaces())

    print("=" * 92)
    print("  재현성 — 도구 계층 (제5장 C2 · 판정 기준 §4.6.2 ⓐ)")
    print("=" * 92)
    print("  고정한 조건")
    print(f"    도면          {drawing} (서피스 {n_surfaces}개)")
    print(f"    도면 단위     {units['name']} (INSUNITS={units['insunits']})")
    print(f"    도구 버전     커밋 {commit}")
    print(f"    반복 횟수     n = {N}")
    print(f"    비교 제외     {sorted(VOLATILE_KEYS)}")
    print()
    print("  ⚠ 워크플로 계층의 재현성은 MCP 호스트가 필요하므로 여기 포함되지 않는다.")
    print()

    def with_state(fn):
        """호출 후 도면 상태(서피스 수·임시객체 잔여)를 함께 담는다."""
        def _inner():
            r = dict(fn())
            names = [s["name"] for s in client.list_surfaces()]
            leftovers = [x for x in names if x.startswith("__mcp_vol_")]
            r["__state__"] = f"surfaces={len(names)} leftovers={leftovers}"
            return r
        return _inner

    targets = [
        ("L1 compute_volume_between_surfaces (평면쌍)",
         with_state(lambda: client.compute_volume_between_surfaces(
             "TEST_A_BASE", "TEST_A_COMP"))),
        ("L1 + 경계 클립 (좌표열)",
         with_state(lambda: client.compute_volume_between_surfaces(
             "TEST_A_BASE", "TEST_A_COMP",
             [[0, 0], [50, 0], [50, 100], [0, 100]]))),
        ("L2 compute_earthwork_by_rock_quality",
         with_state(lambda: client.compute_earthwork_by_rock_quality(
             "TEST_C_S0_GROUND", "TEST_C_DESIGN", STRATA,
             site_id="repro", below_lowest_unit_cost=25000))),
        ("수리 3종 연쇄 (COM 비접근)",
         lambda: {
             "econ": hc.select_economic_diameter(
                 flow_m3_s=0.35, length_m=4200.0,
                 candidate_diameters_mm=[400, 450, 500, 600, 700, 800],
                 unit_construction_cost=[
                     {"diameter_mm": d, "cost_per_m": c} for d, c in
                     {400: 430_000, 450: 480_000, 500: 550_000,
                      600: 690_000, 700: 860_000, 800: 1_050_000}.items()],
                 power_tariff_won_per_kwh=130, pump_efficiency=0.75,
                 annual_operating_hours=8760, discount_rate=0.045,
                 service_life_years=30),
             "prof": hc.hydraulic_profile(
                 flow_m3_s=0.35, diameter_mm=600, start_elevation_m=42.0,
                 start_pump_head_m=50.0, min_residual_head_m=5.0, pump_head_m=40.0,
                 ground_profile=[
                     {"station": 0, "elevation": 42.0},
                     {"station": 1000, "elevation": 55.0},
                     {"station": 2000, "elevation": 78.0},
                     {"station": 2500, "elevation": 86.0},
                     {"station": 4200, "elevation": 90.0}]),
         }),
    ]

    print("-" * 92)
    print(f"  {'대상':44s} {'일치':6s} {'평균(ms)':>10s} {'편차':>9s}  도면 상태")
    print("-" * 92)

    rows = []
    for label, call in targets:
        r = repeat(label, call)
        rows.append(r)
        mark = "O" if r["identical"] else "X"
        e = r["elapsed_ms"]
        state = ("보존" if r["state_preserved"] else "⚠변함") if r["state_sample"] else "—"
        print(f"  {label:44s} {mark:^6s} {e['mean']:10.1f} {e['stdev']:9.2f}  {state}")
        if not r["identical"]:
            for i, d in r["mismatches"][:3]:
                print(f"       ⚠ {i}회차 불일치: {d}")

    print()
    print("=" * 92)
    print("  결과")
    print("=" * 92)
    det = [r for r in rows if r["identical"]]
    print(f"  {len(det)}/{len(rows)} 대상이 {N}회 반복에서 **완전히 동일한 값**을 반환")
    bad_state = [r for r in rows if r["state_sample"] and not r["state_preserved"]]
    print(f"  도면 상태 보존: {'전부 보존' if not bad_state else '⚠ ' + str([r['label'] for r in bad_state])}")
    print()
    print("  소요시간(§4.2 도구 연산 시간의 재료)")
    for r in rows:
        e = r["elapsed_ms"]
        print(f"    {r['label']:44s} {e['min']:8.1f} ~ {e['max']:8.1f} ms")
    print()
    if len(det) == len(rows):
        print("  → 도구 계층은 결정적이다(원칙 P2). 즉 최종 수치에 편차가 관측된다면")
        print("    그 출처는 도구가 아니라 **에이전트 계층**이다.")

    out = Path(__file__).with_name("reproducibility_result.json")
    out.write_text(json.dumps({
        "layer": "tool", "n": N, "drawing": drawing, "commit": commit,
        "units": units, "excluded_keys": sorted(VOLATILE_KEYS), "rows": rows,
    }, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(f"\n  결과 저장 -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
