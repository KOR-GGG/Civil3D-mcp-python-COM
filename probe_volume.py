"""
probe_volume.py — compute_volume_between_surfaces 단계별 진단·검증 스크립트

⚠ 반드시 **사본 도면**으로 실행할 것. 2023년 연구성과 원본을 Civil 3D 2026으로
   저장하면 하위 버전에서 열리지 않을 수 있다.

사용법 (Civil 3D를 켜고 대상 도면을 연 상태에서):

    cd "$env:USERPROFILE\\source\\civil3d-mcp"

    # 1~3단계만 — 연결·서피스 목록·coclass 확보 (읽기 전용)
    .\\.venv\\Scripts\\python.exe probe_volume.py

    # 4~5단계 — 체적 계산 + 부작용 검사
    .\\.venv\\Scripts\\python.exe probe_volume.py <기준면> <비교면>

    # 6단계 — 기준값 대조까지 (선택)
    .\\.venv\\Scripts\\python.exe probe_volume.py <기준면> <비교면> <절토기준값> <성토기준값>

기준값은 인자로 넘긴다. 프로젝트 실측치를 소스에 넣지 않기 위함이다.
"""
from __future__ import annotations

import sys
import traceback

from civil3d_mcp.client import Civil3DClient, Civil3DError

# 2026-08-19 추가: 한글 Windows 콘솔은 기본이 cp949 라, 아래 출력에 쓰이는
# em-dash 나 경고 기호 하나 때문에 UnicodeEncodeError 로 스크립트가 즉사한다.
# (setup_check.py 에서 같은 결함을 고쳤으나 실험 스크립트에는 남아 있었다.)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:                                          # noqa: BLE001
        pass

SEP = "=" * 68


def step(n: int, title: str) -> None:
    print(f"\n{SEP}\n[{n}] {title}\n{SEP}")


def main() -> int:
    client = Civil3DClient()

    # ------------------------------------------------------------------
    step(1, "연결")
    try:
        client.connect()
    except Civil3DError as exc:
        print(f"FAIL — {exc}")
        print("→ Civil 3D가 실행 중이고 도면이 열려 있는지 확인할 것.")
        return 1
    print(f"OK   ProgID = {client._prog_id}")
    if client._prog_id is None:
        print("FAIL — Civil 3D 인터페이스 획득 실패(일반 AutoCAD로 폴백).")
        print("→ Civil 3D 컬렉션에 접근할 수 없으므로 이후 단계는 무의미하다.")
        return 1

    # ------------------------------------------------------------------
    step(2, "서피스 목록")
    try:
        surfaces = client.list_surfaces()
    except Civil3DError as exc:
        print(f"FAIL — {exc}")
        return 1
    print(f"OK   {len(surfaces)}개 인식")
    for s in surfaces:
        rng = ""
        if "min_elevation" in s:
            rng = f"  EL {s['min_elevation']:.2f} ~ {s['max_elevation']:.2f}"
        print(f"   - {s['name']}{rng}")

    # ------------------------------------------------------------------
    step(3, "AeccTinVolumeCreationData coclass 확보")
    print("시도 순서:", ", ".join(client._aecc_versions_to_try()))
    try:
        data = client._new_aecc_object("AeccTinVolumeCreationData")
    except Civil3DError as exc:
        print(f"FAIL — {exc}")
        print("\n→ 이 단계가 실패하면 COM 경로가 막힌 것이다. 확인할 것:")
        print("   · 레지스트리에서 AeccXLand.AeccTinVolumeCreationData 등록 버전 조회")
        print("     PowerShell: Get-ChildItem 'HKLM:\\SOFTWARE\\Classes' -Name |")
        print("                 Select-String 'AeccTinVolumeCreationData'")
        print("   · 나온 버전을 client.py의 _AECC_LAND_VERSIONS 앞에 추가")
        return 1
    print("OK   coclass 생성 성공")
    writable = []
    for prop in ("Name", "Layer", "BaseLayer", "Style", "Description",
                 "BaseSurface", "ComparisonSurface"):
        writable.append(f"{prop}={'있음' if hasattr(data, prop) else '없음'}")
    print("     속성:", ", ".join(writable))

    if len(sys.argv) < 3:
        print("\n서피스 2개를 인자로 주면 4단계(체적 계산)까지 진행한다.")
        print("예: python probe_volume.py <기준면> <비교면>")
        print("    기준면 = 원지반 또는 암층 경계면 / 비교면 = 계획고(코리더 데이텀)")
        return 0

    base_name, comp_name = sys.argv[1], sys.argv[2]

    # ------------------------------------------------------------------
    step(4, f"체적 계산  base='{base_name}'  comparison='{comp_name}'")
    before = len(surfaces)
    try:
        result = client.compute_volume_between_surfaces(base_name, comp_name)
    except Civil3DError as exc:
        print(f"FAIL — {exc}")
        print("\n→ 가장 흔한 원인 두 가지:")
        print("   (a) data.BaseSurface = <COM 객체> 대입이 거부됨")
        print("       → win32com putref 문제. client.py에서")
        print("         data._oleobj_.Invoke(...) 방식으로 바꿔야 할 수 있다.")
        print("   (b) Style 이름이 도면 템플릿에 없음 → 스타일명 확인")
        traceback.print_exc()
        return 1
    except Exception:
        print("FAIL — 예상치 못한 예외")
        traceback.print_exc()
        return 1

    print("OK")
    for k, v in result.items():
        print(f"     {k}: {v}")

    # ------------------------------------------------------------------
    step(5, "부작용 검사 — 임시 볼륨 서피스가 제거되었는가 (원칙 P4)")
    after_list = client.list_surfaces()
    leftovers = [s["name"] for s in after_list if s["name"].startswith("__mcp_vol_")]
    if leftovers:
        print(f"FAIL — 임시 서피스가 남아 있다: {leftovers}")
        print("→ 수동 삭제 후 Erase() 경로를 점검할 것.")
    elif len(after_list) != before:
        print(f"WARN — 서피스 수가 {before} → {len(after_list)}로 변했다.")
    else:
        print(f"OK   서피스 수 {before}개 유지, 잔여 임시 객체 없음")

    # ------------------------------------------------------------------
    step(6, "기준값 대조")
    cut, fill = result["cut_m3"], result["fill_m3"]
    print(f"     도구 산출  절토 {cut:>14,.2f} ㎥   성토 {fill:>14,.2f} ㎥")

    if len(sys.argv) < 5:
        print("\n     기준값이 주어지지 않아 대조를 건너뛴다.")
        print("     대조하려면: python probe_volume.py <기준면> <비교면> <절토기준값> <성토기준값>")
    else:
        try:
            ref_cut, ref_fill = float(sys.argv[3]), float(sys.argv[4])
        except ValueError:
            print("     기준값 인자를 숫자로 해석할 수 없다. 대조를 건너뛴다.")
        else:
            print(f"     기준값     절토 {ref_cut:>14,.2f} ㎥   성토 {ref_fill:>14,.2f} ㎥")
            for label, got, ref in (("절토", cut, ref_cut), ("성토", fill, ref_fill)):
                if ref:
                    err = (got - ref) / ref * 100
                    mark = "OK  " if abs(err) < 1.0 else "확인"
                    print(f"     {mark} {label} 오차 {err:+.3f}%")
                else:
                    mark = "OK  " if abs(got) < 1e-6 else "확인"
                    print(f"     {mark} {label} 기준값 0 — 산출값 {got:,.2f}")

    print("\n※ 부호가 뒤바뀌어 보이면 base/comparison을 서로 바꿔 재실행할 것.")
    print("   볼륨 서피스는 'comparison - base'로 정의된다.")
    print("※ 오차가 크면 boundary(산정 영역) 문제일 가능성이 높다.")
    print("   계획고 서피스가 코리더 데이텀 서피스인지, 그 외곽이 사면을")
    print("   포함하는지 확인할 것 — 부지 경계로 잘리면 사면부 물량이 빠진다.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    finally:
        pass
