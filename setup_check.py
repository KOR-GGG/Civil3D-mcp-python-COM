"""
setup_check.py  –  civil3d-mcp pre-flight environment checker
==============================================================
Run this BEFORE starting the MCP server to verify that all
dependencies and system conditions are met.

Usage:
    python setup_check.py
    python setup_check.py --fix      # attempt auto-fixes (pip install)
    python setup_check.py --json     # machine-readable output
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable


# ---------------------------------------------------------------------------
# Result model
# ---------------------------------------------------------------------------

@dataclass
class CheckResult:
    name: str
    passed: bool
    detail: str
    fix_hint: str = ""


# ---------------------------------------------------------------------------
# Individual checks
# ---------------------------------------------------------------------------

def check_platform() -> CheckResult:
    ok = platform.system() == "Windows"
    return CheckResult(
        name="Windows OS",
        passed=ok,
        detail=f"Detected: {platform.system()} {platform.release()}",
        fix_hint="COM automation requires Windows 10 or 11." if not ok else "",
    )


def check_python_version() -> CheckResult:
    v = sys.version_info
    ok = (v.major == 3 and v.minor >= 11)
    return CheckResult(
        name="Python >= 3.11",
        passed=ok,
        detail=f"Detected: Python {v.major}.{v.minor}.{v.micro}",
        fix_hint=(
            "Install Python 3.11+ from https://python.org "
            "and re-run this script."
        ) if not ok else "",
    )


def check_python_arch() -> CheckResult:
    bits = "64-bit" if sys.maxsize > 2**32 else "32-bit"
    ok = sys.maxsize > 2**32
    return CheckResult(
        name="Python 64-bit",
        passed=ok,
        detail=f"Detected: {bits}",
        fix_hint=(
            "Civil 3D is 64-bit. Install the 64-bit Python distribution."
        ) if not ok else "",
    )


def _import_ok(module: str) -> tuple[bool, str]:
    try:
        __import__(module)
        mod = sys.modules[module]
        ver = getattr(mod, "__version__", "unknown")
        return True, ver
    except ImportError as exc:
        return False, str(exc)


def check_fastmcp() -> CheckResult:
    ok, ver = _import_ok("mcp")
    return CheckResult(
        name="fastmcp / mcp package",
        passed=ok,
        detail=f"version: {ver}" if ok else ver,
        fix_hint="pip install fastmcp" if not ok else "",
    )


def check_win32com() -> CheckResult:
    ok, ver = _import_ok("win32com.client")
    return CheckResult(
        name="pywin32 (win32com)",
        passed=ok,
        detail=f"version: {ver}" if ok else ver,
        fix_hint="pip install pywin32" if not ok else "",
    )


def check_pythoncom() -> CheckResult:
    ok, ver = _import_ok("pythoncom")
    return CheckResult(
        name="pythoncom",
        passed=ok,
        detail="present" if ok else ver,
        fix_hint="pip install pywin32  (pythoncom ships with pywin32)" if not ok else "",
    )


def check_pythonnet() -> CheckResult:
    ok, ver = _import_ok("clr")
    return CheckResult(
        name="pythonnet (clr)",
        passed=ok,
        detail=f"version: {ver}" if ok else ver,
        fix_hint="pip install pythonnet" if not ok else "",
    )


def check_pydantic() -> CheckResult:
    ok, ver = _import_ok("pydantic")
    return CheckResult(
        name="pydantic",
        passed=ok,
        detail=f"version: {ver}" if ok else ver,
        fix_hint="pip install pydantic" if not ok else "",
    )


# --------------- Civil 3D binary paths ---------------

# 2026-08-18 정정: client.py 의 _AUTODESK_ROOTS 와 어긋나 있었다. 실제 DLL 은
# "AutoCAD 20xx\C3D\" 하위에 있는데 여기서는 상위 폴더만 봐서, Civil 3D 가
# 정상 설치된 PC 에서도 "Missing ... AeccDbMgd.dll" 로 오탐했다. 2026 도 빠져 있었다.
_CANDIDATE_ROOTS = [
    r"C:\Program Files\Autodesk\AutoCAD 2026\C3D",
    r"C:\Program Files\Autodesk\AutoCAD 2026",
    r"C:\Program Files\Autodesk\AutoCAD 2025\C3D",
    r"C:\Program Files\Autodesk\AutoCAD 2025",
    r"C:\Program Files\Autodesk\AutoCAD 2024\C3D",
    r"C:\Program Files\Autodesk\AutoCAD 2024",
    r"C:\Program Files\Autodesk\AutoCAD 2023",
]
# 항상 존재하는 DLL
_REQUIRED_DLLS = ["AeccDbMgd.dll", "acdbmgd.dll"]
# Civil 3D 2023/2024 에는 있고 2025+ 에서는 제거된 DLL — 없어도 정상이다.
_OPTIONAL_DLLS = ["AeccLandMgd.dll"]


def _find_civil3d_root() -> str | None:
    env_path = os.getenv("CIVIL3D_BIN_PATH", "").strip()
    if env_path and Path(env_path).is_dir():
        return env_path
    for root in _CANDIDATE_ROOTS:
        if Path(root).is_dir():
            return root
    return None


def check_civil3d_install() -> CheckResult:
    root = _find_civil3d_root()
    if root is None:
        return CheckResult(
            name="Civil 3D installation",
            passed=False,
            detail="No Civil 3D folder found in default paths.",
            fix_hint=(
                "Install Civil 3D 2023-2025, or set CIVIL3D_BIN_PATH "
                "in .env to the folder containing AeccDbMgd.dll."
            ),
        )
    return CheckResult(
        name="Civil 3D installation",
        passed=True,
        detail=f"Found: {root}",
    )


def check_autodesk_dlls() -> CheckResult:
    root = _find_civil3d_root()
    if root is None:
        return CheckResult(
            name="Autodesk .NET DLLs",
            passed=False,
            detail="Civil 3D root not found (see previous check).",
        )
    # 2026-08-18 정정: DLL 이 한 폴더에 모여 있다고 가정했으나 실제로는 나뉘어 있다.
    # AeccDbMgd.dll 은 "AutoCAD 20xx\C3D\", acdbmgd.dll 은 "AutoCAD 20xx\" 에 있다.
    # client.py 의 _find_dll 처럼 후보 폴더 전체를 DLL 마다 훑는다.
    def _locate(dll: str) -> str | None:
        for cand in [root, *_CANDIDATE_ROOTS]:
            if Path(cand, dll).exists():
                return cand
        return None

    found = {dll: _locate(dll) for dll in _REQUIRED_DLLS + _OPTIONAL_DLLS}
    missing = [d for d in _REQUIRED_DLLS if found[d] is None]
    absent_optional = [d for d in _OPTIONAL_DLLS if found[d] is None]
    if missing:
        return CheckResult(
            name="Autodesk .NET DLLs",
            passed=False,
            detail=f"Missing (searched {len(_CANDIDATE_ROOTS)} folders): {', '.join(missing)}",
            fix_hint=(
                "Ensure a full Civil 3D installation is present, "
                "or set CIVIL3D_BIN_PATH to the correct folder."
            ),
        )
    where = {found[d] for d in _REQUIRED_DLLS}
    detail = "All required found in: " + " + ".join(sorted(where))
    if absent_optional:
        detail += f"  (optional, absent in 2025+: {', '.join(absent_optional)})"
    return CheckResult(
        name="Autodesk .NET DLLs",
        passed=True,
        detail=detail,
    )


# --------------- Civil 3D running ---------------

def check_civil3d_running() -> CheckResult:
    """Try GetActiveObject to verify Civil 3D is currently open."""
    try:
        import win32com.client as w32  # type: ignore
    except ImportError:
        return CheckResult(
            name="Civil 3D running",
            passed=False,
            detail="pywin32 not installed – cannot check.",
            fix_hint="Install pywin32 first.",
        )

    # 2026-08-18 정정: client.py 의 목록과 어긋나 있었다. 레지스트리(HKCR)로 실제
    # 확인된 값은 13.8(2026) / 13.7(2025) / 13.6(2024) 이며, 여기에는 그 셋이
    # 전부 빠져 있어 Civil 3D 가 떠 있어도 AutoCAD.Application 으로만 잡혔다.
    prog_ids = [
        "AeccXUiLand.AeccApplication.13.8",  # Civil 3D 2026
        "AeccXUiLand.AeccApplication.13.7",  # Civil 3D 2025
        "AeccXUiLand.AeccApplication.13.6",  # Civil 3D 2024
        "AeccXUiLand.AeccApplication.14.4",  # 원본 값(미검증) — 폴백용
        "AeccXUiLand.AeccApplication.14.0",  # 원본 값(미검증) — 폴백용
        "AeccXUiLand.AeccApplication.13.0",  # 원본 값(미검증) — 폴백용
        "AutoCAD.Application",
    ]
    for prog_id in prog_ids:
        try:
            app = w32.GetActiveObject(prog_id)
            doc = app.ActiveDocument
            name = getattr(doc, "Name", "(unknown)")
            return CheckResult(
                name="Civil 3D running",
                passed=True,
                detail=f"Connected via {prog_id} — active drawing: {name}",
            )
        except Exception:
            continue

    return CheckResult(
        name="Civil 3D running",
        passed=False,
        detail="Could not connect to a running Civil 3D / AutoCAD instance.",
        fix_hint=(
            "Open Civil 3D and load a drawing before starting the MCP server. "
            "This check is optional — the server will retry on first tool call."
        ),
    )


# --------------- Claude Desktop config ---------------

def check_claude_config() -> CheckResult:
    # 2026-08-18 정정: %APPDATA%\Claude 만 보고 있었다. Microsoft Store 판은
    # 샌드박스라 그 경로를 읽지 않는다(2026-08-17 실측). 두 곳을 모두 살핀다.
    _candidates = [
        Path(os.environ.get("APPDATA", ""), "Claude", "claude_desktop_config.json"),
        Path(os.environ.get("LOCALAPPDATA", ""), "Packages",
             "Claude_pzs8sxrjxfjjc", "LocalCache", "Roaming", "Claude",
             "claude_desktop_config.json"),
    ]
    config_path = next((c for c in _candidates if c.exists()), _candidates[0])
    if not config_path.exists():
        return CheckResult(
            name="Claude Desktop config",
            passed=False,
            detail=f"Not found at: {config_path}",
            fix_hint=(
                "Install Claude Desktop from https://claude.ai/download, "
                "then add the civil3d-mcp entry from claude_desktop_config_snippet.json."
            ),
        )
    try:
        # 2026-08-18 정정: 인코딩을 지정하지 않아 한글 Windows 에서 cp949 로 읽혀
        # "cp949 codec can't decode" 로 실패했다. 설정 파일은 UTF-8 이며, Claude
        # Desktop 이 BOM 을 붙여 저장하는 경우가 있어 utf-8-sig 로 읽는다.
        with open(config_path, encoding="utf-8-sig") as fh:
            cfg = json.load(fh)
        servers = cfg.get("mcpServers", {})
        if "civil3d-mcp" in servers:
            return CheckResult(
                name="Claude Desktop config",
                passed=True,
                detail=f"civil3d-mcp entry found in {config_path}",
            )
        return CheckResult(
            name="Claude Desktop config",
            passed=False,
            detail=f"civil3d-mcp entry missing from {config_path}",
            fix_hint=(
                "Add the block from claude_desktop_config_snippet.json "
                "to the mcpServers section and restart Claude Desktop."
            ),
        )
    except Exception as exc:
        return CheckResult(
            name="Claude Desktop config",
            passed=False,
            detail=f"Could not parse config: {exc}",
            fix_hint="Check that claude_desktop_config.json is valid JSON.",
        )


# ---------------------------------------------------------------------------
# Auto-fix
# ---------------------------------------------------------------------------

_PIP_PACKAGES = ["fastmcp", "pywin32", "pythonnet", "pydantic"]


def auto_fix(results: list[CheckResult]) -> None:
    failed_pip = [
        r for r in results
        if not r.passed and r.fix_hint.startswith("pip install")
    ]
    if not failed_pip:
        print("\n  Nothing to auto-fix via pip.")
        return
    for r in failed_pip:
        pkg = r.fix_hint.replace("pip install", "").split("(")[0].strip()
        print(f"\n  Installing: {pkg}")
        subprocess.run([sys.executable, "-m", "pip", "install", pkg], check=False)


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

CHECKS: list[Callable[[], CheckResult]] = [
    check_platform,
    check_python_version,
    check_python_arch,
    check_fastmcp,
    check_win32com,
    check_pythoncom,
    check_pythonnet,
    check_pydantic,
    check_civil3d_install,
    check_autodesk_dlls,
    check_civil3d_running,
    check_claude_config,
]

_PASS = "  [PASS]"
_FAIL = "  [FAIL]"
_WARN = "  [WARN]"


def run_checks(fix: bool = False, as_json: bool = False) -> int:
    # 2026-08-18 추가: 한글 Windows 콘솔은 기본이 cp949 라, 아래 머리글의 em-dash
    # 하나 때문에 UnicodeEncodeError 로 점검을 시작조차 못 하고 죽었다.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:                                      # noqa: BLE001
            pass

    results: list[CheckResult] = []

    if not as_json:
        print()
        print("=" * 60)
        print("  civil3d-mcp  —  environment check")
        print("=" * 60)

    for check_fn in CHECKS:
        result = check_fn()
        results.append(result)
        if not as_json:
            status = _PASS if result.passed else _FAIL
            print(f"\n{status}  {result.name}")
            print(f"        {result.detail}")
            if not result.passed and result.fix_hint:
                print(f"        → {result.fix_hint}")

    n_pass = sum(1 for r in results if r.passed)
    n_fail = len(results) - n_pass

    if as_json:
        print(json.dumps(
            {
                "summary": {"passed": n_pass, "failed": n_fail},
                "checks": [
                    {
                        "name": r.name,
                        "passed": r.passed,
                        "detail": r.detail,
                        "fix_hint": r.fix_hint,
                    }
                    for r in results
                ],
            },
            indent=2,
        ))
        return 0 if n_fail == 0 else 1

    print()
    print("=" * 60)
    print(f"  Result: {n_pass} passed, {n_fail} failed")
    print("=" * 60)

    # Civil 3D running is advisory — doesn't block
    hard_failures = [
        r for r in results
        if not r.passed and r.name != "Civil 3D running"
    ]

    if not hard_failures:
        print()
        print("  All required checks passed.")
        if any(not r.passed for r in results):
            print("  (Civil 3D running check is advisory — start Civil 3D before")
            print("   launching the MCP server.)")
        print()
        print("  Next steps:")
        print("  1. Open Civil 3D and load a drawing")
        print("  2. Start Claude Desktop")
        print("  3. Look for the hammer icon (🔨) in the toolbar")
        print()
    else:
        print()
        print("  Fix the failures above before running the server.")
        if fix:
            auto_fix(results)
        else:
            print("  Run with --fix to attempt automatic pip installs.")
        print()

    return 0 if not hard_failures else 1


def main() -> None:
    parser = argparse.ArgumentParser(
        description="civil3d-mcp pre-flight environment checker"
    )
    parser.add_argument(
        "--fix",
        action="store_true",
        help="Attempt auto-fixes (pip install missing packages)",
    )
    parser.add_argument(
        "--json",
        dest="as_json",
        action="store_true",
        help="Output results as JSON",
    )
    args = parser.parse_args()
    sys.exit(run_checks(fix=args.fix, as_json=args.as_json))


if __name__ == "__main__":
    main()
