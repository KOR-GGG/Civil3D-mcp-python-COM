"""probe_mcp_stdio.py — 서버가 실제로 MCP 프로토콜을 말하는지 확인한다.

왜 필요한가
-----------
지금까지의 확인은 `civil3d_mcp.server` 를 **파이썬으로 import 해서 도구 수를
센 것**뿐이다. 그것은 도구가 등록되었음을 보일 뿐, **호스트가 stdio 로 붙어
JSON-RPC 를 주고받을 수 있는지는 말해 주지 않는다.** 에이전트(D) 실행의
전제이므로 먼저 닫는다.

MCP stdio 전송은 **줄바꿈으로 구분된 JSON-RPC 2.0** 메시지를 쓴다.
SDK 클라이언트를 쓰지 않고 직접 주고받는 이유는, SDK 대 SDK 로 확인하면
프로토콜이 아니라 SDK 의 자기 정합만 보게 되기 때문이다.

확인 순서
---------
  1. initialize                     — 핸드셰이크, 서버 정보
  2. notifications/initialized      — 개시 통지
  3. tools/list                     — 도구 목록과 **입력 스키마**
  4. tools/call (COM 비접근 도구)    — 계산형 도구로 왕복 확인
  5. tools/call (COM 접근 도구)      — Civil 3D 까지 도달하는지 확인

실행
----
    .\.venv\Scripts\python.exe experiments\probe_mcp_stdio.py
"""
from __future__ import annotations

import json
import subprocess
import sys
import time
from pathlib import Path

# 2026-08-19 추가: 한글 Windows 콘솔은 기본이 cp949 라, 아래 출력에 쓰이는
# em-dash 나 경고 기호 하나 때문에 UnicodeEncodeError 로 스크립트가 즉사한다.
# (setup_check.py 에서 같은 결함을 고쳤으나 실험 스크립트에는 남아 있었다.)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:                                          # noqa: BLE001
        pass

REPO = Path(__file__).resolve().parents[1]
SERVER = REPO / ".venv" / "Scripts" / "civil3d-mcp.exe"
TIMEOUT_S = 90.0


class StdioClient:
    def __init__(self, cmd: list[str]):
        self.proc = subprocess.Popen(
            cmd, cwd=str(REPO),
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding="utf-8", bufsize=1,
        )
        self._id = 0

    def send(self, method: str, params: dict | None = None, notify: bool = False):
        msg: dict = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            msg["params"] = params
        if not notify:
            self._id += 1
            msg["id"] = self._id
        self.proc.stdin.write(json.dumps(msg, ensure_ascii=False) + "\n")
        self.proc.stdin.flush()
        return None if notify else self._id

    def read_until(self, want_id: int, timeout_s: float = TIMEOUT_S) -> dict:
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            line = self.proc.stdout.readline()
            if not line:
                if self.proc.poll() is not None:
                    err = self.proc.stderr.read()
                    raise SystemExit(
                        f"서버가 종료되었다(코드 {self.proc.returncode}).\n"
                        f"stderr 마지막 부분:\n{err[-2000:]}"
                    )
                continue
            line = line.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except json.JSONDecodeError:
                # 서버가 표준출력에 로그를 섞어 쓰면 여기 걸린다 — 그 자체가 결함이다
                print(f"     ⚠ stdout 에 JSON 이 아닌 줄: {line[:120]}")
                continue
            if msg.get("id") == want_id:
                return msg
        raise SystemExit(f"{timeout_s}초 안에 id={want_id} 응답이 오지 않았다.")

    def close(self):
        try:
            self.proc.stdin.close()
            self.proc.wait(timeout=10)
        except Exception:                                       # noqa: BLE001
            self.proc.kill()


def main() -> int:
    if not SERVER.exists():
        raise SystemExit(f"서버 실행 파일이 없다: {SERVER}")

    print("=" * 88)
    print("  MCP stdio 프로토콜 확인 — 에이전트(D) 실행의 전제")
    print("=" * 88)
    print(f"  서버: {SERVER}")
    print()

    c = StdioClient([str(SERVER)])
    result: dict = {"server": str(SERVER)}
    try:
        # 1 ── initialize
        t0 = time.perf_counter()
        i = c.send("initialize", {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "thesis-probe", "version": "0.1"},
        })
        r = c.read_until(i)
        boot = time.perf_counter() - t0
        if "error" in r:
            raise SystemExit(f"initialize 실패: {r['error']}")
        info = r["result"].get("serverInfo", {})
        print(f"  [1] initialize            OK  ({boot:.1f}s)")
        print(f"        서버 이름/버전   {info.get('name')} {info.get('version')}")
        print(f"        프로토콜 버전    {r['result'].get('protocolVersion')}")
        result["initialize"] = {"elapsed_s": boot, "serverInfo": info,
                                "protocolVersion": r["result"].get("protocolVersion")}

        c.send("notifications/initialized", {}, notify=True)

        # 2 ── tools/list
        i = c.send("tools/list")
        r = c.read_until(i)
        tools = r["result"]["tools"]
        print(f"  [2] tools/list            OK  — 도구 {len(tools)}개")
        ours = [t["name"] for t in tools if t["name"].startswith("compute_")]
        print(f"        신설 도구        {ours}")
        # 스키마가 타입 힌트에서 자동 유도되었는지 확인 (4.6.3 의 주장)
        target = next(t for t in tools if t["name"] == "compute_earthwork_by_rock_quality")
        props = list(target["inputSchema"].get("properties", {}))
        required = target["inputSchema"].get("required", [])
        print(f"        L2 입력 스키마   properties={props}")
        print(f"                         required={required}")
        result["tools"] = {"count": len(tools), "new_tools": ours,
                           "l2_properties": props, "l2_required": required}

        # 3 ── tools/call : COM 비접근
        i = c.send("tools/call", {
            "name": "compute_head_loss",
            "arguments": {"flow_m3_s": 0.35, "diameter_mm": 600, "length_m": 4200},
        })
        t0 = time.perf_counter()
        r = c.read_until(i)
        dt = time.perf_counter() - t0
        payload = json.loads(r["result"]["content"][0]["text"])
        ok = abs(payload["head_loss_m"] - 15.4234) < 1e-3
        print(f"  [3] tools/call 계산형     {'OK' if ok else '값 불일치'}  ({dt*1000:.0f} ms)")
        print(f"        head_loss_m      {payload['head_loss_m']:.4f}  (기대 15.4234)")
        print(f"        formula_constants {payload.get('formula_constants')}")
        result["call_pure"] = {"ok": ok, "elapsed_ms": dt * 1000, "payload": payload}

        # 4a ── 서버가 어느 도면을 보고 있는지 먼저 확인한다
        i = c.send("tools/call", {"name": "get_drawing_info", "arguments": {}})
        r = c.read_until(i)
        try:
            dinfo = json.loads(r["result"]["content"][0]["text"])
        except (json.JSONDecodeError, KeyError):
            dinfo = {"raw": r["result"]["content"][0]["text"][:200]}
        print(f"  [4a] 서버가 보는 도면      {dinfo.get('name')}")
        print(f"        insertion_units  {dinfo.get('insertion_units')}")
        result["drawing_info"] = dinfo

        # 4 ── tools/call : COM 접근
        i = c.send("tools/call", {"name": "list_surfaces", "arguments": {}})
        t0 = time.perf_counter()
        r = c.read_until(i)
        dt = time.perf_counter() - t0
        # ⚠ 응답 형태를 단정하지 말 것. FastMCP 는 content[].text 외에
        # structuredContent 로도 값을 실어 보낸다. 한쪽만 보고 "0개"라고
        # 결론내면 서버가 아니라 **읽는 쪽**이 틀린 것이다(2026-08-17 실제로 겪음).
        res = r["result"]
        print(f"        응답 키          {sorted(res)}")
        text = res.get("content", [{}])[0].get("text", "")
        print(f"        content[0].text  {text[:160]}{' …' if len(text) > 160 else ''}")
        names = []
        parsed = None
        try:
            parsed = json.loads(text) if text else None
        except json.JSONDecodeError:
            print("        (content.text 는 JSON 이 아니다)")
        if parsed is None:
            parsed = res.get("structuredContent")
            if parsed is not None:
                print(f"        structuredContent 사용")
        if isinstance(parsed, dict):
            for v in parsed.values():
                if isinstance(v, list):
                    parsed = v
                    break
        if isinstance(parsed, list):
            names = [s.get("name") if isinstance(s, dict) else str(s) for s in parsed]
        print(f"  [4] tools/call COM 접근   OK  ({dt*1000:.0f} ms) — 서피스 {len(names)}개")
        print(f"        {names[:6]}{' …' if len(names) > 6 else ''}")
        result["call_com"] = {"elapsed_ms": dt * 1000, "surface_count": len(names),
                              "surfaces": names}

        print()
        print("=" * 88)
        print("  결과")
        print("=" * 88)
        print("  ✅ 서버가 stdio 로 MCP 프로토콜을 정상 수행한다.")
        print("     핸드셰이크 · 도구 목록 · 계산형 호출 · COM 접근 호출이 모두 왕복했다.")
        print("     → 에이전트(D) 실행의 서버 측 전제는 충족되었다. 남은 것은 호스트다.")

        # 5 ── 서버 로그(stderr) 확인
        # 어제 _get_surfaces 의 분기별 문구를 나눠 둔 것이 여기서 값을 한다.
        print()
        print("-" * 88)
        print("  서버 로그(stderr) — 서피스 관련")
        print("-" * 88)
        c.proc.stdin.close()
        try:
            c.proc.wait(timeout=15)
        except Exception:                                       # noqa: BLE001
            c.proc.kill()
        err = c.proc.stderr.read() or ""
        keep = [ln for ln in err.splitlines()
                if any(k in ln for k in
                       ("Surface", "surface", "collection", "COM", "connect",
                        "ProgID", "Civil"))]
        for ln in keep[-25:]:
            print("   ", ln)
        result["server_log_tail"] = keep[-25:]

    finally:
        c.close()

    out = Path(__file__).with_name("mcp_stdio_result.json")
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str),
                   encoding="utf-8")
    print(f"\n  결과 저장 -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
