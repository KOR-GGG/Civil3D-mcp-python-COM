"""harness_d.py — 비교군 D(에이전트) 실행 하니스

무엇인가
--------
`PROTOCOL_D.md` 를 자동으로 수행한다. MCP 서버에 stdio 로 붙고, Anthropic API
로 에이전트를 돌리고, 시행마다 §3.4 의 기록 항목을 남긴다.

왜 하니스인가
-------------
수동 실행(Claude Desktop)은 n=5 반복의 **기록이 부실해진다.** 워크플로 전체
시간을 벽시계로 재야 하고(§4.2), 도구 호출 순서·횟수를 시행마다 남겨야 하며
(§3.4), 재현성은 같은 요청을 n회 반복해 변동계수를 봐야 한다(§4.6.2 ⓑ).
전부 스크립트가 정확하다.

★ 측정의 공정성
---------------
이 하니스를 **본 연구를 도운 대화 세션이 작성했다는 사실은 측정을 오염시키지
않는다.** 에이전트 역할은 API 뒤의 모델이 수행하고, 그 모델에게 전달되는
것은 아래 PROMPT 문자열과 서버가 내려주는 도구 스키마뿐이다. **프롬프트에
서피스 이름이 없으므로**(§3.2) 작성자가 아는 정답이 모델에게 새지 않는다.

⚠ 그러므로 **PROMPT 를 편집하지 말 것.** 힌트를 한 줄 더하는 순간 C5 는
측정이 아니라 연출이 된다.

수동 루프를 쓰는 이유
--------------------
도구가 **실행 시점에 서버로부터 발견**되므로 데코레이터 기반 tool_runner 에
정적으로 등록할 수 없고, 시행마다 호출 순서·개별 소요시간을 남겨야 한다.

실행
----
    .\.venv\Scripts\python.exe experiments\harness_d.py            # 전 환경 n=5
    .\.venv\Scripts\python.exe experiments\harness_d.py --n 1      # 예행 1회
    .\.venv\Scripts\python.exe experiments\harness_d.py --env 2    # 환경 2만
"""
from __future__ import annotations

import argparse
import json
import os
import statistics
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import win32com.client as w32
from dotenv import load_dotenv

REPO = Path(__file__).resolve().parents[1]
SERVER = REPO / ".venv" / "Scripts" / "civil3d-mcp.exe"
load_dotenv(REPO / ".env", override=True)

import anthropic  # noqa: E402  (키를 먼저 로드해야 한다)

# ---------------------------------------------------------------------------
# 고정 설정 — §4.6.2 ⓑ 가 기록을 요구하는 항목
# ---------------------------------------------------------------------------
MODEL = "claude-opus-5"
MAX_TOKENS = 16_000
THINKING = {"type": "adaptive", "display": "summarized"}
EFFORT = "high"
# ⚠ Claude Opus 5 는 temperature/top_p/top_k 를 받지 않는다(400). 샘플링
#    파라미터가 없다는 사실 자체가 재현성 조건의 일부이므로 기록해 둔다.
SAMPLING = "none (Opus 5 rejects temperature/top_p/top_k)"

MAX_ROUNDS = 30          # 도구 호출 왕복 상한 (폭주 방지)
MAX_RETRY_PER_TOOL = 3   # §4.6.3 재시도 상한

SYSTEM = (
    "You are assisting a civil engineer inside Autodesk Civil 3D. "
    "Use the available MCP tools to inspect the drawing and perform the "
    "requested engineering computation. Base every number you report on an "
    "actual tool result — do not estimate. If the drawing does not contain "
    "what the task requires, say so plainly instead of guessing."
)

# ★ PROTOCOL_D.md §3.2 의 프롬프트. 두 환경에 한 글자도 다르지 않게 쓴다.
#   서피스 이름을 넣지 말 것 — 식별 자체가 측정 대상이다.
PROMPT = (
    "열려 있는 Civil 3D 도면에서 계획부지 조성에 필요한 암질별 토공량을 "
    "산정해 주세요.\n"
    "토사·리핑·발파로 구분한 물량과 공사비를 알려 주시고, 어떤 서피스를 "
    "원지반·계획고·지층 경계로 보았는지도 함께 밝혀 주세요.\n"
    "단가는 토사 4,500원/㎥, 풍화암 9,000원/㎥, 연암 18,000원/㎥, "
    "최하층 하부(경암) 25,000원/㎥ 입니다."
)

# 환경 식별 — 그 도면에만 있는 서피스 이름(하니스 내부용, 모델에게 가지 않음)
ENVIRONMENTS = {
    1: {"label": "환경 1 (TEST_C_* 명명)", "marker": "TEST_C_S0_GROUND"},
    2: {"label": "환경 2 (원지형/계획부지01 명명 + 미끼)", "marker": "원지형"},
}

# 정답 (두 환경 동일) — 판정용
EXPECTED_TOTAL_CUT = 120_000.0
EXPECTED_BY_METHOD = {"토사": 50_000.0, "리핑": 30_000.0, "발파": 40_000.0}
TOL = 0.5   # ㎥


# ---------------------------------------------------------------------------
# MCP stdio 브리지
# ---------------------------------------------------------------------------
class MCPBridge:
    """MCP 서버를 stdio 로 띄우고 tools/list · tools/call 을 중계한다."""

    def __init__(self) -> None:
        self.proc = subprocess.Popen(
            [str(SERVER)], cwd=str(REPO),
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding="utf-8", bufsize=1,
        )
        self._id = 0
        self.tools: list[dict] = []

    def _send(self, method: str, params: dict | None = None, notify: bool = False):
        msg: dict = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            msg["params"] = params
        if not notify:
            self._id += 1
            msg["id"] = self._id
        self.proc.stdin.write(json.dumps(msg, ensure_ascii=False) + "\n")
        self.proc.stdin.flush()
        return None if notify else self._id

    def _read(self, want_id: int, timeout_s: float = 180.0) -> dict:
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            line = self.proc.stdout.readline()
            if not line:
                if self.proc.poll() is not None:
                    raise RuntimeError(
                        f"MCP 서버 종료(코드 {self.proc.returncode}):\n"
                        f"{(self.proc.stderr.read() or '')[-1500:]}"
                    )
                continue
            line = line.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except json.JSONDecodeError:
                continue          # 서버가 stdout 에 섞어 쓴 로그
            if msg.get("id") == want_id:
                return msg
        raise TimeoutError(f"{timeout_s}s 안에 id={want_id} 응답 없음")

    def start(self) -> list[dict]:
        r = self._read(self._send("initialize", {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "harness-d", "version": "1.0"},
        }))
        if "error" in r:
            raise RuntimeError(f"initialize 실패: {r['error']}")
        self._send("notifications/initialized", {}, notify=True)
        self.tools = self._read(self._send("tools/list"))["result"]["tools"]
        return self.tools

    def call(self, name: str, arguments: dict) -> tuple[str, bool]:
        """(결과 텍스트, 오류 여부)

        ⚠ 응답 형태를 단정하지 말 것. FastMCP 는 content[].text 외에
        structuredContent 로도 값을 싣는다(2026-08-17 실제로 겪은 오류).
        """
        r = self._read(self._send("tools/call",
                                  {"name": name, "arguments": arguments}))
        if "error" in r:
            return json.dumps(r["error"], ensure_ascii=False), True
        res = r["result"]
        parts = [c.get("text", "") for c in res.get("content", [])
                 if c.get("type") == "text"]
        text = "\n".join(p for p in parts if p)
        if not text and res.get("structuredContent") is not None:
            text = json.dumps(res["structuredContent"], ensure_ascii=False)
        return text or "(빈 응답)", bool(res.get("isError"))

    def stop(self) -> str:
        try:
            self.proc.stdin.close()
            self.proc.wait(timeout=10)
        except Exception:                                       # noqa: BLE001
            self.proc.kill()
        try:
            return self.proc.stderr.read() or ""
        except Exception:                                       # noqa: BLE001
            return ""


def to_anthropic_tools(mcp_tools: list[dict]) -> list[dict]:
    """MCP 도구 정의를 Anthropic tools 형식으로. 스키마는 서버 것을 그대로 쓴다."""
    out = []
    for t in mcp_tools:
        schema = t.get("inputSchema") or {"type": "object", "properties": {}}
        out.append({
            "name": t["name"],
            "description": (t.get("description") or t["name"])[:1900],
            "input_schema": schema,
        })
    return out


# ---------------------------------------------------------------------------
# 도면 전환
# ---------------------------------------------------------------------------
def activate_drawing(marker: str) -> str:
    """marker 서피스를 가진 도면을 활성화하고 그 이름을 돌려준다."""
    from civil3d_mcp.client import Civil3DClient

    acad = w32.GetActiveObject("AutoCAD.Application")
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
                if marker in [s["name"] for s in c.list_surfaces()]:
                    return str(doc.Name)
                break
            except Exception:                                   # noqa: BLE001
                time.sleep(0.25)
    raise SystemExit(f"'{marker}' 을 가진 도면을 찾지 못했다.")


# ---------------------------------------------------------------------------
# 한 시행
# ---------------------------------------------------------------------------
def run_trial(client: anthropic.Anthropic, bridge: MCPBridge,
              tools: list[dict], trial_no: int) -> dict:
    messages: list[dict] = [{"role": "user", "content": PROMPT}]
    trace: list[dict] = []          # 도구 호출 순서·인자·시간
    retry_count: dict[str, int] = {}
    usage_in = usage_out = 0
    wall0 = time.perf_counter()
    stop_why = "완주"
    resp = None

    for rnd in range(MAX_ROUNDS):
        try:
            with client.messages.stream(
                model=MODEL, max_tokens=MAX_TOKENS,
                system=SYSTEM, thinking=THINKING,
                output_config={"effort": EFFORT},
                tools=tools, messages=messages,
            ) as stream:
                resp = stream.get_final_message()
        except anthropic.APIStatusError as exc:
            stop_why = f"API 오류 {exc.status_code}"
            break

        usage_in += resp.usage.input_tokens
        usage_out += resp.usage.output_tokens
        messages.append({"role": "assistant", "content": resp.content})

        if resp.stop_reason == "refusal":
            stop_why = "모델 거부"
            break
        if resp.stop_reason != "tool_use":
            break                                   # 최종 답변

        results = []
        for blk in resp.content:
            if blk.type != "tool_use":
                continue
            retry_count[blk.name] = retry_count.get(blk.name, 0) + 1
            t0 = time.perf_counter()
            try:
                text, is_err = bridge.call(blk.name, dict(blk.input))
            except Exception as exc:                            # noqa: BLE001
                text, is_err = f"{type(exc).__name__}: {exc}", True
            dt = (time.perf_counter() - t0) * 1000
            entry = {"round": rnd + 1, "tool": blk.name,
                     "arguments": dict(blk.input), "error": is_err,
                     "elapsed_ms": dt, "result_head": text[:400]}
            # ⚠ 채점은 도구가 낸 값으로 한다. result_head 는 로그용으로 자르되,
            #    판정 대상 도구의 응답은 자르지 않고 따로 보관해야 한다
            #    (자른 문자열을 파싱하면 판정이 늘 "보류"로 떨어진다).
            if blk.name == "compute_earthwork_by_rock_quality" and not is_err:
                entry["result_full"] = text
            trace.append(entry)
            results.append({"type": "tool_result", "tool_use_id": blk.id,
                            "content": text[:60_000], "is_error": is_err})
        messages.append({"role": "user", "content": results})

        over = [k for k, v in retry_count.items() if v > MAX_RETRY_PER_TOOL]
        if over:
            stop_why = f"재시도 상한 초과: {over}"
            break
    else:
        stop_why = f"왕복 상한({MAX_ROUNDS}) 도달"

    wall = time.perf_counter() - wall0
    final_text = ("\n".join(b.text for b in resp.content if b.type == "text")
                  if resp is not None else "")

    return {"trial": trial_no, "stop_why": stop_why, "wall_s": wall,
            "rounds": len(set(t["round"] for t in trace)),
            "tool_calls": len(trace), "trace": trace,
            "final_text": final_text,
            "usage": {"input_tokens": usage_in, "output_tokens": usage_out},
            **score(trace)}


def score(trace: list[dict]) -> dict:
    """도구가 실제로 낸 값으로 판정한다 — 최종 문장의 숫자를 믿지 않는다."""
    earth = [t for t in trace
             if t["tool"] == "compute_earthwork_by_rock_quality" and not t["error"]]
    used_decoy = any("(점)" in json.dumps(t["arguments"], ensure_ascii=False)
                     for t in trace)
    if not earth:
        return {"completed": False, "correct": None, "used_decoy": used_decoy,
                "computed": None}

    # 마지막 성공 호출을 채점 대상으로 (자르지 않은 result_full 을 쓴다)
    try:
        payload = json.loads(earth[-1].get("result_full", ""))
    except json.JSONDecodeError:
        payload = None
    if not isinstance(payload, dict) or "total_cut_m3" not in payload:
        return {"completed": True, "correct": None, "used_decoy": used_decoy,
                "computed": "판정 보류(응답을 파싱하지 못함)"}

    total = payload.get("total_cut_m3")
    by_method = {b["method"]: b["cut_m3"] for b in payload.get("by_method", [])}
    ok = (total is not None and abs(total - EXPECTED_TOTAL_CUT) < TOL and
          all(abs(by_method.get(m, -1) - v) < TOL
              for m, v in EXPECTED_BY_METHOD.items()))
    return {"completed": True, "correct": ok, "used_decoy": used_decoy,
            "computed": {"total_cut_m3": total, "by_method": by_method,
                         "total_cost": payload.get("total_earthwork_cost")}}


# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=5, help="환경별 반복 횟수")
    ap.add_argument("--env", type=int, choices=[1, 2], help="한 환경만")
    args = ap.parse_args()

    if not os.getenv("ANTHROPIC_API_KEY"):
        raise SystemExit("ANTHROPIC_API_KEY 가 없다. .env 를 확인할 것.")

    client = anthropic.Anthropic()
    envs = [args.env] if args.env else [1, 2]

    print("=" * 92)
    print("  비교군 D(에이전트) 실행 — PROTOCOL_D.md")
    print("=" * 92)
    print(f"  모델 {MODEL} · thinking {THINKING['type']}/{THINKING['display']} · "
          f"effort {EFFORT} · max_tokens {MAX_TOKENS}")
    print(f"  샘플링 파라미터: {SAMPLING}")
    print(f"  반복 n={args.n} · 왕복 상한 {MAX_ROUNDS} · 재시도 상한 {MAX_RETRY_PER_TOOL}")
    print()

    out: dict[str, Any] = {
        "config": {"model": MODEL, "thinking": THINKING, "effort": EFFORT,
                   "max_tokens": MAX_TOKENS, "sampling": SAMPLING,
                   "system": SYSTEM, "prompt": PROMPT, "n": args.n},
        "environments": {},
    }

    for env_no in envs:
        env = ENVIRONMENTS[env_no]
        print("-" * 92)
        print(f"  {env['label']}")
        print("-" * 92)
        drawing = activate_drawing(env["marker"])
        print(f"  활성 도면: {drawing}")

        trials = []
        for i in range(1, args.n + 1):
            # 시행마다 서버를 새로 띄운다 — 기동 시점에 활성 도면에 붙기 때문
            bridge = MCPBridge()
            try:
                mcp_tools = bridge.start()
                if i == 1:
                    print(f"  MCP 도구 {len(mcp_tools)}개 인식")
                r = run_trial(client, bridge, to_anthropic_tools(mcp_tools), i)
            finally:
                bridge.stop()
            trials.append(r)

            mark = "완주" if r["completed"] else "미완주"
            acc = {True: "정확", False: "부정확", None: "―"}[r["correct"]]
            print(f"   [{i}/{args.n}] {mark} {acc} · {r['wall_s']:6.1f}s · "
                  f"도구 {r['tool_calls']}회 · {r['stop_why']}"
                  + ("  ⚠미끼사용" if r["used_decoy"] else ""))

        done = [t for t in trials if t["completed"]]
        good = [t for t in trials if t["correct"] is True]
        walls = [t["wall_s"] for t in trials]
        cv = (statistics.stdev(walls) / statistics.fmean(walls) * 100
              if len(walls) > 1 else 0.0)
        print(f"   => 완주 {len(done)}/{len(trials)} · 값 정확 {len(good)}/{len(trials)}"
              f" · 벽시계 평균 {statistics.fmean(walls):.1f}s (CV {cv:.1f}%)")
        print()

        out["environments"][str(env_no)] = {
            "label": env["label"], "drawing": drawing, "trials": trials,
            "summary": {"completed": len(done), "correct": len(good),
                        "n": len(trials), "wall_mean_s": statistics.fmean(walls),
                        "wall_cv_pct": cv},
        }

    dst = Path(__file__).with_name("d_result.json")
    dst.write_text(json.dumps(out, ensure_ascii=False, indent=2, default=str),
                   encoding="utf-8")
    print("=" * 92)
    print(f"  결과 저장 -> {dst}")
    print("  ⚠ 판정 시 §5(PROTOCOL_D.md)의 세 원칙을 지킬 것.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
