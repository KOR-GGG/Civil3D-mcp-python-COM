"""
server.py  –  civil3d-mcp FastMCP entry point
==============================================
Starts the MCP server, connects to Civil 3D via COM on startup,
and registers all tool groups.

Run via:
    python -m civil3d_mcp.server
    # or, after pip install -e .
    civil3d-mcp
"""
from __future__ import annotations

import asyncio
import functools
import logging
from contextlib import asynccontextmanager
from concurrent.futures import ThreadPoolExecutor
from typing import AsyncIterator, Callable, Any

from mcp.server.fastmcp import FastMCP

from .client import Civil3DClient, Civil3DError
from . import tools_drawing, tools_cogo, tools_lines, tools_surfaces, tools_alignments, tools_corridors

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("civil3d_mcp")

# ---------------------------------------------------------------------------
# Shared Civil 3D client  (one COM connection for the server lifetime)
# ---------------------------------------------------------------------------
client = Civil3DClient()


def _com_thread_init() -> None:
    """Initialize COM STA on the executor thread.

    pythoncom.CoInitialize() must be called on every thread that issues COM
    calls.  Without it the Windows COM runtime raises a 'thread marshalling'
    or RPC_E_WRONG_THREAD error when the asyncio event loop hands work to the
    background thread.
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
        log.debug("COM STA initialized on civil3d-com thread")
    except Exception as exc:  # noqa: BLE001
        log.warning("CoInitialize failed: %s", exc)


# Single-threaded executor: all COM calls are serialised on one STA thread.
_executor = ThreadPoolExecutor(
    max_workers=1,
    thread_name_prefix="civil3d-com",
    initializer=_com_thread_init,
)


async def run_com(fn: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
    """Run a blocking Civil3DClient call on the dedicated COM STA executor thread.

    All COM objects are apartment-affine — they must be accessed only from the
    thread that created them (the civil3d-com STA thread).  Calling them
    directly from the asyncio event-loop thread causes RPC_E_WRONG_THREAD
    (-2147417842).  This helper ensures every client call is dispatched to the
    correct thread.
    """
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(_executor, functools.partial(fn, *args, **kwargs))


# ---------------------------------------------------------------------------
# Lifespan context manager (replaces on_startup / on_shutdown)
# ---------------------------------------------------------------------------
@asynccontextmanager
async def lifespan(server: FastMCP) -> AsyncIterator[None]:
    log.info("civil3d-mcp starting — connecting to Civil 3D via COM…")
    loop = asyncio.get_event_loop()
    try:
        await loop.run_in_executor(_executor, client.connect)
        log.info("civil3d-mcp ready  |  19 tools registered")
    except Civil3DError as exc:
        log.error("Could not connect to Civil 3D: %s", exc)
        log.error(
            "Make sure Civil 3D is open with a drawing loaded, "
            "then restart the MCP server."
        )
    try:
        yield
    finally:
        log.info("civil3d-mcp shutting down")
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(_executor, client.disconnect)
        _executor.shutdown(wait=False)


# ---------------------------------------------------------------------------
# FastMCP application
# ---------------------------------------------------------------------------
mcp = FastMCP(
    name="civil3d-mcp",
    instructions=(
        "AI-powered Autodesk Civil 3D automation via COM (pythonnet + win32com). "
        "Supports drawing queries, COGO points, lines/polylines, surfaces "
        "and alignments. Civil 3D must be running on the same Windows machine."
    ),
    lifespan=lifespan,
)

# ---------------------------------------------------------------------------
# Register all tool groups  (pass run_com so tools dispatch via executor)
# ---------------------------------------------------------------------------
tools_drawing.register(mcp, client, run_com)
tools_cogo.register(mcp, client, run_com)
tools_lines.register(mcp, client, run_com)
tools_surfaces.register(mcp, client, run_com)
tools_alignments.register(mcp, client, run_com)
tools_corridors.register(mcp, client, run_com)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
