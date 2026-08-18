# civil3d-mcp Non-Official from Autodesk. Only research

AI-powered Autodesk Civil 3D automation via **FastMCP** (Python 3.11).  
Connects Claude directly to a running Civil 3D instance using **COM automation** — no companion plugin required.

---

## How it works

```
Claude Desktop  ──stdio──►  FastMCP server (Python 3.11)
                               │
                               ├── win32com.client.GetActiveObject("AutoCAD.Application")
                               │       │
                               └── pythonnet clr → AeccDbMgd.dll / AeccLandMgd.dll
                                       │
                                       ▼
                               Civil 3D (acad.exe, running on same machine)
```

Python talks directly to the running Civil 3D process via the Windows COM server that Civil 3D registers automatically when it starts. `pythonnet` additionally loads Autodesk's .NET managed assemblies to access Civil 3D–specific objects (surfaces, alignments, COGO points) that aren't exposed through the base AutoCAD COM interface.

---

## Architecture

See **[ARCHITECTURE.md](ARCHITECTURE.md)** for the full system design:
layer breakdown, COM connection strategy, request lifecycle, threading model,
COM out-parameter handling, error strategy, and the Civil 3D object model reference.

---

## Requirements

| Requirement | Notes |
|---|---|
| Windows 10/11 | COM automation is Windows-only |
| Python 3.11 | 64-bit recommended |
| Autodesk Civil 3D | 2023, 2024, 2025 or 2026 — must be **open with a drawing loaded** |
| pip packages | `fastmcp`, `pythonnet`, `pywin32`, `pydantic` |

---

## Installation

```powershell
# 1. Clone
git clone https://github.com/KOR-GGG/Civil3D-mcp-python-COM.git civil3d-mcp
cd civil3d-mcp

# 2. Create the virtual environment - it MUST be named .venv
#    experiments/activate_env.py and experiments/harness_d.py hardcode
#    .venv\Scripts\, so any other name breaks the Chapter-5 harnesses.
py -3.11 -m venv .venv

# NOTE: do not bother with Activate.ps1 - corporate machines often block it via
# ExecutionPolicy. Call the interpreter directly instead, as every command below
# does. This is also the form used throughout experiments/.

# 3. Install dependencies (Windows, Python 3.11)
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt   # pytest + anthropic

# 4. Install the package in editable mode (creates .venv\Scripts\civil3d-mcp.exe)
.\.venv\Scripts\python.exe -m pip install -e .

# 5. Run the pre-flight environment checker
.\.venv\Scripts\python.exe setup_check.py
# "Civil 3D running" and "Claude Desktop config" may fail at this point - that is
# expected until you start Civil 3D (step 6) and register the server (below).

# 6. Regenerate the synthetic validation drawings (needs Civil 3D running,
#    with one empty drawing open). Skip only if you never run experiments/.
.\.venv\Scripts\python.exe make_test_surfaces.py A B C P --save-as test_surfaces.dwg
.\.venv\Scripts\python.exe make_test_surfaces.py E --new --save-as test_surfaces_env2.dwg
.\.venv\Scripts\python.exe make_test_surfaces.py N --new --save-as test_surfaces_env3.dwg
```

> **`.env` does not configure the server.** Nothing in `src/` calls `load_dotenv`;
> `client.py:68` reads the *process* environment. To point the server at a
> non-default Civil 3D folder, set `CIVIL3D_BIN_PATH` as a real environment
> variable or in the `env` block of the Claude Desktop config below.
> `.env` is read only by `experiments/harness_d.py` (for `ANTHROPIC_API_KEY`).

See **[INSTALL.md](INSTALL.md)** for the step-by-step version with verification
commands and the common failure modes.

---

## Configuration

### Claude Desktop

Config file location - **check which build you have** (`Get-AppxPackage *Claude*`):

| Build | Path |
|---|---|
| Normal installer | `%APPDATA%\Claude\claude_desktop_config.json` |
| **Microsoft Store** | `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json` |

The Store build is sandboxed and **does not read** `%APPDATA%\Claude` (verified 2026-08-17).

```json
{
  "mcpServers": {
    "civil3d-mcp": {
      "command": "C:\\Users\\YOURNAME\\source\\civil3d-mcp\\.venv\\Scripts\\civil3d-mcp.exe"
    }
  }
}
```

> **2026-08-18 correction - the absolute path is required.** This section previously
> offered `"command": "civil3d-mcp"` and a `"command": "python"` variant. **Neither
> works with the venv install above.** The bare name is not on `PATH` (the exe lives
> in `.venv\Scripts\` and we never activate the venv), and bare `python` resolves to
> the system interpreter, which has none of the dependencies. The absolute path to
> `.venv\Scripts\civil3d-mcp.exe` is the only form observed working.
>
> Print your own path with:
> ```powershell
> "$env:USERPROFILE\source\civil3d-mcp\.venv\Scripts\civil3d-mcp.exe"
> ```
> JSON does not expand environment variables, so this must be edited per machine.
> Escape every backslash (`\\`) inside JSON.

If Civil 3D is not in a default location, add `CIVIL3D_BIN_PATH` to the `env` block -
this is one of only two ways the server can actually receive it (see below):

```json
{
  "mcpServers": {
    "civil3d-mcp": {
      "command": "C:\\Users\\YOURNAME\\source\\civil3d-mcp\\.venv\\Scripts\\civil3d-mcp.exe",
      "env": { "CIVIL3D_BIN_PATH": "C:\\Program Files\\Autodesk\\AutoCAD 2026\\C3D" }
    }
  }
}
```

**Save without a BOM.** PowerShell's `Out-File -Encoding utf8` adds one; Claude Desktop
then fails to parse the JSON and **the server silently never starts**, with no error
shown anywhere in the UI. Use `Set-Content -Encoding utf8NoBOM`, or an editor that can
save UTF-8 without BOM.

Restart Claude Desktop, then confirm from the log rather than by hunting for an icon:

```powershell
Select-String "Server started and connected successfully" "$env:APPDATA\Claude\logs\mcp*.log" |
  Select-Object -Last 3
```

### Environment variables

| Variable | Default | Purpose | Read by |
|---|---|---|---|
| `CIVIL3D_BIN_PATH` | Auto-detected | Folder containing `AeccDbMgd.dll` (usually `...\AutoCAD 20xx\C3D`) | server, via the **process** environment |
| `LOG_LEVEL` | `INFO` | `DEBUG` / `INFO` / `WARNING` / `ERROR` | server |
| `ANTHROPIC_API_KEY` | - | Only for `experiments/harness_d.py` | that script, via `.env` |

> **The server never reads `.env`** (corrected 2026-08-18). Nothing under `src/` calls
> `load_dotenv`; `client.py:68` reads the process environment. So there are exactly two
> ways to give the server `CIVIL3D_BIN_PATH`:
> 1. a real environment variable - `$env:CIVIL3D_BIN_PATH = "..."` before launching, or a
>    user/system variable;
> 2. the `env` block of the Claude Desktop config (see above).
>
> `.env` is loaded only by `experiments/harness_d.py:49,53`, which is why
> `ANTHROPIC_API_KEY` belongs there and `CIVIL3D_BIN_PATH` effectively does not.

---

## Available tools (24 total)

### Drawing
| Tool | Description |
|---|---|
| `get_drawing_info` | Drawing name, path, save state, unit settings |
| `list_civil_object_types` | Count objects by type in model space |
| `get_selected_objects_info` | Properties of currently selected objects |

### COGO Points
| Tool | Description |
|---|---|
| `create_cogo_point` | Create a point at (northing, easting, elevation) STILL WIP!!!!|
| `list_cogo_points` | List all COGO points |
| `delete_cogo_point` | Delete a point by point number |

### Lines & Polylines
| Tool | Description |
|---|---|
| `create_line` | Create a 3D line from two points |
| `create_polyline` | Create a 3D polyline from a vertex list |
| `list_lines` | List all lines/polylines, optionally by layer |

### Surfaces
| Tool | Description |
|---|---|
| `list_surfaces` | List all TIN/Grid surfaces with elevation stats |
| `get_surface_info` | Detailed stats: point/triangle count, 2D/3D area |
| `sample_surface_elevation` | Sample elevation at (easting, northing) |
| `list_surface_definition` | List all definition items (boundaries, breaklines, contours, DEM files, etc.) |

### Earthwork
| Tool | Description |
|---|---|
| `compute_volume_between_surfaces` | Cut / fill / net volume between two TIN surfaces, optionally clipped to a closed polygon. Creates a temporary volume surface and erases it again |
| `compute_earthwork_by_rock_quality` | Cut volume split by rock quality (soil / ripping / blasting) and the resulting cost, from the existing ground, the design surface and the stratum boundary surfaces |

### Pipeline hydraulics
*Calculation only — these three do not read the drawing, so they work without Civil 3D running.*

| Tool | Description |
|---|---|
| `compute_head_loss` | Head loss, hydraulic gradient and velocity for one reach (Hazen-Williams). Diameter in mm, converted internally |
| `select_economic_diameter` | Diameter with the lowest construction cost + present value of pumping energy, among candidates inside the allowed velocity range. Returns the full candidate comparison and the reason each excluded diameter was dropped |
| `compute_hydraulic_profile` | Hydraulic grade line, residual head per station, and the reaches where pressurisation is needed |

### Alignments & Profiles
| Tool | Description |
|---|---|
| `list_alignments` | List all alignments with station range and length |
| `get_alignment_info` | Geometry breakdown: tangents, curves, spirals |
| `get_station_offset` | Project a world point to station + offset |
| `list_profiles` | List all profiles attached to an alignment |
| `get_profile_info` | Detailed info for a named profile |

### Corridors
| Tool | Description |
|---|---|
| `get_corridor_info` | Baselines, regions, and assembly tree for a corridor |

---

## Example prompts

```
"What surfaces are in this drawing and what are their elevation ranges?"

"Create a COGO point at northing 1000, easting 500, elevation 12.5 with description 'BM-01'"

"What is the elevation of the EG surface at easting 45200, northing 87300?"

"List all alignments and give me the geometry breakdown of CL-MAIN"

"What is the station and offset of point (45300, 87450) relative to alignment CL-MAIN?"

"Create a polyline on layer BOUNDARY with vertices at (0,0,0), (100,0,0), (100,100,0), (0,100,0) and close it"
```

---

## Project structure

```
civil3d-mcp-python/
├── src/
│   └── civil3d_mcp/
│       ├── __init__.py           # Public API: Civil3DClient, Civil3DError
│       ├── server.py             # FastMCP app, lifecycle, tool registration
│       ├── client.py             # Civil3DClient — all COM/pythonnet calls
│       ├── tools_drawing.py      # Tools 1-3:  drawing info
│       ├── tools_cogo.py         # Tools 4-6:  COGO points
│       ├── tools_lines.py        # Tools 7-9:  lines & polylines
│       ├── tools_surfaces.py     # Tools 10-13: surfaces
│       ├── tools_alignments.py  # Tools 14-18: alignments & profiles
│       ├── tools_corridors.py   # Tool  19:    corridors
│       ├── tools_earthwork.py   # Tools 20-21: earthwork volumes
│       ├── earthwork_core.py    # Earthwork logic, COM-free (unit-testable)
│       ├── tools_hydraulics.py  # Tools 22-24: pipeline hydraulics
│       └── hydraulics_core.py   # Hydraulics logic, COM-free (unit-testable)
├── test_earthwork_core.py        # pytest suite for earthwork_core (COM-free)
├── test_hydraulics_core.py       # pytest suite for hydraulics_core (COM-free)
├── make_test_surfaces.py         # Builds the synthetic _golden/ validation drawings
├── experiments/                  # Chapter-5 validation harnesses and result JSON
├── setup.py                      # Legacy setuptools entry point
├── setup_check.py                # Pre-flight environment checker
├── pyproject.toml                # PEP 517/518 build config
├── requirements.txt              # Runtime dependencies
├── requirements-dev.txt          # Dev/test dependencies
├── conftest.py                   # pytest sys.path injection
├── .env.example                  # Environment variable template
├── claude_desktop_config_snippet.json
├── ARCHITECTURE.md               # Full system design documentation
└── README.md
```

---

## Running tests

```bash
pip install -e ".[dev]"
pytest -v
```

There is no `tests/` directory - the suites live at the repository root as
`test_earthwork_core.py` and `test_hydraulics_core.py` (78 cases total).
Both exercise `earthwork_core` / `hydraulics_core`, which are COM-free, so no
Civil 3D installation is required to run them.

> Note: `probe_surf_api.py` and `probe_volume.py` are **probes, not tests**. They
> were renamed from `test_*` in commit `4f96e42` precisely because pytest collected
> them and they mutated the drawing that happened to be open.

---

## Troubleshooting

**"Could not connect to a running Civil 3D instance"**  
→ Make sure Civil 3D is open and a drawing is loaded before starting the MCP server.

**"Civil 3D .NET assemblies not found"**  
→ Set `CIVIL3D_BIN_PATH` in `.env` to the folder containing `AeccDbMgd.dll`  
→ Default path: `C:\Program Files\Autodesk\AutoCAD 2024`

**"Surfaces / Alignments collection not accessible"**  
→ Ensure you launched Civil 3D (not plain AutoCAD). The server falls back to AutoCAD's ProgID if the Civil 3D ProgID isn't registered.

**Hammer icon not showing in Claude Desktop**  
→ Check Claude Desktop logs: `%APPDATA%\Claude\logs\`  
→ Run `python -m civil3d_mcp.server` manually to see startup errors.

---

## License

MIT
