# INSTALL.md — 새 PC 설치 절차

> 대상: Windows(한글 로케일 포함) · Civil 3D 2024/2025/2026 중 하나 · 아무것도 깔려 있지 않은 상태
> 목표: `experiments\activate_env.py 3` 이 끝까지 도는 상태
> 최초 작성 2026-08-18 — 사무실 PC 에서 실측하며 정리했다.

## 이 문서를 만든 이유

이 저장소는 오래 **한 대의 PC 에서만** 돌았다. 두 번째 PC 로 옮기려 하자 문서와 코드가
갈라져 있던 지점이 한꺼번에 드러났다 — 점검기가 한글 Windows 에서 첫 줄부터 죽고,
README 의 clone URL 이 404 이고, 도면 생성 스크립트에는 저장 기능이 아예 없었다.
여기 적힌 것은 그 정리의 결과다.

**세 원칙이 이 문서 전체를 관통한다.**

1. **venv 를 activate 하지 않는다.** 항상 `.\.venv\Scripts\python.exe` 로 인터프리터를
   명시한다. 회사 PC 는 ExecutionPolicy 로 `Activate.ps1` 을 막는 경우가 많고,
   저장소의 실행 명령이 이미 전부 이 형태다.
2. **venv 이름은 반드시 `.venv`.** `experiments/activate_env.py` 와
   `experiments/harness_d.py` 가 하드코딩한다. 다른 이름이면 **늦게** 실패한다 —
   도면을 열고 환경 확인까지 통과한 뒤 표식 심기에서 `FileNotFoundError` 로 죽는다.
3. **`install.py` 를 그냥 실행하지 않는다.** venv 를 만들지도 검출하지도 않고
   `sys.executable` 로 설치하므로, 일반 셸에서 돌리면 시스템 Python 이 오염된다.

---

## 단계 요약

| # | 하는 일 | Civil 3D 필요 |
|---|---|---|
| 0 | 현황 파악 | 아니오 |
| 1 | Python 3.11 (64-bit) | 아니오 |
| 2 | clone | 아니오 |
| 3 | `.venv` 생성 | 아니오 |
| 4 | 의존성 설치 | 아니오 |
| 5 | editable 설치 | 아니오 |
| 6 | `setup_check.py` | 아니오 |
| 7 | 단위시험 78건 | 아니오 |
| 8 | **`_golden` 도면 3개 생성** ← 새 clone 의 진짜 관문 | **예** |
| 9 | 환경 전환 확인 | 예 |
| 10 | Claude Desktop 등록 | 예 |

---

## 0단계 — 설치 전에 현황부터 찍는다

```powershell
$env:PYTHONIOENCODING = "utf-8"      # 이 창에서만 유효

Get-ExecutionPolicy -List
py -0p
Get-Command git, winget -ErrorAction SilentlyContinue | Select-Object Name, Source
Get-ChildItem "C:\Program Files\Autodesk" -Directory -ErrorAction SilentlyContinue | Select-Object Name
Get-AppxPackage *Claude* | Select-Object Name, Version
Get-ChildItem Registry::HKEY_CLASSES_ROOT |
  Where-Object PSChildName -like "AeccXUiLand.AeccApplication.*" |
  Select-Object -ExpandProperty PSChildName
```

적어 둘 것 — `-V:3.11` 유무 / `MachinePolicy`·`UserPolicy` 가 걸려 있는지 /
`AutoCAD 20xx` 폴더 / **Store 판 Claude 인지**(`Get-AppxPackage` 결과가 있으면 Store 판) /
등록된 ProgID.

ProgID 대응은 `src/civil3d_mcp/client.py` 기준 **13.8 = 2026 · 13.7 = 2025 · 13.6 = 2024**
이며 셋 다 레지스트리로 검증된 값이다. 서버가 이 셋을 자동으로 시도하므로 버전을 몰라도 된다.

---

## 1단계 — Python 3.11 (64-bit)

`py -0p` 에 `-V:3.11` 이 있으면 건너뛴다.

```powershell
winget install --id Python.Python.3.11 -e --scope user `
  --accept-source-agreements --accept-package-agreements
```

winget 이 없거나 막히면 python.org 설치 파일로 사용자 설치한다
(`/passive InstallAllUsers=0 PrependPath=1 Include_launcher=1`).

**확인**

```powershell
py -3.11 -V                                          # -> Python 3.11.x
py -3.11 -c "import sys; print(sys.maxsize > 2**32)" # -> True
```

> `pyproject.toml` 은 `>=3.11` 로 상한이 없지만 실측 검증된 조합은 3.11.9 하나다.
> **절대경로로 venv 를 만들지 말 것** — `%LOCALAPPDATA%\Programs\Python\Python311` 은
> 사용자 설치일 때의 기본값일 뿐이고, IT 가 머신 설치를 했으면 `C:\Program Files\Python311`
> 이다. `py -3.11` 은 레지스트리를 조회하므로 설치 위치와 무관하다.

---

## 2단계 — clone

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\source" | Out-Null
git clone https://github.com/KOR-GGG/Civil3D-mcp-python-COM.git "$env:USERPROFILE\source\civil3d-mcp"
cd "$env:USERPROFILE\source\civil3d-mcp"
git log --oneline -1
```

**위치를 `%USERPROFILE%\source\civil3d-mcp` 로 고정한다.** 코드 자체는 위치 무관이지만
`experiments/기록지_D.md` 의 명령을 그대로 복사할 수 있어야 한다.

⚠ **Dropbox·OneDrive 안에 두지 말 것.** `.git` 이 동기화되면 충돌 사본이 생겨 저장소가 깨진다.

---

## 3단계 — 가상환경 (이름 `.venv` 고정)

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -c "import sys; print(sys.version); print(sys.maxsize > 2**32)"
```

이미 다른 이름으로 만들었다면 **rename 하지 말고** 폴더를 지우고 다시 만든다
(생성 시 경로가 `pyvenv.cfg` 와 `Scripts\*.exe` 에 박힌다).

---

## 4단계 — 의존성

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade pip setuptools wheel
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

**확인**

```powershell
.\.venv\Scripts\python.exe -c "import win32com.client, pythoncom, fastmcp, pydantic; print('runtime ok')"
.\.venv\Scripts\python.exe -c "import clr; print('clr ok')"
```

`requirements-dev.txt` 를 같이 까는 이유 — pytest 78건(7단계)과 `anthropic`(API 하니스)이
거기에만 있다. 수동 시행 경로만 쓸 거면 `anthropic` 은 없어도 된다.

> 회사망이 SSL 을 가로채면 pip 가 `CERTIFICATE_VERIFY_FAILED` 로 죽는다.
> 사내 인증서를 `pip config set global.cert <경로>` 로 지정하거나 IT 에 문의한다.

---

## 5단계 — editable 설치

```powershell
.\.venv\Scripts\python.exe -m pip install -e .
Test-Path .\.venv\Scripts\civil3d-mcp.exe        # -> True
```

이 exe 가 만들어져야 한다. Claude Desktop 설정과 `experiments/harness_d.py` 가
**절대경로로** 이 파일을 가리킨다.

---

## 6단계 — 사전 점검

```powershell
.\.venv\Scripts\python.exe setup_check.py
```

**이 시점에 FAIL 이어도 되는 것**

| 항목 | 판단 |
|---|---|
| Windows OS · Python >= 3.11 · 64-bit · fastmcp · win32com · pythoncom · pydantic | **FAIL 이면 안 된다** → 1·4단계로 |
| Civil 3D installation · Autodesk .NET DLLs | Civil 3D 가 깔려 있다면 **FAIL 이면 안 된다** |
| `AeccLandMgd.dll` "optional, absent in 2025+" | 정상. 2025+ 에서 제거된 DLL 이다 |
| **Civil 3D running** | **무시** — 아직 안 켰다 |
| **Claude Desktop config** | **무시** — 10단계 전이다 |

⚠ 위 둘 때문에 **종료코드가 1 이고 "Fix the failures above" 가 뜬다.** 이 단계에서는 정상이다.

---

## 7단계 — 단위시험 78건 (Civil 3D 없이 도는 유일한 검증)

```powershell
.\.venv\Scripts\python.exe -m pytest test_earthwork_core.py test_hydraulics_core.py -v
```

마지막 줄이 **`78 passed`** 여야 한다(토공 30 + 수리 48). 숫자가 다르면 그 설치를 의심한다.

`PytestConfigWarning: Unknown config option: asyncio_mode` 는 무시한다 — async 시험은 0건이다.

> 시험 파일은 저장소 **루트**에 있다. `tests/` 디렉터리는 존재한 적이 없다.

---

## 8단계 — `_golden` 도면 3개 생성 ★ 진짜 관문

`.gitignore` 가 `_golden/` 과 `*.dwg` 를 제외하므로 **새 clone 에는 도면이 하나도 없다.**
없으면 `experiments/activate_env.py` 가 `도면 파일이 없다` 로 즉시 종료한다.

**8-1.** Civil 3D 를 실행하고 **도면을 하나 연다.** 시작 탭만으로는 안 된다.

**8-2.** 연결 확인 — `setup_check.py` 의 `Civil 3D running` 이 PASS 하고 detail 에
`Connected via AeccXUiLand.AeccApplication.13.x` 가 떠야 한다.
`Connected via AutoCAD.Application` 이면 Civil 3D 계층을 못 잡은 것이다.

**8-3.** 한 줄씩, 앞 명령이 끝난 뒤 다음을 친다.

```powershell
.\.venv\Scripts\python.exe make_test_surfaces.py A B C P --new --save-as test_surfaces.dwg
.\.venv\Scripts\python.exe make_test_surfaces.py E --new --save-as test_surfaces_env2.dwg
.\.venv\Scripts\python.exe make_test_surfaces.py N --new --save-as test_surfaces_env3.dwg
```

**확인**

```powershell
Get-ChildItem .\_golden\*.dwg | Select-Object Name, Length, LastWriteTime
```

- 세 파일이 다 있어야 하고 **파일명은 한 글자도 바꾸면 안 된다** — `activate_env.py` 가 그대로 찾는다.
- 각 실행 끝에 `저장 완료: ...\_golden\...dwg` 가 나온다.
- 서피스 개수: 환경 1 = **14개**, 환경 2 = **8개**, 환경 3 = **8개**.

> **세 명령 모두 `--new` 를 붙인다.** `--new` 는 Civil 3D 미터 템플릿으로 빈 새 도면을
> 스스로 만든다. 붙이지 않으면 **열려 있던 도면에** 시험 서피스가 들어가고, 이어지는
> `SaveAs` 가 **그 도면의 저장 경로까지** `_golden\...` 으로 바꿔 버린다.
> 실무 도면이 열려 있을 때의 사고를 막으려고, 서피스가 이미 있는 도면에서 `--save-as` 를
> `--new` 없이 쓰면 스크립트가 **중단**하도록 해 두었다(2026-08-18).
>
> **`--save-as` 를 빠뜨리면 저장되지 않는다.** 그때는 스크립트가 세 파일명을 출력하며 경고한다.
> 2026-08-18 이전에는 저장 기능 자체가 없어 손으로 「다른 이름으로 저장」을 해야 했다.

---

## 9단계 — 환경 전환 확인

```powershell
.\.venv\Scripts\python.exe experiments\activate_env.py          # 상태만 본다(표식 안 심음)
```

실제 시행 준비는 인자를 준다.

```powershell
.\.venv\Scripts\python.exe experiments\activate_env.py 3        # 1 · 2 · 3
```

끝에 ① 전환 안내 ② 서피스 목록(14/8/8) ③ `ZZ_PROBE_...` 표식 ④ 시행 전 점검표가 나와야 한다.

**설치 검증만 할 목적이면 인자 없는 형태를 쓴다.** 1·2·3 을 돌리면 도면에 표식이 남는다.
남았으면 `activate_env.py clean` 으로 지우고 **도면을 저장하지 않는다.**

---

## 10단계 — Claude Desktop 등록

⚠ **기동 순서를 지킨다: ① Civil 3D 실행 → ② 도면 열기 → ③ Claude Desktop 실행.**
`client.py` 의 `_ensure_connected` 는 연결이 끊겨 있으면 **예외만 올리고 재연결하지 않는다.**
Claude Desktop 을 먼저 켜면 서버는 뜨고 도구 24개도 보이지만 COM 도구가 전부 실패하며,
나중에 Civil 3D 를 켜도 **복구되지 않는다.** 서버를 다시 띄워야 한다.

**10-1. 설정 파일 경로**

```powershell
$store = Get-AppxPackage *Claude* -ErrorAction SilentlyContinue
$cfg = if ($store) {
  "$env:LOCALAPPDATA\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json"
} else {
  "$env:APPDATA\Claude\claude_desktop_config.json"
}
$cfg
```

Store 판은 샌드박스라 `%APPDATA%\Claude` 를 **읽지 않는다**(2026-08-17 실측).

**10-2. `mcpServers` 병합 — 기존 키를 보존하고 BOM 없이 저장**

```json
{
  "mcpServers": {
    "civil3d-mcp": {
      "command": "C:\\Users\\YOURNAME\\source\\civil3d-mcp\\.venv\\Scripts\\civil3d-mcp.exe"
    }
  }
}
```

**반드시 절대경로를 쓴다.** `"command": "civil3d-mcp"` 는 `.venv\Scripts\` 가 PATH 에 없어
찾지 못하고, `"command": "python"` 변형은 시스템 인터프리터를 불러 의존성이 하나도 없다.
JSON 은 환경변수를 풀지 않으므로 PC 마다 손으로 고쳐야 하고, 백슬래시는 `\\` 로 이스케이프한다.

자기 경로는 이렇게 얻는다.

```powershell
"$env:USERPROFILE\source\civil3d-mcp\.venv\Scripts\civil3d-mcp.exe"
```

⚠ **BOM 없이 저장.** PowerShell 의 `Out-File -Encoding utf8` 은 BOM 을 붙이는데, 그러면
Claude Desktop 이 JSON 파싱에 실패하고 **아무 오류 없이 서버만 안 뜬다.**
`Set-Content -Encoding utf8NoBOM` 을 쓴다.

**10-3. 확인** — 아이콘을 찾지 말고 로그를 본다.

```powershell
Select-String "Server started and connected successfully" "$env:APPDATA\Claude\logs\mcp*.log" |
  Select-Object -Last 3
```

---

## 자주 막히는 지점

| 증상 | 원인 | 조치 |
|---|---|---|
| `Activate.ps1` 실행 불가 | ExecutionPolicy(회사 GPO) | activate 하지 말고 `.\.venv\Scripts\python.exe` 직접 호출 |
| `activate_env.py` 가 표식 심기에서 `FileNotFoundError` | venv 이름이 `.venv` 가 아님 | venv 폴더를 지우고 `py -3.11 -m venv .venv` |
| `도면 파일이 없다: ..._golden\...dwg` | 8단계 미수행 또는 파일명 불일치 | 8-3 을 파일명 그대로 다시 |
| pip `CERTIFICATE_VERIFY_FAILED` | 회사망 SSL 가로채기 | `pip config set global.cert <사내 인증서>` |
| `setup_check.py` 종료코드 1 | Civil 3D 미실행 / Desktop 미등록 | 6단계 표 참조 — 그 둘은 정상 |
| Claude Desktop 에 서버가 안 뜸(오류도 없음) | 설정 파일 BOM | `Set-Content -Encoding utf8NoBOM` 으로 다시 저장 |
| 도구는 보이는데 전부 실패 | Claude Desktop 을 Civil 3D 보다 먼저 켬 | Civil 3D → 도면 → Desktop 순으로 다시 |
| `Connected via AutoCAD.Application` | Civil 3D 계층 미획득 | Civil 3D(AutoCAD 아님)로 실행했는지 확인 |
| `make_test_surfaces.py` 가 "중단 — --save-as 는..." | 실무 도면이 열려 있음 | `--new` 를 붙이거나 빈 도면을 연 뒤 다시 |

## `.env` 에 대하여

**서버는 `.env` 를 읽지 않는다.** `src/` 어디에도 `load_dotenv` 가 없고, `client.py` 는
프로세스 환경변수만 본다. `CIVIL3D_BIN_PATH` 를 주려면 둘 중 하나다.

1. 진짜 환경변수 — `$env:CIVIL3D_BIN_PATH = "..."` 또는 사용자/시스템 변수
2. Claude Desktop 설정의 `env` 블록

`.env` 는 `experiments/harness_d.py` 가 `ANTHROPIC_API_KEY` 를 읽을 때만 쓰인다.
