#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
00_RUN_ALL.py

One-click runner for the extraction agents (phase 1) and optional
dashboard preparation (phase 2 placeholder).
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional
from urllib.error import HTTPError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen


BASE_DIR = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = BASE_DIR / "02_SCRIPTS"
INPUT_DIR = BASE_DIR / "01_INPUT"
OUTPUT_DIR = BASE_DIR / "03_OUTPUT"
DASHBOARD_DIR = BASE_DIR / "04_DASHBOARD"
LOG_DIR = BASE_DIR / "05_LOGS"
STATE_FILE = LOG_DIR / "_state_input.json"

PYTHON_EXE = Path(sys.executable)

# Update these if your local install differs.
POPPLER_BIN = Path(r"C:\Users\Usuario\Desktop\poppler-25.12.0\Library\bin")
TESSERACT_EXE: Optional[Path] = None

DEBUG_OCR = False
STOP_ON_ERROR = False
PROCESS_ONLY_CHANGED = True

# Phase 2 placeholder
RUN_CONSOLIDATORS = False
CONSOLIDATOR_SCRIPTS: List[Path] = []

# Optional: copy approved bases to dashboard folder
COPY_BASES_TO_DASHBOARD = False
DASHBOARD_BASES_DIR = DASHBOARD_DIR / "BASES"

# SharePoint upload (Microsoft Graph)
UPLOAD_TO_SHAREPOINT = False
UPLOAD_SOURCE_DIR = OUTPUT_DIR
UPLOAD_SKIP_EXISTING = False
SP_FOLDER_PATH = "Comercial/1. Analise de Credito"
UPLOAD_OVERWRITE_KEYWORDS = ["IMBIPARK", "JACIEL"]

# Env vars for Graph auth (do not hardcode secrets)
ENV_TENANT_ID = "SP_TENANT_ID"
ENV_CLIENT_ID = "SP_CLIENT_ID"
ENV_CLIENT_SECRET = "SP_CLIENT_SECRET"
ENV_SITE_URL = "SP_SITE_URL"
ENV_LIBRARY_NAME = "SP_LIBRARY_NAME"
ENV_FOLDER_PATH = "SP_FOLDER_PATH"


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _find_output_dir(prefix: str, preferred_name: Optional[str] = None) -> Path:
    if preferred_name:
        preferred = OUTPUT_DIR / preferred_name
        if preferred.exists():
            return preferred
    if OUTPUT_DIR.exists():
        matches = [p for p in OUTPUT_DIR.iterdir() if p.is_dir() and p.name.startswith(prefix)]
        if matches:
            return matches[0]
    if preferred_name:
        return OUTPUT_DIR / preferred_name
    return OUTPUT_DIR / prefix


def _list_company_dirs(input_dir: Path) -> List[Path]:
    if not input_dir.exists():
        return []
    return sorted([p for p in input_dir.iterdir() if p.is_dir()])


def _scan_input_state(input_dir: Path) -> dict:
    state = {}
    if not input_dir.exists():
        return state
    for p in input_dir.rglob("*"):
        if not p.is_file():
            continue
        try:
            stat = p.stat()
        except OSError:
            continue
        rel = p.relative_to(input_dir).as_posix()
        state[rel] = [int(stat.st_mtime), stat.st_size]
    return state


def _load_state(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_state(path: Path, state: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, ensure_ascii=True, indent=2), encoding="utf-8")


def _detect_changed_companies(prev_state: dict, curr_state: dict) -> List[str]:
    changed = set()
    for rel, info in curr_state.items():
        if prev_state.get(rel) != info:
            company = rel.split("/", 1)[0]
            if company:
                changed.add(company)
    return sorted(changed)


def _count_recent_xlsx(outdir: Path, since_ts: float) -> int:
    if not outdir.exists():
        return 0
    count = 0
    for p in outdir.rglob("*.xlsx"):
        try:
            if p.stat().st_mtime >= since_ts:
                count += 1
        except OSError:
            continue
    return count


def _log_write(log_fh, msg: str) -> None:
    print(msg)
    log_fh.write(msg + "\n")


def _run_cmd(
    log_fh,
    label: str,
    args: List[str],
    cwd: Optional[Path] = None,
    outdir: Optional[Path] = None,
) -> bool:
    start_ts = time.time()
    _log_write(log_fh, f"\n== {label} ==")
    _log_write(log_fh, "CMD: " + " ".join(args))

    result = subprocess.run(
        args,
        cwd=str(cwd or BASE_DIR),
        capture_output=True,
        text=True,
    )

    suppress_markers = (
        "Could not get FontBBox",
    )

    if result.stdout:
        for line in result.stdout.splitlines():
            if any(marker in line for marker in suppress_markers):
                continue
            _log_write(log_fh, line.rstrip())
    if result.stderr:
        for line in result.stderr.splitlines():
            if any(marker in line for marker in suppress_markers):
                continue
            _log_write(log_fh, line.rstrip())

    ok = result.returncode == 0
    if not ok:
        _log_write(log_fh, f"[ERROR] {label} exited with code {result.returncode}")
        if STOP_ON_ERROR:
            raise SystemExit(result.returncode)
    else:
        _log_write(log_fh, f"[OK] {label}")

    if outdir:
        recent = _count_recent_xlsx(outdir, start_ts)
        _log_write(log_fh, f"[INFO] Recent .xlsx in {outdir.name}: {recent}")

    return ok


def _find_endividamento_files(input_dir: Path) -> List[Path]:
    exts = {".xlsx", ".xls", ".csv", ".pdf", ".docx", ".png", ".jpg", ".jpeg"}
    files: List[Path] = []
    for p in input_dir.rglob("*"):
        if not p.is_file():
            continue
        if p.suffix.lower() not in exts:
            continue
        if "endividamento" in p.name.lower():
            files.append(p)
    return sorted(files)


def _copy_bases_to_dashboard(log_fh, outdirs: Iterable[Path]) -> None:
    DASHBOARD_BASES_DIR.mkdir(parents=True, exist_ok=True)
    copied = 0
    for outdir in outdirs:
        if not outdir.exists():
            continue
        for p in outdir.rglob("*.xlsx"):
            target = DASHBOARD_BASES_DIR / p.name
            shutil.copy2(p, target)
            copied += 1
    _log_write(log_fh, f"[INFO] Copied {copied} .xlsx files to {DASHBOARD_BASES_DIR}")


def _get_env_required(log_fh, name: str) -> Optional[str]:
    val = (os.environ.get(name) or "").strip()
    if not val:
        _log_write(log_fh, f"[WARN] Missing env var: {name}")
        return None
    return val


def _http_json(method: str, url: str, headers: dict, payload: Optional[dict] = None) -> dict:
    data = None
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
    req = Request(url, data=data, headers=headers, method=method)
    with urlopen(req) as resp:
        body = resp.read().decode("utf-8")
        return json.loads(body) if body else {}


def _http_put_bytes(url: str, headers: dict, data: bytes) -> dict:
    req = Request(url, data=data, headers=headers, method="PUT")
    with urlopen(req) as resp:
        body = resp.read().decode("utf-8")
        return json.loads(body) if body else {}


def _http_form(method: str, url: str, headers: dict, payload: dict) -> dict:
    data = urlencode(payload).encode("utf-8")
    req = Request(url, data=data, headers=headers, method=method)
    with urlopen(req) as resp:
        body = resp.read().decode("utf-8")
        return json.loads(body) if body else {}


def _norm_ascii_upper(s: str) -> str:
    s2 = unicodedata.normalize("NFD", s or "")
    s2 = "".join(ch for ch in s2 if unicodedata.category(ch) != "Mn")
    return s2.upper()


def _should_overwrite(name: str) -> bool:
    if not UPLOAD_OVERWRITE_KEYWORDS:
        return False
    up = _norm_ascii_upper(name)
    return any(_norm_ascii_upper(k) in up for k in UPLOAD_OVERWRITE_KEYWORDS)


def _remote_item_exists(log_fh, token: str, drive_id: str, remote_path: str) -> bool:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{remote_path}"
    try:
        _http_json("GET", url, headers)
        return True
    except HTTPError as exc:
        if exc.code == 404:
            return False
        _log_write(log_fh, f"[ERROR] Item lookup failed: {remote_path} -> {exc}")
        return False
    except Exception as exc:
        _log_write(log_fh, f"[ERROR] Item lookup failed: {remote_path} -> {exc}")
        return False


def _get_graph_token(log_fh) -> Optional[str]:
    tenant_id = _get_env_required(log_fh, ENV_TENANT_ID)
    client_id = _get_env_required(log_fh, ENV_CLIENT_ID)
    client_secret = _get_env_required(log_fh, ENV_CLIENT_SECRET)
    if not tenant_id or not client_id or not client_secret:
        return None
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    try:
        resp = _http_form("POST", url, headers, data)
    except Exception as exc:
        _log_write(log_fh, f"[ERROR] Token request failed: {exc}")
        return None
    token = resp.get("access_token")
    if not token:
        _log_write(log_fh, "[ERROR] Token response missing access_token.")
    return token


def _resolve_site_id(log_fh, token: str, site_url: str) -> tuple[Optional[str], Optional[str]]:
    site_url = site_url.rstrip("/")
    if not site_url.startswith("https://"):
        _log_write(log_fh, "[ERROR] SP_SITE_URL must start with https://")
        return None, None
    host = site_url.split("/")[2]
    path = "/" + "/".join(site_url.split("/")[3:])
    headers = {"Authorization": f"Bearer {token}"}

    def _fetch_site_id(site_path: str) -> Optional[str]:
        url = f"https://graph.microsoft.com/v1.0/sites/{host}:{site_path}"
        resp = _http_json("GET", url, headers)
        return resp.get("id")

    try:
        site_id = _fetch_site_id(path)
        if site_id:
            return site_id, path
    except Exception:
        site_id = None

    if "/" in path.strip("/"):
        parent = "/" + path.strip("/").rsplit("/", 1)[0]
        try:
            site_id = _fetch_site_id(parent)
            if site_id:
                _log_write(log_fh, f"[WARN] Using parent site path: {parent}")
                return site_id, parent
        except Exception as exc:
            _log_write(log_fh, f"[ERROR] Site lookup failed: {exc}")
            return None, None

    _log_write(log_fh, "[ERROR] Site lookup missing id.")
    return None, None


def _get_drive_id(log_fh, token: str, site_id: str) -> Optional[str]:
    headers = {"Authorization": f"Bearer {token}"}
    lib_name = (os.environ.get(ENV_LIBRARY_NAME) or "").strip()
    if lib_name:
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        try:
            resp = _http_json("GET", url, headers)
        except Exception as exc:
            _log_write(log_fh, f"[ERROR] Drive list failed: {exc}")
            return None
        for drive in resp.get("value", []):
            if (drive.get("name") or "").strip().lower() == lib_name.lower():
                _log_write(log_fh, f"[INFO] Using library: {drive.get('name')}")
                return drive.get("id")
        _log_write(log_fh, f"[WARN] Library not found: {lib_name}. Using default.")

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
    try:
        resp = _http_json("GET", url, headers)
    except Exception as exc:
        _log_write(log_fh, f"[ERROR] Default drive lookup failed: {exc}")
        return None
    drive_id = resp.get("id")
    if not drive_id:
        _log_write(log_fh, "[ERROR] Default drive lookup missing id.")
    return drive_id


def _ensure_remote_folder(log_fh, drive_id: str, token: str, folder_path: Path) -> None:
    if not folder_path.parts:
        return
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    current: List[str] = []
    for part in folder_path.parts:
        current.append(part)
        remote = "/".join(quote(p) for p in current)
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{remote}"
        payload = {"folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
        try:
            _http_json("PUT", url, headers, payload)
        except HTTPError as exc:
            if exc.code == 409:
                continue
            _log_write(log_fh, f"[ERROR] Create folder failed: {remote} -> {exc}")
            return
        except Exception as exc:
            _log_write(log_fh, f"[ERROR] Create folder failed: {remote} -> {exc}")
            return


def _upload_folder_to_sharepoint(log_fh, source_dir: Path) -> None:
    site_url = _get_env_required(log_fh, ENV_SITE_URL)
    if not site_url:
        return
    token = _get_graph_token(log_fh)
    if not token:
        return

    site_id, _site_path = _resolve_site_id(log_fh, token, site_url)
    if not site_id:
        return
    drive_id = _get_drive_id(log_fh, token, site_id)
    if not drive_id:
        return

    if not source_dir.exists():
        _log_write(log_fh, f"[WARN] Upload source not found: {source_dir}")
        return

    folder_path = (os.environ.get(ENV_FOLDER_PATH) or SP_FOLDER_PATH).strip().strip("/")
    headers = {"Authorization": f"Bearer {token}"}
    uploaded = 0
    skipped = 0
    _ensure_remote_folder(log_fh, drive_id, token, Path(folder_path))
    for p in source_dir.rglob("*.xlsx"):
        rel = p.relative_to(source_dir)
        remote_folder = Path(folder_path) / rel.parent
        _ensure_remote_folder(log_fh, drive_id, token, remote_folder)
        remote_path = "/".join(quote(part) for part in (remote_folder / p.name).parts)
        exists = _remote_item_exists(log_fh, token, drive_id, remote_path)
        if exists and not _should_overwrite(p.name):
            skipped += 1
            _log_write(log_fh, f"[INFO] Skip existing (no overwrite): {p.name}")
            continue

        if exists:
            conflict = "?@microsoft.graph.conflictBehavior=replace"
        else:
            conflict = "?@microsoft.graph.conflictBehavior=fail"
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{remote_path}:/content{conflict}"
        try:
            _http_put_bytes(url, {**headers, "Content-Type": "application/octet-stream"}, p.read_bytes())
            uploaded += 1
        except HTTPError as exc:
            if exc.code == 409 and UPLOAD_SKIP_EXISTING:
                skipped += 1
                _log_write(log_fh, f"[INFO] Skip existing: {p.name}")
                continue
            _log_write(log_fh, f"[ERROR] Upload failed: {p.name} -> {exc}")
        except Exception as exc:
            _log_write(log_fh, f"[ERROR] Upload failed: {p.name} -> {exc}")
    _log_write(log_fh, f"[INFO] Uploaded {uploaded} file(s) to SharePoint.")
    if UPLOAD_SKIP_EXISTING:
        _log_write(log_fh, f"[INFO] Skipped existing files: {skipped}")


def main() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_path = LOG_DIR / f"run_{_timestamp()}.log"
    with log_path.open("w", encoding="utf-8") as log_fh:
        run_ok = True
        _log_write(log_fh, f"Runner started at {datetime.now().isoformat(timespec='seconds')}")
        _log_write(log_fh, f"Base dir: {BASE_DIR}")

        if not INPUT_DIR.exists():
            _log_write(log_fh, f"[ERROR] Input dir not found: {INPUT_DIR}")
            return

        if not POPPLER_BIN.exists():
            _log_write(log_fh, f"[WARN] POPPLER_BIN not found: {POPPLER_BIN}")
        if TESSERACT_EXE and not TESSERACT_EXE.exists():
            _log_write(log_fh, f"[WARN] TESSERACT_EXE not found: {TESSERACT_EXE}")

        out_01 = _find_output_dir("1.", "1. Relatório de Visita")
        out_02 = _find_output_dir("2.", "2. SERASA CEDENTE")
        out_03 = _find_output_dir("3.", "3. Endividamento")
        out_04 = _find_output_dir("4.", "4. SCR CEDENTE")
        out_05 = _find_output_dir("5.", "5. VADU")
        out_06 = _find_output_dir("6.", "6. SERASA SÓCIO")
        out_07 = _find_output_dir("7.", "7. SCR SÓCIO")
        out_10 = _find_output_dir("10.", "10. SERASA SACADO")
        out_11 = _find_output_dir("11.", "11. SCR CURVA ABC")

        prev_state = _load_state(STATE_FILE)
        curr_state = _scan_input_state(INPUT_DIR)
        if PROCESS_ONLY_CHANGED:
            changed_companies = _detect_changed_companies(prev_state, curr_state)
        else:
            changed_companies = [p.name for p in _list_company_dirs(INPUT_DIR)]

        if not changed_companies:
            _log_write(log_fh, "[INFO] No changes detected in 01_INPUT. Skipping agents.")
        else:
            for company in changed_companies:
                company_dir = INPUT_DIR / company
                if not company_dir.exists():
                    continue

                # Phase 1 - extraction agents
                ok = _run_cmd(
                    log_fh,
                    f"AGENTE 01 - VISITA STRICT ({company})",
                    [
                        str(PYTHON_EXE),
                        str(SCRIPTS_DIR / "agente_01_visita_strict.py"),
                        "--input",
                        str(company_dir),
                        "--outdir",
                        str(out_01),
                    ],
                    outdir=out_01,
                )
                run_ok = run_ok and ok

                ok = _run_cmd(
                    log_fh,
                    f"AGENTE 02 - SERASA CEDENTE ({company})",
                    [
                        str(PYTHON_EXE),
                        str(SCRIPTS_DIR / "agente_02_serasacedente.py"),
                        "--input",
                        str(company_dir),
                        "--outdir",
                        str(out_02),
                    ],
                    outdir=out_02,
                )
                run_ok = run_ok and ok

                endiv_files = _find_endividamento_files(company_dir)
                if endiv_files:
                    for p in endiv_files:
                        endiv_cmd = [
                            str(PYTHON_EXE),
                            str(SCRIPTS_DIR / "agente_03_endividamento.py"),
                            str(p),
                        ]
                        if POPPLER_BIN and POPPLER_BIN.exists():
                            endiv_cmd += ["--poppler", str(POPPLER_BIN)]
                        ok = _run_cmd(
                            log_fh,
                            f"AGENTE 03 - ENDIVIDAMENTO ({company}) {p.name}",
                            endiv_cmd,
                            outdir=out_03,
                        )
                        run_ok = run_ok and ok
                else:
                    _log_write(log_fh, f"[INFO] AGENTE 03 - no endividamento files found for {company}.")

                vadu_cmd = [
                    str(PYTHON_EXE),
                    str(SCRIPTS_DIR / "agente_04_vadu.py"),
                    "--input-dir",
                    str(company_dir),
                    "--outdir",
                    str(out_05),
                ]
                if TESSERACT_EXE:
                    vadu_cmd += ["--tesseract", str(TESSERACT_EXE)]
                ok = _run_cmd(
                    log_fh,
                    f"AGENTE 04 - VADU ({company})",
                    vadu_cmd,
                    outdir=out_05,
                )
                run_ok = run_ok and ok

                if POPPLER_BIN.exists():
                    ok = _run_cmd(
                        log_fh,
                        f"AGENTE 05 - SCR CEDENTE ({company})",
                        [
                            str(PYTHON_EXE),
                            str(SCRIPTS_DIR / "agente_05_scr_cedente.py"),
                            "--input",
                            str(company_dir),
                            "--outdir",
                            str(out_04),
                            "--poppler",
                            str(POPPLER_BIN),
                        ]
                        + (["--tesseract", str(TESSERACT_EXE)] if TESSERACT_EXE else [])
                        + (["--debug"] if DEBUG_OCR else []),
                        outdir=out_04,
                    )
                    run_ok = run_ok and ok
                else:
                    _log_write(log_fh, "[WARN] AGENTE 05 - skipped (missing POPPLER_BIN).")

                ok = _run_cmd(
                    log_fh,
                    f"AGENTE 06 - SERASA SOCIO ({company})",
                    [
                        str(PYTHON_EXE),
                        str(SCRIPTS_DIR / "agente_06_serasa_socio.py"),
                        "--input",
                        str(company_dir),
                        "--outdir",
                        str(out_06),
                    ]
                    + (["--poppler", str(POPPLER_BIN)] if POPPLER_BIN.exists() else [])
                    + (["--tesseract", str(TESSERACT_EXE)] if TESSERACT_EXE else []),
                    outdir=out_06,
                )
                run_ok = run_ok and ok

                ok = _run_cmd(
                    log_fh,
                    f"AGENTE 07 - SERASA SACADO ({company})",
                    [
                        str(PYTHON_EXE),
                        str(SCRIPTS_DIR / "agente_07_serasa_sacado.py"),
                        "--input",
                        str(company_dir),
                        "--outdir",
                        str(out_10),
                    ]
                    + (["--poppler", str(POPPLER_BIN)] if POPPLER_BIN.exists() else [])
                    + (["--tesseract", str(TESSERACT_EXE)] if TESSERACT_EXE else []),
                    outdir=out_10,
                )
                run_ok = run_ok and ok

                if POPPLER_BIN.exists():
                    ok = _run_cmd(
                        log_fh,
                        f"AGENTE 08 - SCR SOCIO ({company})",
                        [
                            str(PYTHON_EXE),
                            str(SCRIPTS_DIR / "agente_08_scr_socio.py"),
                            "--input",
                            str(company_dir),
                            "--outdir",
                            str(out_07),
                            "--poppler",
                            str(POPPLER_BIN),
                        ]
                        + (["--tesseract", str(TESSERACT_EXE)] if TESSERACT_EXE else [])
                        + (["--debug"] if DEBUG_OCR else []),
                        outdir=out_07,
                    )
                    run_ok = run_ok and ok
                else:
                    _log_write(log_fh, "[WARN] AGENTE 08 - skipped (missing POPPLER_BIN).")

                if POPPLER_BIN.exists():
                    ok = _run_cmd(
                        log_fh,
                        f"AGENTE 09 - SCR SACADO ({company})",
                        [
                            str(PYTHON_EXE),
                            str(SCRIPTS_DIR / "agente_09_scr_sacado.py"),
                            "--input",
                            str(company_dir),
                            "--outdir",
                            str(out_11),
                            "--poppler",
                            str(POPPLER_BIN),
                        ]
                        + (["--tesseract", str(TESSERACT_EXE)] if TESSERACT_EXE else [])
                        + (["--debug"] if DEBUG_OCR else []),
                        outdir=out_11,
                    )
                    run_ok = run_ok and ok
                else:
                    _log_write(log_fh, "[WARN] AGENTE 09 - skipped (missing POPPLER_BIN).")

        # Phase 2 - consolidators (placeholder)
        if RUN_CONSOLIDATORS and CONSOLIDATOR_SCRIPTS:
            for script in CONSOLIDATOR_SCRIPTS:
                _run_cmd(
                    log_fh,
                    f"CONSOLIDATOR - {script.name}",
                    [str(PYTHON_EXE), str(script)],
                )
        else:
            _log_write(log_fh, "[INFO] Phase 2 skipped (no consolidators configured).")

        if COPY_BASES_TO_DASHBOARD:
            _copy_bases_to_dashboard(
                log_fh,
                [out_01, out_02, out_03, out_04, out_05, out_06, out_07, out_10, out_11],
            )
        else:
            _log_write(log_fh, "[INFO] Dashboard copy disabled.")

        if UPLOAD_TO_SHAREPOINT:
            _upload_folder_to_sharepoint(log_fh, UPLOAD_SOURCE_DIR)
        else:
            _log_write(log_fh, "[INFO] SharePoint upload disabled.")

        if run_ok and curr_state:
            _save_state(STATE_FILE, curr_state)

        _log_write(log_fh, "Runner finished.")


if __name__ == "__main__":
    main()
