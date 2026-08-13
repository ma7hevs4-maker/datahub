# -*- coding: utf-8 -*-
"""Auto-update do DataHub a partir de GitHub Releases.

Fluxo:
  1. check_for_update(repo) -> (new_version:str, download_url:str) | None
  2. prepare_update(download_url, progress) -> baixa o zip, extrai em %TEMP%
     e gera um .bat que copia a pasta DataHub por cima da instalada e
     relanca o exe.
  3. launch_updater(bat_path) -> dispara o .bat em processo separado e
     encerra este app (para liberar o exe em uso).

O pacote publicado num Release do GitHub deve ser um zip cujo conteudo tem
uma pasta "DataHub" (ou seja: zipar a pasta dist\DataHub, nao o seu conteudo).
O asset deve chamar-se "DataHub.zip" (ou ser o unico .zip do Release).
"""
import os
import sys
import json
import shutil
import zipfile
import tempfile
import subprocess
import urllib.request
import urllib.error
from pathlib import Path

# >>>>> PREENCHA COM "seu-usuario/seu-repo" do GitHub <<<<<
REPO = "SEU_USUARIO/SEU_REPO"

APP_VERSION_FALLBACK = "0.0.0"
API_URL = "https://api.github.com/repos/{repo}/releases/latest"
ASSET_NAME = "DataHub.zip"  # nome esperado do asset no Release


def local_version():
    """Versao embedada pelo build (version.txt na pasta do exe)."""
    try:
        p = Path(sys.executable).parent / "version.txt"
        if p.exists():
            v = p.read_text(encoding="utf-8").strip()
            if v:
                return v
    except Exception:
        pass
    return APP_VERSION_FALLBACK


def is_installed():
    """True somente quando rodando do build (dist/DataHub) com version.txt.
    Evita que o update sobrescreva a pasta do Python ao rodar via 'python app.py'."""
    try:
        return (Path(sys.executable).parent / "version.txt").exists()
    except Exception:
        return False


def _norm_tag(tag):
    return (tag or "").lstrip("vV").strip() or (tag or "")


def parse_version(v):
    parts = []
    for p in _norm_tag(v).split("."):
        num = ""
        for ch in p:
            if ch.isdigit():
                num += ch
            else:
                break
        parts.append(int(num) if num else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def is_newer(remote, current):
    try:
        return parse_version(remote) > parse_version(current)
    except Exception:
        return False


def check_for_update(repo=None, timeout=10):
    """Retorna (tag, download_url) se houver versao mais nova, senao None.
    Qualquer erro de rede/JSON e tratado como 'sem atualizacao' (silencio safe)."""
    repo = repo or REPO
    try:
        url = API_URL.format(repo=repo)
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": "DataHub-Updater",
                "Accept": "application/vnd.github+json",
            },
        )
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = json.loads(r.read().decode("utf-8"))
        tag = data.get("tag_name")
        assets = data.get("assets") or []
        asset = next((a for a in assets if a.get("name") == ASSET_NAME), None)
        if not asset:
            asset = next(
                (a for a in assets if (a.get("name") or "").lower().endswith(".zip")),
                None,
            )
        if not tag or not asset:
            return None
        dl = asset.get("browser_download_url")
        if not dl:
            return None
        if not is_newer(tag, local_version()):
            return None
        return (tag, dl)
    except (urllib.error.URLError, urllib.error.HTTPError, ValueError,
            OSError, json.JSONDecodeError):
        return None
    except Exception:
        return None


def _download(url, dest, progress=None):
    req = urllib.request.Request(url, headers={"User-Agent": "DataHub-Updater"})
    with urllib.request.urlopen(req, timeout=120) as r:
        total = int(r.headers.get("Content-Length", 0) or 0)
        downloaded = 0
        chunk = 64 * 1024
        with open(dest, "wb") as f:
            while True:
                buf = r.read(chunk)
                if not buf:
                    break
                f.write(buf)
                downloaded += len(buf)
                if progress and total:
                    progress(downloaded, total)


def _extract(zip_path, dest_dir):
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(dest_dir)


def prepare_update(download_url, progress=None):
    """Baixa e extrai; retorna o caminho do script de update (.bat) pronto.
    install_dir = pasta que contem o DataHub.exe (pasta 'DataHub')."""
    install_dir = Path(sys.executable).parent
    tmp = Path(tempfile.gettempdir()) / "datahub_upd"
    if tmp.exists():
        shutil.rmtree(tmp, ignore_errors=True)
    tmp.mkdir(parents=True, exist_ok=True)
    zip_path = tmp / "DataHub.zip"
    _download(download_url, zip_path, progress)
    extract_dir = tmp / "extracted"
    extract_dir.mkdir(parents=True, exist_ok=True)
    _extract(zip_path, extract_dir)

    # acha a pasta DataHub extraida
    src = None
    if (extract_dir / "DataHub").is_dir():
        src = extract_dir / "DataHub"
    else:
        for d in extract_dir.rglob("DataHub.exe"):
            src = d.parent
            break
    if src is None:
        raise FileNotFoundError("Pasta DataHub nao encontrada no pacote de atualizacao.")

    bat = tmp / "update.bat"
    install_escaped = str(install_dir).replace('"', '""')
    src_escaped = str(src).replace('"', '""')
    bat_text = (
        "@echo off\n"
        "chcp 65001 >nul\n"
        "timeout /t 2 /nobreak >nul\n"
        f'robocopy "{src_escaped}" "{install_escaped}" /E /R:2 /W:2 /NFL /NDL /NJH /NJS\n'
        f'start "" "{install_escaped}\\DataHub.exe"\n'
        "del \"%~f0\"\n"
    )
    bat.write_text(bat_text, encoding="utf-8")
    return str(bat)


def launch_updater(bat_path):
    """Dispara o updater em processo separado e encerra o app atual."""
    subprocess.Popen(
        ["cmd.exe", "/c", bat_path],
        shell=False,
        creationflags=0x00000200,  # CREATE_NEW_PROCESS_GROUP
        close_fds=True,
    )
    os._exit(0)
