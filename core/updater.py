# -*- coding: utf-8 -*-
"""Auto-update do DataHub a partir de GitHub Releases.

Fluxo:
  1. check_for_update(repo) -> dict com status:
       {"status":"update", "tag":str, "url":str}   -> ha versao nova
       {"status":"uptodate", "tag":str}            -> ja e o mais novo
       {"status":"error", "error":str}             -> falha de rede/API
  2. prepare_update(download_url, expected_version) -> baixa o zip, extrai em
     %TEMP% e gera um .bat que:
       - encerra TODAS as instancias do DataHub.exe (evita arquivo travado),
       - copia a pasta por cima da instalada (robocopy com varias tentativas),
       - CONFERE se o version.txt instalado bate com a versao esperada,
       - so entao relanca o exe; se nao bater, grava update_failed.flag.
  3. launch_updater(bat_path) -> dispara o .bat em processo separado e encerra
     este app (para liberar o exe em uso).

O pacote publicado num Release do GitHub deve ser um zip cujo conteudo tenha
uma pasta "DataHub" (ou seja: zipar a pasta dist/DataHub, nao o seu conteudo).
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
import ssl
from pathlib import Path

# >>>>> slug do repositorio GitHub (usuario/repo) <<<<<
REPO = "ma7hevs4-maker/datahub"

APP_VERSION_FALLBACK = "0.0.0"
API_URL = "https://api.github.com/repos/{repo}/releases/latest"
ASSET_NAME = "DataHub.zip"  # nome esperado do asset no Release
UPDATE_LOG = "update.log"
UPDATE_FAILED_FLAG = "update_failed.flag"


def _make_ssl_ctx():
    """Contexto SSL: cert store do SO (pega CA de proxy corporativo no Windows),
    depois certifi, e em ultimo caso sem verificacao."""
    try:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        ctx.load_default_certs()
        return ctx
    except Exception:
        pass
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        pass
    try:
        return ssl.create_default_context()
    except Exception:
        return ssl._create_unverified_context()


def _version_file_candidates():
    exe_dir = Path(sys.executable).parent
    # onedir (PyInstaller 6.x) coloca os datas em _internal/, enquanto o
    # _MEIPASS so existe no onefile. Checa ambos para o auto-update funcionar.
    candidates = [exe_dir / "version.txt", exe_dir / "_internal" / "version.txt"]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "version.txt")
    return candidates


def local_version():
    """Versao embedada pelo build (version.txt, na pasta do exe ou em _internal)."""
    for p in _version_file_candidates():
        try:
            if p.exists():
                v = p.read_text(encoding="utf-8").strip()
                if v:
                    return v
        except Exception:
            pass
    return APP_VERSION_FALLBACK


def is_installed():
    """True quando rodando do build (dist/DataHub). Evita que o update sobrescreva
    a pasta do Python ao rodar via 'python app.py'.

    Considera instalado tambem se existir o exe empacotado ou a pasta _internal,
    mesmo que o version.txt tenha sumido (assim o botao nao se recusa a atualizar).
    """
    try:
        if any(p.exists() for p in _version_file_candidates()):
            return True
        exe_dir = Path(sys.executable).parent
        if (exe_dir / "DataHub.exe").exists() or (exe_dir / "_internal").exists():
            return True
    except Exception:
        pass
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
    """Retorna um dict de estado (veja docstring do modulo).

    Qualquer erro de rede/JSON e tratado como status 'error' (nao como
    'uptodate') para a UI poder avisar 'nao foi possivel verificar' em vez de
    mentir 'esta atualizado'.
    """
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
        with urllib.request.urlopen(req, timeout=timeout, context=_make_ssl_ctx()) as r:
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
            return {"status": "error", "error": "release sem asset DataHub.zip"}
        dl = asset.get("browser_download_url")
        if not dl:
            return {"status": "error", "error": "asset sem url de download"}
        if not is_newer(tag, local_version()):
            return {"status": "uptodate", "tag": tag}
        return {"status": "update", "tag": tag, "url": dl}
    except (urllib.error.URLError, urllib.error.HTTPError, ValueError,
            OSError, json.JSONDecodeError) as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:  # noqa: BLE001
        return {"status": "error", "error": str(e)}


def _download(url, dest, progress=None):
    req = urllib.request.Request(url, headers={"User-Agent": "DataHub-Updater"})
    with urllib.request.urlopen(req, timeout=120, context=_make_ssl_ctx()) as r:
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


def _log_dir():
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    d = Path(base) / "DataHub"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return d


def prepare_update(download_url, progress=None, expected_version=None):
    """Baixa e extrai; retorna o caminho do script de update (.bat) pronto.

    install_dir = pasta que contem o DataHub.exe (pasta 'DataHub').
    O .bat gerado encerra o app, copia e CONFERE a versao antes de relancar.
    """
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

    exp = _norm_tag(expected_version or "")
    install_escaped = str(install_dir).replace('"', '""')
    src_escaped = str(src).replace('"', '""')

    bat_text = (
        "@echo off\n"
        "chcp 65001 >nul\n"
        'set "LOG=%LOCALAPPDATA%\\DataHub\\' + UPDATE_LOG + '"\n'
        'echo [%DATE% %TIME%] === UPDATE INICIADO p/ {EXP} === >> "%LOG%"\n'
        'echo install_dir={INSTALL} >> "%LOG%"\n'
        'echo src={SRC} >> "%LOG%"\n'
        ":kill\n"
        'taskkill /F /IM DataHub.exe /T >> "%LOG%" 2>&1\n'
        "timeout /t 3 /nobreak >nul\n"
        'set "OK=0"\n'
        "for /L %%i in (1,1,3) do (\n"
        '  echo [%DATE% %TIME%] robocopy tentativa %%i >> "%LOG%"\n'
        '  robocopy "{SRC}" "{INSTALL}" /E /R:5 /W:3 /NFL /NDL /NJS >> "%LOG%" 2>&1\n'
        '  set "GOT="\n'
        '  for /f "usebackq delims=" %%v in (`type "{INSTALL}\\version.txt" 2^>nul`) do set "GOT=%%v"\n'
        '  set "GOT=%GOT: =%"\n'
        '  if "%GOT%"=="{EXP}" (\n'
        '    echo [%DATE% %TIME%] VERSAO CONFERE: %GOT% >> "%LOG%"\n'
        "    goto done\n"
        "  )\n"
        '  echo [%DATE% %TIME%] tentativa %%i: versao instalada=%GOT% (esperado {EXP}) >> "%LOG%"\n'
        ")\n"
        'echo [%DATE% %TIME%] FALHA: versao nao confere apos 3 tentativas >> "%LOG%"\n'
        'echo {EXP} > "%LOCALAPPDATA%\\DataHub\\' + UPDATE_FAILED_FLAG + '"\n'
        "goto end\n"
        ":done\n"
        'echo [%DATE% %TIME%] SUCESSO >> "%LOG%"\n'
        'if exist "%LOCALAPPDATA%\\DataHub\\' + UPDATE_FAILED_FLAG + '" del "%LOCALAPPDATA%\\DataHub\\' + UPDATE_FAILED_FLAG + '"\n'
        'start "" "{INSTALL}\\DataHub.exe"\n'
        ":end\n"
        'del "%~f0"\n'
    )
    bat_text = (
        bat_text.replace("{INSTALL}", install_escaped)
        .replace("{SRC}", src_escaped)
        .replace("{EXP}", exp)
    )

    bat = tmp / "update.bat"
    bat.write_text(bat_text, encoding="utf-8")
    return str(bat)


def launch_updater(bat_path):
    """Dispara o updater em processo separado e encerra o app atual.

    O .bat (gerado em prepare_update) e quem encerra de verdade o app, copia e
    relanca — assim a ordem nunca troca (app morto ANTES da copia).
    """
    subprocess.Popen(
        ["cmd.exe", "/c", bat_path],
        shell=False,
        creationflags=0x00000200,  # CREATE_NEW_PROCESS_GROUP
        close_fds=True,
    )
    os._exit(0)
