import json
import os
# Proxy corporativo com certificado self-signed: permite ao webdriver-manager
# baixar/validar o ChromeDriver sem falhar na verificação SSL.
os.environ.setdefault("WDM_SSL_VERIFY", "0")
import threading
import time
import unicodedata
import urllib3
import calendar
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import urlparse, quote

import pandas as pd
import requests
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from config import SpotfireConfig
from core.base_mensal import atualizar_base_mensal
from core.processors.scanner import tratar_m300

NOME_SCANNER = "scanner_tabela_completa.xlsx"
NOME_DESLOCAMENTOS = "scanner_deslocamentos.xlsx"
NOME_M300_HOJE = "m300_deslocamentos_hoje.csv"
NOME_M300_GERAL = "m300_deslocamentos_geral.csv"
TIMEOUT_DOWNLOAD = 300
TIMEOUT_DOWNLOAD_GERAL = 1800
_SUFIXOS_EXPORT = {".csv", ".xlsx", ".xls", ".zip", ""}

# Análise de Produtividade (dia atual). A do Scanner/geral vem da config.
_NOME_ANALISE_PRODUTIVIDADE = "/M300/Produtividade UO TR - RJ"

# Fallback: id do filtro "Período" na análise de Produtividade. Só é usado
# se o filtro não for encontrado pelo título (o id pode mudar se a análise
# for editada no Spotfire).
_PERIODO_FILTRO_ID = "1be0f48b-fe8f-4d13-bf4e-5bd4341d3870"

# Arquivo onde os cookies de sessão do Spotfire são salvos entre execuções
_COOKIES_FILE = os.path.join(
    os.environ.get("APPDATA", ""),
    "DataHub", "spotfire_session.json",
)

_MESES_PT = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr",
    5: "mai", 6: "jun", 7: "jul", 8: "ago",
    9: "set", 10: "out", 11: "nov", 12: "dez",
}
# Mapeamento inverso para ordenar meses na ordem do calendário (não alfabética)
_MESES_NUM = {v: k for k, v in _MESES_PT.items()}


def _watchdog_driver(driver: uc.Chrome, stop_event: threading.Event) -> None:
    """Mata o driver assim que stop_event for sinalizado."""
    stop_event.wait()
    try:
        driver.quit()
    except Exception:
        pass


_PERFIL_SPOTFIRE = os.path.join(
    os.environ.get("APPDATA", ""),
    "DataHub", "spotfire_profile",
)


def _limpar_locks_perfil() -> None:
    """Remove locks, sessão e cookies do perfil para iniciar sempre limpo."""
    for nome in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
        lock = os.path.join(_PERFIL_SPOTFIRE, nome)
        try:
            if os.path.exists(lock):
                os.remove(lock)
        except Exception:
            pass

    default_dir = os.path.join(_PERFIL_SPOTFIRE, "Default")
    # Last Session/Last Tabs: evita restauro de abas (ex.: intranet da Enel).
    # Os Cookies NÃO são deletados para preservar a sessão do Spotfire entre
    # execuções — depois do 1º login manual o usuário não precisa logar de novo.
    for nome in ("Last Session", "Last Tabs"):
        f = os.path.join(default_dir, nome)
        try:
            if os.path.exists(f):
                os.remove(f)
        except Exception:
            pass


def _criar_driver(pasta_download: str) -> uc.Chrome:
    _limpar_locks_perfil()
    opts = uc.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")
    opts.add_argument(
        "--disable-features="
        "BlockInsecurePrivateNetworkRequests,"
        "PrivateNetworkAccessSendPreflights,"
        "PrivateNetworkAccessPermissionPrompt"
    )
    opts.add_experimental_option("prefs", {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "profile.default_content_settings.popups": 0,
        # Pré-autoriza enelcom.sharepoint.com a acessar rede local sem popup
        "profile.content_settings.exceptions.local_network_access": {
            "https://enelcom.sharepoint.com:443,*": {
                "setting": 1,
                "last_modified": "13000000000000000",
            },
        },
    })
    try:
        driver_path = ChromeDriverManager().install()
    except Exception:
        # Fallback: usa o ChromeDriver já em cache (funciona offline/atrás de proxy).
        candidatos = sorted(
            Path.home().glob(
                ".wdm/drivers/chromedriver/win64/*/chromedriver-win32/chromedriver.exe"
            )
        )
        if not candidatos:
            raise
        driver_path = str(candidatos[-1])
    driver = uc.Chrome(
        driver_executable_path=driver_path,
        options=opts,
    )
    driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": pasta_download,
    })
    return driver


def _salvar_sessao(driver: uc.Chrome, log_fn=print) -> None:
    """Persiste cookies do Spotfire em disco para reutilizar na próxima run."""
    try:
        os.makedirs(os.path.dirname(_COOKIES_FILE), exist_ok=True)
        cookies = driver.get_cookies()
        with open(_COOKIES_FILE, "w", encoding="utf-8") as f:
            json.dump(cookies, f)
        log_fn(f"  💾 Sessão salva ({len(cookies)} cookies).")
    except Exception as e:
        log_fn(f"  ⚠️  Não foi possível salvar sessão: {e}")


def _restaurar_sessao(driver: uc.Chrome, url: str, log_fn=print) -> bool:
    """
    Injeta cookies salvos via CDP (funciona em about:blank) e navega para a
    URL do dashboard. Retorna True se a sessão ainda for válida.
    """
    if not os.path.exists(_COOKIES_FILE):
        return False
    try:
        with open(_COOKIES_FILE, encoding="utf-8") as f:
            cookies = json.load(f)
        if not cookies:
            return False

        log_fn(f"  🔄 Tentando restaurar sessão anterior ({len(cookies)} cookies)...")
        for c in cookies:
            try:
                driver.execute_cdp_cmd("Network.setCookie", {
                    "name":     c["name"],
                    "value":    c["value"],
                    "domain":   c.get("domain", ""),
                    "path":     c.get("path", "/"),
                    "httpOnly": c.get("httpOnly", False),
                    "secure":   c.get("secure", False),
                })
            except Exception:
                pass

        if not url.startswith(("http://", "https://")):
            url = "http://" + url
        parsed_r = urlparse(url)
        hostname_r = parsed_r.hostname or ""

        driver.get(url)
        time.sleep(5)

        cur = driver.current_url.lower()
        if hostname_r in cur and "login" not in cur:
            log_fn(f"  ✅ Sessão restaurada — {driver.current_url}")
            return True

        log_fn("  ⚠️  Sessão expirada — necessário novo login.")
        return False
    except Exception as e:
        log_fn(f"  ⚠️  Erro ao restaurar sessão: {e}")
        return False


def _limpar_cookies_sso(driver: uc.Chrome, spotfire_host: str, log_fn=print) -> None:
    """Remove cookies de domínios externos (SSO/intranet) preservando sessão Spotfire."""
    try:
        todos = driver.execute_cdp_cmd("Network.getAllCookies", {}).get("cookies", [])
        removidos = 0
        for c in todos:
            domain = c.get("domain", "").lstrip(".")
            if spotfire_host not in domain:
                try:
                    driver.execute_cdp_cmd("Network.deleteCookies", {
                        "name": c["name"],
                        "domain": c.get("domain", ""),
                        "path": c.get("path", "/"),
                    })
                    removidos += 1
                except Exception:
                    pass
        if removidos:
            log_fn(f"  🔄 {removidos} cookie(s) de SSO/intranet removidos.")
    except Exception as e:
        log_fn(f"  ⚠️  Limpeza de cookies SSO falhou: {e}")


def _login_http(url: str, username: str, password: str, log_fn=print) -> tuple[dict, bool]:
    """
    Autentica no Spotfire via HTTP direto (form POST) e devolve os cookies
    de sessão. verify=False ignora o certificado corporativo.
    Spotfire usa o padrão Angular de CSRF: XSRF-TOKEN no cookie deve ser
    reenviado como header X-XSRF-TOKEN no POST.
    """
    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    login_url = f"{base}/spotfire/login.html"

    session = requests.Session()
    session.verify = False

    try:
        session.get(login_url, timeout=20)

        xsrf = session.cookies.get("XSRF-TOKEN", "")
        if xsrf:
            log_fn(f"  🔑 XSRF-TOKEN capturado ({xsrf[:8]}...)")

        resp = session.post(
            login_url,
            data={"j_username": username, "j_password": password},
            headers={"X-XSRF-TOKEN": xsrf} if xsrf else {},
            timeout=30,
            allow_redirects=True,
        )
        cookies = session.cookies.get_dict()
        ok = resp.status_code == 200 and "login.html" not in resp.url
        log_fn(f"  🔑 Login HTTP → {resp.status_code} | URL final: {resp.url} | ok={ok}")
        return cookies, ok
    except Exception as e:
        log_fn(f"  ⚠️  Login HTTP falhou: {e}")
        return {}, False


def _preencher_login_spotfire(driver: uc.Chrome, username: str, password: str,
                              log_fn=print) -> bool:
    """Preenche usuário/senha na tela de login do Spotfire e clica em Log in.

    Usa seletores tolerantes (name/id/placeholder/aria-label) porque a tela
    pode variar entre versões do Spotfire/SSO. Retorna True quando conseguiu
    preencher e clicar; a confirmação do login é feita por quem chama.
    """
    def _achar(seletores, timeout=20):
        fim = time.time() + timeout
        while time.time() < fim:
            for sel in seletores:
                try:
                    for el in driver.find_elements(*sel):
                        try:
                            if el.is_displayed() and el.is_enabled():
                                return el
                        except Exception:
                            pass
                except Exception:
                    pass
            time.sleep(0.5)
        return None

    user_sels = [
        (By.NAME, "j_username"),
        (By.NAME, "username"),
        (By.XPATH,
         "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
         "'abcdefghijklmnopqrstuvwxyz'),'user') "
         "or contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
         "'abcdefghijklmnopqrstuvwxyz'),'user') "
         "or contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
         "'abcdefghijklmnopqrstuvwxyz'),'user') "
         "or contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
         "'abcdefghijklmnopqrstuvwxyz'),'user')]"),
        (By.XPATH,
         "//label[contains(translate(normalize-space(.),"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'user')]"
         "//input | //label[contains(translate(normalize-space(.),"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'user')]"
         "/following::input[1]"),
        (By.CSS_SELECTOR, "input[type='email']"),
        (By.CSS_SELECTOR, "input[type='text']"),
    ]
    pass_sels = [
        (By.NAME, "j_password"),
        (By.CSS_SELECTOR, "input[type='password']"),
        (By.XPATH,
         "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
         "'abcdefghijklmnopqrstuvwxyz'),'pass')]"),
        (By.XPATH,
         "//label[contains(translate(normalize-space(.),"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'pass')]"
         "//input | //label[contains(translate(normalize-space(.),"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'pass')]"
         "/following::input[1]"),
    ]
    btn_sels = [
        (By.XPATH,
         "//*[self::button or self::a or self::input or @role='button']"
         "[contains(translate(normalize-space(.),"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'log in') "
         "or contains(translate(normalize-space(.),"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'login') "
         "or contains(translate(normalize-space(.),"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'entrar') "
         "or contains(translate(normalize-space(.),"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'sign in') "
         "or contains(translate(@value,"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'log in') "
         "or contains(translate(@value,"
         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'login')]"),
        (By.CSS_SELECTOR, "input[type='submit']"),
        (By.CSS_SELECTOR, "button[type='submit']"),
    ]

    campo_user = _achar(user_sels, timeout=20)
    if campo_user is None:
        log_fn("  ⚠️  Campo de usuário não encontrado na tela de login.")
        return False
    try:
        log_fn(
            "  🔎 Campo usuário: "
            f"<{campo_user.tag_name} name='{campo_user.get_attribute('name')}' "
            f"id='{campo_user.get_attribute('id')}'>"
        )
    except Exception:
        pass

    campo_pass = _achar(pass_sels, timeout=10)
    if campo_pass is None:
        log_fn("  ⚠️  Campo de senha não encontrado na tela de login.")
        return False
    try:
        log_fn(
            "  🔎 Campo senha: "
            f"<{campo_pass.tag_name} name='{campo_pass.get_attribute('name')}' "
            f"id='{campo_pass.get_attribute('id')}'>"
        )
    except Exception:
        pass

    def _preencher(el, valor):
        try:
            el.click()
        except Exception:
            pass
        try:
            el.clear()
        except Exception:
            pass
        try:
            el.send_keys(valor)
        except Exception:
            pass
        # Fallback para inputs controlados por React/Angular.
        try:
            atual = el.get_attribute("value") or ""
            if atual != valor:
                driver.execute_script(
                    "arguments[0].value=arguments[1];"
                    "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                    "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                    el, valor,
                )
        except Exception:
            pass

    log_fn("  🔐 Preenchendo usuário e senha...")
    _preencher(campo_user, username)
    _preencher(campo_pass, password)
    time.sleep(0.3)

    botao = _achar(btn_sels, timeout=5)
    if botao is not None:
        try:
            log_fn(
                "  🔎 Botão: "
                f"<{botao.tag_name} text='{(botao.text or '')[:30]}' "
                f"value='{botao.get_attribute('value')}'>"
            )
        except Exception:
            pass
        try:
            ActionChains(driver).move_to_element(botao).click().perform()
        except Exception:
            try:
                botao.click()
            except Exception:
                try:
                    driver.execute_script("arguments[0].click();", botao)
                except Exception:
                    pass
        log_fn("  🔐 Botão de login clicado.")
    else:
        # Fallback: submit do formulário; se não houver form, Enter na senha.
        try:
            driver.execute_script("arguments[0].form.submit();", campo_pass)
            log_fn("  🔐 Formulário de login enviado (fallback).")
        except Exception:
            try:
                campo_pass.send_keys(Keys.ENTER)
                log_fn("  🔐 Enter enviado no campo de senha.")
            except Exception:
                log_fn("  ⚠️  Botão de login não encontrado.")
    return True


def _fazer_login(driver: uc.Chrome, url: str, username: str, password: str,
                 log_fn=print) -> None:
    t0 = time.time()
    # Normaliza URL — garante que tem esquema http(s)://
    if not url.startswith(("http://", "https://")):
        url = "http://" + url
    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    login_url = f"{base}/spotfire/login.html"
    hostname  = parsed.hostname or ""

    def _logado(d: uc.Chrome) -> bool:
        """Retorna True quando estamos no Spotfire fora da tela de login."""
        cur = d.current_url.lower()
        return hostname in cur and "login" not in cur

    # ── Inicia em about:blank e tenta restaurar sessão salva ─────────────────
    driver.get("about:blank")
    time.sleep(1)

    if _restaurar_sessao(driver, url, log_fn):
        log_fn(f"  ⏱️ Login (sessão salva): {time.time() - t0:.1f}s")
        return

    # ── Tenta login HTTP direto (mais rápido que a tela) ─────────────────────
    if username and password:
        log_fn("  🔐 Tentando login HTTP direto...")
        cookies, ok = _login_http(url, username, password, log_fn)
        if ok and cookies:
            try:
                driver.get(base + "/")
                for nome, valor in cookies.items():
                    try:
                        driver.add_cookie({"name": nome, "value": valor})
                    except Exception:
                        pass
                driver.get(url)
                if _logado(driver):
                    log_fn(
                        f"  ✅ Login HTTP concluído — {driver.current_url} "
                        f"({time.time() - t0:.1f}s)"
                    )
                    _salvar_sessao(driver, log_fn)
                    return
            except Exception as e:
                log_fn(f"  ⚠️  Injeção dos cookies falhou: {e}")
        else:
            log_fn("  ⚠️  Login HTTP não confirmou; tentando pela tela.")

    # ── Sessão expirada/inexistente: abre tela de login ───────────────────────
    log_fn("  🌐 Abrindo tela de login do Spotfire...")
    driver.get(login_url)

    # Aguarda a página sair de about:blank (redirect SSO pode demorar ~1 min)
    try:
        WebDriverWait(driver, 120).until(
            lambda d: not d.current_url.startswith("about:")
        )
    except TimeoutException:
        pass

    log_fn(f"  🌐 URL: {driver.current_url}")

    # Se o SSO já logou (sessão Windows ativa), pode estar no dashboard direto
    if _logado(driver):
        log_fn("  ✅ Sessão Windows ativa — já no dashboard.")
        _salvar_sessao(driver, log_fn)
        log_fn(f"  ⏱️ Login (SSO): {time.time() - t0:.1f}s")
        return

    # ── Login automático com as credenciais da configuração ──────────────────
    if username and password:
        log_fn("  🔐 Tentando login automático...")
        if _preencher_login_spotfire(driver, username, password, log_fn):
            try:
                WebDriverWait(driver, 90).until(_logado)
            except TimeoutException:
                pass
            if _logado(driver):
                log_fn(f"  ✅ Login automático concluído — {driver.current_url}")
                _salvar_sessao(driver, log_fn)
                log_fn(f"  ⏱️ Login (automático): {time.time() - t0:.1f}s")
                time.sleep(3)
                return
            log_fn("  ⚠️  Login automático não confirmado; aguardando login manual.")
        else:
            log_fn("  ⚠️  Tela de login não reconhecida; aguardando login manual.")

    # ── Aguarda login manual ──────────────────────────────────────────────────
    try:
        driver.maximize_window()
    except Exception:
        pass

    log_fn("  🔐 Faça o login manualmente no Chrome que abriu.")
    log_fn("  ⏳ Aguardando login manual (5 min)...")
    try:
        WebDriverWait(driver, 300).until(_logado)
        log_fn(f"  ✅ Login detectado — {driver.current_url}")
        _salvar_sessao(driver, log_fn)
        log_fn(f"  ⏱️ Login (manual): {time.time() - t0:.1f}s")
    except TimeoutException:
        raise TimeoutError("Login no Spotfire não concluído em 5 minutos.")

    time.sleep(3)


def _aguardar_carregamento(driver: uc.Chrome) -> None:
    """Aguarda o dashboard Spotfire terminar de carregar."""
    time.sleep(8)
    try:
        WebDriverWait(driver, 60).until_not(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(@class,'loading') or contains(@class,'spinner')]")
            )
        )
    except TimeoutException:
        pass
    time.sleep(3)


def _achar_label(driver: uc.Chrome, nome: str):
    """
    Localiza o elemento de label de um filtro.
    Tenta: (1) match exato, (2) exato + ':', (3) starts-with no nó folha.
    Retorna o elemento ou None.
    """
    for xpath in [
        f"//*[normalize-space()='{nome}']",
        f"//*[normalize-space()='{nome}:']",
        f"//*[starts-with(normalize-space(),'{nome}') "
        f"and not(descendant::*[starts-with(normalize-space(),'{nome}')])]",
    ]:
        try:
            return driver.find_element(By.XPATH, xpath)
        except Exception:
            pass
    return None


def _clicar_all_em_filtro(driver: uc.Chrome, nome: str, scope=None) -> None:
    """
    Clica em '(All)' no filtro `nome` usando o eixo following:: do XPath.
    Encontra o label do filtro e pega o primeiro '(All)' que vem depois dele
    em ordem de documento — que é sempre o '(All)' do próprio filtro.
    """
    label = None
    if scope is not None:
        try:
            label = scope.find_element(
                By.XPATH,
                f".//*[starts-with(normalize-space(),'{nome}')]",
            )
        except Exception:
            pass
    if label is None:
        label = _achar_label(driver, nome)
    if label is None:
        return
    try:
        all_el = label.find_element(
            By.XPATH, "following::*[starts-with(normalize-space(),'(All)')][1]"
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", all_el)
        time.sleep(0.2)
        all_el.click()
        time.sleep(0.3)
    except Exception:
        pass


def _achar_container_painel(driver: uc.Chrome, nome: str):
    """
    Localiza o container do painel de filtros com header `nome` (ex: 'Filtros', 'Extração').
    Sobe na DOM a partir do header até encontrar um ancestral que contenha
    pelo menos 2 itens '(All)' — isso identifica o painel completo.
    Retorna o container ou None se não encontrado.
    """
    for xpath in [
        f"//*[normalize-space()='{nome}']",
        f"//*[normalize-space()='{nome}:']",
        f"//*[starts-with(normalize-space(),'{nome}') "
        f"and not(descendant::*[starts-with(normalize-space(),'{nome}')])]",
    ]:
        try:
            header = driver.find_element(By.XPATH, xpath)
        except Exception:
            continue
        el = header
        for _ in range(12):
            try:
                el = el.find_element(By.XPATH, "..")
                alls = el.find_elements(
                    By.XPATH, ".//*[starts-with(normalize-space(),'(All)')]"
                )
                if len(alls) >= 2:
                    return el
            except Exception:
                break
    return None


def _resetar_todos_filtros(driver: uc.Chrome, log_fn=print) -> None:
    """
    Reseta todos os filtros via reload da página.
    O Spotfire recarrega a análise com o estado publicado (todos em All).
    """
    log_fn("  🔄 Recarregando página para resetar filtros...")
    driver.refresh()
    _aguardar_carregamento(driver)
    log_fn("  ✅ Filtros resetados.")


def _achar_valor_no_filtro(all_el, val: str):
    """
    Acha o elemento de valor `val` relativo ao item '(All)' do mesmo filtro.
    Tenta primeiro como irmão direto de '(All)' (mesmo nível da lista),
    depois sobe até 3 níveis e busca descendentes.
    """
    try:
        return all_el.find_element(
            By.XPATH, f"following-sibling::*[normalize-space()='{val}']"
        )
    except Exception:
        pass
    el = all_el
    for _ in range(3):
        try:
            el = el.find_element(By.XPATH, "..")
            return el.find_element(By.XPATH, f".//*[normalize-space()='{val}']")
        except Exception:
            pass
    return None


def _selecionar_intervalo(
    driver: uc.Chrome,
    nome: str,
    valores: set[str],
    log_fn=print,
    scope=None,
    sort_key=None,
) -> None:
    """
    Seleciona um intervalo contíguo de valores no filtro `nome` usando
    click no primeiro item + Shift+click no último.
    Usa following::(All)[1] como âncora para achar os valores via
    _achar_valor_no_filtro, evitando pegar elementos de gráficos/outros painéis.
    """
    if not valores:
        return

    label = None
    if scope is not None:
        try:
            label = scope.find_element(By.XPATH, f".//*[normalize-space()='{nome}']")
        except Exception:
            pass
    if label is None:
        label = _achar_label(driver, nome)
    if label is None:
        log_fn(f"  ⚠️  Filtro '{nome}' não encontrado.")
        return

    all_el = None
    try:
        all_el = label.find_element(
            By.XPATH, "following::*[starts-with(normalize-space(),'(All)')][1]"
        )
    except Exception:
        log_fn(f"  ℹ️  '(All)' não encontrado para filtro '{nome}' — usando busca por container.")

    ordered = sorted(valores, key=sort_key)

    def _encontrar_valor(val: str):
        """
        Tenta 2 estratégias para achar o elemento de valor `val` no filtro.
        Estratégia 1: via âncora (All) — falha quando (All) não está no DOM
                      (scroll virtual do Spotfire empurra (All) para fora).
        Estratégia 2: sobe a partir do label até encontrar um container que
                      contenha o valor como descendente — funciona independente
                      de scroll.
        """
        el = _achar_valor_no_filtro(all_el, val) if all_el is not None else None
        if el is not None:
            return el
        # Busca direta subindo do label (até 6 níveis)
        container = label
        for _ in range(6):
            try:
                container = container.find_element(By.XPATH, "..")
                candidatos = container.find_elements(
                    By.XPATH, f".//*[normalize-space()='{val}']"
                )
                if candidatos:
                    return candidatos[0]
            except Exception:
                break
        return None

    primeiro_el = _encontrar_valor(ordered[0])
    if primeiro_el is None:
        log_fn(f"  ⚠️  '{ordered[0]}' não encontrado em filtro '{nome}'.")
        return
    driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", primeiro_el)
    time.sleep(0.2)
    primeiro_el.click()
    time.sleep(0.3)

    if len(ordered) > 1:
        ultimo_el = _encontrar_valor(ordered[-1])
        if ultimo_el is None:
            log_fn(f"  ⚠️  '{ordered[-1]}' não encontrado em filtro '{nome}'.")
            return
        driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", ultimo_el)
        time.sleep(0.2)
        ActionChains(driver).key_down(Keys.SHIFT).click(ultimo_el).key_up(Keys.SHIFT).perform()
        time.sleep(0.3)

    label_fim = ordered[-1] if len(ordered) > 1 else ordered[0]
    log_fn(f"  ✅ {nome}: {ordered[0]} → {label_fim}")


def _aplicar_filtro_data(
    driver: uc.Chrome,
    data_ini: datetime,
    data_fim: datetime,
    log_fn=print,
) -> None:
    log_fn("  📅 Aplicando filtros de data no Spotfire...")

    # Calcula anos e meses que cobrem o período
    anos: set[str] = set()
    meses: set[str] = set()
    cur = data_ini.replace(day=1)
    while cur <= data_fim:
        anos.add(str(cur.year))
        meses.add(_MESES_PT[cur.month])
        if cur.month == 12:
            cur = cur.replace(year=cur.year + 1, month=1)
        else:
            cur = cur.replace(month=cur.month + 1)

    # Reset via reload — aguarda a página recarregar completamente
    _resetar_todos_filtros(driver, log_fn)

    # Aguarda painel de filtros reaparecer após reload
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(normalize-space(),'Filtros')]")
            )
        )
    except TimeoutException:
        log_fn("  ⚠️  Painel de filtros não encontrado após reload.")
        return

    painel_filtros = _achar_container_painel(driver, "Filtros")
    painel_extracao = (
        _achar_container_painel(driver, "Extração")
        or _achar_container_painel(driver, "Extracao")
    )
    log_fn(
        f"  ℹ️  Painéis — Filtros: {'✅' if painel_filtros else '❌'}  "
        f"Extração: {'✅' if painel_extracao else '❌'}"
    )

    # ── Ano ───────────────────────────────────────────────────────────────────
    _selecionar_intervalo(driver, "Ano", anos, log_fn, scope=painel_filtros)
    time.sleep(1.5)

    # ── Mês ───────────────────────────────────────────────────────────────────
    # Shift+click do primeiro ao último mês em ordem de calendário.
    # Clicar no primeiro mês desejado já deseleciona o mês auto-selecionado
    # pelo Spotfire ao filtrar o Ano — não precisamos resetar o Mês antes.
    _selecionar_intervalo(
        driver, "Mês", meses, log_fn,
        scope=painel_filtros,
        sort_key=lambda m: _MESES_NUM.get(m, 99),
    )
    time.sleep(2)  # aguarda lista "Data Referência" atualizar no painel Extração

    # ── Data Referência (painel direito "Extração") ───────────────────────────
    # (All) só quando o período começa no dia 1 E termina no último dia do
    # mês final — ou seja, meses completos sem nenhum dia faltando nas bordas.
    # Para qualquer período parcial (incluindo 30/abr→01/mai) usa Shift+click.
    ultimo_dia = calendar.monthrange(data_fim.year, data_fim.month)[1]
    periodo_completo = data_ini.day == 1 and data_fim.day == ultimo_dia

    if periodo_completo:
        _clicar_all_em_filtro(driver, "Data Ref", scope=painel_extracao)
        log_fn("  📅 Data Referência: todas as datas do período.")
    else:
        # Período parcial: click na primeira data + Shift+click na última.
        # sort_key por data real — evita ordenação alfabética errada (ex:
        # "01/05/2026" < "30/04/2026" alfabeticamente mas é a data posterior).
        primeira_data = data_ini.strftime("%d/%m/%Y")
        ultima_data = data_fim.strftime("%d/%m/%Y")
        _selecionar_intervalo(
            driver, "Data Ref", {primeira_data, ultima_data}, log_fn,
            scope=painel_extracao,
            sort_key=lambda d: datetime.strptime(d, "%d/%m/%Y"),
        )

    time.sleep(2)
    log_fn(f"  📅 Período: {data_ini.strftime('%d/%m/%Y')} → {data_fim.strftime('%d/%m/%Y')}")


def _clicar_aba_tab_completa(driver: uc.Chrome, log_fn=print) -> None:
    wait = WebDriverWait(driver, 30)
    log_fn("  🗂️  Navegando para aba 'Tab Completa'...")
    aba = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//*[normalize-space()='Tab Completa']")
    ))
    aba.click()
    time.sleep(4)
    _aguardar_carregamento(driver)


def _ultimo_match(driver, xpath):
    """Retorna o ÚLTIMO elemento que casa o xpath (mais interno), ou None."""
    els = driver.find_elements(By.XPATH, xpath)
    return els[-1] if els else None


def _achar_visualizacao(driver: uc.Chrome, titulo_el) -> Any:
    try:
        return driver.execute_script(
            """var e=arguments[0];
            function cls(x){return (x&&x.className||'').toString().toLowerCase();}
            while(e && e!==document.body){
                var c=cls(e);
                // O título também tem 'visual' na classe; precisa subir para o
                // container real da visualização (sf-element-visual/sfc-table).
                if(c.indexOf('visual-title')>=0){ e=e.parentElement; continue; }
                if(c.indexOf('sf-element-visual')>=0||c.indexOf('sfc-table')>=0
                   ||c.indexOf('sfc-graphical-table')>=0||c.indexOf('chart')>=0
                   ||c.indexOf('grid')>=0||c.indexOf('table')>=0) return e;
                e=e.parentElement;
            }
            return null;""",
            titulo_el,
        )
    except Exception:
        return None


def _substituir_com_retry(origem: Path, destino: Path, log_fn=print,
                          tentativas: int = 60, espera: float = 2.0) -> None:
    """Move o arquivo baixado para o destino com retry (Excel/OneDrive/Chrome podem travar)."""
    origem = Path(origem)
    destino = Path(destino)
    for i in range(tentativas):
        try:
            os.replace(str(origem), str(destino))
            return
        except (PermissionError, OSError) as e:
            if i == tentativas - 1:
                raise RuntimeError(
                    f"Não foi possível substituir '{destino.name}'. "
                    f"Feche o arquivo no Excel/OneDrive e tente novamente: {e}"
                )
            if i == 0 or (i + 1) % 10 == 0:
                log_fn(
                    f"  ⚠️  Arquivo ainda em uso; aguardando liberar "
                    f"({i + 1}/{tentativas})..."
                )
            time.sleep(espera)


def _achar_titulo_visual(driver, texto: str):
    """Acha o título de uma visualização (ex.: a tabela 'Deslocamentos').

    O Spotfire renderiza o título como div.sf-element-visual-title. O texto
    pode ter espaços/NBSP, então usamos contains(normalize-space()). Se houver
    mais de um (páginas ocultas), fica com o visível mais abaixo na tela.
    """
    xpath = (
        "//div[contains(concat(' ',@class,' '),' sf-element-visual-title ')]"
        f"[contains(normalize-space(),'{texto}')]"
    )
    try:
        els = driver.find_elements(By.XPATH, xpath)
    except Exception:
        return None
    visiveis = []
    for e in els:
        try:
            if e.is_displayed():
                visiveis.append((e.rect.get("y", 0), e))
        except Exception:
            pass
    if not visiveis:
        return None
    return max(visiveis, key=lambda t: t[0])[1]


def _item_menu_visivel(driver, texto: str):
    """Item VISÍVEL de menu de contexto cujo texto é exatamente `texto`."""
    alvo = texto.strip().lower()
    ci = (
        "translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"
    )
    for classe in ("contextMenuItem", "contextMenuItemLabel"):
        xpath = (
            f"//*[contains(concat(' ',@class,' '),' {classe} ') "
            f"and {ci}='{alvo}']"
        )
        visiveis = []
        try:
            for e in driver.find_elements(By.XPATH, xpath):
                try:
                    if e.is_displayed():
                        visiveis.append(e)
                except Exception:
                    pass
        except Exception:
            pass
        if visiveis:
            return visiveis[-1]
    return None


def _esperar_item_menu(driver, texto: str, timeout: float = 10.0):
    """Espera o item de menu aparecer, tentando o frame atual e o principal."""
    fim = time.time() + timeout
    while time.time() < fim:
        item = _item_menu_visivel(driver, texto)
        if item is not None:
            return item
        try:
            driver.switch_to.default_content()
            item = _item_menu_visivel(driver, texto)
            if item is not None:
                return item
        except Exception:
            pass
        time.sleep(0.3)
    return None


def _exportar_tabela(
    driver: uc.Chrome,
    texto_titulo: str,
    pasta_download: str,
    nome_final: str,
    log_fn=print,
    timeout_download: int = TIMEOUT_DOWNLOAD,
    estabilidade: float = 2.0,
    seletor_alvo: str | None = None,
) -> Path:
    log_fn(f"  📤 Exportando '{texto_titulo}'...")
    # Captura snapshot da pasta ANTES de qualquer ação — evita race condition
    # onde o download termina antes de _aguardar_download começar a observar.
    antes_download = set(Path(pasta_download).iterdir())
    driver.switch_to.default_content()

    alvo = None

    # ── Estratégia 0: seletor CSS direto (ex.: #id46 da página de hoje) ───────
    if seletor_alvo:
        try:
            alvo = driver.find_element(By.CSS_SELECTOR, seletor_alvo)
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", alvo
            )
            time.sleep(1)
            log_fn(f"  ✅ Alvo por seletor: {seletor_alvo}")
        except Exception as e:
            raise RuntimeError(
                f"Seletor '{seletor_alvo}' não encontrado na página: {e}"
            )

    # ── Estratégia 1: título no DOM principal ─────────────────────────────────
    # Match EXATO: evita achar "Filtro Deslocamentos" quando buscamos "Deslocamentos",
    # ou qualquer outro elemento pai que contenha o texto como substring.
    if alvo is None and texto_titulo == "Deslocamentos":
        # Caminho específico da tabela de deslocamentos (título de visualização).
        # Evita o fallback "contains" pegar o filtro "Qtd Deslocamentos".
        try:
            titulo_el = WebDriverWait(driver, 60).until(
                lambda d: _achar_titulo_visual(d, texto_titulo)
            )
            log_fn("  ✅ Título da visualização 'Deslocamentos' encontrado.")
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", titulo_el
            )
            time.sleep(1)
            viz = _achar_visualizacao(driver, titulo_el)
            if viz is None:
                try:
                    viz = titulo_el.find_element(By.XPATH, "..")
                except Exception:
                    viz = titulo_el
            alvo = viz
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
        except TimeoutException:
            log_fn("  ⚠️  Título visual 'Deslocamentos' não apareceu; tentando busca genérica.")

    if alvo is None:
        try:
            # Pega o ÚLTIMO elemento com esse texto (ordem de documento = o mais interno /
            # folha do título), evitando casar o wrapper gigante cujo texto também é o título.
            # Observação: normalize-space(current()) NÃO existe fora de XSLT no Selenium.
            titulo_el = WebDriverWait(driver, 20).until(
                lambda d: _ultimo_match(
                    d,
                    f"//*[normalize-space()='{texto_titulo}' "
                    "and not(ancestor::*[contains(@class,'sf-element-filter') "
                    "or contains(@class,'sfc-filter-panel')])]",
                )
            )
            log_fn("  ✅ Título no DOM principal (exato).")
        except TimeoutException:
            # Fallback: o texto exato pode variar ("todas Colunas" vs "todas as Colunas").
            base = texto_titulo.replace(" todas as Colunas", "").replace(" todas Colunas", "")
            titulo_el = WebDriverWait(driver, 20).until(
                lambda d: _ultimo_match(
                    d,
                    f"//*[contains(normalize-space(), '{base}') "
                    "and not(ancestor::*[contains(@class,'sf-element-filter') "
                    "or contains(@class,'sfc-filter-panel')])]",
                )
            )
            log_fn(f"  ✅ Título no DOM principal (contains '{base}').")

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", titulo_el)
        time.sleep(1)
        viz = _achar_visualizacao(driver, titulo_el)
        if viz is not None:
            # sobe do título até o container da visualização e clica no centro dele
            alvo = viz
        else:
            # clica 5px ABAIXO da borda inferior do título — dentro do corpo da
            # visualização, nunca no page-tab (que ficava a +80px)
            rect = driver.execute_script("return arguments[0].getBoundingClientRect();", titulo_el)
            cx = rect["left"] + rect["width"] / 2
            cy = rect["bottom"] + 5
            alvo = driver.execute_script(
                "return document.elementFromPoint(arguments[0],arguments[1]);", cx, cy
            ) or titulo_el
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

    # ── Estratégia 2: busca dentro de cada iframe ─────────────────────────────
    if alvo is None:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        log_fn(f"  ℹ️  {len(iframes)} iframe(s) na página.")
        for i, ifr in enumerate(iframes):
            try:
                driver.switch_to.frame(ifr)
                els = driver.find_elements(
                    By.XPATH, f"//*[normalize-space()='{texto_titulo}']"
                )
                if els:
                    log_fn(f"  ✅ Título no iframe {i}.")
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", els[0])
                    time.sleep(1)
                    rect = driver.execute_script("return arguments[0].getBoundingClientRect();", els[0])
                    cx = rect["left"] + rect["width"] / 2
                    cy = rect["bottom"] + 80
                    alvo = driver.execute_script(
                        "return document.elementFromPoint(arguments[0],arguments[1]);", cx, cy
                    ) or els[0]
                    break
                driver.switch_to.default_content()
            except Exception as e:
                log_fn(f"    ⚠️  iframe {i}: {e}")
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass

    # ── Estratégia 3: fallback por coordenada de viewport ─────────────────────
    if alvo is None:
        log_fn("  ⚠️  Fallback: coordenada de viewport.")
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        size = driver.get_window_size()
        w, h = size["width"], size["height"]
        # "Tabela Completa" fica na faixa inferior da página (~82% de altura).
        # "Deslocamentos" fica no centro-direito (~45% de altura).
        if "Tabela" in texto_titulo:
            frac_x, frac_y = 0.50, 0.82
        else:
            frac_x, frac_y = 0.70, 0.45
        cx, cy = int(w * frac_x), int(h * frac_y)
        alvo = driver.execute_script(
            f"return document.elementFromPoint({cx},{cy});"
        ) or driver.find_element(By.TAG_NAME, "body")

    # Loga elemento alvo para diagnóstico
    try:
        tag = driver.execute_script("return arguments[0].tagName.toLowerCase();", alvo)
        cls = (driver.execute_script("return arguments[0].className;", alvo) or "")[:80]
        log_fn(f"  ℹ️  Right-click em <{tag} class='{cls}'>")
    except Exception:
        pass

    ActionChains(driver).context_click(alvo).perform()
    time.sleep(2)

    # ── Menu de contexto: Export → Export table (itens VISÍVEIS e exatos) ─────
    # Não usa contains('export') porque isso casava "Export to PDF"/itens
    # ocultos e acabava clicando em outra coisa (às vezes marcando uma linha).
    export_item = _esperar_item_menu(driver, "Export", timeout=10)
    if export_item is None:
        raise RuntimeError("Menu 'Export' não apareceu após right-click.")

    log_fn("  ✅ Item 'Export' visível encontrado.")
    ActionChains(driver).move_to_element(export_item).perform()
    time.sleep(1.2)

    export_table = _esperar_item_menu(driver, "Export table", timeout=12)
    if export_table is None:
        # Fallback: "Export table (without value formatting)".
        export_table = _esperar_item_menu(
            driver, "Export table (without value formatting)", timeout=5
        )
    if export_table is None:
        try:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        except Exception:
            pass
        raise RuntimeError("Item 'Export table' não apareceu no submenu.")

    try:
        log_fn(
            "  ✅ Clicando em 'Export table' "
            f"(<{export_table.tag_name} class='{export_table.get_attribute('class')}'>)."
        )
    except Exception:
        pass
    try:
        export_table.click()
    except Exception:
        driver.execute_script("arguments[0].click();", export_table)
    log_fn(f"  ⏳ Aguardando download de '{texto_titulo}'...")

    arquivo = _aguardar_download(
        pasta_download, log_fn, antes=antes_download,
        timeout=timeout_download, estabilidade=estabilidade,
    )
    destino = Path(pasta_download) / nome_final
    _substituir_com_retry(arquivo, destino, log_fn)
    log_fn(f"  ✅ Salvo: {nome_final}")
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    return destino


def _aguardar_download(pasta: str, log_fn=print, timeout: int = TIMEOUT_DOWNLOAD,
                       antes: set | None = None, estabilidade: float = 2.0) -> Path:
    """Aguarda o download terminar, checando a ESTABILIDADE do tamanho do arquivo.

    Resolve o caso de arquivos grandes (ex.: 55 MB): o código anterior
    devolvia o arquivo no primeiro instante em que aparecia, podendo estar
    pela metade. Agora exigimos que o st_size pare de crescer entre duas
    leituras (~2 s) antes de retornar. Para a base geral, use uma janela
    maior (ex.: 5 s), porque a exportação pode ser bem maior.
    """
    pasta_path = Path(pasta)
    if antes is None:
        antes = set(pasta_path.iterdir())
    inicio = time.time()
    ultimo: tuple[str, int] | None = None
    while time.time() - inicio < timeout:
        time.sleep(estabilidade)
        # Enquanto o Chrome estiver escrevendo um .crdownload NOVO, o arquivo
        # final ainda não está pronto (evita publicar base truncada). Um
        # .crdownload antigo, de execução anterior, é ignorado.
        if any(
            p.suffix == ".crdownload" and p not in antes
            for p in pasta_path.iterdir()
        ):
            continue
        agora = set(pasta_path.iterdir())
        novos = [
            f for f in (agora - antes)
            if f.suffix.lower() in _SUFIXOS_EXPORT
        ]
        if not novos:
            continue
        f = novos[0]
        try:
            sz = f.stat().st_size
        except OSError:
            continue
        if ultimo == (f.name, sz):
            return f
        ultimo = (f.name, sz)
    raise TimeoutError(f"Download Spotfire não concluído em {timeout}s")


def _tipo_export(caminho: Path) -> str:
    """Detecta o formato real do export pelas magic bytes (não pela extensão)."""
    try:
        with open(caminho, "rb") as f:
            magic = f.read(8)
        if magic.startswith(b"PK\x03\x04"):
            return "xlsx"
        if magic.startswith(b"\xd0\xcf\x11\xe0"):
            return "xls"
    except OSError:
        pass
    return "csv"


def _validar_export(caminho: Path, log_fn=print) -> None:
    """Confere se o arquivo exportado existe, tem tamanho e tem linhas de dados."""
    caminho = Path(caminho)
    if not caminho.exists():
        raise RuntimeError(f"Arquivo exportado não existe: {caminho}")
    tamanho = caminho.stat().st_size
    if tamanho <= 0:
        raise RuntimeError(f"Arquivo exportado está vazio: {caminho}")

    tipo = _tipo_export(caminho)
    if tipo in ("xlsx", "xls"):
        try:
            xl = pd.ExcelFile(str(caminho))
            if not xl.sheet_names:
                raise RuntimeError("planilha sem abas")
            df = xl.parse(xl.sheet_names[0], nrows=3)
            if df.empty or len(df.columns) == 0:
                raise RuntimeError("planilha sem linhas/colunas")
        except Exception as e:
            raise RuntimeError(f"Export xlsx inválido ({caminho.name}): {e}")
        log_fn(f"  OK Export validado: {caminho.name} ({tamanho} bytes, xlsx)")
        return

    # CSV: escolhe o encoding pelo BOM/cheiro do arquivo (evita decodificar
    # UTF-8 como UTF-16 e "validar" lixo).
    try:
        with open(caminho, "rb") as f:
            amostra = f.read(200)
    except OSError:
        amostra = b""
    if amostra.startswith((b"\xff\xfe", b"\xfe\xff")):
        encodings = ("utf-16",)
    elif amostra.startswith(b"\xef\xbb\xbf"):
        encodings = ("utf-8-sig",)
    elif amostra.count(b"\x00") > max(8, len(amostra) // 4):
        encodings = ("utf-16-le", "utf-16-be")
    else:
        encodings = ("utf-8", "latin-1")

    for enc in encodings:
        try:
            with open(caminho, "r", encoding=enc, errors="strict") as f:
                linhas = [f.readline() for _ in range(3)]
        except UnicodeError:
            continue
        if not linhas or not linhas[0].strip():
            continue
        primeira = linhas[0].strip().lower()
        if primeira.startswith("<html") or primeira.startswith("<!doctype"):
            raise RuntimeError("o export veio como HTML, não como tabela")
        if not any(l.strip() for l in linhas[1:]):
            raise RuntimeError("o export não tem linhas de dados")
        log_fn(
            f"  OK Export validado: {caminho.name} "
            f"({tamanho} bytes, csv/{enc})"
        )
        return
    raise RuntimeError(f"Export csv inválido ou vazio: {caminho.name}")


def _preparar_pasta_download(pasta_local: str | Path) -> Path:
    """Pasta exclusiva para downloads do Spotfire, limpa a cada execução.

    Evita que um CSV antigo da pasta Downloads (ou um `(1).csv` de execução
    anterior) seja confundido com o arquivo novo.
    """
    d = Path(pasta_local) / "_datahub_m300_download"
    d.mkdir(parents=True, exist_ok=True)
    for f in d.iterdir():
        if f.is_file():
            try:
                f.unlink()
            except OSError:
                pass
    return d


def _limpar_pasta_download(pasta: Path) -> None:
    p = Path(pasta)
    for f in p.iterdir():
        if f.is_file():
            try:
                f.unlink()
            except OSError:
                pass
    try:
        p.rmdir()
    except OSError:
        pass


def _normalizar_export(temp: Path, destino: Path, log_fn=print) -> Path:
    """Publica o export no nome final.

    Se o Spotfire baixou um xlsx e o destino é .csv, converte para CSV UTF-16
    (mesmo padrão dos exports do M300). Se já é CSV, move como está.
    """
    temp = Path(temp)
    destino = Path(destino)
    tipo = _tipo_export(temp)

    if tipo in ("xlsx", "xls") and destino.suffix.lower() == ".csv":
        log_fn("  Convertendo export xlsx para csv...")
        tmp_csv = destino.with_suffix(".tmp_csv.csv")
        try:
            df = pd.read_excel(str(temp), dtype=str)
            df.to_csv(str(tmp_csv), sep="\t", index=False, encoding="utf-16")
            _validar_export(tmp_csv, log_fn)
            temp.unlink(missing_ok=True)
            _substituir_com_retry(tmp_csv, destino, log_fn)
        except Exception:
            for p in (temp, tmp_csv):
                try:
                    p.unlink(missing_ok=True)
                except OSError:
                    pass
            raise
    else:
        _validar_export(temp, log_fn)
        _substituir_com_retry(temp, destino, log_fn)

    return destino


def _exportar_com_fallback(
    driver: uc.Chrome,
    titulos: list[str],
    pasta_download: str,
    nome_final: str,
    log_fn=print,
    timeout_download: int = TIMEOUT_DOWNLOAD,
    estabilidade: float = 2.0,
) -> Path:
    """Exporta tentando vários títulos de tabela e publica de forma atômica.

    O download vai para um nome temporário; só depois de validado ele
    substitui o arquivo final. Assim, uma falha não corrompe o arquivo
    anterior (importante para o SharePoint).
    """
    pasta = Path(pasta_download)
    destino = pasta / nome_final
    temp_nome = f".tmp_{nome_final}"
    temp = pasta / temp_nome
    if temp.exists():
        try:
            temp.unlink()
        except OSError:
            pass

    ultimo_erro = None
    try:
        for titulo in titulos:
            try:
                _exportar_tabela(
                    driver, titulo, str(pasta), temp_nome, log_fn,
                    timeout_download=timeout_download,
                    estabilidade=estabilidade,
                    seletor_alvo=(titulo if titulo.startswith("#") else None),
                )
                return _normalizar_export(temp, destino, log_fn)
            except Exception as e:
                ultimo_erro = e
                log_fn(f"  ⚠️  Export '{titulo}' falhou: {e}")
                try:
                    if temp.exists():
                        temp.unlink()
                except OSError:
                    pass
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                try:
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                except Exception:
                    pass
                time.sleep(1)
    finally:
        if temp.exists():
            try:
                temp.unlink()
            except OSError:
                pass

    raise RuntimeError(
        f"Não foi possível exportar '{nome_final}'. "
        f"Títulos tentados: {titulos}. Último erro: {ultimo_erro}"
    )


def baixar_scanner(
    cfg: SpotfireConfig,
    pasta_local: str,
    data_ini: datetime,
    data_fim: datetime,
    log_fn=print,
    stop_event: threading.Event | None = None,
) -> tuple[Path, Path]:
    """
    Autentica no Spotfire via HTTP, injeta cookie no browser e exporta
    os dois quadros da aba 'Tab Completa'.
    Retorna (path_scanner, path_deslocamentos).
    """
    pasta = str(Path(pasta_local).resolve())
    log_fn("🌐 Spotfire: abrindo sessão...")

    driver = _criar_driver(pasta)
    if stop_event:
        threading.Thread(target=_watchdog_driver, args=(driver, stop_event), daemon=True).start()

    try:
        _fazer_login(driver, cfg.url_scanner, cfg.username, cfg.password, log_fn)
        _aguardar_carregamento(driver)
        _aplicar_filtro_data(driver, data_ini, data_fim, log_fn)
        _clicar_aba_tab_completa(driver, log_fn)

        # Diagnóstico: confirma quais títulos estão no DOM (exact e contains)
        for titulo_diag in ("Tabela Completa todas as Colunas", "Deslocamentos",
                            "Filtro Deslocamentos"):
            n_exact = len(driver.find_elements(
                By.XPATH, f"//*[normalize-space()='{titulo_diag}']"
            ))
            n_contains = len(driver.find_elements(
                By.XPATH, f"//*[contains(normalize-space(),'{titulo_diag}')]"
            ))
            log_fn(f"  🔍 '{titulo_diag}': exact={n_exact}  contains={n_contains}")

        path_scanner = _exportar_tabela(
            driver,
            "Tabela Completa todas as Colunas",
            pasta,
            NOME_SCANNER,
            log_fn,
        )
        time.sleep(3)  # aguarda página estabilizar antes do segundo export
        path_deslocamentos = _exportar_tabela(
            driver,
            "Deslocamentos",
            pasta,
            NOME_DESLOCAMENTOS,
            log_fn,
        )
        return path_scanner, path_deslocamentos

    finally:
        try:
            driver.quit()
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────
# M300 — Scanner 4.0 - RJ (pasta /M300/ do Spotfire)
# Ref: spotfire_scanner_rj_elementos.md
# ─────────────────────────────────────────────────────────────────────────────

# Container estável (GUID) dos filtros esquerdo da análise. Pode mudar se a
# análise for editada no Spotfire — revalidar caso os filtros sumam.
_NOME_ANALISE_M300 = "/M300/Scanner 4.0 - RJ"

# Polo (UI DataHub) -> (Gerência, [Base(s) no filtro Spotfire])
# "Base" no Spotfire lista as UAMs (NITERÓI, SÃO GONÇALO, ...). Para polos que
# agregam mais de uma UAM, selecionamos múltiplos itens (ctrl+click).
_POLO_BASE_M300: dict[str, tuple[str | None, list[str]]] = {
    "Todos os polos": (None, ["(All)"]),
    "UTN": ("Norte", ["(All)"]),
    "UTS": ("Sul", ["(All)"]),
    "Niterói": ("Sul", ["NITERÓI"]),
    "São Gonçalo": ("Sul", ["SÃO GONÇALO"]),
    "Magé": ("Sul", ["MAGÉ"]),
    "Serrana": ("Sul", ["PETRÓPOLIS", "TERESÓPOLIS"]),
    "Sul": ("Sul", ["ANGRA", "RESENDE"]),
    "Campos": ("Norte", ["CAMPOS"]),
    "Lagos": ("Norte", ["ARARUAMA", "CABO FRIO"]),
    "Macaé": ("Norte", ["MACAÉ"]),
    "Noroeste": ("Norte", ["ITAPERUNA", "PÁDUA", "CANTAGALO"]),
}


def _url_analise_m300(url_base: str) -> str:
    """Constrói a URL da análise M300 a partir da configuração.

    Se cfg.url já aponta para a análise (contém 'analysis' ou 'file='), usa
    direto; caso contrário anexa o caminho padrão do M300.
    """
    if "analysis" in url_base or "file=" in url_base:
        return url_base
    return f"{url_base.rstrip('/')}/wp/analysis?file={_NOME_ANALISE_M300}"


def _url_produtividade(url_base: str) -> str:
    """Constrói a URL da análise de Produtividade (dia atual).

    Se `url_hoje` estiver configurada, o chamador usa direto. Esta função
    serve de fallback: aproveita o host da URL do Scanner e monta a análise
    de Produtividade com o parâmetro `file=`.
    """
    url_base = (url_base or "").strip()
    if not url_base:
        raise ValueError("URL do Spotfire não configurada.")
    if "Produtividade" in url_base:
        return url_base
    if not url_base.startswith(("http://", "https://")):
        url_base = "http://" + url_base
    parsed = urlparse(url_base)
    base = f"{parsed.scheme}://{parsed.netloc}"
    return (
        f"{base}/spotfire/wp/analysis"
        f"?file={quote(_NOME_ANALISE_PRODUTIVIDADE)}"
    )


def _garantir_painel_filtros(driver, log_fn=print, timeout: int = 45):
    """Aguarda o painel de filtros (.sfc-filter-panel) carregar; abre se necessário.

    O painel carrega de forma assíncrona após trocar de aba, então esperamos até
    `timeout` segundos antes de qualquer clique. Só tenta o botão toggle (id
    'Spotfire.FilterPanel') se o painel realmente não aparecer, e nunca quebra
    o fluxo por isso.
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.find_elements(By.CSS_SELECTOR, ".sfc-filter-panel")
        )
        return
    except Exception:
        pass
    log_fn("  🔧 Painel de filtros não detectado; tentando abrir...")
    for sel in (
        (By.ID, "Spotfire.FilterPanel"),
        (By.CSS_SELECTOR, "[id='Spotfire.FilterPanel']"),
        (By.CSS_SELECTOR, "#Spotfire\\.FilterPanel"),
    ):
        try:
            driver.find_element(*sel).click()
            time.sleep(2)
            if driver.find_elements(By.CSS_SELECTOR, ".sfc-filter-panel"):
                return
        except Exception:
            continue
    log_fn("  ⚠️  Não foi possível abrir o painel de filtros; seguindo sem ele.")


def _norm(s: str) -> str:
    """Normaliza para comparação: minúsculas + remoção de acentos."""
    return "".join(
        c for c in unicodedata.normalize("NFKD", s or "")
        if not unicodedata.combining(c)
    ).lower()


def _filtros_esquerda(driver):
    # No DOM real do Spotfire (M300) NÃO há iframe e o container #610ffb12 do mapa
    # não existe. Os filtros esquerdos são todos os span.sf-element-filter-content
    # (inclui spans de rótulo, que têm classe sf-element-filter-title e 0 itens).
    return driver.find_elements(By.CSS_SELECTOR, "span.sf-element-filter-content")


def _itens_filtro(filtro_el):
    return [
        (it.get_attribute("title") or it.text or "").strip()
        for it in filtro_el.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item")
    ]


_UAMS_M300 = [
    "angra", "araruama", "cabofrio", "cantagalo", "campos", "italo",
    "macae", "mage", "niteroi", "saogoncalo", "serrana", "lagos",
    "noroeste", "resende", "iteperuna", "padua", "petropolis", "teresopolis",
]
_MESES_M300 = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]


def _blob_filtro(filtro_el):
    """Texto-resumo do filtro (visível mesmo quando recolhido) + títulos dos itens.

    Quando o filtro está expandido, os valores aparecem em div.sf-element-list-box-item;
    quando recolhido, aparecem no texto do próprio span como '(All) N values VALOR1 VALOR2'.
    Juntar ambos torna a identificação robusta aos dois estados.
    """
    txt = filtro_el.text or filtro_el.get_attribute("textContent") or ""
    itens = " ".join(_itens_filtro(filtro_el))
    return _norm(txt + " " + itens)


_LABEL_FILTRO_ESQ = {
    "ano": "ano", "mês": "mês", "gerência": "gerência",
    "atuação": "atuação", "base": "base",
}


def _achar_filtro_esquerdo(driver, tipo: str, tentativas: int = 12):
    """Localiza um filtro esquerdo. Tenta primeiro pelo RÓTULO (sf-element-filter-title),
    que não depende de a lista estar expandida (após re-render o filtro pode aparecer
    recolhido, sem os itens no DOM). Fallback: pelo conteúdo. Repete algumas vezes
    porque a página re-renderiza após aplicar o Polo.
    """
    label = _LABEL_FILTRO_ESQ.get(tipo, tipo)
    ultimo_erro = ""
    for _ in range(tentativas):
        # 1) pelo rótulo (robusto ao recolhimento)
        for t in driver.find_elements(By.CSS_SELECTOR, ".sf-element-filter-title"):
            if _norm(t.text or "") == label:
                try:
                    card = t.find_element(
                        By.XPATH,
                        "ancestor::*[contains(concat(' ',@class,' '),'sf-element-filter')][1]",
                    )
                except Exception:
                    card = t
                cont = card.find_elements(By.CSS_SELECTOR, "span.sf-element-filter-content")
                return cont[0] if cont else card
        # 2) fallback pelo conteúdo
        try:
            for f in _filtros_esquerda(driver):
                cls = f.get_attribute("class") or ""
                if "sf-element-filter-title" in cls:
                    continue
                blob = _blob_filtro(f)
                if not blob:
                    continue
                tokens = blob.split()
                if tipo == "gerência":
                    if "norte" in tokens and "sul" in tokens:
                        return f
                elif tipo == "base":
                    if any(u in tokens for u in _UAMS_M300):
                        return f
                elif tipo == "ano":
                    anos = [tk for tk in tokens if tk.isdigit() and len(tk) == 4]
                    if anos and not any(m in tokens for m in _MESES_M300):
                        return f
                elif tipo == "mês":
                    if any(m in tokens for m in _MESES_M300):
                        return f
                elif tipo == "atuação":
                    if any(k in blob for k in ["ampla", "chip", "corte", "extra", "operacao",
                                              "cadastrar", "ligacao", "religacao", "inspecao", "troca"]):
                        return f
        except Exception as e:
            ultimo_erro = str(e)
        time.sleep(1)
    # diagnóstico: lista o que encontrou, para ajustar se o rótulo for diferente
    titulos = [_norm(t.text or "") for t in driver.find_elements(By.CSS_SELECTOR, ".sf-element-filter-title")]
    raise RuntimeError(
        f"Filtro '{tipo}' não encontrado à esquerda. Rótulos vistos: {titulos}. ({ultimo_erro})"
    )


def _achar_filtro_por_titulo(driver, titulo: str):
    """Localiza um filtro pelo título (ex.: 'Período'), com fallback por conteúdo.

    Diferente de `_achar_filtro_esquerdo`, serve para filtros que não fazem
    parte do conjunto fixo Ano/Mês/Gerência/Atuação/Base.
    """
    alvo = _norm(titulo).strip().rstrip(":")
    for t in driver.find_elements(By.CSS_SELECTOR, ".sf-element-filter-title"):
        if _norm(t.text or "").strip().rstrip(":") == alvo:
            try:
                card = t.find_element(
                    By.XPATH,
                    "ancestor::*[contains(concat(' ',@class,' '),'sf-element-filter')][1]",
                )
            except Exception:
                card = t
            cont = card.find_elements(By.CSS_SELECTOR, "span.sf-element-filter-content")
            cont = [
                c for c in cont
                if "sf-element-filter-title" not in (c.get_attribute("class") or "")
            ]
            return cont[0] if cont else card
    for f in _filtros_esquerda(driver):
        if alvo in _blob_filtro(f):
            return f
    return None


def _itens_all_visiveis(filtro):
    """Itens '(All)' visíveis do filtro (o Spotfire mantém cópias ocultas)."""
    itens = []
    try:
        for it in filtro.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
            titulo = (it.get_attribute("title") or it.text or "").strip()
            if not titulo.lower().startswith("(all)"):
                continue
            try:
                if it.is_displayed():
                    itens.append(it)
            except Exception:
                pass
    except Exception:
        pass
    return itens


def _selecionar_all_visivel(driver, filtro, log_fn=print) -> bool:
    """Seleciona o '(All)' VISÍVEL do filtro. Retorna True se já estava/ficou ok."""
    # Expande se a lista ainda não estiver renderizada.
    if not filtro.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
        try:
            filtro.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", filtro)
            except Exception:
                pass
        time.sleep(0.6)

    for it in _itens_all_visiveis(filtro):
        try:
            if _item_selecionado(it):
                log_fn("  ✓ Período já está em (All) — nada a fazer.")
                return True
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", it
            )
            time.sleep(0.2)
            ActionChains(driver).move_to_element(it).click().perform()
            time.sleep(0.5)
            if _item_selecionado(it):
                log_fn("  ✓ Período = (All) confirmado.")
            return True
        except Exception:
            continue

    # Fallback para DOM sem a classe de list-box: procura o texto visível
    # "(All) N values" dentro do próprio filtro.
    try:
        candidatos = filtro.find_elements(
            By.XPATH, ".//*[starts-with(normalize-space(),'(All)')]"
        )
        for it in candidatos:
            try:
                if not it.is_displayed():
                    continue
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});", it
                )
                time.sleep(0.2)
                ActionChains(driver).move_to_element(it).click().perform()
                time.sleep(0.5)
                log_fn("  ✓ Período = (All) clicado (fallback por texto).")
                return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _revelar_filtro_painel(driver, termo: str = "per", log_fn=print) -> bool:
    """Digita no campo 'Type to search filters' para revelar o filtro desejado.

    Algumas análises só renderizam o card do filtro depois da busca no painel.
    É o mesmo passo que o Codegen do Playwright registrou.
    """
    xpath_input = (
        "//input[contains(translate(@placeholder,"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'search filters') "
        "or contains(translate(@aria-label,"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'search filters')]"
    )
    try:
        campo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, xpath_input))
        )
        campo.clear()
        campo.send_keys(termo)
        time.sleep(0.7)
    except Exception:
        return False

    for seletor in (".FilterPanelMagnifier > svg", ".FilterPanelMagnifier"):
        try:
            driver.find_element(By.CSS_SELECTOR, seletor).click()
            time.sleep(0.7)
            break
        except Exception:
            try:
                driver.execute_script(
                    "arguments[0].click();",
                    driver.find_element(By.CSS_SELECTOR, seletor),
                )
                time.sleep(0.7)
                break
            except Exception:
                continue
    log_fn(f"  ℹ️  Busca de filtros aplicada ('{termo}').")
    return True


def _aplicar_filtro_periodo(driver, log_fn=print):
    """Aplica Período = (All) na análise de Produtividade (dia atual).

    Tenta primeiro o filtro pelo título; se não achar, usa o id conhecido
    como fallback (o id pode mudar se a análise for editada no Spotfire).
    Prioriza o item '(All)' VISÍVEL, porque o Spotfire mantém cópias ocultas.
    """
    _garantir_painel_filtros(driver, log_fn)

    filtro = _achar_filtro_por_titulo(driver, "Período")
    if filtro is None:
        # Em algumas análises o card só é renderizado após a busca do painel.
        if _revelar_filtro_painel(driver, "per", log_fn):
            filtro = _achar_filtro_por_titulo(driver, "Período")
    if filtro is None:
        try:
            filtro = driver.find_element(
                By.CSS_SELECTOR, f'[id="{_PERIODO_FILTRO_ID}"]'
            )
            log_fn("  ℹ️  Filtro 'Período' localizado pelo id de fallback.")
        except Exception:
            filtro = None

    if filtro is None:
        vistos = []
        try:
            vistos = [
                (t.text or "").strip()
                for t in driver.find_elements(
                    By.CSS_SELECTOR, ".sf-element-filter-title"
                )
            ]
        except Exception:
            pass
        raise RuntimeError(
            "Filtro 'Período' não encontrado na análise de Produtividade. "
            f"Rótulos vistos: {vistos}. "
            "Verifique se a análise foi alterada no Spotfire."
        )

    if _selecionar_all_visivel(driver, filtro, log_fn):
        return

    # Fallback para a lógica genérica de list-box (inclui rolagem virtual).
    try:
        if _filtro_tem_selecionado(filtro, "(All)"):
            log_fn("  ✓ Período já está em (All) — nada a fazer.")
            return
    except Exception:
        pass

    _clicar_item_filtro(driver, filtro, "(All)", log_fn)
    time.sleep(1)
    try:
        ok = _filtro_tem_selecionado(filtro, "(All)")
    except Exception:
        ok = False
    if ok:
        log_fn("  ✓ Período = (All) confirmado.")
    else:
        # Não falha: a classe de seleção do Spotfire pode variar; o download
        # ainda é validado depois (arquivo com dados).
        log_fn("  ⚠️  Período = (All) não confirmado pela classe; seguindo.")


def _aguardar_filtros(driver, log_fn=print, timeout: int = 40):
    """Aguarda os filtros esquerdos aparecerem (o painel carrega de forma assíncrona)."""
    for _ in range(timeout):
        fs = _filtros_esquerda(driver)
        if fs:
            return fs
        time.sleep(1)
    return _filtros_esquerda(driver)


def _aguardar_dados(driver, log_fn=print, timeout: int = 90):
    """Aguarda a tabela de dados popular os filtros (ex.: Gerência com Norte/Sul).

    Sem os valores carregados, os filtros aparecem vazios ('(All) 0 values') e não
    dá pra selecionar nada. Espera até o filtro Gerência ter itens de fato.
    """
    for _ in range(timeout):
        try:
            f = _achar_filtro_esquerdo(driver, "gerência")
            if f.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
                return
        except Exception:
            pass
        time.sleep(1)
    log_fn("  ⚠️  Timeout aguardando os filtros receberem valores (dados não carregaram?).")


def _achar_filtro_direito(driver, titulo: str):
    """Localiza o card do filtro pelo título DENTRO do painel direito (.sfc-filter-panel)."""
    painel = driver.find_element(By.CSS_SELECTOR, ".sfc-filter-panel")
    titulos = [e for e in painel.find_elements(By.CSS_SELECTOR, ".sf-element-filter-title")
               if e.text.strip() == titulo]
    if not titulos:
        raise RuntimeError(f"Filtro '{titulo}' não encontrado no painel direito.")
    # Card cujo título (sf-element-filter-title) é exatamente `titulo`, IGNORANDO
    # grupos (sf-element-filter-table-group) — que também contêm 'sf-element-filter'
    # e fariam o XPath anterior casar um container largo com vários filtros.
    xpath = (
        ".//div[contains(concat(' ', @class, ' '), 'sf-element-filter ') "
        "and not(contains(@class,'table-group')) "
        "and .//*[contains(@class,'sf-element-filter-title') and normalize-space(.)=%r]]"
        % titulo
    )
    return painel.find_element(By.XPATH, xpath)


def _m300_diagnosticar(driver, log_fn=print):
    """Imprime os filtros esquerdos identificados pelo conteúdo (texto-resumo + itens)."""
    lin = ["  🔬 Filtros esquerdos (por conteúdo):"]
    for tipo in ["ano", "mês", "gerência", "atuação", "base"]:
        try:
            f = _achar_filtro_esquerdo(driver, tipo)
            lin.append(f"    {tipo}: {_blob_filtro(f)[:120]}")
        except Exception as e:
            lin.append(f"    {tipo}: NÃO ENCONTRADO ({e})")
    try:
        painel = driver.find_element(By.CSS_SELECTOR, ".sfc-filter-panel")
        tit = [e.text.strip() for e in
               painel.find_elements(By.CSS_SELECTOR, ".sf-element-filter-title")]
        lin.append("PANEL: " + " / ".join(tit))
    except Exception as e:
        lin.append(f"PANEL: erro {e}")
    log_fn("\n".join(lin))


def _localizar_item_exato(filtro_el, valor: str):
    """Re-localiza o item pelo title exato (nova referência de nó), desviando de nós
    reciclados pela virtualização do Spotfire ao rolar a lista."""
    if valor.startswith("(All)"):
        for it in filtro_el.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
            if (it.get_attribute("title") or "").strip().lower().startswith("(all)"):
                return it
        return None
    try:
        return filtro_el.find_element(
            By.CSS_SELECTOR, f"div.sf-element-list-box-item[title='{valor}']"
        )
    except Exception:
        for it in filtro_el.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
            if (it.get_attribute("title") or it.text or "").strip() == valor:
                return it
        return None


def _clicar_item_filtro(driver, filtro_el, valor: str, log_fn=print, ctrl: bool = False):
    """Seleciona um valor no filtro (clique real do Selenium, com retry e verificação).

    Tenta até 3x; se o item não for encontrado a lista pode estar vazia momentaneamente
    (re-render em cascata). Após clicar, confere se o valor ficou selecionado (warn-only,
    pois o marcador de seleção do Spotfire pode variar).
    ctrl=True faz ctrl+click (seleção múltipla).
    """
    ultimo_erro = ""
    for _ in range(3):
        try:
            _clicar_item_filtro_tentativa(driver, filtro_el, valor, log_fn, ctrl)
        except (RuntimeError, StaleElementReferenceException) as e:
            ultimo_erro = str(e)
            if "não encontrado" in ultimo_erro:
                time.sleep(1.5)
                continue
            raise
        # verificação honesta da seleção (NÃO re-clica em caso de falha para evitar
        # toggling: se já estiver selecionado, um novo clique desmarcaria).
        try:
            ok = _filtro_tem_selecionado(filtro_el, valor)
            log_fn(f"    {'✓' if ok else '⚠️'} '{valor}' confirmado_selecionado={ok}")
        except Exception as e:
            log_fn(f"    ⚠️  não foi possível confirmar seleção de '{valor}': {e}")
        return
    raise RuntimeError(ultimo_erro)


def _item_selecionado(it) -> bool:
    """True se o item de list-box do Spotfire está marcado (selecionado)."""
    try:
        cls = (it.get_attribute("class") or "").lower()
        if "selected" in cls or "checked" in cls or "active" in cls:
            return True
        if (it.get_attribute("aria-selected") or "").lower() == "true":
            return True
        # Spotfire costuma ter um <input> de checkbox dentro do item
        for c in it.find_elements(By.CSS_SELECTOR, "input"):
            if c.get_attribute("checked") is not None:
                return True
    except Exception:
        pass
    return False


def _filtro_tem_selecionado(filtro_el, valor) -> bool:
    it = _localizar_item_exato(filtro_el, valor)
    if it is None:
        return False
    return _item_selecionado(it)


def _clicar_item_filtro_tentativa(driver, filtro_el, valor: str, log_fn=print, ctrl: bool = False):
    """Seleciona um valor rolando a lista (NÃO usa a busca do filtro — a busca, ao ser
    limpa, pode reverter a seleção no Spotfire). Re-localiza o item pelo title/texto
    exato logo antes de clicar para desviar de nós reciclados pela virtualização."""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", filtro_el)
    except Exception:
        pass
    time.sleep(0.2)
    # Expande se necessário (itens só existem no DOM quando expandido)
    if not filtro_el.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
        try:
            filtro_el.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", filtro_el)
            except Exception:
                pass
        time.sleep(0.6)

    # Procura o item: visível -> busca (se houver input) -> rola a lista virtualizada.
    alvo = _achar_item_na_lista(filtro_el, valor)
    if alvo is None and not valor.startswith("(All)"):
        try:
            busca = filtro_el.find_element(By.CSS_SELECTOR, "input.SearchInput")
            busca.clear()
            busca.send_keys(valor)
            time.sleep(1.2)
            alvo = _localizar_item_exato(filtro_el, valor)
        except Exception:
            alvo = None
    if alvo is None:
        alvo = _rolar_ate_achar(driver, filtro_el, valor)
    if alvo is None:
        disp = [(it.get_attribute("title") or it.text or "").strip() for it in
                filtro_el.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item")][:15]
        diag = ""
        try:
            sc = _achar_scroll_container(driver, filtro_el)
            if sc is not None:
                diag = (f" [scroll={sc.tag_name}.{sc.get_attribute('class')} "
                        f"ovf={driver.execute_script('return getComputedStyle(arguments[0]).overflowY', sc)} "
                        f"sh={sc.get_attribute('scrollHeight')} ch={sc.get_attribute('clientHeight')}]")
        except Exception as e:
            diag = f" [erro scroll: {e}]"
        raise RuntimeError(f"Filtro '{valor}' não encontrado. Itens: {disp}{diag}")
    # Re-localiza pelo title/texto exato logo antes de clicar (desvia de nó reciclado).
    alvo = _localizar_item_exato(filtro_el, valor) or alvo
    try:
        rotulo = (alvo.get_attribute("title") or alvo.text or "").strip() or valor
    except Exception:
        rotulo = valor
    # Clique com ActionChains (eventos de mouse reais/trusted). O Spotfire alterna a
    # seleção no mousedown, não no click, e ignora .click() via JS — por isso o
    # ActionChains (que faz press+release de verdade) é o que funciona.
    if ctrl:
        ActionChains(driver).key_down(Keys.CONTROL).click(alvo).key_up(Keys.CONTROL).perform()
    else:
        ActionChains(driver).move_to_element(alvo).click().perform()
    time.sleep(0.4)
    log_fn(f"    → clique em '{valor}' (alvo={rotulo})")


def _achar_item_na_lista(filtro_el, valor: str):
    """Procura o item pelo title (case/acento-insensível); '(All)' casa por prefixo."""
    v = _norm(valor)
    for it in filtro_el.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
        t = _norm(it.get_attribute("title") or it.text or "")
        if (valor.startswith("(All)") and t.startswith("(all)")) or (t == v):
            return it
    return None


def _achar_scroll_container(driver, filtro_el):
    """Retorna o elemento efetivamente rolável (viewport da lista virtualizada).

    Seletores óbvios costumam apontar para o conteúdo interno (que não rola de
    fato); o viewport real é o primeiro ancestral cujo scrollHeight > clientHeight.
    """
    js = """function vis(e){
               var s=getComputedStyle(e);
               var sc = (s.overflowY==='auto'||s.overflowY==='scroll'||s.overflow==='auto'||s.overflow==='scroll');
               return sc && e.scrollHeight>e.clientHeight+2;
            }
            var it=arguments[0].querySelector('div.sf-element-list-box-item')||arguments[0];
            var e=it; while(e&&e!==document.body){ if(vis(e)) return e; e=e.parentElement; }
            // fallback: qualquer ancestral com conteúdo rolável (mesmo que clipado)
            e=it; while(e&&e!==document.body){ if(e.scrollHeight>e.clientHeight+2) return e; e=e.parentElement; }
            return null;"""
    try:
        return driver.execute_script(js, filtro_el)
    except Exception:
        return None


def _rolar_ate_achar(driver, filtro_el, valor: str, max_passos: int = 80):
    """Rola a lista virtualizada (VirtualListBox) até achar o item pelo title/texto.

    Lists do Spotfire só renderizam na tela os itens visíveis; rolar o viewport
    real dispara a re-renderização da lista virtual. Vai ao fim (últimos valores)
    e, se não achar, volta ao topo e desce devagar. Se o scrollTop não mexer, usa
    foco + PAGE_DOWN como fallback (Spotfire reage a teclas).
    """
    it = _achar_item_na_lista(filtro_el, valor)
    if it:
        return it
    scroll = _achar_scroll_container(driver, filtro_el)
    if scroll is None:
        return None

    def _topo(v):
        try:
            driver.execute_script("arguments[0].scrollTop = arguments[1];", scroll, v)
        except Exception:
            pass

    def _atual():
        try:
            return driver.execute_script("return arguments[0].scrollTop;", scroll)
        except Exception:
            return 0

    # vai ao fim (últimos valores) e confere
    _topo(driver.execute_script("return arguments[0].scrollHeight;", scroll))
    time.sleep(0.5)
    it = _achar_item_na_lista(filtro_el, valor)
    if it:
        return it
    # volta ao topo e desce devagar
    _topo(0)
    time.sleep(0.4)
    for _ in range(max_passos):
        it = _achar_item_na_lista(filtro_el, valor)
        if it:
            return it
        antes = _atual()
        _topo(min(antes + 80, driver.execute_script("return arguments[0].scrollHeight;", scroll)))
        if _atual() == antes:  # scrollTop não mexeu -> tenta teclado
            try:
                driver.execute_script("arguments[0].focus();", scroll)
                scroll.send_keys(Keys.PAGE_DOWN)
            except Exception:
                pass
        time.sleep(0.2)
    return _achar_item_na_lista(filtro_el, valor)


def _abrir_menu_view(driver, log_fn=print) -> bool:
    """Abre o menu View do Spotfire (onde ficam os comandos de reset)."""
    for sel in ((By.XPATH, "//*[@title='View']"),
                (By.XPATH, "//*[normalize-space()='View']")):
        try:
            el = driver.find_element(*sel)
            try:
                el.click()
            except Exception:
                driver.execute_script("arguments[0].click();", el)
            time.sleep(0.8)
            return True
        except Exception:
            continue
    return False


def _resetar_spotfire(driver, log_fn=print) -> int:
    """Reseta TODOS os filtros e TODAS as marcações pelo menu View.

    Usa os comandos nativos do Spotfire:
      - Reset all markings  (limpa o '1 marked' que filtra a tabela de cima)
      - Reset all filters   (zera os filtros, inclusive Equipe)
      - Reset visible filters
    Não clica em itens de filtro, não rola listas e não seleciona nada.
    """
    comandos = (
        "Spotfire.ResetAllMarkings",
        "Spotfire.ResetAllFilters",
        "Spotfire.ResetVisibleFilters",
    )
    clicados = 0

    def _comando_visivel(cmd):
        try:
            return driver.find_element(
                By.CSS_SELECTOR, f"[id='{cmd}']"
            ).is_displayed()
        except Exception:
            return False

    for cmd in comandos:
        if not _comando_visivel(cmd):
            if not _abrir_menu_view(driver, log_fn):
                break
        try:
            el = driver.find_element(By.CSS_SELECTOR, f"[id='{cmd}']")
            try:
                el.click()
            except Exception:
                driver.execute_script("arguments[0].click();", el)
            clicados += 1
            log_fn(f"  ✓ {cmd}")
            time.sleep(1)
        except Exception as e:
            log_fn(f"  ⚠️  Comando '{cmd}' não encontrado: {e}")
            try:
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            except Exception:
                pass

    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    except Exception:
        pass

    if clicados == 0:
        # Fallback: botão visível do painel de filtros.
        try:
            btn = driver.find_element(By.CSS_SELECTOR, "#id83, .ResetButton")
            btn.click()
            clicados = 1
            log_fn("  ✓ Reset Visible Filters (fallback).")
        except Exception:
            log_fn("  ⚠️  Não foi possível resetar filtros/marcações; seguindo.")
    else:
        log_fn(
            f"  ✓ Reset Spotfire: {clicados} comando(s) executado(s) "
            "(filtros e marcações)."
        )
    return clicados


def _resetar_filtros_visiveis(driver, log_fn=print) -> int:
    """Clica nos botões de reset do painel e não seleciona mais nada.

    É o fluxo pedido: abrir a página, resetar TODOS os filtros para (All) e
    exportar. Usa clique via JS para não rolar a página nem interagir com
    itens de lista (o que estava fazendo o robô se perder nos filtros).
    """
    candidatos = []
    try:
        candidatos = driver.find_elements(By.CSS_SELECTOR, ".ResetFilters")
    except Exception:
        pass

    if not candidatos:
        # Fallback: ícones de reset/limpar dentro do painel de filtros.
        try:
            paineis = driver.find_elements(By.CSS_SELECTOR, ".sfc-filter-panel")
            escopo = paineis[0] if paineis else driver
            xpath = (
                "//*[contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                "'abcdefghijklmnopqrstuvwxyz'),'reset') "
                "or contains(translate(@aria-label,"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'reset') "
                "or contains(translate(@title,"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'limpar') "
                "or contains(translate(@aria-label,"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'limpar')]"
            )
            candidatos = escopo.find_elements(By.XPATH, xpath)
        except Exception:
            candidatos = []

    clicados = 0
    vistos = set()
    for b in candidatos:
        try:
            chave = getattr(b, "id", id(b))
            if chave in vistos:
                continue
            vistos.add(chave)
            try:
                b.click()
            except Exception:
                driver.execute_script("arguments[0].click();", b)
            clicados += 1
            time.sleep(0.3)
        except Exception:
            continue

    if clicados:
        log_fn(f"  ✓ reset de filtros: {clicados} botão(ões) clicado(s) (sem selecionar itens).")
    else:
        log_fn("  ⚠️  Nenhum botão de reset encontrado; exportando a visão atual.")
    return clicados


def _resetar_filtros_m300(driver, log_fn=print):
    """Restaura todos os filtros para (All).

    Usa o botão global 'reset visible filters' (.ResetFilters) do painel de filtros,
    que zera TUDO de uma vez — incluindo filtros de intervalo (Fim Calendário,
    Log In/Off etc.) que NÃO são list-box e não respondem ao clique em '(All)' da
    lista. Faz backup zerando também cada filtro esquerdo de list-box. Roda ANTES
    de aguardar os dados, para destravar a carga (a sessão restaurada pode trazer
    estado podre que deixa a tabela vazia).
    """
    _garantir_painel_filtros(driver, log_fn)
    # Reset global de todos os filtros visíveis (list-box E intervalo) numa só ação.
    # 'reset visible filters' zera inclusive o painel esquerdo (Ano/Mês/Gerência/...).
    for r in driver.find_elements(By.CSS_SELECTOR, ".ResetFilters"):
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", r)
            r.click()
            time.sleep(0.4)
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", r)
            except Exception:
                pass
    log_fn("    ✓ reset visible filters (global)")


def _aplicar_filtro_polo(driver, polo: str, log_fn=print):
    """Aplica Gerência, Atuação=(All) e Base — DEPOIS dos filtros de data.

    Ordem do usuário: reset -> Ano -> Mês -> Data Referência -> Gerência -> Atuação -> Base.
    A Atuação é deixada em (All) (sem restrição de tipo de serviço).
    """
    gerencia, bases = _POLO_BASE_M300.get(polo, _POLO_BASE_M300["Todos os polos"])
    if not gerencia:
        log_fn(f"  ⚠️  Polo '{polo}' sem mapeamento de Gerência/Base — pulando filtro de polo.")
        return
    # Re-localiza CADA filtro na hora do clique: o painel re-renderiza após cada
    # seleção e invalida as referências guardadas (StaleElementReferenceException).
    _clicar_item_filtro(driver, _achar_filtro_esquerdo(driver, "gerência"), gerencia, log_fn)
    time.sleep(0.8)
    try:
        _clicar_item_filtro(driver, _achar_filtro_esquerdo(driver, "atuação"), "(All)", log_fn)
    except Exception as e:
        log_fn(f"  ⚠️  Atuação não aplicada ({e})")
    time.sleep(0.8)
    if bases == ["(All)"]:
        _clicar_item_filtro(driver, _achar_filtro_esquerdo(driver, "base"), "(All)", log_fn)
    else:
        _clicar_item_filtro(driver, _achar_filtro_esquerdo(driver, "base"), bases[0], log_fn)
        for base in bases[1:]:
            _clicar_item_filtro(driver, _achar_filtro_esquerdo(driver, "base"), base, log_fn, ctrl=True)


def _janela_meses(meses: int = 2, hoje: datetime | None = None) -> list[tuple[int, int]]:
    """Lista [(ano, mês)] do mês atual e dos anteriores (padrão: 2 meses).

    Em janeiro, devolve [(2027, 1), (2026, 12)] — cobre a virada do ano.
    """
    hoje = hoje or datetime.now()
    cur = hoje.replace(day=1)
    alvos: list[tuple[int, int]] = []
    for _ in range(max(1, int(meses))):
        alvos.append((cur.year, cur.month))
        cur = (cur - timedelta(days=1)).replace(day=1)
    return alvos


def _clicar_filtro_com_retry(driver, tipo: str, valor: str, log_fn=print,
                             ctrl: bool = False, tentativas: int = 4) -> None:
    """Re-localiza o filtro e tenta de novo quando o Spotfire re-renderiza."""
    ultimo = None
    for i in range(tentativas):
        try:
            filtro = _achar_filtro_esquerdo(driver, tipo)
            _clicar_item_filtro(driver, filtro, valor, log_fn, ctrl=ctrl)
            return
        except (StaleElementReferenceException, RuntimeError) as e:
            ultimo = e
            log_fn(
                f"    ⚠️  Filtro '{tipo}' re-renderizou/instável; "
                f"tentando de novo ({i + 1}/{tentativas})..."
            )
            time.sleep(1.2)
    raise ultimo or RuntimeError(f"Filtro '{tipo}' não pôde ser selecionado.")


def _selecionar_janela_filtros(driver, alvos: list[tuple[int, int]],
                               log_fn=print) -> None:
    """Seleciona Ano/Mês no Spotfire para a janela informada.

    Deve ser chamado DEPOIS do reset total. Seleciona o mês atual e o anterior
    juntos (ctrl+click). Em janeiro, seleciona também o ano anterior.
    """
    anos = sorted({y for y, _ in alvos})
    meses: list[int] = []
    for _, m in alvos:
        if m not in meses:
            meses.append(m)

    for i, ano in enumerate(anos):
        _clicar_filtro_com_retry(
            driver, "ano", str(ano), log_fn, ctrl=(i > 0)
        )
        time.sleep(1)

    for i, mes in enumerate(meses):
        _clicar_filtro_com_retry(
            driver, "mês", _MESES_PT[mes], log_fn, ctrl=(i > 0)
        )
        time.sleep(1)

    log_fn(
        "  ✓ Janela selecionada: "
        + " + ".join(f"{_MESES_PT[m]}/{y}" for y, m in alvos)
    )


def _segmentar_por_mes(data_ini: datetime, data_fim: datetime) -> list[tuple[datetime, datetime, str]]:
    """Divide o intervalo em segmentos de 1 mês (cada mês = 1 export)."""
    segs: list[tuple[datetime, datetime, str]] = []
    cur = data_ini.replace(day=1)
    while cur <= data_fim:
        if cur.month == 12:
            proximo = cur.replace(year=cur.year + 1, month=1, day=1)
        else:
            proximo = cur.replace(month=cur.month + 1, day=1)
        ultimo = proximo - timedelta(days=1)
        seg_ini = max(data_ini, cur)
        seg_fim = min(data_fim, ultimo)
        segs.append((seg_ini, seg_fim, f"{cur.year}-{cur.month:02d}"))
        cur = proximo
    return segs


def _m300_aplicar_filtro_data(driver, data_ini: datetime, data_fim: datetime, log_fn=print):
    """Aplica os filtros de DATA na ordem: (All) já feito no reset ->
    Ano -> Mês -> dias no 'Data Referência' (range: 1º dia + shift+clique no último).

    O Polo (Gerência/Atuação/Base) é aplicado DEPOIS, em baixar_m300, para não
    fazer o efeito cascata esconder as opções de Data Referência.
    """
    _garantir_painel_filtros(driver, log_fn)
    # 1) Ano
    fa = _achar_filtro_esquerdo(driver, "ano")
    _clicar_item_filtro(driver, fa, str(data_ini.year), log_fn)
    # 2) Mês
    fm = _achar_filtro_esquerdo(driver, "mês")
    _clicar_item_filtro(driver, fm, _MESES_PT[data_ini.month], log_fn)

    # Data Referência: a busca do próprio filtro NÃO filtra de forma confiável
    # (a lista continua com todos os ~642 dias). Então varremos a lista virtualizada
    # e clicamos CADA dia do intervalo [data_ini, data_fim] via clique real
    # (ActionChains: o Spotfire alterna a seleção no mousedown, não no click JS).
    dr0 = _achar_filtro_direito(driver, "Data Referência")
    if not dr0.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
        dr0.click()
        time.sleep(0.6)

    dias = []
    d = data_ini
    while d <= data_fim:
        dias.append(d)
        d += timedelta(days=1)
    clicados = 0
    for i, dia in enumerate(dias):
        el = _achar_item_data(driver, _achar_filtro_direito(driver, "Data Referência"), dia)
        if not el:
            log_fn(f"    ⚠️  dia {dia:%d/%m/%Y} não encontrado na lista")
            continue
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.2)
        if i == 0:
            ActionChains(driver).move_to_element(el).click().perform()  # âncora
        else:
            ActionChains(driver).key_down(Keys.CONTROL).click(el).key_up(Keys.CONTROL).perform()
        clicados += 1
        time.sleep(0.2)
    if clicados == 0:
        log_fn(f"    ⚠️  nenhum dia selecionado em Data Referência (lista vazia?)")

    # Lê de VOLTA a seleção real no filtro para confirmar o intervalo.
    filtro_dr = _achar_filtro_direito(driver, "Data Referência")
    sel = _ler_range_dr(filtro_dr)
    if sel:
        log_fn(f"    ✓ Data Referência (filtro) = {sel[0]:%d/%m/%Y}..{sel[1]:%d/%m/%Y}")
        if sel[0].date() != data_ini.date() or sel[1].date() < data_fim.date():
            log_fn(f"    ⚠️  seleção real {sel[0]:%d/%m/%Y}..{sel[1]:%d/%m/%Y} "
                   f"(esperado início {data_ini:%d/%m/%Y}, fim até {data_fim:%d/%m/%Y})")
            _log_itens_dr(driver, filtro_dr, log_fn)
    else:
        log_fn(f"    ℹ️  Data Referência (esperado) = {data_ini:%d/%m/%Y}..{data_fim:%d/%m/%Y}")
        _log_itens_dr(driver, filtro_dr, log_fn)
    log_fn(f"    ✓ Ano={data_ini.year} Mês={_MESES_PT[data_ini.month]} | "
           f"Data Referência {data_ini:%d/%m/%Y}..{data_fim:%d/%m/%Y}")
    _aguardar_carregamento(driver)


def _achar_item_data(driver, filtro_el, alvo: datetime, log_fn=print, max_passos=200):
    """Acha o item de DATA (dd/mm/yyyy) num filtro virtualizado; rola a lista se preciso.

    Faz uma varredura nos DOIS sentidos (cima->baixo e baixo->cima) em passos de
    ~80% da altura visível, para garantir que o último item (seja no topo ou no
    fim da lista) seja renderizado pela virtualização.
    """
    v = alvo.strftime("%d/%m/%Y")

    def _visivel():
        for it in filtro_el.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
            t = (it.get_attribute("title") or it.text or "").strip()
            if t == v or t.startswith(v + " "):
                return it
        return None

    it = _visivel()
    if it:
        return it
    scroll = _achar_scroll_container(driver, filtro_el)
    if not scroll:
        return None

    def _passo(delta):
        driver.execute_script(
            "var s=arguments[0], d=arguments[1];"
            "var pass=Math.max(20, s.clientHeight*0.8);"
            "s.scrollTop = Math.min(Math.max(0, s.scrollTop + d*pass), s.scrollHeight);",
            scroll, delta,
        )
        time.sleep(0.12)

    # Varre de cima -> baixo
    driver.execute_script("arguments[0].scrollTop = 0;", scroll)
    time.sleep(0.3)
    for _ in range(max_passos):
        it = _visivel()
        if it:
            return it
        _passo(1)
    # Varre de baixo -> cima (lista em ordem descendente)
    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", scroll)
    time.sleep(0.3)
    for _ in range(max_passos):
        it = _visivel()
        if it:
            return it
        _passo(-1)
    return _visivel()


def _ler_range_dr(filtro_el) -> tuple[datetime, datetime] | None:
    """Lê os dias efetivamente SELECIONADOS no filtro Data Referência.

    Itens selecionados costumam carregar a classe 'selected' (ou 'sf-element-list-box-item selected').
    Retorna (min, max) ou None se nada legível for encontrado.
    """
    datas = []
    for it in filtro_el.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
        cls = (it.get_attribute("class") or "").lower()
        if "selected" in cls:
            t = (it.get_attribute("title") or it.text or "").strip()
            try:
                datas.append(datetime.strptime(t, "%d/%m/%Y"))
            except Exception:
                pass
    if not datas:
        return None
    return (min(datas), max(datas))


def _log_itens_dr(driver, filtro_el, log_fn=print):
    """Loga cada item da Data Referência com sua classe exata (para auditar a seleção real)."""
    try:
        for it in filtro_el.find_elements(By.CSS_SELECTOR, "div.sf-element-list-box-item"):
            t = (it.get_attribute("title") or it.text or "").strip()
            c = (it.get_attribute("class") or "")[:60]
            log_fn(f"      [DR] {t!r} class={c!r}")
    except Exception as e:
        log_fn(f"      ⚠️  _log_itens_dr: {e}")


def _validar_periodo_csv(path_csv: Path, data_ini: datetime, data_fim: datetime, log_fn=print):
    """Confere se a coluna 'Data Referência' (1ª coluna) está dentro do segmento."""
    try:
        df = pd.read_csv(str(path_csv), sep="\t", encoding="utf-16", nrows=200, dtype=str)
        col = df.columns[0]
        datas = pd.to_datetime(df[col].dropna(), format="%d/%m/%Y %H:%M:%S", errors="coerce")
        datas = datas.dropna()
        if datas.empty:
            log_fn("  ⚠️  Validação: sem datas para conferir.")
            return
        if datas.min().to_pydatetime() < data_ini or datas.max().to_pydatetime() > data_fim:
            raise RuntimeError(
                f"Período exportado ({datas.min()}..{datas.max()}) "
                f"não bate com o segmento ({data_ini}..{data_fim})."
            )
        log_fn(f"  ✅ Período OK: {datas.min()}..{datas.max()}")
    except Exception as e:
        raise RuntimeError(f"Falha na validação do período exportado: {e}")


def _merge_chunks_csv(chunks: list[Path], destino: Path, log_fn=print) -> Path:
    dfs = [pd.read_csv(str(c), sep="\t", dtype=str, encoding="utf-16") for c in chunks]
    df = pd.concat(dfs, ignore_index=True).drop_duplicates()
    df.to_csv(str(destino), sep="\t", index=False, encoding="utf-16")
    for c in chunks:
        try:
            c.unlink()
        except Exception:
            pass
    log_fn(f"  📦 Merge M300: {len(df)} linhas -> {destino.name}")
    return destino


def baixar_m300(
    cfg: "SpotfireConfig",
    pasta_local: str,
    data_ini: datetime,
    data_fim: datetime,
    log_fn=print,
    stop_event: threading.Event | None = None,
    polo: str = "Todos os polos",
) -> tuple[Path, Path]:
    """Baixa as duas tabelas do M300 (Scanner 4.0 - RJ) com segmentação mensal.

    Fluxo:
      1. login em cfg.url (cfg.login/cfg.senha)
      2. abre análise /M300/Scanner 4.0 - RJ -> aba 'Tab Completa'
      3. aplica filtro de Polo (Gerência + Base)
      4. para cada mês do intervalo: Ano+Mês + Data Referência=(All); exporta
         'Tabela Completa todas Colunas' e 'Deslocamentos'; valida o período
      5. faz merge dos chunks (dedup) -> 2 CSVs finais (UTF-16, tab)

    Retorna (path_tabela_completa, path_deslocamentos).
    """
    # M300 só carrega dados até o dia anterior ao atual (hoje - 1 dia). Capa o fim
    # para não pedir um dia que ainda não existe no servidor.
    ontem = (datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
             - timedelta(days=1))
    if data_fim > ontem:
        log_fn(f"  📅 M300 só tem dados até {ontem:%d/%m/%Y} (ontem); "
               f"fim ajustado de {data_fim:%d/%m/%Y} -> {ontem:%d/%m/%Y}")
        data_fim = ontem
    if data_ini > ontem:
        raise RuntimeError(
            f"Período inicial {data_ini:%d/%m/%Y} é maior que a data máxima "
            f"disponível no M300 ({ontem:%d/%m/%Y}). Nada a baixar."
        )
    pasta = str(Path(pasta_local).resolve())
    Path(pasta).mkdir(parents=True, exist_ok=True)
    driver = _criar_driver(pasta)
    if stop_event:
        threading.Thread(target=_watchdog_driver, args=(driver, stop_event), daemon=True).start()
    try:
        url = _url_analise_m300(cfg.url)
        _fazer_login(driver, url, cfg.login, cfg.senha, log_fn)
        _abrir_analise(driver, url, log_fn)
        _clicar_aba_tab_completa(driver, log_fn)
        _garantir_painel_filtros(driver, log_fn)
        _aguardar_filtros(driver, log_fn)  # painel de filtros carrega de forma assíncrona
        # Sessão restaurada traz o estado podre das rodadas anteriores (tabela vazia) ->
        # zera TUDO para (All) ANTES de aguardar os dados, para destravar a carga.
        log_fn("  🧹 Resetando filtros para (All)...")
        _resetar_filtros_m300(driver, log_fn)
        _aguardar_dados(driver, log_fn)    # aguarda a tabela popular os valores dos filtros
        try:
            _m300_diagnosticar(driver, log_fn)
        except Exception as e:
            log_fn(f"  ⚠️  Diagnóstico falhou: {e}")
        segs = _segmentar_por_mes(data_ini, data_fim)
        chunks_t, chunks_d = [], []
        for i, (d0, d1, label) in enumerate(segs, 1):
            log_fn(f"  🗓️  Segmento {i}/{len(segs)}: {label}")
            # Ordem: Ano -> Mês -> Data Referência (primeiro), depois Polo.
            _m300_aplicar_filtro_data(driver, d0, d1, log_fn)
            _aplicar_filtro_polo(driver, polo, log_fn)
            pt = _exportar_tabela(
                driver, "Tabela Completa todas Colunas", pasta,
                f"m300_tabela_{label}.csv", log_fn,
            )
            time.sleep(3)
            pd_ = _exportar_tabela(
                driver, "Deslocamentos", pasta,
                f"m300_desloc_{label}.csv", log_fn,
            )
            time.sleep(2)
            _validar_periodo_csv(pt, d0, d1, log_fn)
            chunks_t.append(pt)
            chunks_d.append(pd_)

        dest_t = _merge_chunks_csv(chunks_t, Path(pasta) / "m300_tabela_completa.csv", log_fn)
        dest_d = _merge_chunks_csv(chunks_d, Path(pasta) / "m300_deslocamentos.csv", log_fn)
        return dest_t, dest_d
    finally:
        try:
            driver.quit()
        except Exception:
            pass


def _abrir_analise(driver, url: str, log_fn=print):
    log_fn("  🔗 Abrindo análise M300...")
    driver.get(url)
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.sf-element-page-tab[title='Tab Completa']")
        )
    )
    _aguardar_carregamento(driver)


# ─────────────────────────────────────────────────────────────────────────────
# M300 — deslocamentos (arquivos para o SharePoint)
# ─────────────────────────────────────────────────────────────────────────────

def _norm_coluna(nome) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", str(nome))
        .encode("ascii", "ignore").decode().lower()
        if c.isalnum()
    )


def _gps_label(valor) -> str:
    """Converte o ícone de GPS do Painel para Sim/Não."""
    v = str(valor or "").strip().lower()
    if v == "cross,ffffab9e":
        return "Não"
    if v == "circle,ff37db7e":
        return "Sim"
    return ""


def _merge_painel_desloc(painel_path: Path, desloc_path: Path,
                         destino: Path, log_fn=print) -> Path:
    """Acrescenta ao deslocamento de hoje as colunas do Painel Acompanhamento.

    Merge por 'Equipe' (o Painel tem 1 linha por equipe). Colunas adicionadas:
    Inicio Calendário, Login, Fim calendário, Log Off, GPS (Cross=Não /
    Circle=Sim) e Ext Calendário.
    """
    painel = pd.read_csv(str(painel_path), sep="\t", encoding="utf-16", dtype=str)
    desloc = pd.read_csv(str(desloc_path), sep="\t", encoding="utf-16", dtype=str)

    mapa = {
        "iniciocalendario": "Inicio Calendário",
        "login": "Login",
        "fimcalendario": "Fim calendário",
        "logoff": "Log Off",
        "gps": "GPS",
        "extcalendario": "Ext Calendário",
    }
    pcols = {_norm_coluna(c): c for c in painel.columns}
    dcols = {_norm_coluna(c): c for c in desloc.columns}

    p_equipe = pcols.get("equipe")
    d_equipe = dcols.get("equipe")
    if not p_equipe or not d_equipe:
        raise RuntimeError("Coluna 'Equipe' não encontrada no Painel ou no Deslocamentos.")

    painel_sel = pd.DataFrame({
        "_equipe_key": painel[p_equipe].astype(str).str.strip().str.upper()
    })
    for chave_norm, nome_final in mapa.items():
        origem = pcols.get(chave_norm)
        if origem is None:
            log_fn(f"  ⚠️  Painel sem coluna '{nome_final}'; preenchendo vazio.")
            painel_sel[nome_final] = ""
        else:
            painel_sel[nome_final] = (
                painel[origem].fillna("").astype(str).str.strip()
            )

    # GPS: o Painel traz o ícone; converter para Sim/Não.
    gps_src = pcols.get("gps")
    if gps_src:
        gps_vals = painel[gps_src].fillna("").astype(str).str.strip()
        painel_sel["GPS"] = gps_vals.map(_gps_label)
        nao_mapeados = sorted({
            v for v in gps_vals
            if v and v.strip().lower() not in ("(empty)", "none", "nan")
            and _gps_label(v) == ""
        })
        if nao_mapeados:
            log_fn(f"  ⚠️  GPS sem mapeamento (ficaram vazios): {nao_mapeados[:10]}")

    duplicadas = int(painel_sel["_equipe_key"].duplicated().sum())
    if duplicadas:
        log_fn(f"  ⚠️  Painel com {duplicadas} equipe(s) duplicada(s); mantendo a última.")
    painel_sel = painel_sel.drop_duplicates("_equipe_key", keep="last")

    desloc = desloc.copy()
    desloc["_equipe_key"] = desloc[d_equipe].astype(str).str.strip().str.upper()
    for nome_final in mapa.values():
        if nome_final in desloc.columns:
            desloc = desloc.drop(columns=[nome_final])

    merged = desloc.merge(painel_sel, on="_equipe_key", how="left")
    merged = merged.drop(columns=["_equipe_key"])

    eq_desloc = set(desloc["_equipe_key"])
    eq_painel = set(painel_sel["_equipe_key"])
    casadas = len(eq_desloc & eq_painel)
    sem = len(eq_desloc - eq_painel)
    log_fn(
        f"  ✓ Merge Painel: {len(merged)} linha(s); "
        f"{casadas} equipe(s) casada(s), {sem} sem correspondência."
    )
    if sem:
        exemplos = sorted(eq_desloc - eq_painel)[:10]
        log_fn(f"    ℹ️  Sem painel (exemplos): {exemplos}")

    merged.to_csv(str(destino), sep="\t", index=False, encoding="utf-16")
    _validar_export(destino, log_fn)
    return destino


def _filtrar_janela(df: pd.DataFrame, alvos: list[tuple[int, int]],
                    log_fn=print, rotulo: str = "") -> pd.DataFrame:
    if "Data Referência" not in df.columns:
        raise RuntimeError(f"{rotulo}CSV sem a coluna 'Data Referência'.")
    dt = pd.to_datetime(df["Data Referência"], dayfirst=True, errors="coerce")
    alvo_set = set(alvos)
    dentro = pd.Series(
        [
            (int(a), int(m)) in alvo_set
            for a, m in zip(dt.dt.year.fillna(0), dt.dt.month.fillna(0))
        ],
        index=df.index,
    )
    descartadas = int((~dentro).sum())
    out = df[dentro].copy()
    log_fn(
        f"  🗓️  {rotulo}linhas na janela: {len(out)} "
        f"({descartadas} fora da janela descartada(s))."
    )
    return out


def baixar_bases_scanner_mensais(
    cfg: "SpotfireConfig",
    pasta_local: str,
    log_fn=print,
    stop_event: threading.Event | None = None,
    meses: int = 2,
) -> tuple[list[Path], str | None]:
    """Baixa a janela de 2 meses do Scanner e faz upsert nas bases mensais.

    Exporta DUAS tabelas:
      - Deslocamentos        -> bases_mensais/deslocamentos/YYYY-MM.xlsx (cru)
      - Tabela Completa      -> bases_mensais/tabela_completa/YYYY-MM.xlsx
        (com as 34 colunas verdes de tratamento do modelo)

    Retorna (lista de arquivos modificados, última data presente na base).
    """
    pasta = Path(pasta_local).resolve()
    pasta.mkdir(parents=True, exist_ok=True)
    temp_download = _preparar_pasta_download(pasta)
    alvos = _janela_meses(meses)
    t_total = time.time()
    log_fn(
        "  🗓️  Janela do Scanner: "
        + " + ".join(f"{m:02d}/{y}" for y, m in alvos)
    )

    driver = _criar_driver(str(temp_download))
    if stop_event:
        threading.Thread(
            target=_watchdog_driver, args=(driver, stop_event), daemon=True
        ).start()

    temp_desloc = None
    temp_tc = None
    try:
        url = _url_analise_m300(cfg.url)
        log_fn("🌐 Spotfire (bases mensais): abrindo análise do Scanner...")
        t = time.time()
        _fazer_login(driver, url, cfg.login, cfg.senha, log_fn)
        log_fn(f"  ⏱️ Login: {time.time() - t:.1f}s")
        t = time.time()
        _abrir_analise(driver, url, log_fn)
        _clicar_aba_tab_completa(driver, log_fn)
        log_fn(f"  ⏱️ Abertura + Tab Completa: {time.time() - t:.1f}s")

        # Sempre resetar ANTES de filtrar Ano/Mês.
        log_fn("  🧹 Resetando filtros e marcações pelo menu View...")
        t = time.time()
        _resetar_spotfire(driver, log_fn)
        _aguardar_carregamento(driver)
        log_fn(f"  ⏱️ Reset filtros/marcações: {time.time() - t:.1f}s")

        log_fn("  🗓️  Selecionando Ano/Mês da janela...")
        t = time.time()
        _selecionar_janela_filtros(driver, alvos, log_fn)
        _aguardar_carregamento(driver)
        log_fn(f"  ⏱️ Seleção Ano/Mês: {time.time() - t:.1f}s")

        log_fn("  📤 Exportando Deslocamentos (Scanner)...")
        t = time.time()
        temp_desloc = _exportar_com_fallback(
            driver,
            ["Deslocamentos"],
            str(temp_download),
            "_desloc_scanner.csv",
            log_fn,
            timeout_download=900,
            estabilidade=10.0,
        )
        log_fn(f"  ⏱️ Export Deslocamentos: {time.time() - t:.1f}s")
        log_fn("  📤 Exportando Tabela Completa (Scanner)...")
        t = time.time()
        temp_tc = _exportar_com_fallback(
            driver,
            ["Tabela Completa todas Colunas"],
            str(temp_download),
            "_tabela_completa_scanner.csv",
            log_fn,
            timeout_download=900,
            estabilidade=10.0,
        )
        log_fn(f"  ⏱️ Export Tabela Completa: {time.time() - t:.1f}s")
    finally:
        try:
            _resetar_spotfire(driver, log_fn)
        except Exception:
            pass
        try:
            driver.quit()
        except Exception:
            pass

    arquivos: list[Path] = []

    # 1) Deslocamentos crus (sem colunas de tratamento).
    t = time.time()
    df_desloc = pd.read_csv(
        str(temp_desloc), sep="\t", encoding="utf-16", dtype=str
    )
    df_desloc = _filtrar_janela(
        df_desloc, alvos, log_fn, "Deslocamentos: "
    )
    base_ate = None
    if not df_desloc.empty:
        try:
            dtmax = pd.to_datetime(
                df_desloc["Data Referência"], dayfirst=True, errors="coerce"
            ).max()
            if pd.notna(dtmax):
                base_ate = dtmax.strftime("%Y-%m-%d")
        except Exception:
            pass
    if df_desloc.empty:
        log_fn("  ⚠️  Deslocamentos sem linhas na janela.")
    else:
        arquivos += atualizar_base_mensal(
            df_desloc, "deslocamentos", str(pasta), "Data Referência", log_fn,
            cache_hash=True,
        )
    log_fn(f"  ⏱️ Deslocamentos (processamento+upsert): {time.time() - t:.1f}s")

    # 2) Tabela Completa com as 34 colunas verdes de tratamento.
    t = time.time()
    df_tc = tratar_m300(temp_tc, log_fn)
    log_fn(f"  ⏱️ Tratamento Tabela Completa: {time.time() - t:.1f}s")
    t = time.time()
    df_tc = _filtrar_janela(
        df_tc, alvos, log_fn, "Tabela Completa: "
    )
    if base_ate is None and not df_tc.empty:
        try:
            dtmax = pd.to_datetime(
                df_tc["Data Referência"], dayfirst=True, errors="coerce"
            ).max()
            if pd.notna(dtmax):
                base_ate = dtmax.strftime("%Y-%m-%d")
        except Exception:
            pass
    if df_tc.empty:
        log_fn("  ⚠️  Tabela Completa sem linhas na janela.")
    else:
        arquivos += atualizar_base_mensal(
            df_tc, "tabela_completa", str(pasta), "Data Referência", log_fn,
            cache_hash=True,
        )
    log_fn(f"  ⏱️ Tabela Completa (upsert): {time.time() - t:.1f}s")

    _limpar_pasta_download(temp_download)
    log_fn(f"  ⏱️ Tempo total Scanner: {time.time() - t_total:.1f}s")
    if arquivos:
        log_fn(
            "  ✅ Bases mensais atualizadas: "
            + ", ".join(f"{p.parent.name}/{p.name}" for p in arquivos)
        )
    else:
        log_fn("  ⚠️  Nenhuma base mensal alterada.")
    return arquivos, base_ate


def baixar_deslocamentos_hoje(
    cfg: "SpotfireConfig",
    pasta_local: str,
    log_fn=print,
    stop_event: threading.Event | None = None,
) -> Path:
    """Baixa a tabela de deslocamentos do DIA ATUAL (análise de Produtividade).

    Salva em `m300_deslocamentos_hoje.csv` (ou converte para CSV se o
    Spotfire exportar xlsx). O arquivo é validado antes de substituir o
    anterior.
    """
    pasta = Path(pasta_local).resolve()
    pasta.mkdir(parents=True, exist_ok=True)
    temp_download = _preparar_pasta_download(pasta)
    destino = None
    driver = _criar_driver(str(temp_download))
    if stop_event:
        threading.Thread(
            target=_watchdog_driver, args=(driver, stop_event), daemon=True
        ).start()

    try:
        url = (cfg.url_hoje or "").strip() or _url_produtividade(cfg.url)
        log_fn("🌐 Spotfire (hoje): abrindo análise de Produtividade...")
        _fazer_login(driver, url, cfg.login, cfg.senha, log_fn)
        # Recarregar a análise restaura o estado publicado (filtros em All)
        # sem clicar em nenhum filtro/item.
        log_fn("  🔄 Recarregando a análise para resetar os filtros...")
        driver.get(url)
        _aguardar_carregamento(driver)
        log_fn("  🧹 Resetando filtros e marcações pelo menu View...")
        _resetar_spotfire(driver, log_fn)
        _aguardar_carregamento(driver)

        painel_temp = None
        desloc_temp = None
        try:
            log_fn("  📤 Exportando Painel Acompanhamento...")
            painel_temp = _exportar_com_fallback(
                driver,
                ["Painel Acompanhamento"],
                str(temp_download),
                "_painel_hoje.csv",
                log_fn,
                timeout_download=600,
                estabilidade=5.0,
            )
            log_fn("  📤 Exportando Deslocamentos do dia...")
            desloc_temp = _exportar_com_fallback(
                driver,
                ["Deslocamentos"],
                str(temp_download),
                "_desloc_hoje.csv",
                log_fn,
                timeout_download=900,
                estabilidade=5.0,
            )
        finally:
            # Deixa a análise limpa (sem equipe marcada) após os downloads.
            try:
                _resetar_spotfire(driver, log_fn)
            except Exception:
                pass

        # Fecha o Chrome antes de ler/mover: ele pode manter os CSVs abertos.
        try:
            driver.quit()
        except Exception:
            pass
        driver = None

        # Junta as colunas do Painel no arquivo de deslocamentos de hoje.
        merged_temp = temp_download / "_m300_hoje_merged.csv"
        _merge_painel_desloc(painel_temp, desloc_temp, merged_temp, log_fn)

        final = pasta / NOME_M300_HOJE
        _substituir_com_retry(merged_temp, final, log_fn)
        destino = final
        log_fn(f"  ✅ Salvo: {destino.name} ({destino.stat().st_size} bytes)")
        return destino
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        if destino is not None and destino.exists():
            _limpar_pasta_download(temp_download)


def baixar_deslocamentos_geral(
    cfg: "SpotfireConfig",
    pasta_local: str,
    log_fn=print,
    stop_event: threading.Event | None = None,
) -> Path:
    """Baixa a tabela de deslocamentos GERAL (todos os dias, sem filtros).

    Usa a análise do Scanner 4.0 - RJ, reseta TODOS os filtros para (All) e
    exporta `Deslocamentos` para `m300_deslocamentos_geral.csv`. Não aplica
    filtro de data nem de polo: é a base histórica completa disponível.
    """
    pasta = Path(pasta_local).resolve()
    pasta.mkdir(parents=True, exist_ok=True)
    temp_download = _preparar_pasta_download(pasta)
    destino = None
    driver = _criar_driver(str(temp_download))
    if stop_event:
        threading.Thread(
            target=_watchdog_driver, args=(driver, stop_event), daemon=True
        ).start()

    try:
        url = _url_analise_m300(cfg.url)
        log_fn("🌐 Spotfire (geral): abrindo análise do Scanner...")
        _fazer_login(driver, url, cfg.login, cfg.senha, log_fn)
        # Abrir a análise de novo restaura o estado publicado (filtros em All),
        # sem clicar em filtros. Depois basta ir para a aba 'Tab Completa'.
        log_fn("  🔄 Recarregando a análise para resetar os filtros...")
        _abrir_analise(driver, url, log_fn)
        _clicar_aba_tab_completa(driver, log_fn)
        log_fn("  🧹 Resetando filtros e marcações pelo menu View...")
        _resetar_spotfire(driver, log_fn)
        _aguardar_carregamento(driver)

        temp = None
        try:
            temp = _exportar_com_fallback(
                driver,
                ["Deslocamentos"],
                str(temp_download),
                NOME_M300_GERAL,
                log_fn,
                timeout_download=TIMEOUT_DOWNLOAD_GERAL,
                estabilidade=15.0,
            )
        finally:
            # Deixa a análise limpa (sem equipe marcada) após o download.
            try:
                _resetar_spotfire(driver, log_fn)
            except Exception:
                pass

        # Fecha o Chrome antes de mover: em arquivos grandes o Chrome pode
        # manter o CSV aberto e travar o rename no Windows.
        try:
            driver.quit()
        except Exception:
            pass
        driver = None

        destino = pasta / NOME_M300_GERAL
        _substituir_com_retry(temp, destino, log_fn)
        log_fn(f"  ✅ Salvo: {destino.name} ({destino.stat().st_size} bytes)")
        return destino
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        if destino is not None and destino.exists():
            _limpar_pasta_download(temp_download)
