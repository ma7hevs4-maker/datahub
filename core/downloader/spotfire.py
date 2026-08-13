import json
import os
import threading
import time
import urllib3
import calendar
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

import requests
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from config import SpotfireConfig

NOME_SCANNER = "scanner_tabela_completa.xlsx"
NOME_DESLOCAMENTOS = "scanner_deslocamentos.xlsx"
TIMEOUT_DOWNLOAD = 300

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
    driver = uc.Chrome(
        driver_executable_path=ChromeDriverManager().install(),
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


def _fazer_login(driver: uc.Chrome, url: str, username: str, password: str,
                 log_fn=print) -> None:
    # Normaliza URL — garante que tem esquema http(s)://
    if not url.startswith(("http://", "https://")):
        url = "http://" + url
    parsed = urlparse(url)
    login_url = f"{parsed.scheme}://{parsed.netloc}/spotfire/login.html"
    hostname  = parsed.hostname or ""

    def _logado(d: uc.Chrome) -> bool:
        """Retorna True quando estamos no Spotfire fora da tela de login."""
        cur = d.current_url.lower()
        return hostname in cur and "login" not in cur

    # ── Inicia em about:blank e tenta restaurar sessão salva ─────────────────
    driver.get("about:blank")
    time.sleep(1)

    if _restaurar_sessao(driver, url, log_fn):
        return

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
        return

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


def _exportar_tabela(
    driver: uc.Chrome,
    texto_titulo: str,
    pasta_download: str,
    nome_final: str,
    log_fn=print,
) -> Path:
    log_fn(f"  📤 Exportando '{texto_titulo}'...")
    # Captura snapshot da pasta ANTES de qualquer ação — evita race condition
    # onde o download termina antes de _aguardar_download começar a observar.
    antes_download = set(Path(pasta_download).iterdir())
    driver.switch_to.default_content()

    try:
        driver.save_screenshot(str(Path(pasta_download) / f"debug_pre_{nome_final}.png"))
        log_fn(f"  📸 debug_pre_{nome_final}.png")
    except Exception:
        pass

    alvo = None

    # ── Estratégia 1: título no DOM principal ─────────────────────────────────
    # Match EXATO: evita achar "Filtro Deslocamentos" quando buscamos "Deslocamentos",
    # ou qualquer outro elemento pai que contenha o texto como substring.
    try:
        titulo_el = WebDriverWait(driver, 20).until(EC.presence_of_element_located((
            By.XPATH, f"//*[normalize-space()='{texto_titulo}']",
        )))
        log_fn("  ✅ Título no DOM principal.")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", titulo_el)
        time.sleep(1)
        rect = driver.execute_script("return arguments[0].getBoundingClientRect();", titulo_el)
        cx = rect["left"] + rect["width"] / 2
        viewport_h = driver.execute_script("return window.innerHeight;")
        cy = min(rect["bottom"] + 80, viewport_h - 10)
        alvo = driver.execute_script(
            "return document.elementFromPoint(arguments[0],arguments[1]);", cx, cy
        )
        if alvo is None:
            # Título muito perto da borda inferior — usa o próprio título
            alvo = titulo_el
        else:
            tag = driver.execute_script("return arguments[0].tagName.toLowerCase();", alvo)
            if tag == "iframe":
                ir = driver.execute_script("return arguments[0].getBoundingClientRect();", alvo)
                driver.switch_to.frame(alvo)
                alvo = driver.execute_script(
                    "return document.elementFromPoint(arguments[0],arguments[1]);",
                    cx - ir["left"], cy - ir["top"],
                ) or alvo
    except TimeoutException:
        log_fn("  ⚠️  Título não no DOM principal — buscando nos iframes...")

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

    try:
        driver.save_screenshot(str(Path(pasta_download) / f"debug_menu_{nome_final}.png"))
        log_fn(f"  📸 debug_menu_{nome_final}.png")
    except Exception:
        pass

    # ── Localiza item "Export" — tenta frame atual e depois principal ─────────
    xpath_export = "//*[normalize-space()='Export' or normalize-space()='Exportar']"
    export_item = None
    for ctx in ("atual", "principal"):
        if ctx == "principal":
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
        try:
            export_item = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.XPATH, xpath_export))
            )
            log_fn(f"  ✅ 'Export' encontrado (frame {ctx}).")
            break
        except TimeoutException:
            pass

    if export_item is None:
        raise RuntimeError("Menu 'Export' não apareceu após right-click.")

    ActionChains(driver).move_to_element(export_item).perform()
    time.sleep(1)

    export_table = WebDriverWait(driver, 15).until(EC.element_to_be_clickable(
        (By.XPATH, "//*[normalize-space()='Export table' or normalize-space()='Exportar tabela']")
    ))
    export_table.click()
    log_fn(f"  ⏳ Aguardando download de '{texto_titulo}'...")

    arquivo = _aguardar_download(pasta_download, log_fn, antes=antes_download)
    destino = Path(pasta_download) / nome_final
    os.replace(str(arquivo), str(destino))
    log_fn(f"  ✅ Salvo: {nome_final}")
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    return destino


def _aguardar_download(pasta: str, log_fn=print, timeout: int = TIMEOUT_DOWNLOAD,
                       antes: set | None = None) -> Path:
    """Aguarda qualquer arquivo novo (não .crdownload) aparecer na pasta."""
    pasta_path = Path(pasta)
    if antes is None:
        antes = set(pasta_path.iterdir())
    inicio = time.time()
    while time.time() - inicio < timeout:
        time.sleep(2)
        agora = set(pasta_path.iterdir())
        novos = agora - antes
        novos_reais = [f for f in novos if not f.suffix == ".crdownload"]
        if novos_reais:
            return novos_reais[0]
    raise TimeoutError(f"Download Spotfire não concluído em {timeout}s")


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
