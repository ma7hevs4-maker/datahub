import os
import re
import json
import time
import shutil
import threading
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from config import OperviewConfig

# ── Constantes ────────────────────────────────────────────────────────────────
NOME_ARQUIVO = "extracao_incidencias_operview.xlsx"
TIMEOUT_DOWNLOAD = 600  # sistema lento: até 10 min esperando o arquivo
TIMEOUT_LOGIN = 900     # até 15 min aguardando login manual
SHEET_NAME = "Relatorio_de_Incidencias"
URL_PADRAO = "https://operview-rj.enel.com/"
FMT_DATA = "%d/%m/%Y"


class OperviewNaoImplementado(Exception):
    """Levantado quando falta o arquivo de treinamento de cliques do Operview."""


class _FalhaServidor(Exception):
    pass


# ── Driver (visível — o login é manual) ───────────────────────────────────────
def _criar_driver(pasta_download: str) -> webdriver.Chrome:
    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_experimental_option("prefs", {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_settings.popups": 0,
        "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
    })
    try:
        svc = Service(ChromeDriverManager(cache_valid_range=365).install())
        driver = webdriver.Chrome(service=svc, options=opts)
    except Exception:
        driver = webdriver.Chrome(options=opts)
    time.sleep(0.5)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return driver


# ── Captura dos cliques do Operview (embedida; fluxo comum, sem modo de treino) ─
_CAPTURA_JSON = r"""
[
  {
    "html": "<span _ngcontent-ng-c3230066810=\"\" mattooltip=\"Menu\" mattooltipposition=\"right\" class=\"mat-mdc-tooltip-trigger enel-icon hamburger cursor-pointer\" aria-describedby=\"cdk-describedby-message-ng-1-2\" cdk-describedby-host=\"ng-1\"></span>",
    "id": "",
    "name": "",
    "placeholder": "",
    "tag": "span",
    "text": "",
    "type": "coord",
    "x": 33,
    "y": 38,
    "passo": "abrir_lateral",
    "instrucao": "Clique na barra lateral para expandi-la (botão de menu/hamburguer)"
  },
  {
    "html": "<mat-panel-title _ngcontent-ng-c3230066810=\"\" class=\"mat-expansion-panel-header-title neutrals-white ng-tns-c2690051721-3\"><span _ngcontent-ng-c3230066810=\"\" class=\"enel-icon document cursor-pointer bg-neutrals-white\"></span><span _ngcontent-ng-c3230066810=\"\" style=\"margin-left: 0.5em;\"> Auditoria</span></mat-panel-title>",
    "id": "",
    "name": "",
    "placeholder": "",
    "tag": "mat-panel-title",
    "text": "Auditoria",
    "type": "coord",
    "x": 125,
    "y": 188,
    "passo": "menu_auditoria",
    "instrucao": "Clique no menu 'Auditoria' (já expandido)"
  },
  {
    "html": "<span _ngcontent-ng-c3230066810=\"\">Consulta Incidência</span>",
    "id": "",
    "name": "",
    "placeholder": "",
    "tag": "span",
    "text": "Consulta Incidência",
    "type": "coord",
    "x": 126,
    "y": 385,
    "passo": "consulta_incidencia",
    "instrucao": "Clique em 'Consulta Incidência'"
  },
  {
    "html": "<span class=\"mat-mdc-select-placeholder mat-mdc-select-min-line ng-tns-c3393473648-23 ng-star-inserted\"></span>",
    "id": "",
    "name": "",
    "placeholder": "",
    "tag": "span",
    "text": "",
    "type": "coord",
    "x": 841,
    "y": 155,
    "passo": "filtro_polo",
    "instrucao": "Clique no filtro 'Polo' (selecione o polo desejado ou 'Todos os polos')"
  },
  {
    "html": "<input _ngcontent-ng-c1737117349=\"\" matinput=\"\" placeholder=\"Selecione uma data\" class=\"mat-mdc-input-element mat-datepicker-input ng-tns-c1205077789-20 ng-untouched ng-pristine ng-valid mat-mdc-form-field-input-control mdc-text-field__input cdk-text-field-autofill-monitored\" aria-describedby=\"mat-mdc-hint-0\" id=\"mat-input-0\" aria-invalid=\"false\" aria-required=\"false\" aria-haspopup=\"dialog\" data-mat-calendar=\"mat-datepicker-0\">",
    "id": "mat-input-0",
    "name": "",
    "placeholder": "Selecione uma data",
    "tag": "input",
    "text": "",
    "type": "coord",
    "x": 245,
    "y": 158,
    "passo": "data_inicio",
    "instrucao": "Clique no campo 'Data Início' (De)"
  },
  {
    "html": "<input _ngcontent-ng-c1737117349=\"\" matinput=\"\" placeholder=\"Selecione uma data\" class=\"mat-mdc-input-element mat-datepicker-input ng-tns-c1205077789-21 ng-untouched ng-pristine ng-valid mat-mdc-form-field-input-control mdc-text-field__input cdk-text-field-autofill-monitored\" id=\"mat-input-1\" aria-invalid=\"false\" aria-required=\"false\" aria-haspopup=\"dialog\" data-mat-calendar=\"mat-datepicker-1\">",
    "id": "mat-input-1",
    "name": "",
    "placeholder": "Selecione uma data",
    "tag": "input",
    "text": "",
    "type": "coord",
    "x": 540,
    "y": 158,
    "passo": "data_fim",
    "instrucao": "Clique no campo 'Data Fim' (Até)"
  },
  {
    "html": "<button _ngcontent-ng-c1737117349=\"\" mat-button=\"\" class=\"btn btn-primary me-2\"><i _ngcontent-ng-c1737117349=\"\" class=\"fa fa-filter\"></i><!----> Filtrar </button>",
    "id": "",
    "name": "",
    "placeholder": "",
    "tag": "button",
    "text": "Filtrar",
    "type": "coord",
    "x": 1841,
    "y": 564,
    "passo": "botao_filtrar",
    "instrucao": "Clique no botão 'Filtrar'"
  },
  {
    "html": "<button _ngcontent-ng-c1567543337=\"\" id=\"btn-exportar-datatable\" class=\"btn btn-primary float-end isFormButton top-button mc-first ng-star-inserted\"><i _ngcontent-ng-c1567543337=\"\" class=\"fa fa-file-excel\" style=\"margin-right: 10px;\"></i> Exportar <!----></button>",
    "id": "btn-exportar-datatable",
    "name": "",
    "placeholder": "",
    "tag": "button",
    "text": "Exportar",
    "type": "coord",
    "x": 1804,
    "y": 655,
    "passo": "botao_exportar",
    "instrucao": "Clique no botão 'Exportar'"
  }
]
"""
_CAPTURA = {item["passo"]: item for item in json.loads(_CAPTURA_JSON)}


# ── Diagnóstico: salva o HTML real da página + print para inspeção ─────────────
def _salvar_diagnostico(driver, pasta_local, log_fn=print, sufixo=""):
    try:
        pasta = Path(pasta_local)
        nome = f"operview_diagnostico{sufixo}.html"
        (pasta / nome).write_text(driver.page_source or "", encoding="utf-8", errors="ignore")
        log_fn(f"  🩺 diagnóstico salvo: {nome}")
        try:
            driver.save_screenshot(str(pasta / f"operview_diagnostico{sufixo}.png"))
        except Exception:
            pass
    except Exception as e:
        log_fn(f"  ⚠️  falha ao salvar diagnóstico: {e}")


# ── Esperas (estilo GeoOnline: WebDriverWait em seletores estáveis) ───────────
def _esperar_elemento(driver, xpath, timeout=40, log_fn=print):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except Exception:
        return False


def _esperar_texto(driver, texto, timeout=20, log_fn=print):
    seguro = texto.replace('"', "'")
    return _esperar_elemento(
        driver, f"//*[contains(normalize-space(.), \"{seguro}\")]", timeout, log_fn)


# ── Localizar (id/texto estável quando possível; coordenada como fallback) ─────
def _locator_xpath(cap):
    """Constrói um XPath estável a partir da identidade real do elemento capturado
    (id > tag+texto+classe), não de coordenadas. Resolution-independent."""
    if not cap:
        return None
    tag = cap.get("tag", "")
    cid = cap.get("id")
    if cid and tag in ("input", "button", "select", "a"):
        return f"//{tag}[@id='{cid}']"
    texto = (cap.get("text", "") or "").strip()
    cls = ""
    m = re.search(r'class="([^"]*)"', cap.get("html", "") or "")
    if m:
        cls = m.group(1)
    tokens = [c for c in cls.split()
              if c and not c.startswith("ng-") and c != "cursor-pointer"]
    seg = texto.replace('"', "'")
    if texto and tag in ("input", "button", "span", "mat-panel-title", "select", "a", "div"):
        if tokens:
            conds = " and ".join(f"contains(@class,'{t}')" for t in tokens)
            return f"//{tag}[{conds} and contains(normalize-space(.), \"{seg}\")]"
        return f"//{tag}[contains(normalize-space(.), \"{seg}\")]"
    if tokens:
        conds = " and ".join(f"contains(@class,'{t}')" for t in tokens)
        return f"//{tag}[{conds}]"
    if texto:
        return f"//*[contains(normalize-space(.), \"{seg}\")]"
    return None


def _localizar(driver, cap):
    xp = _locator_xpath(cap)
    if not xp:
        return None
    for el in driver.find_elements(By.XPATH, xp):
        try:
            if el.is_displayed():
                return el
        except Exception:
            pass
    return None


def _coords_viewport(driver, el):
    """Centro do elemento em coordenadas de viewport (CDP usa viewport)."""
    return driver.execute_script(
        "var r=arguments[0].getBoundingClientRect();"
        "return [r.left + r.width/2, r.top + r.height/2];", el)


def _mouse_clicar(driver, x, y):
    """Clique de cursor REAL via CDP (mouseMoved + mousePressed + mouseReleased).

    Equivale a um clique humano de verdade na coordenada — gera mousedown/
    mouseup/pointerdown/pointerup/click trusted. Não é o element.click() do
    Selenium (que causa scroll/scroll-into-view e falha em SPAs do Angular).
    """
    xi, yi = int(round(x)), int(round(y))
    driver.execute_cdp_cmd("Input.dispatchMouseEvent",
                           {"type": "mouseMoved", "x": xi, "y": yi})
    time.sleep(0.15)
    driver.execute_cdp_cmd("Input.dispatchMouseEvent",
        {"type": "mousePressed", "x": xi, "y": yi, "button": "left", "clickCount": 1})
    driver.execute_cdp_cmd("Input.dispatchMouseEvent",
        {"type": "mouseReleased", "x": xi, "y": yi, "button": "left", "clickCount": 1})
    time.sleep(0.1)


def _clicar(driver, cap, log_fn=print, timeout=25):
    if not cap:
        raise OperviewNaoImplementado("passo de treinamento ausente")
    xp = _locator_xpath(cap)
    el = None
    if xp:
        try:
            el = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, xp)))
        except Exception:
            el = None
    if el is None:
        # fallback: coordenada bruta (usado só quando não há identidade estável)
        if "x" in cap and "y" in cap:
            x, y = cap["x"], cap["y"]
            _mouse_clicar(driver, x, y)
            time.sleep(1.3)
            return
        raise OperviewNaoImplementado(
            f"não foi possível localizar o passo '{cap.get('passo')}' na página")
    x, y = _coords_viewport(driver, el)
    _mouse_clicar(driver, x, y)
    time.sleep(1.3)


# ── Helpers de login ───────────────────────────────────────────────────────────
def _achar_por_texto(driver, texto, tag="*"):
    seguro = texto.replace('"', "'")
    xpath = f"//{tag}[contains(normalize-space(.), \"{seguro}\")]"
    for e in driver.find_elements(By.XPATH, xpath):
        try:
            if e.is_displayed():
                return e
        except Exception:
            pass
    return None


def _tem_form_login(driver) -> bool:
    if driver.find_elements(By.CSS_SELECTOR, "input[type=password]"):
        return True
    url = (driver.current_url or "").lower()
    if any(k in url for k in ("login", "oauth", "adfs", "sts", "signin", "auth", "identity")):
        return True
    if _achar_por_texto(driver, "Entrar") or _achar_por_texto(driver, "Acessar"):
        return True
    return False


def _precisa_login(driver) -> bool:
    return _tem_form_login(driver)


def _aguardar_spa(driver, log_fn=print, tentativas=20):
    for _ in range(tentativas):
        try:
            if driver.find_element(By.TAG_NAME, "body").text.strip():
                return
        except Exception:
            pass
        time.sleep(1)
    log_fn("  ⚠️  SPA demorou a renderizar.")


def _aguardar_login(driver, stop_event=None, login_event=None, log_fn=print):
    deadline = time.time() + TIMEOUT_LOGIN
    while time.time() < deadline:
        if stop_event and stop_event.is_set():
            return
        if login_event and login_event.is_set():
            log_fn("  ✅ Login confirmado pelo usuário — seguindo o fluxo.")
            return
        if _achar_por_texto(driver, "Consulta Incidência") is not None:
            log_fn("  ✅ Sessão do Operview detectada automaticamente.")
            return
        time.sleep(5)
    log_fn("  ⚠️  Tempo de espera de login esgotado (continuando mesmo assim).")


# ── Navegação: lateral → Auditoria → Consulta Incidência ──────────────────────
def _esperar_capa(driver, cap, timeout=25, log_fn=print):
    xp = _locator_xpath(cap)
    if not xp:
        return True
    return _esperar_elemento(driver, xp, timeout, log_fn)


def _navegar_consulta(driver, caps, pasta_local, log_fn=print, stop_event=None):
    _clicar(driver, caps.get("abrir_lateral"), log_fn)
    if not _esperar_capa(driver, caps.get("menu_auditoria"), 25, log_fn):
        log_fn("  ⚠️  menu 'Auditoria' não apareceu após abrir a lateral.")
        _salvar_diagnostico(driver, pasta_local, log_fn, "_falha_lateral")
    _clicar(driver, caps.get("menu_auditoria"), log_fn)
    if not _esperar_capa(driver, caps.get("consulta_incidencia"), 25, log_fn):
        log_fn("  ⚠️  'Consulta Incidência' não apareceu após clicar em Auditoria.")
        _salvar_diagnostico(driver, pasta_local, log_fn, "_falha_auditoria")
    _clicar(driver, caps.get("consulta_incidencia"), log_fn)
    if _esperar_elemento(driver, "//input[@placeholder='Selecione uma data']",
                         timeout=45, log_fn=log_fn):
        log_fn("  ✅ formulário de consulta carregado (campos de data visíveis).")
    else:
        log_fn("  ⚠️  formulário de consulta não carregou (campos de data ausentes).")
        _salvar_diagnostico(driver, pasta_local, log_fn, "_falha_consulta")


# ── Preencher filtros (datas + polo) ──────────────────────────────────────────
_MESES = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
}


def _parse_mes_ano(label):
    label = (label or "").strip()
    # formato MM/AAAA (ex.: "08/2026") — é o que o Operview usa
    if "/" in label:
        partes = label.split("/")
        try:
            return int(partes[1]), int(partes[0])
        except (ValueError, IndexError):
            pass
    # fallback: "NOME_DO_MES AAAA"
    ano = mes = None
    for p in label.upper().split():
        if p in _MESES:
            mes = _MESES[p]
        elif p.isdigit():
            ano = int(p)
    return ano, mes


def _navegar_calendario(driver, data, log_fn=print):
    """Navega o datepicker do Material até o mês/ano alvo usando os botões prev/next."""
    alvo = (data.year, data.month)
    for _ in range(36):
        try:
            label = driver.find_element(
                By.CSS_SELECTOR, ".mat-calendar-period-button").text.strip()
        except Exception:
            return False
        parsed = _parse_mes_ano(label)
        if parsed == alvo:
            return True
        if parsed[0] is None or parsed[1] is None:
            time.sleep(0.3)  # rótulo ainda não pintou; tenta de novo
            continue
        ano, mes = parsed
        seletor = (".mat-calendar-next-button"
                   if alvo > (ano, mes) else ".mat-calendar-previous-button")
        try:
            btn = driver.find_element(By.CSS_SELECTOR, seletor)
            _mouse_clicar(driver, *_coords_viewport(driver, btn))
        except Exception:
            return False
        time.sleep(0.4)
    return False


def _tem_calendario(driver):
    return bool(driver.find_elements(
        By.CSS_SELECTOR, ".mat-calendar, .mat-datepicker-content"))


def _abrir_calendario(driver, el, log_fn=print):
    """Abre o datepicker do Material. O toggle fica em <mat-datepicker-toggle> (não no
    <button>), então clicamos no botão interno via data-mat-calendar."""
    if _tem_calendario(driver):
        return True
    cal_id = el.get_attribute("data-mat-calendar") or ""
    # 1) botão do toggle (mais confiável) — usa o data-mat-calendar do input
    tentativas = []
    if cal_id:
        tentativas.append(
            f"//mat-datepicker-toggle[@data-mat-calendar='{cal_id}']//button")
    tentativas += [
        "ancestor::div[contains(@class,'mat-mdc-form-field-infix')]"
        "/following-sibling::div[contains(@class,'mat-mdc-form-field-icon-suffix')]"
        "//button[contains(@class,'mat-mdc-icon-button')]",
        "following-sibling::mat-datepicker-toggle//button",
        "ancestor::mat-form-field[1]//mat-datepicker-toggle//button",
        "//mat-datepicker-toggle//button",
    ]
    for xp in tentativas:
        try:
            tg = el.find_element(By.XPATH, xp)
            _mouse_clicar(driver, *_coords_viewport(driver, tg))
            time.sleep(1.5)
            if _tem_calendario(driver):
                return True
        except Exception:
            continue
    # 2) fallback: clique no próprio input
    try:
        _mouse_clicar(driver, *_coords_viewport(driver, el))
        time.sleep(1.5)
        if _tem_calendario(driver):
            return True
        el.click()
        time.sleep(1.5)
    except Exception:
        pass
    return _tem_calendario(driver)


def _setar_data(driver, qual, data, log_fn=print, pasta_local="."):
    """No Operview a data NÃO é digitada: abre o calendário e clica o dia."""
    try:
        inputs = WebDriverWait(driver, 40).until(
            lambda d: d.find_elements(By.XPATH, "//input[@placeholder='Selecione uma data']")
            if len(d.find_elements(By.XPATH, "//input[@placeholder='Selecione uma data']")) >= 2
            else None
        )
    except Exception:
        raise OperviewNaoImplementado("campos de data 'Selecione uma data' não encontrados")
    el = inputs[0] if qual == "inicial" else inputs[1]
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    if not _abrir_calendario(driver, el, log_fn):
        log_fn(f"  ⚠️  não abriu o calendário ({qual})")
        _salvar_diagnostico(driver, pasta_local, log_fn, f"_calendario_{qual}")
        return
    if not _navegar_calendario(driver, data, log_fn):
        log_fn(f"  ⚠️  não conseguiu navegar o calendário até {data:%m/%Y}")
        _salvar_diagnostico(driver, pasta_local, log_fn, f"_calendario_{qual}")
    dia = str(data.day)
    xpath = ("//td[contains(@class,'mat-calendar-body-cell') and "
             "not(contains(@class,'mat-calendar-body-disabled'))]"
             f"//button[normalize-space()='{dia}']")
    try:
        botao = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, xpath)))
        _mouse_clicar(driver, *_coords_viewport(driver, botao))
        log_fn(f"  📅 data selecionada ({qual}): {data:%d/%m/%Y}")
    except Exception:
        log_fn(f"  ⚠️  não encontrou o dia {dia} no calendário de {data:%m/%Y}")
        _salvar_diagnostico(driver, pasta_local, log_fn, f"_calendario_{qual}")
    time.sleep(0.8)


# ── Mapeamento de polos (igual ao GeoOnline: UTS/UTN são regiões, não opções) ──
_UTN_POLOS = ["CAMPOS", "LAGOS", "MACAÉ", "NOROESTE"]
_UTS_POLOS = ["MAGÉ", "NITERÓI", "SÃO GONÇALO", "SERRANA", "SUL"]


def _resolver_polos(polo):
    """Mapeia UTS/UTN para as cidades. Retorna None para 'Todos os polos'."""
    if not polo or str(polo).strip().lower() in ("todos os polos", "todos", "all", ""):
        return None
    p = str(polo).strip().upper()
    if p == "UTN":
        return list(_UTN_POLOS)
    if p == "UTS":
        return list(_UTS_POLOS)
    return [p]


def _selecionar_polo(driver, polo, cap=None, log_fn=print, pasta_local="."):
    polos = _resolver_polos(polo)
    if polos is None:
        log_fn("  ℹ️  Polo: Todos os polos (sem filtro).")
        return
    # abre o dropdown de Polo (multi-select, igual ao GeoOnline)
    trigger = None
    for xp in (
        "//mat-form-field[.//*[normalize-space()='Polo']]//mat-select",
        "//mat-form-field[.//*[normalize-space()='Polo']]//div[contains(@class,'mat-select')]",
        "//span[normalize-space()='Polo']/ancestor::*[contains(@class,'mat-form-field') or contains(@class,'mat-select')][1]",
    ):
        try:
            trigger = driver.find_element(By.XPATH, xp)
            break
        except Exception:
            continue
    if trigger is None:
        trigger = _achar_por_texto(driver, "Polo")
    if trigger is None:
        log_fn("  ⚠️  campo de polo não encontrado — não selecionado.")
        _salvar_diagnostico(driver, pasta_local, log_fn, "_polo")
        return
    _mouse_clicar(driver, *_coords_viewport(driver, trigger))
    time.sleep(1.2)
    try:
        WebDriverWait(driver, 8).until(
            lambda d: d.find_elements(
                By.CSS_SELECTOR, "mat-option, .mat-select-panel, [role='option']"))
    except Exception:
        pass
    selecionados = []
    for p in polos:
        opt = None
        for xp in (
            f"//mat-option[normalize-space()='{p}']",
            f"//*[normalize-space()='{p}']",
            f"//li[normalize-space()='{p}']",
        ):
            try:
                opt = driver.find_element(By.XPATH, xp)
                break
            except Exception:
                continue
        if opt is None:
            opt = _achar_por_texto(driver, p)
        if opt:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", opt)
            time.sleep(0.3)
            _mouse_clicar(driver, *_coords_viewport(driver, opt))
            selecionados.append(p)
            time.sleep(0.4)
        else:
            log_fn(f"  ⚠️  polo '{p}' não encontrado no painel.")
    if selecionados:
        log_fn(f"  🏭 polos selecionados ({len(selecionados)}/{len(polos)}): "
               + ", ".join(selecionados))
    else:
        log_fn("  ⚠️  nenhum polo selecionado.")
        _salvar_diagnostico(driver, pasta_local, log_fn, "_polo")
    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except Exception:
        pass
    time.sleep(0.5)


def _preencher_filtros(driver, data_ini, data_fim, polo, caps, log_fn=print,
                       stop_event=None, pasta_local="."):
    _setar_data(driver, "inicial", data_ini, log_fn, pasta_local)
    _setar_data(driver, "final", data_fim, log_fn, pasta_local)
    _selecionar_polo(driver, polo, caps.get("filtro_polo"), log_fn, pasta_local)
    time.sleep(1)


# ── Resultados ─────────────────────────────────────────────────────────────────
def _tem_resultados(driver) -> bool:
    try:
        if len(driver.find_elements(By.XPATH, "//tbody/tr")) > 0:
            return True
        return len(driver.find_elements(By.XPATH, "//table//tr")) > 1
    except Exception:
        return False


def _sem_registros(driver) -> bool:
    t = (driver.page_source or "").lower().replace(" ", "")
    return any(s in t for s in ("nenhumregistro", "semregistros", "nãohádados",
                                "nãopossuiregistros", "nomatchingrecords"))


# ── Exportar + espera de download ──────────────────────────────────────────────
def _aguardar_download(pasta_download, log_fn=print, stop_event=None,
                       timeout=TIMEOUT_DOWNLOAD) -> Path:
    pasta = Path(pasta_download)
    conhecidos = {p.name for p in pasta.glob("*.xlsx")} | {p.name for p in pasta.glob("*.csv")}
    deadline = time.time() + timeout
    while time.time() < deadline:
        if stop_event and stop_event.is_set():
            return None
        candidatos = []
        for ext in ("*.xlsx", "*.csv"):
            for p in pasta.glob(ext):
                if p.name.endswith(".crdownload"):
                    continue
                candidatos.append(p)
        novos = [p for p in candidatos if p.name not in conhecidos]
        if novos:
            p = max(novos, key=lambda x: x.stat().st_mtime)
            try:
                t0 = p.stat().st_size
            except OSError:
                t0 = -1
            time.sleep(2)
            try:
                t1 = p.stat().st_size
            except OSError:
                continue
            if t0 == t1:
                log_fn(f"  📥 download concluído: {p.name}")
                return p
        time.sleep(1)
    log_fn("  ⚠️  Timeout aguardando o download do Operview.")
    return None


def _baixar_intervalo(driver, pasta_download, data_ini, data_fim, caps, log_fn=print,
                      stop_event=None, polo="Todos os polos") -> Path:
    _preencher_filtros(driver, data_ini, data_fim, polo, caps, log_fn, stop_event, pasta_download)
    _clicar(driver, caps.get("botao_filtrar"), log_fn)
    log_fn(f"  ⏳ Filtrando [{data_ini:%d/%m/%Y} → {data_fim:%d/%m/%Y}]...")
    time.sleep(4)
    if "falha interna" in (driver.page_source or "").lower():
        raise _FalhaServidor("falha interna do servidor")
    try:
        WebDriverWait(driver, 150).until(lambda d: _tem_resultados(d) or _sem_registros(d))
        if _sem_registros(driver):
            log_fn("  ℹ️  consulta sem registros para o período.")
    except Exception:
        log_fn("  ⚠️  timeout esperando resultados (exportando assim mesmo).")
    xp_exp = _locator_xpath(caps.get("botao_exportar"))
    if xp_exp:
        try:
            WebDriverWait(driver, 120).until(
                EC.element_to_be_clickable((By.XPATH, xp_exp)))
        except Exception:
            log_fn("  ⚠️  botão 'Exportar' não ficou clicável (tentando mesmo assim).")
    _clicar(driver, caps.get("botao_exportar"), log_fn)
    arquivo = _aguardar_download(pasta_download, log_fn, stop_event)
    if arquivo is None:
        raise OperviewNaoImplementado("download não foi detectado após 'Exportar'")
    chunk = Path(pasta_download) / f"_operview_{data_ini:%Y%m%d}_{data_fim:%Y%m%d}.xlsx"
    shutil.move(str(arquivo), str(chunk))
    return chunk


def _download_chunked(driver, pasta_download, data_ini, data_fim, caps, log_fn=print,
                      stop_event=None, polo="Todos os polos") -> list[Path]:
    if stop_event and stop_event.is_set():
        return []
    try:
        return [_baixar_intervalo(driver, pasta_download, data_ini, data_fim,
                                  caps, log_fn, stop_event, polo)]
    except _FalhaServidor:
        total_days = (data_fim - data_ini).days
        if total_days == 0:
            log_fn("  ⚠️  Dia único não baixou — pulando.")
            return []
        mid = data_ini + timedelta(days=total_days // 2)
        log_fn(f"  🔁 dividindo [{data_ini:%d/%m/%Y}→{data_fim:%d/%m/%Y}] por falha de servidor")
        a = _download_chunked(driver, pasta_download, data_ini, mid, caps, log_fn, stop_event, polo)
        b = _download_chunked(driver, pasta_download, mid + timedelta(days=1), data_fim, caps,
                             log_fn, stop_event, polo)
        return a + b


def _merge_chunks(chunks: list[Path], destino_final: Path, log_fn=print) -> Path:
    dfs = []
    for c in chunks:
        df = pd.read_excel(str(c), header=2, dtype=str)
        dfs.append(df)
        c.unlink()
    df_final = pd.concat(dfs, ignore_index=True).drop_duplicates()
    tmp = destino_final.with_suffix(".tmp_operview.xlsx")
    df_final.to_excel(str(tmp), index=False, sheet_name=SHEET_NAME, startrow=2)
    if destino_final.exists():
        for _ in range(10):
            try:
                destino_final.unlink()
                break
            except (PermissionError, OSError):
                time.sleep(0.5)
    try:
        os.replace(str(tmp), str(destino_final))
    except OSError:
        shutil.copy2(str(tmp), str(destino_final))
        tmp.unlink()
    log_fn(f"  📦 Merge concluído: {len(df_final)} linhas → {destino_final.name}")
    return destino_final


# ── Função pública (mesma assinatura do GeoOnline) ─────────────────────────────
def baixar_incidencias(cfg: OperviewConfig, pasta_local, data_ini: datetime,
                       data_fim: datetime, log_fn=print, stop_event=None,
                       polo: str = "Todos os polos", notificar=None) -> Path:
    pasta = Path(pasta_local)
    pasta.mkdir(parents=True, exist_ok=True)
    url = cfg.url or URL_PADRAO
    caps = _CAPTURA
    login_event = threading.Event()

    driver = _criar_driver(str(pasta))
    try:
        log_fn(f"🌐 Abrindo Operview: {url}")
        driver.get(url)
        _aguardar_spa(driver, log_fn)

        if _precisa_login(driver):
            log_fn("🔔 Login necessário no Operview.")
            if notificar:
                notificar(login_event)
            log_fn("⏳ Aguardando login (faça o login e clique OK, ou até 15 min)...")
            _aguardar_login(driver, stop_event, login_event, log_fn)
        else:
            log_fn("✅ Sessão do Operview ativa (sem login manual necessário).")

        _navegar_consulta(driver, caps, pasta_local, log_fn, stop_event)
        chunks = _download_chunked(driver, str(pasta), data_ini, data_fim, caps,
                                   log_fn, stop_event, polo)

        destino = pasta / NOME_ARQUIVO
        if not chunks:
            raise OperviewNaoImplementado("nenhum arquivo baixado do Operview")
        if len(chunks) == 1:
            shutil.move(str(chunks[0]), str(destino))
        else:
            _merge_chunks(chunks, destino, log_fn)
        log_fn(f"  💾 Salvo: {destino}")
        return destino
    finally:
        driver.quit()
