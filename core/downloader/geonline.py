import base64
import ctypes
import os
import tempfile
import threading
import time
import shutil
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from config import GeoOnlineConfig

# ── Mapeamento de polos ───────────────────────────────────────────────────────

_UTN_POLOS = ["CAMPOS", "LAGOS", "MACAÉ", "NOROESTE"]
_UTS_POLOS = ["MAGÉ", "NITERÓI", "SÃO GONÇALO", "SERRANA", "SUL"]


def _resolver_polos(polo: str) -> list[str] | None:
    """Retorna lista de nomes de polo para selecionar no GeoOnline, ou None para TODOS."""
    if not polo or polo == "Todos os polos":
        return None
    if polo == "UTN":
        return list(_UTN_POLOS)
    if polo == "UTS":
        return list(_UTS_POLOS)
    return [polo.upper()]

NOME_ARQUIVO = "extracao_incidencias_geonline.xlsx"
SHEET_NAME = "Relatorio_de_Incidencias"
TIMEOUT_DOWNLOAD = 180  # segundos esperando o arquivo aparecer


def _watchdog_driver(driver: webdriver.Chrome, stop_event: threading.Event) -> None:
    """Mata o driver assim que stop_event for sinalizado."""
    stop_event.wait()
    try:
        driver.quit()
    except Exception:
        pass


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

    # Envia a janela do Chrome para trás de todas as outras janelas
    HWND_BOTTOM = 1
    SWP_NOMOVE = 0x0002
    SWP_NOSIZE = 0x0001
    try:
        title = driver.title
        if title:
            hwnd = ctypes.windll.user32.FindWindowW(None, title)
            if hwnd:
                ctypes.windll.user32.SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0,
                                                  SWP_NOMOVE | SWP_NOSIZE)
    except Exception:
        pass

    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return driver


def _fazer_login(driver: webdriver.Chrome, cfg: GeoOnlineConfig) -> None:
    wait = WebDriverWait(driver, 30)
    driver.get(cfg.url)
    wait.until(EC.element_to_be_clickable((By.ID, "login-conta"))).send_keys(cfg.conta)
    wait.until(EC.element_to_be_clickable((By.ID, "login-email"))).send_keys(cfg.login)
    senha = wait.until(EC.element_to_be_clickable((By.ID, "login-senha")))
    senha.send_keys(cfg.senha)
    senha.send_keys(Keys.ENTER)
    time.sleep(5)


def _selecionar_polo_filtro(driver: webdriver.Chrome, polos: list[str]) -> None:
    """
    Abre o multi-select 'Polo' no formulário de filtro do GeoOnline e marca
    os polos desejados. Os filtros começam em branco, não é preciso resetar nada.
    """
    # ── Abre o dropdown ─────────────────────────────────────────────────────
    aberto = False
    xpaths_trigger = [
        "//div[contains(@class,'dropdown-btn')][.//*[normalize-space()='Polo'] or normalize-space()='Polo']",
        "//mat-form-field[.//*[normalize-space()='Polo']]//mat-select",
        "//mat-form-field[.//*[normalize-space()='Polo']]//div[contains(@class,'trigger')]",
        "//span[normalize-space()='Polo']/parent::*",
        "//span[normalize-space()='Polo']/ancestor::div[contains(@class,'select') or contains(@class,'dropdown') or contains(@class,'multiselect')][1]",
        "//input[@placeholder='Polo']",
    ]
    for xpath in xpaths_trigger:
        try:
            el = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", el)
            el.click()
            time.sleep(0.8)
            itens = driver.find_elements(By.XPATH, "//*[normalize-space()='CAMPOS' or normalize-space()='MAGÉ' or normalize-space()='NITERÓI']")
            if itens:
                aberto = True
                break
        except Exception:
            continue

    if not aberto:
        try:
            driver.execute_script("""
                var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
                var node;
                while ((node = walker.nextNode())) {
                    if (node.textContent.trim() === 'Polo') {
                        var el = node.parentElement;
                        el.click();
                        break;
                    }
                }
            """)
            time.sleep(0.8)
        except Exception:
            pass

    # ── Marca cada polo ────────────────────────────────────────────────────
    for polo in polos:
        for xpath_item, t in [
            (f"//*[normalize-space()='{polo}']", 3),
            (f"//li[.//*[normalize-space()='{polo}']]", 1),
            (f"//li[normalize-space()='{polo}']", 1),
            (f"//div[normalize-space()='{polo}']", 1),
        ]:
            try:
                item = WebDriverWait(driver, t).until(
                    EC.presence_of_element_located((By.XPATH, xpath_item))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", item)
                time.sleep(0.1)
                try:
                    cb = item.find_element(By.XPATH, ".//input[@type='checkbox']")
                    driver.execute_script("arguments[0].click();", cb)
                except Exception:
                    driver.execute_script("arguments[0].click();", item)
                time.sleep(0.2)
                break
            except Exception:
                continue

    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except Exception:
        pass
    time.sleep(0.4)


def _navegar_consulta(driver: webdriver.Chrome) -> None:
    wait = WebDriverWait(driver, 30)
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//span[normalize-space()='Relatórios']")
    )).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//span[normalize-space()='Consulta Incidência']")
    )).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//span[normalize-space()='Filtro']")
    )).click()
    # Garante que o formulário carregou
    wait.until(EC.presence_of_element_located(
        (By.XPATH, "//input[@placeholder='Data Inicial'] | //input[contains(@placeholder, 'Data inicial')]")
    ))
    time.sleep(0.5)


def _preencher_datas(driver: webdriver.Chrome, data_ini: str, data_fim: str) -> None:
    """data_ini e data_fim no formato DD/MM/YYYY."""
    wait = WebDriverWait(driver, 30)

    campo_ini = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//input[@placeholder='Data Inicial']")
    ))
    campo_ini.click()
    campo_ini.send_keys(Keys.CONTROL + "a")
    campo_ini.send_keys(Keys.DELETE)
    campo_ini.send_keys(data_ini)
    campo_ini.send_keys(Keys.ENTER)
    time.sleep(0.5)

    campo_fim = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//input[@placeholder='Data Final']")
    ))
    campo_fim.click()
    campo_fim.send_keys(Keys.CONTROL + "a")
    campo_fim.send_keys(Keys.DELETE)
    campo_fim.send_keys(data_fim)
    campo_fim.send_keys(Keys.ENTER)
    time.sleep(0.5)


def _clicar_extrair(driver: webdriver.Chrome) -> None:
    wait = WebDriverWait(driver, 30)
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[normalize-space()='Extrair Dados'] | //span[normalize-space()='Extrair Dados']")
    )).click()
    time.sleep(2)


def _preparar_formulario(driver: webdriver.Chrome, data_ini: str, data_fim: str,
                          pasta_download: str, log_fn=print) -> bool:
    """
    Garante que estamos na página de Filtro com o formulário pronto.
    Se a página está em estado inconsistente (pós-30k, pós-download),
    recarrega a navegação do zero.
    """
    # Tenta preencher diretamente primeiro
    try:
        _preencher_datas(driver, data_ini, data_fim)
        return True
    except Exception:
        log_fn("  ⚠️  Formulário não respondeu — recarregando navegação...")

    # Recarrega a navegação do zero
    try:
        _navegar_consulta(driver)
        _preencher_datas(driver, data_ini, data_fim)
        return True
    except Exception as e:
        log_fn(f"  ❌ Falha ao preparar formulário: {e}")
        return False


def _tentar_download(
    driver: webdriver.Chrome,
    pasta_download: str,
    data_ini: datetime,
    data_fim: datetime,
    log_fn=print,
    stop_event: threading.Event | None = None,
    polos: list[str] | None = None,
) -> Path | None:
    """
    Tenta baixar [data_ini → data_fim].
    Retorna o Path do arquivo se sucesso, None se 30k ou cancelamento.

    O formulário JÁ ESTÁ ABERTO — só preenche datas e clica Extrair.
    Nunca re-navega (baixar_incidencias já navegou antes de chamar isto).
    """
    if stop_event and stop_event.is_set():
        return None

    ini_str = data_ini.strftime("%d/%m/%Y")
    fim_str = data_fim.strftime("%d/%m/%Y")
    log_fn(f"  Tentando [{ini_str} → {fim_str}]...")

    pasta_p = Path(pasta_download)
    destino = pasta_p / f"_dl_{data_ini.strftime('%Y%m%d')}_{data_fim.strftime('%Y%m%d')}_{int(time.time())}.xlsx"

    xlsx_antes = {p.resolve() for p in pasta_p.glob("*.xlsx")}
    crdownloads_antes = {p.resolve() for p in pasta_p.glob("*.crdownload")}

    # Injeta interceptor para capturar blob Excel como base64
    driver.execute_script("""
        if (!window.__blobHookInstalled) {
            window.__blobHookInstalled = true;
            window.__blobBase64 = null;
            const _orig = URL.createObjectURL;
            URL.createObjectURL = function(blob) {
                const url = _orig.apply(this, arguments);
                if (blob && (
                    blob.type.includes('excel') ||
                    blob.type.includes('spreadsheet') ||
                    blob.type.includes('ms-excel') ||
                    blob.type.includes('octet-stream') ||
                    blob.type.includes('application/')
                )) {
                    window.__blobBase64 = null;
                    const reader = new FileReader();
                    reader.onloadend = function() { window.__blobBase64 = reader.result; };
                    reader.readAsDataURL(blob);
                }
                return url;
            };
        }
        window.__blobBase64 = null;
    """)

    # Tenta preencher formulário. Se falhar, recarrega navegação.
    try:
        if not _preparar_formulario(driver, ini_str, fim_str, pasta_download, log_fn):
            return None
        _clicar_extrair(driver)
    except Exception as e:
        log_fn(f"  ❌ Erro ao submeter formulário [{ini_str}→{fim_str}]: {e}")
        return None

    inicio = time.time()
    download_iniciou = False

    while True:
        if stop_event and stop_event.is_set():
            return None

        # Popup 30k — verificado ANTES de tudo
        try:
            popups = driver.find_elements(
                By.XPATH, "//*[contains(text(),'Aviso Exportação Incidências')]"
            )
            if popups and popups[0].is_displayed():
                try:
                    driver.find_element(
                        By.XPATH, "//button[normalize-space()='Fechar']"
                    ).click()
                    time.sleep(1)
                except Exception:
                    pass
                log_fn(f"  ⚠️  30k atingido para [{ini_str} → {fim_str}]")
                return None
        except Exception:
            pass

        # Espera popup de carregamento ("Aguarde...") sumir antes de verificar download
        loading = driver.find_elements(By.XPATH, "//*[contains(text(),'Aguarde') or contains(text(),'Carregando') or contains(text(),'gerando') or contains(text(),'Gerando')]")
        if loading and loading[0].is_displayed():
            for _ in range(120):
                loading = driver.find_elements(By.XPATH, "//*[contains(text(),'Aguarde') or contains(text(),'Carregando') or contains(text(),'gerando') or contains(text(),'Gerando')]")
                if not loading or not loading[0].is_displayed():
                    time.sleep(1)
                    break
                time.sleep(1)
            continue

        # Verifica se o JS capturou o blob como base64
        b64 = driver.execute_script("return window.__blobBase64;")
        if b64:
            if "," in b64:
                b64_data = b64.split(",", 1)[1]
            else:
                b64_data = b64
            raw = base64.b64decode(b64_data)

            for cd in list(pasta_p.glob("*.crdownload")):
                try:
                    cd.unlink()
                except (PermissionError, OSError):
                    pass

            tmp_path = destino.with_suffix(".tmp.xlsx")
            try:
                if tmp_path.exists():
                    tmp_path.unlink()
                tmp_path.write_bytes(raw)
                os.replace(str(tmp_path), str(destino))
                log_fn(f"  ✅ Capturado via JS base64 ({len(raw)} bytes) [{ini_str} → {fim_str}]")
                time.sleep(3)
                return destino
            except Exception as e:
                log_fn(f"  ❌ Falha ao salvar arquivo final: {e}")
                return None

        crdownloads_novos = {p.resolve() for p in pasta_p.glob("*.crdownload")} - crdownloads_antes
        xlsx_novos = {p.resolve() for p in pasta_p.glob("*.xlsx")} - xlsx_antes

        if crdownloads_novos:
            download_iniciou = True

        if xlsx_novos and not crdownloads_novos:
            novo = Path(list(xlsx_novos)[0])
            time.sleep(1)
            if novo.resolve() != destino.resolve():
                try:
                    os.replace(str(novo), str(destino))
                except Exception:
                    destino = novo
            log_fn(f"  ✅ Baixado [{ini_str} → {fim_str}]")
            return destino

        elapsed = time.time() - inicio

        if elapsed > TIMEOUT_DOWNLOAD and not download_iniciou:
            log_fn(f"  ⚠️  Sem resposta para [{ini_str} → {fim_str}] após {TIMEOUT_DOWNLOAD}s")
            return None

        if elapsed > 1800:
            log_fn(f"  ⚠️  Timeout 30min para [{ini_str} → {fim_str}]")
            return None

        time.sleep(2)


def _download_chunked(
    driver: webdriver.Chrome,
    pasta_download: str,
    data_ini: datetime,
    data_fim: datetime,
    log_fn=print,
    stop_event: threading.Event | None = None,
    polos: list[str] | None = None,
) -> list[Path]:
    """
    Baixa o período. Se bater no limite de 30k (ou sem resposta), divide ao
    meio e baixa cada metade recursivamente, até chegar em períodos que cabem.

    O formulário JÁ ESTÁ ABERTO quando esta função é chamada.
    Nunca re-navega — só muda datas e clica Extrair.
    """
    if stop_event and stop_event.is_set():
        return []

    ini_str = data_ini.strftime("%d/%m/%Y")
    fim_str = data_fim.strftime("%d/%m/%Y")

    resultado = _tentar_download(driver, pasta_download, data_ini, data_fim, log_fn, stop_event, polos)
    if resultado is not None:
        chunk_nome = Path(pasta_download) / f"_chunk_{data_ini.strftime('%Y%m%d')}_{data_fim.strftime('%Y%m%d')}.xlsx"
        os.replace(str(resultado), str(chunk_nome))
        log_fn(f"  💾 Chunk salvo: {chunk_nome.name}")
        return [chunk_nome]

    if stop_event and stop_event.is_set():
        return []

    total_days = (data_fim - data_ini).days
    if total_days == 0:
        log_fn(f"  ⚠️  Dia único [{ini_str}] não baixou — pulando.")
        return []

    mid = data_ini + timedelta(days=total_days // 2)
    log_fn(f"  🔁 [{ini_str}→{fim_str}] dividido em [{ini_str}→{mid.strftime('%d/%m/%Y')}] e [{(mid + timedelta(days=1)).strftime('%d/%m/%Y')}→{fim_str}]")

    chunks_a = _download_chunked(driver, pasta_download, data_ini, mid, log_fn, stop_event, polos)
    if stop_event and stop_event.is_set():
        return chunks_a
    chunks_b = _download_chunked(driver, pasta_download, mid + timedelta(days=1), data_fim, log_fn, stop_event, polos)
    return chunks_a + chunks_b


def _merge_chunks(chunks: list[Path], destino_final: Path, log_fn=print) -> Path:
    """Une todos os chunks em um único arquivo e remove os temporários."""
    dfs = []
    for c in chunks:
        df = pd.read_excel(str(c), header=2, dtype=str)
        dfs.append(df)
        c.unlink()

    df_final = pd.concat(dfs, ignore_index=True).drop_duplicates()
    tmp_path = destino_final.with_suffix(".tmp.xlsx")
    df_final.to_excel(str(tmp_path), index=False,
                      sheet_name=SHEET_NAME, startrow=2)
    if destino_final.exists():
        for _ in range(10):
            try:
                destino_final.unlink()
                break
            except (PermissionError, OSError):
                time.sleep(0.5)
    try:
        os.replace(str(tmp_path), str(destino_final))
    except OSError:
        shutil.copy2(str(tmp_path), str(destino_final))
        tmp_path.unlink()
    log_fn(f"  📦 Merge concluído: {len(df_final)} linhas → {destino_final.name}")
    return destino_final


def baixar_incidencias(
    cfg: GeoOnlineConfig,
    pasta_local: str,
    data_ini: datetime,
    data_fim: datetime,
    log_fn=print,
    stop_event: threading.Event | None = None,
    polo: str = "Todos os polos",
) -> Path:
    """
    Ponto de entrada público. Faz login, aplica filtro de polo (se informado),
    baixa o período (com chunking automático se > 30k linhas) e retorna o Path
    do arquivo final unificado.
    """
    pasta = str(Path(pasta_local).resolve())
    log_fn(f"🌐 GeoOnline: baixando incidências {data_ini.strftime('%d/%m/%Y')} → {data_fim.strftime('%d/%m/%Y')}")

    driver = _criar_driver(pasta)
    if stop_event:
        threading.Thread(target=_watchdog_driver, args=(driver, stop_event), daemon=True).start()

    try:
        _fazer_login(driver, cfg)
        _navegar_consulta(driver)

        polos = _resolver_polos(polo)
        if polos:
            _selecionar_polo_filtro(driver, polos)

        chunks = _download_chunked(driver, pasta, data_ini, data_fim, log_fn, stop_event, polos)
        if not chunks:
            raise RuntimeError("Nenhum dado baixado do GeoOnline.")

        destino_final = Path(pasta) / NOME_ARQUIVO
        if len(chunks) == 1:
            # Libera lock do arquivo de execução anterior antes de renomear
            if destino_final.exists():
                for _ in range(10):
                    try:
                        destino_final.unlink()
                        break
                    except (PermissionError, OSError):
                        time.sleep(0.5)
            try:
                os.replace(str(chunks[0]), str(destino_final))
            except OSError:
                # Fallback: copy2 tolera mais locks que rename no Windows
                shutil.copy2(str(chunks[0]), str(destino_final))
                chunks[0].unlink()
            log_fn(f"  ✅ Arquivo final: {destino_final.name}")
            return destino_final

        return _merge_chunks(chunks, destino_final, log_fn)

    finally:
        try:
            driver.quit()
        except Exception:
            pass
