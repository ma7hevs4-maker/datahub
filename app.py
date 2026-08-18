"""
DataHub v2 — Entry point.
Conecta GUI → Scheduler → Downloaders → Processors → Integrations.
"""
import os
import shutil
import threading
import time
import unicodedata
from datetime import datetime
from pathlib import Path

from config import load_config, AppConfig
from core.scheduler import Scheduler
from core.downloader.geonline import baixar_incidencias
from core.downloader.operview import baixar_incidencias as baixar_operview
from core.processors.incidencias import tratar_incidencias, gerar_analises_n8n
from core.processors.dashboard_html import gerar_dashboard_html
from core.base_mensal import atualizar_base_mensal
from integrations.sharepoint import sincronizar
from integrations.n8n import enviar
from ui.main_window import MainWindow
from PyQt6.QtWidgets import QApplication

# ── Polo filter helpers ───────────────────────────────────────────────────────

_UTN = {"CAMPOS", "LAGOS", "MACAE", "NOROESTE"}
_UTS = {"MAGE", "NITEROI", "SAO GONCALO", "SERRANA", "SUL"}


def _norm(s: str) -> str:
    return (
        unicodedata.normalize("NFKD", str(s).upper())
        .encode("ASCII", "ignore")
        .decode("ASCII")
        .strip()
    )


def _filtrar_polo(df, polo: str, log_fn=print):
    """Filtra o DataFrame pelo polo selecionado. Retorna df inteiro se 'Todos os polos'."""
    if not polo or polo == "Todos os polos":
        return df

    # Procura coluna "Polo" por match parcial no nome (cobre "Polo", "POLO", "Polo Op.", etc.)
    col = next((c for c in df.columns if "polo" in c.strip().lower()), None)
    if col is None:
        log_fn(f"  ⚠️  Coluna 'Polo' não encontrada no arquivo.")
        log_fn(f"  ℹ️  Colunas disponíveis: {', '.join(df.columns.tolist()[:20])}")
        return df

    if polo == "UTN":
        targets = _UTN
    elif polo == "UTS":
        targets = _UTS
    else:
        targets = {_norm(polo)}

    valores_norm = df[col].astype(str).apply(_norm)
    unicos = sorted(valores_norm.unique())
    log_fn(f"  ℹ️  Coluna polo: '{col}' | Valores encontrados: {unicos[:15]}")
    log_fn(f"  ℹ️  Buscando: {sorted(targets)}")

    resultado = df[valores_norm.isin(targets)]
    log_fn(f"  🔍 Polo={polo}: {len(resultado)}/{len(df)} linhas mantidas.")
    return resultado


# ── App ───────────────────────────────────────────────────────────────────────

class App:
    def __init__(self):
        self._cfg: AppConfig = load_config()
        self._scheduler: Scheduler | None = None
        self._cancelado = threading.Event()
        self._window = MainWindow(
            cfg=self._cfg,
            on_executar=self._iniciar,
            on_parar=self._parar,
            on_tratar=self._tratar,
        )
        self._window.quit_signal.connect(self._encerrar)

    def run(self):
        self._window.mainloop()

    # ── Controle do scheduler ─────────────────────────────────────────────────

    def _iniciar(self, selecionados: dict, data_ini: datetime, data_fim: datetime,
                 modo: str, polo: str = "Todos os polos"):
        if self._scheduler and self._scheduler.rodando:
            self._scheduler.stop()

        self._cancelado.set()
        self._cancelado = threading.Event()
        token = self._cancelado

        def fluxo():
            if token.is_set():
                return
            # Recalcula datas a cada iteração para detectar virada de dia
            agora = datetime.now()
            if modo == "uma_vez":
                ini, fim = data_ini, data_fim
            else:
                # Execução contínua: usa data_ini relativa + hoje
                fim = agora.replace(hour=23, minute=59, second=59)
                delta = data_fim - data_ini
                ini = (fim - delta).replace(hour=0, minute=0, second=0)

            self._executar_fluxo(selecionados, ini, fim, token, polo=polo)
            self._window.set_rodando(False)
            if modo == "uma_vez" and not token.is_set():
                self._window.mostrar_concluido("Fluxo concluído com sucesso!")

        self._scheduler = Scheduler(fn=fluxo, modo=modo, log_fn=self._log)
        self._scheduler.start()

    def _parar(self):
        self._cancelado.set()
        if self._scheduler:
            self._scheduler.stop()
        self._log("⏹ Execução interrompida pelo usuário.")

    def _encerrar(self):
        self._parar()
        try:
            self._window._tray.hide()
        except Exception:
            pass
        app = QApplication.instance()
        if app:
            app.quit()

    def _notificar_operview(self, login_event=None):
        """Dispara a notificação de login na GUI (thread-safe via sinal)."""
        try:
            self._window.operview_login_signal.emit(login_event)
        except Exception:
            pass

    # ── Fluxo principal ───────────────────────────────────────────────────────

    def _executar_fluxo(self, selecionados: dict, data_ini: datetime, data_fim: datetime,
                        token: threading.Event | None = None,
                        polo: str = "Todos os polos"):
        cfg = self._cfg
        pasta = str(Path(cfg.pasta_local).resolve())

        if not pasta:
            self._log("❌ Pasta local não configurada. Vá em Configurações.")
            return

        Path(pasta).mkdir(parents=True, exist_ok=True)
        self._log(f"  📁 Pasta de saída: {pasta}")

        # Limpa artefatos de execuções anteriores
        for nome in ("incidencias_tratado.xlsx", "incidencias_tratado.json", "dashboard_incidencias.html"):
            p = Path(pasta) / nome
            try:
                if p.exists():
                    p.unlink()
            except PermissionError:
                pass

        self._log(
            f"🚀 Iniciando — {data_ini.strftime('%d/%m/%Y')} → {data_fim.strftime('%d/%m/%Y')}"
        )
        if polo != "Todos os polos":
            self._log(f"  🔍 Filtro de polo: {polo}")

        if selecionados.get("geonline"):
            try:
                if cfg.origem_relatorio == "operview":
                    self._log("📥 Baixando Operview — Incidências...")
                    baixar = baixar_operview
                    cfg_dl = cfg.operview
                    kwargs = {"notificar": self._notificar_operview}
                else:
                    self._log("📥 Baixando GeoOnline — Incidências...")
                    baixar = baixar_incidencias
                    cfg_dl = cfg.geonline
                    kwargs = {}
                arquivo = baixar(
                    cfg_dl, pasta, data_ini, data_fim, self._log, token,
                    polo=polo, **kwargs,
                )

                if token and token.is_set():
                    self._log("⏹ Cancelado após download.")
                    return

                self._log("⚙️  Tratando incidências...")
                df_full = tratar_incidencias(arquivo, self._log)
                df = _filtrar_polo(df_full, polo, self._log)

                if token and token.is_set():
                    self._log("⏹ Cancelado após tratamento.")
                    return

                saida = Path(pasta) / "incidencias_tratado.xlsx"
                try:
                    for attempt in range(6):
                        try:
                            tmp_saida = saida.with_suffix(".tmp_%d.xlsx" % attempt)
                            df.to_excel(str(tmp_saida), index=False)
                            try:
                                os.replace(str(tmp_saida), str(saida))
                            except OSError:
                                shutil.copy2(str(tmp_saida), str(saida))
                            if tmp_saida.exists():
                                try:
                                    tmp_saida.unlink()
                                except OSError:
                                    pass
                            break
                        except (OSError, PermissionError) as e:
                            try:
                                if tmp_saida.exists():
                                    tmp_saida.unlink()
                            except OSError:
                                pass
                            if attempt == 5:
                                raise
                            time.sleep(1)
                    self._log(f"  💾 Salvo: {saida} ({len(df)} linhas)")
                except (OSError, PermissionError) as e:
                    self._log(
                        f"  ⚠️ Não foi possível salvar o xlsx "
                        f"(arquivo pode estar aberto no Excel/OneDrive): {e}"
                    )

                if cfg.base_mensal_enabled:
                    atualizar_base_mensal(df, "incidencias", pasta, "Data Início", self._log)
                else:
                    self._log("  ⏭️  Base mensal desativada.")

                try:
                    self._log("  📊 Gerando dashboard HTML...")
                    dash = Path(pasta) / "dashboard_incidencias.html"
                    gerar_dashboard_html(df_full, dash, self._log)
                    self._log(f"  💾 Salvo: {dash}")
                except Exception as e:
                    self._log(f"❌ Dashboard falhou: {e}")

                if cfg.n8n.enabled and cfg.n8n.webhook:
                    self._log("  📡 Enviando análises para n8n...")
                    payload = gerar_analises_n8n(df)
                    enviar(payload, cfg.n8n.webhook, self._log)
                else:
                    self._log("  ⏭️  Envio n8n desativado.")

            except Exception as e:
                self._log(f"❌ Falha no download de incidências: {e}")

        if cfg.sharepoint.enabled and cfg.sharepoint.pasta:
            if token and token.is_set():
                return
            self._log("📤 Sincronizando com SharePoint...")
            try:
                # Coleta arquivos gerados nesta execução
                arquivos = []
                tratado = Path(pasta) / "incidencias_tratado.xlsx"
                if tratado.exists():
                    arquivos.append(tratado)
                dash = Path(pasta) / "dashboard_incidencias.html"
                if dash.exists():
                    arquivos.append(dash)
                base_dir = Path(pasta) / "bases_mensais" / "incidencias"
                if base_dir.exists():
                    agora = time.time()
                    for f in base_dir.iterdir():
                        if f.is_file() and (agora - f.stat().st_mtime) < 600:
                            arquivos.append(f)
                sincronizar(arquivos, cfg.sharepoint.pasta, self._log)
            except Exception as e:
                self._log(f"❌ SharePoint falhou: {e}")

        self._log("🏁 Fluxo concluído.")

    # ── Tratamento avulso ─────────────────────────────────────────────────────

    def _tratar(self, arquivo: Path, polo: str, log_fn):
        def _run():
            try:
                log_fn(f"📄 Tratando: {arquivo.name}")
                df = tratar_incidencias(arquivo, log_fn)
                _win = getattr(self, "_window", None)
                if _win is not None and hasattr(_win, "formato_signal"):
                    try:
                        _win.formato_signal.emit(df.attrs.get("formato", "Desconhecido"))
                    except Exception:
                        pass
                df = _filtrar_polo(df, polo, log_fn)
                if polo != "Todos os polos":
                    log_fn(f"  🔍 Polo: {polo} — {len(df)} linhas após filtro")
                saida = arquivo.parent / (arquivo.stem + "_tratado.xlsx")
                tmp_saida = saida.with_suffix(".tmp.xlsx")
                df.to_excel(str(tmp_saida), index=False)
                try:
                    os.replace(str(tmp_saida), str(saida))
                except OSError:
                    if saida.exists():
                        for _ in range(10):
                            try:
                                saida.unlink()
                                break
                            except OSError:
                                time.sleep(0.5)
                    shutil.copy2(str(tmp_saida), str(saida))
                    tmp_saida.unlink()
                log_fn(f"✅ Salvo: {saida.name} ({len(df)} linhas)")
            except Exception as e:
                log_fn(f"❌ Erro no tratamento: {e}")
            finally:
                self._window.set_tratando(False)

        self._window.set_tratando(True)
        threading.Thread(target=_run, daemon=True).start()

    # ── Log helper ────────────────────────────────────────────────────────────

    def _log(self, msg: str):
        self._window.log(msg)


if __name__ == "__main__":
    App().run()
