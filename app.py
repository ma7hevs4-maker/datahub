"""
DataHub v2 — Entry point.
Conecta GUI → Scheduler → Downloaders → Processors → Integrations.
"""
import os
import json
import shutil
import threading
import time
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

from config import load_config, AppConfig
from core.scheduler import Scheduler
from core.downloader.geonline import baixar_incidencias
from core.downloader.operview import baixar_incidencias as baixar_operview
try:
    from core.downloader.spotfire import (
        baixar_deslocamentos_hoje,
        baixar_bases_scanner_mensais,
        NOME_M300_HOJE,
    )
except Exception:
    # Se o build não incluir o módulo, o Operview continua funcionando.
    baixar_deslocamentos_hoje = None
    baixar_bases_scanner_mensais = None
    NOME_M300_HOJE = "m300_deslocamentos_hoje.csv"
from core.processors.incidencias import tratar_incidencias, gerar_analises_n8n
from core.processors.dashboard_html import gerar_dashboard_html
from core.base_mensal import atualizar_base_mensal
from integrations.sharepoint import sincronizar
from integrations.n8n import enviar
try:
    from db_sync import gravar_incidencias  # banco PostgreSQL (opcional)
except Exception:
    gravar_incidencias = None
try:
    from db_m300 import (  # banco PostgreSQL — tabelas do M300 (opcional)
        gravar_deslocamentos_hoje,
        gravar_deslocamentos_total,
        gravar_tabela_completa,
        ler_csv_hoje,
    )
except Exception:
    gravar_deslocamentos_hoje = gravar_deslocamentos_total = None
    gravar_tabela_completa = ler_csv_hoje = None
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


# ── Estado do M300 (por horário: >=08:00 / >=12:00) ──────────────────────

def _m300_state_path() -> Path:
    base = Path(os.environ.get("APPDATA") or Path.home()) / "DataHub"
    base.mkdir(parents=True, exist_ok=True)
    return base / "m300_estado.json"


def _ler_estado_m300() -> dict:
    try:
        p = _m300_state_path()
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
    except Exception:
        pass
    return {}


def _salvar_estado_m300(estado: dict) -> None:
    try:
        p = _m300_state_path()
        tmp = p.with_suffix(".tmp")
        tmp.write_text(
            json.dumps(estado, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(str(tmp), str(p))
    except Exception:
        pass


def _m300_loop_atual() -> dict:
    """Incrementa e devolve o contador de loops do M300 no dia.

    O estado fica em %APPDATA%\\DataHub\\m300_estado.json para sobreviver a
    reinícios do app. O contador zera na virada do dia.
    """
    hoje = datetime.now().strftime("%Y-%m-%d")
    estado = _ler_estado_m300()
    if estado.get("data") != hoje:
        estado = {
            "data": hoje,
            "loops": 0,
            "geral_em": None,
            "base_ate": None,
            "geral_08h": False,
            "geral_12h": False,
        }
    if "geral_08h" not in estado:
        estado["geral_08h"] = False
    if "geral_12h" not in estado:
        estado["geral_12h"] = False
    try:
        estado["loops"] = int(estado.get("loops", 0)) + 1
    except Exception:
        estado["loops"] = 1
    _salvar_estado_m300(estado)
    return estado


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

    def _iniciar(self, selecionados: list, data_ini: datetime, data_fim: datetime,
                 modo: str, polo: str = "Todos os polos"):
        if self._scheduler and self._scheduler.rodando:
            self._scheduler.stop()

        self._cancelado.set()
        self._cancelado = threading.Event()
        token = self._cancelado
        self._n8n_loop = 0

        def fluxo():
            self._window.set_rodando(True)
            if token.is_set():
                return
            self._n8n_loop += 1
            loop_atual = self._n8n_loop
            # Recalcula datas a cada iteração para detectar virada de dia
            agora = datetime.now()
            if modo == "uma_vez":
                ini, fim = data_ini, data_fim
            else:
                # Execução contínua: usa data_ini relativa + hoje
                fim = agora.replace(hour=23, minute=59, second=59)
                delta = data_fim - data_ini
                ini = (fim - delta).replace(hour=0, minute=0, second=0)

            self._executar_fluxo(
                selecionados, ini, fim, token, polo=polo, modo=modo,
                loop_atual=loop_atual,
            )
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

    def _salvar_xlsx_seguro(self, df, saida: Path, log_fn=print):
        for attempt in range(6):
            try:
                tmp = saida.with_suffix(".tmp_%d.xlsx" % attempt)
                df.to_excel(str(tmp), index=False)
                try:
                    os.replace(str(tmp), str(saida))
                except OSError:
                    shutil.copy2(str(tmp), str(saida))
                if tmp.exists():
                    try:
                        tmp.unlink()
                    except OSError:
                        pass
                log_fn(f"  💾 Salvo: {saida} ({len(df)} linhas)")
                return
            except (OSError, PermissionError) as e:
                try:
                    if tmp.exists():
                        tmp.unlink()
                except OSError:
                    pass
                if attempt == 5:
                    log_fn(
                        f"  ⚠️ Não foi possível salvar {saida} "
                        f"(arquivo pode estar aberto no Excel/OneDrive): {e}"
                    )
                    return
                time.sleep(1)

    def _executar_fluxo(self, selecionados: list, data_ini: datetime, data_fim: datetime,
                        token: threading.Event | None = None,
                        polo: str = "Todos os polos",
                        modo: str = "uma_vez",
                        loop_atual: int = 1):
        cfg = self._cfg
        pasta = str(Path(cfg.pasta_local).resolve())

        if not pasta:
            self._log("❌ Pasta local não configurada. Vá em Configurações.")
            return

        Path(pasta).mkdir(parents=True, exist_ok=True)
        self._log(f"  📁 Pasta de saída: {pasta}")

        # Limpa artefatos de execuções anteriores
        for nome in ("incidencias_tratado.xlsx", "incidencias_tratado.json",
                     "dashboard_incidencias.html"):
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

        # Fontes independentes: uma falha não impede a outra nem o SharePoint.
        if "operview" in selecionados:
            try:
                self._executar_operview(
                    data_ini, data_fim, token, polo=polo,
                    modo=modo, loop_atual=loop_atual,
                )
            except Exception as e:
                self._log(f"❌ Operview falhou: {e}")
                import traceback as _tb2
                _trace2 = _tb2.format_exc()
                self._log(_trace2)
                try:
                    with open(r"C:\Users\BR0163806927\Downloads\operview_erro.txt", "w", encoding="utf-8") as _f:
                        _f.write(_trace2)
                except Exception:
                    pass

        if "m300" in selecionados:
            if token and token.is_set():
                return
            try:
                self._executar_m300(token, modo=modo)
            except Exception as e:
                self._log(f"❌ M300 falhou: {e}")
                import traceback as _tb3
                self._log(_tb3.format_exc())

        if cfg.sharepoint.enabled and cfg.sharepoint.pasta:
            if token and token.is_set():
                return
            self._log("📤 Sincronizando com SharePoint...")
            try:
                arquivos = []
                for nome in (
                    "incidencias_tratado.xlsx",
                    "dashboard_incidencias.html",
                    NOME_M300_HOJE,
                ):
                    f = Path(pasta) / nome
                    if f.exists():
                        arquivos.append(f)
                for sub in ("incidencias", "deslocamentos", "tabela_completa"):
                    base_dir = Path(pasta) / "bases_mensais" / sub
                    if base_dir.exists():
                        for f in base_dir.iterdir():
                            if f.is_file() and f.suffix.lower() == ".xlsx":
                                arquivos.append(f)
                if arquivos:
                    sincronizar(arquivos, cfg.sharepoint.pasta, self._log)
                else:
                    self._log("  ⚠️  Nenhum arquivo para sincronizar.")
            except Exception as e:
                self._log(f"❌ SharePoint falhou: {e}")

        self._log("🏁 Fluxo concluído.")

    def _executar_operview(self, data_ini, data_fim, token,
                           polo: str = "Todos os polos",
                           modo: str = "uma_vez",
                           loop_atual: int = 1):
        cfg = self._cfg
        pasta = str(Path(cfg.pasta_local).resolve())
        self._log("📥 Baixando Operview — Incidências...")
        arquivo = baixar_operview(
            cfg.operview, pasta, data_ini, data_fim, self._log, token,
            polo=polo, notificar=self._notificar_operview,
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
        self._salvar_xlsx_seguro(df, saida, self._log)

        if cfg.base_mensal_enabled:
            atualizar_base_mensal(df, "incidencias", pasta, "Data Início", self._log)
        else:
            self._log("  ⏭️  Base mensal desativada.")

        if cfg.db_enabled:
            if gravar_incidencias is not None:
                try:
                    gravar_incidencias(df_full, log_fn=self._log)
                except Exception as e:
                    self._log(f"❌ Banco falhou (o fluxo continua): {e}")
            else:
                self._log("  ⚠️  Banco habilitado, mas módulo db_sync não está disponível.")
        else:
            self._log("  ⏭️  Envio para PostgreSQL desativado.")

        try:
            self._log("  📊 Gerando dashboard HTML...")
            dash = Path(pasta) / "dashboard_incidencias.html"
            gerar_dashboard_html(df_full, dash, self._log)
            self._log(f"  💾 Salvo: {dash}")
        except Exception as e:
            self._log(f"❌ Dashboard falhou: {e}")

        enviar_n8n = True
        if modo == "3min":
            enviar_n8n = (loop_atual % 3 == 0)

        if not enviar_n8n:
            self._log(
                f"  ⏭️  n8n: modo 3 min, loop {loop_atual}; "
                "envio apenas a cada 3 loops."
            )
        elif cfg.n8n.enabled and cfg.n8n.webhook:
            self._log("  📡 Enviando análises para n8n...")
            payload = gerar_analises_n8n(df)
            enviar(payload, cfg.n8n.webhook, self._log)
        else:
            self._log("  ⏭️  Envio n8n desativado.")

    # ── M300 (deslocamentos) ──────────────────────────────────────────────────

    def _executar_m300(self, token, modo: str = "uma_vez"):
        """Baixa os deslocamentos do dia e as bases mensais do Scanner.

        Hoje (Produtividade): Deslocamentos + Painel Acompanhamento, com merge
        por Equipe. Scanner: janela de 2 meses com upsert em
        `bases_mensais/deslocamentos` e `bases_mensais/tabela_completa`.
        Em 'uma_vez', o Scanner roda sempre. Nos modos de loop, o Scanner roda
        pela primeira vez a partir das 08:00; se às 12:00 ainda não tiver dados
        do dia anterior (d-1) nas bases mensais, faz conferência/baixa novamente.
        O estado fica em %APPDATA%\\DataHub\\m300_estado.json.
        """
        cfg = self._cfg
        pasta = str(Path(cfg.pasta_local).resolve())
        if not (cfg.spotfire.url or cfg.spotfire.url_hoje):
            self._log("  ⚠️  M300: URL do Spotfire não configurada — pulando.")
            return
        if (baixar_deslocamentos_hoje is None
                or baixar_bases_scanner_mensais is None):
            self._log("  ⚠️  M300 indisponível neste build (módulo Spotfire não incluído).")
            return

        estado = _m300_loop_atual()
        loops = estado.get("loops", 1)
        data = estado.get("data", datetime.now().strftime("%Y-%m-%d"))
        agora = datetime.now()
        ontem = (
            agora.replace(hour=0, minute=0, second=0, microsecond=0)
            - timedelta(days=1)
        ).strftime("%Y-%m-%d")
        base_ate = str(estado.get("base_ate") or "")
        desloc_dir = Path(pasta) / "bases_mensais" / "deslocamentos"
        tc_dir = Path(pasta) / "bases_mensais" / "tabela_completa"
        tem_bases = (
            desloc_dir.exists() and any(desloc_dir.glob("*.xlsx"))
            and tc_dir.exists() and any(tc_dir.glob("*.xlsx"))
        )
        self._log(f"  🔁 M300 — loop {loops} do dia ({data}) | {agora.strftime('%H:%M:%S')}")

        def _liberado():
            tentativa = estado.get("geral_tentativa")
            if not tentativa:
                return True
            try:
                return (
                    agora - datetime.fromisoformat(tentativa)
                ) >= timedelta(minutes=30)
            except Exception:
                return True

        deve_geral = False
        if modo == "uma_vez":
            deve_geral = True
            self._log("  ℹ️  Modo 'Uma vez': Scanner também será executado.")
        elif not tem_bases:
            deve_geral = True
            self._log("  ℹ️  Bases mensais ainda não existem — baixando nesta execução.")
        else:
            passou_08h = agora.hour >= 8
            passou_12h = agora.hour >= 12
            ja_rodou_08h = bool(estado.get("geral_08h"))
            ja_rodou_12h = bool(estado.get("geral_12h"))
            ja_rodou_hoje = bool(estado.get("geral_em")) and estado.get("geral_em") == data

            if passou_08h and not ja_rodou_08h and not ja_rodou_hoje:
                if _liberado():
                    deve_geral = True
                    estado["geral_08h"] = True
                    self._log("  ℹ️  A partir das 08:00 e ainda não rodou hoje: executando as bases mensais do Scanner.")
                else:
                    self._log("  ⏳ Aguardando 30 min desde a última tentativa para baixar as bases mensais.")
            elif passou_12h and not ja_rodou_12h:
                if tem_bases and base_ate and base_ate >= ontem:
                    estado["geral_12h"] = True
                    _salvar_estado_m300(estado)
                    self._log(
                        f"  ✅ Conferência às 12:00: a base já contém {ontem} (d-1); nada a baixar."
                    )
                elif _liberado():
                    deve_geral = True
                    estado["geral_12h"] = True
                    self._log(
                        f"  ⚠️  Às 12:00 a base não contém {ontem} (d-1); executando o Scanner."
                    )
                else:
                    self._log("  ⏳ Aguardando 30 min desde a última tentativa para conferência às 12:00.")
            elif passou_08h and not ja_rodou_hoje and not ja_rodou_08h:
                if _liberado():
                    deve_geral = True
                    estado["geral_08h"] = True
                    self._log("  ℹ️  A partir das 08:00: executando as bases mensais do Scanner.")
                else:
                    self._log("  ⏳ Aguardando 30 min desde a última tentativa para baixar as bases mensais.")

        if deve_geral:
            estado["geral_tentativa"] = agora.isoformat(timespec="seconds")
            _salvar_estado_m300(estado)

        self._log("📥 M300 — deslocamentos do dia atual (Deslocamentos + Painel)...")
        try:
            p = baixar_deslocamentos_hoje(cfg.spotfire, pasta, self._log, token)
            self._log(f"  ✅ Hoje: {p.name} ({p.stat().st_size} bytes)")
            self._gravar_m300_banco("hoje", p)
        except Exception as e:
            self._log(f"  ❌ M300 hoje falhou: {e}")

        if not deve_geral:
            return
        if token and token.is_set():
            return

        self._log("📥 M300 — bases mensais (mês atual + anterior, sem filtros extras)...")
        try:
            arquivos, base_ate_novo = baixar_bases_scanner_mensais(
                cfg.spotfire, pasta, self._log, token
            )
            estado["geral_em"] = data
            if base_ate_novo:
                estado["base_ate"] = base_ate_novo
            _salvar_estado_m300(estado)
            self._log(
                "  ✅ Bases mensais: "
                + ", ".join(f"{p.parent.name}/{p.name}" for p in arquivos)
            )
            if base_ate_novo:
                self._log(f"  🗓️  Última data na base: {base_ate_novo}")
            for p in arquivos:
                self._gravar_m300_banco(p.parent.name, p)
        except Exception as e:
            self._log(f"  ❌ M300 bases mensais falharam: {e}")

    def _gravar_m300_banco(self, tipo: str, caminho: Path):
        """Grava um arquivo do M300 no PostgreSQL. Nunca derruba o fluxo."""
        if not self._cfg.db_enabled:
            return
        if gravar_deslocamentos_hoje is None:
            self._log("  ⚠️  Banco habilitado, mas módulo db_m300 não está disponível.")
            return
        try:
            if tipo == "hoje":
                gravar_deslocamentos_hoje(ler_csv_hoje(str(caminho)), log_fn=self._log)
            else:
                import pandas as _pd
                df = _pd.read_excel(str(caminho))
                if tipo == "deslocamentos":
                    gravar_deslocamentos_total(df, log_fn=self._log)
                elif tipo == "tabela_completa":
                    gravar_tabela_completa(df, log_fn=self._log)
        except Exception as e:
            self._log(f"❌ Banco (M300 {tipo}) falhou (o fluxo continua): {e}")

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
    import sys
    if "--gravar-login" in sys.argv[1:]:
        from config import load_config
        from core.downloader.operview import gravar_login_operview
        _cfg = load_config()
        gravar_login_operview(cfg=_cfg.operview, log_fn=print)
        sys.exit(0)
    App().run()


