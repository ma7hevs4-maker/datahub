# -*- coding: utf-8 -*-
"""Gera um dashboard HTML interativo replicando o conceito do Hub de Emergencia.

Duas abas (Dashboard 1 polo / Dashboard 2 todos os polos). No topo de cada aba
ficam apenas os filtros universais UT / Polo / NT, aplicados a toda a pagina.

Abaixo dos KPIs e do Top 10, os blocos segmentados ficam todos na mesma pagina
(Entrada, Saida, Entrada&Saida, Em Aberto, +24h, Ordem 2, 5 Regras na aba 1;
Entrada, Saida, Em Aberto, TMA, Processos na aba 2); cada bloco tem seus
proprios filtros (Sucursal, Entrada, Saida) condicionais ao polo/UT escolhido,
e os botoes do segbar rolam a pagina ate o bloco correspondente.

Labels dentro das barras: fonte branca sem fundo; totais no topo das barras:
caixa branca com fonte preta. Graficos de linha: fundo branco e fonte preta.
O KPI "Abertos Agora" fica a esquerda de "Abertos +24h". O Top 10 (abertos)
mostra Número, Polo, Alimentador, Causa, Clts Atual e a % do total de clientes.

Duplo clique em barras, KPIs, Top 10 ou em um grafico maximizado abre um modal
com TODAS as colunas do incidente correspondente (payload f/c), com botoes de
exportar (.xlsx via SheetJS / .csv).

A partir do arquivo de incidencias ja tratado pelo DataHub
(incidencias_tratado.xlsx).

Uso standalone:
    python dashboard_html.py <incidencias_tratado.xlsx> [saida.html]

Integrado no DataHub, chamar:
    gerar_dashboard_html(df, caminho_saida, log_fn)

Regras globais:
    - Desconsiderar incidentes com Estado = programado, programada ou anulado.
"""

import json
import os
import sys
import shutil
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

if sys.stdout and sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


ESTADO_DESCARTAR = {"programado", "programada", "anulado", "anulada"}


def _norm(v):
    if v is None:
        return ""
    return str(v).strip()


def _carregar_template():
    """Template do dashboard. Prioriza o arquivo externo (versao validada);
    cai no _TEMPLATE embutido se o arquivo nao existir."""
    candidatos = [
        Path(__file__).parent / "dashboard_template.html",
        Path(sys.executable).parent / "dashboard_template.html",
        Path.cwd() / "dashboard_template.html",
    ]
    for c in candidatos:
        try:
            if c.exists():
                return c.read_text(encoding="utf-8")
        except OSError:
            pass
    return _TEMPLATE


def _write_retry(saida, content, encoding="utf-8", tries=6, delay=1.0):
    """Grava o arquivo com tentativas — contorna locks transitivos (WinError 32)."""
    saida = Path(saida)
    last = None
    for _ in range(tries):
        try:
            tmp = saida.with_suffix(saida.suffix + ".tmp")
            tmp.write_text(content, encoding=encoding)
            try:
                os.replace(str(tmp), str(saida))
            except OSError:
                shutil.copy2(str(tmp), str(saida))
            if tmp.exists():
                try:
                    tmp.unlink()
                except OSError:
                    pass
            return
        except (OSError, PermissionError) as e:
            last = e
            time.sleep(delay)
    raise last


def gerar_dashboard_html(df, saida=None, log_fn=print):
    if saida is None:
        saida = Path.cwd() / "dashboard_incidencias.html"
    saida = Path(saida)
    saida.parent.mkdir(parents=True, exist_ok=True)

    log_fn("  [*] Gerando dashboard HTML...")

    df = df.copy()

    def col(*nomes):
        for n in nomes:
            if n in df.columns:
                return n
        return None

    # ---- Descartar Estado programado / programada / anulado -----------------
    c_estado = col("Estado")
    if c_estado:
        mask = (
            df[c_estado].astype(str).str.strip().str.lower()
            .isin(ESTADO_DESCARTAR)
        )
        n_desc = int(mask.sum())
        if n_desc:
            df = df[~mask]
            log_fn(f"  [i] {n_desc} linhas descartadas por Estado (programado/anulado)")

    # ---- Descartar Causa atend cancelado cr / atendimento cancelado / programada manutenção / programada obras ----
    c_causa = col("Causa")
    if c_causa:
        mask_c = (
            df[c_causa].astype(str).str.strip().str.lower()
            .isin({"atend cancelado cr", "atendimento cancelado", "programada manutenção", "programada obras"})
        )
        n_c = int(mask_c.sum())
        if n_c:
            df = df[~mask_c]
            log_fn(f"  [i] {n_c} linhas descartadas por Causa (atend cancelado cr / atendimento cancelado / programada manutenção / programada obras)")

    c_polo = col("Polo")
    c_tens = col("Nivel de Tensao", "Nivel de Tensão")
    c_nt = col("NT")
    c_status = col("Status")
    c_entrada = col("Entrada")
    c_saida = col("Saída")
    c_hora_e = col("Hora entrada")
    c_hora_s = col("Hora saída", "Hora saida")
    c_turno = col("Turno abertura")
    c_turno_s = col("Turno fechamento")
    c_cluster = col("Cluster duração", "Cluster duracao")
    c_dur = col("Duração em minutos", "Duracao em minutos")
    c_rein = col("Reincidente")
    c_rein_tipo = col("Reincidente tipo")
    c_causa = col("Causa")
    c_al = col("AL")
    c_ordem = col("Ordem 2 Consolidada")
    c_num = col("Número", "Numero")
    c_cli = col("Nº Cliente", "N Cliente")
    c_tma = col("TMA")
    c_tme = col("TME")
    c_tme_ana = col("TME Análise", "TME Analise")
    c_tmp = col("TMP")
    c_tmd = col("TMD")
    c_considera = col("Considerar TMA")
    c_prod = col("Produtivo/Improdutivo")
    c_proc = col("Processo (Equipe atrib.)")
    c_grupo = col("Grupo Processos (Equipe atrib.)")
    c_grupo_d = col("Grupo Processos (Equipe desloc.)")
    c_dia = col("Dia")
    c_mes = col("Mês")
    c_mes_sd = col("Mês Saída")
    c_dia_sd = col("Dia Saída")
    c_equipe = col("Equipe Atribuída", "Equipe Atribuida")
    c_equi_d = col("Equipe Desl.", "Equipe Desloc.")
    c_cli_af = col("Cli Af Max")
    c_cli_af3 = col("Cli Af 3 min")
    c_clts = col("Clts Atual")
    c_cli_tmp = col("Cli x Tmp")
    c_equi_op = col("Equipe/Operador")
    c_regras = col("Cumpre Regras Ouro")
    c_osm = col("OSM")
    c_suc = col("Sucursal")
    c_atr = col("Atribuição")

    horas = [str(h) for h in range(24)]

    full_cols = [str(c) for c in df.columns]

    def sv(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        return str(v).strip()

    rows = []
    full_rows = []
    for i, r in df.iterrows():
        polo = _norm(r.get(c_polo))
        if polo in ("", "---", "nan", "None"):
            continue

        def val(c):
            if c is None:
                return None
            v = r.get(c)
            if pd.isna(v):
                return None
            return v

        full_rows.append([sv(r.get(c)) for c in df.columns])

        tens = _norm(val(c_tens))
        if tens.upper() == "OSM":
            tens = "OSM"
        elif tens.upper() in ("BT", "MT", "AT"):
            tens = tens.upper()
        else:
            tens = "Outro"

        status = _norm(val(c_status))
        entrada = _norm(val(c_entrada))
        saida_col = _norm(val(c_saida))
        hora_e = _norm(val(c_hora_e))
        if hora_e not in horas:
            hora_e = ""
        hora_s = _norm(val(c_hora_s))
        if hora_s not in horas:
            hora_s = ""

        def _band_turno(h):
            try:
                x=int(str(h).split(":")[0])
            except (TypeError, ValueError, AttributeError):
                return ""
            if x<8: return "A"
            if x<16: return "B"
            return "C"
        turno_a = _band_turno(hora_e) or _norm(val(c_turno))
        turno_s = _band_turno(hora_s) or _norm(val(c_turno_s))
        cluster = _norm(val(c_cluster))
        rein_tipo = _norm(val(c_rein_tipo))
        causa = _norm(val(c_causa))
        ordem = _norm(val(c_ordem)).lower()

        num = _norm(val(c_num))
        cli = _norm(val(c_cli))

        def to_int(v, d=0):
            if v is None:
                return d
            try:
                return int(float(v))
            except (TypeError, ValueError):
                return d

        def to_num(v, d=0.0):
            if v is None:
                return d
            try:
                return float(v)
            except (TypeError, ValueError):
                return d

        cli_af = to_int(val(c_cli_af))
        cli_af3 = to_int(val(c_cli_af3))
        clts = to_int(val(c_clts))
        cxt = to_num(val(c_cli_tmp))
        dur = to_int(val(c_dur))
        tma = to_int(val(c_tma))
        tme = to_int(val(c_tme))
        tme_ana = _norm(val(c_tme_ana))

        considera = _norm(val(c_considera))
        prod = _norm(val(c_prod))
        proc = _norm(val(c_proc))
        grupo = _norm(val(c_grupo))
        grupo_d = _norm(val(c_grupo_d))
        dia = _norm(val(c_dia))
        mes = _norm(val(c_mes))
        equipe = _norm(val(c_equipe))
        equi_d = _norm(val(c_equi_d))
        equi_op = _norm(val(c_equi_op)).lower()
        regras = _norm(val(c_regras)).upper()
        if regras in ("S", "SIM", "Y", "YES", "1"):
            regras = "S"
        elif regras in ("N", "NAO", "NÃO", "NO", "0"):
            regras = "N"
        osm = _norm(val(c_osm))
        suc = _norm(val(c_suc))
        atr = _norm(val(c_atr))
        nt = _norm(val(c_nt)).upper()
        al = _norm(val(c_al))

        rows.append({
            "n": num,
            "p": polo,
            "nt": nt,
            "t": tens,
            "s": status,
            "e": entrada,
            "sd": saida_col,
            "he": hora_e,
            "hs": hora_s,
            "ta": turno_a,
            "ts": turno_s,
            "c": cluster,
            "dur": dur,
            "rt": rein_tipo,
            "ca": causa,
            "o": ordem,
            "cli": cli,
            "af": cli_af,
            "a3": cli_af3,
            "ct": clts,
            "cxt": cxt,
            "tma": tma,
            "tmp": to_int(val(c_tmp)),
            "tmd": to_int(val(c_tmd)),
            "tme": tme,
            "tmea": tme_ana,
            "cg": considera,
            "pd": prod,
            "pr": proc,
            "gp": grupo,
            "gpd": grupo_d,
            "d": dia,
            "mes": mes,
            "d_sd": _norm(val(c_dia_sd)),
            "mes_sd": _norm(val(c_mes_sd)),
            "diaMes": dia,
            "eq": equipe,
            "eqd": equi_d,
            "eo": equi_op,
            "ro": regras,
            "osm": osm,
            "suc": suc,
            "atr": atr,
            "al": al,
            "ea": equipe,
        })

    html = _carregar_template().replace(
        "__DADOS__",
        json.dumps({"m": rows, "f": full_rows, "c": full_cols},
                   ensure_ascii=True, separators=(",", ":")),
    ).replace(
        "__BUILD_AT__",
        datetime.now().strftime("%d/%m/%Y %H:%M"),
    )
    _write_retry(saida, html)
    log_fn(f"  [OK] Dashboard gerado: {saida} ({len(rows)} linhas)")
    return str(saida)


_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DataHub — Dashboard de Incidências</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
<style>
:root{
  --bg:#eef2f7; --card:#ffffff; --card2:#f7fafc; --border:#d8e0ea;
  --txt:#1f2d3d; --mut:#5b6b7b; --acc:#1f6feb; --acc2:#dceafd; --head:#1c2b3d;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--txt);font-family:'Segoe UI',Arial,sans-serif;font-size:14px}
header{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:16px 22px;background:var(--head);color:#fff;flex-wrap:wrap}
header h1{font-size:18px;margin:0}
header .sub{color:#c8d4e3;font-size:12px}
#meta-info{padding-right:62px}
.tabs{display:flex;gap:4px;padding:12px 22px 0}
.tab{padding:9px 18px;border:1px solid var(--border);border-bottom:none;border-radius:8px 8px 0 0;background:var(--card2);color:var(--txt);cursor:pointer;font-weight:600}
.tab.active{background:var(--card);color:var(--acc);border-color:var(--border)}
.tab:hover:not(.active){background:#e6ecf5}
.panel{display:none;padding:14px 22px 30px}
.panel.active{display:block}
.segbar{display:flex;gap:6px;flex-wrap:wrap;padding:8px 0 14px}
.segbtn{padding:7px 13px;border:1px solid var(--border);border-radius:7px;background:var(--card);color:var(--txt);cursor:pointer;font-weight:600;font-size:13px}
.segbtn:hover{background:var(--acc2);color:var(--acc);border-color:var(--acc)}
.segbtn .dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px;vertical-align:baseline}
.filters{display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;padding:12px 16px;background:var(--card);border:1px solid var(--border);border-radius:10px;margin-bottom:14px}
.filters .flabel{display:flex;flex-direction:column;gap:4px;font-size:12px;color:var(--mut)}
.msel{position:relative;display:inline-block}
.msel-btn{background:#fff;color:var(--txt);border:1px solid var(--border);border-radius:6px;padding:6px 10px;font-size:12px;cursor:pointer;max-width:260px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;text-align:left}
.msel-btn:hover{border-color:var(--acc)}
.export-btn{margin-left:auto;background:var(--acc);color:#fff;border:none;border-radius:6px;padding:7px 14px;font-size:12px;font-weight:600;cursor:pointer}
.export-btn:hover{filter:brightness(1.12)}
.msel-pop{position:absolute;z-index:500;top:100%;left:0;margin-top:4px;background:#fff;border:1px solid var(--border);border-radius:8px;box-shadow:0 6px 22px rgba(16,40,73,.18);padding:6px;max-height:260px;overflow:auto;min-width:170px}
.msel-opt{display:flex;align-items:center;gap:6px;padding:4px 6px;font-size:12px;color:var(--txt);cursor:pointer;border-radius:4px}
.msel-opt:hover{background:var(--acc2)}
.msel-opt.all{border-bottom:1px solid var(--border);padding-bottom:6px;margin-bottom:2px;font-weight:600}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:12px}
.kpi{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:11px 13px;box-shadow:0 1px 2px rgba(16,40,73,.06);cursor:pointer}
.kpi .v{font-size:24px;font-weight:700;color:#1a7f37}
.kpi.blue .v{color:var(--acc)} .kpi.red .v{color:#d1242f} .kpi.orange .v{color:#f0820f} .kpi.purple .v{color:#8250df} .kpi.cyan .v{color:#0aa2a2}
.kpi .l{color:var(--mut);font-size:11px}

.topbox{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:12px 14px;margin-bottom:14px;box-shadow:0 1px 2px rgba(16,40,73,.06)}
.tophead{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.tophead h4{margin:0;font-size:13px;color:var(--acc)}
.tophead .export-btn{margin-left:auto;padding:6px 12px}
.topseg{display:flex;gap:2px;background:var(--card2);border:1px solid var(--border);border-radius:7px;padding:2px}
.tsbtn{background:transparent;border:none;border-radius:5px;padding:5px 10px;font-size:12px;font-weight:600;cursor:pointer;color:var(--mut)}
.tsbtn.active{background:var(--acc);color:#fff}
.tsbtn:hover:not(.active){color:var(--acc)}
table.top{width:100%;border-collapse:collapse;font-size:12px}
table.top th,table.top td{padding:5px 8px;border-bottom:1px solid var(--border);text-align:left}
table.top th{color:var(--mut);font-weight:600}
table.top th.num{text-align:right}
table.top td.num{text-align:right}
table.top tbody tr{cursor:pointer}
table.top tbody tr:hover{background:var(--acc2)}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:14px}
.card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:14px;box-shadow:0 1px 2px rgba(16,40,73,.06);position:relative}
.card h3{margin:0 0 10px;font-size:13px;color:var(--acc);padding-right:26px}
.card canvas{max-height:290px;cursor:pointer}
.grid.g5{grid-template-columns:repeat(5,minmax(0,1fr))}
.card.span2x2{grid-column:span 2;grid-row:span 2}
.card.span2x2 canvas{height:592px!important;max-height:none}
.grid.g-tma{grid-template-columns:repeat(3,minmax(0,1fr))}
.card.span-all{grid-column:1 / -1}
@media(max-width:1000px){.grid.g-tma{grid-template-columns:1fr}.card.span-all{grid-column:auto}}
.stack{display:grid;grid-template-rows:1fr 1fr;gap:14px}
.vcard{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;min-height:150px;text-align:center}
.vcard .v{font-size:40px;font-weight:700;line-height:1}
.vcard .l{font-size:13px;color:var(--mut)}
#to-top{position:fixed;top:14px;right:16px;z-index:900;width:46px;height:46px;border-radius:50%;border:1px solid var(--border);background:var(--head);color:#fff;font-size:22px;cursor:pointer;box-shadow:0 3px 10px rgba(0,0,0,.35);display:flex;align-items:center;justify-content:center}
#to-top:hover{background:var(--acc)}
@media(max-width:1750px){.grid.g5{grid-template-columns:repeat(auto-fit,minmax(340px,1fr))}}
.segblock{padding:18px 0 8px;scroll-margin-top:10px}
.seghead{display:flex;align-items:center;gap:8px;font-size:15px;font-weight:700;color:var(--acc);padding:10px 2px 12px;flex-wrap:wrap}
.seghead .dot{display:inline-block;width:11px;height:11px;border-radius:50%}
.block-hint{font-size:11px;color:var(--mut);font-weight:400;margin-left:auto}
.print-btn{margin-left:8px;background:var(--card2);border:1px solid var(--border);border-radius:6px;padding:4px 10px;font-size:12px;font-weight:600;cursor:pointer;color:var(--txt)}
.print-btn:hover{background:var(--acc2);color:var(--acc);border-color:var(--acc)}
.search-inp{width:150px;padding:5px 9px;border:1px solid var(--border);border-radius:6px;font-size:12.5px;font-family:inherit;outline:none}
.search-btn{background:#fff;border:1px solid var(--border);border-radius:6px;cursor:pointer;padding:4px 8px;font-size:13px;line-height:1;color:#333}
.search-btn:hover{background:var(--acc2)}
.maxbtn{position:absolute;top:10px;right:10px;width:28px;height:28px;line-height:24px;text-align:center;
  background:var(--card2);border:1px solid var(--border);border-radius:6px;color:var(--mut);cursor:pointer;font-size:14px}
.maxbtn:hover{background:var(--acc2);color:var(--acc);border-color:var(--acc)}
.btn{border:1px solid var(--border);border-radius:6px;padding:6px 12px;font-size:12px;font-weight:600;cursor:pointer;background:var(--card2);color:var(--txt)}
.btn:hover{background:var(--acc2);color:var(--acc);border-color:var(--acc)}
.modal{display:none;position:fixed;inset:0;z-index:1000;background:rgba(15,25,40,.6);padding:12px;align-items:center;justify-content:center}
#drill-modal{z-index:2000}
.modal-box{background:#fff;border-radius:12px;box-shadow:0 12px 40px rgba(0,0,0,.35);width:100%;max-width:1400px;height:96vh;display:flex;flex-direction:column;overflow:hidden}
.modal-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:10px 16px;background:var(--head);color:#fff}
.modal-head h3{margin:0;font-size:15px;color:#fff;padding:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.modal-hint{font-size:12px;color:var(--mut);padding:7px 16px;background:var(--card2);border-bottom:1px solid var(--border)}
.modal-body{flex:1 1 auto;min-height:0;padding:14px 16px 16px;display:flex;flex-direction:column}
.modal-body canvas{width:100%!important;height:100%!important;flex:1 1 auto;min-height:0}
.drill-tools{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.drill-body{flex:1 1 auto;min-height:0;overflow:auto;padding:12px 16px 16px}
.dtwrap{overflow:auto;max-height:100%}
table.dt{border-collapse:collapse;font-size:12px;white-space:nowrap}
table.dt th,table.dt td{border:1px solid var(--border);padding:5px 8px;text-align:left}
table.dt th{background:var(--head);color:#fff;position:sticky;top:0;font-weight:600}
table.dt tr:nth-child(even) td{background:var(--card2)}
table.dt tr.team-row td{cursor:pointer}
table.dt tr.team-row:hover td{background:var(--acc2);color:var(--acc)}
table.dt tr.team-row td:nth-child(1){font-weight:600}
.foot{color:var(--mut);font-size:11px;text-align:center;padding:14px}
@media(max-width:760px){.grid{grid-template-columns:1fr}}
.inst{max-width:1000px;margin:0 auto;line-height:1.55;font-size:14px}
.inst h2{color:var(--acc);border-bottom:2px solid var(--border);padding-bottom:6px;margin:30px 0 10px}
.inst h3{color:var(--txt);margin:20px 0 6px}
.inst h4{color:#0a4bbf;margin:14px 0 4px}
.inst p{margin:6px 0;color:var(--txt)}
.inst ul,.inst ol{margin:6px 0 8px 20px}
.inst li{margin:4px 0;color:var(--txt)}
.inst code{background:var(--card2);border:1px solid var(--border);border-radius:4px;padding:1px 6px;font-size:12.5px;color:#9a3b00;white-space:nowrap}
.inst table{border-collapse:collapse;width:100%;margin:10px 0}
.inst th,.inst td{border:1px solid var(--border);padding:7px 10px;text-align:left;vertical-align:top}
.inst th{background:var(--card2);color:var(--acc)}
.inst tr:nth-child(even) td{background:rgba(255,255,255,.02)}
.inst .note{background:var(--card2);border-left:3px solid var(--acc);padding:8px 12px;border-radius:0 6px 6px 0;margin:10px 0;color:var(--txt)}
.inst .badge{display:inline-block;background:var(--card2);border:1px solid var(--border);border-radius:10px;padding:1px 9px;font-size:12px;color:var(--acc);margin:1px 2px}
</style>
</head>
<body>

<header>
  <div>
    <h1>DataHub — Dashboard de Incidências</h1>
    <div class="sub" id="sub-info"></div>
  </div>
  <div class="sub" id="meta-info"></div>
</header>

<div class="tabs">
  <button class="tab active" data-tab="d1" onclick="sw('d1')">Dashboard (1 polo)</button>
  <button class="tab" data-tab="d2" onclick="sw('d2')">Dashboard 2 (todos os polos)</button>
  <button class="tab" data-tab="d3" onclick="sw('d3')">Dashboard (Mês/Dia)</button>
  <button class="tab" data-tab="inst" onclick="sw('inst')">Regras</button>
</div>

<!-- ============================= ABA 1 ============================= -->
<div class="panel active" id="d1">
  <div class="segbar" id="segbar1">
    <button class="segbtn" data-seg="e1"><span class="dot" style="background:#1f6feb"></span>Entrada</button>
    <button class="segbtn" data-seg="s1"><span class="dot" style="background:#f0820f"></span>Saída</button>
  <button class="segbtn" data-seg="m1"><span class="dot" style="background:#5f6b7a"></span>Entrada &amp; Saída</button>
  <button class="segbtn" data-seg="bk1"><span class="dot" style="background:#466067"></span>Backlog</button>
  <button class="segbtn" data-seg="o1"><span class="dot" style="background:#1a7f37"></span>Em Aberto</button>
      <button class="segbtn" data-seg="x1"><span class="dot" style="background:#d1242f"></span>+24h</button>
      <button class="segbtn" data-seg="xO"><span class="dot" style="background:#0969da"></span>OSM</button>
      <button class="segbtn" data-seg="x2"><span class="dot" style="background:#8250df"></span>Ordem 2</button>
      <button class="segbtn" data-seg="x3"><span class="dot" style="background:#0aa2a2"></span>5 Regras</button>
  </div>

  <div class="filters" id="f-top1"></div>

  <div id="resumo1"><div class="kpis" id="k1"></div>
  <div class="topbox"><div class="tophead"><h4 id="top-title1">Top 10 Afetações (abertos)</h4><div style="display:flex;gap:8px;align-items:center;margin-left:auto"><div class="topseg" id="topseg1"><button type="button" class="tsbtn active" data-tm="afet" onclick="setTopMode('d1','afet')">Afetação</button><button type="button" class="tsbtn" data-tm="conh" onclick="setTopMode('d1','conh')">Cli x Tmp</button></div><button type="button" class="export-btn" onclick="exportTop10('d1')">⬇ Exportar Top 10</button></div></div><div id="top1"></div></div></div>

  <section class="segblock" id="seg-e1">
    <div class="seghead"><span class="dot" style="background:#1f6feb"></span>Entrada<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-e1"></div>
    <div class="grid">
      <div class="card"><h3>Entrada por Hora — Incidente × Reincidente</h3><canvas id="e1-1"></canvas></div>
      <div class="card"><h3>Entrada por Turno — Incidente × Reincidente</h3><canvas id="e1-2"></canvas></div>
      <div class="card"><h3>Análise de Reincidentes — Equipe × Operador</h3><canvas id="e1-3"></canvas></div>
      <div class="card"><h3>Primeira Causa de Reincidência (ranking)</h3><canvas id="e1-4"></canvas></div>
      <div class="card"><h3>Top Equipes com Reincidência</h3><canvas id="e1-5"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-s1">
    <div class="seghead"><span class="dot" style="background:#f0820f"></span>Saída<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-s1"></div>
    <div class="grid">
      <div class="card"><h3>Saída por Hora — Equipe × Operador (fechados)</h3><canvas id="s1-1"></canvas></div>
      <div class="card"><h3>Saída por Turno — Equipe × Operador</h3><canvas id="s1-2"></canvas></div>
      <div class="card"><h3>Saída por Hora — por categoria de Saída</h3><canvas id="s1-3"></canvas></div>
      <div class="card"><h3>Atribuição de Incidentes por Entrada (abertos)</h3><canvas id="s1-4"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-m1">
    <div class="seghead"><span class="dot" style="background:#5f6b7a"></span>Entrada &amp; Saída<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-m1"></div>
    <div class="grid">
      <div class="card"><h3>Entrada × Saída por Turno</h3><canvas id="m1-1"></canvas></div>
      <div class="card"><h3>Entrada × Saída por Hora</h3><canvas id="m1-2"></canvas></div>
      <div class="card"><h3>Entrada × Saída por Hora — Clientes</h3><canvas id="m1-3"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-bk1">
      <div class="seghead"><span class="dot" style="background:#466067"></span>Backlog<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-bk1"></div>
    <div class="grid">
      <div class="card"><h3>Backlog (dia anterior × dia atual) — em aberto por hora</h3><canvas id="cb1-1"></canvas></div>
      <div class="card"><h3>Backlog por Hora (em aberto ao final da hora)</h3><canvas id="b1-1"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-o1">
    <div class="seghead"><span class="dot" style="background:#1a7f37"></span>Em Aberto<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-o1"></div>
    <div class="grid g5">
      <div class="card"><h3>Abertas Hoje — BT</h3><canvas id="o1-1"></canvas></div>
      <div class="card"><h3>Abertas Ontem — BT</h3><canvas id="o1-2"></canvas></div>
      <div class="card"><h3>Abertas Dias Anteriores — BT</h3><canvas id="o1-3"></canvas></div>
      <div class="card span2x2"><h3>Clientes em Aberto (Clts Atual)</h3><canvas id="o1-7"></canvas></div>
      <div class="card"><h3>Abertas Hoje — MT</h3><canvas id="o1-4"></canvas></div>
      <div class="card"><h3>Abertas Ontem — MT</h3><canvas id="o1-5"></canvas></div>
      <div class="card"><h3>Abertas Dias Anteriores — MT</h3><canvas id="o1-6"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-x1">
    <div class="seghead"><span class="dot" style="background:#d1242f"></span>+24h<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-x1"></div>
    <div class="grid">
      <div class="card"><h3>+24h — Abertos × Fechados Hoje</h3><div class="stack"><div class="card vcard"><div class="v" id="x24-sim">0</div><div class="l">Abertos</div></div><div class="card vcard"><div class="v" id="x24-nao">0</div><div class="l">Fechados Hoje</div></div></div></div>
      <div class="card"><h3>+24h Ontem — Aberto × Fechado Hoje</h3><canvas id="x1b"></canvas></div>
      <div class="card"><h3>+24h Dias Anteriores — Aberto × Fechado Hoje</h3><canvas id="x1c"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-xO">
    <div class="seghead"><span class="dot" style="background:#0969da"></span>OSM<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-xO"></div>
    <div class="grid">
      <div class="card"><h3>OSM — Entradas × Saídas</h3><div class="stack"><div class="card vcard"><div class="v" id="xO-ent">0</div><div class="l">Entradas (OSM)</div></div><div class="card vcard"><div class="v" id="xO-sai">0</div><div class="l">Saídas (OSM)</div></div></div></div>
      <div class="card"><h3>Entrada × Saída por Hora — OSM (Sim)</h3><canvas id="xO-1"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-x2">
    <div class="seghead"><span class="dot" style="background:#8250df"></span>Ordem 2<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-x2"></div>
    <div class="grid">
      <div class="card"><h3>Ordem 2 — Fechados Hoje (sim × não)</h3><div class="stack"><div class="card vcard"><div class="v" id="x2-sim">0</div><div class="l">Sim</div></div><div class="card vcard"><div class="v" id="x2-nao">0</div><div class="l">Não</div></div></div></div>
      <div class="card"><h3>Ordem 2 — Fechados Hoje (sim × não)</h3><canvas id="x1-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-x3">
    <div class="seghead"><span class="dot" style="background:#0aa2a2"></span>5 Regras<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-x3"></div>
    <div class="grid">
      <div class="card"><h3>5 Regras de Ouro — Fechados Hoje MT (S × N)</h3><div class="stack"><div class="card vcard"><div class="v" id="x3-sim">0</div><div class="l">S</div></div><div class="card vcard"><div class="v" id="x3-nao">0</div><div class="l">N</div></div></div></div>
      <div class="card"><h3>5 Regras de Ouro — Fechados Hoje MT (S × N)</h3><canvas id="x1-3"></canvas></div>
    </div>
  </section>
</div>

<!-- ============================= ABA 2 ============================= -->
<div class="panel" id="d2">
  <div class="segbar" id="segbar2">
    <button class="segbtn" data-seg="e2"><span class="dot" style="background:#1f6feb"></span>Entrada</button>
    <button class="segbtn" data-seg="s2"><span class="dot" style="background:#f0820f"></span>Saída</button>
    <button class="segbtn" data-seg="o2"><span class="dot" style="background:#1a7f37"></span>Em Aberto</button>
    <button class="segbtn" data-seg="t2"><span class="dot" style="background:#8250df"></span>TMA</button>
    <button class="segbtn" data-seg="p2"><span class="dot" style="background:#0aa2a2"></span>Processos</button>
  </div>

  <div class="filters" id="f-top2"></div>

  <div id="resumo2"><div class="kpis" id="k2"></div>
  <div class="topbox"><div class="tophead"><h4 id="top-title2">Top 10 Afetações (abertos)</h4><div style="display:flex;gap:8px;align-items:center;margin-left:auto"><div class="topseg" id="topseg2"><button type="button" class="tsbtn active" data-tm="afet" onclick="setTopMode('d2','afet')">Afetação</button><button type="button" class="tsbtn" data-tm="conh" onclick="setTopMode('d2','conh')">Cli x Tmp</button></div><button type="button" class="export-btn" onclick="exportTop10('d2')">⬇ Exportar Top 10</button></div></div><div id="top2"></div></div></div>

  <section class="segblock" id="seg-e2">
    <div class="seghead"><span class="dot" style="background:#1f6feb"></span>Entrada<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-e2"></div>
    <div class="grid">
      <div class="card"><h3>Entrada Total por Polo — Incidente × Reincidente</h3><canvas id="e2-1"></canvas></div>
      <div class="card"><h3>Entrada por Hora — por Polo</h3><canvas id="e2-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-s2">
    <div class="seghead"><span class="dot" style="background:#f0820f"></span>Saída<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-s2"></div>
    <div class="grid">
      <div class="card"><h3>Saída Total por Polo — Equipe × Operador (fechados)</h3><canvas id="s2-1"></canvas></div>
      <div class="card"><h3>Saída por Hora — por Polo (fechados)</h3><canvas id="s2-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-o2">
    <div class="seghead"><span class="dot" style="background:#1a7f37"></span>Em Aberto<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-o2"></div>
    <div class="grid">
      <div class="card"><h3>Abertas por Polo × Cluster — Atribuição</h3><canvas id="o2-1"></canvas></div>
      <div class="card"><h3>Clientes em Aberto por Polo × Cluster</h3><canvas id="o2-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-t2">
    <div class="seghead"><span class="dot" style="background:#8250df"></span>TMA<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-t2"></div>
    <div class="grid">
      <div class="card"><h3>Fechados considerados p/ TMA — por Polo</h3><canvas id="t2-1"></canvas></div>
      <div class="card"><h3>Fechados por Equipe (Equipe Desl.)</h3><canvas id="t2-2"></canvas></div>
      <div class="card"><h3>Média TMA — considerados</h3><canvas id="t2-3"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-p2">
    <div class="seghead"><span class="dot" style="background:#0aa2a2"></span>Processos<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-p2"></div>
    <div class="grid">
      <div class="card"><h3>Atribuição por Polo × Entrada</h3><canvas id="p2-1" data-drill-field="eq"></canvas></div>
      <div class="card"><h3>Atribuição +24h por Polo × Entrada</h3><canvas id="p2-2" data-drill-field="eq"></canvas></div>
      <div class="card"><h3>Atribuído por Processo — Abertos</h3><canvas id="p2-3" data-drill-field="eq"></canvas></div>
      <div class="card"><h3>Improdutivo por Processo — Fechados</h3><canvas id="p2-4" data-drill-field="eqd"></canvas></div>
      <div class="card"><h3>Fechados por Processo</h3><canvas id="p2-5" data-drill-field="eqd"></canvas></div>
    </div>
  </section>
</div>

<!-- ============================= ABA INSTRUÇÕES ============================= -->
  <!-- ============================= ABA 3 (Mês/Dia) ============================= -->
  <div class="panel" id="d3">
    <div class="segbar" id="segbar3">
      <button class="segbtn" data-seg="e3"><span class="dot" style="background:#1f6feb"></span>Entrada</button>
      <button class="segbtn" data-seg="s3"><span class="dot" style="background:#f0820f"></span>Saída</button>
  <button class="segbtn" data-seg="m3"><span class="dot" style="background:#5f6b7a"></span>Entrada &amp; Saída</button>
  <button class="segbtn" data-seg="bk3"><span class="dot" style="background:#466067"></span>Backlog</button>
  <button class="segbtn" data-seg="t3"><span class="dot" style="background:#b08a00"></span>TMA</button>
      <button class="segbtn" data-seg="o3"><span class="dot" style="background:#1a7f37"></span>Reincidentes</button>
      <button class="segbtn" data-seg="o4"><span class="dot" style="background:#db2777"></span>Improdutivo</button>
      <button class="segbtn" data-seg="x31"><span class="dot" style="background:#d1242f"></span>+24h</button>
      <button class="segbtn" data-seg="xO3"><span class="dot" style="background:#0969da"></span>OSM</button>
      <button class="segbtn" data-seg="x32"><span class="dot" style="background:#8250df"></span>Ordem 2</button>
      <button class="segbtn" data-seg="x33"><span class="dot" style="background:#0aa2a2"></span>5 Regras</button>
    </div>

    <div class="filters" id="f-top3"></div>

    <div id="resumo3"><div class="kpis" id="k3"></div>
    <div class="topbox"><div class="tophead"><h4 id="top-title3">Top 10 Afetações (fechados)</h4><div style="display:flex;gap:8px;align-items:center;margin-left:auto"><div class="topseg" id="topseg3"><button type="button" class="tsbtn active" data-tm="afet" onclick="setTopMode('d3','afet')">Afetação</button><button type="button" class="tsbtn" data-tm="conh" onclick="setTopMode('d3','conh')">Cli x Tmp</button></div><button type="button" class="export-btn" onclick="exportTop10('d3')">⬇ Exportar Top 10</button></div></div><div id="top3"></div></div></div>

    <section class="segblock" id="seg-e3">
      <div class="seghead"><span class="dot" style="background:#1f6feb"></span>Entrada<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
      <div class="filters" id="f-e3"></div>
      <div class="grid">
        <div class="card"><h3>Entradas BT por Dia — Incidente × Reincidente</h3><canvas id="d3-e1-1"></canvas></div>
        <div class="card"><h3>Entradas MT por Dia — Incidente × Reincidente</h3><canvas id="d3-e1-2"></canvas></div>
        <div class="card"><h3>Entradas por Turno (A/B/C) por Dia</h3><canvas id="d3-e1-3"></canvas></div>
        <div class="card"><h3>Entradas por Afetação (Sem / Individual / Coletivo) por Dia</h3><canvas id="d3-e1-4"></canvas></div>
      </div>
    </section>

    <section class="segblock" id="seg-s3">
      <div class="seghead"><span class="dot" style="background:#f0820f"></span>Saída<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
      <div class="filters" id="f-s3"></div>
      <div class="grid">
        <div class="card"><h3>Saída por Dia — Equipe × Operador (fechados)</h3><canvas id="d3-s3-1"></canvas></div>
        <div class="card"><h3>Produtivo × Improdutivo × Produtivo s/ afetação (equipe, por dia)</h3><canvas id="d3-s3-2"></canvas></div>
        <div class="card"><h3>Improdutivos por Turno (A/B/C) por Dia</h3><canvas id="d3-s3-3"></canvas></div>
        <div class="card"><h3>Top 10 Causa — Fechados (período)</h3><canvas id="d3-s3-4"></canvas></div>
      </div>
    </section>

    <section class="segblock" id="seg-m3">
      <div class="seghead"><span class="dot" style="background:#5f6b7a"></span>Entrada &amp; Saída<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
      <div class="filters" id="f-m3"></div>
      <div class="grid">
        <div class="card"><h3>Entrada × Saída por Dia</h3><canvas id="d3-m1-1"></canvas></div>
        <div class="card"><h3>Entrada × Saída por Dia (c/ Saída Equipe)</h3><canvas id="d3-m1-2"></canvas></div>
        <div class="card"><h3>Entrada × Saída por Dia (saída = fechados)</h3><canvas id="d3-m1-3"></canvas></div>
      </div>
    </section>

    <section class="segblock" id="seg-bk3">
    <div class="seghead"><span class="dot" style="background:#466067"></span>Backlog<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
      <div class="filters" id="f-bk3"></div>
      <div class="grid">
        <div class="card"><h3>Backlog (carreamento dia a dia) — em aberto no fim do dia</h3><canvas id="d3-m1-4"></canvas></div>
        <div class="card"><h3>Backlog (carreamento) por Hora — média do período</h3><canvas id="d3-m1-5"></canvas></div>
      </div>
    </section>

    <section class="segblock" id="seg-t3">
      <div class="seghead"><span class="dot" style="background:#b08a00"></span>TMA<span class="block-hint">clique nas barras para detalhar</span></div>
      <div class="filters" id="f-t3"></div>
      <div class="grid g-tma">
        <div class="card span-all"><h3>Tempos Médios (TMA / TMP / TMD / TME)</h3><canvas id="d3-t3-1"></canvas></div>
        <div class="card span-all"><h3>Tempos Médios (TMA / TMP / TMD / TME) por Tipo</h3><canvas id="d3-t3-5"></canvas></div>
        <div class="card"><h3>Tempos Médios - Turno A (com régua de origem)</h3><canvas id="d3-t3-2"></canvas></div>
        <div class="card"><h3>Tempos Médios - Turno B (com régua de origem)</h3><canvas id="d3-t3-3"></canvas></div>
        <div class="card"><h3>Tempos Médios - Turno C (com régua de origem)</h3><canvas id="d3-t3-4"></canvas></div>
        <div class="card span-all"><h3>Tempos Médios por hora de saída (com régua de origem)</h3><canvas id="d3-t3-7"></canvas></div>
        <div class="card span-all"><h3>Tempos Médios por Hora de entrada</h3><canvas id="d3-t3-6"></canvas></div>
      </div>
    </section>

    <section class="segblock" id="seg-o3">
      <div class="seghead"><span class="dot" style="background:#1a7f37"></span>Reincidentes<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
      <div class="filters" id="f-o3"></div>
      <div class="grid">
        <div class="card"><h3>Reincidentes (1º) por Dia — Equipe × Remoto</h3><canvas id="d3-r3-1"></canvas></div>
        <div class="card"><h3>Top 10 Primeiras Causas (1º) — Reincidência</h3><canvas id="d3-r3-2"></canvas></div>
        <div class="card"><h3>Top 10 Equipes com Reincidência (1º)</h3><canvas id="d3-r3-3"></canvas></div>
      </div>
    </section>

    <section class="segblock" id="seg-o4">
      <div class="seghead"><span class="dot" style="background:#db2777"></span>Improdutivo<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
      <div class="filters" id="f-o4"></div>
      <div class="grid">
        <div class="card"><h3>Improdutivos por Dia — Equipe × Remoto</h3><canvas id="d3-r4-1"></canvas></div>
        <div class="card"><h3>Top 10 Causas — Improdutivo</h3><canvas id="d3-r4-2"></canvas></div>
        <div class="card"><h3>Top 10 Equipes com Improdutivo</h3><canvas id="d3-r4-3"></canvas></div>
      </div>
    </section>

    <section class="segblock" id="seg-x31">
      <div class="seghead"><span class="dot" style="background:#d1242f"></span>+24h<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
      <div class="filters" id="f-x31"></div>
      <div class="grid">
        <div class="card"><h3>+24h (duração) — Fechados por tempo de atendimento</h3><div class="stack"><div class="card vcard"><div class="v" id="d3-x24-sim">0</div><div class="l">Fechados >24h</div></div><div class="card vcard"><div class="v" id="d3-x24-nao">0</div><div class="l">Fechados ≤24h</div></div></div></div>
        <div class="card"><h3>+24h por Dia — Fechados com duração >24h</h3><canvas id="d3-x1b"></canvas></div>
      </div>
    </section>

  <section class="segblock" id="seg-xO3">
    <div class="seghead"><span class="dot" style="background:#0969da"></span>OSM<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-xO3"></div>
    <div class="grid">
      <div class="card"><h3>OSM — Entradas × Saídas</h3><div class="stack"><div class="card vcard"><div class="v" id="d3-xO-ent">0</div><div class="l">Entradas (OSM)</div></div><div class="card vcard"><div class="v" id="d3-xO-sai">0</div><div class="l">Saídas (OSM)</div></div></div></div>
      <div class="card"><h3>Entrada × Saída por Dia — OSM (Sim)</h3><canvas id="d3-xO-1"></canvas></div>
    </div>
  </section>

    <section class="segblock" id="seg-x32">
      <div class="seghead"><span class="dot" style="background:#8250df"></span>Ordem 2<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
      <div class="filters" id="f-x32"></div>
      <div class="grid">
        <div class="card"><h3>Ordem 2 — Fechados Hoje (sim × não)</h3><div class="stack"><div class="card vcard"><div class="v" id="d3-x2-sim">0</div><div class="l">Sim</div></div><div class="card vcard"><div class="v" id="d3-x2-nao">0</div><div class="l">Não</div></div></div></div>
        <div class="card"><h3>Ordem 2 por Dia — Fechados (sim × não)</h3><canvas id="d3-x1-2"></canvas></div>
      </div>
    </section>

    <section class="segblock" id="seg-x33">
      <div class="seghead"><span class="dot" style="background:#0aa2a2"></span>5 Regras<span class="block-hint">clique nas barras/KPIs/Top 10 para detalhar</span></div>
      <div class="filters" id="f-x33"></div>
      <div class="grid">
        <div class="card"><h3>5 Regras de Ouro — Fechados Hoje MT (S × N)</h3><div class="stack"><div class="card vcard"><div class="v" id="d3-x3-sim">0</div><div class="l">S</div></div><div class="card vcard"><div class="v" id="d3-x3-nao">0</div><div class="l">N</div></div></div></div>
        <div class="card"><h3>5 Regras de Ouro por Dia — Fechados MT (S × N)</h3><canvas id="d3-x1-3"></canvas></div>
      </div>
    </section>

  </div>

  <div class="panel" id="inst">
  <div class="inst">
    <h2>Regras — Conceitos e lógicas dos KPIs e gráficos</h2>
    <p>Esta aba documenta o significado de cada indicador e a regra exata usada para calculá-lo no dashboard. Toda a lógica reflete o que está implementado no relatório.</p>

    <h3>1. Filtros</h3>
    <h4>1.1 Filtros universais (valem para toda a aba)</h4>
    <p>Definidos em <code>TOP_CONF</code>. Aplicam-se a todos os segmentos da aba:</p>
    <ul>
      <li><b>UT</b> — Unidade (UTS / UTN).</li>
      <li><b>Polo</b> — polo do incidente (<code>p</code>).</li>
      <li><b>Sucursal</b> — sucursal (<code>suc</code>).</li>
      <li><b>NT</b> — nível de tensão (BT, MT, AT, OSM, Outro).</li>
      <li><b>Grupo Processos (Eq. Desl.)</b> — grupo do processo pela equipe deslocada (<code>gpd</code>).</li>
      <li><b>Turno Entrada</b> (<code>ta</code>) e <b>Turno Saída</b> (<code>ts</code>) — turnos A, B ou C.</li>
    </ul>
    <p>Também há o campo <b>Buscar Equipe / Incidente</b> (procura por nº ou nome de equipe <code>eq</code>/<code>eqd</code>) e o botão <b>Exportar Resumo</b> (consolida KPIs, vCards e gráficos em CSV).</p>

    <h4>1.2 Filtros por segmento</h4>
    <p>Definidos em <code>SUB_CONF</code>. Cada segmento pode refinar ainda mais:</p>
    <table>
      <tr><th>Segmento</th><th>Filtros disponíveis</th><th>Default</th></tr>
      <tr><td>Entrada (e1/e2)</td><td>Entrada, Saída, Turno Entrada, Turno Saída</td><td>Entrada = Hoje</td></tr>
      <tr><td>Saída (s1/s2)</td><td>Entrada, Saída, Turno Entrada, Turno Saída</td><td>Saída = Hoje</td></tr>
      <tr><td>Entrada &amp; Saída (m1)</td><td>Dia (Hoje / Ontem)</td><td>Dia = Hoje</td></tr>
      <tr><td>Em Aberto (o1/o2)</td><td>o1: nenhum · o2: Entrada</td><td>—</td></tr>
      <tr><td>+24h (x1), Ordem 2 (x2), 5 Regras (x3)</td><td>nenhum</td><td>—</td></tr>
      <tr><td>TMA (t2), Processos (p2)</td><td>Entrada, Saída, Turno Entrada, Turno Fechamento</td><td>—</td></tr>
      <tr><td colspan="3" style="background:#1c2733;color:#9fb3c8;font-weight:600">Dashboard 3 (visão de período / mês-dia)</td></tr>
      <tr><td>Entrada (e3)</td><td>nenhum</td><td>—</td></tr>
      <tr><td>Saída (s3)</td><td>Turno Entrada, Turno Saída</td><td>—</td></tr>
      <tr><td>Entrada &amp; Saída (m3)</td><td>nenhum (gráficos já são por dia)</td><td>—</td></tr>
      <tr><td>Reincidentes (o3)</td><td>Turno Entrada, Turno Saída</td><td>—</td></tr>
      <tr><td>Improdutivo (o4)</td><td>Turno Entrada, Turno Saída</td><td>—</td></tr>
      <tr><td>TMA (t3)</td><td>— (recorte interno por <code>Considerar TMA = Considerar</code>)</td><td>—</td></tr>
      <tr><td>+24h (x31), OSM (xO3), Ordem 2 (x32), 5 Regras (x33)</td><td>nenhum</td><td>—</td></tr>
    </table>

    <h3>2. Conceitos de base</h3>
    <ul>
      <li><b>Entrada</b> (<code>e</code>) — momento em que o incidente <i>abriu</i>: <code>Hoje</code> / <code>Ontem</code> / <code>Dias Anteriores</code>.</li>
      <li><b>Saída</b> (<code>sd</code>) — momento em que <i>fechou</i>: <code>Hoje</code> / <code>Ontem</code> / <code>Dias Anteriores</code>.</li>
      <li><b>Turnos</b> — A, B e C (Turno de Entrada <code>ta</code> e Turno de Saída/Fechamento <code>ts</code>).</li>
      <li><b>NT (Nível de Tensão)</b> — BT, MT, AT, OSM, Outro.</li>
      <li><b>Polo / Sucursal</b> — hierarquia geográfica do incidente.</li>
      <li><b>Grupo Processos (Eq. Desl.)</b> (<code>gpd</code>) — classificação do processo conforme a equipe <i>deslocada</i> (ex.: EMERGÊNCIA, COMERCIAL, PERDAS, PODA, LINHA VIVA).</li>
    </ul>

    <h3>3. Campos e regras de negócio</h3>
    <ul>
      <li><b>Equipe Atribuída</b> (<code>eq</code>) <i>versus</i> <b>Equipe Deslocada</b> (<code>eqd</code>): <code>eq</code> é a equipe a quem o incidente foi atribuído; <code>eqd</code> é a equipe que de fato foi deslocada/executou. A origem (<code>eo</code>) indica <code>equipe</code> (executado por equipe deslocada) ou <code>operador</code> (atendimento remoto/OSM).</li>
      <li class="note">Incidentes com <code>eqd = ---</code> <b>não são contados</b> na Produção nem nos gráficos de Processos (barras de deslocamento). Uma equipe em branco significa "sem equipe deslocada".</li>
      <li><b>Incidente × Reincidência</b>: um incidente é <b>Reincidência</b> quando seu tipo (<code>rt</code>) está em <code>REIN</code> (ex.: "BT – Reincidência", "BT – Reincidência Coletiva", "MT TRONCO – Reincidência", "MT RAMAL – Reincidência"). Tudo o que não está em <code>REIN</code> é <b>Incidente</b>.</li>
      <li><b>Causas ignoradas</b> (<code>isIgnCausa</code>): não entram nos gráficos de "Em Aberto" nem em alguns de atribuição. São: <code>ATENDIMENTO REMOTO</code>, <code>PROGRAMADA MANUTENÇÃO</code>, e qualquer causa que contenha <code>INSTANT</code> ou <code>CANCELADO</code>.</li>
      <li><b>TMA</b> — Tempo Médio de Atendimento (em minutos, campo <code>tma</code>). <b>TME</b> — Tempo Médio de Execução (<code>tme</code>); a análise de TME considera apenas registros com <code>tmea = sim</code>.</li>
      <li><b>Cluster de duração</b> (<code>c</code>): faixas de 2 em 2 horas (<code>0h-2h</code> … <code>22h-24h</code>) mais a faixa <code>24h+</code>.</li>
      <li><b>Duração em minutos</b> (<code>dur</code>): tempo total em minutos (abertura → fechamento). No Dashboard 3, o segmento <b>+24h</b> usa diretamente essa coluna para contar fechados com <code>dur &gt; 1440</code> (mais de 24h), em vez do cluster.</li>
      <li><b>Ordem 2</b> (<code>o</code>): <code>sim</code> / <code>não</code>.</li>
      <li><b>5 Regras de Ouro</b> (<code>ro</code>): <code>S</code> / <code>N</code> (aplicado a incidentes <b>MT</b>).</li>
      <li><b>OSM</b> (<code>osm</code>): <code>SIM</code> / <code>NÃO</code>.</li>
      <li><b>Atribuição</b> (<code>atr</code>): <code>Atribuído</code> / <code>Não Atribuído</code>.</li>
    </ul>

    <h3>4. KPIs (Dashboard 1 e 2)</h3>
    <table>
      <tr><th>KPI</th><th>Lógica de cálculo</th></tr>
      <tr><td>Abertos Agora</td><td>Registros com status <code>Aberto</code> (todos).</td></tr>
      <tr><td>Abertos Hoje / Ontem / Dias Anteriores</td><td><code>Aberto</code> e <code>e</code> = Hoje / Ontem / Dias Anteriores.</td></tr>
      <tr><td>Não Atribuídos</td><td><code>Aberto</code> e <code>atr = Não Atribuído</code>.</td></tr>
      <tr><td>Clientes Afetados (abertos)</td><td>Soma de <code>ct</code> (Clts Atual) dos abertos.</td></tr>
      <tr><td>Abertos +24h</td><td><code>Aberto</code> e cluster <code>c = 24h+</code>.</td></tr>
      <tr><td>Fechados Hoje / Ontem / Dias Anteriores</td><td><code>Fechado</code> e <code>sd</code> = Hoje / Ontem / Dias Anteriores.</td></tr>
      <tr><td>TMA Médio (fechados)</td><td>Média de <code>tma</code> dos fechados.</td></tr>
      <tr><td>TME Médio (fechados)</td><td>Média de <code>tme</code> dos fechados (com <code>tmea = sim</code> na apuração).</td></tr>
      <tr><td>Equipes Atribuídas</td><td>Nº de valores distintos de <code>eq</code> entre os abertos (ignora vazio/---).</td></tr>
      <tr><td>Equipes em Execução</td><td>Nº de valores distintos de <code>eqd</code> entre os abertos (ignora vazio/---).</td></tr>
      <tr><td>Equipes Totais Deslocadas</td><td>Nº de valores distintos de <code>eqd</code> entre os fechados hoje (ignora vazio/---).</td></tr>
    </table>

    <h4>4.1 Produção</h4>
    <p>Fórmula: <code>fechados hoje por equipe ÷ nº de equipes deslocadas</code>. O numerador considera fechados hoje com origem <code>equipe</code> e <code>eqd ≠ ---</code>; o denominador é o nº de <code>eqd</code> distintos.</p>
    <ul>
      <li><b>Dashboard 1</b>: Produção Total, Produção por Turno (A/B/C, filtrado por turno de saída <code>ts</code>) e Produção por Grupo (EMERGÊNCIA, COMERCIAL, PERDAS, PODA, LINHA VIVA, filtrado por <code>gpd</code>).</li>
      <li><b>Dashboard 2</b>: Produção por Polo (fechados hoje por equipe daquele polo ÷ equipes deslocadas daquele polo).</li>
    </ul>

    <h3>5. Segmentos e gráficos</h3>

    <h4>5.1 Dashboard 1</h4>
    <p><span class="badge">Entrada (e1)</span></p>
    <ul>
      <li>Entrada por Hora — Incidente × Reincidente: barras por hora de entrada (<code>he</code>), divididas em Incidente e Reincidente.</li>
      <li>Entrada por Turno — Incidente × Reincidente: por turno de entrada (<code>ta</code>).</li>
      <li>Análise de Reincidentes — Equipe × Operador: reincidências por tipo (<code>rt</code>), separando execução por <code>equipe</code> vs <code>operador</code>.</li>
      <li>Primeira Causa de Reincidência (ranking): top 15 causas (<code>ca</code>) das reincidências.</li>
      <li>Top Equipes com Reincidência: top 10 equipes deslocadas (<code>eqd</code>) com reincidência, por tipo.</li>
    </ul>
    <p><span class="badge">Saída (s1)</span></p>
    <ul>
      <li>Saída por Hora — Equipe × Operador (fechados): por hora de saída (<code>hs</code>).</li>
      <li>Saída por Turno — Equipe × Operador: por turno de saída (<code>ts</code>).</li>
      <li>Saída por Hora — por categoria de Saída: linha por Entrada (<code>e</code>) ao longo das horas.</li>
      <li>Atribuição de Incidentes por Entrada (abertos): por Entrada (<code>e</code>), Atribuído vs Não Atribuído, ignorando causas ignoradas.</li>
    </ul>
    <p><span class="badge">Entrada &amp; Saída (m1)</span></p>
    <ul>
      <li>por Turno: Entrada (<code>ta</code>) vs Saída (<code>ts</code>).</li>
      <li>por Hora (linha): Entrada, Saída e Saída Equipe (por <code>hs</code>).</li>
      <li>por Hora — Clientes (linha): igual, somando <code>ct</code>.</li>
      <li>por Hora — OSM (linha): Entrada OSM (<code>osm=SIM</code>) e Saída OSM.</li>
    </ul>
    <p><span class="badge">Em Aberto (o1)</span></p>
    <ul>
      <li>Abertas Hoje / Ontem / Dias Anteriores — BT e — MT: por cluster (<code>c</code>), Atribuído vs Não Atribuído, ignorando causas ignoradas.</li>
      <li>Clientes em Aberto (Clts Atual): por cluster, somando <code>ct</code>.</li>
    </ul>
    <p><span class="badge">+24h (x1)</span></p>
    <ul>
      <li>KPIs +24h (Abertos / Fechados Hoje) e gráficos para Ontem e Dias Anteriores: por tipo (BT/MT) × turno de entrada (<code>ta</code>), Aberto vs Fechado Hoje, referentes a <code>c = 24h+</code>.</li>
    </ul>
    <p><span class="badge">Ordem 2 (x2)</span></p>
    <ul>
      <li>KPIs Ordem 2 (sim/não) e gráfico: por Entrada (<code>e</code>), <code>sim</code> vs <code>não</code> (<code>o</code>), fechados hoje.</li>
    </ul>
    <p><span class="badge">5 Regras (x3)</span></p>
    <ul>
      <li>KPIs 5 Regras (S/N) e gráfico: MT com <code>ro</code> S/N, por Entrada × NT.</li>
    </ul>

    <h4>5.2 Dashboard 2 (todos os polos)</h4>
    <p><span class="badge">Entrada (e2)</span></p>
    <ul>
      <li>Entrada por Polo — Incidente × Reincidente; Entrada por Hora (linha por polo).</li>
    </ul>
    <p><span class="badge">Saída (s2)</span></p>
    <ul>
      <li>Saída por Polo — Equipe × Operador; Saída por Hora (linha por polo).</li>
    </ul>
    <p><span class="badge">Em Aberto (o2)</span></p>
    <ul>
      <li>por Polo × Cluster (Atribuído/Não Atribuído) e por Polo × Cluster (clientes <code>ct</code>), ignorando causas ignoradas.</li>
    </ul>
    <p><span class="badge">TMA (t2)</span></p>
    <ul>
      <li>Qtd de fechados "Considerar" por polo; Top 20 equipes deslocadas (<code>eqd</code>) por quantidade; Média TMA por polo (fechados "Considerar").</li>
    </ul>
    <p><span class="badge">Processos (p2)</span></p>
    <ul>
      <li>Atribuição por Polo × Entrada e Atribuição +24h por Polo × Entrada: abertos, Atribuído vs Não Atribuído.</li>
      <li>Atribuído por Processo — Abertos: por grupo de processo (<code>gp</code>, equipe atribuída).</li>
      <li>Improdutivo por Processo — Fechados e Fechados por Processo: por grupo (<code>gpd</code>, equipe deslocada <code>eqd ≠ ---</code>).</li>
    </ul>

    <h4>5.3 Dashboard 3 (Mês/Dia — visão de período)</h4>
    <p>Todos os gráficos do Dashboard 3 são <b>por dia</b> (eixo X = dia, ordenado por mês/dia) e trazem <b>linha de média</b> (vermelha para o total, ou na cor de cada série). Clique em qualquer barra/KPI para detalhar. Os filtros universais (UT, Polo, NT, Grupo, Turnos, Mês, Dia) aplicam-se normalmente; os filtros por segmento só existem onde indicado abaixo.</p>
    <p><span class="badge">Entrada (e3)</span></p>
    <ul>
      <li>Entrada BT / MT por Dia — Incidente × Reincidente (barras empilhadas) + média do total.</li>
      <li>Entrada por Turno (A/B/C) por Dia — barras lado a lado + média por turno.</li>
      <li>Entrada por Afetação (Cli Af 3 min) por Dia — Sem afetação / Individual / Coletivo + médias por categoria.</li>
    </ul>
    <p><span class="badge">Saída (s3)</span></p>
    <ul>
      <li>Saída Equipe × Operador por Dia (fechados) + média do total.</li>
      <li>Produtivo × Improdutivo × Produtivo s/ afetação (equipe) por Dia + médias por categoria.</li>
      <li>Improdutivos por Turno (A/B/C) por Dia + médias por turno.</li>
      <li>Top 10 Causa (fechados, equipe) — barras horizontais.</li>
    </ul>
    <p><span class="badge">Entrada &amp; Saída (m3)</span></p>
    <ul>
      <li>Entrada × Saída por Dia (barras) + médias.</li>
      <li>Entrada × Saída por Dia (linha, c/ Saída Equipe) + médias.</li>
      <li>Entrada × Saída por Dia — Cli Af 3 min (saída = fechados) + médias. Substitui o "Clientes Atual" do Dashboard 1 pela métrica de afetação de 3 minutos (<code>a3</code>).</li>
    </ul>
    <p><span class="badge">TMA (t3)</span></p>
    <ul>
      <li><b>Geral (por Dia)</b>: barras empilhadas de TMP + TMD + TME (média por dia) + 4 retas tracejadas de média (TMA vermelho, TMP azul, TMD laranja, TME verde) sobre o período. Apenas incidentes com <code>Considerar TMA = Considerar</code>.</li>
      <li><b>Por Turno (A/B/C)</b>: um gráfico empilhado (TMP + TMD + TME, média por dia) para cada turno, cada um com <b>suas próprias</b> 4 retas tracejadas de média — permite comparar as médias entre turnos.</li>
      <li><b>Layout</b>: o gráfico Geral ocupa a linha superior (largura total) e os 3 gráficos por turno ficam lado a lado na linha inferior (grade 2×3).</li>
      <li>A reta vermelha (TMA) aparece na <b>legenda como "TMA"</b>; o rótulo de <b>total</b> sobre as barras é vermelho com fonte branca (representa o TMA).</li>
      <li><b>Filtro</b>: nenhum filtro por segmento — o recorte é interno por <code>Considerar TMA = Considerar</code>.</li>
    </ul>
    <p><span class="badge">Reincidentes (o3)</span> — substitui o antigo segmento "Em Aberto".</p>
    <ul>
      <li>Reincidentes (1º) Equipe × Remoto por Dia (rt contém "1°", primeiras causas) + média do total.</li>
      <li>Top 10 Primeiras Causas (reincidência "1°") por causa — barras horizontais.</li>
      <li>Top 10 Equipes (eqd) com Reincidência (1°) — barras horizontais.</li>
    </ul>
    <p><span class="badge">Improdutivo (o4)</span></p>
    <ul>
      <li>Improdutivos por Dia — Equipe × Remoto (coluna <code>Produtivo/Improdutivo = Improdutivo</code>) + média do total.</li>
      <li>Top 10 Causas — Improdutivo — barras horizontais.</li>
      <li>Top 10 Equipes (eqd) com Improdutivo — barras horizontais.</li>
      <li>Cor tema: magenta (diferente do verde dos Reincidentes).</li>
    </ul>
    <p><span class="badge">+24h (x31)</span></p>
    <ul>
      <li>KPIs (Fechados &gt;24h / Fechados ≤24h) e gráfico por Dia: fechados com <code>dur &gt; 1440</code> min, calculado pela coluna Duração em minutos (não mais pelo cluster).</li>
    </ul>
    <p><span class="badge">OSM (xO3)</span></p>
    <ul>
      <li>KPIs (Entradas/Saídas OSM) e linha Entrada × Saída OSM por Dia + médias.</li>
    </ul>
    <p><span class="badge">Ordem 2 (x32) · 5 Regras (x33)</span></p>
    <ul>
      <li>Ordem 2: fechados (sim/não) por Dia + média do total.</li>
      <li>5 Regras: MT fechados (S/N) por Dia + média do total.</li>
    </ul>

    <h3>6. Interações</h3>
    <ul>
      <li><b>Clique em barra de gráfico / KPI / Top 10 / linha de equipe</b> → abre o <i>drill-down</i> (tabela de incidentes) já filtrado.</li>
      <li><b>Botão "Ver Equipes"</b> (no modal): alterna para a visão agrupada por equipe (campo <code>eq</code> ou <code>eqd</code> conforme o contexto). Clicar numa equipe filtra os incidentes dela.</li>
      <li><b>KPIs de Produção e de Equipe</b> já abrem com o botão "Ver Equipes" disponível.</li>
      <li><b>Botão Print</b> (em cada segmento e na barra de filtros): gera a imagem (html2canvas) dos gráficos e copia para a <i>área de transferência</i> (Ctrl+V para colar). Ao clicar, o botão fica verde confirmando o print.</li>
      <li><b>Botão Ent+Saída</b> (no segmento Entrada): copia os gráficos de Entrada e Saída juntos.</li>
      <li><b>Maximizar gráfico</b>: duplo clique ou botão de zoom — scroll amplia, arrastar move, +/−/⟲ controlam o zoom.</li>
      <li><b>Exportar</b>: <code>.xlsx</code>/<code>.csv</code> no modal de incidentes; <code>Exportar Resumo</code> e <code>Exportar Top 10</code> nos filtros.</li>
    </ul>
  </div>
</div>

<div class="foot">Gerado automaticamente pelo DataHub</div>

<button id="to-top" title="Voltar ao topo" onclick="window.scrollTo({top:0,behavior:'smooth'})">↑</button>

<!-- modal maximizar gráfico -->
<div class="modal" id="modal" onclick="if(event.target===this)closeModal()">
  <div class="modal-box">
    <div class="modal-head">
      <h3 id="modal-title"></h3>
      <div style="display:flex;gap:8px">
        <button class="maxbtn" title="Aproximar" style="position:static;color:#fff;border-color:#3b4a5c;background:transparent;font-size:16px;line-height:20px" onclick="modalZoomBtn(true)">+</button>
        <button class="maxbtn" title="Afastar" style="position:static;color:#fff;border-color:#3b4a5c;background:transparent;font-size:16px;line-height:20px" onclick="modalZoomBtn(false)">−</button>
        <button class="maxbtn" id="modal-reset" title="Resetar zoom" style="position:static;color:#fff;border-color:#3b4a5c;background:transparent" onclick="modalReset()">⟲</button>
        <button class="maxbtn" title="Fechar" style="position:static;color:#fff;border-color:#3b4a5c;background:transparent" onclick="closeModal()">✕</button>
      </div>
    </div>
    <div class="modal-hint">🖱️ Scroll: zoom&nbsp;&nbsp;•&nbsp;&nbsp;Arrastar: mover o gráfico&nbsp;&nbsp;•&nbsp;&nbsp;+ / − / ⟲: controle de zoom&nbsp;&nbsp;•&nbsp;&nbsp;Clique: detalhar</div>
    <div class="modal-body"><canvas id="modal-canvas"></canvas></div>
  </div>
</div>

<!-- modal drill-down (tabela de incidentes) -->
<div class="modal" id="drill-modal" onclick="if(event.target===this)closeDrillModal()">
  <div class="modal-box">
    <div class="modal-head">
      <h3 id="drill-title"></h3>
      <div class="drill-tools">
        <button class="btn" id="drill-toggle" style="display:none;color:#fff;border-color:#3b4a5c;background:transparent" onclick="toggleDrillMode()">Ver Equipes</button>
        <button class="btn" style="color:#fff;border-color:#3b4a5c;background:transparent" onclick="exportXLSX()">Exportar .xlsx</button>
        <button class="btn" style="color:#fff;border-color:#3b4a5c;background:transparent" onclick="exportCSV()">Exportar .csv</button>
        <button class="btn" style="color:#fff;border-color:#3b4a5c;background:transparent" onclick="closeDrillModal()">✕ Fechar</button>
      </div>
    </div>
    <div class="modal-hint" id="drill-count"></div>
    <div class="drill-body" id="drill-body"></div>
  </div>
</div>

<script>
Chart.register(ChartDataLabels);
  const DATA=__DADOS__;
const RAW=DATA.m;
RAW.forEach(r=>{
  const bandH=h=>{const x=parseInt(h,10); if(!Number.isFinite(x)) return null; return x<8?'A':(x<16?'B':'C');};
  const ta=bandH(r.he), ts=bandH(r.hs);
  if(ta) r.ta=ta;
  if(ts) r.ts=ts;
});
const FULL=DATA.f;
const FULL_COLS=DATA.c;
const IDXOF=new Map(RAW.map((r,i)=>[r,i]));

function esc(s){return (s==null?'':String(s)).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}

// ---- utilidades ----
function sumKey(rows, key){let s=0;for(const r of rows){s+=Number(r[key])||0;}return s;}
function meanKey(rows, key){if(!rows.length)return 0;return Math.round(sumKey(rows,key)/rows.length);}
function horasArr(rows,key,wt){
  const a=new Array(24).fill(0);
  for(const r of rows){const h=r[key];if(h===''||h==null)continue;const i=parseInt(h,10);if(i>=0&&i<24)a[i]+=(wt&&Number(r[wt])>0)?Number(r[wt]):1;}
  return a;
}
function tally(rows,keys,wt){
  const m={};
  for(const r of rows){const lab=keys.map(k=>r[k]).join(' · ');m[lab]=(m[lab]||0)+(wt?(Number(r[wt])||0):1);}
  return m;
}

const HOURS=Array.from({length:24},(_,i)=>String(i));
const HOUR_LAB=Array.from({length:24},(_,i)=>i+'h');
const CLUSTER=['0h-2h','2h-4h','4h-6h','6h-8h','8h-10h','10h-12h','12h-14h','14h-16h','16h-18h','18h-20h','20h-22h','22h-24h','24h+'];
const ENTR_ORD=['Hoje','Ontem','Dias Anteriores'];
const SAID_ORD=['Hoje','Ontem','Dias Anteriores','Aberto'];
const TURN=['A','B','C'];
const TENS_ORD=['BT','MT','AT','OSM','Outro'];
const NT_ORD=['BT','MT RAMAL','MT TRONCO','AT'];
const POLO_ORD=['MAGÉ','NITERÓI','SÃO GONÇALO','SERRANA','SUL','CAMPOS','LAGOS','MACAÉ','NOROESTE'];
const UTS=['MAGÉ','NITERÓI','SÃO GONÇALO','SERRANA','SUL'];
const UTN=['CAMPOS','LAGOS','MACAÉ','NOROESTE'];
const REIN=['BT – Reincidência','BT – Reincidência Coletiva','MT TRONCO – Reincidência','MT RAMAL – Reincidência'];
const isReinc=r=>REIN.includes(r.rt);
const isIncid=r=>!REIN.includes(r.rt);
const isPrim=r=>/1[°º]/.test(String(r.rt||''));   // primeira causa (Reincidente tipo contém "1°")
const SEM_EQUIPE=['---','-','','SEM EQUIPE','GENERICA'];
var okEqd=r=>{const v=String(r.eqd||'').trim();return v!==''&&!SEM_EQUIPE.includes(v);};
const dayKeyOf=r=>String(r.d)+'/'+String(r.mes);
function mkDays(base){
  const dk=r=>((r.mes!==undefined&&r.mes!==null)?String(r.mes):'?')+'-'+((r.d!==undefined&&r.d!==null)?String(r.d):'?');
  const dm={};
  base.forEach(r=>{const k=dk(r); if(!dm[k]) dm[k]={mes:r.mes,d:r.d};});
  const keys=Object.keys(dm).sort((a,b)=>{const A=dm[a],B=dm[b];return (Number(A.mes)-Number(B.mes))||(Number(A.d)-Number(B.d));});
  const labels=keys.map(k=>String(dm[k].d)+'/'+String(dm[k].mes));
  if(!labels.length) labels.push('—');
  return {dk,keys,labels,mean:a=>{const t=a.reduce((x,y)=>x+y,0);return a.length?t/a.length:0;}};
}
const isIgnCausa=r=>{const c=String(r.ca||'').toUpperCase();return c==='ATENDIMENTO REMOTO'||c==='PROGRAMADA MANUTENÇÃO'||c.includes('INSTANT')||c.includes('CANCELADO');};

const BLUE=['#1f6feb','#86b7f2','#4d8ceb','#b0cff7','#2f6fce','#a9cdf5','#5b8ad4','#c9dcf7'];
const ORANGE=['#f0820f','#f5b041','#f9c979','#d96c06','#f0a94a','#fcd9a0'];
const GREEN=['#1a7f37','#8fcba4','#3c9d5a','#b5dcbf','#0f6b2e','#d2ecda'];
const PALETTE=['#1f6feb','#f0820f','#1a7f37','#8250df','#d1242f','#0aa2a2','#bf3f7f','#b08a00','#5f6b7a','#00a86b','#e07856','#7b2ff7'];

function lum(hex){
  const n=parseInt(String(hex).replace('#',''),16);
  if(!isFinite(n)) return 128;
  return .2126*(n>>16)+.7152*((n>>8)&255)+.0722*(n&255);
}

function polos(){
  const s=new Set(RAW.map(r=>r.p));
  return POLO_ORD.filter(p=>s.has(p)).concat([...s].filter(p=>!POLO_ORD.includes(p)));
}
function distinct(key,order){
  const s=new Set(RAW.map(r=>r[key]).filter(v=>v&&v!==''));
  if(order) return order.filter(v=>s.has(v)).concat([...s].filter(v=>!order.includes(v)));
  return [...s].sort((a,b)=>String(a).localeCompare(String(b)));
}
function ents(){return distinct('e',ENTR_ORD);}
function sds(){return distinct('sd',SAID_ORD);}
function nts(){return distinct('nt',NT_ORD);}
function metaDias(){
  const s=new Set(RAW.map(r=>r.d));
  const ds=[...s].sort((a,b)=>Number(a)-Number(b));
  return ds.length?('Dia '+ds[0]+' a '+ds[ds.length-1]):'—';
}

// ---- navegação ----
let tab='d1';
  const BUILD_AT='__BUILD_AT__';

// ---- filtros universais (UT, Polo, NT) ----
const UF={d1:{},d2:{}};
function passes(r,F){
  if(F.p&&!F.p.all&&!F.p.sel.has(r.p)) return false;
  if(F.suc&&!F.suc.all&&!F.suc.sel.has(r.suc)) return false;
  if(F.e&&!F.e.all&&!F.e.sel.has(entKey(r))) return false;
  if(F.sd&&!F.sd.all&&!F.sd.sel.has(saiKey(r))) return false;
  if(F.ta&&!F.ta.all&&!F.ta.sel.has(r.ta)) return false;
  if(F.ts&&!F.ts.all&&!F.ts.sel.has(r.ts)) return false;
  if(F.nt&&!F.nt.all&&!F.nt.sel.has(r.nt)) return false;
  if(F.gpd&&!F.gpd.all&&!F.gpd.sel.has(r.gpd)) return false;
  if(F.mes&&!F.mes.all&&!F.mes.sel.has(String(r.mes))) return false;
  if(F.diaMes&&!F.diaMes.all&&!F.diaMes.sel.has(String(r.diaMes))) return false;
  if(F.ut){
    if(!F.ut.all){
      const a=F.ut.sel.has('UTS')&&UTS.includes(r.p);
      const b=F.ut.sel.has('UTN')&&UTN.includes(r.p);
      if(!a&&!b) return false;
    }
  }
  return true;
}

const TOP_CONF={
  d1:[['ut','UT','ut',['UTS','UTN']],['p','Polo','p',polos()],['suc','Sucursal','suc',()=>sucPool()],['nt','NT','nt',nts()],['gpd','Grupo Processos (Eq. Desl.)','gpd',()=>distinct('gpd')],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],
  d2:[['ut','UT','ut',['UTS','UTN']],['p','Polo','p',polos()],['suc','Sucursal','suc',()=>sucPool()],['nt','NT','nt',nts()],['gpd','Grupo Processos (Eq. Desl.)','gpd',()=>distinct('gpd')],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],
  d3:[['ut','UT','ut',['UTS','UTN']],['p','Polo','p',polos()],['suc','Sucursal','suc',()=>sucPool()],['nt','NT','nt',nts()],['gpd','Grupo Processos (Eq. Desl.)','gpd',()=>distinct('gpd')],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN],['mes','Mês','mes',distinct('mes').map(String)],['diaMes','Dia','diaMes',distinct('diaMes').map(String)]],
};

// ---- filtros multi-seleção ----
let clickHandler=false;
function closeMsels(){document.querySelectorAll('.msel-pop').forEach(p=>p.hidden=true);}
function mselWidget(container, label, cur, pool, onChange){
  const wrap=document.createElement('div'); wrap.className='msel';
  const fl=document.createElement('div'); fl.className='flabel'; fl.textContent=label;
  const btn=document.createElement('button'); btn.type='button'; btn.className='msel-btn';
  const pop=document.createElement('div'); pop.className='msel-pop'; pop.hidden=true;
  const render=()=>{
    btn.textContent=(cur.all||cur.sel.size===0)?'Todos':[...cur.sel].join(', ');
    btn.title=btn.textContent;
  };
  const allL=document.createElement('label'); allL.className='msel-opt all';
  const allCb=document.createElement('input'); allCb.type='checkbox';
  allL.appendChild(allCb); allL.appendChild(document.createTextNode(' Todos'));
  pop.appendChild(allL);
  const boxes=[];
  pool.forEach(v=>{
    const l=document.createElement('label'); l.className='msel-opt';
    const cb=document.createElement('input'); cb.type='checkbox';
    l.appendChild(cb); l.appendChild(document.createTextNode(' '+v));
    boxes.push({v,cb}); pop.appendChild(l);
  });
  const sync=()=>{
    allCb.checked=cur.all;
    boxes.forEach(b=>b.cb.checked=cur.sel.has(b.v));
    render();
  };
  allCb.onchange=()=>{
    if(allCb.checked){cur.all=true;cur.sel.clear();}else{cur.all=false;}
    sync(); onChange();
  };
  boxes.forEach(b=>{
    b.cb.onchange=()=>{
      if(b.cb.checked){cur.sel.add(b.v);}else{cur.sel.delete(b.v);}
      cur.all=cur.sel.size===0;
      sync(); onChange();
    };
  });
  btn.onclick=(e)=>{
    e.stopPropagation();
    const show=pop.hidden;
    closeMsels();
    pop.hidden=!show;
  };
  wrap.appendChild(fl); wrap.appendChild(btn); wrap.appendChild(pop);
  container.appendChild(wrap);
  if(!clickHandler){
    document.addEventListener('click',e=>{if(!e.target.closest('.msel'))closeMsels();});
    clickHandler=true;
  }
  sync();
  return wrap;
}
function makeFilter(cur, pool){
  return {all:!(cur&&cur.length), sel:new Set(cur||[])};
}

// ---- filtros universais por aba ----
function buildTopFilters(t){
  const cont=document.getElementById(t==='d1'?'f-top1':(t==='d2'?'f-top2':'f-top3'));
  cont.innerHTML='';
  const F={};
  TOP_CONF[t].forEach(([k,label,prop,pool])=>{
    const p=typeof pool==='function'?pool():pool;
    const cur=makeFilter(undefined,p);
    F[prop]=cur;
    mselWidget(cont,label,cur,p,()=>onTopChange(t));
  });
  UF[t]=F;
  const sw=document.createElement('div'); sw.className='flabel';
  const swlab=document.createElement('span'); swlab.textContent='Buscar Equipe / Incidente';
  sw.appendChild(swlab);
  const swr=document.createElement('div'); swr.style.cssText='display:flex;gap:6px';
  const inp=document.createElement('input'); inp.className='search-inp'; inp.placeholder='Nº, equipe...';
  const btn=document.createElement('button'); btn.type='button'; btn.className='search-btn'; btn.title='Buscar'; btn.innerHTML='&#128269;';
  const go=()=>{const q=inp.value.trim(); if(q) searchOpen(q);};
  btn.onclick=go;
  inp.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();go();}});
  swr.appendChild(inp); swr.appendChild(btn);
  sw.appendChild(swr);
  cont.appendChild(sw);
  const b=document.createElement('button');
  b.type='button'; b.className='export-btn';
  b.title='Exportar arquivo com todos os números consolidados dos KPIs e gráficos';
  b.textContent='⬇ Exportar Resumo';
  b.onclick=()=>exportResumo(t);
  cont.appendChild(b);
}
function onTopChange(t){
  Object.keys(SUB_CONF).forEach(sub=>buildFilters(sub));
  renderPage();
}

// ---- filtros por segmento ----
function entKey(r){const d=parseInt(r.d,10),m=parseInt(r.mes,10);return (Number.isFinite(d)&&Number.isFinite(m))?(d+'/'+m):'';}
function saiKey(r){const d=parseFloat(r.d_sd),m=parseFloat(r.mes_sd);return (Number.isFinite(d)&&Number.isFinite(m))?(Math.trunc(d)+'/'+Math.trunc(m)):'';}
function dateVal(s){const p=String(s).split('/');return (+p[1])*100+(+p[0]);}
function prevDayKey(s){const p=String(s).split('/');const dt=new Date(2025,(+p[1])-1,+p[0]);dt.setDate(dt.getDate()-1);return dt.getDate()+'/'+(dt.getMonth()+1);}
function lastHourOfDate(base,dk){let mx=-1;base.filter(r=>entKey(r)===dk).forEach(r=>{[r.he,r.hs].forEach(v=>{const n=parseInt(v,10);if(Number.isFinite(n)&&n>mx)mx=n;});});return mx<0?23:mx;}
const DATE_POOL_E=[...new Set(RAW.map(entKey))].filter(v=>v).sort((a,b)=>dateVal(b)-dateVal(a));
const DATE_POOL_S=[...new Set(RAW.map(saiKey))].filter(v=>v).sort((a,b)=>dateVal(b)-dateVal(a));
const DATE_POOL=[...new Set([...DATE_POOL_E,...DATE_POOL_S])].sort((a,b)=>dateVal(b)-dateVal(a));
function mostEnt(){return DATE_POOL_E[0]||'';}
function mostSai(){return DATE_POOL_S[0]||'';}
function mostDay(){return DATE_POOL[0]||'';}
const SUB_CONF={
  e1:{keys:[['e','Dia Entrada','e',DATE_POOL_E],['sd','Dia Saída','sd',DATE_POOL_S],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{e:[mostEnt()]}},
  s1:{keys:[['e','Dia Entrada','e',DATE_POOL_E],['sd','Dia Saída','sd',DATE_POOL_S],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{sd:[mostSai()]}},
  m1:{keys:[['dia','Dia','dia',DATE_POOL]],defs:{dia:[mostDay()]}},
  bk1:{keys:[['dia','Dia','dia',DATE_POOL]],defs:{dia:[mostDay()]}},
  o1:{keys:[],defs:{}},
  x1:{keys:[],defs:{}},
  xO:{keys:[['dia','Dia','dia',DATE_POOL]],defs:{dia:[mostDay()]}},
  x2:{keys:[],defs:{}},
  x3:{keys:[],defs:{}},
  e2:{keys:[['e','Dia Entrada','e',DATE_POOL_E],['sd','Dia Saída','sd',DATE_POOL_S],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{e:[mostEnt()]}},
  s2:{keys:[['e','Dia Entrada','e',DATE_POOL_E],['sd','Dia Saída','sd',DATE_POOL_S],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{sd:[mostSai()]}},
  o2:{keys:[['e','Dia Entrada','e',DATE_POOL_E]],defs:{}},
  t2:{keys:[['e','Dia Entrada','e',DATE_POOL_E],['sd','Dia Saída','sd',DATE_POOL_S],['ta','Turno Entrada','ta',TURN],['ts','Turno Fechamento','ts',TURN]],defs:{}},
  p2:{keys:[['e','Dia Entrada','e',DATE_POOL_E],['sd','Dia Saída','sd',DATE_POOL_S],['ta','Turno Entrada','ta',TURN],['ts','Turno Fechamento','ts',TURN]],defs:{}},
  e3:{keys:[],defs:{}},
  s3:{keys:[['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{}},
  m3:{keys:[],defs:{}},
  bk3:{keys:[],defs:{}},
  t3:{keys:[],defs:{}},
  o3:{keys:[['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{}},
  o4:{keys:[['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{}},
  x31:{keys:[],defs:{}},
  xO3:{keys:[],defs:{}},
  x32:{keys:[],defs:{}},
  x33:{keys:[],defs:{}},
};
const FSTATE={};
let CUR_SUB='';
function sucPool(){
  return distinct('suc');
}
function buildFilters(sub){
  const conf=SUB_CONF[sub];
  const cont=document.getElementById('f-'+sub);
  if(!cont){ FSTATE[sub]={}; return; }
  if(!conf.keys.length){ cont.style.display='none'; FSTATE[sub]={}; return; }
  cont.style.display='';
  cont.innerHTML='';
  const F={};
  conf.keys.forEach(([k,label,prop,pool])=>{
    const p=typeof pool==='function'?pool():pool;
    const cur=makeFilter(conf.defs[prop],p);
    for(const v of [...cur.sel]) if(!p.includes(v)) cur.sel.delete(v);
    if(cur.sel.size===0) cur.all=true;
    F[prop]=cur;
    mselWidget(cont,label,cur,p,()=>renderSub(sub));
  });
  FSTATE[sub]=F;
}
function mergeF(a,b){
  if(!a||a.all) return b||{all:true,sel:new Set()};
  if(!b||b.all) return a;
  const sel=new Set([...a.sel].filter(v=>b.sel.has(v)));
  return {all:false,sel:sel};
}
function segBase(sub){
  const U=UF[tab]||{};
  const F=FSTATE[sub]||{};
  const C={ut:U.ut,p:U.p,suc:U.suc,nt:U.nt,gpd:U.gpd,e:F.e,sd:F.sd,ta:mergeF(U.ta,F.ta),ts:mergeF(U.ts,F.ts),mes:U.mes,diaMes:U.diaMes};
  return RAW.filter(r=>passes(r,C));
}
function plOf(base){ return polos().filter(p=>base.some(r=>r.p===p)); }

// ---- gráficos ----
let charts={};
let lastCfg={};
function mk(id, cfg){
  if(charts[id]) charts[id].destroy();
  cfg.options = cfg.options||{};
  cfg.options.maintainAspectRatio=false;
  cfg.options.plugins = cfg.options.plugins||{};
  cfg.options.plugins.legend = Object.assign({position:'bottom',labels:{boxWidth:12,font:{size:11},color:'#5b6b7b'}}, cfg.options.plugins.legend||{});
  cfg.options.plugins.legend.onClick=function(e,item,legend){ if(item.datasetIndex==null) return; const ci=legend.chart; ci.setDatasetVisibility(item.datasetIndex, !ci.isDatasetVisible(item.datasetIndex)); ci.update(); };
  cfg.options.plugins.legend.labels.generateLabels=function(chart){
    const ds=chart.data.datasets||[];
    const idx=ds.findIndex(d=>d.label&&d.label.indexOf('Entrada A / Entrada B / Entrada C')===0);
    const out=[];
    ds.forEach((d,i)=>{
      if(i===idx) return;
      const isLine=d.type==='line';
      const col=isLine?(d.borderColor||'#888'):(Array.isArray(d.backgroundColor)?d.backgroundColor[0]:(d.backgroundColor||'#888'));
      out.push({text:d.label||('Dataset '+(i+1)),fillStyle:col,strokeStyle:d.borderColor||col,lineWidth:isLine?2:1,datasetIndex:i,hidden:!chart.isDatasetVisible(i)});
    });
    if(idx>=0) ['A','B','C'].forEach(t=>out.push({text:'Turno '+t,fillStyle:ENTRADA_COLORS[t],strokeStyle:ENTRADA_COLORS[t],lineWidth:1,datasetIndex:idx,hidden:!chart.isDatasetVisible(idx)}));
    const note=chart.options.plugins.legend.__note;
    if(note) out.push({text:note,fillStyle:'transparent',strokeStyle:'transparent',lineWidth:0,datasetIndex:null});
    return out;
  };
  cfg.options.plugins.tooltip = {callbacks:{label:c=>{const v=c.parsed.y??c.parsed.x??c.parsed;return ' '+Number(v).toLocaleString('pt-BR',{maximumFractionDigits:1});}}};
  cfg.options.layout = cfg.options.layout||{};
  cfg.options.layout.padding = Object.assign({top:26,right:40}, cfg.options.layout.padding||{});
  const el=document.getElementById(id);
  el.style.height='290px';
  el.onclick=ev=>handleChartClick(id, charts[id], ev);
  el.onmousemove=ev=>handleChartHover(id, charts[id], ev);
  el.onmouseleave=()=>{ const t=document.getElementById('ov-tip'); if(t) t.style.display='none'; if(el) el.style.cursor='default'; };
  lastCfg[id]=cloneCfg(cfg);
  charts[id]=new Chart(el, cfg);
  return charts[id];
}
function ensureOvTip(){
  let t=document.getElementById('ov-tip');
  if(!t){ t=document.createElement('div'); t.id='ov-tip'; document.body.appendChild(t); }
  t.style.position='fixed'; t.style.zIndex='9999'; t.style.pointerEvents='none';
  t.style.background='rgba(31,45,61,.96)'; t.style.color='#fff';
  t.style.padding='4px 8px'; t.style.borderRadius='6px'; t.style.fontSize='11px';
  t.style.fontFamily='"Segoe UI",Arial'; t.style.boxShadow='0 2px 8px rgba(0,0,0,.25)';
  t.style.display='none';
  return t;
}
function handleChartHover(id, chart, ev){
  if(!chart) return;
  const t=ensureOvTip();
  const cnv=chart.canvas;
  const rect=cnv.getBoundingClientRect();
  const x=(ev.clientX-rect.left);
  const y=(ev.clientY-rect.top);
  let tip='';
  if(chart.__countRects&&chart.__countRects.length){
    const hit=chart.__countRects.find(b=>x>=b.x1&&x<=b.x2&&y>=b.y1&&y<=b.y2);
    if(hit){ tip=hit.tip||''; const label=(chart.data.labels||[])[hit.i]; if(label!=null) tip+=' · '+label; }
  }
  if(tip){ t.textContent=tip; t.style.display='block'; t.style.left=(ev.clientX+12)+'px'; t.style.top=(ev.clientY+12)+'px'; cnv.style.cursor='pointer'; }
  else { t.style.display='none'; cnv.style.cursor='default'; }
}
function cloneCfg(cfg){
  return {
    type: cfg.type,
    data:{labels:(cfg.data.labels||[]).slice(),datasets:(cfg.data.datasets||[]).map(ds=>Object.assign({},ds,{data:(ds.data||[]).slice()}))},
    options:Object.assign({}, cfg.options, {plugins:Object.assign({}, cfg.options.plugins), scales:Object.assign({}, cfg.options.scales||{}), layout:Object.assign({}, cfg.options.layout)})
  };
}
function axes(opts){
  return {scales:{
    x:Object.assign({grid:{color:'#e5ebf3'},ticks:{color:'#5b6b7b',font:{size:10}}},opts&&opts.x||{}),
    y:Object.assign({beginAtZero:true,grid:{color:'#e5ebf3'},ticks:{color:'#5b6b7b'}},opts&&opts.y||{})
  }};
}
function dlabels(display, extra){
  const base={color:'#1f2d3d',font:{size:11,weight:'600'},anchor:'end',align:'end',offset:4,display:display===false?'auto':true,
    backgroundColor:'#ffffff',borderRadius:3,padding:{top:2,bottom:2,left:4,right:4},
    formatter:v=>Number.isFinite(v)?(v===Math.round(v)?String(v):v.toFixed(1)):v};
  return Object.assign(base, extra||{});
}
function dins(){
  return {color:'#ffffff',font:{size:10,weight:'700'},anchor:'center',align:'center',display:'auto',
    backgroundColor:c=>{
      const ds=c.dataset;if(!ds)return '#5b6b7b';
      if(ds.type==='line') return ds.borderColor||'#5b6b7b';
      const b=Array.isArray(ds.backgroundColor)?ds.backgroundColor[c.dataIndex]:ds.backgroundColor;
      return b&&b!=='rgba(0,0,0,0)'?b:'#5b6b7b';
    },
    borderRadius:3,padding:{top:1,bottom:1,left:3,right:3},
    formatter:v=>Number.isFinite(v)?(v===Math.round(v)?String(v):v.toFixed(1)):v};
}
function stacked(id,cfg){
  mk(id,Object.assign({},cfg,{options:Object.assign({},cfg.options||{},axes({x:{stacked:true},y:{stacked:true}}))}));
}

// plugin: total no topo (fundo branco, fonte preta) — orientação vertical/horizontal
const totalsPlugin={id:'totals',afterDatasetsDraw(chart,args,opts){
  if(!opts||!opts.enabled) return;
  const isH=chart.options.indexAxis==='y';
  const isStacked=isH?!!(chart.options.scales&&chart.options.scales.x&&chart.options.scales.x.stacked):!!(chart.options.scales&&chart.options.scales.y&&chart.options.scales.y.stacked);
  const cat=isH?chart.scales.y:chart.scales.x;
  const val=isH?chart.scales.x:chart.scales.y;
  if(!cat||!val) return;
  const labels=chart.data.labels||[];
  const meta=chart.getDatasetMeta(0);
  if(!meta||!meta.data||!meta.data.length) return;
  const ctx=chart.ctx;
  const ds=chart.data.datasets;
  const tmaIdx=ds.findIndex(d=>d.label==='TMA');
  const tmaVisible=tmaIdx<0||chart.isDatasetVisible(tmaIdx);
const n=Math.min(labels.length,meta.data.length);
  chart.__totRects=[];
  ctx.save();
  ctx.textAlign='center';
  ctx.textBaseline='middle';
  ctx.font='700 11px "Segoe UI",Arial';
  const items=[];
  for(let i=0;i<n;i++){
    let tot=0,maxV=0;
    for(const d of ds){
      if(d.type==='line'||d.order===99) continue;
      const v=Number(d.data[i])||0;tot+=v;if(v>maxV)maxV=v;}
    if(tot<=0) continue;
    const txt=tot>=1000?String(Math.round(tot/100)/10)+'k':(Number.isInteger(tot)?String(tot):tot.toFixed(1));
    const tw=ctx.measureText(txt).width;
    const px=isH?val.getPixelForValue(tot):cat.getPixelForValue(labels[i]);
    const anchorV=isStacked?tot:maxV;
    const py=isH?cat.getPixelForValue(labels[i]):val.getPixelForValue(anchorV);
    items.push({i:i,txt:txt,tw:tw,px:px,py:py});
  }
  const placed=[];
  const ok=(x1,y1,x2,y2)=>{for(const p of placed){if(x1<p.x2&&x2>p.x1&&y1<p.y2&&y2>p.y1)return false;}return true;};
  const top=chart.chartArea?chart.chartArea.top+4:0;
  for(const it of items){
    const tl=(opts&&opts.redLabel)&&tmaVisible;
    if((opts&&opts.redLabel)&&!tmaVisible) continue;
    const bh=16,step=18;
    let cy=(isH?it.py:it.py-15);
    let x1=it.px-it.tw/2-3,y1=cy-bh/2,x2=it.px+it.tw/2+3,y2=cy+bh/2;
    let guard=0;
    while(!ok(x1,y1,x2,y2)){
      cy-=step;y1=cy-bh/2;y2=cy+bh/2;
      if(cy<top||++guard>80){cy=(isH?it.py:it.py-15);y1=cy-bh/2;y2=cy+bh/2;break;}
    }
    placed.push({x1:x1,y1:y1,x2:x2,y2:y2});
    chart.__totRects.push({x1:x1,y1:y1,x2:x2,y2:y2,i:it.i});
     ctx.fillStyle= tl?'#d1242f':'#ffffff';
     ctx.fillRect(x1,y1,it.tw+6,16);
     ctx.strokeStyle= tl?'#d1242f':'#d8e0ea';
     ctx.lineWidth=1;
     ctx.strokeRect(x1,y1,it.tw+6,16);
     ctx.fillStyle= tl?'#ffffff':'#1f2d3d';
     ctx.fillText(it.txt,it.px,cy);
  }
  ctx.restore();
}};
Chart.register(totalsPlugin);
const meanLinesPlugin={id:'meanLines',afterDatasetsDraw(chart){
  const cfg=chart.options.plugins&&chart.options.plugins.meanLines;
  if(!cfg||!cfg.lines||!cfg.lines.length) return;
  const {ctx,chartArea:{left,right},scales:{y}}=chart;
  ctx.save();ctx.font='700 10px "Segoe UI",Arial';ctx.textBaseline='middle';ctx.textAlign='left';
  const visibleLines=cfg.lines.filter(ml=>{
    const v=ml.value;
    if(v==null||!isFinite(v)) return false;
    const color=ml.color||'#d1242f';
    return (chart.data.datasets||[]).some((d,i)=>chart.isDatasetVisible(i)&&(d.borderColor===color||d.backgroundColor===color));
  });
  visibleLines.forEach((ml,idx)=>{
    const v=ml.value;
    const color=ml.color||'#d1242f';
    const yp=y.getPixelForValue(v);
    if(yp<chart.chartArea.top-1||yp>chart.chartArea.bottom+1) return;
    if(!ml.noLine){
      ctx.beginPath();ctx.setLineDash([6,4]);ctx.lineWidth=1.5;ctx.strokeStyle=color;
      ctx.moveTo(left,yp);ctx.lineTo(right,yp);ctx.stroke();
    }
    let label;
    if(ml.label) label=ml.label+' : ';
    else { const dm=(chart.data.datasets||[]).find(d=>d.borderColor===color||d.backgroundColor===color); label=(dm&&dm.label?dm.label:'Média')+': '; }
    const txt=label+Number(v).toLocaleString('pt-BR',{maximumFractionDigits:1});
    const tw=ctx.measureText(txt).width;
    const yLabel=yp-idx*18-8;
    ctx.setLineDash([]);ctx.fillStyle='rgba(255,255,255,.9)';ctx.fillRect(right-tw-7,yLabel-8,tw+6,16);
    ctx.fillStyle='#1f2d3d';ctx.fillText(txt,right-tw-4,yLabel);
  });
  ctx.restore();
}};
Chart.register(meanLinesPlugin);
const countOverlayPlugin={id:'countOverlay',afterDatasetsDraw(chart){
  const datasets=(chart.data.datasets||[]).filter(d=>d.__countOverlay);
  if(!datasets.length) return;
  const y1=chart.scales&&chart.scales.y1;
  if(!y1) return;
  const {ctx,chartArea:ca}=chart;
  const x=chart.scales.x;
  ctx.save();
  chart.__countRects=[];
  const hourTop={};
  datasets.forEach(ds=>{
    const di=(chart.data.datasets||[]).indexOf(ds);
    if(chart.isDatasetVisible&&!chart.isDatasetVisible(di)) return;
    const isTeams=ds.label&&ds.label.indexOf('equipes')>=0;
    const metric=ds.label==='Média saída'?'saida':ds.label==='Média entrada'?'entrada':isTeams?'equipes':'producao';
    if(ds.noLine){
      // rótulos posicionados ACIMA do rótulo vermelho (total) no topo de cada barra, independente da escala
      const tr=chart.__totRects||[];
      ctx.font='700 10px "Segoe UI",Arial';ctx.textBaseline='middle';ctx.textAlign='center';
      ds.data.forEach((v,i)=>{
        if(v==null||!isFinite(v)||v<=0) return;
        const rect=tr.find(r=>r.i===i);
        if(!rect) return;
        const topY=(hourTop[i]!==undefined)?hourTop[i]:(rect.y1-2);
        const txt=Number(v).toLocaleString('pt-BR',{maximumFractionDigits:1});
        const tw=ctx.measureText(txt).width;
        const w=Math.min(tw+8,90),h=14;
        let by=topY-h;
        if(by<ca.top+1) by=ca.top+1;
        const cx=x.getPixelForValue(i);
        ctx.fillStyle=ds.borderColor||'#5f6b7a';
        ctx.fillRect(cx-w/2,by,w,h);
        ctx.fillStyle='#ffffff';ctx.fillText(txt,cx,by+h/2);
        chart.__countRects.push({x1:cx-w/2,y1:by,x2:cx+w/2,y2:by+h,i:i,metric:metric,tip:(ds.label||'')+': '+txt+(isTeams?' · clique p/ ver equipes':'')});
        hourTop[i]=by-2;
      });
      return;
    }
    const pts=ds.data.map((v,i)=>v!=null&&isFinite(v)?{i:i,x:x.getPixelForValue(i),y:y1.getPixelForValue(v),v:v}:null);
    const drawLine=!ds.noLine;
    if(drawLine){
      ctx.lineWidth=2;ctx.strokeStyle=ds.borderColor||'#8e44ad';
      ctx.beginPath();
      let started=false;
      for(const p of pts){ if(!p||p.v<=0) continue;
        if(!started){ctx.moveTo(p.x,p.y);started=true;}else ctx.lineTo(p.x,p.y);
      }
      if(started) ctx.stroke();
    }
    ctx.fillStyle=ds.borderColor||'#8e44ad';
    ctx.font='700 11px "Segoe UI",Arial';ctx.textBaseline='middle';ctx.textAlign='center';
    for(const p of pts){ if(!p||p.v<=0) continue;
      ctx.beginPath();ctx.arc(p.x,p.y,3,0,Math.PI*2);ctx.fill();
      const txt=Number(p.v).toLocaleString('pt-BR',{maximumFractionDigits:1});
      const tw=ctx.measureText(txt).width;
      const w=Math.min(tw+8,90),h=15;
      let bx=p.x-w/2,by=p.y-h/2-2;
      if(by<ca.top+2) by=ca.top+2;
      if(by+h>ca.bottom-2) by=ca.bottom-2-h;
      ctx.fillRect(bx,by,w,h);
      ctx.fillStyle='#ffffff';ctx.fillText(txt,p.x,by+h/2);ctx.fillStyle=ds.borderColor||'#8e44ad';
      chart.__countRects.push({x1:bx,y1:by,x2:bx+w,y2:by+h,i:p.i,metric:metric,tip:(ds.label||'')+': '+txt});
    }
  });
  ctx.restore();
}};
Chart.register(countOverlayPlugin);

// ---- plugin: composição por turno de ENTRADA (régua ao lado da barra do dia) ----
const ENTRADA_COLORS={A:'#8b5cf6',B:'#fe8da1',C:'#0aa2a2'};
const ENTRADA_LEGEND=['A','B','C'];
const ENTRADA_DRAW=['C','B','A']; // desenha de baixo p/ cima; A fica no topo
const entradaCompPlugin={id:'entradaComp',beforeDatasetsDraw(chart){
  const comp=chart.__entradaComp;
  if(!comp) return;
  const hasVisibleEntrada=(chart.data.datasets||[]).some((d,i)=>{
    if(!chart.isDatasetVisible(i)) return false;
    const lbl=d.label||'';
    return lbl==='Entrada A / Entrada B / Entrada C' || /^\/ Entrada [ABC]$/.test(lbl);
  });
  if(!hasVisibleEntrada) return;
  const {ctx,chartArea:ca}=chart;
  const meta=chart.getDatasetMeta(0);
  if(!meta||!meta.data||!meta.data.length) return;
  const n=Math.min(comp.length,meta.data.length);
  const barW=meta.data[0]?(meta.data[0].width||8):8;
  const gw=Math.max(18,Math.min(30,barW*0.75));
  const barDs=chart.data.datasets.filter(d=>!d.type||d.type==='bar');
  ctx.save();
  chart.__entradaRects=[];
  for(let i=0;i<n;i++){
    const b=meta.data[i];
    if(!b) continue;
    const sum=(barDs.reduce((s,d)=>s+(Number(d.data[i])||0),0));
    const top=yPix(sum);
    if(!(sum>0)||top>=ca.bottom-1) continue;
    const x=b.x+barW/2+3;
    const totalH=ca.bottom-top;
    let y0=ca.bottom;
    for(const t of ENTRADA_DRAW){
      const pct=comp[i][t]||0;
      const segH=totalH*(pct/100);
      const y1=y0-segH;
      ctx.fillStyle=ENTRADA_COLORS[t];
      ctx.fillRect(x,y1,gw,segH);
      ctx.strokeStyle='#ffffff';ctx.lineWidth=1;
      ctx.strokeRect(x,y1,gw,segH);
      if(segH>=14){
        const dl=chart.options.plugins&&chart.options.plugins.datalabels;
        let dsize=(dl&&dl.font&&dl.font.size)||10;
        if(chart.canvas&&chart.canvas.id==='modal-canvas') dsize=Math.round(dsize*0.8);
        ctx.font='700 '+dsize+'px "Segoe UI",Arial';
        const txt=Math.round(pct)+'%';
        const tw=ctx.measureText(txt).width;
        const ch=dsize+2;const cw=tw+6;
        const cx=x+gw/2;
        const cy=y1+segH/2;
        ctx.fillStyle=ENTRADA_COLORS[t];
        roundRectAll(ctx,cx-cw/2,cy-ch/2,cw,ch,3);
        ctx.fillStyle='#ffffff';ctx.textAlign='center';ctx.textBaseline='middle';
        ctx.fillText(txt,cx,cy);
      }
      y0=y1;
    }
    chart.__entradaRects.push({x1:x,y1:top,x2:x+gw,y2:ca.bottom,i:i});
  }
  ctx.restore();
  function yPix(sum){
    const ax=chart.scales.y;
    if(!ax) return ca.bottom;
    let p=ax.getPixelForValue(sum);
    if(p<ca.top) p=ca.top;
    if(p>ca.bottom) p=ca.bottom;
    return p;
  }
}};
Chart.register(entradaCompPlugin);

// ---- plugin: destaque do turno com maior TMA no t3-6 ----
const horaTurnoPlugin={id:'horaTurno',beforeDatasetsDraw(chart){
  const st=chart.__horaTurno;
  if(!st||!st.hours||!st.hours.length) return;
  const tmaIdx=(chart.data.datasets||[]).findIndex(d=>d.label==='TMA');
  if(tmaIdx>=0&&!chart.isDatasetVisible(tmaIdx)) return;
  const {ctx,chartArea:ca}=chart;
  const barIdx=[];chart.data.datasets.forEach((d,j)=>{if(!d.type||d.type==='bar')barIdx.push(j);});
  const bars=barIdx.map(j=>chart.getDatasetMeta(j)).filter(m=>m&&m.data&&m.data.length);
  if(!bars.length) return;
  const base=bars[0];
  const w=base.data[0]?(base.data[0].width||14):14;
  ctx.save();
  ctx.beginPath();
  for(const i of st.hours){
    if(i<0||i>=base.data.length) continue;
    const b=base.data[i];
    let top=ca.bottom;
    for(const m of bars){const e=m.data[i];if(e&&Number.isFinite(e.y))top=Math.min(top,e.y);}
    const x=b.x-w/2-2,y=top-2,ww=w+4,hh=ca.bottom-top+4;
    ctx.rect(x,y,ww,hh);
  }
  ctx.strokeStyle='#d1242f';ctx.lineWidth=1.75;
  ctx.stroke();
  ctx.restore();
},afterDatasetsDraw(chart){
  const st=chart.__horaTurno;
  if(!st||!st.hours||!st.hours.length) return;
  const tmaIdx=(chart.data.datasets||[]).findIndex(d=>d.label==='TMA');
  if(tmaIdx>=0&&!chart.isDatasetVisible(tmaIdx)) return;
  const {ctx,chartArea:ca}=chart;
  ctx.save();
  const txt=st.tlabel+' · TMA médio '+Math.round(st.max)+' min · sem o turno: '+Math.round(st.sem)+' min · impacto +'+Math.round(st.imp)+' min (média geral '+Math.round(st.geral)+')';
  ctx.font='700 10px "Segoe UI",Arial';
  const tw=ctx.measureText(txt).width;
  const bw=tw+16;
  const bx=ca.left+(ca.right-ca.left-bw)/2, by=ca.top-21;
  ctx.fillStyle='#d1242f';
  roundRectAll(ctx,bx,by,bw,19,9);
  ctx.fillStyle='#ffffff';ctx.textAlign='center';ctx.textBaseline='middle';
  ctx.fillText(txt,bx+bw/2,by+9.5);
  ctx.restore();
}};
Chart.register(horaTurnoPlugin);
function roundRectAll(ctx,x,y,w,h,r){
  ctx.beginPath();
  ctx.moveTo(x+r,y);ctx.arcTo(x+w,y,x+w,y+h,r);ctx.arcTo(x+w,y+h,x,y+h,r);
  ctx.arcTo(x,y+h,x,y,r);ctx.arcTo(x,y,x+w,y,r);ctx.closePath();
  ctx.fill();
}

// ---- renderChart genérico + drill ----
let drillSets={};
let overlayDrill={};
let overlayDrillEq={};
function storeDrill(id, labels, fn){
  drillSets[id]=[];
  for(let i=0;i<labels.length;i++) drillSets[id][i]=fn(i);
}
function chartTitle(id){
  const el=document.getElementById(id);
  const card=el&&el.closest('.card');
  return (card&&card.querySelector('h3').textContent)||id;
}
function handleChartClick(id, chart, ev){
  if(!chart) return;
  const df=chart.canvas.getAttribute('data-drill-field')||null;
  const cnv=chart.canvas;
  const rect=cnv.getBoundingClientRect();
  const x=(ev.clientX-rect.left);
  const y=(ev.clientY-rect.top);
  if(chart.__countRects&&chart.__countRects.length){
    const hit=chart.__countRects.find(b=>x>=b.x1&&x<=b.x2&&y>=b.y1&&y<=b.y2);
    if(hit){
      const label=(chart.data.labels||[])[hit.i];
      const rows=(chart.__countDrill&&chart.__countDrill[hit.metric]&&chart.__countDrill[hit.metric][hit.i])||null;
      const field=hit.metric==='equipes'?'eqd':df;
      if(rows&&rows.length) openDrillModal(chartTitle(id)+(label!=null?' — '+label:'')+(hit.metric==='equipes'?' · Equipes':''), rows, field);
      return;
    }
  }
  if(chart.__entradaRects&&chart.__entradaRects.length){
    const hit=chart.__entradaRects.find(b=>x>=b.x1&&x<=b.x2&&y>=b.y1&&y<=b.y2);
    if(hit){
      const sets=(drillSets[id]&&drillSets[id][hit.i])?drillSets[id][hit.i]:null;
      if(sets&&sets.length){
        const rows=sets.reduce((a,b)=>a.concat(b),[]);
        const label=(chart.data.labels||[])[hit.i];
        if(rows.length) openDrillModal(chartTitle(id)+(label!=null?' — '+label:''), rows, df);
      }
      return;
    }
  }
  const isLine=chart.config&&chart.config.type==='line';
  const items=chart.getElementsAtEventForMode(ev,'nearest',isLine?{intersect:false,maxDistance:40}:{intersect:true},false);
  if(items.length){
    const it=items[0];
    const label=(chart.data.labels||[])[it.index];
    const sets=(drillSets[id]&&drillSets[id][it.index])||null;
    let rows=null;
    if(sets){rows=sets[it.datasetIndex]||null;if(!rows||!rows.length)rows=sets.find(s=>s&&s.length);}
    if(rows&&rows.length){openDrillModal(chartTitle(id)+(label!=null?' — '+label:''),rows,df);return;}
  }
  if(chart.__totRects&&chart.__totRects.length){
    const hit=chart.__totRects.find(b=>x>=b.x1&&x<=b.x2&&y>=b.y1&&y<=b.y2);
    if(hit){
      const sets=(drillSets[id]&&drillSets[id][hit.i])?drillSets[id][hit.i]:null;
      if(sets&&sets.length){
        const rows=chart.__totSingle?(sets[0]||[]):sets.reduce((a,b)=>a.concat(b),[]);
        const label=(chart.data.labels||[])[hit.i];
        if(rows.length) openDrillModal(chartTitle(id)+(label!=null?' — '+label:''), rows, df);
      }
      return;
    }
  }
  const ex=chart.$datalabels&&chart.$datalabels._labels;
  if(ex){
    for(let i=ex.length-1;i>=0;i--){
      const l=ex[i];
      if(l.$layout&&l.$layout._visible&&l.$layout._box.contains({x:x,y:y})){
        const sets=(drillSets[id]&&drillSets[id][l._index])||null;
        let rows=null;
        if(sets){rows=sets[l.$groups._set]||null;if(!rows||!rows.length)rows=sets.find(s=>s&&s.length);}
        if(rows&&rows.length) openDrillModal(chartTitle(id)+' — '+(chart.data.labels||[])[l._index], rows, df);
        return;
      }
    }
  }
}
function renderChart(id, labels, series, o){
  if(tab==='d3') id='d3-'+id;
  o=o||{};
  const datasets=series.map(s=>{
    const d={label:s.name,data:labels.map(lab=>(s.val?s.val(s.rows(lab)):s.wt?sumKey(s.rows(lab),s.wt):s.rows(lab).length))};
    d.backgroundColor=s.color;
    if(o.type==='line'){
      d.borderColor=s.color; d.borderWidth=2; d.pointRadius=(o.pts===false?0:2);
      d.fill=s.fill===true; d.tension=(s.tension===undefined?0.3:s.tension);
      d.backgroundColor=s.area||'rgba(0,0,0,0)';
    }
    return d;
  });
  storeDrill(id, labels, i=>series.map(s=>s.rows(labels[i])));
  const isLine=o.type==='line';
  const plug={datalabels:isLine?dlabels(false):dins(),totals:{enabled:!!o.stacked}};
  if(o.meanLines) plug.meanLines={lines:o.meanLines};
  const cfg={type:o.type||'bar',data:{labels:labels,datasets:datasets},options:{plugins:plug}};
  if(o.indexAxis==='y') cfg.options.indexAxis='y';
  if(o.stacked){
    stacked(id,cfg);
  }else{
    if(o.axes!==false) cfg.options=Object.assign(cfg.options,axes(o.axes));
    mk(id,cfg);
  }
}

function comboKeys(rows,keys){
  const axes=keys.map(k=>k==='h'?HOURS:
    k==='p'?polos():
    k==='c'?CLUSTER:
    k==='e'?ENTR_ORD:
    k==='sd'?SAID_ORD:
    k==='t'?TENS_ORD:
    k==='ta'?TURN:k==='ts'?TURN:
    k==='nt'?NT_ORD:TURN);
  const present=new Set(rows.map(r=>keys.map(k=>r[k]).join(' · ')));
  const out=[];
  (function rec(i,acc){
    if(i===keys.length){const lab=acc.join(' · ');if(present.has(lab))out.push(lab);return;}
    for(const v of axes[i]) rec(i+1,acc.concat(v));
  })(0,[]);
  return out;
}

// ---- blocos aba 1 ----
function buildE1(base){
  const horas=base;
  renderChart('e1-1',HOUR_LAB,[
    {name:'Incidente',color:BLUE[0],rows:i=>horas.filter(isIncid).filter(r=>parseInt(r.he,10)===parseInt(i,10))},
    {name:'Reincidente',color:BLUE[1],rows:i=>horas.filter(isReinc).filter(r=>parseInt(r.he,10)===parseInt(i,10))},
  ],{stacked:true});
  renderChart('e1-2',TURN,[
    {name:'Incidente',color:BLUE[0],rows:t=>horas.filter(r=>isIncid(r)&&r.ta===t)},
    {name:'Reincidente',color:BLUE[1],rows:t=>horas.filter(r=>isReinc(r)&&r.ta===t)},
  ],{});
  const reinc=horas.filter(isReinc);
  const rtK=REIN.filter(v=>reinc.some(r=>r.rt===v));
  renderChart('e1-3',rtK,[
    {name:'Equipe',color:BLUE[0],rows:v=>reinc.filter(r=>r.rt===v&&r.eo==='equipe')},
    {name:'Operador',color:BLUE[1],rows:v=>reinc.filter(r=>r.rt===v&&r.eo==='operador')},
  ],{stacked:true});
  const primR=horas.filter(isPrim);   // reincidencias 1° (rt contem "1°")
  const ca=tally(primR.filter(r=>r.ca&&r.ca!=='---'),['ca']);
  const caK=Object.keys(ca).sort((a,b)=>ca[b]-ca[a]).slice(0,15);
  renderChart('e1-4',caK,[{name:'Qtd',color:BLUE[0],rows:k=>primR.filter(r=>r.ca===k)}],{indexAxis:'y'});
  const eqm=tally(primR.filter(okEqd),['eqd']);
  const eqK=Object.keys(eqm).sort((a,b)=>eqm[b]-eqm[a]).slice(0,10);
  renderChart('e1-5',eqK,[{name:'Qtd',color:BLUE[0],rows:k=>primR.filter(r=>r.eqd===k)}],{indexAxis:'y'});
}

// ---- Dashboard 3 - Entrada (visão de período) ----
function barWithMeans(id, labels, datasets, opts, drillFn){
  opts=opts||{};
  const cid=tab==='d3'?('d3-'+id):id;
  const el=document.getElementById(cid);
  if(!el) return;
  const meanLines=opts.meanLines||[];
  const wantsY1=datasets.some(d=>d.yAxisID==='y1');
  const scales={
    x:{stacked:!!opts.stacked,grid:{color:'#e5ebf3'},ticks:{color:'#5b6b7b',font:{size:10}}},
    y:{stacked:!!opts.stacked,beginAtZero:true,grid:{color:'#e5ebf3'},ticks:{color:'#5b6b7b'}}};
  if(wantsY1){
    scales.y1={position:'right',beginAtZero:true,grid:{drawOnChartArea:false},ticks:{color:'#5b6b7b'},title:{display:false}};
  }
  const cfg={
    type:'bar',
    data:{labels:labels,datasets:datasets},
    options:{
      responsive:true,maintainAspectRatio:false,
      scales:scales,
      plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11},color:'#5b6b7b'}},
               tooltip:{enabled:true},
               datalabels:dins(),
                totals:{enabled:true,redLabel:!!opts.redLabel},
               meanLines:{lines:meanLines}}
    },
    plugins:[]
  };
  storeDrill(cid, labels, drillFn||(()=>[]));
  if(opts.legendNote) cfg.options.plugins.legend.__note=opts.legendNote;
  const chart=mk(cid, cfg);
  if(chart&&opts.totSingle) charts[cid].__totSingle=true;
  return chart;
}
function buildE3(base){
  const rows=base;
  const dk=r=>((r.mes!==undefined&&r.mes!==null)?String(r.mes):'?')+'-'+((r.d!==undefined&&r.d!==null)?String(r.d):'?');
  const dm={};
  rows.forEach(r=>{const k=dk(r); if(!dm[k]) dm[k]={mes:r.mes,d:r.d};});
  const keys=Object.keys(dm).sort((a,b)=>{const A=dm[a],B=dm[b];return (Number(A.mes)-Number(B.mes))||(Number(A.d)-Number(B.d));});
  const labels=keys.map(k=>String(dm[k].d)+'/'+String(dm[k].mes));
  if(!labels.length) labels.push('—');
  const bt=rows.filter(r=>r.t==='BT');
  const mt=rows.filter(r=>r.t==='MT');
  const mean=a=>{const t=a.reduce((x,y)=>x+y,0);return a.length?t/a.length:0;};
  const btInc=keys.map(k=>bt.filter(r=>dk(r)===k&&isIncid(r)).length), btRe=keys.map(k=>bt.filter(r=>dk(r)===k&&isReinc(r)).length);
  const mtInc=keys.map(k=>mt.filter(r=>dk(r)===k&&isIncid(r)).length), mtRe=keys.map(k=>mt.filter(r=>dk(r)===k&&isReinc(r)).length);
  const btTot=btInc.map((v,i)=>v+btRe[i]);
  const mtTot=mtInc.map((v,i)=>v+mtRe[i]);
  barWithMeans('e1-1',labels,[
    {label:'Incidente',backgroundColor:BLUE[0],data:btInc},
    {label:'Reincidente',backgroundColor:BLUE[1],data:btRe},
  ],{stacked:true,meanLines:[{value:mean(btTot),color:'#d1242f'}]},
    i=>[bt.filter(r=>dk(r)===keys[i]&&isIncid(r)),bt.filter(r=>dk(r)===keys[i]&&isReinc(r))]);
  barWithMeans('e1-2',labels,[
    {label:'Incidente',backgroundColor:BLUE[0],data:mtInc},
    {label:'Reincidente',backgroundColor:BLUE[1],data:mtRe},
  ],{stacked:true,meanLines:[{value:mean(mtTot),color:'#d1242f'}]},
    i=>[mt.filter(r=>dk(r)===keys[i]&&isIncid(r)),mt.filter(r=>dk(r)===keys[i]&&isReinc(r))]);
  const turns=['A','B','C'];
  const tcol=[BLUE[0],BLUE[1],BLUE[3]];
  const turnData=turns.map((t,i)=>{const data=keys.map(k=>rows.filter(r=>dk(r)===k&&r.ta===t).length);return {label:'Turno '+t,backgroundColor:tcol[i],data:data,mean:mean(data)};});
  const turnMean=turnData.map(d=>d.mean);
  barWithMeans('e1-3',labels,turnData,{stacked:false,meanLines:tcol.map((c,i)=>({value:turnMean[i],color:c}))},
    i=>turns.map((t,ti)=>rows.filter(r=>dk(r)===keys[i]&&r.ta===t)));
  const cat=r=>{const v=Number(r.a3)||0;return v<=0?'sem':(v===1?'ind':'col');};
  const afSem=keys.map(k=>rows.filter(r=>dk(r)===k&&cat(r)==='sem').length);
  const afInd=keys.map(k=>rows.filter(r=>dk(r)===k&&cat(r)==='ind').length);
  const afCol=keys.map(k=>rows.filter(r=>dk(r)===k&&cat(r)==='col').length);
  const afCol2=[BLUE[1],BLUE[0],'#0a3069'];
  barWithMeans('e1-4',labels,[
    {label:'Sem afetação',backgroundColor:'#8b949e',data:afSem},
    {label:'Individual',backgroundColor:afCol2[1],data:afInd},
    {label:'Coletivo',backgroundColor:afCol2[2],data:afCol},
  ],{stacked:true,meanLines:[
    {value:mean(afSem),color:'#8b949e'},
    {value:mean(afInd),color:afCol2[1]},
    {value:mean(afCol),color:afCol2[2]},
  ]},
    i=>[rows.filter(r=>dk(r)===keys[i]&&cat(r)==='sem'),rows.filter(r=>dk(r)===keys[i]&&cat(r)==='ind'),rows.filter(r=>dk(r)===keys[i]&&cat(r)==='col')]);
}

function buildS1(base){
  const fe=base.filter(r=>r.s==='Fechado');
  renderChart('s1-1',HOUR_LAB,[
    {name:'Equipe',color:ORANGE[0],rows:i=>fe.filter(r=>r.eo==='equipe').filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
    {name:'Operador',color:ORANGE[1],rows:i=>fe.filter(r=>r.eo==='operador').filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
  ],{stacked:true});
  renderChart('s1-2',TURN,[
    {name:'Equipe',color:ORANGE[0],rows:t=>fe.filter(r=>r.eo==='equipe'&&r.ts===t)},
    {name:'Operador',color:ORANGE[1],rows:t=>fe.filter(r=>r.eo==='operador'&&r.ts===t)},
  ],{});
  const eK=ENTR_ORD.filter(v=>fe.some(r=>r.e===v));
  renderChart('s1-3',HOUR_LAB,eK.map((v,i)=>({name:v,color:ORANGE[i%ORANGE.length],rows:h=>fe.filter(r=>r.e===v).filter(r=>parseInt(r.hs,10)===parseInt(h,10))})),{});
  const U=UF[tab]||{};
  const F=FSTATE[CUR_SUB]||{};
  const C4={ut:U.ut,p:U.p,suc:U.suc,nt:U.nt,gpd:U.gpd,e:F.e,ta:mergeF(U.ta,F.ta),ts:mergeF(U.ts,F.ts)};
  const ab4=RAW.filter(r=>r.s==='Aberto'&&passes(r,C4));
  renderChart('s1-4',ENTR_ORD,[
    {name:'Atribuído',color:ORANGE[0],rows:e=>ab4.filter(r=>r.s==='Aberto'&&!isIgnCausa(r)&&r.e===e&&r.atr==='Atribuído')},
    {name:'Não Atribuído',color:ORANGE[1],rows:e=>ab4.filter(r=>r.s==='Aberto'&&!isIgnCausa(r)&&r.e===e&&r.atr==='Não Atribuído')},
  ],{stacked:true});
}

// ---- Dashboard 3 - Saída (visão de período) ----
function buildS3(base){
  const fe=base.filter(r=>r.s==='Fechado');
  const dk=r=>((r.mes!==undefined&&r.mes!==null)?String(r.mes):'?')+'-'+((r.d!==undefined&&r.d!==null)?String(r.d):'?');
  const dm={};
  fe.forEach(r=>{const k=dk(r); if(!dm[k]) dm[k]={mes:r.mes,d:r.d};});
  const keys=Object.keys(dm).sort((a,b)=>{const A=dm[a],B=dm[b];return (Number(A.mes)-Number(B.mes))||(Number(A.d)-Number(B.d));});
  const labels=keys.map(k=>String(dm[k].d)+'/'+String(dm[k].mes));
  if(!labels.length) labels.push('—');
  const mean=a=>{const t=a.reduce((x,y)=>x+y,0);return a.length?t/a.length:0;};

  // 1) Equipe x Operador por dia (empilhado) + media do total
  const eqD=keys.map(k=>fe.filter(r=>dk(r)===k&&r.eo==='equipe').length);
  const opD=keys.map(k=>fe.filter(r=>dk(r)===k&&r.eo==='operador').length);
  const totD=eqD.map((v,i)=>v+opD[i]);
  barWithMeans('s3-1',labels,[
    {label:'Equipe',backgroundColor:ORANGE[0],data:eqD},
    {label:'Operador',backgroundColor:ORANGE[1],data:opD},
  ],{stacked:true,meanLines:[{value:mean(totD),color:'#d1242f'}]},
    i=>[fe.filter(r=>dk(r)===keys[i]&&r.eo==='equipe'),fe.filter(r=>dk(r)===keys[i]&&r.eo==='operador')]);

  // 2) Produtivo x Improdutivo x Produtivo s/ afetacao (equipe only) + medias por categoria
  const eqOnly=fe.filter(r=>r.eo==='equipe');
  const imp=keys.map(k=>eqOnly.filter(r=>dk(r)===k&&r.pd==='Improdutivo').length);
  const pro=keys.map(k=>eqOnly.filter(r=>dk(r)===k&&r.pd==='Produtivo').length);
  const proS=keys.map(k=>eqOnly.filter(r=>dk(r)===k&&r.pd==='Produtivo sem afetação').length);
  const cImp=ORANGE[3], cPro=ORANGE[0], cProS=ORANGE[1];
  barWithMeans('s3-2',labels,[
    {label:'Improdutivo',backgroundColor:cImp,data:imp},
    {label:'Produtivo',backgroundColor:cPro,data:pro},
    {label:'Produtivo s/ afetação',backgroundColor:cProS,data:proS},
  ],{stacked:true,meanLines:[
    {value:mean(imp),color:cImp},
    {value:mean(pro),color:cPro},
    {value:mean(proS),color:cProS},
  ]},
    i=>[eqOnly.filter(r=>dk(r)===keys[i]&&r.pd==='Improdutivo'),eqOnly.filter(r=>dk(r)===keys[i]&&r.pd==='Produtivo'),eqOnly.filter(r=>dk(r)===keys[i]&&r.pd==='Produtivo sem afetação')]);

  // 3) Improdutivos por turno (A/B/C) por dia (lado a lado) + medias por turno
  const turns=['A','B','C'];
  const tcol=[ORANGE[3],ORANGE[0],ORANGE[1]];
  const turnData=turns.map((t,i)=>{const data=keys.map(k=>fe.filter(r=>dk(r)===k&&r.ts===t&&r.pd==='Improdutivo').length);return {label:'Turno '+t,backgroundColor:tcol[i],data:data,mean:mean(data)};});
  const turnMean=turnData.map(d=>d.mean);
  barWithMeans('s3-3',labels,turnData,{stacked:false,meanLines:tcol.map((c,i)=>({value:turnMean[i],color:c}))},
    i=>turns.map(t=>fe.filter(r=>dk(r)===keys[i]&&r.ts===t&&r.pd==='Improdutivo')));

  // 4) Top 10 causa (fechados, apenas equipe) - barras horizontais
  const caFe=fe.filter(r=>r.eo==='equipe');
  const ca=tally(caFe.filter(r=>r.ca&&r.ca!=='---'),['ca']);
  const caK=Object.keys(ca).sort((a,b)=>ca[b]-ca[a]).slice(0,10);
  renderChart('s3-4',caK,[{name:'Qtd',color:ORANGE[0],rows:k=>caFe.filter(r=>r.ca===k)}],{indexAxis:'y'});
}

function buildM1(base){
  if(CUR_SUB==='m3') return buildM1_d3(base);
  const F=FSTATE[CUR_SUB]||{};
  const dia=(F.dia&&!F.dia.all&&F.dia.sel.size)?[...F.dia.sel]:[mostDay()];
  const ent=base.filter(r=>dia.includes(entKey(r)));
  const sai=base.filter(r=>r.s==='Fechado'&&dia.includes(saiKey(r)));
  const saiEq=sai.filter(r=>r.eo==='equipe');
  renderChart('m1-1',TURN,[
    {name:'Entrada',color:'#1f6feb',rows:t=>ent.filter(r=>r.ta===t)},
    {name:'Saída',color:'#f0820f',rows:t=>sai.filter(r=>r.ts===t)},
  ],{});
  renderChart('m1-2',HOUR_LAB,[
    {name:'Entrada',color:'#1f6feb',area:'rgba(31,111,235,.12)',rows:i=>ent.filter(r=>parseInt(r.he,10)===parseInt(i,10))},
    {name:'Saída',color:'#f0820f',area:'rgba(240,130,15,.12)',rows:i=>sai.filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
    {name:'Saída Equipe',color:'#1a7f37',area:'rgba(26,127,55,.12)',rows:i=>saiEq.filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
  ],{type:'line'});
  renderChart('m1-3',HOUR_LAB,[
    {name:'Entrada',color:'#1f6feb',area:'rgba(31,111,235,.12)',wt:'a3',rows:i=>ent.filter(r=>parseInt(r.he,10)===parseInt(i,10))},
    {name:'Saída',color:'#f0820f',area:'rgba(240,130,15,.12)',wt:'a3',rows:i=>sai.filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
    {name:'Saída Equipe',color:'#1a7f37',area:'rgba(26,127,55,.12)',wt:'a3',rows:i=>saiEq.filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
  ],{type:'line'});
}

function buildBk1(base){
  buildCarry1(base);
  buildBacklog1(base);
}

function lastHourOf(base){
  const cats=new Set(base.map(r=>r.e));
  const most=cats.has('Hoje')?'Hoje':(cats.has('Ontem')?'Ontem':'Dias Anteriores');
  let mx=-1;
  base.filter(r=>r.e===most).forEach(r=>{
    [r.he,r.hs].forEach(v=>{const n=parseInt(v,10);if(Number.isFinite(n)&&n>mx)mx=n;});
  });
  return mx<0?23:mx;
}

function buildBacklog1(base){
  const F=FSTATE[CUR_SUB]||{};
  const dia=(F.dia&&!F.dia.all&&F.dia.sel.size)?[...F.dia.sel]:[mostDay()];
  const curK=dia.slice().sort((a,b)=>dateVal(b)-dateVal(a))[0];
  const ent=base.filter(r=>dia.includes(entKey(r)));
  const LH=(dia.length===1)?lastHourOfDate(base,curK):23;
  const openAtEnd=(r,H)=>{
    const he=parseInt(r.he,10);
    if(!Number.isFinite(he)||he>H) return false;
    if(r.s==='Aberto') return true;
    const sk=saiKey(r);
    if(!sk) return false;
    if(dateVal(sk)>dateVal(curK)) return true;
    if(sk===curK){ const hs=parseInt(r.hs,10); return Number.isFinite(hs)&&hs>H; }
    return false;
  };
  const newH=HOUR_LAB.map((h,i)=>ent.filter(r=>parseInt(r.he,10)===i&&openAtEnd(r,i)).length);
  const carryH=HOUR_LAB.map((h,i)=>ent.filter(r=>{const he=parseInt(r.he,10);return Number.isFinite(he)&&he<i&&openAtEnd(r,i);}).length);
  const pctH=HOUR_LAB.map((h,i)=>{
    if(i===0) return null;
    const prev=ent.filter(r=>parseInt(r.he,10)===i-1);
    if(!prev.length) return null;
    const done=prev.filter(r=>r.s==='Fechado'&&parseInt(r.hs,10)===i).length;
    return Math.round(100*done/prev.length);
  });
  const newC=newH.map((v,i)=>i<=LH?v:0);
  const carC=carryH.map((v,i)=>i<=LH?v:0);
  const pctC=pctH.map((v,i)=>i<=LH?v:null);
  barWithMeans('b1-1',HOUR_LAB,[
    {label:'Backlog hora atual',backgroundColor:'#466067',data:newC,order:1},
    {label:'Backlog hora anterior',backgroundColor:'#34baab',data:carC,order:2},
    {label:'% resolvidos da hora anterior',type:'line',yAxisID:'y1',data:pctC,borderColor:'#a7cd2c',backgroundColor:'#a7cd2c',borderWidth:2,pointRadius:2,fill:false,tension:0.3,order:0,datalabels:{display:true,color:'#000',backgroundColor:'rgba(255,255,255,.9)',borderRadius:3,padding:2,font:{size:9,weight:'700'},formatter:v=>v==null?'':v+'%'}},
  ],{stacked:true,redLabel:false,meanLines:[]},
    i=>{if(i>LH)return[[],[],[]];return[ent.filter(r=>parseInt(r.he,10)===i),ent.filter(r=>parseInt(r.he,10)<i),[]];});
}

function buildCarry1(base){
  const mean=a=>{const t=a.reduce((x,y)=>x+y,0);return a.length?t/a.length:0;};
  const F=FSTATE[CUR_SUB]||{};
  const dia=(F.dia&&!F.dia.all&&F.dia.sel.size)?[...F.dia.sel]:[mostDay()];
  const curK=dia.slice().sort((a,b)=>dateVal(b)-dateVal(a))[0];
  const antK=prevDayKey(curK);
  const LH=(dia.length===1)?lastHourOfDate(base,curK):23;
  const atual=base.filter(r=>entKey(r)===curK);
  const ant=base.filter(r=>entKey(r)===antK);
  const openSai=(r,H)=>{
    if(r.s==='Aberto') return true;
    const sk=saiKey(r);
    if(!sk) return false;
    if(dateVal(sk)>dateVal(curK)) return true;
    if(sk===curK){ const hs=parseInt(r.hs,10); return Number.isFinite(hs)&&hs>H; }
    return false;
  };
  const openAtu=(r,H)=>{const he=parseInt(r.he,10);if(!Number.isFinite(he)||he>H)return false;return openSai(r,H);};
  const openAnt=(r,H)=>openSai(r,H);
  const antH=HOUR_LAB.map((h,i)=>ant.filter(r=>openAnt(r,i)).length);
  const atuH=HOUR_LAB.map((h,i)=>atual.filter(r=>openAtu(r,i)).length);
  const antC=antH.map((v,i)=>i<=LH?v:0);
  const atuC=atuH.map((v,i)=>i<=LH?v:0);
  barWithMeans('cb1-1',HOUR_LAB,[
    {label:'Backlog dia anterior',backgroundColor:'#34baab',data:antC},
    {label:'Backlog dia atual',backgroundColor:'#466067',data:atuC},
  ],{stacked:true,redLabel:false,meanLines:[{value:mean(antC),color:'#34baab'},{value:mean(atuC),color:'#466067'}]},
    i=>{if(i>LH)return[[],[]];return[ant.filter(r=>openAnt(r,i)),atual.filter(r=>openAtu(r,i))];});
}

// ---- Dashboard 3 - Entrada & Saída (visão de período) ----
function buildM1_d3(base){
  const ent=base;                                   // entradas (abertos) no período
  const sai=base.filter(r=>r.s==='Fechado');        // saídas = fechados
  const saiEq=sai.filter(r=>r.eo==='equipe');
  // dkEnt: agrupa entradas pela data de início
  const dkEnt=r=>((r.mes!==undefined&&r.mes!==null)?String(r.mes):'?')+'-'+((r.d!==undefined&&r.d!==null)?String(r.d):'?');
  // dkSai: agrupa saídas pela data fim (fallback: data início)
  const _ms=r=>{const v=r.mes_sd;return(v!==undefined&&v!==null&&String(v).trim()!=='')?Number(v):Number(r.mes);};
  const _ds=r=>{const v=r.d_sd;return(v!==undefined&&v!==null&&String(v).trim()!=='')?Number(v):Number(r.d);};
  const dkSai=r=>String(_ms(r))+'-'+String(_ds(r));
  const dm={};
  ent.forEach(r=>{const k=dkEnt(r); if(!dm[k]) dm[k]={mes:Number(r.mes)||0,d:Number(r.d)||0};});
  sai.forEach(r=>{const k=dkSai(r); const m=_ms(r);const d=_ds(r);if(!dm[k]) dm[k]={mes:m,d:d};});
  const keys=Object.keys(dm).sort((a,b)=>{const A=dm[a],B=dm[b];return(A.mes-B.mes)||(A.d-B.d);});
  const labels=keys.map(k=>String(dm[k].d)+'/'+String(dm[k].mes));
  if(!labels.length) labels.push('—');
  const dayKeyOfEnt=r=>String(r.d)+'/'+String(r.mes);
  const dayKeyOfSai=r=>String(_ds(r))+'/'+String(_ms(r));
  const mean=a=>{const t=a.reduce((x,y)=>x+y,0);return a.length?t/a.length:0;};
  const entD=keys.map(k=>ent.filter(r=>dkEnt(r)===k).length);
  const saiD=keys.map(k=>sai.filter(r=>dkSai(r)===k).length);
  const saiEqD=keys.map(k=>saiEq.filter(r=>dkSai(r)===k).length);
  const entA3=keys.map(k=>sumKey(ent.filter(r=>dkEnt(r)===k),'a3'));
  const saiA3=keys.map(k=>sumKey(sai.filter(r=>dkSai(r)===k),'a3'));
  const saiEqA3=keys.map(k=>sumKey(saiEq.filter(r=>dkSai(r)===k),'a3'));

  // m1-1: Entrada x Saída por Dia (barras) + médias
  barWithMeans('m1-1',labels,[
    {label:'Entrada',backgroundColor:BLUE[0],data:entD},
    {label:'Saída',backgroundColor:ORANGE[0],data:saiD},
  ],{stacked:false,meanLines:[{value:mean(entD),color:BLUE[0]},{value:mean(saiD),color:ORANGE[0]}]},
    i=>[ent.filter(r=>dkEnt(r)===keys[i]),sai.filter(r=>dkSai(r)===keys[i])]);

  // m1-2: Entrada x Saída por Dia (linha, c/ Saída Equipe) + médias
  renderChart('m1-2',labels,[
    {name:'Entrada',color:BLUE[0],rows:lab=>ent.filter(r=>dayKeyOfEnt(r)===lab)},
    {name:'Saída',color:ORANGE[0],rows:lab=>sai.filter(r=>dayKeyOfSai(r)===lab)},
    {name:'Saída Equipe',color:'#1a7f37',rows:lab=>saiEq.filter(r=>dayKeyOfSai(r)===lab)},
  ],{type:'line',meanLines:[{value:mean(entD),color:BLUE[0]},{value:mean(saiD),color:ORANGE[0]},{value:mean(saiEqD),color:'#1a7f37'}]});

  // m1-3: Entrada x Saída por Dia — Cli Af 3 min (saída = fechados) + médias
  renderChart('m1-3',labels,[
    {name:'Entrada',color:BLUE[0],wt:'a3',rows:lab=>ent.filter(r=>dayKeyOfEnt(r)===lab)},
    {name:'Saída',color:ORANGE[0],wt:'a3',rows:lab=>sai.filter(r=>dayKeyOfSai(r)===lab)},
    {name:'Saída Equipe',color:'#1a7f37',wt:'a3',rows:lab=>saiEq.filter(r=>dayKeyOfSai(r)===lab)},
  ],{type:'line',meanLines:[{value:mean(entA3),color:BLUE[0]},{value:mean(saiA3),color:ORANGE[0]},{value:mean(saiEqA3),color:'#1a7f37'}]});

}

function buildCarryD3(base){
  const ent=base;
  const sai=base.filter(r=>r.s==='Fechado');
  const entGroups={}; ent.forEach(r=>{const k=entKey(r);(entGroups[k]=entGroups[k]||[]).push(r);});
  const mean=a=>{const t=a.reduce((x,y)=>x+y,0);return a.length?t/a.length:0;};
  const _ms=r=>{const v=r.mes_sd;return(v!==undefined&&v!==null&&String(v).trim()!=='')?Number(v):Number(r.mes);};
  const _ds=r=>{const v=r.d_sd;return(v!==undefined&&v!==null&&String(v).trim()!=='')?Number(v):Number(r.d);};
  const dkEnt=r=>((r.mes!==undefined&&r.mes!==null)?String(r.mes):'?')+'-'+((r.d!==undefined&&r.d!==null)?String(r.d):'?');
  const dkSai=r=>String(_ms(r))+'-'+String(_ds(r));
  const _openSai=(r,H,curK)=>{
    if(r.s==='Aberto') return true;
    const sk=saiKey(r); if(!sk) return false;
    if(dateVal(sk)>dateVal(curK)) return true;
    if(sk===curK){ const hs=parseInt(r.hs,10); return Number.isFinite(hs)&&hs>H; }
    return false;
  };
  const _openAtu=(r,H,curK)=>{const he=parseInt(r.he,10);if(!Number.isFinite(he)||he>H)return false;return _openSai(r,H,curK);};
  const _openAnt=(r,H,curK)=>_openSai(r,H,curK);
  const dm={};
  ent.forEach(r=>{const k=dkEnt(r); if(!dm[k]) dm[k]={mes:Number(r.mes)||0,d:Number(r.d)||0};});
  sai.forEach(r=>{const k=dkSai(r); const m=_ms(r);const d=_ds(r);if(!dm[k]) dm[k]={mes:m,d:d};});
  const keys=Object.keys(dm).sort((a,b)=>{const A=dm[a],B=dm[b];return(A.mes-B.mes)||(A.d-B.d);});
  const labels=keys.map(k=>String(dm[k].d)+'/'+String(dm[k].mes));
  if(!labels.length) labels.push('—');
  const perDay={};
  labels.forEach(k=>{
    const ant=entGroups[prevDayKey(k)]||[];
    const atu=entGroups[k]||[];
    const antH=HOUR_LAB.map((_,h)=>ant.filter(r=>_openAnt(r,h,k)).length);
    const atuH=HOUR_LAB.map((_,h)=>atu.filter(r=>_openAtu(r,h,k)).length);
    perDay[k]={antH,atuH};
  });
  // m1-4: carreamento por DIA (valor no fim do dia = hora 23)
  const carryD=labels.map(k=>perDay[k].antH[23]);
  const openD=labels.map(k=>perDay[k].atuH[23]);
  barWithMeans('m1-4',labels,[
    {label:'Backlog dia anterior',backgroundColor:'#34baab',data:carryD},
    {label:'Backlog dia atual',backgroundColor:'#466067',data:openD},
  ],{stacked:true,meanLines:[{value:mean(carryD),color:'#34baab'},{value:mean(openD),color:'#466067'}]},
    i=>[ (entGroups[prevDayKey(labels[i])]||[]).filter(r=>_openAnt(r,23,labels[i])), (entGroups[labels[i]]||[]).filter(r=>_openAtu(r,23,labels[i])) ]);
  // m1-5: carreamento por HORA — média do período
  const avgAnt=HOUR_LAB.map((_,h)=>mean(labels.map(k=>perDay[k].antH[h])));
  const avgAtu=HOUR_LAB.map((_,h)=>mean(labels.map(k=>perDay[k].atuH[h])));
  barWithMeans('m1-5',HOUR_LAB,[
    {label:'Backlog dia anterior',backgroundColor:'#34baab',data:avgAnt},
    {label:'Backlog dia atual',backgroundColor:'#466067',data:avgAtu},
  ],{stacked:true,meanLines:[{value:mean(avgAnt),color:'#34baab'},{value:mean(avgAtu),color:'#466067'}]},
    i=>[ [].concat(...labels.map(k=>(entGroups[prevDayKey(k)]||[]).filter(r=>_openAnt(r,i,k)))), [].concat(...labels.map(k=>(entGroups[k]||[]).filter(r=>_openAtu(r,i,k)))) ]);
}

function buildBk3(base){ buildCarryD3(base); }

// ---- Dashboard 3 - Reincidentes (visão de período) ----
function buildR3(base){
  const prim=base.filter(isPrim);                   // primeiras causas (1°)
  const {dk,keys,labels,mean}=mkDays(base);
  // 1) Reincidentes (1°) Equipe x Remoto por dia (primeiras causas) + media do total
  const eqD=keys.map(k=>prim.filter(r=>dk(r)===k&&r.eo==='equipe').length);
  const remD=keys.map(k=>prim.filter(r=>dk(r)===k&&r.eo==='operador').length);
  const totD=eqD.map((v,i)=>v+remD[i]);
  barWithMeans('r3-1',labels,[
    {label:'Equipe',backgroundColor:GREEN[0],data:eqD},
    {label:'Remoto',backgroundColor:GREEN[1],data:remD},
  ],{stacked:true,meanLines:[{value:mean(totD),color:'#d1242f'}]},
    i=>[prim.filter(r=>dk(r)===keys[i]&&r.eo==='equipe'),prim.filter(r=>dk(r)===keys[i]&&r.eo==='operador')]);
  // 2) Top 10 primeiras causas (1°) por causa - horizontal
  const ca=tally(prim.filter(r=>r.ca&&r.ca!=='---'),['ca']);
  const caK=Object.keys(ca).sort((a,b)=>ca[b]-ca[a]).slice(0,10);
  renderChart('r3-2',caK,[{name:'Qtd',color:GREEN[0],rows:k=>prim.filter(r=>r.ca===k)}],{indexAxis:'y'});
  // 3) Top 10 equipes (eqd) com reincidencia (1°) - horizontal
  const eqm=tally(prim.filter(okEqd),['eqd']);
  const eqK=Object.keys(eqm).sort((a,b)=>eqm[b]-eqm[a]).slice(0,10);
  renderChart('r3-3',eqK,[{name:'Qtd',color:GREEN[0],rows:k=>prim.filter(r=>r.eqd===k)}],{indexAxis:'y'});
}

// ---- Dashboard 3 - Improdutivo (visão de período) ----
const IMP_COL='#db2777', IMP_COL2='#f295c7';
function buildR4(base){
  const imp=base.filter(r=>String(r.pd||'')==='Improdutivo');   // improdutivos (coluna Produtivo/Improdutivo)
  const {dk,keys,labels,mean}=mkDays(base);
  // 1) Improdutivos Equipe x Remoto por dia + media do total
  const eqD=keys.map(k=>imp.filter(r=>dk(r)===k&&r.eo==='equipe').length);
  const remD=keys.map(k=>imp.filter(r=>dk(r)===k&&r.eo==='operador').length);
  const totD=eqD.map((v,i)=>v+remD[i]);
  barWithMeans('r4-1',labels,[
    {label:'Equipe',backgroundColor:IMP_COL,data:eqD},
    {label:'Remoto',backgroundColor:IMP_COL2,data:remD},
  ],{stacked:true,meanLines:[{value:mean(totD),color:'#d1242f'}]},
    i=>[imp.filter(r=>dk(r)===keys[i]&&r.eo==='equipe'),imp.filter(r=>dk(r)===keys[i]&&r.eo==='operador')]);
  // 2) Top 10 causas (improdutivo) - horizontal
  const ca=tally(imp.filter(r=>r.ca&&r.ca!=='---'),['ca']);
  const caK=Object.keys(ca).sort((a,b)=>ca[b]-ca[a]).slice(0,10);
  renderChart('r4-2',caK,[{name:'Qtd',color:IMP_COL,rows:k=>imp.filter(r=>r.ca===k)}],{indexAxis:'y'});
  // 3) Top 10 equipes (eqd) com improdutivo - horizontal
  const eqm=tally(imp.filter(okEqd),['eqd']);
  const eqK=Object.keys(eqm).sort((a,b)=>eqm[b]-eqm[a]).slice(0,10);
  renderChart('r4-3',eqK,[{name:'Qtd',color:IMP_COL,rows:k=>imp.filter(r=>r.eqd===k)}],{indexAxis:'y'});
}

function buildO1(base){
  const at=base.filter(r=>!isIgnCausa(r)&&r.s==='Aberto');
  const clus=CLUSTER.filter(v=>at.some(r=>r.c===v));
  const mkBt=(id,ent)=>{
    const rows=at.filter(r=>r.e===ent&&r.t==='BT');
    renderChart(id,clus,[
      {name:'Atribuído',color:GREEN[0],rows:c=>rows.filter(r=>r.c===c&&r.atr==='Atribuído')},
      {name:'Não Atribuído',color:GREEN[1],rows:c=>rows.filter(r=>r.c===c&&r.atr==='Não Atribuído')},
    ],{stacked:true});
  };
  const mkMt=(id,ent)=>{
    const rows=at.filter(r=>r.e===ent&&r.t==='MT');
    renderChart(id,clus,[
      {name:'Atribuído',color:GREEN[0],rows:c=>rows.filter(r=>r.c===c&&r.atr==='Atribuído')},
      {name:'Não Atribuído',color:GREEN[1],rows:c=>rows.filter(r=>r.c===c&&r.atr==='Não Atribuído')},
    ],{stacked:true});
  };
  mkBt('o1-1','Hoje'); mkBt('o1-2','Ontem'); mkBt('o1-3','Dias Anteriores');
  mkMt('o1-4','Hoje'); mkMt('o1-5','Ontem'); mkMt('o1-6','Dias Anteriores');
  const ab=at.filter(r=>r.s==='Aberto');
  renderChart('o1-7',clus,[
    {name:'Atribuído',color:GREEN[0],wt:'ct',rows:c=>ab.filter(r=>r.c===c&&r.atr==='Atribuído')},
    {name:'Não Atribuído',color:GREEN[1],wt:'ct',rows:c=>ab.filter(r=>r.c===c&&r.atr==='Não Atribuído')},
  ],{stacked:true});
}

function buildXO(base){
  if(CUR_SUB==='xO3') return buildXO_d3(base);
  const F=FSTATE[CUR_SUB]||{};
  const dia=(F.dia&&!F.dia.all&&F.dia.sel.size)?[...F.dia.sel]:[mostDay()];
  const ent=base.filter(r=>dia.includes(entKey(r)));
  const sai=base.filter(r=>r.s==='Fechado'&&dia.includes(saiKey(r)));
  const osmEnt=ent.filter(r=>String(r.osm).toUpperCase()==='SIM');
  const osmSai=sai.filter(r=>String(r.osm).toUpperCase()==='SIM');
  renderChart('xO-1',HOUR_LAB,[
    {name:'Entrada OSM',color:'#1f6feb',area:'rgba(31,111,235,.12)',rows:i=>osmEnt.filter(r=>parseInt(r.he,10)===parseInt(i,10))},
    {name:'Saída OSM',color:'#d1242f',area:'rgba(209,36,47,.12)',rows:i=>osmSai.filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
  ],{type:'line'});
  setVCard('xO-ent',osmEnt.length);
  setVCard('xO-sai',osmSai.length);
  vcardRows['xO-ent']=osmEnt;
  vcardRows['xO-sai']=osmSai;
}

function fillKpis(id,items){
  const el=document.getElementById(id);
  if(!el) return;
  el.innerHTML=items.map(b=>`<div class="kpi ${b.cls}"><div class="v">${b.v}</div><div class="l">${b.l}</div></div>`).join('');
}
const fmtN=n=>Number(n||0).toLocaleString('pt-BR');
function dCount(rows,key){const s=new Set(rows.map(r=>r[key]).filter(v=>v&&v!==''&&v!=='---'));return s.size;}
function setVCard(id,v){
  if(tab==='d3') id='d3-'+id;
  const el=document.getElementById(id);
  if(el) el.textContent=fmtN(v);
}
function buildX1(base){
  if(CUR_SUB==='x31') return buildX1_d3(base);
  const mk=(id,ent)=>{
    const p24=base.filter(r=>r.c==='24h+'&&r.e===ent);
    const lab=comboKeys(p24,['t','ta']);
    if(!lab.length) return;
    renderChart(id,lab,[
      {name:'Aberto',color:'#d1242f',rows:l=>p24.filter(r=>r.t+' · '+r.ta===l&&r.s==='Aberto')},
      {name:'Fechado Hoje',color:'#1a7f37',rows:l=>p24.filter(r=>r.t+' · '+r.ta===l&&r.s==='Fechado'&&r.sd==='Hoje')},
    ],{stacked:true});
  };
  const p24=base.filter(r=>r.c==='24h+');
  vcardRows['x24-sim']=p24.filter(r=>r.s==='Aberto');
  vcardRows['x24-nao']=p24.filter(r=>r.s==='Fechado'&&r.sd==='Hoje');
  setVCard('x24-sim',vcardRows['x24-sim'].length);
  setVCard('x24-nao',vcardRows['x24-nao'].length);
  mk('x1b','Ontem'); mk('x1c','Dias Anteriores');
}
function buildX2(base){
  if(CUR_SUB==='x32') return buildX2_d3(base);
  const fe=base.filter(r=>r.s==='Fechado'&&r.sd==='Hoje');
  const lab=ENTR_ORD.filter(v=>fe.some(r=>r.e===v));
  renderChart('x1-2',lab,[
    {name:'sim',color:'#1a7f37',rows:e=>fe.filter(r=>r.e===e&&r.o==='sim')},
    {name:'não',color:'#d1242f',rows:e=>fe.filter(r=>r.e===e&&r.o==='não')},
  ],{stacked:true});
  vcardRows['x2-sim']=fe.filter(r=>r.o==='sim');
  vcardRows['x2-nao']=fe.filter(r=>r.o==='não');
  setVCard('x2-sim',vcardRows['x2-sim'].length);
  setVCard('x2-nao',vcardRows['x2-nao'].length);
}
function buildX3(base){
  if(CUR_SUB==='x33') return buildX3_d3(base);
  const fe=base.filter(r=>r.s==='Fechado'&&r.sd==='Hoje');
  const mt=fe.filter(r=>r.t==='MT'&&(r.ro==='S'||r.ro==='N'));
  const lab=comboKeys(mt,['e','nt']);
  renderChart('x1-3',lab,[
    {name:'S',color:'#1a7f37',rows:l=>mt.filter(r=>r.e+' · '+r.nt===l&&r.ro==='S')},
    {name:'N',color:'#d1242f',rows:l=>mt.filter(r=>r.e+' · '+r.nt===l&&r.ro==='N')},
  ],{stacked:true});
  vcardRows['x3-sim']=mt.filter(r=>r.ro==='S');
  vcardRows['x3-nao']=mt.filter(r=>r.ro==='N');
  setVCard('x3-sim',vcardRows['x3-sim'].length);
  setVCard('x3-nao',vcardRows['x3-nao'].length);
}

// ---- d3 variantes per-day + média ----
function buildX1_d3(base){
  const dm=r=>{const n=parseInt(r.dur,10);return isNaN(n)?0:n;};
  const p24=base.filter(r=>r.s==='Fechado'&&dm(r)>1440);            // fechados com +24h (duração em min)
  const men=base.filter(r=>r.s==='Fechado'&&dm(r)<=1440);           // fechados com -24h
  const {dk,keys,labels,mean}=mkDays(base);
  const mais=keys.map(k=>p24.filter(r=>dk(r)===k).length);
  barWithMeans('x1b',labels,[
    {label:'Fechados >24h',backgroundColor:'#d1242f',data:mais},
  ],{meanLines:[{value:mean(mais),color:'#d1242f'}]},
    i=>[p24.filter(r=>dk(r)===keys[i])]);
  vcardRows['x24-sim']=p24;
  vcardRows['x24-nao']=men;
  setVCard('x24-sim',vcardRows['x24-sim'].length);
  setVCard('x24-nao',vcardRows['x24-nao'].length);
}

function buildXO_d3(base){
  const ent=base, sai=base.filter(r=>r.s==='Fechado');
  const osmEnt=ent.filter(r=>String(r.osm).toUpperCase()==='SIM');
  const osmSai=sai.filter(r=>String(r.osm).toUpperCase()==='SIM');
  const {dk,keys,labels,mean}=mkDays(base);
  const eD=keys.map(k=>osmEnt.filter(r=>dk(r)===k).length);
  const sD=keys.map(k=>osmSai.filter(r=>dk(r)===k).length);
  renderChart('xO-1',labels,[
    {name:'Entrada OSM',color:BLUE[0],rows:lab=>osmEnt.filter(r=>dayKeyOf(r)===lab)},
    {name:'Saída OSM',color:ORANGE[0],rows:lab=>osmSai.filter(r=>dayKeyOf(r)===lab)},
  ],{type:'line',meanLines:[{value:mean(eD),color:BLUE[0]},{value:mean(sD),color:ORANGE[0]}]});
  setVCard('xO-ent',osmEnt.length);
  setVCard('xO-sai',osmSai.length);
  vcardRows['xO-ent']=osmEnt;
  vcardRows['xO-sai']=osmSai;
}

function buildX2_d3(base){
  const fe=base.filter(r=>r.s==='Fechado');
  const {dk,keys,labels,mean}=mkDays(base);
  const sim=keys.map(k=>fe.filter(r=>dk(r)===k&&r.o==='sim').length);
  const nao=keys.map(k=>fe.filter(r=>dk(r)===k&&r.o==='não').length);
  const tot=sim.map((v,i)=>v+nao[i]);
  barWithMeans('x1-2',labels,[
    {label:'Sim',backgroundColor:GREEN[0],data:sim},
    {label:'Não',backgroundColor:'#d1242f',data:nao},
  ],{stacked:true,meanLines:[{value:mean(tot),color:'#5f6b7a'}]},
    i=>[fe.filter(r=>dk(r)===keys[i]&&r.o==='sim'),fe.filter(r=>dk(r)===keys[i]&&r.o==='não')]);
  vcardRows['x2-sim']=fe.filter(r=>r.o==='sim');
  vcardRows['x2-nao']=fe.filter(r=>r.o==='não');
  setVCard('x2-sim',vcardRows['x2-sim'].length);
  setVCard('x2-nao',vcardRows['x2-nao'].length);
}

function buildX3_d3(base){
  const fe=base.filter(r=>r.s==='Fechado');
  const mt=fe.filter(r=>r.t==='MT'&&(r.ro==='S'||r.ro==='N'));
  const {dk,keys,labels,mean}=mkDays(base);
  const sim=keys.map(k=>mt.filter(r=>dk(r)===k&&r.ro==='S').length);
  const nao=keys.map(k=>mt.filter(r=>dk(r)===k&&r.ro==='N').length);
  const tot=sim.map((v,i)=>v+nao[i]);
  barWithMeans('x1-3',labels,[
    {label:'S',backgroundColor:GREEN[0],data:sim},
    {label:'N',backgroundColor:'#d1242f',data:nao},
  ],{stacked:true,meanLines:[{value:mean(tot),color:'#5f6b7a'}]},
    i=>[mt.filter(r=>dk(r)===keys[i]&&r.ro==='S'),mt.filter(r=>dk(r)===keys[i]&&r.ro==='N')]);
  vcardRows['x3-sim']=mt.filter(r=>r.ro==='S');
  vcardRows['x3-nao']=mt.filter(r=>r.ro==='N');
  setVCard('x3-sim',vcardRows['x3-sim'].length);
  setVCard('x3-nao',vcardRows['x3-nao'].length);
}

// ---- blocos aba 2 ----
function buildE2(base){
  const pl=plOf(base);
  renderChart('e2-1',pl,[
    {name:'Incidente',color:BLUE[0],rows:p=>base.filter(r=>r.p===p&&isIncid(r))},
    {name:'Reincidente',color:BLUE[1],rows:p=>base.filter(r=>r.p===p&&isReinc(r))},
  ],{stacked:true});
  const e2h=pl.map((p,i)=>({name:p,color:BLUE[i%BLUE.length],rows:h=>base.filter(r=>r.p===p).filter(r=>parseInt(r.he,10)===parseInt(h,10))}));
  e2h.sort((a,b)=>lum(a.color)-lum(b.color));
  renderChart('e2-2',HOUR_LAB,e2h,{stacked:true});
}
function buildS2(base){
  const fe=base.filter(r=>r.s==='Fechado');
  const pl=plOf(base);
  renderChart('s2-1',pl,[
    {name:'Equipe',color:ORANGE[0],rows:p=>fe.filter(r=>r.p===p&&r.eo==='equipe')},
    {name:'Operador',color:ORANGE[1],rows:p=>fe.filter(r=>r.p===p&&r.eo==='operador')},
  ],{stacked:true});
  const s2h=pl.map((p,i)=>({name:p,color:ORANGE[i%ORANGE.length],rows:h=>fe.filter(r=>r.p===p).filter(r=>parseInt(r.hs,10)===parseInt(h,10))}));
  s2h.sort((a,b)=>lum(a.color)-lum(b.color));
  renderChart('s2-2',HOUR_LAB,s2h,{stacked:true});
}
function buildO2(base){
  const ab=base.filter(r=>!isIgnCausa(r)&&r.s==='Aberto');
  const lab1=comboKeys(ab,['p','c']);
  const fA=r=>r.atr==='Atribuído';
  renderChart('o2-1',lab1,[
    {name:'Atribuído',color:GREEN[0],rows:l=>ab.filter(r=>r.p+' · '+r.c===l&&fA(r))},
    {name:'Não Atribuído',color:GREEN[1],rows:l=>ab.filter(r=>r.p+' · '+r.c===l&&!fA(r))},
  ],{stacked:true});
  renderChart('o2-2',lab1,[
    {name:'Atribuído',color:GREEN[0],wt:'ct',rows:l=>ab.filter(r=>r.p+' · '+r.c===l&&fA(r))},
    {name:'Não Atribuído',color:GREEN[1],wt:'ct',rows:l=>ab.filter(r=>r.p+' · '+r.c===l&&!fA(r))},
  ],{stacked:true});
}
function buildT3(base){
  const cons=base.filter(r=>r.cg==='Considerar');
  const {dk,keys,labels}=mkDays(cons);
  const tmpD=keys.map(k=>meanKey(cons.filter(r=>dk(r)===k),'tmp'));
  const tmdD=keys.map(k=>meanKey(cons.filter(r=>dk(r)===k),'tmd'));
  const tmeD=keys.map(k=>meanKey(cons.filter(r=>dk(r)===k),'tme'));
  const tmaM=meanKey(cons,'tma');
  const tmpM=meanKey(cons,'tmp');
  const tmdM=meanKey(cons,'tmd');
  const tmeM=meanKey(cons,'tme');
  barWithMeans('t3-1',labels,[
    {label:'TMP',backgroundColor:BLUE[0],data:tmpD},
    {label:'TMD',backgroundColor:ORANGE[0],data:tmdD},
    {label:'TME',backgroundColor:GREEN[0],data:tmeD},
    {label:'TMA',type:'line',data:labels.map(()=>null),borderColor:'#d1242f',borderDash:[6,4],borderWidth:2,pointRadius:0,fill:false,order:99},
  ],{stacked:true,redLabel:true,meanLines:[
    {value:tmaM,color:'#d1242f',label:'TMA'},
    {value:tmpM,color:BLUE[0],label:'TMP'},
    {value:tmdM,color:ORANGE[0],label:'TMD'},
    {value:tmeM,color:GREEN[0],label:'TME'},
  ]},
    i=>[cons.filter(r=>dk(r)===keys[i])]);
  buildT3Turno(base,'A','t3-2');
  buildT3Turno(base,'B','t3-3');
  buildT3Turno(base,'C','t3-4');
  buildT3HoraTurno(base);
  buildT3Tipos(base);
  buildT3Hora(base);
}

function buildT3HoraTurno(base){
  const cons=base.filter(r=>r.cg==='Considerar');
  const tmpD=HOUR_LAB.map((h,i)=>meanKey(cons.filter(r=>parseInt(r.hs,10)===i),'tmp'));
  const tmdD=HOUR_LAB.map((h,i)=>meanKey(cons.filter(r=>parseInt(r.hs,10)===i),'tmd'));
  const tmeD=HOUR_LAB.map((h,i)=>meanKey(cons.filter(r=>parseInt(r.hs,10)===i),'tme'));
  const tmaM=meanKey(cons,'tma');
  const tmpM=meanKey(cons,'tmp');
  const tmdM=meanKey(cons,'tmd');
  const tmeM=meanKey(cons,'tme');
  const chart=barWithMeans('t3-7',HOUR_LAB,[
    {label:'TMP',backgroundColor:BLUE[0],data:tmpD},
    {label:'TMD',backgroundColor:ORANGE[0],data:tmdD},
    {label:'TME',backgroundColor:GREEN[0],data:tmeD},
    {label:'TMA',type:'line',data:HOUR_LAB.map(()=>null),borderColor:'#d1242f',borderDash:[6,4],borderWidth:2,pointRadius:0,fill:false,order:99},
    {label:'Entrada A / Entrada B / Entrada C',backgroundColor:'#8b5cf6',data:HOUR_LAB.map(()=>null),stack:'_ent',order:98,borderWidth:0,__countOverlay:true},
  ],{stacked:true,redLabel:true,legendNote:'16-18h = transição (faixa EXT, fora dos turnos B/C)',meanLines:[
    {value:tmaM,color:'#d1242f',label:'TMA'},
    {value:tmpM,color:BLUE[0],label:'TMP'},
    {value:tmdM,color:ORANGE[0],label:'TMD'},
    {value:tmeM,color:GREEN[0],label:'TME'},
  ]},
    i=>[cons.filter(r=>parseInt(r.hs,10)===i)]);
  if(chart){
    chart.__entradaComp=HOUR_LAB.map((h,i)=>{
      const hour=cons.filter(r=>parseInt(r.hs,10)===i);
      const n=hour.length||1;
      const pct=t=>Math.round(100*hour.filter(r=>r.ta===t).length/n);
      return {A:pct('A'),B:pct('B'),C:pct('C')};
    });
  }
}

function buildT3Tipos(base){
  const cons=base.filter(r=>r.cg==='Considerar');
  const tip=[['BT',r=>String(r.t||'').toUpperCase()==='BT'],['MT',r=>String(r.t||'').toUpperCase()==='MT'],['5RO=S',r=>String(r.ro||'').toUpperCase()==='S'],['OSM=SIM',r=>/sim/i.test(String(r.osm||''))]];
  const labs=tip.map(q=>q[0]);
  const rows=l=>cons.filter(tip.find(q=>q[0]===l)[1]);
  const tmpD=labs.map(l=>meanKey(rows(l),'tmp'));
  const tmdD=labs.map(l=>meanKey(rows(l),'tmd'));
  const tmeD=labs.map(l=>meanKey(rows(l),'tme'));
  const tmaM=meanKey(cons,'tma');
  const tmpM=meanKey(cons,'tmp');
  const tmdM=meanKey(cons,'tmd');
  const tmeM=meanKey(cons,'tme');
barWithMeans('t3-5',labs,[
    {label:'TMP',backgroundColor:BLUE[0],data:tmpD},
    {label:'TMD',backgroundColor:ORANGE[0],data:tmdD},
    {label:'TME',backgroundColor:GREEN[0],data:tmeD},
    {label:'TMA',type:'line',data:labs.map(()=>null),borderColor:'#d1242f',borderDash:[6,4],borderWidth:2,pointRadius:0,fill:false,order:99},
  ],{stacked:true,redLabel:true,meanLines:[
    {value:tmaM,color:'#d1242f',label:'TMA'},
    {value:tmpM,color:BLUE[0],label:'TMP'},
    {value:tmdM,color:ORANGE[0],label:'TMD'},
    {value:tmeM,color:GREEN[0],label:'TME'},
  ]},
    i=>[rows(labs[i])]);
}

function buildT3Hora(base){
  const cons=base.filter(r=>r.cg==='Considerar');
  const fech=cons.filter(r=>r.s==='Fechado');
  const tmpD=HOUR_LAB.map((h,i)=>meanKey(cons.filter(r=>parseInt(r.he,10)===i),'tmp'));
  const tmdD=HOUR_LAB.map((h,i)=>meanKey(cons.filter(r=>parseInt(r.he,10)===i),'tmd'));
  const tmeD=HOUR_LAB.map((h,i)=>meanKey(cons.filter(r=>parseInt(r.he,10)===i),'tme'));
  const fechHs=fech.filter(r=>r.eo==='equipe'&&okEqd(r));
  const days=cons.length?(new Set(cons.map(r=>r.d!==undefined&&r.d!==null?r.d:r.diaMes))).size||1:1;
  const medSaidaD=HOUR_LAB.map((h,i)=>fech.filter(r=>parseInt(r.hs,10)===i).length/days);
  const medEntradaD=HOUR_LAB.map((h,i)=>fech.filter(r=>parseInt(r.he,10)===i).length/days);
  const turnOfHour=h=>h<8?'A':(h<15?'B':'C');
  const bandOfHour=h=>h<8?'A':(h<16?'B':(h<19?'EXT':'C'));
  const dayKey=r=>(r.d!==undefined&&r.d!==null?r.d:r.diaMes);
  const daysList=[...new Set(cons.map(dayKey))];
  const byDay={};
  fech.forEach(r=>{const k=dayKey(r);(byDay[k]=byDay[k]||[]).push(r);});
  const teamByBandDay={A:[],B:[],C:[],EXT:[]};
  daysList.forEach(k=>{const arr=byDay[k]||[];['A','B','C','EXT'].forEach(T=>{const s=new Set();for(const r of arr){if(bandOfHour(parseInt(r.hs,10))===T&&r.ea)s.add(r.ea);}teamByBandDay[T].push(s.size);});});
  const teamAvg={};['A','B','C','EXT'].forEach(T=>{const a=teamByBandDay[T];teamAvg[T]=a.length?a.reduce((x,y)=>x+y,0)/a.length:0;});
  const eqD=HOUR_LAB.map((h,i)=>teamAvg[bandOfHour(i)]);
  const prD=HOUR_LAB.map((h,i)=>{let incSum=0;for(const k of daysList){incSum+=(byDay[k]||[]).filter(r=>parseInt(r.hs,10)===i).length;}const incAvg=daysList.length?incSum/daysList.length:0;const t=teamAvg[bandOfHour(i)];return t>0?incAvg/t:0;});
  const tmaM=meanKey(cons,'tma');
  const tmpM=meanKey(cons,'tmp');
  const tmdM=meanKey(cons,'tmd');
  const tmeM=meanKey(cons,'tme');
  const medSaidaM=fech.length/days;
  const medEntradaM=fech.length/days;
  overlayDrill['t3-6']=[];
  overlayDrillEq['t3-6']=[];
  const chart=barWithMeans('t3-6',HOUR_LAB,[
    {label:'TMP',backgroundColor:BLUE[0],data:tmpD},
    {label:'TMD',backgroundColor:ORANGE[0],data:tmdD},
    {label:'TME',backgroundColor:GREEN[0],data:tmeD},
    {label:'TMA',type:'line',data:HOUR_LAB.map(()=>null),borderColor:'#d1242f',borderDash:[6,4],borderWidth:2,pointRadius:0,fill:false,order:99},
    {label:'Média saída',type:'line',yAxisID:'y1',data:medSaidaD,borderColor:'#8e44ad',backgroundColor:'#8e44ad',borderWidth:2,pointRadius:0,fill:false,order:98,datalabels:{display:false},__countOverlay:true},
    {label:'Média entrada',type:'line',yAxisID:'y1',data:medEntradaD,borderColor:'#0aa2a2',backgroundColor:'#0aa2a2',borderWidth:2,pointRadius:0,fill:false,order:97,datalabels:{display:false},__countOverlay:true},
    {label:'Média equipes/hora',type:'line',yAxisID:'y1',data:eqD,borderColor:'#5f6b7a',backgroundColor:'#5f6b7a',borderWidth:0,pointRadius:0,fill:false,order:96,datalabels:{display:false},__countOverlay:true,noLine:true},
    {label:'Média produção/hora',type:'line',yAxisID:'y1',data:prD,borderColor:'#bf3f7f',backgroundColor:'#bf3f7f',borderWidth:0,pointRadius:0,fill:false,order:95,datalabels:{display:false},__countOverlay:true,noLine:true},
  ],{stacked:true,redLabel:true,legendNote:'16-18h = transição (faixa EXT, fora dos turnos B/C)',meanLines:[
    {value:tmaM,color:'#d1242f',label:'TMA'},
    {value:tmpM,color:BLUE[0],label:'TMP'},
    {value:tmdM,color:ORANGE[0],label:'TMD'},
    {value:tmeM,color:GREEN[0],label:'TME'},
  ],
    totSingle:true},
    i=>{const c=cons.filter(r=>parseInt(r.he,10)===i);const f=fech.filter(r=>parseInt(r.he,10)===i);const eq=fechHs.filter(r=>parseInt(r.hs,10)===i);overlayDrill['t3-6'][i]=f;overlayDrillEq['t3-6'][i]=eq;return [c,c,c,c,f,f,eq,eq];});
  if(chart){
    chart.__countDrill={saida:[],entrada:[],equipes:[],producao:[]};
    HOUR_LAB.forEach((h,i)=>{
      chart.__countDrill.saida[i]=fech.filter(r=>parseInt(r.hs,10)===i);
      chart.__countDrill.entrada[i]=fech.filter(r=>parseInt(r.he,10)===i);
      chart.__countDrill.equipes[i]=fech.filter(r=>parseInt(r.hs,10)===i);
      chart.__countDrill.producao[i]=fech.filter(r=>parseInt(r.hs,10)===i);
    });
    const tmaByBand={A:meanKey(cons.filter(r=>bandOfHour(parseInt(r.he,10))==='A'),'tma'),
                     B:meanKey(cons.filter(r=>bandOfHour(parseInt(r.he,10))==='B'),'tma'),
                     C:meanKey(cons.filter(r=>bandOfHour(parseInt(r.he,10))==='C'),'tma'),
                     EXT:meanKey(cons.filter(r=>bandOfHour(parseInt(r.he,10))==='EXT'),'tma')};
    let dom='A';
    ['B','C','EXT'].forEach(T=>{ if(tmaByBand[T]>tmaByBand[dom]) dom=T; });
    const domRows=cons.filter(r=>bandOfHour(parseInt(r.he,10))===dom);
    const other=cons.filter(r=>bandOfHour(parseInt(r.he,10))!==dom);
    const peso=cons.length?Math.round(100*domRows.length/cons.length):0;
    const tlabel=dom==='EXT'?'Transição 16-18h':'Turno '+dom;
    chart.__horaTurno={
      turno:dom,
      tlabel:tlabel,
      max:tmaByBand[dom]||0,
      imp:(tmaM-(meanKey(other,'tma')||0)),
      peso:peso,
      nDom:domRows.length,
      nGeral:cons.length,
      geral:tmaM,
      sem:meanKey(other,'tma'),
      hours:HOUR_LAB.map((h)=>bandOfHour(parseInt(h,10))===dom).map((on,i)=>on?i:-1).filter(i=>i>=0)
    };
  }
}

function buildT3Turno(base,turno,id){
  const cons=base.filter(r=>r.cg==='Considerar'&&r.ts===turno);
  const {dk,keys,labels}=mkDays(cons);
  const tmpD=keys.map(k=>meanKey(cons.filter(r=>dk(r)===k),'tmp'));
  const tmdD=keys.map(k=>meanKey(cons.filter(r=>dk(r)===k),'tmd'));
  const tmeD=keys.map(k=>meanKey(cons.filter(r=>dk(r)===k),'tme'));
  const tmaM=meanKey(cons,'tma'),tmpM=meanKey(cons,'tmp'),tmdM=meanKey(cons,'tmd'),tmeM=meanKey(cons,'tme');
  const chart=barWithMeans(id,labels,[
    {label:'TMP',backgroundColor:BLUE[0],data:tmpD},
    {label:'TMD',backgroundColor:ORANGE[0],data:tmdD},
    {label:'TME',backgroundColor:GREEN[0],data:tmeD},
    {label:'TMA',type:'line',data:labels.map(()=>null),borderColor:'#d1242f',borderDash:[6,4],borderWidth:2,pointRadius:0,fill:false,order:99},
    {label:'Entrada A / Entrada B / Entrada C',backgroundColor:'#8b5cf6',data:labels.map(()=>null),stack:'_ent',order:98,borderWidth:0,__countOverlay:true},
  ],{stacked:true,redLabel:true,meanLines:[
    {value:tmaM,color:'#d1242f',label:'TMA'},
    {value:tmpM,color:BLUE[0],label:'TMP'},
    {value:tmdM,color:ORANGE[0],label:'TMD'},
    {value:tmeM,color:GREEN[0],label:'TME'},
  ]},
    i=>[cons.filter(r=>dk(r)===keys[i])]);
  if(chart){
    chart.__entradaComp=keys.map(k=>{
      const day=cons.filter(r=>dk(r)===k);
      const n=day.length||1;
      const pct=t=>Math.round(100*day.filter(r=>r.ta===t).length/n);
      return {A:pct('A'),B:pct('B'),C:pct('C')};
    });
  }
}

function buildT2(base){
  const cons=base.filter(r=>r.s==='Fechado'&&r.cg==='Considerar');
  const pl=plOf(base);
  renderChart('t2-1',pl,[{name:'Qtd',color:'#1f6feb',rows:p=>cons.filter(r=>r.p===p)}],{});
  const fe=base.filter(r=>r.s==='Fechado');
  const eqm=tally(fe.filter(okEqd),['eqd']);
  const eqK=Object.keys(eqm).sort((a,b)=>eqm[b]-eqm[a]).slice(0,20);
  renderChart('t2-2',eqK,[{name:'Qtd',color:'#f0820f',rows:k=>cons.filter(r=>r.eqd===k&&okEqd(r))}],{indexAxis:'y'});
  renderChart('t2-3',pl,[{name:'Média TMA',color:'#1f6feb',rows:p=>cons.filter(r=>r.p===p),val:rows=>meanKey(rows,'tma')}],{});
}
function buildP2(base){
const pl=plOf(base);
  const at=base.filter(r=>!isIgnCausa(r)&&r.s==='Aberto');
  const lab1=comboKeys(at,['p','e']);
  const fA=r=>r.atr==='Atribuído';
  renderChart('p2-1',lab1,[
    {name:'Atribuído',color:PALETTE[0],rows:l=>at.filter(r=>r.p+' · '+r.e===l&&fA(r))},
    {name:'Não Atribuído',color:PALETTE[1],rows:l=>at.filter(r=>r.p+' · '+r.e===l&&!fA(r))},
  ],{stacked:true});
  const ab24=base.filter(r=>!isIgnCausa(r)&&r.c==='24h+'&&r.s==='Aberto');
  const lab2=comboKeys(ab24,['p','e']);
  renderChart('p2-2',lab2,[
    {name:'Atribuído',color:PALETTE[0],rows:l=>ab24.filter(r=>r.p+' · '+r.e===l&&fA(r))},
    {name:'Não Atribuído',color:PALETTE[1],rows:l=>ab24.filter(r=>r.p+' · '+r.e===l&&!fA(r))},
  ],{stacked:true});
  const abEquipe=base.filter(r=>r.s==='Aberto'&&r.eo==='equipe');
  const gpK=[...new Set(abEquipe.map(r=>r.gp))].filter(v=>v&&v!=='').sort();
  renderChart('p2-3',pl,gpK.map((g,i)=>({name:g,color:PALETTE[i%PALETTE.length],rows:p=>abEquipe.filter(r=>r.p===p&&r.gp===g)})),{});
  const imp=base.filter(r=>r.s==='Fechado'&&r.pd==='Improdutivo'&&r.eo==='equipe'&&okEqd(r));
  const gpdI=[...new Set(imp.map(r=>r.gpd))].filter(v=>v&&v!=='').sort();
  renderChart('p2-4',pl,gpdI.map((g,i)=>({name:g,color:PALETTE[i%PALETTE.length],rows:p=>imp.filter(r=>r.p===p&&r.gpd===g)})),{});
  const feEq=base.filter(r=>r.s==='Fechado'&&r.eo==='equipe'&&okEqd(r));
  const gpdF=[...new Set(feEq.map(r=>r.gpd))].filter(v=>v&&v!=='').sort();
  renderChart('p2-5',pl,gpdF.map((g,i)=>({name:g,color:PALETTE[i%PALETTE.length],rows:p=>feEq.filter(r=>r.p===p&&r.gpd===g)})),{});
}

// ---- KPIs ----
const KPI_TITLES={abH:'Abertos Hoje',abO:'Abertos Ontem',abA:'Abertos Dias Anteriores',feH:'Fechados Hoje',feO:'Fechados Ontem',feA:'Fechados Dias Anteriores',naoAtrib:'Não Atribuídos',abAgora:'Abertos Agora',mais24:'Abertos +24h',cliAf:'Clientes Afetados (abertos)',tmaMed:'TMA Médio (fechados)',eqAtrib:'Equipes Atribuídas (abertos)',eqExec:'Equipes em Execução (abertos)',eqTotDesl:'Equipes Totais Deslocadas (fechados hoje)',prod:'Produção Total (fechados hoje / equipes deslocadas)',tmeMed:'TME Médio — TME Análise=sim'};
let kpiRows={};
let kpiTitles=Object.assign({},KPI_TITLES);
kpiTitles['d3_abertos']='Abertos (período)';
kpiTitles['d3_abertos_dia']='Abertos (média/dia)';
kpiTitles['d3_cliAf']='Clientes Afetados (período)';
kpiTitles['d3_cliAf_dia']='Clientes Afetados (média/dia)';
kpiTitles['d3_fechados']='Fechados (período)';
kpiTitles['d3_fechados_dia']='Fechados (média/dia)';
kpiTitles['d3_tma']='TMA Médio';
kpiTitles['d3_tme']='TME Médio';
kpiTitles['d3_tmd']='TMD Médio';
kpiTitles['d3_eqDia']='Média de Equipes Deslocadas';
kpiTitles['d3_prod']='Produção média dia por equipes deslocadas';
['A','B','C'].forEach(t=>{kpiTitles['d3_prod_t'+t]='Produção Turno '+t; kpiTitles['d3_eqDia_t'+t]='Média de Equipes Deslocadas Turno '+t;});
['EMERGÊNCIA','COMERCIAL','PERDAS','PODA','LINHA VIVA'].forEach(g=>{kpiTitles['d3_prod_g_'+g]='Produção '+g;});
function buildKPIs(){
  const U=UF[tab]||{};
  const base=RAW.filter(r=>passes(r,U));
  const ab=base.filter(r=>r.s==='Aberto');
  const fe=base.filter(r=>r.s==='Fechado');
  const hoje=base.filter(r=>r.e==='Hoje');
  const feHoje=fe.filter(r=>r.sd==='Hoje');
  const hasEq=r=>{const v=String(r.eq||'').trim();return v!==''&&v!=='---';};
  const hasEqd=r=>{const v=String(r.eqd||'').trim();return v!==''&&v!=='---';};
  kpiRows={
    abH:ab.filter(r=>r.e==='Hoje'),
    abO:ab.filter(r=>r.e==='Ontem'),
    abA:ab.filter(r=>r.e==='Dias Anteriores'),
    feH:feHoje,
    feO:fe.filter(r=>r.sd==='Ontem'),
    feA:fe.filter(r=>r.sd==='Dias Anteriores'),
    naoAtrib:base.filter(r=>r.s==='Aberto'&&r.atr==='Não Atribuído'),
    abAgora:ab,
    mais24:ab.filter(r=>r.c==='24h+'),
    cliAf:ab,
    tmaMed:fe.filter(r=>r.cg==='Considerar'),
    eqAtrib:ab.filter(hasEq),
    eqExec:ab.filter(hasEqd),
    eqTotDesl:feHoje.filter(hasEqd),
    tmeMed:fe.filter(r=>r.tmea==='sim'),
    tmdMed:fe,
  };
  const fmt=n=>Number(n||0).toLocaleString('pt-BR');
  const fmt1=n=>Number(n||0).toLocaleString('pt-BR',{maximumFractionDigits:1});
  const boxes=[
    {cls:'blue',k:'abAgora',v:fmt(kpiRows.abAgora.length),l:'Abertos Agora'},
    {cls:'blue',k:'abH',v:fmt(kpiRows.abH.length),l:'Abertos Hoje'},
    {cls:'blue',k:'abO',v:fmt(kpiRows.abO.length),l:'Abertos Ontem'},
    {cls:'blue',k:'abA',v:fmt(kpiRows.abA.length),l:'Abertos Dias Ant.'},
    {cls:'purple',k:'naoAtrib',v:fmt(kpiRows.naoAtrib.length),l:'Não Atribuídos'},
    {cls:'red',k:'cliAf',v:fmt(sumKey(kpiRows.cliAf,'ct')),l:'Clientes Afetados (abertos)'},
    {cls:'red',k:'mais24',v:fmt(kpiRows.mais24.length),l:'Abertos +24h'},
    {cls:'orange',k:'feH',v:fmt(kpiRows.feH.length),l:'Fechados Hoje'},
    {cls:'orange',k:'feO',v:fmt(kpiRows.feO.length),l:'Fechados Ontem'},
    {cls:'',k:'tmaMed',v:fmt(meanKey(kpiRows.tmaMed,'tma')),l:'TMA Médio (fechados)'},
    {cls:'',k:'tmdMed',v:fmt(meanKey(fe,'tmd')),l:'TMD Médio (fechados)'},
    {cls:'',k:'tmeMed',v:fmt(meanKey(fe,'tme')),l:'TME Médio (fechados)'},
    {cls:'',k:'eqAtrib',v:fmt(dCount(kpiRows.eqAtrib,'eq')),l:'Equipes Atribuídas'},
    {cls:'',k:'eqExec',v:fmt(dCount(kpiRows.eqExec,'eqd')),l:'Equipes em Execução'},
    {cls:'',k:'eqTotDesl',v:fmt(dCount(kpiRows.eqTotDesl,'eqd')),l:'Equipes Totais Deslocadas'},
  ];
  if(tab==='d1'){
    const feHojeEq=feHoje.filter(r=>r.eo==='equipe'&&hasEqd(r));
    const den=dCount(feHojeEq,'eqd')||1;
    kpiRows.prod=feHojeEq;
    boxes.push({cls:'',k:'prod',v:fmt1(feHojeEq.length/den),l:'Produção Total'});
    ['A','B','C'].forEach(t=>{
      const rp=feHojeEq.filter(r=>r.ts===t);
      const d=dCount(rp,'eqd')||1;
      kpiRows['prod_t'+t]=rp;
      kpiTitles['prod_t'+t]='Produção Turno '+t+' (fechados hoje / equipes deslocadas)';
      boxes.push({cls:'',k:'prod_t'+t,v:fmt1(rp.length/d),l:'Produção Turno '+t});
    });
    ['EMERGÊNCIA','COMERCIAL','PERDAS','PODA','LINHA VIVA'].forEach(g=>{
      const rp=feHojeEq.filter(r=>r.gpd===g);
      const d=dCount(rp,'eqd')||1;
      kpiRows['prod_g_'+g]=rp;
      kpiTitles['prod_g_'+g]='Produção '+g+' (fechados hoje / equipes deslocadas)';
      boxes.push({cls:'',k:'prod_g_'+g,v:fmt1(rp.length/d),l:'Produção '+g});
    });
  }else{
    polos().forEach(p=>{
      const rp=feHoje.filter(r=>r.p===p&&r.eo==='equipe'&&hasEqd(r));
      const den=dCount(rp,'eqd')||1;
      kpiRows['prod_'+p]=rp;
      boxes.push({cls:'',k:'prod_'+p,v:fmt1(rp.length/den),l:'Produção '+p});
    });
  }
  const el=document.getElementById('k'+(tab==='d1'?'1':(tab==='d2'?'2':'3')));
  el.innerHTML=boxes.map(b=>`<div class="kpi ${b.cls}" data-k="${b.k}"><div class="v">${b.v}</div><div class="l">${b.l}</div></div>`).join('');
}

// ---- KPIs Dashboard 3 (visão de período / incidentes fechados) ----
function buildKPIsD3(){
  const U=UF['d3']||{};
  const base=RAW.filter(r=>passes(r,U));
  const fe=base.filter(r=>r.s==='Fechado');
  const ab=base;
  const tmaRows=fe.filter(r=>r.cg==='Considerar');
  const hasEqd=r=>{const v=String(r.eqd||'').trim();return v!==''&&v!=='---';};
  const dayKey=r=>((r.mes!==undefined&&r.mes!==null)?String(r.mes):'?')+'-'+((r.d!==undefined&&r.d!==null)?String(r.d):'?');
  const daysA={}; ab.forEach(r=>{daysA[dayKey(r)]=1;});
  const daysB={}; fe.forEach(r=>{daysB[dayKey(r)]=1;});
  const nA=Object.keys(daysA).length||1;
  const nF=Object.keys(daysB).length||1;
  const days={};
  fe.forEach(r=>{ if(!hasEqd(r)) return; const k=dayKey(r); if(!days[k]) days[k]=new Set(); days[k].add(r.eqd); });
  const dayVals=Object.keys(days).map(k=>days[k].size);
  const eqDiaMed=dayVals.length? dayVals.reduce((a,b)=>a+b,0)/dayVals.length : 0;
  const feEq=fe.filter(r=>r.eo==='equipe'&&hasEqd(r));
  kpiRows={
    d3_abertos:ab,
    d3_abertos_dia:ab,
    d3_cliAf:fe,
    d3_cliAf_dia:fe,
    d3_fechados:fe,
    d3_fechados_dia:fe,
    d3_tma:tmaRows,
    d3_tme:fe,
    d3_tmd:fe,
    d3_prod:feEq,
  };
  const fmt=n=>Number(n||0).toLocaleString('pt-BR');
  const fmt1=n=>Number(n||0).toLocaleString('pt-BR',{maximumFractionDigits:1});
  const cliAf=sumKey(fe,'a3');
  const eqDiaDiv=eqDiaMed>0?eqDiaMed*nF:1;
  const boxes=[
    {cls:'blue',k:'d3_abertos',v:fmt(ab.length),l:'Abertos (período)'},
    {cls:'blue',k:'d3_abertos_dia',v:fmt1(ab.length/nA),l:'Abertos (média/dia)'},
    {cls:'red',k:'d3_cliAf',v:fmt(cliAf),l:'Clientes Afetados (período)'},
    {cls:'red',k:'d3_cliAf_dia',v:fmt1(cliAf/nF),l:'Clientes Afetados (média/dia)'},
    {cls:'orange',k:'d3_fechados',v:fmt(fe.length),l:'Fechados (período)'},
    {cls:'orange',k:'d3_fechados_dia',v:fmt1(fe.length/nF),l:'Fechados (média/dia)'},
    {cls:'',k:'d3_tma',v:fmt(meanKey(tmaRows,'tma')),l:'TMA Médio'},
    {cls:'',k:'d3_tmd',v:fmt(meanKey(fe,'tmd')),l:'TMD Médio'},
    {cls:'',k:'d3_tme',v:fmt1((()=>{const dm={};fe.forEach(r=>{const k=dayKey(r);if(!dm[k])dm[k]=[];dm[k].push(Number(r.tme)||0);});const mns=Object.values(dm).map(a=>a.reduce((x,y)=>x+y,0)/a.length);return mns.length?mns.reduce((a,b)=>a+b,0)/mns.length:0;})()),l:'TME Médio'},
    {cls:'',k:'d3_eqDia',v:fmt1(eqDiaMed),l:'Média de Equipes Deslocadas'},
    {cls:'',k:'d3_prod',v:fmt1(feEq.length/eqDiaDiv),l:'Produção média dia por equipes deslocadas'},
  ];
  ['A','B','C'].forEach(t=>{
    const feT=fe.filter(r=>r.eo==='equipe'&&hasEqd(r)&&r.ts===t);
    const dt={};
    feT.forEach(r=>{const k=dayKey(r); if(!dt[k]) dt[k]=new Set(); dt[k].add(r.eqd);});
    const dv=Object.keys(dt).map(k=>dt[k].size);
    const eqDiaT=dv.length? dv.reduce((a,b)=>a+b,0)/dv.length : 0;
    kpiRows['d3_eqDia_t'+t]=feT;
    boxes.push({cls:'',k:'d3_eqDia_t'+t,v:fmt1(eqDiaT),l:'Média de Equipes Deslocadas Turno '+t});
    const rp=feEq.filter(r=>r.ts===t);
    const div=eqDiaT>0?eqDiaT*nF:1;
    kpiRows['d3_prod_t'+t]=rp;
    boxes.push({cls:'',k:'d3_prod_t'+t,v:fmt1(rp.length/div),l:'Produção Turno '+t});
  });
  ['EMERGÊNCIA','COMERCIAL','PERDAS','PODA','LINHA VIVA'].forEach(g=>{
    const rp=feEq.filter(r=>r.gpd===g);
    const d=dCount(rp,'eqd')||1;
    kpiRows['d3_prod_g_'+g]=rp;
    boxes.push({cls:'',k:'d3_prod_g_'+g,v:fmt1(rp.length/d),l:'Produção '+g});
  });
  const el=document.getElementById('k3');
  if(el) el.innerHTML=boxes.map(b=>`<div class="kpi ${b.cls}" data-k="${b.k}"><div class="v">${b.v}</div><div class="l">${b.l}</div></div>`).join('');
}

// ---- Top 10 ----
let topData=[];
let topMode={d1:'afet',d2:'afet',d3:'afet'};
function topTitleFor(t,m){const st=(t==='d3')?'fechados':'abertos';return m==='conh'?('Top 10 Cli x Tmp ('+st+')'):('Top 10 Afetações ('+st+')');}
function fmtCxt(v){return Number(v||0).toLocaleString('pt-BR',{maximumFractionDigits:2});}
function buildTop10(){
  const U=UF[tab]||{};
  const base=RAW.filter(r=>passes(r,U));
  const ab=(tab==='d3')?base.filter(r=>r.s==='Fechado'):base.filter(r=>r.s==='Aberto');
  const m=topMode[tab]||'afet';
  const key=m==='conh'?'cxt':(tab==='d3'?'a3':'ct');
  const totalCt=sumKey(ab,'ct')||1;
  const totalCxt=sumKey(ab,'cxt')||1;
  const total=key==='cxt'?totalCxt:sumKey(ab,key)||1;
  topData=ab.slice().sort((a,b)=>(Number(b[key])||0)-(Number(a[key])||0)).slice(0,10);
  const topSum=topData.reduce((a,r)=>a+(Number(r[key])/total*100),0);
  const topCt=topData.reduce((a,r)=>a+Number(r.ct),0);
  const topCxt=topData.reduce((a,r)=>a+Number(r.cxt||0),0);
  const topCaf=topData.reduce((a,r)=>a+Number(r.a3||0),0);
  const cliLabel=(tab==='d3')?'Cli Af 3 min Σ '+topCaf.toLocaleString('pt-BR'):'Clts Atual Σ '+topCt.toLocaleString('pt-BR');
  const cliVal=r=>(tab==='d3')?Number(r.a3||0).toLocaleString('pt-BR'):Number(r.ct).toLocaleString('pt-BR');
  const el=document.getElementById('top'+(tab==='d1'?'1':(tab==='d2'?'2':'3')));
  el.innerHTML='<table class="top"><tr><th>Número</th><th>NT</th><th>Polo</th><th>Alimentador</th><th>Causa</th><th>Atribuição</th><th>OSM</th><th class="num">'+cliLabel+'</th><th class="num">Cli x Tmp Σ '+fmtCxt(topCxt)+'</th><th class="num">% Σ '+topSum.toLocaleString('pt-BR',{maximumFractionDigits:1})+'%</th></tr>'+
    topData.map((r,i)=>`<tr data-i="${i}"><td>${esc(r.n)}</td><td>${esc(r.nt)}</td><td>${esc(r.p)}</td><td>${esc(r.al)}</td><td>${esc(r.ca)}</td><td>${esc(r.eq)}</td><td>${esc(r.osm)}</td><td class="num">${cliVal(r)}</td><td class="num">${fmtCxt(r.cxt)}</td><td class="num">${(Number(r[key])/total*100).toFixed(1)}%</td></tr>`).join('')+
    '</table>';
  el.querySelectorAll('tr[data-i]').forEach(tr=>{
    tr.style.cursor='pointer';
    tr.addEventListener('click',()=>{const r=topData[Number(tr.dataset.i)]; if(r) openDrillModal(r.n+' — detalhe',[r],null);});
  });
}
function topRows(t){
  const U=UF[t]||{};
  const base=RAW.filter(r=>passes(r,U));
  const ab=base.filter(r=>r.s==='Aberto');
  const m=topMode[t]||'afet';
  const key=m==='conh'?'cxt':(t==='d3'?'a3':'ct');
  const total=sumKey(ab,key)||1;
  const rows=ab.slice().sort((a,b)=>(Number(b[key])||0)-(Number(a[key])||0)).slice(0,10);
  return {rows,total};
}
function exportTop10(t){
  const {rows,total}=topRows(t);
  const m=topMode[t]||'afet';
  const key=m==='conh'?'cxt':(t==='d3'?'a3':'ct');
  const cliHeader=(t==='d3')?'Cli Af 3 min':'Clts Atual';
  const cols=['Número','NT','Polo','Alimentador','Causa','Atribuição','OSM',cliHeader,'Cli x Tmp','%','Status','Observação','Previsão'];
  const lines=[cols.map(csvCell).join(';')];
  rows.forEach(r=>{
    lines.push([r.n,r.nt,r.p,r.al,r.ca,r.eq,r.osm,Number(t==='d3'?r.a3||0:r.ct),Number(r.cxt||0),(Number(r[key])/total*100).toFixed(1)+'%','','',''].map(csvCell).join(';'));
  });
  downloadFile('top10_incidencias_'+t+'_'+m+'_'+new Date().toISOString().slice(0,10)+'.csv','\ufeff'+lines.join('\r\n'));
}
function setTopMode(t,m){
  topMode[t]=m;
  const seg=document.getElementById('topseg'+(t==='d1'?'1':(t==='d2'?'2':'3')));
  if(seg) seg.querySelectorAll('.tsbtn').forEach(b=>b.classList.toggle('active',b.dataset.tm===m));
  const title=document.getElementById('top-title'+(t==='d1'?'1':(t==='d2'?'2':'3')));
  if(title) title.textContent=topTitleFor(t,m);
  if(tab===t) buildTop10();
}

// ---- modal maximizar ----
let modalChart=null;
let modalOrig=null;
let modalDrag=null;
let box=null;
function modScaleX(){return modalChart&&modalChart.options&&modalChart.options.scales&&modalChart.options.scales.x;}
function modScaleY(){return modalChart&&modalChart.options&&modalChart.options.scales&&modalChart.options.scales.y;}
function captureOrig(){
  const o={};
  if(modalChart.scales.x) o.x={min:modalChart.scales.x.min,max:modalChart.scales.x.max};
  if(modalChart.scales.y) o.y={min:modalChart.scales.y.min,max:modalChart.scales.y.max};
  return o;
}
function setX(min,max){const s=modScaleX();if(!s)return;s.min=Math.round(min);s.max=Math.round(max);}
function setY(min,max){const s=modScaleY();if(!s)return;s.min=min;s.max=max;}
function clampAxis(axis,min,max,len){
  const o=modalOrig[axis];
  if(o){
    if(min<o.min){max+=o.min-min;min=o.min;}
    if(max>o.max){min-=max-o.max;max=o.max;}
    if(min<o.min)min=o.min;
    if(max>o.max)max=o.max;
    if(axis==='x'&&len&&max-min<len)max=min+len;
  }
  return {min:min,max:max};
}
function zoomAt(chart,factor,px,py){
  const s=chart.scales;
  if(s.x){
    const len=(s.x.max-s.x.min)||1;
    const frac=(px-(s.x.left||0))/(s.x.width||1);
    const nl=Math.max(0.5,len*factor);
    const r=clampAxis('x',s.x.min+(len-nl)*frac,s.x.min+(len-nl)*frac+nl,len);
    setX(r.min,r.max);
  }
  if(s.y){
    const len=(s.y.max-s.y.min)||1;
    const frac=1-((py-(s.y.top||0))/(s.y.height||1));
    const nl=Math.max(1e-6,len*factor);
    const r=clampAxis('y',s.y.min+(len-nl)*frac,s.y.min+(len-nl)*frac+nl);
    setY(r.min,r.max);
  }
  chart.update('none');
}
function modalWheel(ev){
  if(!modalChart)return;
  ev.preventDefault();
  const r=box.getBoundingClientRect();
  zoomAt(modalChart, ev.deltaY>0?1.12:1/1.12, ev.clientX-r.left, ev.clientY-r.top);
}
function modalDown(ev){
  if(!modalChart||ev.button!==0)return;
  modalDrag={x:ev.clientX,y:ev.clientY};
  document.addEventListener('pointermove',modalMove);
  document.addEventListener('pointerup',modalUp);
  document.addEventListener('pointercancel',modalUp);
}
function modalMove(ev){
  if(!modalDrag||!modalChart)return;
  const dx=ev.clientX-modalDrag.x;
  const dy=ev.clientY-modalDrag.y;
  modalDrag={x:ev.clientX,y:ev.clientY};
  const s=modalChart.scales;
  if(s.x){
    const len=(s.x.max-s.x.min)||1;
    const ratio=s.x.width/len||1;
    const r=clampAxis('x',s.x.min-dx/ratio,s.x.max-dx/ratio,len);
    setX(r.min,r.max);
  }
  if(s.y){
    const len=(s.y.max-s.y.min)||1;
    const ratio=s.y.height/len||1;
    const r=clampAxis('y',s.y.min+dy/ratio,s.y.max+dy/ratio);
    setY(r.min,r.max);
  }
  modalChart.update('none');
}
function modalUp(){modalDrag=null;document.removeEventListener('pointermove',modalMove);document.removeEventListener('pointerup',modalUp);document.removeEventListener('pointercancel',modalUp);}
function modalReset(){
  if(!modalChart||!modalOrig)return;
  if(modalOrig.x){const s=modScaleX();if(s){s.min=modalOrig.x.min;s.max=modalOrig.x.max;}}
  if(modalOrig.y){const s=modScaleY();if(s){s.min=modalOrig.y.min;s.max=modalOrig.y.max;}}
  modalChart.update('none');
}
function modalZoomBtn(zoomIn){
  if(!modalChart)return;
  const r=box.getBoundingClientRect();
  zoomAt(modalChart, zoomIn?1/1.4:1.4, r.width/2, r.height/2);
}
function openModal(id,title){
  const base=lastCfg[id];
  if(!base) return;
  document.getElementById('modal-title').textContent=title;
  box=document.getElementById('modal-canvas');
  const cfg=cloneCfg(base);
  cfg.options.layout.padding={top:40,right:70,left:24,bottom:14};
  if(cfg.options.scales){const sc={};for(const k in cfg.options.scales){sc[k]=JSON.parse(JSON.stringify(cfg.options.scales[k]));}cfg.options.scales=sc;}
  const stacked=!!(cfg.options.scales&&cfg.options.scales.x&&cfg.options.scales.x.stacked);
  cfg.options.plugins.totals=Object.assign({}, (base.options.plugins&&base.options.plugins.totals)||{enabled:stacked});
  if(cfg.type==='bar'||cfg.type==='line'){
    cfg.options.plugins.datalabels=Object.assign({}, cfg.options.plugins.datalabels, {font:{size:15,weight:'700'}, offset:8});
  } else {
    cfg.options.plugins.datalabels=cfg.options.plugins.datalabels||{};
  }
  modalChart=new Chart(box, cfg);
  const srcChart=charts[id];
  if(srcChart){
    if(srcChart.__entradaComp) modalChart.__entradaComp=srcChart.__entradaComp;
    if(srcChart.__horaTurno) modalChart.__horaTurno=srcChart.__horaTurno;
    if(srcChart.__totSingle) modalChart.__totSingle=true;
    if(srcChart.__countDrill) modalChart.__countDrill=srcChart.__countDrill;
  }
  box.onclick=ev=>handleChartClick(id, modalChart, ev);
  box.onmousemove=ev=>handleChartHover(id, modalChart, ev);
  box.onmouseleave=()=>{ const t=document.getElementById('ov-tip'); if(t) t.style.display='none'; if(box) box.style.cursor='default'; };
  document.getElementById('modal').style.display='flex';
  setTimeout(()=>{modalOrig=captureOrig();},80);
}
function closeModal(){
  modalUp();
  if(modalChart){modalChart.destroy();modalChart=null;}
  modalOrig=null;
  box.onclick=null;
  box.onmousemove=null;
  box.onmouseleave=null;
  document.getElementById('modal').style.display='none';
}
function addMaxBtns(){
  box=document.getElementById('modal-canvas');
  box.addEventListener('wheel',modalWheel,{passive:false});
  box.addEventListener('pointerdown',modalDown);
  document.querySelectorAll('.card').forEach(card=>{
    if(card.querySelector('.maxbtn')) return;
    const cnv=card.querySelector('canvas');
    if(!cnv) return;
    const b=document.createElement('button');
    b.className='maxbtn'; b.title='Maximizar'; b.innerHTML='&#x26F6;';
    b.onclick=()=>openModal(cnv.id, (card.querySelector('h3')||{}).textContent||cnv.id);
    card.appendChild(b);
  });
}

// ---- modal drill-down + export ----
let drillIdx=[];
let drillRows=[];
let drillMode='inc';
let drillField=null;
let drillTeam=null;
let drillTitleBase='';
function drillBodyHtml(rows){
  const head=FULL_COLS.map(l=>'<th>'+esc(l)+'</th>').join('');
  const MAX=500;
  const shown=rows.slice(0,MAX);
  const trs=shown.map(r=>{const f=FULL[IDXOF.get(r)];return '<tr>'+FULL_COLS.map((c,j)=>'<td>'+esc(f?f[j]:'')+'</td>').join('')+'</tr>';}).join('');
  return '<div class="dtwrap"><table class="dt"><tr>'+head+'</tr>'+(shown.length<rows.length?'<tr><td colspan="'+FULL_COLS.length+'" style="color:#5b6b7b">Mostrando '+shown.length+' de '+rows.length.toLocaleString('pt-BR')+' — use Exportar para obter todos.</td></tr>':'')+trs+'</table></div>';
}
function openDrillModal(title, rows, field){
  if(!rows||!rows.length) return;
  drillRows=rows;
  drillField=field||null;
  drillMode='inc';
  drillTeam=null;
  drillTitleBase=title;
  document.getElementById('drill-title').textContent=title;
  const tg=document.getElementById('drill-toggle');
  if(tg){
    tg.style.display=drillField?'':'none';
    tg.textContent='Ver Equipes';
  }
  renderDrill();
  document.getElementById('drill-modal').style.display='flex';
}
function renderDrill(){
  const body=document.getElementById('drill-body');
  const count=document.getElementById('drill-count');
  if(drillMode==='eq'&&drillField){
    renderDrillTeams(body);
    count.textContent=drillRows.length.toLocaleString('pt-BR')+' incidentes · clique numa equipe para ver os incidentes dela';
  }else{
    const rows=drillIncRows();
    body.innerHTML=drillBodyHtml(rows);
    let prefix='';
    if(drillField){const nEq=dCount(rows,drillField);prefix=nEq.toLocaleString('pt-BR')+(nEq===1?' equipe e ':' equipes e ');}
    count.textContent=prefix+rows.length.toLocaleString('pt-BR')+' incidentes · exportar para ver todos';
    drillIdx=rows.map(r=>IDXOF.get(r)).filter(i=>i!=null);
  }
}
function drillIncRows(){
  if(drillTeam==null) return drillRows;
  return drillRows.filter(r=>r[drillField]===drillTeam);
}
function renderDrillTeams(body){
  const seen=new Map();
  drillRows.forEach(r=>{
    const v=r[drillField];
    if(!v||v===''||v==='---') return;
    if(!seen.has(v)) seen.set(v,[]);
    seen.get(v).push(String(r.n||''));
  });
  const items=[...seen.entries()].sort((a,b)=>b[1].length-a[1].length);
  body.innerHTML='<div class="dtwrap"><table class="dt"><tr><th>Equipe</th><th class="num">Incidentes</th><th>Nºs</th></tr>'+
    items.map(([v,nums])=>'<tr class="team-row"><td>'+esc(v)+'</td><td class="num">'+nums.length.toLocaleString('pt-BR')+'</td><td>'+esc(nums.join(', '))+'</td></tr>').join('')+
    '</table></div>';
  body.querySelectorAll('.team-row').forEach((tr,i)=>{
    tr.addEventListener('click',()=>drillPickTeam(items[i][0]));
  });
}
function drillPickTeam(v){
  drillTeam=v;
  drillMode='inc';
  document.getElementById('drill-title').textContent=drillTitleBase+' — '+v;
  renderDrill();
}
function toggleDrillMode(){
  if(!drillField) return;
  const tg=document.getElementById('drill-toggle');
  if(drillMode==='inc'){
    drillMode='eq';
    drillTeam=null;
    tg.textContent='Ver Incidentes';
    document.getElementById('drill-title').textContent=drillTitleBase;
  }else{
    drillMode='inc';
    tg.textContent='Ver Equipes';
    document.getElementById('drill-title').textContent=drillTitleBase;
  }
  renderDrill();
}
function closeDrillModal(){
  drillIdx=[];
  drillRows=[];
  drillField=null;
  drillTeam=null;
  document.getElementById('drill-modal').style.display='none';
}
function csvCell(v){const s=String(v==null?'':v);return /[";\r\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;}
function downloadFile(name,text){
  const b=new Blob([text],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(b); a.download=name;
  document.body.appendChild(a); a.click();
  setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();},0);
}
function exportCSV(){
  if(!drillIdx.length) return;
  const lines=[FULL_COLS.join(';')];
  for(const i of drillIdx) lines.push(FULL[i].map(csvCell).join(';'));
  downloadFile('incidencias.csv','\ufeff'+lines.join('\r\n'));
}
function exportResumo(t){
  const tid=t||tab;
  const lines=[];
  lines.push('DataHub Dashboard de Incidências — Resumo consolidado');
  lines.push('Exportado em;'+new Date().toLocaleString('pt-BR'));
  lines.push('Aba;'+(tid==='d1'?'Dashboard (1 polo)':(tid==='d3'?'Dashboard (Mês/Dia)':'Dashboard 2 (todos os polos)')));
  lines.push('Período;'+metaDias());
  lines.push('Registros (base);'+RAW.filter(r=>passes(r,UF[tid]||{})).length);
  lines.push('');
  lines.push('KPIs;Valor');
  const kcont=document.getElementById('k'+(tid==='d1'?'1':(tid==='d2'?'2':'3')));
  kcont.querySelectorAll('.kpi').forEach(k=>{
    lines.push(csvCell(k.querySelector('.l').textContent.trim())+';'+csvCell(k.querySelector('.v').textContent.trim()));
  });
  document.getElementById(tid).querySelectorAll('.vcard').forEach(vc=>{
    const l=vc.querySelector('.l').textContent.trim();
    const v=vc.querySelector('.v').textContent.trim();
    lines.push(csvCell(l)+';'+csvCell(v));
  });
  lines.push('');
  Object.keys(charts).forEach(id=>{
    const ch=charts[id];
    const cv=document.getElementById(id);
    const card=cv?cv.closest('.card'):null;
    const title=card&&card.querySelector('h3')?card.querySelector('h3').textContent.trim():id;
    const names=ch.data.datasets.map(d=>d.label||'');
    const labels=ch.data.labels||[];
    lines.push('## '+id+' - '+title);
    lines.push(['Rótulo'].concat(names).concat(['Total']).map(csvCell).join(';'));
    const colTot=names.map(()=>0);
    labels.forEach((l,i)=>{
      const row=[l];
      let t=0;
      ch.data.datasets.forEach((d,j)=>{
        const v=typeof d.data[i]==='number'?d.data[i]:Number(d.data[i]||0);
        row.push(v); t+=v; colTot[j]+=v;
      });
      row.push(t);
      lines.push(row.map(csvCell).join(';'));
    });
    const row=['TOTAL'];
    let t=0;
    colTot.forEach(v=>{row.push(v);t+=v;});
    row.push(t);
    lines.push(row.map(csvCell).join(';'));
    lines.push('');
  });
  downloadFile('resumo_incidencias_'+tid+'_'+new Date().toISOString().slice(0,10)+'.csv','\ufeff'+lines.join('\r\n'));
}
function exportXLSX(){
  if(!drillIdx.length) return;
  if(typeof XLSX==='undefined'){alert('Biblioteca xlsx indisponível (offline) — exportando .csv.');exportCSV();return;}
  const objs=drillIdx.map(i=>{const o={};FULL_COLS.forEach((c,j)=>{o[c]=FULL[i][j];});return o;});
  const ws=XLSX.utils.json_to_sheet(objs);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Incidencias');
  XLSX.writeFile(wb,'incidencias.xlsx');
}
function searchOpen(q){
  const Q=String(q||'').trim();
  if(!Q) return;
  const byNum=RAW.filter(r=>String(r.n||'')===Q);
  if(byNum.length){openDrillModal('Busca — Incidente '+Q,byNum);return;}
  const L=Q.toLowerCase();
  const byTeam=RAW.filter(r=>{
    const a=String(r.eq||'').toLowerCase(), b=String(r.eqd||'').toLowerCase();
    return (a&&a.includes(L))||(b&&b.includes(L));
  });
  if(byTeam.length){openDrillModal('Busca — '+Q+' (equipe)',byTeam);return;}
  const byNumPart=RAW.filter(r=>String(r.n||'').includes(Q));
  if(byNumPart.length){openDrillModal('Busca — Nº contendo '+Q,byNumPart);return;}
  alert('Nada encontrado para "'+Q+'"');
}
function bindDrill(){
  ['k1','k2','k3'].forEach(id=>{
    document.getElementById(id).addEventListener('click',e=>{
      const el=e.target.closest('.kpi');
      if(!el) return;
      const k=el.dataset.k;
      const r=kpiRows[k];
      const fld=(k.indexOf('eq')>=0)?'eqd':(k.indexOf('prod')>=0?'eqd':(k==='eqAtrib'?'eq':null));
      if(r&&r.length) openDrillModal(kpiTitles[k]||k,r,fld);
    });
  });
  ['top1','top2','top3'].forEach(id=>{
    document.getElementById(id).addEventListener('click',e=>{
      const tr=e.target.closest('tr[data-i]');
      if(!tr) return;
      const r=topData[Number(tr.dataset.i)];
      if(r){const tt=(topTitle[topMode[tab]]||'Top 10').replace(/\s*\(abertos\)/,'');openDrillModal(tt+' — '+r.n,[r]);}
    });
  });
}

// ---- navegação e render ----
function sw(id){
  tab=id;
  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('active',b.dataset.tab===id));
  ['d1','d2','d3','inst'].forEach(p=>{const el=document.getElementById(p);if(el)el.classList.toggle('active',id===p);});
  if(id==='d1'||id==='d2'||id==='d3') renderPage();
}
function renderSub(sub){
  CUR_SUB=sub;
  const base=segBase(sub);
  const fns={e1:buildE1,s1:buildS1,m1:buildM1,bk1:buildBk1,o1:buildO1,x1:buildX1,x2:buildX2,x3:buildX3,e2:buildE2,s2:buildS2,o2:buildO2,t2:buildT2,p2:buildP2,xO:buildXO,e3:buildE3,s3:buildS3,m3:buildM1,bk3:buildBk3,o3:buildR3,o4:buildR4,t3:buildT3,x31:buildX1,xO3:buildXO,x32:buildX2,x33:buildX3};
  if(fns[sub]) fns[sub](base);
}
function renderPage(){
  if(tab==='d3') buildKPIsD3(); else buildKPIs();
  buildTop10();
  const segs=(tab==='d1')?['e1','s1','m1','bk1','o1','x1','xO','x2','x3']:(tab==='d3')?['e3','s3','m3','bk3','t3','o3','o4','x31','xO3','x32','x33']:['e2','s2','o2','t2','p2'];
  segs.forEach(sub=>renderSub(sub));
  const U=UF[tab]||{};
  document.getElementById('sub-info').textContent=(tab==='d1')?'Dashboard (1 polo)':(tab==='d3')?'Dashboard (Mês/Dia)':'Dashboard 2 (todos os polos)';
  document.getElementById('meta-info').textContent='Período: últimos 7 dias — Última atualização: '+BUILD_AT;
}
function bindSegBars(){
  ['segbar1','segbar2','segbar3'].forEach(id=>{
    document.getElementById(id).addEventListener('click',e=>{
      const b=e.target.closest('.segbtn');
      if(!b) return;
      const el=document.getElementById('seg-'+b.dataset.seg);
      if(el) el.scrollIntoView({behavior:'smooth',block:'start'});
    });
  });
}
document.addEventListener('keydown',e=>{if(e.key==='Escape'){closeModal();closeDrillModal();}});

// ---- print (html2canvas -> area de transferencia) ----
function downloadBlob(blob,name){
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;
  document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},0);
}
function flashBtn(btn,color){
  if(!btn) return;
  btn.style.background=color;btn.style.color='#fff';btn.style.borderColor=color;
  setTimeout(()=>{btn.style.background='';btn.style.color='';btn.style.borderColor='';},900);
}
function copyCanvas(canvas,label,btn){
  canvas.toBlob(blob=>{
    if(!blob){flashBtn(btn,'#d1242f');return;}
    const ok=()=>flashBtn(btn,'#1a7f37');
    const fb=()=>{downloadBlob(blob,'print.png');flashBtn(btn,'#d1242f');};
    if(navigator.clipboard&&window.ClipboardItem){
      try{navigator.clipboard.write([new ClipboardItem({'image/png':blob})]).then(ok).catch(fb);}catch(e){fb();}
    }else{fb();}
  },'image/png');
}
function printEl(el,label,btn){
  if(!el) return;
  if(typeof html2canvas==='undefined'){flashBtn(btn,'#d1242f');return;}
  html2canvas(el,{backgroundColor:'#ffffff',scale:2,logging:false,useCORS:true}).then(canvas=>copyCanvas(canvas,label,btn)).catch(()=>flashBtn(btn,'#d1242f'));
}
function printSegment(sb,btn){
  const grid=sb.querySelector('.grid');
  const head=sb.querySelector('.seghead');
  let title='Segmento';
  if(head){const cl=head.cloneNode(true);const pb=cl.querySelector('.print-btn');if(pb)pb.remove();title=(cl.textContent||'').replace(/\s+/g,' ').trim();}
  printEl(grid||sb,title,btn);
}
function printEntradaSaida(entSb,btn){
  const saidaSb=document.getElementById(entSb.id.replace('seg-e','seg-s'));
  const entGrid=entSb.querySelector('.grid');
  const saidaGrid=saidaSb?saidaSb.querySelector('.grid'):null;
  const els=[entGrid,saidaGrid].filter(Boolean);
  if(!els.length){flashBtn(btn,'#d1242f');return;}
  if(typeof html2canvas==='undefined'){flashBtn(btn,'#d1242f');return;}
  Promise.all(els.map(el=>html2canvas(el,{backgroundColor:'#ffffff',scale:2,logging:false,useCORS:true}))).then(canvases=>{
    const gap=14;
    const w=Math.max.apply(null,canvases.map(c=>c.width));
    const h=canvases.reduce((a,c)=>a+c.height,0)+gap*(canvases.length-1);
    const out=document.createElement('canvas');out.width=w;out.height=h;
    const ctx=out.getContext('2d');ctx.fillStyle='#ffffff';ctx.fillRect(0,0,w,h);
    let y=0;canvases.forEach(c=>{ctx.drawImage(c,0,y);y+=c.height+gap;});
    copyCanvas(out,'Entrada + Saída',btn);
  }).catch(()=>flashBtn(btn,'#d1242f'));
}
function addPrintBtns(){
  document.querySelectorAll('.segblock').forEach(sb=>{
    const head=sb.querySelector('.seghead');
    if(!head||head.querySelector('.print-btn')) return;
    if(/^seg-e\d+$/.test(sb.id)){
      const b2=document.createElement('button');
      b2.type='button'; b2.className='print-btn'; b2.title='Copiar print dos gráficos de Entrada + Saída'; b2.innerHTML='&#128424; Ent+Saída';
      b2.onclick=e=>{e.stopPropagation();printEntradaSaida(sb,b2);};
      head.appendChild(b2);
    }
    const b=document.createElement('button');
    b.type='button'; b.className='print-btn'; b.title='Copiar print dos gráficos do segmento'; b.innerHTML='&#128424; Print';
    b.onclick=e=>{e.stopPropagation();printSegment(sb,b);};
    head.appendChild(b);
  });
  [['f-top1','resumo1'],['f-top2','resumo2']].forEach(pair=>{
    const fid=pair[0], rid=pair[1];
    const cont=document.getElementById(fid);
    if(!cont||cont.querySelector('.print-btn')) return;
    const b=document.createElement('button');
    b.type='button'; b.className='print-btn'; b.title='Copiar print dos KPIs + Top 10'; b.innerHTML='&#128424; Print';
    b.onclick=e=>{e.stopPropagation();printEl(document.getElementById(rid),'KPIs + Top 10');};
    cont.appendChild(b);
  });
}
// ---- download da base ORIGINAL filtrada (Reincidentes / Improdutivo) ----
function addExportBtns(){
  const map={
    'seg-o3':{fn:isPrim, fname:'reincidentes_base.csv'},
    'seg-o4':{fn:r=>String(r.pd||'').toLowerCase().trim()==='improdutivo', fname:'improdutivo_base.csv'},
  };
  Object.keys(map).forEach(segId=>{
    const sb=document.getElementById(segId);
    if(!sb) return;
    const head=sb.querySelector('.seghead');
    if(!head||head.querySelector('.export-base-btn')) return;
    const b=document.createElement('button');
    b.type='button'; b.className='print-btn export-base-btn'; b.title='Baixar a base original filtrada por este segmento (CSV)';
    b.innerHTML='&#11015; Base';
    b.onclick=e=>{e.stopPropagation();
      const lines=[FULL_COLS.join(';')];
      RAW.forEach((r,i)=>{ if(map[segId].fn(r)) lines.push(FULL[i].map(csvCell).join(';')); });
      if(lines.length>1) downloadFile(map[segId].fname,'\ufeff'+lines.join('\r\n'));
      flashBtn(b,'#1a7f37');
    };
    const pb=head.querySelector('.print-btn');
    if(pb) head.insertBefore(b,pb); else head.appendChild(b);
  });
}
// ---- vcards clicaveis (+24h, Ordem 2, 5 Regras) ----
let vcardRows={};
const VCARD_TITLES={'x24-sim':'+24h — Abertos','x24-nao':'+24h — Fechados Hoje','x2-sim':'Ordem 2 — Fechados Hoje (sim)','x2-nao':'Ordem 2 — Fechados Hoje (não)','x3-sim':'5 Regras — Fechados Hoje MT (S)','x3-nao':'5 Regras — Fechados Hoje MT (N)','d3-x24-sim':'+24h — Fechados >24h (duração)','d3-x24-nao':'+24h — Fechados ≤24h (duração)','d3-xO-ent':'Entradas OSM','d3-xO-sai':'Saídas OSM','d3-x2-sim':'Ordem 2 — Fechados (sim)','d3-x2-nao':'Ordem 2 — Fechados (não)','d3-x3-sim':'5 Regras — Fechados MT (S)','d3-x3-nao':'5 Regras — Fechados MT (N)'};
function bindVcards(){
  document.querySelectorAll('.vcard').forEach(vc=>{
    const vEl=vc.querySelector('.v');
    if(!vEl||vc.__bound) return;
    vc.__bound=true;
    vc.style.cursor='pointer';
    vc.addEventListener('click',()=>{
      const id=vEl.id;
      const key=id.replace(/^d3-/,'');
      const rows=vcardRows[key]||vcardRows[id];
      if(rows&&rows.length) openDrillModal(VCARD_TITLES[id]||VCARD_TITLES[key]||id, rows);
    });
  });
}

// ---- init ----
buildTopFilters('d1');
buildTopFilters('d2');
buildTopFilters('d3');
Object.keys(SUB_CONF).forEach(sub=>buildFilters(sub));
bindSegBars();
bindDrill();
addMaxBtns();
addPrintBtns();
addExportBtns();
bindVcards();
sw('d1');
</script>
</body>
</html>

"""


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Gera dashboard HTML de incidências.")
    ap.add_argument("arquivo", help="Caminho do xlsx tratado")
    ap.add_argument("saida", nargs="?", default=None, help="Arquivo HTML de saída (opcional)")
    args = ap.parse_args()

    df = pd.read_excel(args.arquivo)
    saida = args.saida or (Path(args.arquivo).parent / "dashboard_incidencias.html")
    gerar_dashboard_html(df, Path(saida))

