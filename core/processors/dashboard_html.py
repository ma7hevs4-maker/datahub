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

        turno_a = _norm(val(c_turno))
        turno_s = _norm(val(c_turno_s))
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
        regras = _norm(val(c_regras))
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
.kpi.blue .v{color:var(--acc)} .kpi.red .v{color:#d1242f} .kpi.orange .v{color:#f0820f} .kpi.purple .v{color:#8250df}
.kpi .l{color:var(--mut);font-size:11px}

.topbox{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:12px 14px;margin-bottom:14px;box-shadow:0 1px 2px rgba(16,40,73,.06)}
.tophead{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.tophead h4{margin:0;font-size:13px;color:var(--acc)}
.tophead .export-btn{margin-left:auto;padding:6px 12px}
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
.maxbtn{position:absolute;top:10px;right:10px;width:28px;height:28px;line-height:24px;text-align:center;
  background:var(--card2);border:1px solid var(--border);border-radius:6px;color:var(--mut);cursor:pointer;font-size:14px}
.maxbtn:hover{background:var(--acc2);color:var(--acc);border-color:var(--acc)}
.btn{border:1px solid var(--border);border-radius:6px;padding:6px 12px;font-size:12px;font-weight:600;cursor:pointer;background:var(--card2);color:var(--txt)}
.btn:hover{background:var(--acc2);color:var(--acc);border-color:var(--acc)}
.modal{display:none;position:fixed;inset:0;z-index:1000;background:rgba(15,25,40,.6);padding:12px;align-items:center;justify-content:center}
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
.print-btn{margin-left:8px;background:var(--card2);border:1px solid var(--border);border-radius:6px;padding:4px 10px;font-size:12px;font-weight:600;cursor:pointer;color:var(--txt)}
.print-btn:hover{background:var(--acc2);color:var(--acc);border-color:var(--acc)}
.search-inp{width:150px;padding:5px 9px;border:1px solid var(--border);border-radius:6px;font-size:12.5px;font-family:inherit;outline:none}
.search-btn{background:#fff;border:1px solid var(--border);border-radius:6px;cursor:pointer;padding:4px 8px;font-size:13px;line-height:1;color:#333}
.search-btn:hover{background:var(--acc2)}
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
  <button class="tab" data-tab="d3" onclick="sw('d3')">Dashboard 3 (Mês/Dia)</button>
  <button class="tab" data-tab="inst" onclick="sw('inst')">Regras</button>
</div>

<!-- ============================= ABA 1 ============================= -->
<div class="panel active" id="d1">
  <div class="segbar" id="segbar1">
    <button class="segbtn" data-seg="e1"><span class="dot" style="background:#1f6feb"></span>Entrada</button>
    <button class="segbtn" data-seg="s1"><span class="dot" style="background:#f0820f"></span>Saída</button>
    <button class="segbtn" data-seg="m1"><span class="dot" style="background:#5f6b7a"></span>Entrada &amp; Saída</button>
    <button class="segbtn" data-seg="o1"><span class="dot" style="background:#1a7f37"></span>Em Aberto</button>
    <button class="segbtn" data-seg="x1"><span class="dot" style="background:#d1242f"></span>+24h</button>
    <button class="segbtn" data-seg="xO"><span class="dot" style="background:#0969da"></span>OSM</button>
    <button class="segbtn" data-seg="x2"><span class="dot" style="background:#8250df"></span>Ordem 2</button>
    <button class="segbtn" data-seg="x3"><span class="dot" style="background:#0aa2a2"></span>5 Regras</button>
  </div>

  <div class="filters" id="f-top1"></div>

  <div id="resumo1"><div class="kpis" id="k1"></div>
  <div class="topbox"><div class="tophead"><h4>Top 10 Afetações (abertos)</h4><button type="button" class="export-btn" onclick="exportTop10('d1')">⬇ Exportar Top 10</button></div><div id="top1"></div></div></div>

  <section class="segblock" id="seg-e1">
    <div class="seghead"><span class="dot" style="background:#1f6feb"></span>Entrada<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
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
    <div class="seghead"><span class="dot" style="background:#f0820f"></span>Saída<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-s1"></div>
    <div class="grid">
      <div class="card"><h3>Saída por Hora — Equipe × Operador (fechados)</h3><canvas id="s1-1"></canvas></div>
      <div class="card"><h3>Saída por Turno — Equipe × Operador</h3><canvas id="s1-2"></canvas></div>
      <div class="card"><h3>Saída por Hora — por categoria de Saída</h3><canvas id="s1-3"></canvas></div>
      <div class="card"><h3>Atribuição de Incidentes por Entrada (abertos)</h3><canvas id="s1-4"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-m1">
    <div class="seghead"><span class="dot" style="background:#5f6b7a"></span>Entrada &amp; Saída<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-m1"></div>
    <div class="grid">
      <div class="card"><h3>Entrada × Saída por Turno</h3><canvas id="m1-1"></canvas></div>
      <div class="card"><h3>Entrada × Saída por Hora</h3><canvas id="m1-2"></canvas></div>
      <div class="card"><h3>Entrada × Saída por Hora — Clientes (Clts Atual)</h3><canvas id="m1-3"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-xO">
    <div class="seghead"><span class="dot" style="background:#0969da"></span>OSM<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-xO"></div>
    <div class="grid">
      <div class="card"><h3>Entrada × Saída por Hora — OSM (Sim)</h3><canvas id="xO-1"></canvas></div>
      <div class="card"><h3>OSM — Sim × Não</h3><div class="stack"><div class="card vcard"><div class="v" id="xO-sim">0</div><div class="l">Sim</div></div><div class="card vcard"><div class="v" id="xO-nao">0</div><div class="l">Não</div></div></div></div>
    </div>
  </section>

  <section class="segblock" id="seg-o1">
    <div class="seghead"><span class="dot" style="background:#1a7f37"></span>Em Aberto<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
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
    <div class="seghead"><span class="dot" style="background:#d1242f"></span>+24h<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-x1"></div>
    <div class="grid">
      <div class="card"><h3>+24h — Abertos × Fechados Hoje</h3><div class="stack"><div class="card vcard"><div class="v" id="x24-sim">0</div><div class="l">Abertos</div></div><div class="card vcard"><div class="v" id="x24-nao">0</div><div class="l">Fechados Hoje</div></div></div></div>
      <div class="card"><h3>+24h Ontem — Aberto × Fechado Hoje</h3><canvas id="x1b"></canvas></div>
      <div class="card"><h3>+24h Dias Anteriores — Aberto × Fechado Hoje</h3><canvas id="x1c"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-x2">
    <div class="seghead"><span class="dot" style="background:#8250df"></span>Ordem 2<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-x2"></div>
    <div class="grid">
      <div class="card"><h3>Ordem 2 — Fechados Hoje (sim × não)</h3><div class="stack"><div class="card vcard"><div class="v" id="x2-sim">0</div><div class="l">Sim</div></div><div class="card vcard"><div class="v" id="x2-nao">0</div><div class="l">Não</div></div></div></div>
      <div class="card"><h3>Ordem 2 — Fechados Hoje (sim × não)</h3><canvas id="x1-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-x3">
    <div class="seghead"><span class="dot" style="background:#0aa2a2"></span>5 Regras<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
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
  <div class="topbox"><div class="tophead"><h4>Top 10 Afetações (abertos)</h4><button type="button" class="export-btn" onclick="exportTop10('d2')">⬇ Exportar Top 10</button></div><div id="top2"></div></div></div>

  <section class="segblock" id="seg-e2">
    <div class="seghead"><span class="dot" style="background:#1f6feb"></span>Entrada<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-e2"></div>
    <div class="grid">
      <div class="card"><h3>Entrada Total por Polo — Incidente × Reincidente</h3><canvas id="e2-1"></canvas></div>
      <div class="card"><h3>Entrada por Hora — por Polo</h3><canvas id="e2-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-s2">
    <div class="seghead"><span class="dot" style="background:#f0820f"></span>Saída<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-s2"></div>
    <div class="grid">
      <div class="card"><h3>Saída Total por Polo — Equipe × Operador (fechados)</h3><canvas id="s2-1"></canvas></div>
      <div class="card"><h3>Saída por Hora — por Polo (fechados)</h3><canvas id="s2-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-o2">
    <div class="seghead"><span class="dot" style="background:#1a7f37"></span>Em Aberto<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-o2"></div>
    <div class="grid">
      <div class="card"><h3>Abertas por Polo × Cluster — Atribuição</h3><canvas id="o2-1"></canvas></div>
      <div class="card"><h3>Clientes em Aberto por Polo × Cluster</h3><canvas id="o2-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-t2">
    <div class="seghead"><span class="dot" style="background:#8250df"></span>TMA<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-t2"></div>
    <div class="grid">
      <div class="card"><h3>Fechados considerados p/ TMA — por Polo</h3><canvas id="t2-1"></canvas></div>
      <div class="card"><h3>Média TMA — considerados</h3><canvas id="t2-3"></canvas></div>
      <div class="card"><h3>Fechados por Equipe (Equipe Desl.)</h3><canvas id="t2-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-p2">
    <div class="seghead"><span class="dot" style="background:#0aa2a2"></span>Processos<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-p2"></div>
    <div class="grid">
      <div class="card"><h3>Atribuição por Polo × Entrada</h3><canvas id="p2-1"></canvas></div>
      <div class="card"><h3>Atribuição +24h por Polo × Entrada</h3><canvas id="p2-2"></canvas></div>
      <div class="card"><h3>Atribuído por Processo — Abertos</h3><canvas id="p2-3"></canvas></div>
      <div class="card"><h3>Improdutivo por Processo — Fechados</h3><canvas id="p2-4"></canvas></div>
      <div class="card"><h3>Fechados por Processo</h3><canvas id="p2-5"></canvas></div>
    </div>
  </section>
</div>

<!-- ============================= ABA 3 ============================= -->
<div class="panel" id="d3">
  <div class="segbar" id="segbar3">
    <button class="segbtn" data-seg="e3"><span class="dot" style="background:#1f6feb"></span>Entrada</button>
    <button class="segbtn" data-seg="s3"><span class="dot" style="background:#f0820f"></span>Saída</button>
    <button class="segbtn" data-seg="m3"><span class="dot" style="background:#5f6b7a"></span>Entrada &amp; Saída</button>
    <button class="segbtn" data-seg="o3"><span class="dot" style="background:#1a7f37"></span>Em Aberto</button>
    <button class="segbtn" data-seg="x31"><span class="dot" style="background:#d1242f"></span>+24h</button>
    <button class="segbtn" data-seg="xO3"><span class="dot" style="background:#0969da"></span>OSM</button>
    <button class="segbtn" data-seg="x32"><span class="dot" style="background:#8250df"></span>Ordem 2</button>
    <button class="segbtn" data-seg="x33"><span class="dot" style="background:#0aa2a2"></span>5 Regras</button>
  </div>

  <div class="filters" id="f-top3"></div>

  <div id="resumo3"><div class="kpis" id="k3"></div>
  <div class="topbox"><div class="tophead"><h4>Top 10 Afetações (abertos)</h4><button type="button" class="export-btn" onclick="exportTop10('d3')">⬇ Exportar Top 10</button></div><div id="top3"></div></div></div>

  <section class="segblock" id="seg-e3">
    <div class="seghead"><span class="dot" style="background:#1f6feb"></span>Entrada<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-e3"></div>
    <div class="grid">
      <div class="card"><h3>Entrada por Hora — Incidente × Reincidente</h3><canvas id="d3-e1-1"></canvas></div>
      <div class="card"><h3>Entrada por Turno — Incidente × Reincidente</h3><canvas id="d3-e1-2"></canvas></div>
      <div class="card"><h3>Análise de Reincidentes — Equipe × Operador</h3><canvas id="d3-e1-3"></canvas></div>
      <div class="card"><h3>Primeira Causa de Reincidência (ranking)</h3><canvas id="d3-e1-4"></canvas></div>
      <div class="card"><h3>Top Equipes com Reincidência</h3><canvas id="d3-e1-5"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-s3">
    <div class="seghead"><span class="dot" style="background:#f0820f"></span>Saída<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-s3"></div>
    <div class="grid">
      <div class="card"><h3>Saída por Hora — Equipe × Operador (fechados)</h3><canvas id="d3-s1-1"></canvas></div>
      <div class="card"><h3>Saída por Turno — Equipe × Operador</h3><canvas id="d3-s1-2"></canvas></div>
      <div class="card"><h3>Saída por Hora — por categoria de Saída</h3><canvas id="d3-s1-3"></canvas></div>
      <div class="card"><h3>Atribuição de Incidentes por Entrada (abertos)</h3><canvas id="d3-s1-4"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-m3">
    <div class="seghead"><span class="dot" style="background:#5f6b7a"></span>Entrada &amp; Saída<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-m3"></div>
    <div class="grid">
      <div class="card"><h3>Entrada × Saída por Turno</h3><canvas id="d3-m1-1"></canvas></div>
      <div class="card"><h3>Entrada × Saída por Hora</h3><canvas id="d3-m1-2"></canvas></div>
      <div class="card"><h3>Entrada × Saída por Hora — Clientes (Clts Atual)</h3><canvas id="d3-m1-3"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-o3">
    <div class="seghead"><span class="dot" style="background:#1a7f37"></span>Em Aberto<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-o3"></div>
    <div class="grid g5">
      <div class="card"><h3>Abertas Hoje — BT</h3><canvas id="d3-o1-1"></canvas></div>
      <div class="card"><h3>Abertas Ontem — BT</h3><canvas id="d3-o1-2"></canvas></div>
      <div class="card"><h3>Abertas Dias Anteriores — BT</h3><canvas id="d3-o1-3"></canvas></div>
      <div class="card span2x2"><h3>Clientes em Aberto (Clts Atual)</h3><canvas id="d3-o1-7"></canvas></div>
      <div class="card"><h3>Abertas Hoje — MT</h3><canvas id="d3-o1-4"></canvas></div>
      <div class="card"><h3>Abertas Ontem — MT</h3><canvas id="d3-o1-5"></canvas></div>
      <div class="card"><h3>Abertas Dias Anteriores — MT</h3><canvas id="d3-o1-6"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-x31">
    <div class="seghead"><span class="dot" style="background:#d1242f"></span>+24h<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-x31"></div>
    <div class="grid">
      <div class="card"><h3>+24h — Abertos × Fechados Hoje</h3><div class="stack"><div class="card vcard"><div class="v" id="d3-x24-sim">0</div><div class="l">Abertos</div></div><div class="card vcard"><div class="v" id="d3-x24-nao">0</div><div class="l">Fechados Hoje</div></div></div></div>
      <div class="card"><h3>+24h Ontem — Aberto × Fechado Hoje</h3><canvas id="d3-x1b"></canvas></div>
      <div class="card"><h3>+24h Dias Anteriores — Aberto × Fechado Hoje</h3><canvas id="d3-x1c"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-xO3">
    <div class="seghead"><span class="dot" style="background:#0969da"></span>OSM<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-xO3"></div>
    <div class="grid">
      <div class="card"><h3>Entrada × Saída por Hora — OSM (Sim)</h3><canvas id="d3-xO-1"></canvas></div>
      <div class="card"><h3>OSM — Sim × Não</h3><div class="stack"><div class="card vcard"><div class="v" id="d3-xO-sim">0</div><div class="l">Sim</div></div><div class="card vcard"><div class="v" id="d3-xO-nao">0</div><div class="l">Não</div></div></div></div>
    </div>
  </section>

  <section class="segblock" id="seg-x32">
    <div class="seghead"><span class="dot" style="background:#8250df"></span>Ordem 2<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-x32"></div>
    <div class="grid">
      <div class="card"><h3>Ordem 2 — Fechados Hoje (sim × não)</h3><div class="stack"><div class="card vcard"><div class="v" id="d3-x2-sim">0</div><div class="l">Sim</div></div><div class="card vcard"><div class="v" id="d3-x2-nao">0</div><div class="l">Não</div></div></div></div>
      <div class="card"><h3>Ordem 2 — Fechados Hoje (sim × não)</h3><canvas id="d3-x1-2"></canvas></div>
    </div>
  </section>

  <section class="segblock" id="seg-x33">
    <div class="seghead"><span class="dot" style="background:#0aa2a2"></span>5 Regras<span class="block-hint">duplo clique nas barras/KPIs/Top 10 para detalhar</span></div>
    <div class="filters" id="f-x33"></div>
    <div class="grid">
      <div class="card"><h3>5 Regras de Ouro — Fechados Hoje MT (S × N)</h3><div class="stack"><div class="card vcard"><div class="v" id="d3-x3-sim">0</div><div class="l">S</div></div><div class="card vcard"><div class="v" id="d3-x3-nao">0</div><div class="l">N</div></div></div></div>
      <div class="card"><h3>5 Regras de Ouro — Fechados Hoje MT (S × N)</h3><canvas id="d3-x1-3"></canvas></div>
    </div>
  </section>
</div>

<!-- ============================= ABA INSTRUÇÕES ============================= -->
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
    <div class="modal-hint">🖱️ Scroll: zoom&nbsp;&nbsp;•&nbsp;&nbsp;Arrastar: mover o gráfico&nbsp;&nbsp;•&nbsp;&nbsp;+ / − / ⟲: controle de zoom&nbsp;&nbsp;•&nbsp;&nbsp;Duplo clique: detalhar</div>
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
let PID='d1';
let CUR_SUB='';
const BUILD_AT='__BUILD_AT__';

// ---- filtros universais (UT, Polo, NT) ----
const UF={d1:{},d2:{},d3:{}};
function passes(r,F){
  if(F.p&&!F.p.all&&!F.p.sel.has(r.p)) return false;
  if(F.suc&&!F.suc.all&&!F.suc.sel.has(r.suc)) return false;
  if(F.e&&!F.e.all&&!F.e.sel.has(r.e)) return false;
  if(F.sd&&!F.sd.all&&!F.sd.sel.has(r.sd)) return false;
  if(F.ta&&!F.ta.all&&!F.ta.sel.has(r.ta)) return false;
  if(F.ts&&!F.ts.all&&!F.ts.sel.has(r.ts)) return false;
  if(F.nt&&!F.nt.all&&!F.nt.sel.has(r.nt)) return false;
  if(F.gpd&&!F.gpd.all&&!F.gpd.sel.has(r.gpd)) return false;
  if(F.mes&&!F.mes.all&&!F.mes.sel.has(r.mes)) return false;
  if(F.dia&&!F.dia.all&&!F.dia.sel.has(r.dia)) return false;
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
  d1:[['ut','UT','ut',['UTS','UTN']],['p','Polo','p',polos()],['suc','Sucursal','suc',()=>sucPool()],['nt','NT','nt',nts()],['gpd','Grupo Processos (Eq. Desl.)','gpd',()=>distinct('gpd')]],
  d2:[['ut','UT','ut',['UTS','UTN']],['p','Polo','p',polos()],['suc','Sucursal','suc',()=>sucPool()],['nt','NT','nt',nts()],['gpd','Grupo Processos (Eq. Desl.)','gpd',()=>distinct('gpd')]],
  d3:[['mes','Mês','mes',()=>distinct('mes')],['dia','Dia','dia',()=>distinct('dia')]],
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
const SUB_CONF={
  e1:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{e:['Hoje']}},
  s1:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{sd:['Hoje']}},
  m1:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()]],defs:{e:['Hoje'],sd:['Hoje']}},
  o1:{keys:[],defs:{}},
  x1:{keys:[],defs:{}},
  xO:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()]],defs:{e:['Hoje'],sd:['Hoje']}},
  x2:{keys:[],defs:{}},
  x3:{keys:[],defs:{}},
  e2:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{e:['Hoje']}},
  s2:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{sd:['Hoje']}},
  o2:{keys:[['e','Entrada','e',ents()]],defs:{}},
  t2:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()],['ta','Turno Entrada','ta',TURN],['ts','Turno Fechamento','ts',TURN]],defs:{}},
  p2:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()],['ta','Turno Entrada','ta',TURN],['ts','Turno Fechamento','ts',TURN]],defs:{}},
  e3:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{e:['Hoje']}},
  s3:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()],['ta','Turno Entrada','ta',TURN],['ts','Turno Saída','ts',TURN]],defs:{sd:['Hoje']}},
  m3:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()]],defs:{e:['Hoje'],sd:['Hoje']}},
  o3:{keys:[],defs:{}},
  x31:{keys:[],defs:{}},
  xO3:{keys:[['e','Entrada','e',ents()],['sd','Saída','sd',sds()]],defs:{e:['Hoje'],sd:['Hoje']}},
  x32:{keys:[],defs:{}},
  x33:{keys:[],defs:{}},
};
const FSTATE={};
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
function segBase(sub){
  const U=UF[tab]||{};
  const F=FSTATE[sub]||{};
  const C={ut:U.ut,p:U.p,suc:U.suc,nt:U.nt,gpd:U.gpd,mes:U.mes,dia:U.dia,e:F.e,sd:F.sd,ta:F.ta,ts:F.ts};
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
  cfg.options.plugins.legend = cfg.options.plugins.legend||{position:'bottom',labels:{boxWidth:12,font:{size:11},color:'#5b6b7b'}};
  cfg.options.plugins.tooltip = {callbacks:{label:c=>{const v=c.parsed.y??c.parsed.x??c.parsed;return ' '+v;}}};
  cfg.options.layout = cfg.options.layout||{};
  cfg.options.layout.padding = Object.assign({top:26,right:40}, cfg.options.layout.padding||{});
  const el=document.getElementById(id);
  el.style.height='290px';
  el.ondblclick=ev=>handleChartDbl(id, charts[id], ev);
  el.onclick=ev=>handleTotClick(id, charts[id], ev);
  lastCfg[id]=cloneCfg(cfg);
  charts[id]=new Chart(el, cfg);
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
    backgroundColor:c=>{const ds=c.dataset&&c.dataset.backgroundColor;const col=Array.isArray(ds)?ds[c.dataIndex]:ds;return col||'#5b6b7b';},
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
  const cat=isH?chart.scales.y:chart.scales.x;
  const val=isH?chart.scales.x:chart.scales.y;
  if(!cat||!val) return;
  const labels=chart.data.labels||[];
  const meta=chart.getDatasetMeta(0);
  if(!meta||!meta.data||!meta.data.length) return;
  const ctx=chart.ctx;
  const ds=chart.data.datasets;
  const n=Math.min(labels.length,meta.data.length);
  chart.__totRects=[];
  ctx.save();
  ctx.textAlign='center';
  ctx.textBaseline='middle';
  for(let i=0;i<n;i++){
    let tot=0;
    for(const d of ds){tot+=Number(d.data[i])||0;}
    if(tot<=0) continue;
    const txt=tot>=1000?String(Math.round(tot/100)/10)+'k':String(tot);
    ctx.font='700 11px "Segoe UI",Arial';
    const tw=ctx.measureText(txt).width;
    const px=isH?val.getPixelForValue(tot):cat.getPixelForValue(labels[i]);
    const py=isH?cat.getPixelForValue(labels[i]):val.getPixelForValue(tot);
    const cy=isH?py:py-15;
    const x1=px-tw/2-3,y1=cy-8,x2=px+tw/2+3,y2=cy+8;
    chart.__totRects.push({x1:x1,y1:y1,x2:x2,y2:y2,i:i});
    ctx.fillStyle='#ffffff';
    ctx.fillRect(x1,y1,tw+6,16);
    ctx.strokeStyle='#d8e0ea';
    ctx.lineWidth=1;
    ctx.strokeRect(x1,y1,tw+6,16);
    ctx.fillStyle='#1f2d3d';
    ctx.fillText(txt,px,cy);
  }
  ctx.restore();
}};
Chart.register(totalsPlugin);

// ---- renderChart genérico + drill ----
let drillSets={};
function storeDrill(id, labels, fn){
  drillSets[id]=[];
  for(let i=0;i<labels.length;i++) drillSets[id][i]=fn(i);
}
function chartTitle(id){
  const el=document.getElementById(id);
  const card=el&&el.closest('.card');
  return (card&&card.querySelector('h3').textContent)||id;
}
function handleChartDbl(id, chart, ev){
  if(!chart) return;
  const isLine=chart.config&&chart.config.type==='line';
  const items=chart.getElementsAtEventForMode(ev,'nearest',isLine?{intersect:false,maxDistance:40}:{intersect:true},false);
  if(!items.length) return;
  const it=items[0];
  const label=(chart.data.labels||[])[it.index];
  const rows=(drillSets[id]&&drillSets[id][it.index])?drillSets[id][it.index][it.datasetIndex]:null;
  if(!rows||!rows.length) return;
  openDrillModal(chartTitle(id)+(label!=null?' — '+label:''), rows);
}
function handleTotClick(id, chart, ev){
  if(!chart) return;
  const cnv=chart.canvas;
  const rect=cnv.getBoundingClientRect();
  const x=(ev.clientX-rect.left)*(cnv.width/rect.width);
  const y=(ev.clientY-rect.top)*(cnv.height/rect.height);
  if(chart.__totRects&&chart.__totRects.length){
    const hit=chart.__totRects.find(b=>x>=b.x1&&x<=b.x2&&y>=b.y1&&y<=b.y2);
    if(hit){
      const sets=(drillSets[id]&&drillSets[id][hit.i])?drillSets[id][hit.i]:null;
      if(sets&&sets.length){
        const rows=sets.reduce((a,b)=>a.concat(b),[]);
        const label=(chart.data.labels||[])[hit.i];
        if(rows.length) openDrillModal(chartTitle(id)+(label!=null?' — '+label:''), rows);
      }
      return;
    }
  }
  const ex=chart.$datalabels&&chart.$datalabels._labels;
  if(ex){
    for(let i=ex.length-1;i>=0;i--){
      const l=ex[i];
      if(l.$layout&&l.$layout._visible&&l.$layout._box.contains({x:x,y:y})){
        const rows=(drillSets[id]&&drillSets[id][l._index])?drillSets[id][l._index][l.$groups._set]:null;
        if(rows&&rows.length) openDrillModal(chartTitle(id)+' — '+(chart.data.labels||[])[l._index], rows);
        return;
      }
    }
  }
}
function renderChart(id, labels, series, o){
  if(PID==='d3') id='d3-'+id;
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
  const cfg={type:o.type||'bar',data:{labels:labels,datasets:datasets},options:{plugins:{datalabels:isLine?dlabels(false):dins(),totals:{enabled:!!o.stacked}}}};
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
  const ca=tally(reinc.filter(r=>!['---','-'].includes(r.ca)),['ca']);
  const caK=Object.keys(ca).sort((a,b)=>ca[b]-ca[a]).slice(0,15);
  renderChart('e1-4',caK,[{name:'Qtd',color:BLUE[0],rows:k=>reinc.filter(r=>r.ca===k)}],{indexAxis:'y'});
  const eqm=tally(reinc.filter(r=>!['---','-','','SEM EQUIPE','GENERICA'].includes(r.eqd)),['eqd']);
  const eqK=Object.keys(eqm).sort((a,b)=>eqm[b]-eqm[a]).slice(0,10);
  const re4=REIN.filter(v=>reinc.some(r=>r.rt===v));
  renderChart('e1-5',eqK,re4.map((v,i)=>({name:v,color:BLUE[i%BLUE.length],rows:eq=>reinc.filter(r=>r.eqd===eq&&r.rt===v)})),{});
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
  const C4={ut:U.ut,p:U.p,suc:U.suc,nt:U.nt,gpd:U.gpd,e:F.e,sd:{all:false,sel:new Set(['Aberto'])},ta:F.ta,ts:F.ts};
  const ab4=RAW.filter(r=>passes(r,C4));
  renderChart('s1-4',ENTR_ORD,[
    {name:'Atribuído',color:ORANGE[0],rows:e=>ab4.filter(r=>r.s==='Aberto'&&!isIgnCausa(r)&&r.e===e&&r.atr==='Atribuído')},
    {name:'Não Atribuído',color:ORANGE[1],rows:e=>ab4.filter(r=>r.s==='Aberto'&&!isIgnCausa(r)&&r.e===e&&r.atr==='Não Atribuído')},
  ],{stacked:true});
}

function buildM1(base){
  renderChart('m1-1',TURN,[
    {name:'Entrada',color:'#1f6feb',rows:t=>base.filter(r=>r.ta===t)},
    {name:'Saída',color:'#f0820f',rows:t=>base.filter(r=>r.s==='Fechado'&&r.ts===t)},
  ],{});
  renderChart('m1-2',HOUR_LAB,[
    {name:'Entrada',color:'#1f6feb',area:'rgba(31,111,235,.12)',rows:i=>base.filter(r=>parseInt(r.he,10)===parseInt(i,10))},
    {name:'Saída',color:'#d1242f',area:'rgba(209,36,47,.12)',rows:i=>base.filter(r=>r.s==='Fechado').filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
    {name:'Saída Equipe',color:'#1a7f37',area:'rgba(26,127,55,.12)',rows:i=>base.filter(r=>r.s==='Fechado'&&r.eo==='equipe').filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
  ],{type:'line'});
  renderChart('m1-3',HOUR_LAB,[
    {name:'Entrada (clientes)',color:'#1f6feb',area:'rgba(31,111,235,.12)',wt:'ct',rows:i=>base.filter(r=>parseInt(r.he,10)===parseInt(i,10))},
    {name:'Saída (Cli Af 3 min)',color:'#d1242f',area:'rgba(209,36,47,.12)',wt:'a3',rows:i=>base.filter(r=>r.s==='Fechado').filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
    {name:'Saída Equipe (Cli Af 3 min)',color:'#1a7f37',area:'rgba(26,127,55,.12)',wt:'a3',rows:i=>base.filter(r=>r.s==='Fechado'&&r.eo==='equipe').filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
  ],{type:'line'});
}

function buildXO(base){
  const osm=base.filter(r=>String(r.osm).toUpperCase()==='SIM');
  renderChart('xO-1',HOUR_LAB,[
    {name:'Entrada OSM',color:'#1f6feb',area:'rgba(31,111,235,.12)',rows:i=>osm.filter(r=>parseInt(r.he,10)===parseInt(i,10))},
    {name:'Saída OSM',color:'#d1242f',area:'rgba(209,36,47,.12)',rows:i=>osm.filter(r=>r.s==='Fechado').filter(r=>parseInt(r.hs,10)===parseInt(i,10))},
  ],{type:'line'});
  setVCard('xO-sim',base.filter(r=>String(r.osm).toUpperCase()==='SIM').length);
  setVCard('xO-nao',base.filter(r=>String(r.osm).toUpperCase()==='NÃO').length);
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

function fillKpis(id,items){
  const el=document.getElementById(id);
  if(!el) return;
  el.innerHTML=items.map(b=>`<div class="kpi ${b.cls}"><div class="v">${b.v}</div><div class="l">${b.l}</div></div>`).join('');
}
const fmtN=n=>Number(n||0).toLocaleString('pt-BR');
function setVCard(id,v){
  if(PID==='d3') id='d3-'+id;
  const el=document.getElementById(id);
  if(el) el.textContent=fmtN(v);
}
function buildX1(base){
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
  setVCard('x24-sim',p24.filter(r=>r.s==='Aberto').length);
  setVCard('x24-nao',p24.filter(r=>r.s==='Fechado'&&r.sd==='Hoje').length);
  mk('x1b','Ontem'); mk('x1c','Dias Anteriores');
}
function buildX2(base){
  const fe=base.filter(r=>r.s==='Fechado'&&r.sd==='Hoje');
  const lab=ENTR_ORD.filter(v=>fe.some(r=>r.e===v));
  renderChart('x1-2',lab,[
    {name:'sim',color:'#1a7f37',rows:e=>fe.filter(r=>r.e===e&&r.o==='sim')},
    {name:'não',color:'#d1242f',rows:e=>fe.filter(r=>r.e===e&&r.o==='não')},
  ],{stacked:true});
  setVCard('x2-sim',fe.filter(r=>r.o==='sim').length);
  setVCard('x2-nao',fe.filter(r=>r.o==='não').length);
}
function buildX3(base){
  const fe=base.filter(r=>r.s==='Fechado'&&r.sd==='Hoje');
  const mt=fe.filter(r=>r.t==='MT'&&(r.ro==='S'||r.ro==='N'));
  const lab=comboKeys(mt,['e','nt']);
  renderChart('x1-3',lab,[
    {name:'S',color:'#1a7f37',rows:l=>mt.filter(r=>r.e+' · '+r.nt===l&&r.ro==='S')},
    {name:'N',color:'#d1242f',rows:l=>mt.filter(r=>r.e+' · '+r.nt===l&&r.ro==='N')},
  ],{stacked:true});
  setVCard('x3-sim',mt.filter(r=>r.ro==='S').length);
  setVCard('x3-nao',mt.filter(r=>r.ro==='N').length);
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
function buildT2(base){
  const cons=base.filter(r=>r.s==='Fechado'&&r.cg==='Considerar');
  const pl=plOf(base);
  renderChart('t2-1',pl,[{name:'Qtd',color:'#1f6feb',rows:p=>cons.filter(r=>r.p===p)}],{});
  const fe=base.filter(r=>r.s==='Fechado');
  const eqm=tally(fe.filter(r=>!['---','-','','SEM EQUIPE','GENERICA'].includes(r.eqd)),['eqd']);
  const eqK=Object.keys(eqm).sort((a,b)=>eqm[b]-eqm[a]).slice(0,20);
  renderChart('t2-2',eqK,[{name:'Qtd',color:'#f0820f',rows:k=>fe.filter(r=>r.eqd===k&&!['---','-','','SEM EQUIPE','GENERICA'].includes(r.eqd))}],{indexAxis:'y'});
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
  const imp=base.filter(r=>r.s==='Fechado'&&r.pd==='Improdutivo'&&r.eo==='equipe');
  const gpdI=[...new Set(imp.map(r=>r.gpd))].filter(v=>v&&v!=='').sort();
  renderChart('p2-4',pl,gpdI.map((g,i)=>({name:g,color:PALETTE[i%PALETTE.length],rows:p=>imp.filter(r=>r.p===p&&r.gpd===g)})),{});
  const feEq=base.filter(r=>r.s==='Fechado'&&r.eo==='equipe');
  const gpdF=[...new Set(feEq.map(r=>r.gpd))].filter(v=>v&&v!=='').sort();
  renderChart('p2-5',pl,gpdF.map((g,i)=>({name:g,color:PALETTE[i%PALETTE.length],rows:p=>feEq.filter(r=>r.p===p&&r.gpd===g)})),{});
}

// ---- KPIs ----
const KPI_TITLES={abH:'Abertos Hoje',abO:'Abertos Ontem',abA:'Abertos Dias Anteriores',feH:'Fechados Hoje',feO:'Fechados Ontem',naoAtrib:'Não Atribuídos',abAgora:'Abertos Agora',mais24:'Abertos +24h',cliAf:'Clientes Afetados (abertos)',tmaMed:'TMA Médio (fechados)',tmdMed:'TMD Médio (fechados)',eqAtrib:'Equipes Atribuídas (entrada hoje)',eqDesl:'Equipes Deslocadas (entrada hoje)',prod:'Produção Total (fechados hoje / equipes deslocadas)',tmeMed:'TME Médio — TME Análise=sim'};
let kpiRows={};
let kpiTitles=Object.assign({},KPI_TITLES);
function buildKPIs(){
  const U=UF[tab]||{};
  const base=RAW.filter(r=>passes(r,U));
  const ab=base.filter(r=>r.s==='Aberto');
  const fe=base.filter(r=>r.s==='Fechado');
  const hoje=base.filter(r=>r.e==='Hoje');
  const feHoje=fe.filter(r=>r.sd==='Hoje');
  const dCount=(rows,key)=>{const s=new Set(rows.map(r=>r[key]).filter(v=>v&&v!==''&&v!=='---'&&v!=='-'));return s.size;};
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
    tmaMed:fe,
    eqAtrib:hoje,
    eqDesl:hoje,
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
    {cls:'',k:'eqDesl',v:fmt(dCount(kpiRows.eqDesl,'eqd')),l:'Equipes Deslocadas'},
  ];
  if(tab==='d1'){
    const feHojeEq=feHoje.filter(r=>r.eo==='equipe');
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
      const rp=feHoje.filter(r=>r.p===p&&r.eo==='equipe');
      const den=dCount(rp,'eqd')||1;
      kpiRows['prod_'+p]=rp;
      boxes.push({cls:'',k:'prod_'+p,v:fmt1(rp.length/den),l:'Produção '+p});
    });
  }
  const el=document.getElementById(tab==='d3'?'k3':(tab==='d1'?'k1':'k2'));
  el.innerHTML=boxes.map(b=>`<div class="kpi ${b.cls}" data-k="${b.k}"><div class="v">${b.v}</div><div class="l">${b.l}</div></div>`).join('');
}

// ---- Top 10 ----
let topData=[];
function buildTop10(){
  const U=UF[tab]||{};
  const base=RAW.filter(r=>passes(r,U));
  const ab=base.filter(r=>r.s==='Aberto');
  const totalCt=sumKey(ab,'ct')||1;
  topData=ab.slice().sort((a,b)=>(Number(b.ct)||0)-(Number(a.ct)||0)).slice(0,10);
  const topSum=topData.reduce((a,r)=>a+(Number(r.ct)/totalCt*100),0);
  const topClts=topData.reduce((a,r)=>a+Number(r.ct),0);
  const el=document.getElementById(tab==='d3'?'top3':(tab==='d1'?'top1':'top2'));
  el.innerHTML='<table class="top"><tr><th>Número</th><th>NT</th><th>Polo</th><th>Alimentador</th><th>Causa</th><th>Atribuição</th><th>OSM</th><th class="num">Clts Atual Σ '+topClts.toLocaleString('pt-BR')+'</th><th class="num">% Σ '+topSum.toLocaleString('pt-BR',{maximumFractionDigits:1})+'%</th></tr>'+
    topData.map((r,i)=>`<tr data-i="${i}"><td>${esc(r.n)}</td><td>${esc(r.nt)}</td><td>${esc(r.p)}</td><td>${esc(r.al)}</td><td>${esc(r.ca)}</td><td>${esc(r.eq)}</td><td>${esc(r.osm)}</td><td class="num">${Number(r.ct).toLocaleString('pt-BR')}</td><td class="num">${(Number(r.ct)/totalCt*100).toFixed(1)}%</td></tr>`).join('')+
    '</table>';
}
function topRows(t){
  const U=UF[t]||{};
  const base=RAW.filter(r=>passes(r,U));
  const ab=base.filter(r=>r.s==='Aberto');
  const totalCt=sumKey(ab,'ct')||1;
  const rows=ab.slice().sort((a,b)=>(Number(b.ct)||0)-(Number(a.ct)||0)).slice(0,10);
  return {rows,totalCt};
}
function exportTop10(t){
  const {rows,totalCt}=topRows(t);
  const cols=['Número','NT','Polo','Alimentador','Causa','Atribuição','OSM','Clts Atual','%','Status','Observação','Previsão'];
  const lines=[cols.map(csvCell).join(';')];
  rows.forEach(r=>{
    lines.push([r.n,r.nt,r.p,r.al,r.ca,r.eq,r.osm,Number(r.ct),(Number(r.ct)/totalCt*100).toFixed(1)+'%','','',''].map(csvCell).join(';'));
  });
  downloadFile('top10_incidencias_'+t+'_'+new Date().toISOString().slice(0,10)+'.csv','\ufeff'+lines.join('\r\n'));
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
function handleModalDbl(id, ev){
  if(!modalChart) return;
  const isLine=modalChart.config&&modalChart.config.type==='line';
  const items=modalChart.getElementsAtEventForMode(ev,'nearest',isLine?{intersect:false,maxDistance:40}:{intersect:true},false);
  if(!items.length) return;
  const it=items[0];
  const label=(modalChart.data.labels||[])[it.index];
  const rows=(drillSets[id]&&drillSets[id][it.index])?drillSets[id][it.index][it.datasetIndex]:null;
  if(rows&&rows.length) openDrillModal(document.getElementById('modal-title').textContent+(label!=null?' — '+label:''), rows);
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
  cfg.options.plugins.totals={enabled:stacked};
  if(cfg.type==='bar'||cfg.type==='line'){
    cfg.options.plugins.datalabels=Object.assign({}, cfg.options.plugins.datalabels, {font:{size:15,weight:'700'}, offset:8});
  } else {
    cfg.options.plugins.datalabels=cfg.options.plugins.datalabels||{};
  }
  modalChart=new Chart(box, cfg);
  box.ondblclick=ev=>handleModalDbl(id, ev);
  box.onclick=ev=>handleTotClick(id, modalChart, ev);
  document.getElementById('modal').style.display='flex';
  setTimeout(()=>{modalOrig=captureOrig();},80);
}
function closeModal(){
  modalUp();
  if(modalChart){modalChart.destroy();modalChart=null;}
  modalOrig=null;
  box.ondblclick=null;
  box.onclick=null;
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
    count.textContent=rows.length.toLocaleString('pt-BR')+' incidentes · exportar para ver todos';
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
    if(!v||v===''||v==='---'||v==='-') return;
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
  lines.push('Aba;'+(tid==='d1'?'Dashboard (1 polo)':(tid==='d3')?'Dashboard 3 (Mês/Dia)':'Dashboard 2 (todos os polos)'));
  lines.push('Período;'+metaDias());
  lines.push('Registros (base);'+RAW.filter(r=>passes(r,UF[tid]||{})).length);
  lines.push('');
  lines.push('KPIs;Valor');
  const kcont=document.getElementById(tid==='d3'?'k3':(tid==='d1'?'k1':'k2'));
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
function bindDrill(){
  ['k1','k2'].forEach(id=>{
    document.getElementById(id).addEventListener('dblclick',e=>{
      const el=e.target.closest('.kpi');
      if(!el) return;
      const k=el.dataset.k;
      const r=kpiRows[k];
      const fld=k==='eqAtrib'?'eq':k==='eqDesl'?'eqd':k.indexOf('prod')===0?'eqd':null;
      if(r&&r.length) openDrillModal(kpiTitles[k]||(k.indexOf('prod_')===0?'Produção '+k.slice(5):k),r,fld);
    });
  });
  ['top1','top2'].forEach(id=>{
    document.getElementById(id).addEventListener('dblclick',e=>{
      const tr=e.target.closest('tr[data-i]');
      if(!tr) return;
      const r=topData[Number(tr.dataset.i)];
      if(r) openDrillModal('Top 10 Afetações — '+r.n,[r]);
    });
  });
}

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

// ---- navegação e render ----
function sw(id){
  tab=id;
  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('active',b.dataset.tab===id));
  document.getElementById('d1').classList.toggle('active',id==='d1');
  document.getElementById('d2').classList.toggle('active',id==='d2');
  document.getElementById('d3').classList.toggle('active',id==='d3');
  document.getElementById('inst').classList.toggle('active',id==='inst');
  renderPage();
}
function renderSub(sub){
  CUR_SUB=sub;
  const base=segBase(sub);
  const fns={e1:buildE1,s1:buildS1,m1:buildM1,o1:buildO1,x1:buildX1,xO:buildXO,x2:buildX2,x3:buildX3,e2:buildE2,s2:buildS2,o2:buildO2,t2:buildT2,p2:buildP2,e3:buildE1,s3:buildS1,m3:buildM1,o3:buildO1,x31:buildX1,xO3:buildXO,x32:buildX2,x33:buildX3};
  if(fns[sub]) fns[sub](base);
}
function renderPage(){
  if(tab==='inst'){
    document.getElementById('sub-info').textContent='Regras — Conceitos e lógicas dos KPIs e gráficos';
    document.getElementById('meta-info').textContent='';
    return;
  }
  PID=(tab==='d3')?'d3':'d1';
  buildKPIs();
  buildTop10();
  const segs=(tab==='d1')?['e1','s1','m1','o1','x1','xO','x2','x3']:(tab==='d3')?['e3','s3','m3','o3','x31','xO3','x32','x33']:['e2','s2','o2','t2','p2'];
  segs.forEach(sub=>renderSub(sub));
  const U=UF[tab]||{};
  document.getElementById('sub-info').textContent=(tab==='d1')?'Dashboard (1 polo)':(tab==='d3')?'Dashboard 3 (Mês/Dia)':'Dashboard 2 (todos os polos)';
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

// ---- init ----
buildTopFilters('d1');
buildTopFilters('d2');
buildTopFilters('d3');
Object.keys(SUB_CONF).forEach(sub=>buildFilters(sub));
bindSegBars();
bindDrill();
  addMaxBtns();
  addPrintBtns();
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

