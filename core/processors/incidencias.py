"""
Tratamento do relatório de Incidências do GeoOnline.
Baseado em Tratamento/tratamento.py, adaptado para receber Path e log_fn.
"""
import json
import urllib.request
import warnings
import unicodedata
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

SHEET_NAME = "Relatorio_de_Incidencias"


def _sem_acento(texto: str) -> str:
    return unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")


# ── Mapas de causa ─────────────────────────────────────────────────────────────

MAPA_CONSIDERAR = {_sem_acento(k): v for k, v in {
    "ATEND CANCELADO CR": False, "ATENDIMENTO REMOTO": False, "BORNE FUNDIDO": True,
    "CASA FECHADA": True, "DEF INT COM CORREÇÃO COELCE": True, "DEFEITO EM CONEXAO": True,
    "DEFEITO EM CONEXAO DE MEDIDOR": True, "DEFEITO EM MEDIDOR": True,
    "DEFEITO EM RAMAL DE LIGAÇÃO": True, "DEFEITO INTERNO CLIENTE": True,
    "DEFEITO RAMAL DE LIGAÇÃO CABO AL": True, "DEFEITO TEMPORÁRIO NÃO IDENTIFICADO": True,
    "DEGRADAÇÃO MATERIAL": True, "ESTAVA NORMAL": True, "GALHO/FOLHA": True,
    "INSTANTANEA": False, "JUMP PARTIDO": True, "MANUTENÇAO ENTORNO": False,
    "MEDIDOR QUEIMADO SUBSTITUÍDO": True, "PADRAO QUEBRADO": True, "PIPA": True,
    "PROGRAMADA MANUTENÇÃO": False, "RAMAL DE LIGAÇÃO SUBSTITUÍDO": True,
    "TESTE REMOTO": False, "TRANSFERÊNCIA E RETRANSFERÊNCIA DE CARGAS": False,
    "VEGETAÇÃO": True, "DEFEITO EM MEDIDOR RAMAL CONCÊNTRICO": True,
    "INCIDENCIA SEM AFETAÇÃO": True, "PÁSSARO": True,
    "DEFEITO EM CONEXAO RAMAL CONCENTRICO": True, "DEFEITO EM TRANSFORMADOR": True,
    "ANIMAIS": True, "ÁRVORE TOMBADA": True, "NIVEL DE TENSÃO NORMAL": False,
    "CONDUTOR PARTIDO": True, "SOBRECARGA": True, "NOVAS INSTALAÇÕES": False,
    "CONDUTOR TRANÇADO": True, "PROGRAMADA URGENTE": False,
    "REGISTRO INDEVIDO DA RECLAMAÇÃO": False, "PROGRAMADA OBRAS": False,
    "DEFEITO EM CONCENTRADOR - REDE DAT": True, "RAMAL BAMBO": True,
    "ENDEREÇO NÃO LOCALIZADO": True, "DEFEITO EM CONEXÃO": True, "ROUBO": True,
    "CORPO ESTRANHO": True, "CS COM MAL CONTATO": True, "DEFEITO EM CHAVE FUSIVEL": True,
    "DEFEITO EM MEDIÇÃO SELADA": True, "DESLIGAMENTO EMERGENCIAL": True,
    "GRANDE CLIENTE DEFEITO INTERNO": True, "LUZ CORTADA": False,
    "OSCILAÇÃO": False, "OSCILAÇÃO PROVOCADA POR TERCEIROS": False,
    "OUTRAS CAUSAS DE TERCEIROS": True, "POSTE ABALROADO": True,
    "RESELAGEM QUADRO": True, "TERCEIROS ACIDENTAL": True,
    "TRABALHO PROGRAMADO": False, "DEFEITO EM CONEXAO RAMAL CONVECIONAL": True,
    "VANDALISMO": True, "FALHA DE AJUSTE DA PROTEÇÃO": True,
    "DEFEITO EM RAMAL CONCÊNTRICO": True, "CONDUTOR BAMBO": True,
    "INCÊNDIO": True, "DESCARGAS ATMOSFÉRICAS": True,
    "DEFEITO EM IP COM CORREÇAO": True, "DEF EM CONEXAO MED RAMAL CONCÊNTRICO": True,
    "CHUVA": True, "DESLG/RELIGAMENTO POR SOLICITAÇÃO DO CLIENTE": False,
    "CS QUEIMADO": True, "VERIFICAÇÃO NIVEL DE TENSAO": False,
    "DEFEITO EM ISOLADOR": True, "PROGRAMADA PARA MUDANÇA DE TAPS": False,
    "RECLAMAÇÃO NIVEL DE TENSAO": False, "RELIGAÇÃO NORMAL": False,
    "RELIGAÇÃO ESPECIAL": False, "DEFEITO EM IP SEM CORREÇÃO": True,
    "CORTE INDEVIDO": False, "DEFEITO NO DISJUNTOR NO TRAFO": True,
    "ERRO NA MONTAGEM OU INSTALAÇÃO": True, "ATENDIMENTO COMERCIAL": True,
    "AREA RISCO": True, "MARESIA": True, "DEFEITO EM RELIGADOR": True,
    "DEFEITO EM TP": True, "TRANSFERENCIA DE CLIENTE DEFINITIVA": True,
    "POLUIÇÃO": True, "CONDUTOR SOLTO DO ISOLADOR": True,
    "DESLIGAMENTO PROGRAMADO": True, "SUSPEITA DE FRAUDE": True,
    "FALHA DE MANOBRA": True, "DEFEITO EM DISJUNTOR": True,
    "FALHA EQUIPAMENTO PROTEÇÃO": True, "ERRO DE OPERAÇÃO CC": True,
    "ENTREGA DE VIATURA": True, "UC FECHADA": True,
    "FALHA EQ. CONTROLE, REGULAÇÃO OU COMUNICAÇÃO": True, "INUNDAÇÃO": True,
    "DEFEITO EM CONEXAO RAMAL SUBTERRANEO": True, "ELO INADEQUADO": True,
    "DIA CRITICO INTERNO INT": False, "CALAMIDADE PUBLICA": False,
    "DIA CRITICO INTERNO PROG": False, "ATUAÇÃO ESQUEMA ALIVIO CARGA": False,
    "QUEIMADAS": True,
}.items()}

MAPA_TIPO = {_sem_acento(k): "Remoto" for k in [
    "ATEND CANCELADO CR", "ATENDIMENTO REMOTO", "INSTANTANEA",
    "PROGRAMADA MANUTENÇÃO", "TRANSFERÊNCIA E RETRANSFERÊNCIA DE CARGAS",
    "PROGRAMADA URGENTE", "PROGRAMADA OBRAS", "NOVAS INSTALAÇÕES",
    "ATUAÇÃO ESQUEMA ALIVIO CARGA", "TRABALHO PROGRAMADO",
]}

CAUSAS_IMPRODUTIVO = [_sem_acento(c) for c in [
    "CASA FECHADA", "DEFEITO INTERNO CLIENTE", "ENDEREÇO NÃO LOCALIZADO",
    "ESTAVA NORMAL", "GRANDE CLIENTE DEFEITO INTERNO", "INCIDENCIA SEM AFETAÇÃO",
    "LUZ CORTADA", "NIVEL DE TENSÃO NORMAL", "OSCILAÇÃO",
    "OSCILAÇÃO PROVOCADA POR TERCEIROS", "OUTRAS CAUSAS DE TERCEIROS",
    "REGISTRO INDEVIDO DA RECLAMAÇÃO", "UC FECHADA",
]]

MAPA_SUCURSAL = {
    "PTI": "Paraty", "PTM": "Paraty", "SRO": "Paraty", "TAT": "Paraty", "MAM02": "Paraty",
    "MUR": "Mangaratiba", "JAC05": "Mangaratiba",
    "ANG": "Angra dos Reis", "JAC01": "Angra dos Reis", "JAC02": "Angra dos Reis",
    "JAC03": "Angra dos Reis", "JAC04": "Angra dos Reis", "FRADE-ANG": "Angra dos Reis",
    "ITO": "Angra dos Reis", "MAM01": "Angra dos Reis", "PR.BRAVA": "Angra dos Reis",
    "V.RESID": "Angra dos Reis",
}

MAPA_PROCESSO = {
    "A": "AMPLA CHIP", "R": "CORTE E RELIGAÇÃO", "L": "NOVAS LIGAÇÕES",
    "F": "GESTOR NOVAS LIGAÇÕES", "P": "PERDAS", "M": "MANUTENÇÃO",
    "D": "PODA", "V": "LINHA VIVA", "T": "TELECONTROLE",
    "B": "OBRAS", "O": "OPERAÇÃO", "X": "EXTRA OPERAÇÃO",
}

MAPA_GRUPO = {
    "PODA": "PODA", "MANUTENÇÃO": "PODA",
    "PERDAS": "PERDAS", "AMPLA CHIP": "PERDAS",
    "OPERAÇÃO": "EMERGÊNCIA", "EXTRA OPERAÇÃO": "EMERGÊNCIA",
    "EQUIPE GENÉRICA": "EMERGÊNCIA", "SEM EQUIPE": "EMERGÊNCIA",
    "CORTE E RELIGAÇÃO": "COMERCIAL", "NOVAS LIGAÇÕES": "COMERCIAL",
    "GESTOR NOVAS LIGAÇÕES": "COMERCIAL",
    "LINHA VIVA": "LINHA VIVA",
}


def _converter_duracao(valor) -> int:
    s = str(valor).strip()
    if s in ("", "---", "nan", "None", "0"):
        return 0
    if ":" in s:
        partes = s.split(":")
        try:
            return int(partes[0]) * 60 + int(partes[1])
        except Exception:
            return 0
    try:
        return int(float(s))
    except Exception:
        return 0


def _validar_prefixo(cod: str) -> str:
    c = str(cod).strip().upper()
    if c in ("", "NAN", "NONE", "---", "-", "0", "NULL"):
        return "SEM EQUIPE"
    if len(c) != 10 or c[8] != "-" or not c[2:5].isdigit():
        return "GENERICA"
    return c


def _extrair_processo(cod: str) -> str:
    p = _validar_prefixo(cod)
    if p == "SEM EQUIPE":
        return "SEM EQUIPE"
    if p == "GENERICA":
        return "EQUIPE GENÉRICA"
    letra = p[6] if len(p) > 6 else ""
    return MAPA_PROCESSO.get(letra, "EQUIPE GENÉRICA")


def _extrair_enel(cod: str) -> str:
    p = _validar_prefixo(cod)
    if p == "SEM EQUIPE":
        return "Sem equipe"
    if p == "GENERICA":
        return "EQUIPE GENÉRICA"
    letra = p[7] if len(p) > 7 else ""
    if letra == "E":
        return "Enel"
    if letra == "P":
        return "Parceira"
    return "Sem equipe"


def _grupo_processos(proc: str) -> str:
    return MAPA_GRUPO.get(str(proc).strip().upper(), "")


def tratar_incidencias(caminho: Path, log_fn=print) -> pd.DataFrame:
    log_fn(f"  📖 Lendo arquivo: {caminho.name}")

    xl = pd.ExcelFile(str(caminho), engine="openpyxl")
    sheet = SHEET_NAME if SHEET_NAME in xl.sheet_names else xl.sheet_names[0]

    # Tenta os formatos possíveis: original GeoOnline (2 linhas meta + cabeçalho),
    # arquivo mesclado salvo por _merge_chunks (startrow=2), e fallback direto.
    df = None
    for kwargs in [
        {"header": 1, "skiprows": [0]},
        {"header": 2},
        {"header": 0},
    ]:
        candidate = xl.parse(sheet, **kwargs)
        candidate.columns = [str(c).strip() for c in candidate.columns]
        if any(c in candidate.columns for c in ("Data Início", "Data Inicio", "Número", "Numero")):
            df = candidate
            break
    if df is None:
        df = xl.parse(sheet, header=1, skiprows=[0])
        df.columns = [str(c).strip() for c in df.columns]

    log_fn(f"  📋 {len(df)} linhas, {len(df.columns)} colunas carregadas")

    # Remove PROGRAMADO
    if "Estado" in df.columns:
        antes = len(df)
        df = df[df["Estado"].astype(str).str.strip().str.upper() != "PROGRAMADO"]
        removidos = antes - len(df)
        if removidos > 0:
            log_fn(f"  🗑️  {removidos} PROGRAMADO removidos")

    log_fn("  🧹 Calculando status, duração, turnos...")

    # Normalizar nomes de colunas com/sem acento
    _col_map = {}
    for _orig, _alt in [("Data Início", "Data Inicio"), ("Data Fim", "DataFim"), ("Número", "Numero")]:
        if _orig not in df.columns and _alt in df.columns:
            _col_map[_alt] = _orig
    if _col_map:
        df.rename(columns=_col_map, inplace=True)
        log_fn(f"  🔧 Colunas renomeadas: {list(_col_map.keys())} → {list(_col_map.values())}")

    # Datas
    df["_dt_inicio"] = pd.to_datetime(df.get("Data Início", df.get("Data Inicio")), dayfirst=True, errors="coerce")
    df["_dt_fim"] = pd.to_datetime(df.get("Data Fim", df.get("DataFim")), dayfirst=True, errors="coerce")
    df["_ano_mes"] = df["_dt_inicio"].dt.strftime("%Y-%m")

    # Ano / Mês / Dia (da Data Início)
    df["Ano"] = df["_dt_inicio"].dt.year
    df["Mês"] = df["_dt_inicio"].dt.month
    df["Dia"] = df["_dt_inicio"].dt.day

    # Duração
    if "Duração" in df.columns:
        df["Duração em minutos"] = df["Duração"].apply(_converter_duracao)
    else:
        df["Duração em minutos"] = (
            (df["_dt_fim"] - df["_dt_inicio"]).dt.total_seconds() // 60
        ).fillna(0).clip(lower=0)

    # Status
    df["Status"] = np.where(
        df["_dt_fim"].isna() | (df["Data Fim"].astype(str).str.strip() == "---"),
        "Aberto", "Fechado",
    )

    # Causa normalizada
    df["_causa_norm"] = df["Causa"].astype(str).apply(
        lambda x: _sem_acento(x.strip().upper())
    )

    # Considerar TMA
    def _considerar(row):
        if row["Duração em minutos"] == 0:
            return "Não Considerar"
        causa = row["_causa_norm"]
        return "Considerar" if MAPA_CONSIDERAR.get(causa, True) else "Não Considerar"
    df["Considerar TMA"] = df.apply(_considerar, axis=1)

    # Equipe/Operador
    df["Equipe/Operador"] = df["_causa_norm"].apply(
        lambda c: "operador" if MAPA_TIPO.get(c) == "Remoto" else "equipe"
    )

    # Cluster duração
    bins = [0, 120, 240, 360, 480, 600, 720, 840, 960, 1080, 1200, 1320, 1440, np.inf]
    labels = ["0h-2h", "2h-4h", "4h-6h", "6h-8h", "8h-10h", "10h-12h", "12h-14h",
              "14h-16h", "16h-18h", "18h-20h", "20h-22h", "22h-24h", "24h+"]
    df["Cluster duração"] = pd.cut(df["Duração em minutos"], bins=bins, labels=labels, right=False)

    # Atribuição
    col_eq = "Equipe Atribuída"
    df["Atribuição"] = np.where(
        df[col_eq].isna() | (df[col_eq].astype(str).str.strip().isin(["---", "", "nan"])),
        "Não Atribuído", "Atribuído",
    )

    # Horas e Turnos
    df["Hora entrada"] = df["_dt_inicio"].dt.hour
    df["Hora saída"] = df["_dt_fim"].apply(lambda x: int(x.hour) if pd.notnull(x) else "Aberta")

    def _turno(h):
        if isinstance(h, str):
            return "Aberta"
        if 0 <= h < 8:
            return "A"
        if 8 <= h < 15:
            return "B"
        return "C"

    df["Turno abertura"] = df["Hora entrada"].apply(_turno)
    df["Turno fechamento"] = df["Hora saída"].apply(_turno)

    # Produtivo/Improdutivo
    col_cli = "Cli Af Max" if "Cli Af Max" in df.columns else df.columns[13]
    df["_cli_af_max"] = pd.to_numeric(df[col_cli], errors="coerce").fillna(0)

    def _produtivo(row):
        if row["_causa_norm"] in CAUSAS_IMPRODUTIVO:
            return "Improdutivo"
        if row["_cli_af_max"] == 0:
            return "Produtivo sem afetação"
        return "Produtivo"
    df["Produtivo/Improdutivo"] = df.apply(_produtivo, axis=1)

    # TME
    if "TME" in df.columns:
        df["TME Análise"] = np.where(pd.to_numeric(df["TME"], errors="coerce") > 100, "sim", "não")
    else:
        df["TME Análise"] = "não"

    # Entrada/Saída
    hoje = datetime.now().date()
    ontem = hoje - timedelta(days=1)

    def _cat_data(dt):
        if pd.isnull(dt):
            return "Aberto"
        d = dt.date()
        if d == hoje:
            return "Hoje"
        if d == ontem:
            return "Ontem"
        return "Dias Anteriores"

    df["Entrada"] = df["_dt_inicio"].apply(_cat_data)
    df["Saída"] = df["_dt_fim"].apply(lambda x: _cat_data(x) if pd.notnull(x) else "Aberto")

    # Data/Hora Atuação
    col_tmp = "TMP"
    if col_tmp in df.columns:
        tmp_min = pd.to_numeric(df[col_tmp], errors="coerce").fillna(0)
        df["Data Atuação"] = df["_dt_inicio"] + pd.to_timedelta(tmp_min, unit="m")
    else:
        df["Data Atuação"] = df["_dt_inicio"]
    df["Hora Atuação"] = df["Data Atuação"].apply(
        lambda x: x.strftime("%H:%M") if pd.notnull(x) else ""
    )

    # Ponto Elétrico (Dispositivo)
    col_pe = "Ponto Elétrico" if "Ponto Elétrico" in df.columns else df.columns[18]
    _INVALIDOS_PE = {"N/D", "", "NAN", "NONE", "---", "-"}
    def _extrair_ponto(val):
        s = str(val).strip()
        if s.upper() in _INVALIDOS_PE:
            return ""
        return s.split("-", 1)[0].strip() if "-" in s else s
    df["Ponto Elétrico (Dispositivo)"] = df[col_pe].apply(_extrair_ponto)
    df["_pe_tratado"] = df["Ponto Elétrico (Dispositivo)"]

    # Nível de tensão
    col_nt = "NT" if "NT" in df.columns else df.columns[1]
    df["Nivel de Tensão"] = (
        df[col_nt].astype(str).str.strip().str.upper()
        .replace({"MT TRONCO": "MT", "MT RAMAL": "MT"})
    )

    # Cluster clientes
    col_clts = "Clts Atual" if "Clts Atual" in df.columns else None
    if col_clts:
        cli = pd.to_numeric(df[col_clts], errors="coerce").fillna(0)
        df["Cluster clientes"] = pd.cut(
            cli,
            bins=[0, 50, 100, 400, 1000, np.inf],
            labels=["Até 50 clientes", "51-100 clientes", "101-400 clientes",
                    "401-1000 clientes", "1000+ clientes"],
            right=True,
        )
    else:
        df["Cluster clientes"] = ""

    # Sucursal Real
    col_al = "AL" if "AL" in df.columns else None
    if col_al:
        df["Sucursal Real"] = (
            df[col_al].astype(str).str.strip().str.upper().map(MAPA_SUCURSAL).fillna("")
        )
    else:
        df["Sucursal Real"] = ""

    # Processo das equipes
    col_eq_attr = "Equipe Atribuída" if "Equipe Atribuída" in df.columns else df.columns[30]
    df["Processo (Equipe atrib.)"] = df[col_eq_attr].apply(_extrair_processo)
    df["Enel/Parceira (Equipe atrib.)"] = df[col_eq_attr].apply(_extrair_enel)
    df["Grupo Processos (Equipe atrib.)"] = df["Processo (Equipe atrib.)"].apply(_grupo_processos)

    col_eq_desl = "Equipe Desl." if "Equipe Desl." in df.columns else None
    if col_eq_desl:
        df["Processo (Equipe desloc.)"] = df[col_eq_desl].apply(_extrair_processo)
        df["Enel/Parceira (Equipe desloc.)"] = df[col_eq_desl].apply(_extrair_enel)
        df["Grupo Processos (Equipe desloc.)"] = df["Processo (Equipe desloc.)"].apply(_grupo_processos)
    else:
        df["Processo (Equipe desloc.)"] = ""
        df["Enel/Parceira (Equipe desloc.)"] = ""
        df["Grupo Processos (Equipe desloc.)"] = ""

    # Rankings / Reincidência
    col_num_cli = next(
        (c for c in ["Nº Cliente", "UC", "Cliente", "N_Cliente"] if c in df.columns),
        None,
    )
    df["_num_cli_temp"] = (
        df[col_num_cli] if col_num_cli else df.iloc[:, 21]
    ).astype(str).str.strip().replace(["nan", "None", "", "0", "0.0"], "---")

    mask_base = (
        df["Produtivo/Improdutivo"].isin(["Produtivo", "Produtivo sem afetação"])
        & (df["Equipe/Operador"] == "equipe")
    )
    for mask, label, gcols in [
        (mask_base & (df["Nivel de Tensão"] == "BT") & (df["_cli_af_max"] > 1),
         "coletivo", ["_ano_mes", "_pe_tratado"]),
        (mask_base & (df["Nivel de Tensão"] == "BT") & (df["_cli_af_max"] == 1) & (df["_num_cli_temp"] != "---"),
         "individual", ["_ano_mes", "_num_cli_temp"]),
        (mask_base & df["Nivel de Tensão"].str.contains("MT", na=False) & (df["_pe_tratado"] != ""),
         "mt", ["_ano_mes", "_pe_tratado", "Nivel de Tensão"]),
    ]:
        if mask.any():
            sub = df[mask].copy().sort_values("_dt_inicio")
            df.loc[mask, f"_rank_{label}"] = sub.groupby(gcols).cumcount() + 1
            df.loc[mask, f"_total_{label}"] = df.loc[mask].groupby(gcols)[f"_rank_{label}"].transform("max")

    def _reincidente(row):
        try:
            nt = row["NT"]
            if nt == "AT":
                return "Incidência de AT"
            if row["Produtivo/Improdutivo"] == "Improdutivo":
                return f"{nt} – Improdutivo"
            if row["Equipe/Operador"] == "operador":
                return f"{nt} – Incidência"
            if "MT" in nt:
                return "Incidência de MT"
            if nt == "BT":
                if row["_cli_af_max"] > 1:
                    return "BT – Incidência Coletiva"
                if row.get("_num_cli_temp", "---") == "---":
                    return "Incidência Individual"
                r = row.get("_rank_individual", 0) or 0
                t = row.get("_total_individual", 0) or 0
                if r > 1:
                    return "BT – Reincidência"
                if r == 1 and t > 1:
                    return "BT – 1° Incidência individual"
                return "Incidência Individual"
            return "Incidência"
        except Exception:
            return "Incidência"

    def _reincidente_tipo(row):
        try:
            nt = row["NT"]
            if row["Equipe/Operador"] == "operador" or row["Produtivo/Improdutivo"] == "Improdutivo":
                if "MT" in nt:
                    pref = "MT RAMAL" if "RAMAL" in nt else "MT TRONCO"
                    tipo = "Incidência" if row["Equipe/Operador"] == "operador" else "Improdutivo"
                    return f"{pref} – {tipo}"
                return row["Reincidente"]
            if "MT" in nt:
                pref = "MT RAMAL" if "RAMAL" in nt else "MT TRONCO"
                if not row.get("_pe_tratado", ""):
                    return f"{pref} – Incidência"
                r = row.get("_rank_mt", 0) or 0
                t = row.get("_total_mt", 0) or 0
                if r == 1 and t > 1:
                    return f"{pref} – 1° Incidência"
                if r > 1:
                    return f"{pref} – Reincidência"
                return f"{pref} – Incidência"
            if nt == "BT" and row["_cli_af_max"] > 1:
                r = row.get("_rank_coletivo", 0) or 0
                t = row.get("_total_coletivo", 0) or 0
                if r == 1 and t > 1:
                    return "BT – 1° Incidência Coletiva"
                if r > 1:
                    return "BT – Reincidência Coletiva"
                return "BT – Incidência Coletiva"
            return row["Reincidente"]
        except Exception:
            return "Incidência"

    df["Reincidente"] = df.apply(_reincidente, axis=1)
    df["Reincidente tipo"] = df.apply(_reincidente_tipo, axis=1)

    # Ordem 2 Consolidada
    def _ordem2(row):
        try:
            as_v = str(row.iloc[44]).lower() if pd.notna(row.iloc[44]) else ""
            ad_v = str(row.iloc[29]).lower() if pd.notna(row.iloc[29]) else ""
            ae_v = str(row.iloc[30]).lower() if pd.notna(row.iloc[30]) else ""
            az_v = str(row.iloc[51]).lower() if pd.notna(row.iloc[51]) else ""
            if "sim" in as_v or "ord" in ad_v or "ord" in ae_v:
                return "sim"
            if any(t in az_v for t in ["ord2", "ord 2", "ordem 2", "ordem2"]):
                return "sim"
            return "não"
        except Exception:
            return "não"
    df["Ordem 2 Consolidada"] = df.apply(_ordem2, axis=1)

    # Montar DataFrame final
    cols_tratamento = [
        "Duração em minutos", "Status", "Considerar TMA", "Equipe/Operador",
        "Cluster duração", "Atribuição", "Hora entrada", "Turno abertura",
        "Hora saída", "Turno fechamento", "Produtivo/Improdutivo",
        "TME Análise", "Entrada", "Saída", "Data Atuação", "Hora Atuação",
        "Ponto Elétrico (Dispositivo)", "Nivel de Tensão", "Cluster clientes",
        "Sucursal Real", "Processo (Equipe atrib.)", "Enel/Parceira (Equipe atrib.)",
        "Grupo Processos (Equipe atrib.)", "Processo (Equipe desloc.)",
        "Enel/Parceira (Equipe desloc.)", "Grupo Processos (Equipe desloc.)",
        "Reincidente", "Reincidente tipo", "Ordem 2 Consolidada",
        "Ano", "Mês", "Dia",
    ]
    cols_temp = [c for c in df.columns if c.startswith("_") and c not in ("_dt_inicio",)]
    cols_orig = [c for c in df.columns if c not in cols_temp and c not in cols_tratamento]
    return df[cols_orig + cols_tratamento].copy()


def gerar_analises_n8n(df: pd.DataFrame) -> dict:
    """Gera o payload JSON para envio ao n8n."""
    agora = datetime.now()
    col_polo = "Polo" if "Polo" in df.columns else None
    if not col_polo:
        return {}

    UTS = ["MAGÉ", "NITERÓI", "SÃO GONÇALO", "SERRANA", "SUL"]
    UTN = ["CAMPOS", "LAGOS", "MACAÉ", "NOROESTE"]
    ORDEM_CLUSTERS = [
        "0h-2h", "2h-4h", "4h-6h", "6h-8h", "8h-10h", "10h-12h",
        "12h-14h", "14h-16h", "16h-18h", "18h-20h", "20h-22h", "22h-24h", "24h+",
    ]

    # Exclui incidentes de "DT" (coluna "Número") e "ATEND CANCELADO CR" (coluna "Causa")
    col_num = "Número" if "Número" in df.columns else None
    if col_num:
        df = df[~df[col_num].astype(str).str.upper().str.startswith("DT")]

    col_causa = "Causa" if "Causa" in df.columns else None
    if col_causa:
        mask_causa = df[col_causa].astype(str).str.upper().str.strip()
        df = df[~mask_causa.str.contains("ATEND.*CANCEL", regex=True, na=False)]

    df_ab = df[df["Status"] == "Aberto"].copy()
    df_ab["_polo"] = (
        df_ab[col_polo].astype(str).str.strip().str.upper()
        .replace({"": "SEM POLO", "NAN": "SEM POLO", "---": "SEM POLO"})
    )

    col_clts = "Clts Atual" if "Clts Atual" in df.columns else None

    def _abertas(polos):
        res = {}
        for p in polos:
            sub = df_ab[df_ab["_polo"] == p]
            if sub.empty:
                continue
            clusters = {
                c: int((sub["Cluster duração"].astype(str) == c).sum())
                for c in ORDEM_CLUSTERS
                if (sub["Cluster duração"].astype(str) == c).sum() > 0
            }
            clientes = int(sub[col_clts].astype(float).sum()) if col_clts else 0
            res[p] = {"total": len(sub), "clientes": clientes, "clusters": clusters}
        return res

    def _atribuicao(polos):
        res = {}
        for p in polos:
            sub = df_ab[df_ab["_polo"] == p]
            if sub.empty:
                continue
            nao_atr = sub[sub["Atribuição"] == "Não Atribuído"]
            por_dia = {}
            if not nao_atr.empty and "Entrada" in nao_atr.columns:
                agora = datetime.now()
                for dia, grupo in nao_atr.groupby("Entrada"):
                    duracao = agora - grupo["_dt_inicio"]
                    mais_24h = int((duracao > pd.Timedelta(hours=24)).sum())
                    por_dia[str(dia)] = {"total": int(len(grupo)), "mais_24h": mais_24h}
            res[p] = {"total": len(sub), "nao_atribuido": int(len(nao_atr)), "por_dia": por_dia}
        return res

    col_clts = "Clts Atual" if "Clts Atual" in df.columns else None

    def _top10(polos):
        res = {}
        for p in polos:
            sub = df_ab[df_ab["_polo"] == p].copy()
            if sub.empty or not col_clts or not col_num:
                continue
            sub["_c"] = pd.to_numeric(sub[col_clts], errors="coerce").fillna(0)
            total = sub["_c"].sum()
            top = sub.nlargest(10, "_c")[[col_num, "_c", "Atribuição"]]
            res[p] = {
                "total_clientes": int(total),
                "top10": [
                    {"numero": str(r[col_num]), "clientes": int(r["_c"]),
                     "perc": round(r["_c"] / total * 100, 1) if total > 0 else 0,
                     "atribuido": str(r["Atribuição"]).strip() == "Atribuído"}
                    for _, r in top.iterrows()
                ],
            }
        return res

    col_tme = "TME" if "TME" in df.columns else None

    def _tme(polos):
        res = {}
        for p in polos:
            sub_polo = df[df[col_polo].astype(str).str.strip().str.upper() == p]
            sub = sub_polo[sub_polo["TME Análise"] == "sim"]
            if sub.empty or not col_num:
                continue
            sub = sub[sub[col_num].astype(str).str.startswith("00")]
            if sub.empty:
                continue
            incidentes = sorted(
                [{"numero": str(r[col_num]),
                  "tme": int(pd.to_numeric(r[col_tme], errors="coerce") or 0) if col_tme else 0}
                 for _, r in sub.iterrows()],
                key=lambda x: x["tme"], reverse=True,
            )
            res[p] = {"total": len(incidentes), "incidentes": incidentes[:10], "extras": max(0, len(incidentes) - 10)}
        return res

    # ── Resumo por Polo ──────────────────────────────────────────────────
    col_reinc = "Reincidente tipo" if "Reincidente tipo" in df.columns else None
    col_eq_desl = "Equipe Desl." if "Equipe Desl." in df.columns else None
    col_tensao = "Nivel de Tensão" if "Nivel de Tensão" in df.columns else None

    def _calc_por_polos(lista_polos):
        res = {}
        for p in lista_polos:
            mask_p = df[col_polo].astype(str).str.strip().str.upper() == p
            sub = df[mask_p]
            sub_ab = sub[sub["Status"] == "Aberto"]

            # Abertos atual = todos em aberto agora (independente de quando abriram)
            abertos_atual = int(len(sub_ab))

            # Abertos hoje = abertos hoje independente de status (abertos + fechados no dia)
            sub_hj = sub[sub["Entrada"] == "Hoje"]
            abertos_hoje = int(len(sub_hj))

            abertos_hj_por_tensao = {}
            if col_tensao and not sub_hj.empty:
                for tensao, grupo in sub_hj.groupby(col_tensao):
                    t = str(tensao).strip()
                    if t and t.upper() not in ("", "NAN"):
                        abertos_hj_por_tensao[t] = int(len(grupo))

            abertos_atual_por_tensao = {}
            if col_tensao and not sub_ab.empty:
                for tensao, grupo in sub_ab.groupby(col_tensao):
                    t = str(tensao).strip()
                    if t and t.upper() not in ("", "NAN"):
                        abertos_atual_por_tensao[t] = int(len(grupo))

            backlog_ontem = int((sub_ab["Entrada"] == "Ontem").sum())
            backlog_anteriores = int(sub_ab[(sub_ab["Entrada"] != "Hoje") & (sub_ab["Entrada"] != "Ontem")].shape[0])

            sub_fech_hj = sub[sub["Saída"] == "Hoje"]
            fechados_hoje = int(len(sub_fech_hj))
            fechados_eq = int((sub_fech_hj["Equipe/Operador"] == "equipe").sum())
            fechados_remoto = int((sub_fech_hj["Equipe/Operador"] == "operador").sum())

            re = {"bt_reincidencia": 0, "bt_reincidencia_coletiva": 0, "mt_tronco_reincidencia": 0}
            if col_reinc:
                rt = sub[col_reinc].astype(str).str.strip()
                re["bt_reincidencia"] = int((rt == "BT – Reincidência").sum())
                re["bt_reincidencia_coletiva"] = int((rt == "BT – Reincidência Coletiva").sum())
                re["mt_tronco_reincidencia"] = int(rt.str.contains(r"MT TRONCO.*Reincidência", regex=True, na=False).sum())

            eq_ativas = 0
            if col_eq_desl:
                eqs = sub_fech_hj[col_eq_desl].dropna().unique()
                eq_ativas = len([e for e in eqs if str(e).strip() not in ("---", "", "nan")])

            media = round(fechados_hoje / eq_ativas, 1) if eq_ativas > 0 else 0

            res[p] = {
                "abertos_atual": abertos_atual,
                "abertos_atual_por_tensao": abertos_atual_por_tensao,
                "abertos_hoje": abertos_hoje,
                "abertos_hj_por_tensao": abertos_hj_por_tensao,
                "backlog_ontem": backlog_ontem,
                "backlog_anteriores": backlog_anteriores,
                "fechados_hoje": fechados_hoje,
                "fechados_equipe": fechados_eq,
                "fechados_remoto": fechados_remoto,
                "reincidentes": re,
                "equipes_ativas": eq_ativas,
                "media_producao": media,
            }
        return res

    col_turno_fech = "Turno fechamento" if "Turno fechamento" in df.columns else None

    def _fechamento(polos):
        res = {}
        if col_turno_fech is None:
            return res
        df_fech = df[df["Saída"] == "Hoje"].copy()
        mask_aberto = df_fech["Saída"].astype(str).str.strip().str.upper().str.contains(r"^ABERTA?$", na=False)
        mask_turno_aberto = df_fech[col_turno_fech].astype(str).str.strip().str.upper().str.contains(r"^ABERTA?$", na=False)
        df_fech = df_fech[~mask_aberto & ~mask_turno_aberto]
        df_fech["_polo"] = (
            df_fech[col_polo].astype(str).str.strip().str.upper()
            .replace({"": "SEM POLO", "NAN": "SEM POLO", "---": "SEM POLO"})
        )
        df_fech["_turno_norm"] = (
            df_fech[col_turno_fech].astype(str).str.strip()
            .replace({"A": "Turno A", "B": "Turno B", "C": "Turno C"})
        )
        for p in polos:
            sub = df_fech[df_fech["_polo"] == p]
            if sub.empty:
                continue
            por_turno = {}
            for turno, grupo in sub.groupby("_turno_norm"):
                t = str(turno).strip()
                if t.upper() in ("ABERTO", "ABERTA", "", "NAN"):
                    continue
                por_tensao = {}
                if col_tensao:
                    for tensao, sub_t in grupo.groupby(col_tensao):
                        tv = str(tensao).strip()
                        if tv and tv.upper() not in ("", "NAN"):
                            por_tensao[tv] = int(len(sub_t))
                por_turno[t] = {"total": int(len(grupo)), "por_tensao": por_tensao}
            res[p] = {"total": int(len(sub)), "por_turno": por_turno}
        return res

    resumo = {"UTS": _calc_por_polos(UTS), "UTN": _calc_por_polos(UTN)}

    abertas_uts = _abertas(UTS)
    abertas_utn = _abertas(UTN)
    resumo_texto = "\n".join(
        [f"{p}: {v['total']} abertos" for p, v in {**abertas_uts, **abertas_utn}.items()]
    )
    sem_polo = int((df_ab.get("_polo", pd.Series()) == "SEM POLO").sum())

    return {
        "data_hora": agora.strftime("%d/%m/%Y %H:%M"),
        "total_registros": int(df.shape[0]),
        "abertas": {
            "total_geral": len(df_ab),
            "UTS": {"total": sum(v["total"] for v in abertas_uts.values()), "polos": abertas_uts},
            "UTN": {"total": sum(v["total"] for v in abertas_utn.values()), "polos": abertas_utn},
            "sem_polo": sem_polo,
        },
        "atribuicao": {"UTS": _atribuicao(UTS), "UTN": _atribuicao(UTN)},
        "top10": {"UTS": _top10(UTS), "UTN": _top10(UTN)},
        "tme": {"UTS": _tme(UTS), "UTN": _tme(UTN)},
        "fechamento": {"UTS": _fechamento(UTS), "UTN": _fechamento(UTN)},
        "resumo_texto": resumo_texto,
        "resumo": resumo,
    }


def enviar_n8n(payload: dict, webhook_url: str, log_fn=print) -> None:
    if not webhook_url:
        return
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        webhook_url, data=data, headers={"Content-Type": "application/json"}
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            if resp.status == 200:
                log_fn("  ✅ Análises enviadas ao n8n")
            else:
                log_fn(f"  ⚠️  n8n respondeu {resp.status}")
    except Exception as e:
        log_fn(f"  ❌ Falha n8n: {e}")
