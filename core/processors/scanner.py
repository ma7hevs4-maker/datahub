"""
Tratamento do relatório M300 (Scanner 4.0 - RJ) exportado do Spotfire.

Reproduz, em pandas, as 38 colunas de tratamento que hoje existem no
modelo "Scanner 4.0 tratamento.xlsx" (fórmulas em BQ..DB). A saída preserva
as 68 colunas originais + as 38 de tratamento, NA MESMA ORDEM do Excel, para
que o dashboard (Lovable) consiga ler pelos nomes das colunas.

Codificação do CSV do Spotfire: UTF-16 LE, separado por TAB.
"""
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

MAPA_PROCESSO = {
    "A": "AMPLA CHIP", "R": "CORTE E RELIGAÇÃO", "L": "NOVAS LIGAÇÕES",
    "F": "GESTOR NOVAS LIGAÇÕES", "P": "PERDAS", "M": "MANUTENÇÃO",
    "D": "PODA", "V": "LINHA VIVA", "T": "TELECONTROLE",
    "B": "OBRAS", "O": "OPERAÇÃO", "X": "EXTRA OPERAÇÃO",
}

# 13 UAMs -> 9 polos. Espelha o VLOOKUP da coluna BZ do modelo Excel.
_MAPA_POLO_M300 = {
    "NITERÓI": "Niterói", "SÃO GONÇALO": "São Gonçalo", "MAGÉ": "Magé",
    "ANGRA": "Sul", "RESENDE": "Sul", "PETRÓPOLIS": "Serrana",
    "TERESÓPOLIS": "Serrana", "ARARUAMA": "Lagos", "CABO FRIO": "Lagos",
    "CAMPOS": "Campos", "CANTAGALO": "Noroeste", "ITAPERUNA": "Noroeste",
    "PÁDUA": "Noroeste", "MACAÉ": "Macaé",
}


# ── Helpers de texto da Equipe (idênticos aos do modelo Excel) ──────────────────

def _prefixo_processo(equipe: str) -> str:
    """LEFT(RIGHT(equipe,4),1)"""
    s = str(equipe).strip()
    return s[-4] if len(s) >= 4 else ""


def _processo(equipe: str) -> str:
    pref = _prefixo_processo(equipe)
    return MAPA_PROCESSO.get(pref.upper(), "EQUIPE GENÉRICA")


def _enel_parceira(equipe: str) -> str:
    """IF(LEFT(RIGHT(equipe,3),1)='P','Parceira',IF(='E','Enel','Fora do Padrão'))"""
    s = str(equipe).strip()
    if len(s) < 3:
        return "Fora do Padrão"
    char = s[-3]
    if char == "P":
        return "Parceira"
    if char == "E":
        return "Enel"
    return "Fora do Padrão"


def _turno(equipe: str) -> str:
    s = str(equipe).strip()
    last = s[-1] if s else ""
    return last if last in ("A", "B", "C") else "Fora do Padrão"


def _equipe_fmt(equipe: str) -> str:
    """LEFT(equipe,5) & RIGHT(equipe,1) -> ex: NI516B"""
    s = str(equipe).strip()
    if len(s) < 6:
        return s
    return s[:5] + s[-1]


def _polo_m300(uam: str) -> str:
    """Mapa UAM -> Polo (coluna BZ do modelo Excel)."""
    return _MAPA_POLO_M300.get(str(uam).strip().upper(), "")


# ── Helpers numéricos ──────────────────────────────────────────────────────────

def _num(series: pd.Series) -> pd.Series:
    """Converte para número, tratando vírgula decimal ('412,00' -> 412.0)."""
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False), errors="coerce"
    )


def _frac_dia(s: pd.Series) -> pd.Series:
    """Fração do dia (0..1) da parte horária de cada datetime (serial Excel)."""
    return (s.dt.hour * 3600 + s.dt.minute * 60 + s.dt.second) / 86400.0


def _fmt_hms(frac: float) -> str:
    """Formata fração de dia como h:mm:ss (mesma saída de TEXT(x,'h:mm:ss'))."""
    frac = frac % 1.0
    total = int(round(frac * 86400))
    h, r = divmod(total, 3600)
    m, s = divmod(r, 60)
    return f"{h}:{m:02d}:{s:02d}"


def _piece(v, c0, c1, a, b, teto, baixo=0):
    """Reproduz o padrão SE(...,...,...) em degraus do Excel."""
    if pd.isna(v):
        return "ERRO"
    if v <= c0:
        return baixo
    if v < c1:
        return v * a + b
    if v >= c1:
        return teto
    return "ERRO"


def tratar_m300(caminho: Path, log_fn=print) -> pd.DataFrame:
    """Lê o CSV UTF-16 do Spotfire e gera as 38 colunas de tratamento."""
    log_fn(f"  📖 Lendo arquivo M300: {caminho.name}")
    df = pd.read_csv(str(caminho), sep="\t", dtype=str, encoding="utf-16")
    df.columns = [str(c).strip() for c in df.columns]
    log_fn(f"  📋 {len(df)} linhas, {len(df.columns)} colunas (originais)")

    # ── Entradas numéricas (com vírgula decimal) ──
    ao = _num(df["Qtd Serviços"])
    aj = _num(df["HT total"])
    bb = _num(df["HD Total"])
    ar = _num(df["Task Time Total TR"])
    au = _num(df["TR Ordem Imp SS equipe"])
    av = _num(df["1º Desloc"])                       # bruto (coluna original)
    ba = _num(df["1º Login Corrigido"])
    bc = _num(df["Retorno a base"])
    bn = _num(df["TR Ordem Sencundário equipe"])
    bo = _num(df["TR_TOTAL_CAL"])
    bp = _num(df["TEMPO_PADRAO_TOTAL_CAL"])

    # ── Entradas de data/hora ──
    dt_ref = pd.to_datetime(df["Data Referência"], format="%d/%m/%Y", errors="coerce")
    c_cal = pd.to_datetime(df["Inicio Calendario"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    ag_rel = pd.to_datetime(df["Hora 1º Deslocamento"], format="%d/%m/%Y %H:%M:%S", errors="coerce")

    equipe = df["Equipe"].astype(str)

    # ── BQ / BR: diferença de horário (Hora 1º Deslocamento - Inicio Calendario) ──
    diff = (_frac_dia(ag_rel) - _frac_dia(c_cal)) % 1.0
    bq = diff.apply(
        lambda x: ("-" + _fmt_hms(1 - x)) if (not pd.isna(x) and x > 0.5)
        else ("" if pd.isna(x) else _fmt_hms(x))
    )

    def _br(x, a):
        if not pd.isna(a) and a > 0:
            return "Após o Calendario"
        if pd.isna(x):
            return ""
        if x <= 0.5:
            return "No horário"
        return "Abaixo de 15 min" if (1 - x) <= (15 / 1440) else "Acima de 15 min"

    br = pd.Series([_br(x, a) for x, a in zip(diff, av)], index=df.index)

    # ── Colunas de texto da Equipe ──
    bs = equipe.apply(_prefixo_processo)
    bt = equipe.apply(_processo)
    bu = equipe.apply(_enel_parceira)
    bv = equipe.apply(_turno)
    by = equipe.apply(_equipe_fmt)
    bz = df["Base"].astype(str).apply(_polo_m300)

    # ── Razões e métricas ──
    cc = (aj / bb).where(bb != 0, other=np.nan)        # Utilização (ratio)
    cd = (ar / bb).where(bb != 0, other=np.nan)        # Task Time (ratio)
    ce = (bp / bo).where(bo != 0, other=np.nan)        # Calc_eficiencia
    cf = dt_ref.dt.month
    ca = pd.Timestamp.today().normalize() - pd.Timedelta(days=1)
    cb = df["Data Referência"].astype(str) + " " + equipe
    bw = dt_ref
    bx = pd.Series([pd.NaT] * len(df), index=df.index)

    # ── Pontuações (piecewise, iguais ao modelo) ──
    ch = ao.apply(lambda v: _piece(v, 1, 5.5, 1.36, 9.02, 16.5, 0))
    cj = ce.apply(lambda v: _piece(v, 0.8, 1.15, 11.33, -1.33, 11.7, 0))
    cm = cc.apply(lambda v: _piece(v, 0.6, 0.88, 40, -24, 11.2, 0))
    cn = ba.apply(lambda v: _piece(v, 7, 12, -2.5, 30, 0, 13.6))
    co = av.apply(lambda v: _piece(v, 20, 30, -0.72, 28, 0, 13.6))
    cs = bc.apply(lambda v: _piece(v, 35, 50, -0.5, 25, 0, 7.5))
    ct = cd.apply(lambda v: _piece(v, 0.4, 0.65, 55.53, -12.77, 23.33, 0))
    cu = bn.apply(lambda v: _piece(v, 45, 72, -0.682, 49.1, 0, 18.41))
    cv = au.apply(lambda v: _piece(v, 17, 28, -1.25, 35, 0, 13.75))

    # Classificação = soma das colunas CH..CX (ignora texto/"ERRO", como o Excel)
    cg = sum(pd.to_numeric(c, errors="coerce") for c in (ch, cj, cm, cn, co, cs, ct, cu, cv))

    # ── Validações ──
    cy = df["Log In"].apply(
        lambda v: "Com Fim de Turno" if str(v).strip() not in ("", "nan", "None") else "Sem Fim de Turno"
    )
    cz = df["Log Off"].apply(
        lambda v: "Com Fim de Turno" if str(v).strip() not in ("", "nan", "None") else "Sem Fim de Turno"
    )
    _intv = _num(df["Intervalo"])
    da = _intv.apply(lambda v: "Sem Intervalo" if (pd.isna(v) or v == 0) else "Com Intervalo")
    db = av.apply(lambda v: "Sem 1º Deslocamento" if (pd.isna(v) or v == 0) else "Com 1º Deslocamento")

    # ── Montagem: 68 originais + 38 de tratamento (ordem do Excel) ──
    trat = [
        ("Dif Deslocamento e Inicio do calendario", bq),
        ("Deslocamento antes do calendario", br),
        ("prefixo processo", bs),
        ("Processo", bt),
        ("Enel/Parceira", bu),
        ("Turno ", bv),
        ("Data", bw),
        ("Data ", bx),
        ("EQUIPE", by),
        ("POLO", bz),
        ("x", ca),
        ("chave", cb),
        ("Utilização", cc),
        ("Task Time ", cd),
        ("Calc_eficiencia", ce),
        ("mês_calculo", cf),
        ("Classificação", cg),
        ("Ativ/Equipe/Dia", ch),
        ("Produtividade", None),
        ("Eficiencia", cj),
        ("Eficiencia Real", None),
        ("Qtd Serv IIº", None),
        ("Utilização", cm),
        ("1º Login", cn),
        ("1º Desloc", co),
        ("1º Despacho", None),
        ("Plataforma", None),
        ("Intervalo", None),
        ("Retorno a Base", cs),
        ("Task Time TR", ct),
        ("TMR Secundário", cu),
        ("TMR Improd.", cv),
        ("TML", None),
        ("H.E.(Min.)", None),
        ("valida login", cy),
        ("valida logoff", cz),
        ("valida intervalo", da),
        ("valida 1° deslocamento", db),
    ]
    frames = [df]
    for nome, serie in trat:
        if serie is None:
            col = pd.Series([pd.NA] * len(df), index=df.index)
        elif isinstance(serie, pd.Series):
            col = serie
        else:
            col = pd.Series([serie] * len(df), index=df.index)
        frames.append(pd.DataFrame({nome: col}))
    out = pd.concat(frames, axis=1)
    log_fn(f"  ✅ M300 tratado ({len(out)} linhas, {len(out.columns)} colunas)")
    return out


def tratar_deslocamentos(caminho: Path, log_fn=print) -> pd.DataFrame:
    """Pass-through: lê o CSV de deslocamentos (UTF-16) sem transformações."""
    log_fn(f"  📖 Lendo arquivo: {caminho.name}")
    df = pd.read_csv(str(caminho), sep="\t", dtype=str, encoding="utf-16")
    df.columns = [str(c).strip() for c in df.columns]
    log_fn(f"  📋 {len(df)} linhas, {len(df.columns)} colunas carregadas (pass-through)")
    return df
