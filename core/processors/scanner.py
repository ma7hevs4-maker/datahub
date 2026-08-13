"""
Tratamento do relatório Scanner 4.0 - Tabela Completa (Spotfire).
Colunas A-BL são dados crus; colunas de tratamento são geradas aqui.
Ignorar colunas BM (x2), BN (x) e BT (Data duplicada).
"""
import warnings
import pandas as pd
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

MAPA_PROCESSO = {
    "A": "AMPLA CHIP", "R": "CORTE E RELIGAÇÃO", "L": "NOVAS LIGAÇÕES",
    "F": "GESTOR NOVAS LIGAÇÕES", "P": "PERDAS", "M": "MANUTENÇÃO",
    "D": "PODA", "V": "LINHA VIVA", "T": "TELECONTROLE",
    "B": "OBRAS", "O": "OPERAÇÃO", "X": "EXTRA OPERAÇÃO",
}

MAPA_POLO = {
    "NITERÓI": "Niterói", "NITEROI": "Niterói",
    "SÃO GONÇALO": "São Gonçalo", "SAO GONCALO": "São Gonçalo",
    "MAGÉ": "Magé", "MAGE": "Magé",
    "ANGRA": "Sul", "RESENDE": "Sul",
    "PETRÓPOLIS": "Serrana", "PETROPOLIS": "Serrana",
    "TERESÓPOLIS": "Serrana", "TERESOPOLIS": "Serrana",
}


def _prefixo_processo(equipe: str) -> str:
    """LEFT(RIGHT(equipe, 4), 1)"""
    s = str(equipe).strip()
    return s[-4] if len(s) >= 4 else ""


def _processo(equipe: str) -> str:
    pref = _prefixo_processo(equipe)
    return MAPA_PROCESSO.get(pref.upper(), "EQUIPE GENÉRICA")


def _enel_parceira(equipe: str) -> str:
    """IF(LEFT(RIGHT(equipe, 3), 1) = 'P', 'Parceira', IF(...='E', 'Enel', 'Fora do Padrão'))"""
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
    """LEFT(equipe, 5) & RIGHT(equipe, 1)  →  ex: NI516B"""
    s = str(equipe).strip()
    if len(s) < 6:
        return s
    return s[:5] + s[-1]


def _polo(uo: str) -> str:
    return MAPA_POLO.get(str(uo).strip().upper(), "")


def tratar_scanner(caminho: Path, log_fn=print) -> pd.DataFrame:
    log_fn(f"  📖 Lendo arquivo: {caminho.name}")

    # O arquivo pode chegar como xlsx ou csv do Spotfire
    if caminho.suffix.lower() == ".csv":
        df = pd.read_csv(str(caminho), sep="\t", dtype=str, encoding="utf-8")
    else:
        sheet = "Scanner 4.0 - RJ - Tabela Compl"
        try:
            df = pd.read_excel(str(caminho), sheet_name=sheet, dtype=str)
        except Exception:
            df = pd.read_excel(str(caminho), dtype=str)

    df.columns = [str(c).strip() for c in df.columns]
    log_fn(f"  📋 {len(df)} linhas, {len(df.columns)} colunas carregadas")

    col_equipe = "Equipe" if "Equipe" in df.columns else df.columns[1]
    col_uo = "UO" if "UO" in df.columns else df.columns[11]
    col_login = "Log In" if "Log In" in df.columns else df.columns[3]

    log_fn("  🧹 Calculando processo, turno, polo...")

    df["prefixo processo"] = df[col_equipe].apply(_prefixo_processo)
    df["Processo"] = df[col_equipe].apply(_processo)
    df["Enel/Parceira"] = df[col_equipe].apply(_enel_parceira)
    df["Turno"] = df[col_equipe].apply(_turno)
    df["Data"] = df[col_login]
    df["EQUIPE"] = df[col_equipe].apply(_equipe_fmt)
    df["POLO"] = df[col_uo].apply(_polo)

    log_fn(f"  ✅ Scanner tratado ({len(df)} linhas)")
    return df


def tratar_deslocamentos(caminho: Path, log_fn=print) -> pd.DataFrame:
    """Pass-through: lê e retorna sem transformações."""
    log_fn(f"  📖 Lendo arquivo: {caminho.name}")

    if caminho.suffix.lower() == ".csv":
        df = pd.read_csv(str(caminho), sep="\t", dtype=str, encoding="utf-8")
    else:
        df = pd.read_excel(str(caminho), dtype=str)

    df.columns = [str(c).strip() for c in df.columns]
    log_fn(f"  📋 {len(df)} linhas, {len(df.columns)} colunas carregadas (pass-through)")
    return df
