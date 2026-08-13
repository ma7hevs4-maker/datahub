"""
Mantém bases mensais por relatório.
Cada relatório tem sua subpasta; arquivos gerados: YYYY-MM.xlsx

Lógica de UPSERT por mês:
  - Se o arquivo já existe: mescla dados existentes + novos,
    deduplica pela chave (ex: "Número"), mantendo a versão mais recente.
  - Se não existe: cria o arquivo com os dados novos.
"""
import os
import warnings
import pandas as pd
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Coluna usada como chave de deduplicação/upsert por relatório
CHAVE_UPSERT = {
    "incidencias": "Número",
    "scanner": None,       # dedup por linha inteira se None
    "deslocamentos": None,
}


def _upsert_mes(
    caminho: Path,
    novos: pd.DataFrame,
    chave: str | None,
    log_fn=print,
) -> tuple[int, int]:
    """
    Faz upsert de `novos` no arquivo `caminho`.
    Retorna (total_final, adicionados, atualizados).
    """
    if not chave or chave not in novos.columns:
        # Sem chave: usa linha inteira para dedup
        if caminho.exists():
            existentes = pd.read_excel(str(caminho), dtype=str)
            combinado = pd.concat([existentes, novos], ignore_index=True)
            antes = len(combinado)
            combinado = combinado.drop_duplicates(keep="last")
            depois = len(combinado)
            adicionados = antes - depois
        else:
            combinado = novos
            adicionados = len(combinado)
    else:
        # Com chave: upsert baseado na coluna chave
        if caminho.exists():
            existentes = pd.read_excel(str(caminho), dtype=str)
            # IDs dos novos
            ids_novos = set(novos[chave].astype(str).str.strip())
            # Remove dos existentes os que serão atualizados
            mantidos = existentes[
                ~existentes[chave].astype(str).str.strip().isin(ids_novos)
            ]
            atualizados = len(existentes) - len(mantidos)
            combinado = pd.concat([mantidos, novos], ignore_index=True)
            adicionados = len(novos) - atualizados
        else:
            combinado = novos
            adicionados = len(combinado)
            atualizados = 0

    tmp_path = caminho.with_suffix(".tmp.xlsx")
    combinado.to_excel(str(tmp_path), index=False)
    os.replace(str(tmp_path), str(caminho))
    return len(combinado), adicionados, atualizados if chave else 0


def atualizar_base_mensal(
    df: pd.DataFrame,
    relatorio: str,
    pasta_local: str,
    col_data: str = None,
    log_fn=print,
) -> list[Path]:
    """
    Divide df por mês (usando col_data) e mantém arquivos mensais
    em pasta_local/bases_mensais/<relatorio>/YYYY-MM.xlsx com upsert.
    Retorna lista de arquivos criados ou modificados.
    """
    base_dir = Path(pasta_local) / "bases_mensais" / relatorio
    base_dir.mkdir(parents=True, exist_ok=True)

    # Detecta coluna de data se não informada
    if col_data is None:
        candidatas = ["Data Início", "_dt_inicio", "Data Referência", "Log In", "Data"]
        col_data = next((c for c in candidatas if c in df.columns), None)
        if col_data is None:
            log_fn(f"  ⚠️  [{relatorio}] Coluna de data não encontrada — base mensal ignorada")
            return []

    df = df.copy()
    df["_data_parsed"] = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")
    df["_ano_mes"] = df["_data_parsed"].dt.strftime("%Y-%m")
    df_valido = df[df["_ano_mes"].notna()].copy()

    chave = CHAVE_UPSERT.get(relatorio)
    modificados = []

    for mes, grupo in df_valido.groupby("_ano_mes"):
        grupo = grupo.drop(columns=["_data_parsed", "_ano_mes"], errors="ignore")
        caminho = base_dir / f"{mes}.xlsx"

        total, adicionados, atualizados = _upsert_mes(caminho, grupo, chave, log_fn)
        modificados.append(caminho)
        if atualizados > 0:
            log_fn(f"  🔄 [{relatorio}] {mes}: {total} linhas ({adicionados} novos, {atualizados} atualizados)")
        else:
            log_fn(f"  ➕ [{relatorio}] {mes}: {total} linhas ({adicionados} novos)")

    return modificados
