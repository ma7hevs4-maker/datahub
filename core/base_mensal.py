"""
Mantém bases mensais por relatório.
Cada relatório tem sua subpasta; arquivos gerados: YYYY-MM.xlsx

Lógica de UPSERT por mês:
  - Se o arquivo já existe: mescla dados existentes + novos,
    deduplica pela chave (ex: "Número"), mantendo a versão mais recente.
  - Se não existe: cria o arquivo com os dados novos.
"""
import hashlib
import json
import os
import time
import warnings
import numpy as np
import pandas as pd
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Coluna(s) usada(s) como chave de deduplicação/upsert por relatório.
# Para deslocamentos, a chave composta evita duplicar a mesma ordem/equipe/dia
# quando o Spotfire reexporta a base com horários atualizados.
CHAVE_UPSERT = {
    "incidencias": "Número",
    "scanner": None,       # dedup por linha inteira se None
    "deslocamentos": [
        "Data Referência", "Equipe", "Nr_Ordem", "Despachada",
    ],
    "tabela_completa": [
        "Data Referência", "Equipe", "Nr_Ordem", "Despachada",
    ],
}


# Regra de negócio: incidente com Status "Aberto" cuja Data Início é anterior aos últimos N dias (a partir de hoje)
# não fica na base mensal de incidências (e, por consequência, não vai para o banco).
DIAS_ABERTO_BASE = 5


def descartar_abertos_antigos(df: pd.DataFrame, col_data: str = "Data Início", dias: int = DIAS_ABERTO_BASE):
    """Retorna (df_sem_abertos_antigos, quantidade_removida)."""
    if df is None or len(df) == 0 or "Status" not in df.columns or col_data not in df.columns:
        return df, 0
    aberto = df["Status"].astype(str).str.strip().str.lower().eq("aberto")
    try:
        ini = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce", format="mixed")
    except (TypeError, ValueError):
        ini = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")
    corte = pd.Timestamp.now().normalize() - pd.Timedelta(days=dias)
    antigo = (aberto & (ini < corte)).fillna(False)
    n = int(antigo.sum())
    return (df[~antigo] if n else df), n


def _chave_series(df: pd.DataFrame, chave):
    """Série com o valor da chave (string simples ou composta por '|')."""
    if not chave:
        return None
    cols = [chave] if isinstance(chave, str) else list(chave)
    if not cols or any(c not in df.columns for c in cols):
        return None
    serie = df[cols[0]].astype(str).str.strip()
    for c in cols[1:]:
        serie = serie + "|" + df[c].astype(str).str.strip()
    return serie


# Colunas voláteis que não devem invalidar o cache de um mês sem mudança real.
_HASH_IGNORAR = {"x", "DATA_ETL"}

# Cache fica fora da pasta do SharePoint, para não sujar bases_mensais.
_CACHE_DIR = (
    Path(os.environ.get("APPDATA") or Path.home())
    / "DataHub" / "cache_bases_mensais"
)


def _hash_grupo(df: pd.DataFrame, chave) -> str:
    """Hash estável do conteúdo do mês (ordena pela chave antes)."""
    d = df.copy()
    ignorar = [c for c in _HASH_IGNORAR if c in d.columns]
    if ignorar:
        d = d.drop(columns=ignorar)
    cols = [chave] if isinstance(chave, str) else (list(chave) if chave else [])
    if cols and all(c in d.columns for c in cols):
        d = d.sort_values(cols, kind="mergesort")
    h = hashlib.sha256()
    h.update(pd.util.hash_pandas_object(d, index=False).values.tobytes())
    return h.hexdigest()


def _concat_dados(a: pd.DataFrame, b: pd.DataFrame) -> pd.DataFrame:
    """Concatena por posição, sem alinhar rótulos.

    A Tabela Completa tem colunas duplicadas (Data, Utilização, ...). O
    pd.concat tenta alinhar esses rótulos e lança InvalidIndexError. Aqui
    usamos numpy e mantemos os cabeçalhos (inclusive duplicados) de `b`.
    """
    if a is None or len(a) == 0:
        return b
    if b is None or len(b) == 0:
        return a
    if len(a.columns) != len(b.columns):
        # Fallback para schemas diferentes.
        return pd.concat([a, b], ignore_index=True)
    arr = np.vstack([
        a.to_numpy(dtype=object),
        b.to_numpy(dtype=object),
    ])
    return pd.DataFrame(arr, columns=b.columns)


def _upsert_mes(
    caminho: Path,
    novos: pd.DataFrame,
    chave,
    log_fn=print,
    limpar_abertos: bool = False,
) -> tuple[int, int, int]:
    """
    Faz upsert de `novos` no arquivo `caminho`.
    `chave` pode ser uma coluna (str), várias colunas (list) ou None.
    Retorna (total_final, adicionados, atualizados).
    """
    t0 = time.time()
    chave_novos = _chave_series(novos, chave)

    if chave_novos is None:
        # Sem chave: usa linha inteira para dedup
        if caminho.exists():
            existentes = pd.read_excel(str(caminho), dtype=str)
            combinado = _concat_dados(existentes, novos)
            antes = len(combinado)
            combinado = combinado.drop_duplicates(keep="last")
            adicionados = antes - len(combinado)
        else:
            combinado = novos.drop_duplicates(keep="last")
            adicionados = len(combinado)
        atualizados = 0
    else:
        # Dedup dentro do próprio lote novo (o Spotfire pode repetir a chave).
        duplicados = int(chave_novos.duplicated(keep="last").sum())
        if duplicados:
            log_fn(f"    🧹 {duplicados} linha(s) duplicada(s) no lote — mantendo a última")
            novos = novos.loc[~chave_novos.duplicated(keep="last")].copy()
            chave_novos = _chave_series(novos, chave)

        if caminho.exists():
            existentes = pd.read_excel(str(caminho), dtype=str)
            chave_existentes = _chave_series(existentes, chave)
            if chave_existentes is None:
                log_fn("    ⚠️  Chave não encontrada no arquivo existente; dedup por linha inteira")
                combinado = _concat_dados(existentes, novos).drop_duplicates(keep="last")
                adicionados = max(0, len(combinado) - len(existentes))
                atualizados = 0
            else:
                ids_novos = set(chave_novos)
                mantidos = existentes[~chave_existentes.isin(ids_novos)]
                atualizados = len(existentes) - len(mantidos)
                combinado = _concat_dados(mantidos, novos)
                adicionados = len(novos) - atualizados
        else:
            combinado = novos
            adicionados = len(combinado)
            atualizados = 0

    if limpar_abertos:
        combinado, removidos = descartar_abertos_antigos(combinado)
        if removidos:
            log_fn(f"    🧹 {removidos} incidente(s) abertos há mais de {DIAS_ABERTO_BASE} dias saíram da base")

    tmp_path = caminho.with_suffix(".tmp.xlsx")
    combinado.to_excel(str(tmp_path), index=False)
    ultimo_erro = None
    for tentativa in range(10):
        try:
            os.replace(str(tmp_path), str(caminho))
            break
        except OSError as e:
            ultimo_erro = e
            time.sleep(1)
    else:
        raise ultimo_erro
    log_fn(
        f"    ⏱️ {caminho.name}: {len(combinado)} linhas "
        f"gravadas em {time.time() - t0:.1f}s"
    )
    return len(combinado), adicionados, atualizados


def atualizar_base_mensal(
    df: pd.DataFrame,
    relatorio: str,
    pasta_local: str,
    col_data: str = None,
    log_fn=print,
    cache_hash: bool = False,
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
    if cache_hash:
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)

    for mes, grupo in df_valido.groupby("_ano_mes"):
        grupo = grupo.drop(columns=["_data_parsed", "_ano_mes"], errors="ignore")
        caminho = base_dir / f"{mes}.xlsx"
        hash_novo = _hash_grupo(grupo, chave) if cache_hash else ""
        sidecar = _CACHE_DIR / f"{relatorio}_{mes}.json"

        if cache_hash:
            # Migração: remove JSON antigo que ficava ao lado do XLSX.
            legacy = caminho.with_suffix(".json")
            if legacy.exists():
                try:
                    legacy.unlink()
                except OSError:
                    pass

        if cache_hash and caminho.exists() and sidecar.exists():
            try:
                meta = json.loads(sidecar.read_text(encoding="utf-8"))
                if meta.get("hash") == hash_novo:
                    log_fn(
                        f"  ⏭️  [{relatorio}] {mes}: sem alterações "
                        f"({meta.get('total', '?')} linhas) — XLSX não reescrito"
                    )
                    continue
            except Exception:
                pass

        total, adicionados, atualizados = _upsert_mes(caminho, grupo, chave, log_fn, limpar_abertos=(relatorio == "incidencias"))
        modificados.append(caminho)
        if cache_hash:
            try:
                sidecar.write_text(
                    json.dumps({"hash": hash_novo, "total": total}),
                    encoding="utf-8",
                )
            except Exception:
                pass
        if atualizados > 0:
            log_fn(f"  🔄 [{relatorio}] {mes}: {total} linhas ({adicionados} novos, {atualizados} atualizados)")
        else:
            log_fn(f"  ➕ [{relatorio}] {mes}: {total} linhas ({adicionados} novos)")

    return modificados
