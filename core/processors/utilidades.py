"""
Ferramentas de processamento avulso para a aba de Tratamento.
"""
import csv
import pandas as pd
from pathlib import Path
import openpyxl


def _save_xlsx_zip64(df, caminho):
    """Salva DataFrame como XLSX com suporte ZIP64 para arquivos >4GB.
    
    O pandas/ExcelWriter não expõe use_zip64 como engine_kwarg
    no openpyxl 3.1.x, então criamos o workbook manualmente.
    """
    from openpyxl import Workbook
    wb = Workbook(write_only=True)
    wb.use_zip64 = True
    ws = wb.create_sheet()
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append([None if pd.isna(v) else v for v in row])
    wb.save(caminho)


def xlsx_para_csv(caminho_xlsx: str, caminho_csv: str = None, sep: str = ";", log_fn=print) -> str:
    """Converte arquivo XLSX para CSV."""
    src = Path(caminho_xlsx)
    if not src.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {src}")

    if caminho_csv is None:
        caminho_csv = str(src.with_suffix(".csv"))

    log_fn("  📖 Lendo arquivo XLSX...")
    df = pd.read_excel(src, dtype=str)
    log_fn(f"  ✅ {len(df)} linhas carregadas")

    log_fn("  📝 Gerando CSV...")
    df.to_csv(caminho_csv, sep=sep, index=False, encoding="utf-8-sig")
    log_fn(f"  💾 Salvo: {caminho_csv}")
    return caminho_csv


def csv_para_xlsx(caminho_csv: str, caminho_xlsx: str = None, sep: str = ";", log_fn=print) -> str:
    """Converte arquivo CSV para XLSX."""
    src = Path(caminho_csv)
    if not src.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {src}")

    if caminho_xlsx is None:
        caminho_xlsx = str(src.with_suffix(".xlsx"))

    log_fn("  📖 Lendo arquivo CSV...")
    try:
        df = pd.read_csv(src, sep=sep, dtype=str, encoding='utf-8')
    except UnicodeDecodeError:
        log_fn("  ⚠️  Encoding utf-8 falhou — tentando latin-1...")
        df = pd.read_csv(src, sep=sep, dtype=str, encoding='latin-1')
    log_fn(f"  ✅ {len(df)} linhas carregadas")

    log_fn("  📝 Gerando XLSX...")
    _save_xlsx_zip64(df, caminho_xlsx)
    log_fn(f"  💾 Salvo: {caminho_xlsx}")
    return caminho_xlsx


def merge_arquivos(lista_caminhos: list[str], caminho_saida: str,
                   col_chave: str = None, log_fn=print) -> str:
    """
    Junta 2+ arquivos com mesmas colunas em um único arquivo.
    Se col_chave for informada, faz dedup pela chave (última versão vence).
    """
    if not lista_caminhos or len(lista_caminhos) < 2:
        raise ValueError("Informe pelo menos 2 arquivos para merge")

    log_fn(f"  📦 Merge: {len(lista_caminhos)} arquivos para unir")

    dfs = []
    for i, caminho in enumerate(lista_caminhos, 1):
        p = Path(caminho)
        if not p.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {p}")

        log_fn(f"  📄 [{i}/{len(lista_caminhos)}] Lendo: {p.name}")
        if p.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(p, dtype=str)
        elif p.suffix.lower() == ".csv":
            try:
                df = pd.read_csv(p, dtype=str)
            except UnicodeDecodeError:
                df = pd.read_csv(p, dtype=str, encoding='latin-1')
        else:
            raise ValueError(f"Formato não suportado: {p.suffix}")
        log_fn(f"     → {len(df)} linhas")
        dfs.append(df)

    log_fn("  🔗 Combinando arquivos...")
    combinado = pd.concat(dfs, ignore_index=True)
    log_fn(f"  ✅ {len(combinado)} linhas no total")

    if col_chave and col_chave in combinado.columns:
        antes = len(combinado)
        log_fn(f"  🧹 Removendo duplicatas (chave: {col_chave})...")
        combinado = combinado.drop_duplicates(subset=[col_chave], keep="last")
        removidos = antes - len(combinado)
        log_fn(f"  🗑️  {removidos} duplicatas removidas")

    saida = Path(caminho_saida)
    log_fn(f"  📝 Gerando arquivo final: {saida.name}")
    if saida.suffix.lower() in (".xlsx", ".xls"):
        _save_xlsx_zip64(combinado, saida)
    elif saida.suffix.lower() == ".csv":
        combinado.to_csv(saida, index=False, encoding="utf-8-sig")
    else:
        saida = saida.with_suffix(".xlsx")
        _save_xlsx_zip64(combinado, saida)

    log_fn(f"  💾 Merge concluído: {saida} ({len(combinado)} linhas)")
    return str(saida)
