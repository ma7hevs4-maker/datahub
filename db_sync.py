"""
db_sync.py — grava o DataFrame tratado do DataHub no PostgreSQL (upsert por Número).

Uso no app.py (depois de gerar `df_full`, igual ao SharePoint/n8n, dentro de try/except):

    from db_sync import gravar_incidencias
    try:
        gravar_incidencias(df_full, log_fn=self._log)
    except Exception as e:
        self._log(f"❌ Banco falhou (o fluxo continua): {e}")

Conexão (nesta ordem): variável de ambiente DATABASE_URL, ou o arquivo %APPDATA%\\DataHub\\database_url.txt
(uma linha: postgresql://usuario:senha@localhost:5432/campus). Se nenhum existir, o banco é simplesmente ignorado.
Dependência: pip install psycopg2-binary

Reincidência: depois de gravar, recalcula Reincidente / Reincidente tipo no banco sobre o MÊS INTEIRO
(função recalcular_reincidencia do schema.sql), porque o DataHub só baixa 7 dias.
"""
import hashlib
import hmac
import os
import secrets
import unicodedata
from pathlib import Path

import pandas as pd
import psycopg2
from psycopg2.extras import execute_values

POLOS = {"CAMPOS", "LAGOS", "MACAE", "NOROESTE", "MAGE", "NITEROI", "SAO GONCALO", "SERRANA", "SUL"}

# (nome na planilha, nome no banco, tipo) — 66 colunas; as demais da planilha NÃO vão para o banco.
# Embutido aqui (e não em arquivo .json) para funcionar dentro do executável do PyInstaller.
COLUNAS = [
    ('Número', 'numero', 'TEXT PRIMARY KEY'),
    ('NT', 'nt', 'TEXT'),
    ('Data Início', 'data_inicio', 'TIMESTAMP'),
    ('Data Fim', 'data_fim', 'TIMESTAMP'),
    ('Data Prev Atendimento', 'data_prev_atendimento', 'TIMESTAMP'),
    ('Duração', 'duracao', 'TEXT'),
    ('Polo', 'polo', 'TEXT'),
    ('Sucursal', 'sucursal', 'TEXT'),
    ('Causa', 'causa', 'TEXT'),
    ('Clts Atual', 'clts_atual', 'INTEGER'),
    ('Norm. < 3 min', 'norm_3_min', 'INTEGER'),
    ('Cli Af 3 min', 'cli_af_3_min', 'INTEGER'),
    ('Cli Af Max', 'cli_af_max', 'INTEGER'),
    ('Cli x Tmp', 'cli_x_tmp', 'DOUBLE PRECISION'),
    ('Conjunto', 'conjunto', 'TEXT'),
    ('CD', 'cd', 'TEXT'),
    ('AL', 'al', 'TEXT'),
    ('Ponto Elétrico', 'ponto_eletrico', 'TEXT'),
    ('Prioridade', 'prioridade', 'TEXT'),
    ('Avisos', 'avisos', 'INTEGER'),
    ('Tipo Reclamação', 'tipo_reclamacao', 'TEXT'),
    ('Callback', 'callback', 'TEXT'),
    ('Resultado Ligação (Produtivo)', 'resultado_ligacao_produtivo', 'TEXT'),
    ('Motivo (Produtivo)', 'motivo_produtivo', 'TEXT'),
    ('Tipo', 'tipo', 'TEXT'),
    ('Equipe Atribuída', 'equipe_atribuida', 'TEXT'),
    ('Equipe Desl.', 'equipe_desl', 'TEXT'),
    ('TMP', 'tmp', 'DOUBLE PRECISION'),
    ('TMD', 'tmd', 'DOUBLE PRECISION'),
    ('TME', 'tme', 'DOUBLE PRECISION'),
    ('TMA', 'tma', 'DOUBLE PRECISION'),
    ('T. Primeira Manobra', 't_primeira_manobra', 'INTEGER'),
    ('Cliente Essencial', 'cliente_essencial', 'INTEGER'),
    ('OSM', 'osm', 'TEXT'),
    ('Mês Saída', 'mes_saida', 'SMALLINT'),
    ('Dia Saída', 'dia_saida', 'SMALLINT'),
    ('Duração em minutos', 'duracao_em_minutos', 'INTEGER'),
    ('Status', 'status', 'TEXT'),
    ('Considerar TMA', 'considerar_tma', 'TEXT'),
    ('Equipe/Operador', 'equipe_operador', 'TEXT'),
    ('Cluster duração', 'cluster_duracao', 'TEXT'),
    ('Atribuição', 'atribuicao', 'TEXT'),
    ('Hora entrada', 'hora_entrada', 'SMALLINT'),
    ('Turno abertura', 'turno_abertura', 'TEXT'),
    ('Hora saída', 'hora_saida', 'TEXT'),
    ('Turno fechamento', 'turno_fechamento', 'TEXT'),
    ('Produtivo/Improdutivo', 'produtivo_improdutivo', 'TEXT'),
    ('TME Análise', 'tme_analise', 'TEXT'),
    ('Data Atuação', 'data_atuacao', 'TIMESTAMP'),
    ('Hora Atuação', 'hora_atuacao', 'TEXT'),
    ('Ponto Elétrico (Dispositivo)', 'ponto_eletrico_dispositivo', 'TEXT'),
    ('Nivel de Tensão', 'nivel_de_tensao', 'TEXT'),
    ('Cluster clientes', 'cluster_clientes', 'TEXT'),
    ('Sucursal Real', 'sucursal_real', 'TEXT'),
    ('Processo (Equipe atrib.)', 'processo_equipe_atrib', 'TEXT'),
    ('Enel/Parceira (Equipe atrib.)', 'enel_parceira_equipe_atrib', 'TEXT'),
    ('Grupo Processos (Equipe atrib.)', 'grupo_processos_equipe_atrib', 'TEXT'),
    ('Processo (Equipe desloc.)', 'processo_equipe_desloc', 'TEXT'),
    ('Enel/Parceira (Equipe desloc.)', 'enel_parceira_equipe_desloc', 'TEXT'),
    ('Grupo Processos (Equipe desloc.)', 'grupo_processos_equipe_desloc', 'TEXT'),
    ('Reincidente', 'reincidente', 'TEXT'),
    ('Reincidente tipo', 'reincidente_tipo', 'TEXT'),
    ('Ordem 2 Consolidada', 'ordem_2_consolidada', 'TEXT'),
    ('Ano', 'ano', 'SMALLINT'),
    ('Mês', 'mes', 'SMALLINT'),
    ('Dia', 'dia', 'SMALLINT'),
    ('Observação', 'observacao', 'TEXT'),
    ('Cumpre Regras Ouro', 'cumpre_regras_ouro', 'TEXT'),
    ('ORD 2', 'ord_2', 'TEXT'),
    ('Estado', 'estado', 'TEXT'),
    ('Entrada', 'entrada', 'TEXT'),
    ('Saída', 'saida', 'TEXT'),
    ('MR', 'mr', 'TEXT'),
    ('Início MR', 'inicio_mr', 'TIMESTAMP'),
    ('Nível Compensação', 'nivel_compensacao', 'TEXT'),
    ('Compensação', 'compensacao', 'TEXT'),
    ('Status URA', 'status_ura', 'TEXT'),
    ('Resultado URA', 'resultado_ura', 'TEXT'),
    ('Resultado BOT', 'resultado_bot', 'TEXT'),
    ('Motivo BOT', 'motivo_bot', 'TEXT'),
    ('Ponto Atenção', 'ponto_atencao', 'TEXT'),
    ('Tipo Agrupamento', 'tipo_agrupamento', 'TEXT'),
    ('Área Risco', 'area_risco', 'TEXT'),
    ('Latitude', 'latitude', 'DOUBLE PRECISION'),
    ('Longitude', 'longitude', 'DOUBLE PRECISION'),
]


def _norm(s) -> str:
    return unicodedata.normalize("NFKD", str(s).upper()).encode("ascii", "ignore").decode().strip()


_SEM_CLIENTE = {"", "nan", "none", "nat", "0", "0.0", "-", "---"}


def _hash_cliente(v, segredo: bytes):
    """Código irreversível do Nº Cliente (HMAC-SHA256). Só serve para saber se é o 'mesmo cliente'."""
    t = str(v).strip()
    if t.endswith(".0"):
        t = t[:-2]
    if t.lower() in _SEM_CLIENTE:
        return None
    return hmac.new(segredo, t.encode(), hashlib.sha256).hexdigest()[:32]


def _segredo(cur) -> bytes:
    cur.execute("INSERT INTO config_interna (chave, valor) VALUES ('hmac_cliente', %s) ON CONFLICT DO NOTHING",
                (secrets.token_hex(32),))
    cur.execute("SELECT valor FROM config_interna WHERE chave = 'hmac_cliente'")
    return cur.fetchone()[0].encode()


def _datas(s: pd.Series) -> pd.Series:
    """dd/mm/aaaa hh:mm:ss (formato do Operview). Rápido com formato fixo; o que não casar tenta o modo flexível."""
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    r = pd.to_datetime(s, format="%d/%m/%Y %H:%M:%S", errors="coerce")
    falhou = r.isna() & s.notna()
    if falhou.any():
        r[falhou] = pd.to_datetime(s[falhou], dayfirst=True, errors="coerce")
    return r


ESTADOS_FORA = {"programado", "programada", "anulado", "anulada"}
CAUSAS_FORA = {"atend cancelado cr", "atendimento cancelado", "programada manutenção", "programada obras"}


def preparar(df: pd.DataFrame, segredo: bytes):
    """Filtra os 9 polos e converte tipos. Retorna (df_pronto, descartadas)."""
    total = len(df)
    polo_norm = df["Polo"].map(_norm)
    df = df[polo_norm.isin(POLOS)].copy()
    descartadas = total - len(df)

    # Fora do banco (mesma regra do dashboard HTML): Estado programado/anulado e causas canceladas/programadas.
    est = df["Estado"].astype(str).str.strip().str.lower() if "Estado" in df.columns else pd.Series("", index=df.index)
    cau = df["Causa"].astype(str).str.strip().str.lower() if "Causa" in df.columns else pd.Series("", index=df.index)
    excluir = est.isin(ESTADOS_FORA) | cau.isin(CAUSAS_FORA)
    excluidos = df.loc[excluir, "Número"].dropna().astype(str).str.strip().tolist() if "Número" in df.columns else []
    descartadas += int(excluir.sum())
    df = df[~excluir].copy()

    out = pd.DataFrame(index=df.index)
    for orig, nome, tipo in COLUNAS:
        s = df[orig] if orig in df.columns else pd.Series(None, index=df.index, dtype=object)
        if s.dtype == object or str(s.dtype).startswith("str"):
            s = s.where(~s.astype(str).str.strip().isin(["---", "-", "nan", "NaT", "None", ""]), None)
        if tipo.startswith("TIMESTAMP"):
            s = _datas(s)
        elif tipo in ("INTEGER", "SMALLINT"):
            s = pd.to_numeric(s, errors="coerce").astype("Int64")
        elif tipo.startswith("DOUBLE"):
            s = pd.to_numeric(s, errors="coerce")
        else:
            s = s.astype("object").where(s.notna(), None)
            s = s.map(lambda v: None if v is None else str(v))
        out[nome] = s
    # Operview usa -222077236 / 9090669639 como "sem coordenada": vira NULL.
    out.loc[out["latitude"].abs() > 90, "latitude"] = None
    out.loc[out["longitude"].abs() > 180, "longitude"] = None
    # Arquivos antigos não têm Mês/Dia Saída: deriva da Data Fim.
    for nome, atr in (("mes_saida", "month"), ("dia_saida", "day")):
        vazio = out[nome].isna()
        if vazio.any():
            out.loc[vazio, nome] = getattr(out.loc[vazio, "data_fim"].dt, atr).astype("Int64")
    col_cli = next((c for c in ["Nº Cliente", "UC", "Cliente", "N_Cliente"] if c in df.columns), None)
    out["cliente_hash"] = df[col_cli].map(lambda v: _hash_cliente(v, segredo)) if col_cli else None
    out["polo_regional"] = df["Polo"].astype(str).str.strip()
    out = out[out["numero"].notna()].drop_duplicates(subset="numero", keep="last")
    out.attrs["excluidos"] = excluidos
    return out, descartadas


def _py(v):
    if v is None or v is pd.NaT or (isinstance(v, float) and pd.isna(v)) or v is pd.NA:
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item"):
        return v.item()
    return v


def obter_dsn():
    dsn = os.environ.get("DATABASE_URL", "").strip()
    if dsn:
        return dsn
    appdata = os.environ.get("APPDATA")
    if appdata:
        arq = Path(appdata) / "DataHub" / "database_url.txt"
        if arq.exists():
            return arq.read_text(encoding="utf-8-sig").strip()
    return ""


def gravar_incidencias(df: pd.DataFrame, dsn: str = None, log_fn=print, lote: int = 5000) -> int:
    dsn = dsn or obter_dsn()
    if not dsn:
        log_fn("  ⏭️  Banco não configurado (sem DATABASE_URL / database_url.txt) — ignorado.")
        return 0
    recebidas = len(df)
    conn = psycopg2.connect(dsn)
    try:
        with conn, conn.cursor() as cur:           # uma transação: ou grava tudo ou nada
            for _o, _n, _t in COLUNAS:                  # migração automática: cria colunas novas
                if "PRIMARY KEY" not in _t:
                    cur.execute(f"ALTER TABLE incidencias ADD COLUMN IF NOT EXISTS {_n} {_t}")
            pronto, descartadas = preparar(df, _segredo(cur))
            excl = pronto.attrs.get("excluidos") or []
            apagadas = 0
            if excl:                                    # incidente que virou programado/anulado/cancelado sai do banco
                cur.execute("DELETE FROM incidencias WHERE numero = ANY(%s)", (excl,))
                apagadas = cur.rowcount
            nomes = list(pronto.columns)
            sets = ", ".join(f"{n}=EXCLUDED.{n}" for n in nomes if n != "numero")
            sql = (f"INSERT INTO incidencias ({', '.join(nomes)}, carregado_em) VALUES %s "
                   f"ON CONFLICT (numero) DO UPDATE SET {sets}, carregado_em=now()")
            tpl = "(" + ",".join(["%s"] * len(nomes)) + ", now())"
            registros = [tuple(_py(v) for v in row) for row in pronto.itertuples(index=False, name=None)]
            for i in range(0, len(registros), lote):
                execute_values(cur, sql, registros[i:i + lote], template=tpl, page_size=lote)
            # Reincidência depende do mês inteiro (não só dos 7 dias baixados): recalcula no banco.
            meses = sorted({d.replace(day=1).date() for d in pronto["data_inicio"].dropna()})
            alterados = 0
            for m in meses:
                cur.execute("SELECT recalcular_reincidencia(%s)", (m,))
                alterados += cur.fetchone()[0]
            cur.execute(
                "INSERT INTO cargas (linhas_recebidas, linhas_gravadas, linhas_descartadas) VALUES (%s,%s,%s)",
                (recebidas, len(registros), descartadas))
        log_fn(f"  🗄️  Banco: {len(registros)} linhas gravadas ({descartadas} descartadas: fora dos 9 polos, Estado ou Causa); "
               f"reincidência recalculada em {len(meses)} mês(es), {alterados} linha(s) ajustada(s)"
               + (f"; {apagadas} removida(s) por Estado/Causa" if apagadas else ""))
        return len(registros)
    except Exception as e:
        try:
            with psycopg2.connect(dsn) as c2, c2.cursor() as cur2:
                cur2.execute("INSERT INTO cargas (linhas_recebidas, ok, erro) VALUES (%s, false, %s)",
                             (recebidas, str(e)[:500]))
        except Exception:
            pass
        raise
    finally:
        conn.close()
