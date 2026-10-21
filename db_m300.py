"""
db_m300.py — grava no PostgreSQL as tabelas do M300 geradas pelo DataHub.

Tabelas (criadas sozinhas, se não existirem, no primeiro uso — não precisa rodar SQL à mão):
  deslocamentos_hoje   foto do dia (m300_deslocamentos_hoje.csv). Substituída inteira a cada ciclo.
  deslocamentos_total  um registro por deslocamento (bases_mensais/deslocamentos/AAAA-MM.xlsx). Upsert.
  tabela_completa      um registro por equipe-dia (bases_mensais/tabela_completa/AAAA-MM.xlsx). Upsert.

Uso no app.py, junto de onde já grava as incidências (cada chamada dentro de try/except; o fluxo nunca para):

    from db_m300 import gravar_deslocamentos_hoje, gravar_deslocamentos_total, gravar_tabela_completa
    gravar_deslocamentos_hoje(df_hoje, log_fn=self._log)         # DataFrame do CSV de hoje (ou caminho do .csv)
    gravar_deslocamentos_total(df_deslocamentos, log_fn=self._log)
    gravar_tabela_completa(df_tabela_completa, log_fn=self._log)

Aceita o DataFrame com as colunas originais (como estão nos arquivos). Coluna que não existir vira NULL;
coluna nova que não está no esquema é ignorada e avisada no log.
Conexão: a mesma do db_sync (DATABASE_URL ou %APPDATA%\\DataHub\\database_url.txt).
"""
import re
import unicodedata
from datetime import date, datetime, time as dtime

import pandas as pd
import psycopg2
from psycopg2.extras import execute_values

from db_sync import obter_dsn, _py

# (nome no arquivo, nome no banco, tipo)  tipos: text, num (double), int, date, ts (timestamp), time
COLUNAS_DESL = [
    ('Data Referência', 'data_referencia', 'date'),
    ('Equipe', 'equipe', 'text'),
    ('Inicio Calendario', 'inicio_calendario', 'ts'),
    ('Log In', 'log_in', 'ts'),
    ('Fim Calendario', 'fim_calendario', 'ts'),
    ('Log Off', 'log_off', 'ts'),
    ('1º Login', 'primeiro_login', 'num'),
    ('Inicio Intervalo', 'inicio_intervalo', 'ts'),
    ('Fim Intervalo', 'fim_intervalo', 'ts'),
    ('Intervalo', 'intervalo', 'int'),
    ('Base', 'base', 'text'),
    ('UO', 'uo', 'text'),
    ('Tipo Equipe', 'tipo_equipe', 'text'),
    ('Período', 'periodo', 'text'),
    ('FiltroRepetidoEquipeData', 'filtrorepetidoequipedata', 'num'),
    ('Nr_Ordem', 'nr_ordem', 'text'),
    ('Despachada', 'despachada', 'ts'),
    ('A_Caminho', 'a_caminho', 'ts'),
    ('Liberada', 'liberada', 'ts'),
    ('No_Local', 'no_local', 'ts'),
    ('HT Ordem', 'ht_ordem', 'num'),
    ('CLASSE', 'classe', 'text'),
    ('CAUSA', 'causa', 'text'),
    ('TR Ordem', 'tr_ordem', 'num'),
    ('TL Ordem', 'tl_ordem', 'num'),
    ('TR Ordem Imp SS', 'tr_ordem_imp_ss', 'num'),
    ('status', 'status', 'text'),
    ('tempo_padrao', 'tempo_padrao', 'num'),
    ('VerificaRepetidosHT', 'verificarepetidosht', 'num'),
    ('Conta Task Time tr', 'conta_task_time_tr', 'num'),
    ('Conta task Time TL', 'conta_task_time_tl', 'num'),
    ('Qtd Equipes OS', 'qtd_equipes_os', 'num'),
    ('Hora 1º Deslocamento', 'hora_1o_deslocamento', 'ts'),
    ('Hora 1º Despacho', 'hora_1o_despacho', 'ts'),
    ('Hora Ultima Ordem', 'hora_ultima_ordem', 'ts'),
    ('HT total', 'ht_total', 'num'),
    ('HL TOTAL', 'hl_total', 'num'),
    ('TR TOTAL', 'tr_total', 'num'),
    ('TL TOTAL', 'tl_total', 'num'),
    ('HP Ordem', 'hp_ordem', 'num'),
    ('Qtd Serviços', 'qtd_servicos', 'num'),
    ('Qtd Deslocamentos', 'qtd_deslocamentos', 'num'),
    ('Soma De Desl/Ativ', 'soma_de_desl_ativ', 'num'),
    ('Task Time Total TR', 'task_time_total_tr', 'num'),
    ('Task Time Total TL', 'task_time_total_tl', 'num'),
    ('Média Qtd Serviços', 'media_qtd_servicos', 'num'),
    ('TR Ordem Imp SS equipe', 'tr_ordem_imp_ss_equipe', 'num'),
    ('1º Desloc', 'primeiro_desloc', 'num'),
    ('1º Despacho', 'primeiro_despacho', 'num'),
    ('HP Total', 'hp_total', 'num'),
    ('Log In Corrigido', 'log_in_corrigido', 'ts'),
    ('Log Off Corrigido', 'log_off_corrigido', 'ts'),
    ('1º Login Corrigido', 'primeiro_login_corrigido', 'num'),
    ('HD Total', 'hd_total', 'num'),
    ('Retorno a base', 'retorno_a_base', 'num'),
    ('AtuaçãoHD', 'atuacaohd', 'text'),
    ('Horas Extras', 'horas_extras', 'num'),
    ('Desvios', 'desvios', 'text'),
    ('Ano', 'ano', 'int'),
    ('Mês', 'mes', 'text'),
    ('Semana', 'semana', 'int'),
    ('TR Ordem Secundário', 'tr_ordem_secundario', 'num'),
    ('TR_TOTAL_CAL', 'tr_total_cal', 'num'),
    ('TEMPO_PADRAO_TOTAL_CAL', 'tempo_padrao_total_cal', 'num'),
]
COLUNAS_TC = [
    ('Data Referência', 'data_referencia', 'date'),
    ('Equipe', 'equipe', 'text'),
    ('Inicio Calendario', 'inicio_calendario', 'ts'),
    ('Log In', 'log_in', 'ts'),
    ('Fim Calendario', 'fim_calendario', 'ts'),
    ('Log Off', 'log_off', 'ts'),
    ('1º Login', 'primeiro_login', 'num'),
    ('Inicio Intervalo', 'inicio_intervalo', 'ts'),
    ('Fim Intervalo', 'fim_intervalo', 'ts'),
    ('Intervalo', 'intervalo', 'int'),
    ('Base', 'base', 'text'),
    ('UO', 'uo', 'text'),
    ('Tipo Equipe', 'tipo_equipe', 'text'),
    ('Período', 'periodo', 'text'),
    ('FiltroRepetidoEquipeData', 'filtrorepetidoequipedata', 'num'),
    ('Nr_Ordem', 'nr_ordem', 'text'),
    ('Despachada', 'despachada', 'ts'),
    ('A_Caminho', 'a_caminho', 'ts'),
    ('Liberada', 'liberada', 'ts'),
    ('No_Local', 'no_local', 'ts'),
    ('HT Ordem', 'ht_ordem', 'num'),
    ('CLASSE', 'classe', 'text'),
    ('CAUSA', 'causa', 'text'),
    ('TR Ordem', 'tr_ordem', 'num'),
    ('TL Ordem', 'tl_ordem', 'num'),
    ('TR Ordem Imp SS', 'tr_ordem_imp_ss', 'num'),
    ('status', 'status', 'text'),
    ('tempo_padrao', 'tempo_padrao', 'num'),
    ('VerificaRepetidosHT', 'verificarepetidosht', 'num'),
    ('Conta Task Time tr', 'conta_task_time_tr', 'num'),
    ('Conta task Time TL', 'conta_task_time_tl', 'num'),
    ('Qtd Equipes OS', 'qtd_equipes_os', 'num'),
    ('Hora 1º Deslocamento', 'hora_1o_deslocamento', 'ts'),
    ('Hora 1º Despacho', 'hora_1o_despacho', 'ts'),
    ('Hora Ultima Ordem', 'hora_ultima_ordem', 'ts'),
    ('HT total', 'ht_total', 'num'),
    ('HL TOTAL', 'hl_total', 'num'),
    ('TR TOTAL', 'tr_total', 'num'),
    ('TL TOTAL', 'tl_total', 'num'),
    ('HP Ordem', 'hp_ordem', 'num'),
    ('Qtd Serviços', 'qtd_servicos', 'num'),
    ('Qtd Deslocamentos', 'qtd_deslocamentos', 'num'),
    ('Soma De Desl/Ativ', 'soma_de_desl_ativ', 'num'),
    ('Task Time Total TR', 'task_time_total_tr', 'num'),
    ('Task Time Total TL', 'task_time_total_tl', 'num'),
    ('Média Qtd Serviços', 'media_qtd_servicos', 'num'),
    ('TR Ordem Imp SS equipe', 'tr_ordem_imp_ss_equipe', 'num'),
    ('1º Desloc', 'primeiro_desloc', 'num'),
    ('1º Despacho', 'primeiro_despacho', 'num'),
    ('HP Total', 'hp_total', 'num'),
    ('Log In Corrigido', 'log_in_corrigido', 'ts'),
    ('Log Off Corrigido', 'log_off_corrigido', 'ts'),
    ('1º Login Corrigido', 'primeiro_login_corrigido', 'num'),
    ('HD Total', 'hd_total', 'num'),
    ('Retorno a base', 'retorno_a_base', 'num'),
    ('Horas Extras', 'horas_extras', 'num'),
    ('Desvios', 'desvios', 'text'),
    ('Ano', 'ano', 'int'),
    ('Mês', 'mes', 'text'),
    ('Semana', 'semana', 'int'),
    ('AtuaçãoHD', 'atuacaohd', 'text'),
    ('Plataforma', 'plataforma', 'num'),
    ('Tipo_Classe', 'tipo_classe', 'text'),
    ('DATA_ETL', 'data_etl', 'ts'),
    ('TR Ordem Secundário', 'tr_ordem_secundario', 'num'),
    ('TR Ordem Sencundário equipe', 'tr_ordem_sencundario_equipe', 'num'),
    ('TR_TOTAL_CAL', 'tr_total_cal', 'num'),
    ('TEMPO_PADRAO_TOTAL_CAL', 'tempo_padrao_total_cal', 'num'),
    ('Dif Deslocamento e Inicio do calendario', 'dif_deslocamento_e_inicio_do_calendario', 'text'),
    ('Deslocamento antes do calendario', 'deslocamento_antes_do_calendario', 'text'),
    ('prefixo processo', 'prefixo_processo', 'text'),
    ('Processo', 'processo', 'text'),
    ('Enel/Parceira', 'enel_parceira', 'text'),
    ('Turno ', 'turno', 'text'),
    ('Data', 'data', 'date'),
    ('Data.1', 'data_1', 'num'),
    ('EQUIPE', 'equipe_curta', 'text'),
    ('POLO', 'polo', 'text'),
    ('x', 'x', 'date'),
    ('chave', 'chave', 'text'),
    ('Utilização', 'utilizacao', 'num'),
    ('Task Time ', 'task_time', 'num'),
    ('Calc_eficiencia', 'calc_eficiencia', 'num'),
    ('mês_calculo', 'mes_calculo', 'int'),
    ('Classificação', 'classificacao', 'num'),
    ('Ativ/Equipe/Dia', 'ativ_equipe_dia', 'num'),
    ('Produtividade', 'produtividade', 'num'),
    ('Eficiencia', 'eficiencia', 'num'),
    ('Eficiencia Real', 'eficiencia_real', 'num'),
    ('Qtd Serv IIº', 'qtd_serv_iio', 'num'),
    ('Utilização.1', 'utilizacao_1', 'num'),
    ('1º Login.1', 'primeiro_login_1', 'num'),
    ('1º Desloc.1', 'primeiro_desloc_1', 'num'),
    ('1º Despacho.1', 'primeiro_despacho_1', 'num'),
    ('Plataforma.1', 'plataforma_1', 'num'),
    ('Intervalo.1', 'intervalo_1', 'num'),
    ('Retorno a Base', 'retorno_a_base_2', 'num'),
    ('Task Time TR', 'task_time_tr', 'num'),
    ('TMR Secundário', 'tmr_secundario', 'num'),
    ('TMR Improd.', 'tmr_improd', 'num'),
    ('TML', 'tml', 'num'),
    ('H.E.(Min.)', 'h_e_min', 'num'),
]
COLUNAS_HOJE = [
    ('Dia', 'dia', 'date'),
    ('Equipe', 'equipe', 'text'),
    ('Ordem', 'ordem', 'text'),
    ('Despachado', 'despachado', 'ts'),
    ('A Caminho', 'a_caminho', 'ts'),
    ('No Local', 'no_local', 'ts'),
    ('liberada', 'liberada', 'ts'),
    ('InicioOS', 'inicioos', 'ts'),
    ('fimOs', 'fimos', 'ts'),
    ('Qtd Deslocamentos', 'qtd_deslocamentos', 'int'),
    ('Horas Trabalhadas', 'horas_trabalhadas', 'int'),
    ('Em Atendimento', 'em_atendimento', 'int'),
    ('Inicio Calendário', 'inicio_calendario', 'time'),
    ('Login', 'login', 'time'),
    ('Fim calendário', 'fim_calendario', 'time'),
    ('Log Off', 'log_off', 'time'),
    ('GPS', 'gps', 'text'),
    ('Ext Calendário', 'ext_calendario', 'text'),
]

PK = {
    "deslocamentos_total": ["data_referencia", "equipe", "nr_ordem", "despachada"],
    "tabela_completa": ["data_referencia", "equipe"],
}

# Polo pelo nome da base (deslocamentos) e pelo prefixo da equipe (CSV de hoje, que não traz base).
_POLO_NOME = {"Lagos": "LAGOS", "Serrana": "SERRANA", "São Gonçalo": "SÃO GONÇALO", "Magé": "MAGÉ", "Sul": "SUL",
              "Noroeste": "NOROESTE", "Niterói": "NITERÓI", "Campos": "CAMPOS", "Macaé": "MACAÉ"}
BASE_POLO = {"ARARUAMA": "LAGOS", "CABO FRIO": "LAGOS", "PETRÓPOLIS": "SERRANA", "TERESÓPOLIS": "SERRANA",
             "SÃO GONÇALO": "SÃO GONÇALO", "MAGÉ": "MAGÉ", "ANGRA": "SUL", "RESENDE": "SUL",
             "CANTAGALO": "NOROESTE", "ITAPERUNA": "NOROESTE", "PÁDUA": "NOROESTE",
             "NITERÓI": "NITERÓI", "CAMPOS": "CAMPOS", "MACAÉ": "MACAÉ"}
PREFIXO_POLO = {"AR": "LAGOS", "CF": "LAGOS", "PE": "SERRANA", "TE": "SERRANA", "SG": "SÃO GONÇALO", "MG": "MAGÉ",
                "AN": "SUL", "RE": "SUL", "CG": "NOROESTE", "IT": "NOROESTE", "PD": "NOROESTE",
                "NI": "NITERÓI", "CP": "CAMPOS", "MC": "MACAÉ"}

_SQL_TIPO = {"text": "TEXT", "num": "DOUBLE PRECISION", "int": "INTEGER", "date": "DATE", "ts": "TIMESTAMP", "time": "TIME"}


# ------------------------------------------------------------------ esquema
def _ddl(tabela, colunas, extras, pk=None, serial=False):
    cols = [f"    {n} {_SQL_TIPO[t]}" for _, n, t in colunas] + [f"    {e}" for e in extras]
    if serial:
        cols.insert(0, "    id BIGSERIAL PRIMARY KEY")
    elif pk:
        cols.append(f"    PRIMARY KEY ({', '.join(pk)})")
    return f"CREATE TABLE IF NOT EXISTS {tabela} (\n" + ",\n".join(cols) + "\n)"


def _alters(tabela, colunas):
    """Colunas novas no esquema entram em tabelas já existentes sem perder dados."""
    return [f"ALTER TABLE {tabela} ADD COLUMN IF NOT EXISTS {n} {_SQL_TIPO[t]}" for _, n, t in colunas]


def esquema_sql():
    """Todo o SQL necessário (também usado para gerar schema_m300.sql)."""
    ext = ["polo_regional TEXT", "carregado_em TIMESTAMPTZ NOT NULL DEFAULT now()"]
    return _esquema_base(ext) + _alters("deslocamentos_hoje", COLUNAS_HOJE) + _alters("deslocamentos_total", COLUNAS_DESL) + _alters("tabela_completa", COLUNAS_TC)


def _esquema_base(ext):
    return [
        "ALTER TABLE cargas ADD COLUMN IF NOT EXISTS origem TEXT NOT NULL DEFAULT 'incidencias'",
        _ddl("deslocamentos_hoje", COLUNAS_HOJE, ext, serial=True),
        "CREATE INDEX IF NOT EXISTS ix_deshoje_equipe ON deslocamentos_hoje (equipe)",
        "CREATE INDEX IF NOT EXISTS ix_deshoje_ordem ON deslocamentos_hoje (ordem)",
        "CREATE INDEX IF NOT EXISTS ix_deshoje_polo ON deslocamentos_hoje (polo_regional)",
        _ddl("deslocamentos_total", COLUNAS_DESL, ext, pk=PK["deslocamentos_total"]),
        "CREATE INDEX IF NOT EXISTS ix_destot_data ON deslocamentos_total (data_referencia)",
        "CREATE INDEX IF NOT EXISTS ix_destot_polo_data ON deslocamentos_total (polo_regional, data_referencia)",
        "CREATE INDEX IF NOT EXISTS ix_destot_ordem ON deslocamentos_total (nr_ordem)",
        _ddl("tabela_completa", COLUNAS_TC, ext, pk=PK["tabela_completa"]),
        "CREATE INDEX IF NOT EXISTS ix_tabcomp_data ON tabela_completa (data_referencia)",
        "CREATE INDEX IF NOT EXISTS ix_tabcomp_polo_data ON tabela_completa (polo_regional, data_referencia)",
        "CREATE INDEX IF NOT EXISTS ix_tabcomp_equipe ON tabela_completa (equipe)",
    ]


def garantir_tabelas(cur):
    for sql in esquema_sql():
        cur.execute(sql)


# ------------------------------------------------------------------ conversão
def _datas(s, fmt):
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    out = pd.to_datetime(s, format=fmt, errors="coerce")
    falta = out.isna() & s.notna()
    if falta.any():
        out = out.copy()
        out[falta] = pd.to_datetime(s[falta], dayfirst=True, errors="coerce")
    return out


def _conv(s, tipo):
    """Converte uma coluna do DataFrame para o tipo do banco (lista de valores Python)."""
    if tipo == "date":
        r = _datas(s, "%d/%m/%Y").dt.date
    elif tipo == "ts":
        r = _datas(s, "%d/%m/%Y %H:%M:%S")
    elif tipo == "time":
        t = pd.to_datetime(s.astype("string").str.strip(), format="%H:%M", errors="coerce")
        r = t.dt.time
    elif tipo in ("num", "int"):
        if not pd.api.types.is_numeric_dtype(s):
            s = s.astype("string").str.strip().str.replace(",", ".", regex=False)
        r = pd.to_numeric(s, errors="coerce")
        if tipo == "int":
            r = r.round().astype("Int64")
    else:  # text
        r = s.astype("string").str.strip()
        r = r.mask(r.isin(["", "nan", "NaN", "None", "<NA>"]))
    out = []
    for v in r.tolist():
        if v is None or v is pd.NaT or v is pd.NA or (isinstance(v, float) and v != v):
            out.append(None)
        elif isinstance(v, pd.Timestamp):
            out.append(v.to_pydatetime())
        elif isinstance(v, (date, datetime, dtime, str)):
            out.append(v)
        else:
            out.append(v.item() if hasattr(v, "item") else v)
    return out


def _preparar(df, colunas, pk=None, polo_de=None, log_fn=print):
    """DataFrame original -> (lista de colunas do banco, registros, descartadas)."""
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    nomes_esp = {o for o, _, _ in colunas}
    novas = [c for c in df.columns if c not in nomes_esp]
    if novas:
        log_fn(f"  ⚠️  Colunas novas ignoradas (não existem no esquema): {novas[:8]}")
    dados = {}
    for o, n, t in colunas:
        dados[n] = _conv(df[o], t) if o in df.columns else [None] * len(df)
    out = pd.DataFrame(dados, index=range(len(df)))
    if polo_de:
        out["polo_regional"] = polo_de(df)
    else:
        out["polo_regional"] = None
    descartadas = 0
    if pk:
        ok = out[pk].notna().all(axis=1)
        descartadas = int((~ok).sum())
        out = out[ok].drop_duplicates(pk, keep="last")
    cols = list(out.columns)
    regs = [tuple(_py(v) for v in row) for row in out.itertuples(index=False, name=None)]
    return cols, regs, descartadas


def _polo_por_base(df):
    b = df["Base"].astype("string").str.strip().str.upper() if "Base" in df.columns else None
    if b is None:
        return [None] * len(df)
    return [BASE_POLO.get(x) if isinstance(x, str) else None for x in b.tolist()]


def _polo_por_nome(df):
    p = df["POLO"].astype("string").str.strip() if "POLO" in df.columns else None
    if p is None:
        return _polo_por_base(df)
    return [_POLO_NOME.get(x) if isinstance(x, str) else None for x in p.tolist()]


def _polo_por_prefixo(df):
    e = df["Equipe"].astype("string").str.strip().str.upper().str[:2]
    return [PREFIXO_POLO.get(x) if isinstance(x, str) else None for x in e.tolist()]


# ------------------------------------------------------------------ gravação
def _executar(tabela, df, colunas, log_fn, dsn, substituir=False, pk=None, polo_de=None, lote=5000):
    dsn = dsn or obter_dsn()
    if not dsn:
        log_fn(f"  ⏭️  Banco não configurado — {tabela} ignorada.")
        return 0
    if df is None or len(df) == 0:
        log_fn(f"  ⏭️  {tabela}: sem linhas, nada a gravar.")
        return 0
    recebidas = len(df)
    conn = psycopg2.connect(dsn)
    try:
        with conn, conn.cursor() as cur:  # uma transação: ou grava tudo ou nada
            garantir_tabelas(cur)
            cols, regs, descartadas = _preparar(df, colunas, pk, polo_de, log_fn)
            lista = ", ".join(cols)
            if substituir:
                cur.execute(f"DELETE FROM {tabela}")
                sql = f"INSERT INTO {tabela} ({lista}, carregado_em) VALUES %s"
            else:
                sets = ", ".join(f"{c}=EXCLUDED.{c}" for c in cols if c not in pk)
                sql = (f"INSERT INTO {tabela} ({lista}, carregado_em) VALUES %s "
                       f"ON CONFLICT ({', '.join(pk)}) DO UPDATE SET {sets}, carregado_em=now()")
            tpl = "(" + ",".join(["%s"] * len(cols)) + ", now())"
            for i in range(0, len(regs), lote):
                execute_values(cur, sql, regs[i:i + lote], template=tpl, page_size=lote)
            cur.execute("INSERT INTO cargas (origem, linhas_recebidas, linhas_gravadas, linhas_descartadas) VALUES (%s,%s,%s,%s)",
                        (tabela, recebidas, len(regs), descartadas))
        extra = f" ({descartadas} sem chave descartadas)" if descartadas else ""
        log_fn(f"  🗄️  Banco ({tabela}): {len(regs)} linhas gravadas{extra}")
        return len(regs)
    except Exception as e:
        try:
            with psycopg2.connect(dsn) as c2, c2.cursor() as cur2:
                cur2.execute("INSERT INTO cargas (origem, linhas_recebidas, ok, erro) VALUES (%s,%s,false,%s)",
                             (tabela, recebidas, str(e)[:500]))
        except Exception:
            pass
        raise
    finally:
        conn.close()


def ler_csv_hoje(caminho):
    """O CSV do M300 é texto separado por TAB em UTF-16."""
    return pd.read_csv(caminho, sep="\t", encoding="utf-16", dtype=str)


def gravar_deslocamentos_hoje(df, dsn=None, log_fn=print):
    """Foto do dia: apaga e regrava tudo numa transação. Linhas sem 'Dia' (equipe sem ordem) recebem o dia do arquivo."""
    if isinstance(df, (str, bytes)) or hasattr(df, "__fspath__"):
        df = ler_csv_hoje(df)
    if df is None or len(df) == 0:
        log_fn("  ⏭️  deslocamentos_hoje: arquivo vazio, mantendo a foto anterior.")
        return 0
    df = df.copy()
    if "Dia" in df.columns:
        dia = df["Dia"].dropna()
        if len(dia):
            df["Dia"] = df["Dia"].fillna(dia.mode().iloc[0])
    return _executar("deslocamentos_hoje", df, COLUNAS_HOJE, log_fn, dsn, substituir=True, polo_de=_polo_por_prefixo)


def gravar_deslocamentos_total(df, dsn=None, log_fn=print):
    return _executar("deslocamentos_total", df, COLUNAS_DESL, log_fn, dsn, pk=PK["deslocamentos_total"], polo_de=_polo_por_base)


def gravar_tabela_completa(df, dsn=None, log_fn=print):
    # a coluna EQUIPE (código curto, ex.: AR162B) tem nome parecido com Equipe: no banco vira equipe_curta
    return _executar("tabela_completa", df, COLUNAS_TC, log_fn, dsn, pk=PK["tabela_completa"], polo_de=_polo_por_nome)
