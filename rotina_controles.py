# -*- coding: utf-8 -*-
"""
=============================================================================
  CONTROLES DE ROTINA — GDF
  Executa os SQLs da pasta rotina/ contra o Oracle e gera HTML de auditoria.

  Identifica erros de lancamento e de integridade contabil do GDF.

  Dependencias:  pip install oracledb pandas
  Uso:
      python rotina_controles.py --mes 8 --ano 2026
      python rotina_controles.py --mes 8 --ano 2026 --saida painel/rotina.html
      python rotina_controles.py --mes 8 --ano 2026 --controles 01,08,15
=============================================================================
"""
import argparse, re, sys, warnings
from datetime import datetime
from html import escape as _esc
from pathlib import Path

warnings.filterwarnings('ignore', message='.*SQLAlchemy.*')

import oracledb
import pandas as pd
import config_local as _cfg

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIGURACAO  (credenciais em config_local.py — nao commitado)
# ─────────────────────────────────────────────────────────────────────────────
DB_USER     = _cfg.DB_USER
DB_PASSWORD = _cfg.DB_PASSWORD
DB_HOST     = _cfg.DB_HOST
DB_PORT     = _cfg.DB_PORT
DB_SERVICE  = _cfg.DB_SERVICE

INSTANT_CLIENT_DIR = _cfg.INSTANT_CLIENT_DIR
OUTPUT_DIR         = Path(__file__).parent / "painel"
ROTINA_DIR         = Path(__file__).parent / "rotina"

MESES = {1:'Janeiro', 2:'Fevereiro', 3:'Marco', 4:'Abril', 5:'Maio',
         6:'Junho', 7:'Julho', 8:'Agosto', 9:'Setembro',
         10:'Outubro', 11:'Novembro', 12:'Dezembro'}

THRESHOLD    = 0.01   # centavos — diferenca minima para considerar erro
MAX_LINHAS   = 300    # limite de linhas exibidas por tabela no HTML

# ─────────────────────────────────────────────────────────────────────────────
#  METADATA DOS CONTROLES
#
#  check: 'colN_nz'  — coluna de indice N (0-based) != 0 → linha com erro
#         'colN_neg' — coluna de indice N < 0           → linha com erro
#         'all_rows' — toda linha retornada e erro
#
#  Indices das colunas de check (apos COGESTAO/COGESTAOCONTAB/COUGCONTAB/COUG
#  que ocupam os indices 0-3 nos SQLs agrupados):
#    C01: C_84 esta no indice 5 (apos C_10 no indice 4)
#    C02-C06, C08-C10, C13, C14, C16: diferenca no indice 4
#    C15: saldo no indice 2 (SELECT sem GROUP BY: corrconta, conta, saldo, ug)
#    C17: all_rows (filtro WHERE ja seleciona so os erros)
#    C18: saldo no indice 3 (SELECT: nomeconta, conta, gestao, saldo, ug)
# ─────────────────────────────────────────────────────────────────────────────
CONTROLES = [
    {
        'id': '01', 'check': 'col5_nz',
        'nome': 'C01 — Contas 21 Fluxo Principal × 827110101',
        'tipo': 'LANCAMENTO',
        'sql': '01-Controle 1 - Contas 21 F Principal do exercício x 827110101 LANÇAMENTO.sql',
        'descricao': (
            'Verifica se os lancamentos nas contas 21 do Fluxo Principal do exercicio '
            '(211xxx–215xxx) estao em equilibrio com a NE Principal (827110101). '
            'Diferenca != 0 por UG indica lancamento inconsistente.'
        ),
        'rename': {
            'E997027': 'Gestao', 'E997028': 'Gestao Contab',
            'E997036': 'UG Contab', 'E997038': 'UG',
            'C_84': 'DIFERENCA',
            'C_10': '2152301XX', 'C_11': '2112101XX', 'C_12': '2111101XX',
            'C_13': '2112104XX', 'C_14': '2112105XX', 'C_15': '2112106XX',
            'C_16': '2112107XX', 'C_17': '2112305XX', 'C_18': '2153101XX',
            'C_19': '2153102XX', 'C_1': '2142201XX', 'C_2': '2142202XX',
            'C_20': '2153103XX', 'C_21': '2153104XX', 'C_22': '2153105XX',
            'C_23': '2153106XX', 'C_24': '2153199XX', 'C_25': '2153301XX',
            'C_26': '2153303XX', 'C_27': '2153304XX', 'C_28': '2153305XX',
            'C_29': '2153306XX', 'C_3': '2142203XX', 'C_30': '2153399XX',
            'C_31': '2113101XX', 'C_32': '2113103XX', 'C_33': '2113104XX',
            'C_34': '2113105XX', 'C_35': '2114101XX', 'C_36': '2114105XX',
            'C_37': '2114106XX', 'C_38': '2114107XX', 'C_39': '2114108XX',
            'C_4': '2142206XX', 'C_40': '2114109XX', 'C_41': '2114112XX',
            'C_42': '2111104XX', 'C_43': '2114201XX', 'C_44': '2114202XX',
            'C_45': '2114203XX', 'C_46': '2114301XX', 'C_47': '2114303XX',
            'C_48': '2114305XX', 'C_49': '2114306XX', 'C_5': '2142299XX',
            'C_50': '2111105XX', 'C_51': '2114307XX', 'C_52': '2114308XX',
            'C_53': '218910101', 'C_54': '218910102', 'C_55': '218910104',
            'C_56': '218910199', 'C_57': '218910200', 'C_58': '218910500',
            'C_59': '218910700', 'C_6': '2143201XX', 'C_60': '218910800',
            'C_61': '218910900', 'C_62': '218911000', 'C_63': '218911100',
            'C_64': '218911200', 'C_65': '218911400', 'C_66': '218911800',
            'C_67': '218911900', 'C_68': '218912300', 'C_69': '218913000',
            'C_7': '2143202XX', 'C_70': '218913301', 'C_71': '218913401',
            'C_72': '218913402', 'C_73': '218913403', 'C_74': '218919600',
            'C_75': '218920102', 'C_76': '218920400', 'C_77': '218920500',
            'C_78': '218920700', 'C_79': '218929600', 'C_8': '2143502XX',
            'C_80': '218930500', 'C_81': '827110101', 'C_82': '2111106XX',
            'C_83': '2111107XX', 'C_85': '2121102XX', 'C_86': '2121302XX',
            'C_87': '2121304XX', 'C_88': '2122102XX', 'C_89': '2122103XX',
            'C_9': '2152101XX', 'C_90': '2122302XX', 'C_91': '2122303XX',
            'C_92': '2123101XX', 'C_93': '2123301XX', 'C_94': '218910300',
            'C_95': '2124101XX', 'C_96': '2124102XX', 'C_97': '2124301XX',
            'C_98': '2124302XX', 'C_99': '21251XXXX', 'C_100': '21253XXXX',
            'C_101': '21261XXXX', 'C_102': '21263XXXX', 'C_103': '2131101XX',
            'C_104': '2131103XX', 'C_105': '2131105XX', 'C_106': '2131106XX',
            'C_107': '2131107XX', 'C_108': '2131108XX', 'C_109': '2131109XX',
            'C_110': '2131110XX', 'C_111': '2131201XX', 'C_112': '2131203XX',
            'C_113': '2131205XX', 'C_114': '2132101XX', 'C_115': '2132102XX',
            'C_116': '2141301XX', 'C_117': '2141302XX', 'C_118': '2141305XX',
            'C_119': '2141309XX', 'C_120': '2141310XX', 'C_121': '2141311XX',
            'C_122': '2141312XX', 'C_123': '2141399XX',
        },
    },
    {
        'id': '02', 'check': 'col4_nz',
        'nome': 'C02 — Contas 21 RP × 827110201/827110203/631810000',
        'tipo': 'LANCAMENTO',
        'sql': '02-Controle 2 - Contas 21 F Ret do exercício x 827110201,827110203,63181 LANÇAMENTO.sql',
        'descricao': (
            'Equilibrio entre contas 21 de Restos a Pagar e as contas '
            '827110201, 827110203 e 631810000. Diferenca != 0 = RP sem contrapartida.'
        ),
    },
    {
        'id': '03', 'check': 'col4_nz',
        'nome': 'C03 — 21xxx98xx = 63211xxxx',
        'tipo': 'LANCAMENTO',
        'sql': '03-Controle 3 - 21xxx98xx = 63211xxxx LANÇAMENTO.sql',
        'descricao': (
            'Contas de cancelamento de RP (21xxx98xx) devem corresponder a '
            '632110100 e 632110300. Diferenca != 0 = cancelamento sem registro adequado.'
        ),
    },
    {
        'id': '04', 'check': 'col4_nz',
        'nome': 'C04 — Contas 21 sem NE × 827110401',
        'tipo': 'LANCAMENTO',
        'sql': '04-Controle 4 - Contas 21 sem NE x 827110401 LANÇAMENTO.sql',
        'descricao': (
            'Lancamentos em contas 21 sem nota de empenho frente a conta 827110401. '
            'Diferenca != 0 = RP sem empenho identificado.'
        ),
    },
    {
        'id': '05', 'check': 'col4_nz',
        'nome': 'C05 — Contas 21 em liquidacao × 827110196',
        'tipo': 'LANCAMENTO',
        'sql': '05-Controle 5 - Contas 21 em liquidação  x 827110196 LANÇAMENTO.sql',
        'descricao': (
            'RP em liquidacao: contas 218919600/218929600 e 218919896/218929896 '
            'frente a 827110196 e 631200000. Diferenca != 0 = RP em liquidacao inconsistente.'
        ),
    },
    {
        'id': '06', 'check': 'col4_nz',
        'nome': 'C06 — RPNP liquidado × 631300000',
        'tipo': 'LANCAMENTO',
        'sql': '06-Controle 6 - Contas 21 RPNP liquidado x 6313 LANÇAMENTO.sql',
        'descricao': (
            'RPNP liquidado: contas 218914002/218924002 frente a conta 631300000. '
            'Diferenca != 0 = liquidacao de RPNP sem registro correto.'
        ),
    },
    {
        'id': '08', 'check': 'col4_nz',
        'nome': 'C08 — Balancete INTRA',
        'tipo': 'INTEGRIDADE',
        'sql': '08-Controle 8 - Balancete INTRA.sql',
        'descricao': (
            'Equilibrio do balancete INTRA (5o digito = 2). '
            'Soma das classes 1+2+3+4 deve ser zero. '
            'Diferenca != 0 = lancamento sem contrapartida INTRA.'
        ),
    },
    {
        'id': '09', 'check': 'col4_nz',
        'nome': 'C09 — Balancete Nao INTRA',
        'tipo': 'INTEGRIDADE',
        'sql': '09-Controle 9 - Balancete Não INTRA.sql',
        'descricao': (
            'Equilibrio do balancete Nao INTRA. '
            'Soma das classes 1+2+3+4 deve ser zero por UG.'
        ),
    },
    {
        'id': '10', 'check': 'col4_nz',
        'nome': 'C10 — Balanco Financeiro (BF)',
        'tipo': 'INTEGRIDADE',
        'sql': '10-Balanço-BF.sql',
        'descricao': (
            'Equilibrio do Balanco Financeiro por UG (Ingressos = Dispendios + Variacao de Saldo). '
            'Diferenca != 0 indica BF desequilibrado.'
        ),
    },
    {
        'id': '13', 'check': 'col4_nz',
        'nome': 'C13 — Balanco Patrimonial (BP)',
        'tipo': 'INTEGRIDADE',
        'sql': '13-Balanço-BP.sql',
        'descricao': (
            'Equilibrio do BP por UG (Ativo = Passivo + PL). '
            'Diferenca != 0 indica inconsistencia patrimonial.'
        ),
    },
    {
        'id': '14', 'check': 'col4_nz',
        'nome': 'C14 — Contas 72119XXXX × 82119XXXX',
        'tipo': 'LANCAMENTO',
        'sql': '14-72119XXXX.sql',
        'descricao': (
            'Equilibrio entre contas 721190100-400 e 821190100-400 '
            '(excluindo documento 2025NS00005). Diferenca != 0 = contrapartida nao registrada.'
        ),
    },
    {
        'id': '15', 'check': 'col2_neg',
        'nome': 'C15 — Receita Negativa (VSALDOCONTABIL)',
        'tipo': 'INTEGRIDADE',
        'sql': '15-Receita Negativa Saldo.sql',
        'descricao': (
            'Contas de receita orcamentaria (621200000-621399999) com saldo negativo '
            'no VSALDOCONTABIL. Saldo negativo indica estorno maior que o valor lancado.'
        ),
    },
    {
        'id': '16', 'check': 'col4_nz',
        'nome': 'C16 — Contas 5221904XX',
        'tipo': 'LANCAMENTO',
        'sql': '16-5221904XX.sql',
        'descricao': (
            'Movimentacao nas contas 522190401 e 522190409. '
            'Valor != 0 indica lancamentos que precisam ser verificados.'
        ),
    },
    {
        'id': '17', 'check': 'col3_neg',
        'nome': 'C17 — Inversao de Saldo',
        'tipo': 'INTEGRIDADE',
        'sql': '17-Inversão de Saldo.sql',
        'descricao': (
            'Contas ativas (classe 1) com saldo natural Devedor (INSALDOCONTABIL=D, ININVERSAOSALDO=N) '
            'que apresentam saldo Credor (negativo) no VSALDOCONTABIL. '
            'Saldo negativo = conta com saldo invertido.'
        ),
    },
    {
        'id': '18', 'check': 'col3_nz',
        'nome': 'C18 — Previsao Adicional a Lancar',
        'tipo': 'INTEGRIDADE',
        'sql': '18-Previsão Adicional a Lançar.sql',
        'descricao': (
            'Saldo pendente nas contas de previsao adicional (521920500 e 821191201) ate o mes 6. '
            'Saldo != 0 indica previsao adicional ainda nao lancada.'
        ),
    },
]


# ─────────────────────────────────────────────────────────────────────────────
#  CONEXAO
# ─────────────────────────────────────────────────────────────────────────────
def conectar_oracle():
    print(f"  Inicializando Oracle Client: {INSTANT_CLIENT_DIR}")
    try:
        oracledb.init_oracle_client(lib_dir=INSTANT_CLIENT_DIR)
    except Exception as e:
        if "already been initialized" not in str(e):
            raise
    dsn = f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}"
    conn = oracledb.connect(user=DB_USER, password=DB_PASSWORD, dsn=dsn)
    print(f"  Conectado! Oracle {conn.version}")
    return conn


# ─────────────────────────────────────────────────────────────────────────────
#  EXECUCAO DE CADA CONTROLE
# ─────────────────────────────────────────────────────────────────────────────
_COLS_ID = {'E997027', 'E997028', 'E997036', 'E997038',
            'COGESTAO', 'COGESTAOCONTAB', 'COUGCONTAB', 'COUG',
            'ININVERSAOSALDO', 'NOCONTACONTABIL',
            'COCONTACONTABIL', 'COCONTACORRENTE',
            'GESTAO', 'GESTAO CONTAB', 'UG CONTAB', 'UG'}

# Colunas que devem ser zero-padded (UG = 6 digitos, GESTAO = 5 digitos, CONTA = 9 digitos)
_UG_COLS      = {'E997036', 'E997038', 'COUGCONTAB', 'COUG', 'UG CONTAB', 'UG'}
_GESTAO_COLS  = {'E997027', 'E997028', 'COGESTAO', 'COGESTAOCONTAB', 'GESTAO', 'GESTAO CONTAB'}
_CONTA_COLS   = {'COCONTACONTABIL', 'COCONTACORRENTE', 'NOCONTACONTABIL'}


def _is_id_col(col):
    """True se a coluna deve ser tratada como identificador (sem formato numerico)."""
    col_up = col.upper()
    if col_up in _COLS_ID:
        return True
    # Captura expressoes SUBSTR/TRIM sem alias que contenham CONTA ou CORRENTE
    return 'CONTA' in col_up or 'CORRENTE' in col_up

# Mapeamento de aliases de grouping Oracle → nome amigavel
_ALIAS_AMIGAVEL = {
    'E997027': 'Gestao',       'E997028': 'Gestao Contab',
    'E997036': 'UG Contab',    'E997038': 'UG',
    'COGESTAO': 'Gestao',      'COGESTAOCONTAB': 'Gestao Contab',
    'COUGCONTAB': 'UG Contab', 'COUG': 'UG',
}


def _auto_rename_sql(sql_text):
    """Gera dict de rename a partir dos aliases calculados (C_N) no SQL."""
    rename = dict(_ALIAS_AMIGAVEL)
    segs = re.findall(r'((?:SUM|DECODE)\(.+?)\)\s*as\s+(C_\w+|E\w+)', sql_text, re.DOTALL)
    for expr, alias in segs:
        alias_up = alias.upper()
        if alias_up in rename:
            continue
        m = re.search(r'BETWEEN\s+(\d+)\s+AND\s+(\d+)', expr)
        if m:
            lo, hi = m.group(1), m.group(2)
            if lo == hi:
                rename[alias_up] = lo
            else:
                common = ''
                for a, b in zip(lo, hi):
                    if a == b:
                        common += a
                    else:
                        break
                rename[alias_up] = common + 'X' * (len(lo) - len(common))
            continue
        m = re.search(r'DECODE\([^,]+,\s*(\d+),', expr)
        if m:
            rename[alias_up] = m.group(1)
    return rename


def executar_controle(conn, controle, ano):
    sql_path = ROTINA_DIR / controle['sql']
    if not sql_path.exists():
        raise FileNotFoundError(f"SQL nao encontrado: {sql_path}")

    sql = sql_path.read_text(encoding='utf-8')
    sql = sql.replace('MIL2026', f'MIL{ano}').strip().rstrip(';')

    print(f"    [{controle['id']}] executando...", end=' ', flush=True)
    df = pd.read_sql(sql, conn)
    df.columns = [str(c).upper() for c in df.columns]

    # Renomeia colunas: usa dict manual se existir, senao parseia o SQL
    if 'rename' in controle:
        rename_map = {k.upper(): v for k, v in controle['rename'].items()}
    else:
        rename_map = _auto_rename_sql(sql)

    # Marca a coluna de check como DIFERENCA (exceto all_rows e col3_neg/nz que sao saldo/inversao)
    chk = controle['check']
    if chk not in ('all_rows',) and '_' in chk:
        chk_idx = int(chk[3:chk.index('_')])
        chk_col = df.columns[chk_idx]
        if chk_col not in rename_map or rename_map.get(chk_col) == chk_col:
            rename_map[chk_col] = 'DIFERENCA'

    # Desambigua alvos duplicados (ex: dois C_N mapeados ao mesmo codigo de conta)
    seen: dict = {}
    dedup: dict = {}
    for orig, novo in rename_map.items():
        if novo not in seen:
            seen[novo] = 1
            dedup[orig] = novo
        else:
            seen[novo] += 1
            dedup[orig] = f'{novo}_{seen[novo]}'
    df.rename(columns=dedup, inplace=True)

    for col in df.columns:
        if not _is_id_col(col):
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    # Determina linhas com problema
    chk = controle['check']
    if chk == 'all_rows':
        mask = pd.Series(True, index=df.index)
    elif '_nz' in chk:
        idx = int(chk[3:chk.index('_')])
        mask = df.iloc[:, idx].abs() > THRESHOLD
    elif '_neg' in chk:
        idx = int(chk[3:chk.index('_')])
        mask = df.iloc[:, idx] < -THRESHOLD
    else:
        mask = pd.Series(False, index=df.index)

    df_erro = df[mask].copy()
    n_erro  = len(df_erro)
    n_total = len(df)
    print(f"{n_total} linhas, {n_erro} com erro.")
    return df, df_erro, n_erro, n_total


# ─────────────────────────────────────────────────────────────────────────────
#  FORMATACAO
# ─────────────────────────────────────────────────────────────────────────────
def _brl(v):
    try:
        v = float(v)
    except (ValueError, TypeError):
        return str(v)
    neg = v < 0
    s = f"{abs(v):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    return f"({s})" if neg else s


def _fmt(v, col):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return '—'
    col_up = col.upper()
    if not _is_id_col(col):
        try:
            return _brl(float(v))
        except (ValueError, TypeError):
            pass
        return _esc(str(v))
    # Colunas de identificacao: sem separador de milhar, inteiro se possivel
    try:
        iv = int(float(v))
        if col_up in _UG_COLS:
            return _esc(str(iv).zfill(6))
        if col_up in _GESTAO_COLS:
            return _esc(str(iv).zfill(5))
        if col_up in _CONTA_COLS or 'CONTA' in col_up or 'CORRENTE' in col_up:
            return _esc(str(iv).zfill(9))
        return _esc(str(iv))
    except (ValueError, TypeError):
        pass
    return _esc(str(v))


# ─────────────────────────────────────────────────────────────────────────────
#  CSS (mesmo padrao visual de auditoria_consolidada.html)
# ─────────────────────────────────────────────────────────────────────────────
_CSS = """
:root{
  --brand:#1A4A8F;--pg:#F1F4FB;--s1:#FFF;--s2:#E8EDF7;--bd:#C8D5ED;
  --t1:#0D1829;--t2:#4C5C7A;--t3:#8A9BBD;
  --err:#BE1C1C;--err-bg:#FEF2F2;--err-bd:#FECACA;
  --ok:#156030;--ok-bg:#F0FDF4;--ok-bd:#BBF7D0;
  --inf:#1547A0;--inf-bg:#EFF6FF;--inf-bd:#BFDBFE;
  --fn:'Segoe UI',system-ui,-apple-system,BlinkMacSystemFont,sans-serif;
  --fm:'Cascadia Code','SF Mono',Consolas,'Courier New',monospace;
  --r:8px;
}
@media(prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --brand:#4889D8;--pg:#080C18;--s1:#101827;--s2:#182035;--bd:#1D2C48;
  --t1:#D4DFF5;--t2:#6A7EA6;--t3:#3D5070;
  --err:#F87171;--err-bg:#170404;--err-bd:#7F1D1D;
  --ok:#4ADE80;--ok-bg:#031309;--ok-bd:#14532D;
  --inf:#93C5FD;--inf-bg:#0C1A35;--inf-bd:#1E3A5F;
}}
:root[data-theme="dark"]{
  --brand:#4889D8;--pg:#080C18;--s1:#101827;--s2:#182035;--bd:#1D2C48;
  --t1:#D4DFF5;--t2:#6A7EA6;--t3:#3D5070;
  --err:#F87171;--err-bg:#170404;--err-bd:#7F1D1D;
  --ok:#4ADE80;--ok-bg:#031309;--ok-bd:#14532D;
  --inf:#93C5FD;--inf-bg:#0C1A35;--inf-bd:#1E3A5F;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--fn);font-size:14px;line-height:1.5;color:var(--t1);background:var(--pg)}
.hd{background:var(--brand);color:#fff;padding:14px 24px;display:flex;align-items:center;
    justify-content:space-between;gap:16px;position:sticky;top:0;z-index:10}
.hd-title{font-size:15px;font-weight:700;letter-spacing:-.01em}
.hd-sub{font-size:12px;opacity:.72;margin-top:2px}
.wrap{max-width:1100px;margin:0 auto;padding:24px 20px;display:flex;flex-direction:column;gap:14px}
.nav{display:flex;flex-wrap:wrap;gap:6px;padding:10px 14px;background:var(--s1);
     border:1px solid var(--bd);border-radius:var(--r)}
.nav a{font-size:11px;font-weight:600;color:var(--brand);text-decoration:none;
       padding:3px 9px;border-radius:100px;border:1px solid var(--bd)}
.nav a:hover{border-color:var(--brand)}
.kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.kpi{background:var(--s1);border:1px solid var(--bd);border-radius:var(--r);
     padding:13px 15px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;
             background:var(--kpi-stripe,var(--bd))}
.kpi-lbl{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--t3);margin-bottom:5px}
.kpi-val{font-size:22px;font-weight:700;letter-spacing:-.025em;color:var(--kpi-color,var(--t1))}
details.ctrl{background:var(--s1);border:1px solid var(--bd);border-radius:var(--r);overflow:hidden}
details.ctrl summary{padding:11px 18px;font-size:13.5px;font-weight:600;cursor:pointer;
                     list-style:none;display:flex;align-items:center;gap:10px;background:var(--s2)}
details.ctrl summary::-webkit-details-marker{display:none}
details.ctrl summary::after{content:'+';font-size:15px;color:var(--t3);margin-left:auto}
details.ctrl[open] summary::after{content:'\2212'}
.ctrl-body{padding:14px 18px;display:flex;flex-direction:column;gap:10px}
.ctrl-desc{font-size:12.5px;color:var(--t2);line-height:1.55}
.chip{font-size:10px;font-weight:700;padding:1px 8px;border-radius:100px;border:1px solid;white-space:nowrap}
.c-ok{background:var(--ok-bg);color:var(--ok);border-color:var(--ok-bd)}
.c-err{background:var(--err-bg);color:var(--err);border-color:var(--err-bd)}
.c-inf{background:var(--inf-bg);color:var(--inf);border-color:var(--inf-bd)}
.tbl-wrap{overflow-x:auto;border:1px solid var(--bd);border-radius:6px;max-height:420px;overflow-y:auto}
table.tbl{width:100%;border-collapse:collapse;font-size:11.5px;font-family:var(--fm)}
table.tbl th,table.tbl td{padding:4px 9px;border-bottom:1px solid var(--bd);white-space:nowrap}
table.tbl th{background:var(--s2);font-weight:700;font-family:var(--fn);position:sticky;top:0;z-index:1}
table.tbl td.num{text-align:right}
table.tbl tr.err td{background:var(--err-bg)}
table.tbl tfoot td{background:var(--s2);font-weight:700;font-family:var(--fn);border-top:2px solid var(--brand);position:sticky;bottom:0}
.ok-msg{color:var(--ok);font-size:13px;font-weight:600}
.stats{font-size:12px;color:var(--t2)}
.trunc{font-size:11px;color:var(--t3);font-style:italic;margin-top:4px}
.falha{color:var(--err);font-size:13px}
.tbl-toolbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.btn-xls{display:inline-flex;align-items:center;gap:4px;font-size:11px;font-weight:600;padding:3px 11px;border-radius:5px;border:1px solid var(--bd);background:var(--s2);color:var(--t2);cursor:pointer;font-family:var(--fn);transition:color .15s,border-color .15s}
.btn-xls:hover{color:var(--t1);border-color:var(--t2)}
footer{font-size:11px;color:var(--t3);text-align:center;padding:14px}
@media(max-width:680px){.kpi-row{grid-template-columns:repeat(2,1fr)}}
"""


# ─────────────────────────────────────────────────────────────────────────────
#  GERACAO DO HTML
# ─────────────────────────────────────────────────────────────────────────────
def _tabela_html(df, df_erro, cid=''):
    erro_idx = set(df_erro.index)
    cols = list(df.columns)

    ths = ''.join(f'<th>{_esc(c)}</th>' for c in cols)

    exibir = df if len(df) <= MAX_LINHAS else df.iloc[:MAX_LINHAS]
    truncado = len(df) > MAX_LINHAS

    linhas = []
    for idx, row in exibir.iterrows():
        cls = ' class="err"' if idx in erro_idx else ''
        cells = []
        for c in cols:
            v = row[c]
            if not _is_id_col(c):
                try:
                    cells.append(f'<td class="num">{_fmt(v, c)}</td>')
                    continue
                except Exception:
                    pass
            cells.append(f'<td>{_fmt(v, c)}</td>')
        linhas.append(f'<tr{cls}>{"".join(cells)}</tr>')

    # Linha de totais (apenas colunas numéricas com soma != 0)
    tfoot = ''
    totais = {}
    for c in cols:
        if not _is_id_col(c):
            try:
                totais[c] = float(pd.to_numeric(df[c], errors='coerce').fillna(0).sum())
            except Exception:
                totais[c] = 0.0
    if totais and any(abs(v) > 0.005 for v in totais.values()):
        fcs = []
        primeiro_texto = True
        for c in cols:
            if _is_id_col(c):
                fcs.append(f'<td>{"<strong>TOTAL</strong>" if primeiro_texto else ""}</td>')
                primeiro_texto = False
            else:
                v = totais.get(c, 0.0)
                fcs.append(f'<td class="num"><strong>{_brl(v)}</strong></td>')
                primeiro_texto = False
        tfoot = f'<tfoot><tr>{"".join(fcs)}</tr></tfoot>'

    tbl_id = f' id="tbl-{cid}"' if cid else ''
    tbl = (
        f'<div class="tbl-wrap"><table class="tbl"{tbl_id}>'
        f'<thead><tr>{ths}</tr></thead>'
        f'<tbody>{"".join(linhas)}</tbody>'
        f'{tfoot}'
        f'</table></div>'
    )
    nota = (f'<p class="trunc">Exibindo {MAX_LINHAS} de {len(df)} linhas.</p>'
            if truncado else '')
    return tbl + nota


def gerar_html(resultados, mes, ano, saida):
    """resultados: list of (controle, df|None, df_erro|None, n_erro, n_total)"""
    n_err    = sum(1 for _, df, _, ne, _ in resultados if df is not None and ne > 0)
    n_ok     = sum(1 for _, df, _, ne, _ in resultados if df is not None and ne == 0)
    n_falhou = sum(1 for _, df, _, _, _ in resultados if df is None)
    tot_err  = sum((ne or 0) for _, _, _, ne, _ in resultados)

    mes_label = f'{mes:02d}/{ano} — {MESES[mes]}/{ano}'
    agora     = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    # KPIs
    kpi_cor   = 'var(--err)' if n_err else 'var(--ok)'
    kpis = f"""
<div class="kpi-row">
  <div class="kpi" style="--kpi-stripe:{kpi_cor};--kpi-color:{kpi_cor}">
    <div class="kpi-lbl">Controles com Erro</div>
    <div class="kpi-val">{n_err}</div>
  </div>
  <div class="kpi" style="--kpi-stripe:var(--ok);--kpi-color:var(--ok)">
    <div class="kpi-lbl">Controles OK</div>
    <div class="kpi-val">{n_ok}</div>
  </div>
  <div class="kpi">
    <div class="kpi-lbl">Total de Controles</div>
    <div class="kpi-val">{len(resultados)}</div>
  </div>
  <div class="kpi" style="--kpi-stripe:var(--err);--kpi-color:{'var(--err)' if tot_err else 'var(--t1)'}">
    <div class="kpi-lbl">Linhas com Erro</div>
    <div class="kpi-val">{tot_err:,}</div>
  </div>
</div>"""

    # Nav
    nav = '<div class="nav">' + ''.join(
        f'<a href="#ctrl-{c["id"]}">{c["id"]}</a>'
        for c, _, _, _, _ in resultados
    ) + '</div>'

    # Cards
    cards = []
    for controle, df, df_erro, n_erro, n_total in resultados:
        cid = controle['id']

        btn_xls = (f'<button class="btn-xls" '
                   f'onclick="exportarExcel(\'tbl-{cid}\',\'C{cid}\')" '
                   f'title="Exportar tabela para Excel">&#8595; Excel</button>')

        if df is None:
            status = '<span class="chip c-err">FALHA</span>'
            corpo  = '<p class="falha">Falha ao executar o SQL deste controle.</p>'
            aberto = ' open'
        elif n_erro > 0:
            status = '<span class="chip c-err">ERRO</span>'
            corpo  = (f'<div class="tbl-toolbar"><p class="stats">{n_erro} de {n_total} linhas com divergencia</p>'
                      f'{btn_xls}</div>'
                      + _tabela_html(df, df_erro, cid))
            aberto = ' open'
        else:
            status = '<span class="chip c-ok">OK</span>'
            corpo  = (f'<div class="tbl-toolbar">'
                      f'<p class="ok-msg">&#10003; Nenhuma divergencia encontrada ({n_total} linhas verificadas).</p>'
                      f'{btn_xls}</div>'
                      + _tabela_html(df, df_erro, cid))
            aberto = ''

        tipo = f'<span class="chip c-inf">{_esc(controle["tipo"])}</span>'

        cards.append(f"""
<details class="ctrl"{aberto} id="ctrl-{cid}">
  <summary>{status} {tipo} {_esc(controle["nome"])}</summary>
  <div class="ctrl-body">
    <p class="ctrl-desc">{_esc(controle["descricao"])}</p>
    {corpo}
  </div>
</details>""")

    resultado_geral = (
        f'<span style="color:#fff;font-weight:700;opacity:.9">'
        f'{n_err} controle(s) com ERRO</span>'
        if n_err else
        '<span style="color:#a7f3d0;font-weight:700">Todos os controles OK</span>'
    )

    page = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Controles de Rotina GDF {mes_label}</title>
<style>{_CSS}</style>
</head>
<body>
<div class="hd">
  <div>
    <div class="hd-title">Controles de Rotina &#8212; GDF</div>
    <div class="hd-sub">{_esc(mes_label)} &middot; Emitido em {agora}</div>
  </div>
  <div>{resultado_geral}</div>
</div>
<div class="wrap">
  {nav}
  {kpis}
  {"".join(cards)}
</div>
<footer>Gerado por rotina_controles.py &middot; {agora}</footer>
<script>
function exportarExcel(tblId, nome) {{
  var tbl = document.getElementById(tblId);
  if (!tbl) return;
  var esc = function(s) {{ return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }};
  var xml = '<?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?>';
  xml += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ';
  xml += 'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';
  xml += '<Worksheet ss:Name="Dados"><Table>';
  var rows = tbl.rows;
  for (var i = 0; i < rows.length; i++) {{
    xml += '<Row>';
    var cells = rows[i].cells;
    for (var j = 0; j < cells.length; j++) {{
      var t = cells[j].innerText.trim();
      if (cells[j].classList.contains('num') && t !== '—' && t !== '') {{
        var n = parseFloat(t.replace(/\./g,'').replace(',','.'));
        if (!isNaN(n)) {{ xml += '<Cell><Data ss:Type="Number">' + n + '</Data></Cell>'; continue; }}
      }}
      xml += '<Cell><Data ss:Type="String">' + esc(t) + '</Data></Cell>';
    }}
    xml += '</Row>';
  }}
  xml += '</Table></Worksheet></Workbook>';
  var blob = new Blob(['﻿' + xml], {{type:'application/vnd.ms-excel;charset=utf-8'}});
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = nome + '.xls';
  document.body.appendChild(a); a.click();
  setTimeout(function(){{ URL.revokeObjectURL(url); a.remove(); }}, 800);
}}
</script>
</body>
</html>"""

    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(page, encoding='utf-8')
    print(f"  HTML salvo: {saida}")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(
        description='Controles de Rotina GDF — gera HTML de auditoria a partir dos SQLs em rotina/')
    p.add_argument('--mes',       type=int, default=datetime.now().month,
                   help='Mes de referencia 1-12')
    p.add_argument('--ano',       type=int, default=datetime.now().year)
    p.add_argument('--saida',     type=str, default=None,
                   help='Caminho do HTML de saida '
                        '(default: painel/rotina_controles_AAAA_MM.html)')
    p.add_argument('--controles', type=str, default='',
                   help='Executar apenas estes IDs (ex: 01,08,15). Vazio = todos.')
    a = p.parse_args()

    if not 1 <= a.mes <= 12:
        print("ERRO: --mes deve estar entre 1 e 12."); sys.exit(1)

    filtro = {x.strip() for x in a.controles.split(',') if x.strip()}
    lista  = [c for c in CONTROLES if not filtro or c['id'] in filtro]

    saida = (Path(a.saida) if a.saida
             else OUTPUT_DIR / f'rotina_controles_{a.ano}_{a.mes:02d}.html')

    print(f"\n{'='*60}")
    print(f"  CONTROLES DE ROTINA — GDF")
    print(f"  {MESES[a.mes]}/{a.ano}  |  {len(lista)} controles")
    print(f"{'='*60}")

    print("\n[1/3] Conectando ao Oracle...")
    conn = conectar_oracle()

    print(f"\n[2/3] Executando {len(lista)} controles...")
    resultados = []
    for controle in lista:
        try:
            df, df_erro, n_erro, n_total = executar_controle(conn, controle, a.ano)
            resultados.append((controle, df, df_erro, n_erro, n_total))
        except Exception as e:
            print(f"    FALHA: {e}")
            resultados.append((controle, None, None, 0, 0))

    conn.close()
    print("  Conexao Oracle encerrada.")

    print("\n[3/3] Gerando HTML...")
    gerar_html(resultados, a.mes, a.ano, saida)

    n_err = sum(1 for _, df, _, ne, _ in resultados if df is not None and ne > 0)
    print(f"\n{'='*60}")
    print(f"  {'CONCLUIDO COM ' + str(n_err) + ' CONTROLE(S) COM ERRO' if n_err else 'CONCLUIDO — TODOS OK'}")
    print(f"{'='*60}\n")


if __name__ == '__main__':
    main()
