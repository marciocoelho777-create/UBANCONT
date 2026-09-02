# -*- coding: utf-8 -*-
"""
=============================================================================
  BALANCO ORCAMENTARIO - GDF  (PSIAG550 / Anexo 12)
  SQL fixo, equacoes embutidas conforme "Lista Equacoes de Balanco -
  Balanco Orcamentario" (Tipo 04 - arquivo
  Relatorio_Equacao_de_Balanco_Orcamentario.xlsx), seguindo o MESMO padrao
  de extracao Oracle ja validado e em producao nos projetos:
    - dfc_v5_20_06_2026.py                (Demonstracao dos Fluxos de Caixa, Anexo 15)
    - balanco_financeiro_23_06_2026_ok.py (Balanco Financeiro, Anexo 13)

  Layout de saida (linhas/colunas) espelha ListaBalancoOrcamentario.pdf
  (Anexo 12, Governo do Distrito Federal, Maio/2026, "1 - Com Superavit").

  COMPOSICAO (MCASP 9a Edicao, item 2.1):
    a. Quadro Principal (Receitas x Despesas Orcamentarias)
    b. Quadro por Tipos de Creditos Adicionais (Despesas)
    c. Quadro da Execucao de Restos a Pagar Nao Processados
    d. Quadro da Execucao de Restos a Pagar Processados

  LOGICA DE COLUNA  (campo "Coluna" da planilha de equacoes):
    RECEITAS (Quadro Principal):
      col 1 = Previsao Inicial    (a)   col 2 = Previsao Atualizada (b)
      col 3 = Receitas Realizadas (c)   Saldo (d) = (c - b)            [calculado]
    DESPESAS (Quadro Principal):
      col 1 = Dotacao Inicial  (e)   col 2 = Dotacao Atualizada (f)
      col 3 = Desp. Empenhadas (g)   col 4 = Desp. Liquidadas   (h)
      col 5 = Despesas Pagas   (i)   Saldo da Dotacao (j) = (f - g)    [calculado]
    CREDITOS ADICIONAIS:
      col1=Cred.Adic.Suplementar(a)  col2=Cred.Especiais Abertos(b)
      col3=Cred.Especiais Reabertos(c) col4=Cred.Extraord.Reabertos(d)
      col5=(-)Cancel.p/Cred.Suplem.(e) col6=(-)Remanej.Veto LOA(f)
      col7=(-)Cancel.p/Cred.Especial(g)
      Total Alteracoes (h) = (a+b+c+d-e-f-g)                          [calculado]
    RESTOS A PAGAR NAO PROCESSADOS:
      col1=Inscritos Exerc.Ant.(a) col2=Em 31/Dez Exerc.Ant.(b)
      col3=Liquidados(c) col4=Pagos(d) col5=Cancelados(e)
      Saldo (f) = (a+b-d-e)                                           [calculado]
    RESTOS A PAGAR PROCESSADOS:
      col1=Inscritos Exerc.Ant.(a) col2=Em 31/Dez Exerc.Ant.(b)
      col3=Pagos(c) col4=Cancelados(d)
      Saldo (e) = (a+b-c-d)                                           [calculado]

  LOGICA DE MES (INMES) e TIPOS DE MOVIMENTO -- IDENTICA AO DFC/BF:
    0  = saldo de abertura
    1  = movimento acumulado jan..mes
    13 = movimento acumulado + encerramento (padrao receita/despesa)
    SC = Saldo Credor  = C - D     SD = Saldo Devedor = D - C

  RECEITAS -> contas 5211XXXXX/5212XXXXX (previsao) e 6212XXXXX/6213XXXXX
              (realizada), filtradas por SUBSTR(o.COCONTACORRENTE,1,N) =
              prefixo de Natureza da Receita (campo NatReceita da planilha).
  DESPESAS -> contas 5221X/5222X (dotacao) e 6221X/6229X (execucao),
              filtradas por GND = SUBSTR(ne.CONATUREZA,2,1) (2o digito da
              natureza c.G.mm.ee.dd) via JOIN com NOTAEMPENHO, no MESMO
              padrao multi-schema (2023/2024/2025/{ano}) usado no DFC.

  Dependencias:  pip install oracledb openpyxl pandas reportlab
  Uso:
      python balanco_orcamentario.py --mes 5 --ano 2026
      python balanco_orcamentario.py --mes 5 --ano 2026 --ug 130101
      python balanco_orcamentario.py --mes 5 --ano 2026 --formato pdf
      python balanco_orcamentario.py --mes 5 --ano 2026 --diag
=============================================================================
"""
import argparse, calendar, sys
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import oracledb
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, PageBreak)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIGURACAO  (idem DFC / Balanco Financeiro)
# ─────────────────────────────────────────────────────────────────────────────
# Credenciais carregadas de config_local.py (nunca em texto puro aqui)
DB_USER = DB_PASSWORD = DB_HOST = DB_SERVICE = ""
DB_PORT = 1521
INSTANT_CLIENT_DIR = ""
try:
    import importlib.util as _ilu, pathlib as _pl
    _cfg = _pl.Path(__file__).parent / "config_local.py"
    if _cfg.exists():
        _spec = _ilu.spec_from_file_location("config_local", _cfg)
        _mod  = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_mod)
        for _k in ("DB_USER","DB_PASSWORD","DB_HOST","DB_PORT",
                   "DB_SERVICE","INSTANT_CLIENT_DIR"):
            if hasattr(_mod, _k):
                globals()[_k] = getattr(_mod, _k)
except Exception:
    pass
OUTPUT_DIR         = Path(r"C:\balanço 2026 gemini arquivos")

MESES = {1:'Janeiro',2:'Fevereiro',3:'Marco',4:'Abril',5:'Maio',
         6:'Junho',7:'Julho',8:'Agosto',9:'Setembro',
         10:'Outubro',11:'Novembro',12:'Dezembro'}

# ─────────────────────────────────────────────────────────────────────────────
#  SUBQUERY NOTAEMPENHO (multi-schema)
#  Fornece CONATUREZA para derivar o GND (2o digito) que filtra os itens
#  do Quadro de Restos a Pagar (NUNE no formato AAAANEnnnnn, onde AAAA pode
#  ser de exercicios bem anteriores ao corrente -- RP residual visto em
#  diagnostico chegou a 2018). ANOS_NOTAEMPENHO e descoberto dinamicamente
#  em descobrir_anos_notaempenho() (chamado uma vez no inicio de main()/
#  diagnostico()) para nao falhar com ORA-00942 caso algum ano nao tenha
#  schema criado no banco. Fallback fixo 2023..ano se nunca foi descoberto.
# ─────────────────────────────────────────────────────────────────────────────
ANOS_NOTAEMPENHO = None  # preenchido por descobrir_anos_notaempenho()

#  Janela de anos para tras que o UNION ALL do ne_union() cobre, a partir do
#  ano corrente. Verificado empiricamente em 21/08/2026 (mes 8/2026): dos 24
#  anos disponiveis (2003..2026), somente 2020..2026 (ano-6..ano) tinham
#  algum lancamento de RP/paga casado com NOTAEMPENHO -- 2003..2019 davam
#  zero matches e zero valor. ano-8 mantem 2 anos de folga sobre o que foi
#  observado, sem precisar manter todo o historico desde 2003. Se algum mes
#  futuro passar a exigir empenhos mais antigos que isso, aumentar este
#  numero (nao ha garantia formal de prescricao de RP embutida aqui).
JANELA_ANOS_NOTAEMPENHO = 8


def descobrir_anos_notaempenho(conn, ano):
    """Descobre os schemas MIL{aaaa} que existem e tem NOTAEMPENHO, com
    ano-JANELA_ANOS_NOTAEMPENHO <= aaaa <= ano. Preenche a global
    ANOS_NOTAEMPENHO. Chamar uma vez no inicio da execucao (main/diagnostico),
    antes de qualquer busca."""
    global ANOS_NOTAEMPENHO
    cur = conn.cursor()
    cur.execute("""SELECT owner FROM all_tables
                   WHERE table_name = 'NOTAEMPENHO' AND owner LIKE 'MIL%'""")
    owners = [r[0] for r in cur.fetchall()]
    cur.close()
    ano_min = ano - JANELA_ANOS_NOTAEMPENHO
    anos_validos = []
    for o in owners:
        sufixo = o.replace('MIL', '')
        if sufixo.isdigit() and ano_min <= int(sufixo) <= ano:
            anos_validos.append(int(sufixo))
    anos_validos.sort()
    if not anos_validos:
        anos_validos = list(range(max(2023, ano_min), ano + 1))
    ANOS_NOTAEMPENHO = anos_validos
    print(f"  [NOTAEMPENHO] Schemas detectados ({ano_min}..{ano}): "
          f"{anos_validos[0]}..{anos_validos[-1]} ({len(anos_validos)} anos)")
    return anos_validos


def ne_union(ano):
    anos = ANOS_NOTAEMPENHO if ANOS_NOTAEMPENHO else list(range(2023, ano + 1))
    partes = [f"SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{a}.NOTAEMPENHO" for a in anos]
    return "(\n        " + "\n        UNION ALL\n        ".join(partes) + "\n    )"

# ─────────────────────────────────────────────────────────────────────────────
#  ESTRUTURA DE RECEITAS — QUADRO PRINCIPAL
#  (codigo_item, nome, nivel_indent, tipo, prefixos_natrec)
#  tipo: 'titulo' | 'subtotal' | 'item' | 'total_calc'
#  prefixos_natrec: lista de prefixos (SUBSTR COCONTACORRENTE) que somam
#  para este item; None quando o item e' um subtotal/total calculado.
# ─────────────────────────────────────────────────────────────────────────────
RECEITAS = [
    ("RECEITAS CORRENTES (I)",                        0, "subtotal_calc", None),
    ("Impostos, Taxas e Contribuições de Melhoria",   1, "subtotal_calc", None),
    ("Impostos",                                      2, "item", ['111','711']),
    ("Taxas",                                         2, "item", ['112','712']),
    ("Contribuições de Melhoria",                     2, "item", ['113','713']),
    ("Receita de Contribuições",                      1, "subtotal_calc", None),
    ("Contribuições Sociais",                         2, "item", ['121','721']),
    ("Contribuições Econômicas",                      2, "item", ['122','722']),
    ("Contribuições p/ Entidades Priv. Interesse Púb.",2,"item", ['123','723']),
    ("Contrib. p/ Custeio Serv. Iluminação Pública",  2, "item", ['124','724']),
    ("Receita Patrimonial",                           1, "subtotal_calc", None),
    ("Exploração do Patrimônio Imobiliário do Estado",2, "item", ['131','731']),
    ("Valores Mobiliários",                           2, "item", ['132','732']),
    ("Delegação de Serv. Públ. Mediante Concessão",   2, "item", ['133','733']),
    ("Exploração de Recursos Naturais",               2, "item", ['134','734']),
    ("Exploração do Patrimônio Intangível",           2, "item", ['135','735']),
    ("Cessão de Direitos",                            2, "item", ['136','736']),
    ("Demais Receitas Patrimoniais",                  2, "item", ['139','739']),
    ("Receita Agropecuária",                          1, "item", ['140','141','740','741']),
    ("Receita Industrial",                             1, "item", ['150','151','750','751']),
    ("Receita de Serviços",                           1, "subtotal_calc", None),
    ("Serviços Administrativos e Comerciais Gerais",  2, "item", ['161','761']),
    ("Serviços e Atividades Ref. à Navegação e Transp.",2,"item", ['162','762']),
    ("Serviços e Atividades Referentes à Saúde",      2, "item", ['163','763']),
    ("Serviços e Atividades Financeiras",             2, "item", ['164','764']),
    ("Outros Serviços",                               2, "item", ['169','769']),
    ("Transferências Correntes",                      1, "subtotal_calc", None),
    ("Transferências da União e de suas Entidades",   2, "item", ['171','771']),
    ("Transferências dos Estados, DF e de suas Entidades",2,"item",['172','772']),
    ("Transferências dos Municípios e de suas Entidades",2,"item",['173','773']),
    ("Transferências de Instituições Privadas",       2, "item", ['174','774']),
    ("Transferências de Outras Instituições Públicas",2, "item", ['175','775']),
    ("Transferências do Exterior",                    2, "item", ['176','776']),
    ("Transferências de Pessoas Físicas",             2, "item", ['177','777']),
    ("Transferências Provenientes de Convênios",      2, "item", ['178','778']),
    ("Demais Transferências Correntes",               2, "item", ['179','779']),
    ("Outras Receitas Correntes",                     1, "subtotal_calc", None),
    ("Multas Administrativas, Contratuais e Judiciais",2,"item", ['191','791']),
    ("Indenizações, Restituições e Ressarcimentos",   2, "item", ['192','792']),
    ("Bens, Direitos e Valores Incorp. ao Patrimônio Públ.",2,"item",['193','793']),
    ("Multas e Juros de Mora das Receitas Tributárias",2,"item", ['194']),
    ("Demais Receitas Correntes",                     2, "item", ['199','799']),

    ("RECEITAS DE CAPITAL (II)",                      0, "subtotal_calc", None),
    ("Operações de Crédito",                          1, "subtotal_calc", None),
    ("Operações de Crédito - Mercado Interno",        2, "item", ['211','811']),
    ("Operações de Crédito - Mercado Externo",        2, "item", ['212','812']),
    ("Alienação de Bens",                             1, "subtotal_calc", None),
    ("Alienação de Bens Móveis",                      2, "item", ['221','821']),
    ("Alienação de Bens Imóveis",                     2, "item", ['222','822']),
    ("Alienação de Bens Intangíveis",                 2, "item", ['223','823']),
    ("Amortizações de Empréstimos",                   1, "item", ['231']),
    ("Transferências de Capital",                     1, "subtotal_calc", None),
    ("Transferências da União e de suas Entidades (Capital)",   2, "item", ['241','841']),
    ("Transferências dos Estados, DF e de suas Entidades (Capital)",2,"item",['242','842']),
    ("Transferências dos Municípios e de suas Entidades (Capital)",2,"item",['243','843']),
    ("Transferências de Instituições Privadas (Capital)",       2, "item", ['244','844']),
    ("Transferências de Outras Instituições Públicas (Capital)",2, "item", ['245','845']),
    ("Transferências do Exterior (Capital)",                    2, "item", ['246','846']),
    ("Demais Transferências de Capital",              2, "item", ['249','849']),
    ("Outras Receitas de Capital",                    1, "subtotal_calc", None),
    ("Integralização do Capital Social",              2, "item", ['291','891']),
    ("Remuneração das Disponibilidades do Tesouro",   2, "item", ['293','893']),
    ("Demais Receitas de Capital",                    2, "item", ['299','899']),
]

# Itens de fechamento/saldos de exercicios anteriores (linhas fixas, fora do
# loop generico de RECEITAS porque tem regras proprias e nao usam NatRec):
#   SUBTOTAL (III) = (I+II) | SUBTOTAL C/ REFIN (V) | DEFICIT (VI)
#   TOTAL (VII) | SALDOS EXERC. ANTERIORES | Recursos Arrec. Exerc. Ant.
#   Superavit Financeiro | Reabertura Creditos Adicionais
# NOTA: a planilha de equacoes nao tem uma linha propria para "Operacoes de
# Credito / Refinanciamento (IV)" -- no PDF modelo essa linha aparece
# zerada (R$0,00) em todas as colunas, pois o GDF nao pratica refinanciamento
# de dívida na execucao corrente. Mantido por consistencia estrutural com
# os mesmos prefixos de Operacoes de Credito (211/212), mas na pratica
# deve sair zerado.
OP_CREDITO_REFIN = [
    ("Operações de Crédito Internas",  ['211','811']),
    ("Operações de Crédito Externas",  ['212','812']),
]

# ─────────────────────────────────────────────────────────────────────────────
#  ESTRUTURA DE DESPESAS — QUADRO PRINCIPAL  (campo GND da planilha)
#  (nome, nivel, gnd)  -- GND = 2o digito de CONATUREZA (c.G.mm.ee.dd)
# ─────────────────────────────────────────────────────────────────────────────
DESPESAS_CORRENTES = [
    ("Pessoal e Encargos Sociais", '1'),
    ("Juros e Encargos da Dívida", '2'),
    ("Outras Despesas Correntes",  '3'),
]
DESPESAS_CAPITAL = [
    ("Investimentos",          '4'),
    ("Inversões Financeiras",  '5'),
    ("Amortização da Dívida",  '6'),
]
GND_RESERVA = '9'       # Reserva de Contingência (cod. 9.9.99.99.99)
GND_RESERVA_RPPS = '7'  # Reserva do RPPS

# ─────────────────────────────────────────────────────────────────────────────
#  RESTOS A PAGAR — mesma quebra por GND, contas proprias (531/631 NP;
#  532/632 Processados) conforme planilha de equacoes blocos 4 e 5.
# ─────────────────────────────────────────────────────────────────────────────
RP_GRUPOS = [
    ("Pessoal e Encargos Sociais", '1', 'corrente'),
    ("Outras Despesas Correntes",  '3', 'corrente'),
    ("Investimentos",              '4', 'capital'),
    ("Inversões Financeiras",      '5', 'capital'),
]

# ─────────────────────────────────────────────────────────────────────────────
#  CONEXAO  (identica ao DFC / Balanco Financeiro)
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
#  RECEITAS  —  busca por item via SUBSTR(COCONTACORRENTE, 1, N) = prefixo
#  As contas 521100000/521200000 sao contas-mae SEM tipo de conta corrente
#  (nunca recebem lancamento direto) -- confirmado no relatorio de conta
#  corrente: os lancamentos reais ocorrem nas contas-filha (521110000,
#  521120xxx, 521210000, 521220xxx, 521290000 etc.), todas tipo
#  "11 - Natureza da Receita". Por isso o filtro usa BETWEEN na faixa
#  completa da conta-mae, que cobre automaticamente todas as filhas.
#  col 1 (Previsao Inicial)    : SD 521100000-521199999 (faixa)
#  col 2 (Previsao Atualizada) : SD 521100000-521299999 (faixa)
#  col 3 (Receita Realizada)   : SC 621200000 - SD 6213XXXXX
# ─────────────────────────────────────────────────────────────────────────────
def _sql_receita_item(prefixos, mes):
    """Gera os 3 SUM(CASE WHEN...) de um item de receita p/ uma lista de
    prefixos de Natureza da Receita (podem ter comprimentos distintos)."""
    cond_parts = []
    for p in prefixos:
        n = len(p)
        cond_parts.append(f"SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,{n}) = '{p}'")
    cond = "(" + " OR ".join(cond_parts) + ")"

    prev_inicial = f"""
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 521100000 AND 521199999
              AND {cond}
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)"""

    prev_atual = f"""
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 521100000 AND 521299999
              AND {cond}
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)"""

    realizada = f"""
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL = 621200000
              AND {cond}
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND {cond}
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)"""

    return prev_inicial, prev_atual, realizada


def buscar_receitas(conn, mes, ano, coug):
    """
    Uma unica consulta cobrindo todos os itens de receita -- antes eram
    ~51 queries separadas (uma por item, escolhido originalmente para
    facilitar depuracao item a item), cada uma um full-scan de
    LANCAMENTOCONTABIL; agora os blocos de 3 colunas (inicial/atualizada/
    realizada) de cada item ficam na mesma SELECT, preservando as mesmas
    expressoes CASE, para varrer a tabela uma unica vez. Era o maior
    gargalo de bo.py (mais itens que qualquer outra busca do script).
    Retorna dict {nome_item: {'inicial':v,'atualizada':v,'realizada':v}}.
    """
    filtro_o = f"AND o.COUG = {coug}" if coug else ""
    cur = conn.cursor()
    resultado = {}

    itens = [r for r in RECEITAS if r[2] == "item"]
    print(f"  [receitas] MIL{ano}.LANCAMENTOCONTABIL — {len(itens)} itens em 1 consulta...")

    selects = []
    for nome, nivel, tipo, prefixos in itens:
        selects.extend(_sql_receita_item(prefixos, mes))

    sql = f"""
        SELECT /*+ PARALLEL(o, 4) */  {', '.join(selects)}
        FROM MIL{ano}.LANCAMENTOCONTABIL o
        WHERE 1=1 {filtro_o}
    """
    cur.execute(sql)
    row = cur.fetchone()

    for idx, (nome, nivel, tipo, prefixos) in enumerate(itens):
        p_ini, p_atu, p_rea = row[idx*3:idx*3+3]
        resultado[nome] = {
            'inicial':    float(p_ini or 0),
            'atualizada': float(p_atu or 0),
            'realizada':  float(p_rea or 0),
        }

    cur.close()
    return resultado


def buscar_op_credito_refinanciamento(conn, mes, ano, coug):
    """Operações de Crédito Internas/Externas (linha IV, Subtotal c/
    Refinanciamento), mesma logica de coluna das receitas -- 1 consulta
    para os itens de OP_CREDITO_REFIN em vez de uma por item."""
    filtro_o = f"AND o.COUG = {coug}" if coug else ""
    cur = conn.cursor()
    resultado = {}

    selects = []
    for nome, prefixos in OP_CREDITO_REFIN:
        selects.extend(_sql_receita_item(prefixos, mes))

    sql = f"""
        SELECT /*+ PARALLEL(o, 4) */  {', '.join(selects)}
        FROM MIL{ano}.LANCAMENTOCONTABIL o
        WHERE 1=1 {filtro_o}
    """
    cur.execute(sql)
    row = cur.fetchone()

    for idx, (nome, prefixos) in enumerate(OP_CREDITO_REFIN):
        p_ini, p_atu, p_rea = row[idx*3:idx*3+3]
        resultado[nome] = {
            'inicial':    float(p_ini or 0),
            'atualizada': float(p_atu or 0),
            'realizada':  float(p_rea or 0),
        }
    cur.close()
    return resultado


def buscar_saldos_exercicios_anteriores(conn, mes, ano, coug):
    """
    SALDOS DE EXERCICIOS ANTERIORES (Recursos Arrecadados em Exerc.
    Anteriores / Superavit Financeiro / Reabertura de Creditos Adicionais).
    Conforme planilha de equacoes (item 1.01.02.08.01.0), atualizada em
    28/08/2026 (a planilha oficial removeu as linhas de 522130101/522130102
    filtradas por grupo de fonte -- essa parcela ficou exclusiva do item
    1.01.02.08.02.0 Superavit Financeiro, que ja cobre a faixa completa
    522130100-522130199 sem filtro de fonte, evitando dupla contagem):
      Recursos Arrecadados, col2 (Atualizada):
        SD 521100000-521299999 com NatRec prefixo '999' (Funcao Receita 999)
      Recursos Arrecadados, col3 (Realizada):
        SC 621200000 - SD 621300000-621399999, com NatRec prefixo '999'
      Superavit Financeiro: col2 = SD 522130100-522130199 (faixa completa,
        sem filtro de fonte -- ja validado e batendo)
      Reabertura Cred. Adic.: col2 = SD 522120202
    """
    filtro_o = f"AND o.COUG = {coug}" if coug else ""
    cur = conn.cursor()

    cond_999 = "SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,3) = '999'"
    # Consolidado em 1 consulta (antes eram 3) -- mesmas expressoes CASE de
    # cada bloco, so reunidas na mesma SELECT para varrer a tabela 1 vez.
    sql = f"""
        SELECT /*+ PARALLEL(o, 4) */
          SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
                    AND o.COCONTACONTABIL BETWEEN 521100000 AND 521299999
                    AND {cond_999}
               THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
               ELSE 0 END) AS ATUALIZADA,
          SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
                    AND o.COCONTACONTABIL = 621200000 AND {cond_999}
               THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
               ELSE 0 END)
        - SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
                    AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999 AND {cond_999}
               THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
               ELSE 0 END) AS REALIZADA,
          SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
                          AND o.COCONTACONTABIL BETWEEN 522130100 AND 522130199
                     THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
                     ELSE 0 END) AS SUPERAVIT_FINANCEIRO,
          SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
                          AND o.COCONTACONTABIL = 522120202
                     THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
                     ELSE 0 END) AS REABERTURA_CREDITOS
        FROM MIL{ano}.LANCAMENTOCONTABIL o
        WHERE 1=1 {filtro_o}
    """
    cur.execute(sql)
    row = cur.fetchone()
    recursos_arrecadados = {'atualizada': float(row[0] or 0), 'realizada': float(row[1] or 0)}
    superavit_financeiro = float(row[2] or 0)
    reabertura_creditos  = float(row[3] or 0)

    cur.close()
    return {
        'recursos_arrecadados': recursos_arrecadados,
        'superavit_financeiro': {'atualizada': superavit_financeiro, 'realizada': 0.0},
        'reabertura_creditos':  {'atualizada': reabertura_creditos, 'realizada': 0.0},
    }


def buscar_saldo_521920500(conn, mes, ano):
    """Saldo SD da conta 521920500 (Previsao Adicional a Lancar) em SALDOCONTABIL
    ate INMES=mes, e tambem isolado so ate o mes ANTERIOR (INMES < mes) --
    esse segundo valor sinaliza saldo de mes(es) ja encerrado(s) que deveria
    ter zerado e nao zerou. Se != 0, explica o gap do equilibrio orcamentario
    (C18). Retorna (saldo_total, saldo_meses_anteriores)."""
    cur = conn.cursor()
    cur.execute(f"""
        SELECT
            NVL(SUM(CASE WHEN INMES <= {mes} THEN VADEBITO - VACREDITO ELSE 0 END), 0),
            NVL(SUM(CASE WHEN INMES <  {mes} THEN VADEBITO - VACREDITO ELSE 0 END), 0)
        FROM MIL{ano}.SALDOCONTABIL
        WHERE COCONTACONTABIL = 521920500
    """)
    row = cur.fetchone()
    val, val_ant = float(row[0] or 0), float(row[1] or 0)
    cur.close()
    return val, val_ant


# ─────────────────────────────────────────────────────────────────────────────
#  DESPESAS — QUADRO PRINCIPAL
#  As contas 622130100/300/400 (Empenhada/Liquidada) usam conta corrente
#  tipo "20 - Celula da Despesa com ND Detalhado" (layout fornecido pelo
#  usuario, confirmado por decomposicao de amostras reais):
#    Esfera (1) + Unidade Orcamentaria (5) + Programa de Trabalho (17)
#    + Fonte de Recursos (9) + Natureza da Despesa detalhada (8) = 40
#  A Natureza da Despesa ocupa os ultimos 8 caracteres (posicoes 33-40,
#  1-indexed para SUBSTR Oracle); o GND e o 2o digito dessa Natureza
#  (estrutura c.g.mm.ee.dd do PCASP) -> SUBSTR(COCONTACORRENTE,34,1).
#  A conta 622920104 (Paga) usa tipo "16 - Numero do Empenho" (11 digitos)
#  e por isso continua filtrada via JOIN com NOTAEMPENHO (NUNE), igual ao
#  padrao do DFC.
#    col1 Dotacao Inicial    : SD 522110000
#    col2 Dotacao Atualizada : SD 522110000 + SD 522120000 - SC 522150000 + SD 522190000
#    col3 Desp. Empenhadas   : SC 622130100-622130699, filtro por GND na celula
#    col4 Desp. Liquidadas   : SC 622130300+622130400+622130700, filtro por GND na celula
#    col5 Desp. Pagas        : SC 622920104, filtro por GND via NOTAEMPENHO (NUNE)
#  Reserva (GND 9, cod. 9.9.99.99.99) nao tem empenho -> so dotacao.
# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
#  DESPESAS — QUADRO PRINCIPAL
#  As contas-mae 522110000/522120000/522150000/522190000 NUNCA recebem
#  lancamento direto (sem tipo de conta corrente, confirmado no relatorio).
#  Os lancamentos reais ocorrem nas contas-filha (522110100/200/300,
#  522120100/201/202/301, 522150100/200/300, 522190101/109/401/409 etc.),
#  todas tipo "13 - Celula Orcamentaria da Despesa" (38 caracteres,
#  mesma estrutura ja decifrada nos Creditos Adicionais): o GND fica na
#  MESMA posicao absoluta da celula de despesa (SUBSTR(...,34,1)).
#  As contas 622130100/300/400 (Empenhada/Liquidada) usam conta corrente
#  tipo "20 - Celula da Despesa com ND Detalhado" (layout fornecido pelo
#  usuario, confirmado por decomposicao de amostras reais):
#    Esfera (1) + Unidade Orcamentaria (5) + Programa de Trabalho (17)
#    + Fonte de Recursos (9) + Natureza da Despesa detalhada (8) = 40
#  A Natureza da Despesa ocupa os ultimos 8 caracteres (posicoes 33-40,
#  1-indexed para SUBSTR Oracle); o GND e o 2o digito dessa Natureza
#  (estrutura c.g.mm.ee.dd do PCASP) -> SUBSTR(COCONTACORRENTE,34,1).
#  A conta 622920104 (Paga) usa tipo "16 - Numero do Empenho" (11 digitos)
#  e por isso continua filtrada via JOIN com NOTAEMPENHO (NUNE), igual ao
#  padrao do DFC.
#    col1 Dotacao Inicial    : SD 522110000-522119999, filtro GND na celula
#    col2 Dotacao Atualizada : SD 522110000-522129999 - SC 522150000-522159999
#                               + SD 522190000-522199999, filtro GND na celula
#    col3 Desp. Empenhadas   : SC 622130100-622130699, filtro por GND na celula
#    col4 Desp. Liquidadas   : SC 622130300+622130400+622130700, filtro por GND na celula
#    col5 Desp. Pagas        : SC 622920104, filtro por GND via NOTAEMPENHO (NUNE)
#  Reserva (GND 9, cod. 9.9.99.99.99) nao tem empenho -> so dotacao.
#
#  NOTA SOBRE PEQUENAS DIFERENCAS RESIDUAIS vs. O RELATORIO OFICIAL:
#  o relatorio oficial (PSIAG550) usa, para Despesas Liquidadas/Empenhadas,
#  a tabela MIL{ano}.BALANCOGERAL (saldos CONSOLIDADOS, filtrados por
#  CONATUREZA real de 6 digitos), enquanto este script usa LANCAMENTO-
#  CONTABIL diretamente (lancamentos individuais, fonte primaria). Em
#  investigacao detalhada (mes fechado, diferenca de R$414,55 em Pessoal),
#  confirmou-se que BALANCOGERAL e atualizada por um JOB PERIODICO (nao em
#  tempo real) -- um documento contabilmente correto e balanceado
#  (D=C, verificado lancamento a lancamento) processado em data tardia do
#  mes (ex.: dia 25) pode ainda nao ter sido propagado para BALANCOGERAL
#  no momento da extracao do relatorio oficial. Logo, residuos pequenos
#  entre este script e o relatorio oficial tendem a indicar que o SCRIPT
#  esta MAIS ATUALIZADO (usa a fonte primaria), nao que ha erro de calculo.
#  Se a diferenca persistir mesmo dias depois (apos novas rodadas do job),
#  vale reabrir a investigacao.
# ─────────────────────────────────────────────────────────────────────────────
def _sql_despesa_item(gnd, mes, ano, reserva=False):
    cond_celula = f"SUBSTR(o.COCONTACORRENTE,34,1) = '{gnd}'"
    cond_ne     = f"SUBSTR(ne.CONATUREZA,2,1) = '{gnd}'"

    dot_inicial = f"""
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 522110000 AND 522119999
              AND {cond_celula}
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)"""

    dot_atual = f"""
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 522110000 AND 522129999
              AND {cond_celula}
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 522150000 AND 522159999
              AND {cond_celula}
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 522190000 AND 522199999
              AND {cond_celula}
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)"""

    if reserva:
        # Reserva de Contingencia / RPPS: nao ha empenho (cod. 9.9.99.99.99)
        return dot_inicial, dot_atual, "0", "0", "0"

    empenhada = f"""
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 622130000 AND 622139999
              AND {cond_celula}
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)"""

    liquidada = f"""
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL IN (622130300,622130400,622130700)
              AND {cond_celula}
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)"""

    paga = f"""
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL = 622920104
              AND EXISTS (SELECT 1 FROM {ne_union(ano)} ne
                          WHERE ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                            AND ne.COUG     = o.COUGCONTAB
                            AND ne.COGESTAO = o.COGESTAOCONTAB
                            AND {cond_ne})
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)"""

    return dot_inicial, dot_atual, empenhada, liquidada, paga


def buscar_despesas(conn, mes, ano, coug):
    """
    Uma unica consulta cobrindo todos os GNDs (Pessoal/Juros/Outras Correntes/
    Investimentos/Inversoes/Amortizacao/Reserva de Contingencia/Reserva RPPS)
    — cada grupo vira um bloco de 5 colunas agregadas na mesma SELECT, em vez
    de uma query por grupo, para varrer LANCAMENTOCONTABIL uma unica vez.
    Retorna dict {nome: {dotacao_inicial, dotacao_atualizada, empenhada,
    liquidada, paga}}.
    """
    filtro_o = f"AND o.COUG = {coug}" if coug else ""
    cur = conn.cursor()
    resultado = {}

    itens = ([(nome, gnd, False) for nome, gnd in DESPESAS_CORRENTES + DESPESAS_CAPITAL]
             + [("Reserva de Contingência", GND_RESERVA, True),
                ("Reserva do RPPS", GND_RESERVA_RPPS, True)])
    print(f"  [despesas] MIL{ano}.LANCAMENTOCONTABIL — {len(itens)} grupos (GND) em 1 consulta...")

    selects = []
    for _, gnd, reserva in itens:
        selects.extend(_sql_despesa_item(gnd, mes, ano, reserva=reserva))

    sql = f"""
        SELECT /*+ PARALLEL(o, 4) */  {', '.join(selects)}
        FROM MIL{ano}.LANCAMENTOCONTABIL o
        WHERE 1=1 {filtro_o}
    """
    cur.execute(sql)
    row = cur.fetchone()

    for idx, (nome, _, _) in enumerate(itens):
        d_ini, d_atu, emp, liq, pag = row[idx*5:idx*5+5]
        resultado[nome] = {
            'dotacao_inicial':    float(d_ini or 0),
            'dotacao_atualizada': float(d_atu or 0),
            'empenhada':          float(emp or 0),
            'liquidada':          float(liq or 0),
            'paga':               float(pag or 0),
        }

    cur.close()
    return resultado


# ─────────────────────────────────────────────────────────────────────────────
#  QUADRO POR TIPOS DE CREDITOS ADICIONAIS
#  Colunas (campo "Coluna" da planilha, contas 522120100..522150300):
#    col1 Cred.Adic.Suplementar(a)        = SD 522120100
#    col2 Cred.Especiais Abertos(b)       = SD 522120201
#    col3 Cred.Especiais Reabertos(c)     = SD 522120202
#    col4 Cred.Extraordinarios Reabertos(d)= SD 522120301
#    col5 (-) Cancel. p/ Cred.Suplem.(e)  = SC 522150100
#    col6 (-) Remanejamento Veto LOA(f)   = SC 522150200
#    col7 (-) Cancel. p/ Cred.Especial(g) = SC 522150300
#  Conta corrente tipo "13 - Celula Orcamentaria da Despesa" (38 caracteres):
#    Esfera(1) + UO(5) + ProgTrabalho(17) + Fonte(9) + ND_parcial(6) = 38
#  Confirmado por decomposicao de amostras reais (categ+GND+modalidade+
#  elemento_parcial nos ultimos 6 digitos). O GND fica na MESMA posicao
#  absoluta que na celula de 40 chars (tipo 20): SUBSTR(...,34,1), pois
#  Esfera+UO+PT+Fonte (32 chars) e identico nos dois tipos.
# ─────────────────────────────────────────────────────────────────────────────
def _sql_creditos_item(gnd, mes, ano):
    cond_celula = f"SUBSTR(o.COCONTACORRENTE,34,1) = '{gnd}'"

    def conta_sd(conta):
        return f"""
        SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes} AND o.COCONTACONTABIL = {conta}
                  AND {cond_celula}
             THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
             ELSE 0 END)"""

    def conta_sc(conta):
        return f"""
        SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes} AND o.COCONTACONTABIL = {conta}
                  AND {cond_celula}
             THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
             ELSE 0 END)"""

    return (conta_sd(522120100), conta_sd(522120201), conta_sd(522120202),
            conta_sd(522120301), conta_sc(522150100), conta_sc(522150200),
            conta_sc(522150300))


def buscar_creditos_adicionais(conn, mes, ano, coug):
    """
    Uma unica consulta para todos os grupos (GND) — mesma logica de antes,
    so que os 7 blocos de 7 colunas ficam na mesma SELECT em vez de uma
    query por grupo, para varrer LANCAMENTOCONTABIL uma unica vez.
    """
    filtro_o = f"AND o.COUG = {coug}" if coug else ""
    cur = conn.cursor()
    resultado = {}

    grupos = DESPESAS_CORRENTES + DESPESAS_CAPITAL + [("Reserva de Contingência", GND_RESERVA)]
    print(f"  [creditos adicionais] MIL{ano} — {len(grupos)} grupos em 1 consulta...")

    selects = []
    for _, gnd in grupos:
        selects.extend(_sql_creditos_item(gnd, mes, ano))

    sql = f"""
        SELECT /*+ PARALLEL(o, 4) */  {', '.join(selects)}
        FROM MIL{ano}.LANCAMENTOCONTABIL o
        WHERE 1=1 {filtro_o}
    """
    cur.execute(sql)
    row = cur.fetchone()

    chaves = ['suplementar','esp_abertos','esp_reabertos','extraord_reabertos',
              'cancel_suplementar','remanej_veto','cancel_especial']
    for idx, (nome, _) in enumerate(grupos):
        vals = row[idx*7:idx*7+7]
        resultado[nome] = {k: float(v or 0) for k, v in zip(chaves, vals)}
    cur.close()
    return resultado


# ─────────────────────────────────────────────────────────────────────────────
#  RESTOS A PAGAR NAO PROCESSADOS (Quadro 3)
#    col1 Inscritos Exerc.Ant.(a)     = SD 531200000  -- tabela BALANCOGERAL
#    col2 Em 31/Dez Exerc.Ant.(b)     = SD 531100000  -- tabela BALANCOGERAL
#    col3 Liquidados(c)               = SC (631300000+631400000+631810000+631820000)
#    col4 Pagos(d)                    = SC (631400000+631820000)
#    col5 Cancelados(e)               = SC 631900000
#  Os saldos "Inscritos"/"Em 31/Dez" (contas 531xxx) vem da tabela
#  MIL{ano}.BALANCOGERAL, NAO de LANCAMENTOCONTABIL -- confirmado via SQL
#  real do Discoverer Oracle fornecido pelo usuario. Essa tabela ja tem o
#  GND pronto na coluna INCATEGORIA (= SUBSTR(CONATUREZA,2,1), validado
#  pela checagem 'correto'/'problemas' da query original), eliminando a
#  necessidade de JOIN com NOTAEMPENHO ou decomposicao de COCONTACORRENTE.
#  As contas de execucao (631xxx) continuam em LANCAMENTOCONTABIL, com
#  JOIN multi-schema via NUNE (igual ao padrao do DFC), pois sao tipo
#  "16 - Numero do Empenho".
# ─────────────────────────────────────────────────────────────────────────────
def _sql_rpnp_saldo_anterior(gnd, mes):
    """Inscritos em Exerc.Anteriores / Em 31-Dez do Exerc.Anterior, via
    MIL{ano}.BALANCOGERAL, filtrando por INCATEGORIA = GND. TO_CHAR torna
    o filtro seguro independente de INCATEGORIA ser NUMBER ou CHAR."""
    def conta_bg(conta):
        return f"""
        SUM(CASE WHEN b.INMES <= {mes} AND b.COCONTACONTABIL = {conta}
                  AND TO_CHAR(b.INCATEGORIA) = '{gnd}'
             THEN (b.VADEBITO - b.VACREDITO)
             ELSE 0 END)"""
    inscritos_ant = conta_bg(531200000)
    em_dez_ant    = conta_bg(531100000)
    return inscritos_ant, em_dez_ant


def _sql_rpnp_item(gnd, mes, ano):
    cond_ne = f"SUBSTR(ne.CONATUREZA,2,1) = '{gnd}'"

    def conta_sc(contas):
        lista = ",".join(str(c) for c in contas)
        return f"""
        SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes} AND o.COCONTACONTABIL IN ({lista})
                  AND EXISTS (SELECT 1 FROM {ne_union(ano)} ne
                              WHERE ne.NUNE = SUBSTR(o.COCONTACORRENTE,1,11)
                                AND ne.COUG = o.COUGCONTAB AND ne.COGESTAO = o.COGESTAOCONTAB
                                AND {cond_ne})
             THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
             ELSE 0 END)"""

    liquidados    = conta_sc([631300000, 631400000, 631810000, 631820000])
    pagos         = conta_sc([631400000, 631820000])
    cancelados    = conta_sc([631900000])
    return liquidados, pagos, cancelados


def buscar_rp_nao_processados(conn, mes, ano, coug):
    """
    Duas consultas no total (uma por tabela-fonte) em vez de uma por grupo:
    os 4 grupos (GND) viram blocos de colunas na mesma SELECT, preservando
    exatamente as mesmas expressoes CASE/EXISTS de antes.
    """
    filtro_o = f"AND o.COUG = {coug}" if coug else ""
    filtro_b = f"AND b.COUG = {coug}" if coug else ""
    cur = conn.cursor()
    resultado = {}

    print(f"  [RP nao processados] MIL{ano} — {len(RP_GRUPOS)} grupos em 2 consultas...")

    selects_saldo = []
    for _, gnd, _cat in RP_GRUPOS:
        selects_saldo.extend(_sql_rpnp_saldo_anterior(gnd, mes))
    sql_saldo = f"""
        SELECT /*+ PARALLEL(b, 4) */  {', '.join(selects_saldo)}
        FROM MIL{ano}.BALANCOGERAL b
        WHERE 1=1 {filtro_b}
    """
    cur.execute(sql_saldo)
    row_saldo = cur.fetchone()

    selects_exec = []
    for _, gnd, _cat in RP_GRUPOS:
        selects_exec.extend(_sql_rpnp_item(gnd, mes, ano))
    sql_exec = f"""
        SELECT /*+ PARALLEL(o, 4) */  {', '.join(selects_exec)}
        FROM MIL{ano}.LANCAMENTOCONTABIL o
        WHERE 1=1 {filtro_o}
    """
    cur.execute(sql_exec)
    row_exec = cur.fetchone()

    for idx, (nome, gnd, _cat) in enumerate(RP_GRUPOS):
        inscritos_ant, em_dez_ant = row_saldo[idx*2:idx*2+2]
        liquidados, pagos, cancelados = row_exec[idx*3:idx*3+3]
        resultado[nome] = {
            'inscritos_ant': float(inscritos_ant or 0),
            'em_dez_ant':    float(em_dez_ant or 0),
            'liquidados':    float(liquidados or 0),
            'pagos':         float(pagos or 0),
            'cancelados':    float(cancelados or 0),
        }
    cur.close()
    return resultado


# ─────────────────────────────────────────────────────────────────────────────
#  RESTOS A PAGAR PROCESSADOS (Quadro 4)
#    col1 Inscritos Exerc.Ant.(a) = SD 532200000  -- tabela BALANCOGERAL
#    col2 Em 31/Dez Exerc.Ant.(b) = SD 532100000  -- tabela BALANCOGERAL
#    col3 Pagos(c)  = SC (632210100+632210200+632210300+632210400)
#    col4 Cancelados(d) = SC 632900000
#  Mesmo padrao do RPNP: saldo anterior via BALANCOGERAL/INCATEGORIA,
#  execucao via LANCAMENTOCONTABIL/JOIN NOTAEMPENHO.
# ─────────────────────────────────────────────────────────────────────────────
def _sql_rpp_saldo_anterior(gnd, mes):
    def conta_bg(conta):
        return f"""
        SUM(CASE WHEN b.INMES <= {mes} AND b.COCONTACONTABIL = {conta}
                  AND TO_CHAR(b.INCATEGORIA) = '{gnd}'
             THEN (b.VADEBITO - b.VACREDITO)
             ELSE 0 END)"""
    inscritos_ant = conta_bg(532200000)
    em_dez_ant    = conta_bg(532100000)
    return inscritos_ant, em_dez_ant


def _sql_rpp_item(gnd, mes, ano):
    cond_ne = f"SUBSTR(ne.CONATUREZA,2,1) = '{gnd}'"

    def conta_sc(contas):
        lista = ",".join(str(c) for c in contas)
        return f"""
        SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes} AND o.COCONTACONTABIL IN ({lista})
                  AND EXISTS (SELECT 1 FROM {ne_union(ano)} ne
                              WHERE ne.NUNE = SUBSTR(o.COCONTACORRENTE,1,11)
                                AND ne.COUG = o.COUGCONTAB AND ne.COGESTAO = o.COGESTAOCONTAB
                                AND {cond_ne})
             THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
             ELSE 0 END)"""

    pagos         = conta_sc([632210100, 632210200, 632210300, 632210400])
    cancelados    = conta_sc([632900000])
    return pagos, cancelados


def buscar_rp_processados(conn, mes, ano, coug):
    filtro_o = f"AND o.COUG = {coug}" if coug else ""
    filtro_b = f"AND b.COUG = {coug}" if coug else ""
    cur = conn.cursor()
    resultado = {}

    print(f"  [RP processados] MIL{ano} — {len(RP_GRUPOS)} grupos em 2 consultas...")

    selects_saldo = []
    for _, gnd, _cat in RP_GRUPOS:
        selects_saldo.extend(_sql_rpp_saldo_anterior(gnd, mes))
    sql_saldo = f"""
        SELECT /*+ PARALLEL(b, 4) */  {', '.join(selects_saldo)}
        FROM MIL{ano}.BALANCOGERAL b
        WHERE 1=1 {filtro_b}
    """
    cur.execute(sql_saldo)
    row_saldo = cur.fetchone()

    selects_exec = []
    for _, gnd, _cat in RP_GRUPOS:
        selects_exec.extend(_sql_rpp_item(gnd, mes, ano))
    sql_exec = f"""
        SELECT /*+ PARALLEL(o, 4) */  {', '.join(selects_exec)}
        FROM MIL{ano}.LANCAMENTOCONTABIL o
        WHERE 1=1 {filtro_o}
    """
    cur.execute(sql_exec)
    row_exec = cur.fetchone()

    for idx, (nome, gnd, _cat) in enumerate(RP_GRUPOS):
        inscritos_ant, em_dez_ant = row_saldo[idx*2:idx*2+2]
        pagos, cancelados = row_exec[idx*2:idx*2+2]
        resultado[nome] = {
            'inscritos_ant': float(inscritos_ant or 0),
            'em_dez_ant':    float(em_dez_ant or 0),
            'pagos':         float(pagos or 0),
            'cancelados':    float(cancelados or 0),
        }
    cur.close()
    return resultado


# ─────────────────────────────────────────────────────────────────────────────
#  CALCULO DOS TOTAIS DERIVADOS
#  Reproduz as somas de linha do PDF modelo: subtotais de bloco (Correntes/
#  Capital), SUBTOTAL (III)=(I+II), SUBTOTAL c/Refin (V), TOTAL (VII),
#  Saldo = Realizada/Liquidada - Atualizada, Deficit/Superavit de ajuste.
# ─────────────────────────────────────────────────────────────────────────────
def _soma_chaves(dic, chaves, campo):
    return sum(dic.get(c, {}).get(campo, 0) for c in chaves)


def calcular_receitas(receitas_raw, opcred_raw, saldos_ant_raw):
    """Agrega itens em subtotais e calcula Saldo = Realizada - Atualizada.
    NOTA: 'opcred_raw' (Operacoes de Credito Internas/Externas) e mantido
    apenas por compatibilidade de assinatura -- NAO e mais somado em lugar
    nenhum, pois causava dupla contagem com as Operacoes de Credito ja
    presentes em RECEITAS DE CAPITAL (II). O balanco oficial do GDF nao
    exibe a linha IV (confirmado pelo usuario); ver OPERAÇÕES DE CRÉDITO /
    REFINANCIAMENTO (IV) abaixo, mantida zerada por consistencia estrutural."""
    r = {nome: v.copy() for nome, v in receitas_raw.items()}

    grupos_correntes = {
        'Impostos, Taxas e Contribuições de Melhoria': ['Impostos','Taxas','Contribuições de Melhoria'],
        'Receita de Contribuições': ['Contribuições Sociais','Contribuições Econômicas',
                                      'Contribuições p/ Entidades Priv. Interesse Púb.',
                                      'Contrib. p/ Custeio Serv. Iluminação Pública'],
        'Receita Patrimonial': ['Exploração do Patrimônio Imobiliário do Estado','Valores Mobiliários',
                                 'Delegação de Serv. Públ. Mediante Concessão','Exploração de Recursos Naturais',
                                 'Exploração do Patrimônio Intangível','Cessão de Direitos',
                                 'Demais Receitas Patrimoniais'],
        'Receita de Serviços': ['Serviços Administrativos e Comerciais Gerais',
                                 'Serviços e Atividades Ref. à Navegação e Transp.',
                                 'Serviços e Atividades Referentes à Saúde','Serviços e Atividades Financeiras',
                                 'Outros Serviços'],
        'Transferências Correntes': ['Transferências da União e de suas Entidades',
                                      'Transferências dos Estados, DF e de suas Entidades',
                                      'Transferências dos Municípios e de suas Entidades',
                                      'Transferências de Instituições Privadas',
                                      'Transferências de Outras Instituições Públicas',
                                      'Transferências do Exterior','Transferências de Pessoas Físicas',
                                      'Transferências Provenientes de Convênios','Demais Transferências Correntes'],
        'Outras Receitas Correntes': ['Multas Administrativas, Contratuais e Judiciais',
                                       'Indenizações, Restituições e Ressarcimentos',
                                       'Bens, Direitos e Valores Incorp. ao Patrimônio Públ.',
                                       'Multas e Juros de Mora das Receitas Tributárias',
                                       'Demais Receitas Correntes'],
    }
    for nome_grp, filhos in grupos_correntes.items():
        for campo in ('inicial','atualizada','realizada'):
            r[nome_grp] = r.get(nome_grp, {'inicial':0,'atualizada':0,'realizada':0})
            r[nome_grp][campo] = _soma_chaves(r, filhos, campo)

    receitas_correntes_filhos = list(grupos_correntes.keys()) + ['Receita Agropecuária','Receita Industrial']
    for campo in ('inicial','atualizada','realizada'):
        r['RECEITAS CORRENTES (I)'] = r.get('RECEITAS CORRENTES (I)', {'inicial':0,'atualizada':0,'realizada':0})
        r['RECEITAS CORRENTES (I)'][campo] = _soma_chaves(r, receitas_correntes_filhos, campo)

    grupos_capital = {
        'Operações de Crédito': ['Operações de Crédito - Mercado Interno','Operações de Crédito - Mercado Externo'],
        'Alienação de Bens': ['Alienação de Bens Móveis','Alienação de Bens Imóveis','Alienação de Bens Intangíveis'],
        'Transferências de Capital': ['Transferências da União e de suas Entidades (Capital)',
                                       'Transferências dos Estados, DF e de suas Entidades (Capital)',
                                       'Transferências dos Municípios e de suas Entidades (Capital)',
                                       'Transferências de Instituições Privadas (Capital)',
                                       'Transferências de Outras Instituições Públicas (Capital)',
                                       'Transferências do Exterior (Capital)','Demais Transferências de Capital'],
        'Outras Receitas de Capital': ['Integralização do Capital Social',
                                        'Remuneração das Disponibilidades do Tesouro','Demais Receitas de Capital'],
    }
    for nome_grp, filhos in grupos_capital.items():
        for campo in ('inicial','atualizada','realizada'):
            r[nome_grp] = r.get(nome_grp, {'inicial':0,'atualizada':0,'realizada':0})
            r[nome_grp][campo] = _soma_chaves(r, filhos, campo)

    receitas_capital_filhos = list(grupos_capital.keys()) + ['Amortizações de Empréstimos']
    for campo in ('inicial','atualizada','realizada'):
        r['RECEITAS DE CAPITAL (II)'] = r.get('RECEITAS DE CAPITAL (II)', {'inicial':0,'atualizada':0,'realizada':0})
        r['RECEITAS DE CAPITAL (II)'][campo] = _soma_chaves(r, receitas_capital_filhos, campo)

    subtotal = {}
    for campo in ('inicial','atualizada','realizada'):
        subtotal[campo] = r['RECEITAS CORRENTES (I)'][campo] + r['RECEITAS DE CAPITAL (II)'][campo]
    r['SUBTOTAL DAS RECEITAS (III)'] = subtotal

    # OPERACOES DE CREDITO / REFINANCIAMENTO (IV): conforme o MCASP 9a Ed.
    # (item 2.4.1, confirmado no texto literal do manual), esta linha
    # representa ESPECIFICAMENTE receita de operacoes de credito destinada
    # ao REFINANCIAMENTO da divida publica -- e NAO duplica as "Operacoes
    # de Credito" genericas que ja estao em RECEITAS DE CAPITAL (II).
    #
    # VERIFICACAO EMPIRICA (nao e suposicao): investigou-se TODAS as
    # naturezas de receita completas (8 digitos) de fato lancadas nas
    # contas 521100000-521299999/621200000 dentro do prefixo 211/212/811/
    # 812 em MIL2026/maio. Apenas 4 naturezas existem: 21125601, 21199901,
    # 21225401, 21299901 -- nenhuma com rubrica '1' (que e o padrao de
    # Refinanciamento, ex. 21112000/81112000 citados na planilha de
    # equacoes). A soma exata dessas 4 naturezas (R$ 409.906.409,00 na
    # Previsao Atualizada) bate 100% com o total de "Operacoes de Credito"
    # ja contabilizado em RECEITAS DE CAPITAL (II) -- ou seja, TODO o
    # credito do exercicio e operacao comum, NENHUM e refinanciamento.
    # Por isso a linha IV e seu detalhamento (Internas/Externas) ficam
    # zerados -- mas permanecem VISIVEIS no Excel por exigencia de
    # evidenciacao do MCASP, diferente do PDF oficial do GDF, que OMITE
    # esta linha por completo (uma inconsistencia de evidenciacao frente
    # ao manual). Se em exercicios futuros aparecer natureza com rubrica
    # '1' nessa faixa, ela devera ser segregada para a linha IV (e
    # subtraida de II), sem alterar o Subtotal (III+IV).
    r['OPERAÇÕES DE CRÉDITO / REFINANCIAMENTO (IV)'] = {
        'inicial': 0.0, 'atualizada': 0.0, 'realizada': 0.0,
    }
    for nome2 in ('Operações de Crédito Internas', 'Operações de Crédito Externas'):
        r[nome2] = {'inicial': 0.0, 'atualizada': 0.0, 'realizada': 0.0, 'saldo': 0.0}

    subtotal_refin = dict(subtotal)
    r['SUBTOTAL COM REFINANCIAMENTO (V)'] = subtotal_refin

    # DEFICIT (VI): so existe se a despesa total executada superar a receita
    # realizada total; calculado depois de termos despesas (ver calcular_tudo)
    r['DEFICIT (VI)'] = {'inicial':0.0,'atualizada':0.0,'realizada':0.0}

    total_vii = {}
    for campo in ('inicial','atualizada','realizada'):
        total_vii[campo] = subtotal_refin[campo] + r['DEFICIT (VI)'][campo]
    r['TOTAL (VII)'] = total_vii

    saldos = saldos_ant_raw
    saldos_total = {'atualizada': 0.0, 'realizada': 0.0}
    for k in ('recursos_arrecadados','superavit_financeiro','reabertura_creditos'):
        saldos_total['atualizada'] += saldos[k]['atualizada']
        saldos_total['realizada']  += saldos[k]['realizada']
    r['_saldos_exercicios_anteriores'] = saldos
    r['SALDOS DE EXERCÍCIOS ANTERIORES'] = {'inicial':0.0, **saldos_total}

    for nome, dic in r.items():
        if 'inicial' in dic:
            dic['saldo'] = dic.get('realizada',0) - dic.get('atualizada',0)
    return r


def calcular_despesas(despesas_raw):
    """Agrega GNDs em Despesas Correntes (VIII) / Capital (IX) e calcula
    Saldo da Dotacao = Dotacao Atualizada - Despesa Empenhada."""
    d = {nome: v.copy() for nome, v in despesas_raw.items()}

    def soma(nomes, campo):
        return sum(d[n][campo] for n in nomes if n in d)

    nomes_corr = [n for n, _ in DESPESAS_CORRENTES]
    nomes_cap  = [n for n, _ in DESPESAS_CAPITAL]
    campos = ['dotacao_inicial','dotacao_atualizada','empenhada','liquidada','paga']

    d['DESPESAS CORRENTES (VIII)'] = {c: soma(nomes_corr, c) for c in campos}
    d['DESPESAS DE CAPITAL (IX)']  = {c: soma(nomes_cap, c) for c in campos}

    subtotal = {c: d['DESPESAS CORRENTES (VIII)'][c] + d['DESPESAS DE CAPITAL (IX)'][c]
                + d.get('Reserva de Contingência', {}).get(c, 0)
                for c in campos}
    d['SUBTOTAL DAS DESPESAS (XI)'] = subtotal

    # Amortizacao da Divida / Refinanciamento (XII): GDF nao opera com
    # refinanciamento na execucao corrente -> mantido zerado (igual ao PDF
    # modelo, onde a linha XII aparece com 0,00 em todas as colunas).
    refin = {c: 0.0 for c in campos}
    d['AMORTIZAÇÃO DA DÍVIDA/REFINANCIAMENTO (XII)'] = refin

    subtotal_refin = {c: subtotal[c] + refin[c] for c in campos}
    d['SUBTOTAL COM REFINANCIAMENTO (XIII)'] = subtotal_refin

    # SUPERAVIT (XIV) e TOTAL (XV) dependem da Receita Realizada Total e
    # sao calculados em calcular_tudo() (junto com o DEFICIT das receitas),
    # pois representam o mesmo ajuste de equilibrio visto de lados opostos.
    # Aqui ficam zerados como placeholder.
    d['SUPERAVIT (XIV)'] = {c: 0.0 for c in campos}
    d['TOTAL (XV)'] = dict(subtotal_refin)

    for nome, dic in d.items():
        if 'dotacao_atualizada' in dic:
            dic['saldo_dotacao'] = dic['dotacao_atualizada'] - dic['empenhada']
    return d


def calcular_creditos_adicionais(creditos_raw):
    c = {nome: v.copy() for nome, v in creditos_raw.items()}
    campos = ['suplementar','esp_abertos','esp_reabertos','extraord_reabertos',
              'cancel_suplementar','remanej_veto','cancel_especial']

    def soma(nomes, campo):
        return sum(c[n][campo] for n in nomes if n in c)

    nomes_corr = [n for n, _ in DESPESAS_CORRENTES]
    nomes_cap  = [n for n, _ in DESPESAS_CAPITAL]
    c['DESPESAS CORRENTES'] = {f: soma(nomes_corr, f) for f in campos}
    c['DESPESAS DE CAPITAL'] = {f: soma(nomes_cap, f) for f in campos}

    total = {f: c['DESPESAS CORRENTES'][f] + c['DESPESAS DE CAPITAL'][f]
             + c.get('Reserva de Contingência', {}).get(f, 0) for f in campos}
    c['TOTAL'] = total

    for nome, dic in c.items():
        if 'suplementar' in dic:
            dic['total_alteracoes'] = (dic['suplementar'] + dic['esp_abertos'] + dic['esp_reabertos']
                                        + dic['extraord_reabertos'] - dic['cancel_suplementar']
                                        - dic['remanej_veto'] - dic['cancel_especial'])
    return c


def calcular_rp_nao_processados(rp_raw):
    rp = {nome: v.copy() for nome, v in rp_raw.items()}
    campos = ['inscritos_ant','em_dez_ant','liquidados','pagos','cancelados']

    def soma(nomes, campo):
        return sum(rp[n][campo] for n in nomes if n in rp)

    correntes = [n for n,_,cat in RP_GRUPOS if cat=='corrente']
    capital   = [n for n,_,cat in RP_GRUPOS if cat=='capital']
    rp['DESPESAS CORRENTES'] = {f: soma(correntes, f) for f in campos}
    rp['DESPESAS DE CAPITAL'] = {f: soma(capital, f) for f in campos}
    total = {f: rp['DESPESAS CORRENTES'][f] + rp['DESPESAS DE CAPITAL'][f] for f in campos}
    rp['TOTAL'] = total

    for nome, dic in rp.items():
        if 'inscritos_ant' in dic:
            dic['saldo'] = dic['inscritos_ant'] + dic['em_dez_ant'] - dic['pagos'] - dic['cancelados']
    return rp


def calcular_rp_processados(rp_raw):
    rp = {nome: v.copy() for nome, v in rp_raw.items()}
    campos = ['inscritos_ant','em_dez_ant','pagos','cancelados']

    def soma(nomes, campo):
        return sum(rp[n][campo] for n in nomes if n in rp)

    correntes = [n for n,_,cat in RP_GRUPOS if cat=='corrente']
    capital   = [n for n,_,cat in RP_GRUPOS if cat=='capital']
    rp['DESPESAS CORRENTES'] = {f: soma(correntes, f) for f in campos}
    rp['DESPESAS DE CAPITAL'] = {f: soma(capital, f) for f in campos}
    total = {f: rp['DESPESAS CORRENTES'][f] + rp['DESPESAS DE CAPITAL'][f] for f in campos}
    rp['TOTAL'] = total

    for nome, dic in rp.items():
        if 'inscritos_ant' in dic:
            dic['saldo'] = dic['inscritos_ant'] + dic['em_dez_ant'] - dic['pagos'] - dic['cancelados']
    return rp


def calcular_tudo(receitas_raw, opcred_raw, saldos_ant_raw, despesas_raw,
                   creditos_raw, rpnp_raw, rpp_raw):
    """
    Combina os 5 quadros e resolve o ajuste de equilibrio entre o lado da
    Receita e o lado da Despesa (MCASP 9a Ed., item 2.4.1):
      - Se Despesa Empenhada > Receita Realizada (SUBTOTAL III/V) ->
        aparece DEFICIT (VI) do lado das receitas, igualando TOTAL (VII)
        a SUBTOTAL COM REFINANCIAMENTO (XIII) das despesas.
      - Se Receita Realizada > Despesa Empenhada (SUBTOTAL XIII) -> aparece
        SUPERAVIT (XIV) do lado das despesas, igualando TOTAL (XV) ao
        TOTAL (VII) da receita.
    Os dois nunca coexistem (um dos dois fica zerado).
    """
    receitas = calcular_receitas(receitas_raw, opcred_raw, saldos_ant_raw)
    despesas = calcular_despesas(despesas_raw)
    creditos = calcular_creditos_adicionais(creditos_raw)
    rpnp     = calcular_rp_nao_processados(rpnp_raw)
    rpp      = calcular_rp_processados(rpp_raw)

    campos_desp = ['dotacao_inicial','dotacao_atualizada','empenhada','liquidada','paga']
    subtotal_refin_receita = receitas['SUBTOTAL COM REFINANCIAMENTO (V)']['realizada']
    subtotal_refin_despesa = despesas['SUBTOTAL COM REFINANCIAMENTO (XIII)']['empenhada']
    ajuste = subtotal_refin_receita - subtotal_refin_despesa

    if ajuste < 0:
        # Despesa empenhada supera a receita realizada -> DEFICIT (receitas)
        receitas['DEFICIT (VI)']['realizada'] = ajuste  # negativo
        for campo in ('inicial','atualizada'):
            receitas['DEFICIT (VI)'][campo] = 0.0
        receitas['TOTAL (VII)'] = {
            'inicial':    receitas['SUBTOTAL COM REFINANCIAMENTO (V)']['inicial'],
            'atualizada': receitas['SUBTOTAL COM REFINANCIAMENTO (V)']['atualizada'],
            'realizada':  subtotal_refin_receita + ajuste,
        }
        despesas['SUPERAVIT (XIV)'] = {c: 0.0 for c in campos_desp}
        despesas['TOTAL (XV)'] = dict(despesas['SUBTOTAL COM REFINANCIAMENTO (XIII)'])
    else:
        # Receita realizada supera (ou iguala) a despesa empenhada -> SUPERAVIT (despesas)
        receitas['DEFICIT (VI)'] = {'inicial':0.0,'atualizada':0.0,'realizada':0.0}
        receitas['TOTAL (VII)'] = dict(receitas['SUBTOTAL COM REFINANCIAMENTO (V)'])
        superavit = {c: 0.0 for c in campos_desp}
        superavit['empenhada'] = ajuste
        despesas['SUPERAVIT (XIV)'] = superavit
        despesas['TOTAL (XV)'] = {c: despesas['SUBTOTAL COM REFINANCIAMENTO (XIII)'][c] + superavit[c]
                                    for c in campos_desp}

    for nome in ('DEFICIT (VI)', 'TOTAL (VII)'):
        receitas[nome]['saldo'] = receitas[nome].get('realizada',0) - receitas[nome].get('atualizada',0)
    for nome in ('SUPERAVIT (XIV)', 'TOTAL (XV)'):
        despesas[nome]['saldo_dotacao'] = despesas[nome]['dotacao_atualizada'] - despesas[nome]['empenhada']

    return {
        'receitas': receitas, 'despesas': despesas,
        'creditos_adicionais': creditos,
        'rp_nao_processados': rpnp, 'rp_processados': rpp,
    }


# ─────────────────────────────────────────────────────────────────────────────
#  AUDITORIA DE INTEGRIDADE  (mesmo espirito do Balanco Financeiro)
# ─────────────────────────────────────────────────────────────────────────────
def _mes_encerrado(mes, ano):
    """True se hoje ja passou do ultimo dia de (mes, ano) -- usado para
    escalar o C18 (521920500) de INFO para ALERTA quando o prazo ("deve
    zerar ate o fim do mes") ja passou e o saldo persiste."""
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    return datetime.now().date() > datetime(ano, mes, ultimo_dia).date()


def auditoria_integridade(t, saldo_521920500=0.0, saldo_521920500_ant=0.0, mes=None, ano=None):
    achados = []
    rec = t['receitas']; des = t['despesas']

    # Controle 1: equilibrio orcamentario (MCASP 2.1) -- TOTAL(VII) Previsao
    # Atualizada + (Superavit Financeiro + Reabertura de Creditos Adicionais)
    # deve igualar a Dotacao Atualizada (Despesas). "Recursos Arrecadados em
    # Exercicios Anteriores" e uma rubrica informativa dentro de "Saldos de
    # Exercicios Anteriores" que NAO entra nesta equacao de equilibrio --
    # confirmado batendo exato (diferenca = 0,00) com os valores do PDF
    # modelo (maio/2026) quando excluido.
    saldos_ant = rec['_saldos_exercicios_anteriores']
    financiamento_atualizada = (saldos_ant['superavit_financeiro']['atualizada']
                                 + saldos_ant['reabertura_creditos']['atualizada'])
    receita_lado = rec['TOTAL (VII)']['atualizada'] + financiamento_atualizada
    despesa_lado = des['SUBTOTAL COM REFINANCIAMENTO (XIII)']['dotacao_atualizada']
    dif = receita_lado - despesa_lado
    if abs(dif) < 1.00:
        achados.append(('OK', 'Equilíbrio Orçamentário (MCASP 2.1)',
                         f'Diferença: {dif:,.2f} (fecha exato)'))
    else:
        achados.append(('ERRO', 'Equilíbrio Orçamentário (MCASP 2.1)',
                         f'Diferença: {dif:,.2f} — Previsão Atualizada + '
                         f'(Superávit Financeiro + Reabertura de Créditos) não '
                         f'igualou a Dotação Atualizada. Recursos Arrecadados em '
                         f'Exercícios Anteriores é excluído desta conta por ser '
                         f'rubrica informativa (não financia dotação adicional).'))
        if saldo_521920500 != 0.0 and abs(dif + saldo_521920500) < 1.00:
            mes_ja_encerrou = mes is not None and ano is not None and _mes_encerrado(mes, ano)
            saldo_de_mes_anterior = abs(saldo_521920500_ant) >= 1.00
            if mes_ja_encerrou or saldo_de_mes_anterior:
                achados.append(('ALERTA', 'C18 — Previsão Adicional a Lançar NÃO resolvida até o fim do mês (521920500)',
                                 f'Saldo 521920500 = {saldo_521920500:,.2f} (dos quais '
                                 f'{saldo_521920500_ant:,.2f} já vem de mês(es) anterior(es) '
                                 f'já encerrado(s)) — deveria ter zerado até o fim do mês e '
                                 f'não zerou. Verificar (ver C18).'))
            else:
                achados.append(('INFO', 'C18 — Previsão Adicional a Lançar (521920500)',
                                 f'Saldo 521920500 = {saldo_521920500:,.2f} corresponde '
                                 f'exatamente à diferença acima. Quando os lançamentos '
                                 f'pendentes forem reclassificados ao fim do mês a '
                                 f'diferença zerará automaticamente (ver C18).'))

    # Controle 2: Despesa Empenhada >= Liquidada >= Paga (cada GND)
    seq_erro = False
    for nome, gnd in DESPESAS_CORRENTES + DESPESAS_CAPITAL:
        d = des.get(nome, {})
        if d.get('empenhada',0) + 0.01 < d.get('liquidada',0):
            achados.append(('ERRO', f'Empenhada < Liquidada em {nome}',
                             f'Emp={d.get("empenhada",0):,.2f}  Liq={d.get("liquidada",0):,.2f}'))
            seq_erro = True
        if d.get('liquidada',0) + 0.01 < d.get('paga',0):
            achados.append(('ERRO', f'Liquidada < Paga em {nome}',
                             f'Liq={d.get("liquidada",0):,.2f}  Paga={d.get("paga",0):,.2f}'))
            seq_erro = True
    if not seq_erro:
        achados.append(('OK', 'Sequência Empenhada >= Liquidada >= Paga', 'Nenhuma inconsistência por GND'))

    # Controle 3: sinais nao-negativos esperados
    sinal_erro = False
    for nome in ('RECEITAS CORRENTES (I)','RECEITAS DE CAPITAL (II)'):
        v = rec.get(nome, {}).get('realizada', 0)
        if v < -0.01:
            achados.append(('ERRO', f'Sinal de {nome}', f'Valor negativo inesperado: {v:,.2f}'))
            sinal_erro = True
    for nome, _ in DESPESAS_CORRENTES + DESPESAS_CAPITAL:
        v = des.get(nome, {}).get('empenhada', 0)
        if v < -0.01:
            achados.append(('ERRO', f'Sinal de Empenhada em {nome}', f'Valor negativo: {v:,.2f}'))
            sinal_erro = True
    if not sinal_erro:
        achados.append(('OK', 'Sinais dos itens principais', 'Nenhum valor negativo inesperado'))

    # Controle 4: sinais nao-negativos em Restos a Pagar (saldo final nunca
    # deveria ser negativo -- indicaria mais pago/cancelado do que inscrito)
    rpnp = t.get('rp_nao_processados', {}); rpp = t.get('rp_processados', {})
    sinal_rp_erro = False
    for nome, dic in list(rpnp.items()) + list(rpp.items()):
        saldo = dic.get('saldo')
        if saldo is not None and saldo < -0.01:
            achados.append(('ERRO', f'Saldo de RP negativo em {nome}', f'Saldo: {saldo:,.2f}'))
            sinal_rp_erro = True
    if not sinal_rp_erro:
        achados.append(('OK', 'Sinais de Restos a Pagar', 'Nenhum saldo final negativo'))

    return achados


def auditoria_defasagem_balancogeral(conn, ano, mes, coug=None):
    """
    Controle adicional (requer conexao ativa): compara o total bruto de
    MIL{ano}.LANCAMENTOCONTABIL com MIL{ano}.BALANCOGERAL para as contas
    de execucao da despesa (622130000-622139999), SEM quebrar por GND/
    natureza. Os dois devem ser EXATAMENTE iguais, pois ambas as tabelas
    deveriam refletir os mesmos lancamentos.
    Se houver diferenca, ela indica que BALANCOGERAL (tabela CONSOLIDADA,
    atualizada por job periodico -- nao em tempo real) ainda nao processou
    todos os lancamentos mais recentes de LANCAMENTOCONTABIL (fonte
    primaria). Avisa o usuario sobre a magnitude da defasagem, mas isso
    NAO invalida o calculo do script (que usa a fonte primaria, mais
    atualizada) -- serve apenas para alertar que o relatorio oficial
    (PSIAG550), que consulta BALANCOGERAL, pode estar temporariamente
    desatualizado em relacao a este script nessa magnitude.
    """
    cur = conn.cursor()
    filtro = f"AND COUG = {coug}" if coug else ""
    q1 = f"""SELECT SUM(DECODE(INDEBITOCREDITO,'C',VALANCAMENTO,'D',-VALANCAMENTO,0))
            FROM MIL{ano}.LANCAMENTOCONTABIL
            WHERE COCONTACONTABIL BETWEEN 622130000 AND 622139999
              AND INMES BETWEEN 1 AND {mes} {filtro}"""
    cur.execute(q1)
    total_lc = float(cur.fetchone()[0] or 0)

    q2 = f"""SELECT SUM(VACREDITO - VADEBITO)
            FROM MIL{ano}.BALANCOGERAL
            WHERE COCONTACONTABIL BETWEEN 622130000 AND 622139999
              AND INMES <= {mes} {filtro}"""
    cur.execute(q2)
    total_bg = float(cur.fetchone()[0] or 0)
    cur.close()

    dif = total_lc - total_bg
    if abs(dif) < 1.00:
        return ('OK', 'Defasagem LANCAMENTOCONTABIL vs BALANCOGERAL',
                f'Diferença: {dif:,.2f} (tabelas sincronizadas)')
    else:
        return ('ERRO', 'Defasagem LANCAMENTOCONTABIL vs BALANCOGERAL',
                f'Diferença: {dif:,.2f} — BALANCOGERAL (job periódico) ainda não '
                f'processou todos os lançamentos recentes de LANCAMENTOCONTABIL '
                f'(fonte primária). O script usa a fonte primária, então está '
                f'correto/atualizado; o relatório oficial pode estar temporariamente '
                f'desatualizado nessa magnitude até a próxima execução do job.')


def imprimir_auditoria(achados):
    icones = {'OK': '✔', 'ERRO': '✘', 'ALERTA': '⚠', 'INFO': 'ℹ'}
    print(f"\n{'─'*60}")
    print("  AUDITORIA DE INTEGRIDADE")
    print(f"{'─'*60}")
    for status, titulo, detalhe in achados:
        ic = icones.get(status, '?')
        print(f"  {ic} [{status:<6}] {titulo}")
        print(f"            {detalhe}")
    n_erro = sum(1 for s,_,_ in achados if s == 'ERRO')
    print(f"{'─'*60}")
    if n_erro:
        print(f"  RESULTADO: {n_erro} pendencia(s) sinalizada(s) - ver detalhes acima.")
    else:
        print(f"  RESULTADO: TODOS OS CONTROLES PASSARAM.")
    print(f"{'─'*60}")


# ─────────────────────────────────────────────────────────────────────────────
#  EXCEL  —  layout espelha ListaBalancoOrcamentario.pdf (Anexo 12)
# ─────────────────────────────────────────────────────────────────────────────
FMT_BRL = '#,##0.00'
AZUL_TIT  = "2E5C8A"
CINZA_HDR = "DCE6F1"
AMARELO   = "9DC3E6"
BRANCO    = "F5F9FD"

def _fill(h): return PatternFill('solid', fgColor=h)
def _side(s='thin'): return Side(border_style=s, color='B8CCE4')
B_THIN = Border(left=_side(), right=_side(), top=_side(), bottom=_side())
B_MED  = Border(left=_side('medium'), right=_side('medium'), top=_side('medium'), bottom=_side('medium'))
F_GRAY = _fill(CINZA_HDR); F_YELL = _fill(AMARELO); F_WHIT = _fill(BRANCO)

def cel(ws, r, c, v='', bold=False, sz=9, bg=None, ha='left', brd=None, fmt=None, ind=0):
    x = ws.cell(row=r, column=c, value=v)
    x.font = Font(name='Arial', bold=bold, size=sz, color='000000')
    x.alignment = Alignment(horizontal=ha, vertical='center', indent=ind)
    if bg: x.fill = bg
    if brd: x.border = brd
    if fmt and v not in ('', None): x.number_format = fmt
    return x


def _linha_receita(ws, r, nome, dic, nivel, bold=False, bg=None):
    ws.row_dimensions[r].height = 14
    cel(ws, r, 1, nome, bold=bold, bg=bg, brd=B_THIN, ha='left', ind=nivel)
    cel(ws, r, 2, dic.get('inicial',0),    bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 3, dic.get('atualizada',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 4, dic.get('realizada',0),  bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 5, dic.get('saldo',0),      bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)


def _linha_despesa(ws, r, nome, dic, nivel, bold=False, bg=None):
    ws.row_dimensions[r].height = 14
    cel(ws, r, 1, nome, bold=bold, bg=bg, brd=B_THIN, ha='left', ind=nivel)
    cel(ws, r, 2, dic.get('dotacao_inicial',0),    bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 3, dic.get('dotacao_atualizada',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 4, dic.get('empenhada',0),          bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 5, dic.get('liquidada',0),          bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 6, dic.get('paga',0),               bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 7, dic.get('saldo_dotacao',0),      bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)


def _cabecalho_pagina(ws, titulo, mes, ano, ug_label, ncols):
    ws.merge_cells(f'A1:{chr(64+ncols)}1')
    c = ws['A1']
    c.value = f'GOVERNO DO DISTRITO FEDERAL\nBALANÇO ORÇAMENTÁRIO — {titulo}'
    c.font = Font(name='Arial', bold=True, size=12, color=AZUL_TIT)
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f'A2:{chr(64+ncols)}2')
    cel(ws, 2, 1, f'Exercício {ano}  |  Mês de Referência: {mes:02d} - {MESES[mes]}  |  {ug_label}'
                  f'  |  Posição em: {datetime.now():%d/%m/%Y %H:%M:%S}', sz=9)
    ws.row_dimensions[3].height = 4


def gerar_excel(t, mes, ano, ug_label, output_path, achados=None):
    wb = openpyxl.Workbook()
    rec = t['receitas']; des = t['despesas']
    cred = t['creditos_adicionais']; rpnp = t['rp_nao_processados']; rpp = t['rp_processados']

    # ── ABA 1: RECEITAS ─────────────────────────────────────────────────
    ws = wb.active; ws.title = 'Receitas'
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'
    for col, w in zip('ABCDE', [42, 18, 18, 18, 18]):
        ws.column_dimensions[col].width = w
    _cabecalho_pagina(ws, 'Quadro Principal — Receitas', mes, ano, ug_label, 5)

    H = 4; ws.row_dimensions[H].height = 28
    for c_idx, txt in enumerate(['RECEITAS ORÇAMENTÁRIAS','Previsão Inicial (a)',
                                  'Previsão Atualizada (b)','Receitas Realizadas (c)',
                                  'Saldo (d)=(c-b)'], 1):
        cel(ws, H, c_idx, txt, bold=True, bg=F_GRAY, brd=B_THIN, ha='center')

    r = H + 1
    for nome, nivel, tipo, _prefixos in RECEITAS:
        dic = rec.get(nome, {'inicial':0,'atualizada':0,'realizada':0,'saldo':0})
        bold = tipo in ("subtotal_calc",) or nivel == 0
        bg = F_GRAY if nivel == 0 else None
        _linha_receita(ws, r, nome, dic, nivel, bold=bold, bg=bg)
        r += 1

    for nome in ('SUBTOTAL DAS RECEITAS (III)',
                 'OPERAÇÕES DE CRÉDITO / REFINANCIAMENTO (IV)'):
        dic = rec.get(nome, {'inicial':0,'atualizada':0,'realizada':0,'saldo':0})
        bold0 = (nome == 'SUBTOTAL DAS RECEITAS (III)')
        _linha_receita(ws, r, nome, dic, 0, bold=True, bg=F_GRAY); r += 1
        if nome == 'OPERAÇÕES DE CRÉDITO / REFINANCIAMENTO (IV)':
            # Detalhamento exigido pelo MCASP 9a Ed. (2.4.1): segregar em
            # Internas/Externas mesmo quando o total da linha IV e zero --
            # o GDF nao opera refinanciamento de divida, mas a linha deve
            # permanecer evidenciada por estrutura (o PDF oficial do GDF
            # omite essa linha, o que e uma inconsistencia de evidenciacao
            # frente ao MCASP; aqui mantemos visivel por completude).
            opcred = {nome2: v for nome2, v in rec.items()
                      if nome2 in ('Operações de Crédito Internas','Operações de Crédito Externas')}
            for nome2 in ('Operações de Crédito Internas','Operações de Crédito Externas'):
                dic2 = opcred.get(nome2, {'inicial':0,'atualizada':0,'realizada':0,'saldo':0})
                _linha_receita(ws, r, nome2, dic2, 1); r += 1

    for nome in ('SUBTOTAL COM REFINANCIAMENTO (V)',
                 'DEFICIT (VI)', 'TOTAL (VII)'):
        dic = rec.get(nome, {'inicial':0,'atualizada':0,'realizada':0,'saldo':0})
        _linha_receita(ws, r, nome, dic, 0, bold=True, bg=F_GRAY); r += 1

    cel(ws, r, 1, 'SALDOS DE EXERCÍCIOS ANTERIORES', bold=True, bg=F_YELL, brd=B_THIN); 
    cel(ws, r, 2, 0, bg=F_YELL, brd=B_THIN, fmt=FMT_BRL)
    cel(ws, r, 3, rec['SALDOS DE EXERCÍCIOS ANTERIORES']['atualizada'], bold=True, bg=F_YELL, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 4, rec['SALDOS DE EXERCÍCIOS ANTERIORES']['realizada'], bold=True, bg=F_YELL, brd=B_THIN, ha='right', fmt=FMT_BRL)
    cel(ws, r, 5, '', bg=F_YELL, brd=B_THIN)
    r += 1
    saldos_ant = rec['_saldos_exercicios_anteriores']
    for label, key in [('  Recursos Arrecadados em Exercícios Anteriores','recursos_arrecadados'),
                        ('  Superávit Financeiro','superavit_financeiro'),
                        ('  Reabertura de Créditos Adicionais','reabertura_creditos')]:
        d = saldos_ant[key]
        cel(ws, r, 1, label, brd=B_THIN, ind=1)
        cel(ws, r, 2, '', brd=B_THIN)
        cel(ws, r, 3, d['atualizada'], brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 4, d['realizada'], brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 5, '', brd=B_THIN)
        r += 1

    # ── ABA 2: DESPESAS ──────────────────────────────────────────────────
    ws2 = wb.create_sheet('Despesas')
    ws2.sheet_view.showGridLines = False
    ws2.page_setup.orientation = 'landscape'
    for col, w in zip('ABCDEFG', [38, 16, 16, 16, 16, 16, 16]):
        ws2.column_dimensions[col].width = w
    _cabecalho_pagina(ws2, 'Quadro Principal — Despesas', mes, ano, ug_label, 7)

    H = 4; ws2.row_dimensions[H].height = 28
    for c_idx, txt in enumerate(['DESPESAS ORÇAMENTÁRIAS','Dotação Inicial (e)',
                                  'Dotação Atualizada (f)','Despesas Empenhadas (g)',
                                  'Despesas Liquidadas (h)','Despesas Pagas (i)',
                                  'Saldo da Dotação (j)=(f-g)'], 1):
        cel(ws2, H, c_idx, txt, bold=True, bg=F_GRAY, brd=B_THIN, ha='center')

    r = H + 1
    _linha_despesa(ws2, r, 'DESPESAS CORRENTES (VIII)', des['DESPESAS CORRENTES (VIII)'], 0, bold=True, bg=F_GRAY); r += 1
    for nome, _gnd in DESPESAS_CORRENTES:
        _linha_despesa(ws2, r, nome, des[nome], 1); r += 1
    _linha_despesa(ws2, r, 'DESPESAS DE CAPITAL (IX)', des['DESPESAS DE CAPITAL (IX)'], 0, bold=True, bg=F_GRAY); r += 1
    for nome, _gnd in DESPESAS_CAPITAL:
        _linha_despesa(ws2, r, nome, des[nome], 1); r += 1
    _linha_despesa(ws2, r, 'RESERVA DE CONTINGÊNCIA (X)', des['Reserva de Contingência'], 0, bold=True, bg=F_GRAY); r += 1
    _linha_despesa(ws2, r, 'RESERVA DO RPPS', des['Reserva do RPPS'], 0, bold=True, bg=F_GRAY); r += 1
    for nome in ('SUBTOTAL DAS DESPESAS (XI)','AMORTIZAÇÃO DA DÍVIDA/REFINANCIAMENTO (XII)',
                 'SUBTOTAL COM REFINANCIAMENTO (XIII)','SUPERAVIT (XIV)','TOTAL (XV)'):
        _linha_despesa(ws2, r, nome, des[nome], 0, bold=True, bg=F_GRAY); r += 1

    D = r + 1
    cel(ws2, D, 1, 'Equilíbrio: TOTAL(VII)+Saldos Exerc.Ant. [Receitas, Atualizada] '
                   'vs Dotação Atualizada (Despesas):', sz=8)
    eq = (rec['TOTAL (VII)']['atualizada'] + rec['SALDOS DE EXERCÍCIOS ANTERIORES']['atualizada']
          - des['SUBTOTAL COM REFINANCIAMENTO (XIII)']['dotacao_atualizada'])
    cel(ws2, D, 4, eq, sz=8, ha='right', fmt=FMT_BRL)

    # ── ABA 3: CRÉDITOS ADICIONAIS ───────────────────────────────────────
    ws3 = wb.create_sheet('Créditos Adicionais')
    ws3.sheet_view.showGridLines = False
    ws3.page_setup.orientation = 'landscape'
    for col, w in zip('ABCDEFGH', [38, 15, 15, 15, 15, 15, 15, 16]):
        ws3.column_dimensions[col].width = w
    _cabecalho_pagina(ws3, 'Quadro por Tipos de Créditos Adicionais', mes, ano, ug_label, 8)

    H = 4; ws3.row_dimensions[H].height = 30
    for c_idx, txt in enumerate(['DESPESAS ORÇAMENTÁRIAS','Créd.Adic.Suplementar (a)',
                                  'Créd.Especiais Abertos (b)','Créd.Especiais Reabertos (c)',
                                  'Créd.Extraordinários Reabertos (d)',
                                  '(-) Cancel.Dotação p/Créd.Suplementar (e)',
                                  '(-) Remanejamento Veto LOA (f)',
                                  '(-) Cancel.Dotação p/Créd.Especial (g) / Total (h)'], 1):
        cel(ws3, H, c_idx, txt, bold=True, bg=F_GRAY, brd=B_THIN, ha='center')

    def _linha_credito(ws, r, nome, dic, nivel, bold=False, bg=None):
        ws.row_dimensions[r].height = 14
        cel(ws, r, 1, nome, bold=bold, bg=bg, brd=B_THIN, ha='left', ind=nivel)
        for c_idx, campo in enumerate(['suplementar','esp_abertos','esp_reabertos',
                                        'extraord_reabertos','cancel_suplementar',
                                        'remanej_veto'], 2):
            cel(ws, r, c_idx, dic.get(campo,0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 8, dic.get('total_alteracoes',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)

    r = H + 1
    _linha_credito(ws3, r, 'DESPESAS CORRENTES', cred['DESPESAS CORRENTES'], 0, bold=True, bg=F_GRAY); r += 1
    for nome, _gnd in DESPESAS_CORRENTES:
        _linha_credito(ws3, r, nome, cred[nome], 1); r += 1
    _linha_credito(ws3, r, 'DESPESAS DE CAPITAL', cred['DESPESAS DE CAPITAL'], 0, bold=True, bg=F_GRAY); r += 1
    for nome, _gnd in DESPESAS_CAPITAL:
        _linha_credito(ws3, r, nome, cred[nome], 1); r += 1
    _linha_credito(ws3, r, 'RESERVA DE CONTINGÊNCIA (X)', cred['Reserva de Contingência'], 0, bold=True, bg=F_GRAY); r += 1
    _linha_credito(ws3, r, 'TOTAL', cred['TOTAL'], 0, bold=True, bg=F_GRAY); r += 1

    # ── ABA 4: RESTOS A PAGAR ────────────────────────────────────────────
    ws4 = wb.create_sheet('Restos a Pagar')
    ws4.sheet_view.showGridLines = False
    ws4.page_setup.orientation = 'landscape'
    for col, w in zip('ABCDEF', [38, 18, 18, 18, 18, 18]):
        ws4.column_dimensions[col].width = w
    _cabecalho_pagina(ws4, 'Execução de Restos a Pagar', mes, ano, ug_label, 6)

    H = 4; ws4.row_dimensions[H].height = 14
    cel(ws4, H, 1, 'NÃO PROCESSADOS', bold=True, bg=F_GRAY, brd=B_THIN)
    H += 1; ws4.row_dimensions[H].height = 28
    for c_idx, txt in enumerate(['DESPESAS ORÇAMENTÁRIAS','Insc.Exerc.Anteriores (a)',
                                  'Em 31/Dez Exerc.Anterior (b)','Liquidados (c)',
                                  'Pagos (d)','Cancelados (e) / Saldo (f)=(a+b-d-e)'], 1):
        cel(ws4, H, c_idx, txt, bold=True, bg=F_GRAY, brd=B_THIN, ha='center')

    def _linha_rpnp(ws, r, nome, dic, nivel, bold=False, bg=None):
        ws.row_dimensions[r].height = 14
        cel(ws, r, 1, nome, bold=bold, bg=bg, brd=B_THIN, ha='left', ind=nivel)
        cel(ws, r, 2, dic.get('inscritos_ant',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 3, dic.get('em_dez_ant',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 4, dic.get('liquidados',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 5, dic.get('pagos',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 6, dic.get('cancelados',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)

    r = H + 1
    _linha_rpnp(ws4, r, 'DESPESAS CORRENTES', rpnp['DESPESAS CORRENTES'], 0, bold=True, bg=F_GRAY); r += 1
    for nome, _gnd, cat in RP_GRUPOS:
        if cat == 'corrente':
            _linha_rpnp(ws4, r, nome, rpnp[nome], 1); r += 1
    _linha_rpnp(ws4, r, 'DESPESAS DE CAPITAL', rpnp['DESPESAS DE CAPITAL'], 0, bold=True, bg=F_GRAY); r += 1
    for nome, _gnd, cat in RP_GRUPOS:
        if cat == 'capital':
            _linha_rpnp(ws4, r, nome, rpnp[nome], 1); r += 1
    _linha_rpnp(ws4, r, 'TOTAL', rpnp['TOTAL'], 0, bold=True, bg=F_GRAY); r += 1

    r += 2
    cel(ws4, r, 1, 'PROCESSADOS', bold=True, bg=F_GRAY, brd=B_THIN); r += 1
    ws4.row_dimensions[r].height = 28
    for c_idx, txt in enumerate(['DESPESAS ORÇAMENTÁRIAS','Insc.Exerc.Anteriores (a)',
                                  'Em 31/Dez Exerc.Anterior (b)','Pagos (c)',
                                  'Cancelados (d) / Saldo (e)=(a+b-c-d)',''], 1):
        cel(ws4, r, c_idx, txt, bold=True, bg=F_GRAY, brd=B_THIN, ha='center')
    r += 1

    def _linha_rpp(ws, r, nome, dic, nivel, bold=False, bg=None):
        ws.row_dimensions[r].height = 14
        cel(ws, r, 1, nome, bold=bold, bg=bg, brd=B_THIN, ha='left', ind=nivel)
        cel(ws, r, 2, dic.get('inscritos_ant',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 3, dic.get('em_dez_ant',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 4, dic.get('pagos',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 5, dic.get('cancelados',0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)

    _linha_rpp(ws4, r, 'DESPESAS CORRENTES', rpp['DESPESAS CORRENTES'], 0, bold=True, bg=F_GRAY); r += 1
    for nome, _gnd, cat in RP_GRUPOS:
        if cat == 'corrente':
            _linha_rpp(ws4, r, nome, rpp[nome], 1); r += 1
    _linha_rpp(ws4, r, 'DESPESAS DE CAPITAL', rpp['DESPESAS DE CAPITAL'], 0, bold=True, bg=F_GRAY); r += 1
    for nome, _gnd, cat in RP_GRUPOS:
        if cat == 'capital':
            _linha_rpp(ws4, r, nome, rpp[nome], 1); r += 1
    _linha_rpp(ws4, r, 'TOTAL', rpp['TOTAL'], 0, bold=True, bg=F_GRAY); r += 1

    # ── ABA 5: AUDITORIA ─────────────────────────────────────────────────
    if achados:
        ws5 = wb.create_sheet('Auditoria')
        ws5.column_dimensions['A'].width = 10
        ws5.column_dimensions['B'].width = 40
        ws5.column_dimensions['C'].width = 80
        cel(ws5, 1, 1, 'Status', bold=True, bg=F_GRAY, brd=B_THIN)
        cel(ws5, 1, 2, 'Controle', bold=True, bg=F_GRAY, brd=B_THIN)
        cel(ws5, 1, 3, 'Detalhe', bold=True, bg=F_GRAY, brd=B_THIN)
        for i, (status, titulo, detalhe) in enumerate(achados, 2):
            bg = F_YELL if status == 'ERRO' else F_WHIT
            cel(ws5, i, 1, status, bold=True, bg=bg, brd=B_THIN)
            cel(ws5, i, 2, titulo, bg=bg, brd=B_THIN)
            cel(ws5, i, 3, detalhe, bg=bg, brd=B_THIN)

    wb.save(output_path)
    print(f"  Excel salvo: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  PDF  —  mesmo padrao estetico do DFC/Balanco Financeiro (reportlab),
#  em paisagem (landscape) por ter mais colunas que os outros relatorios.
# ─────────────────────────────────────────────────────────────────────────────
def _brl(v):
    if v is None:
        return ""
    neg = v < 0
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"({s})" if neg else s


def gerar_pdf(t, mes, ano, ug_label, output_path, achados=None):
    if not REPORTLAB_OK:
        print("  PDF ignorado: instale reportlab -> pip install reportlab")
        return

    doc = SimpleDocTemplate(str(output_path), pagesize=landscape(A4),
          leftMargin=10*mm, rightMargin=10*mm,
          topMargin=10*mm, bottomMargin=10*mm,
          title=f"Balanco Orcamentario GDF {ano} - {MESES[mes]}",
          author="SIAC/SIGGO - GDF")

    AZUL_RL  = colors.HexColor("#2E5C8A")
    VMED_RL  = colors.HexColor("#4472C4")
    VCLR_RL  = colors.HexColor("#D9E2F3")
    CINZA_RL = colors.HexColor("#EBF3FB")
    AMAR_RL  = colors.HexColor("#9DC3E6")

    def st(name, **kw):
        kw.setdefault('fontName', 'Helvetica')
        return ParagraphStyle(name, **kw)

    st_tit  = st("tit", fontName='Helvetica-Bold', fontSize=12,
                 textColor=colors.white, alignment=TA_CENTER)
    st_sub  = st("sub", fontName='Helvetica-Bold', fontSize=9,
                 textColor=colors.white, alignment=TA_CENTER)
    st_meta = st("meta", fontSize=7, textColor=colors.HexColor("#444444"),
                 alignment=TA_CENTER)
    st_hdr  = st("hdr", fontName='Helvetica-Bold', fontSize=7,
                 textColor=colors.white, alignment=TA_CENTER)

    story = []

    def hdr_table(title, sub, meta):
        w = 270*mm
        data = [[Paragraph(title, st_tit)], [Paragraph(sub, st_sub)],
                [Paragraph(meta, st_meta)]]
        tb = Table(data, colWidths=[w])
        tb.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(0,0), AZUL_RL),
            ('BACKGROUND', (0,1),(0,1), VMED_RL),
            ('BACKGROUND', (0,2),(0,2), colors.HexColor("#F5F5F5")),
            ('TOPPADDING',(0,0),(-1,-1), 3),
            ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ]))
        return tb

    def fmt_table(headers, rows_data, col_widths, bold_rows=None, gray_rows=None):
        """rows_data: lista de listas de strings (ja formatadas).
        bold_rows/gray_rows: sets de indices (0-based, sem contar header)."""
        bold_rows = bold_rows or set()
        gray_rows = gray_rows or set()
        header_cells = [Paragraph(f"<b>{h}</b>", st_hdr) for h in headers]
        data = [header_cells]
        for nome_row, vals in rows_data:
            data.append([nome_row] + vals)
        tab = Table(data, colWidths=col_widths, repeatRows=1)
        ts = [
            ('BACKGROUND', (0,0),(-1,0), AZUL_RL),
            ('GRID', (0,0),(-1,-1), 0.3, colors.HexColor("#CCCCCC")),
            ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
            ('TOPPADDING',(0,0),(-1,-1), 2),
            ('BOTTOMPADDING',(0,0),(-1,-1), 2),
            ('LEFTPADDING',(0,0),(-1,-1), 3),
            ('RIGHTPADDING',(0,0),(-1,-1), 3),
        ]
        for i in range(len(rows_data)):
            r = i + 1
            if i in bold_rows:
                ts += [('BACKGROUND',(0,r),(-1,r), VCLR_RL),
                       ('FONTNAME',(0,r),(-1,r), 'Helvetica-Bold')]
            elif i in gray_rows:
                ts.append(('BACKGROUND',(0,r),(-1,r), CINZA_RL))
        tab.setStyle(TableStyle(ts))
        return tab

    def p(txt, bold=False, align=TA_LEFT, ind=0):
        fn = 'Helvetica-Bold' if bold else 'Helvetica'
        return Paragraph(txt, ParagraphStyle("c", fontName=fn, fontSize=7,
                                              alignment=align, leftIndent=ind))

    # ── PAGINA 1: RECEITAS ──────────────────────────────────────────────
    rec = t['receitas']
    story.append(hdr_table("Governo do Distrito Federal",
        "Balanco Orcamentario — Quadro Principal — Receitas (Anexo 12)",
        f"Exercicio {ano} | Mes: {mes:02d}-{MESES[mes]} | {ug_label} | "
        f"Posicao: {datetime.now():%d/%m/%Y %H:%M}"))
    story.append(Spacer(1, 3*mm))

    headers_rec = ["RECEITAS ORÇAMENTÁRIAS", "Previsão Inicial (a)",
                   "Previsão Atualizada (b)", "Receitas Realizadas (c)", "Saldo (d)=(c-b)"]
    rows_rec, bold_rec = [], set()
    idx = 0
    for nome, nivel, tipo, _prefixos in RECEITAS:
        dic = rec.get(nome, {'inicial':0,'atualizada':0,'realizada':0,'saldo':0})
        bold = tipo == "subtotal_calc" or nivel == 0
        if bold:
            bold_rec.add(idx)
        rows_rec.append((p(nome, bold=bold, ind=nivel*8), [
            p(_brl(dic.get('inicial',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('atualizada',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('realizada',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('saldo',0)), bold=bold, align=TA_RIGHT)]))
        idx += 1
    for nome in ('SUBTOTAL DAS RECEITAS (III)', 'OPERAÇÕES DE CRÉDITO / REFINANCIAMENTO (IV)',
                 'SUBTOTAL COM REFINANCIAMENTO (V)', 'DEFICIT (VI)', 'TOTAL (VII)'):
        dic = rec.get(nome, {'inicial':0,'atualizada':0,'realizada':0,'saldo':0})
        bold_rec.add(idx)
        rows_rec.append((p(nome, bold=True), [
            p(_brl(dic.get('inicial',0)), bold=True, align=TA_RIGHT),
            p(_brl(dic.get('atualizada',0)), bold=True, align=TA_RIGHT),
            p(_brl(dic.get('realizada',0)), bold=True, align=TA_RIGHT),
            p(_brl(dic.get('saldo',0)), bold=True, align=TA_RIGHT)]))
        idx += 1

    # SALDOS DE EXERCÍCIOS ANTERIORES + sublinhas (Recursos Arrecadados /
    # Superávit Financeiro / Reabertura de Créditos Adicionais). Antes esse
    # bloco so era escrito no Excel (gerar_excel); faltava aqui no PDF.
    saldos_tot = rec.get('SALDOS DE EXERCÍCIOS ANTERIORES',
                          {'inicial':0,'atualizada':0,'realizada':0})
    bold_rec.add(idx)
    rows_rec.append((p('SALDOS DE EXERCÍCIOS ANTERIORES', bold=True), [
        p(_brl(saldos_tot.get('inicial',0)), bold=True, align=TA_RIGHT),
        p(_brl(saldos_tot.get('atualizada',0)), bold=True, align=TA_RIGHT),
        p(_brl(saldos_tot.get('realizada',0)), bold=True, align=TA_RIGHT),
        p('', bold=True, align=TA_RIGHT)]))
    idx += 1
    saldos_ant_pdf = rec.get('_saldos_exercicios_anteriores', {})
    for label, key in [('  Recursos Arrecadados em Exercícios Anteriores','recursos_arrecadados'),
                        ('  Superávit Financeiro','superavit_financeiro'),
                        ('  Reabertura de Créditos Adicionais','reabertura_creditos')]:
        d = saldos_ant_pdf.get(key, {'atualizada':0,'realizada':0})
        rows_rec.append((p(label, ind=8), [
            p('', align=TA_RIGHT),
            p(_brl(d.get('atualizada',0)), align=TA_RIGHT),
            p(_brl(d.get('realizada',0)), align=TA_RIGHT),
            p('', align=TA_RIGHT)]))
        idx += 1

    col_w_rec = [110*mm, 40*mm, 40*mm, 40*mm, 40*mm]
    story.append(fmt_table(headers_rec, rows_rec, col_w_rec, bold_rows=bold_rec))
    story.append(PageBreak())

    # ── PAGINA 2: DESPESAS ───────────────────────────────────────────────
    des = t['despesas']
    story.append(hdr_table("Governo do Distrito Federal",
        "Balanco Orcamentario — Quadro Principal — Despesas (Anexo 12)",
        f"Exercicio {ano} | Mes: {mes:02d}-{MESES[mes]} | {ug_label} | "
        f"Posicao: {datetime.now():%d/%m/%Y %H:%M}"))
    story.append(Spacer(1, 3*mm))

    headers_des = ["DESPESAS ORÇAMENTÁRIAS", "Dotação Inicial (e)", "Dotação Atualizada (f)",
                   "Despesas Empenhadas (g)", "Despesas Liquidadas (h)", "Despesas Pagas (i)",
                   "Saldo da Dotação (j)=(f-g)"]
    def linha_des(nome, dic, bold=False, ind=0):
        return (p(nome, bold=bold, ind=ind), [
            p(_brl(dic.get('dotacao_inicial',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('dotacao_atualizada',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('empenhada',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('liquidada',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('paga',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('saldo_dotacao',0)), bold=bold, align=TA_RIGHT)])

    rows_des, bold_des = [], set()
    idx = 0
    rows_des.append(linha_des('DESPESAS CORRENTES (VIII)', des['DESPESAS CORRENTES (VIII)'], bold=True)); bold_des.add(idx); idx+=1
    for nome, _ in DESPESAS_CORRENTES:
        rows_des.append(linha_des(nome, des[nome], ind=8)); idx+=1
    rows_des.append(linha_des('DESPESAS DE CAPITAL (IX)', des['DESPESAS DE CAPITAL (IX)'], bold=True)); bold_des.add(idx); idx+=1
    for nome, _ in DESPESAS_CAPITAL:
        rows_des.append(linha_des(nome, des[nome], ind=8)); idx+=1
    rows_des.append(linha_des('RESERVA DE CONTINGÊNCIA (X)', des['Reserva de Contingência'], bold=True)); bold_des.add(idx); idx+=1
    rows_des.append(linha_des('RESERVA DO RPPS', des['Reserva do RPPS'], bold=True)); bold_des.add(idx); idx+=1
    for nome in ('SUBTOTAL DAS DESPESAS (XI)', 'AMORTIZAÇÃO DA DÍVIDA/REFINANCIAMENTO (XII)',
                 'SUBTOTAL COM REFINANCIAMENTO (XIII)', 'SUPERAVIT (XIV)', 'TOTAL (XV)'):
        rows_des.append(linha_des(nome, des[nome], bold=True)); bold_des.add(idx); idx+=1
    col_w_des = [70*mm, 35*mm, 35*mm, 35*mm, 35*mm, 30*mm, 35*mm]
    story.append(fmt_table(headers_des, rows_des, col_w_des, bold_rows=bold_des))

    eq = (rec['TOTAL (VII)']['atualizada'] + rec['SALDOS DE EXERCÍCIOS ANTERIORES']['atualizada']
          - des['SUBTOTAL COM REFINANCIAMENTO (XIII)']['dotacao_atualizada'])
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        f"<b>Equilíbrio:</b> TOTAL(VII)+Saldos Exerc.Ant. [Receitas, Atualizada] "
        f"vs Dotação Atualizada (Despesas): {_brl(eq)}",
        ParagraphStyle("conf", fontName='Helvetica', fontSize=7,
                       textColor=colors.HexColor("#555555"))))
    story.append(PageBreak())

    # ── PAGINA 3: CREDITOS ADICIONAIS ────────────────────────────────────
    cred = t['creditos_adicionais']
    story.append(hdr_table("Governo do Distrito Federal",
        "Balanco Orcamentario — Quadro por Tipos de Créditos Adicionais (Anexo 12)",
        f"Exercicio {ano} | Mes: {mes:02d}-{MESES[mes]} | {ug_label} | "
        f"Posicao: {datetime.now():%d/%m/%Y %H:%M}"))
    story.append(Spacer(1, 3*mm))

    headers_cred = ["DESPESAS ORÇAMENTÁRIAS", "Créd.Suplementar (a)", "Créd.Esp.Abertos (b)",
                     "Créd.Esp.Reabertos (c)", "Créd.Extraord.Reab. (d)",
                     "(-) Cancel.Suplem. (e)", "(-) Remanej.Veto (f)", "Total Alt. (h)"]
    def linha_cred(nome, dic, bold=False, ind=0):
        return (p(nome, bold=bold, ind=ind), [
            p(_brl(dic.get('suplementar',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('esp_abertos',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('esp_reabertos',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('extraord_reabertos',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('cancel_suplementar',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('remanej_veto',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('total_alteracoes',0)), bold=bold, align=TA_RIGHT)])

    rows_cred, bold_cred = [], set()
    idx = 0
    rows_cred.append(linha_cred('DESPESAS CORRENTES', cred['DESPESAS CORRENTES'], bold=True)); bold_cred.add(idx); idx+=1
    for nome, _ in DESPESAS_CORRENTES:
        rows_cred.append(linha_cred(nome, cred[nome], ind=8)); idx+=1
    rows_cred.append(linha_cred('DESPESAS DE CAPITAL', cred['DESPESAS DE CAPITAL'], bold=True)); bold_cred.add(idx); idx+=1
    for nome, _ in DESPESAS_CAPITAL:
        rows_cred.append(linha_cred(nome, cred[nome], ind=8)); idx+=1
    rows_cred.append(linha_cred('RESERVA DE CONTINGÊNCIA (X)', cred['Reserva de Contingência'], bold=True)); bold_cred.add(idx); idx+=1
    rows_cred.append(linha_cred('TOTAL', cred['TOTAL'], bold=True)); bold_cred.add(idx); idx+=1
    col_w_cred = [55*mm, 30*mm, 30*mm, 30*mm, 32*mm, 30*mm, 30*mm, 30*mm]
    story.append(fmt_table(headers_cred, rows_cred, col_w_cred, bold_rows=bold_cred))
    story.append(PageBreak())

    # ── PAGINA 4: RESTOS A PAGAR ──────────────────────────────────────────
    rpnp = t['rp_nao_processados']; rpp = t['rp_processados']
    story.append(hdr_table("Governo do Distrito Federal",
        "Balanco Orcamentario — Execução de Restos a Pagar (Anexo 12)",
        f"Exercicio {ano} | Mes: {mes:02d}-{MESES[mes]} | {ug_label} | "
        f"Posicao: {datetime.now():%d/%m/%Y %H:%M}"))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph("<b>NÃO PROCESSADOS</b>", st("h", fontName='Helvetica-Bold', fontSize=9)))
    story.append(Spacer(1, 1*mm))

    headers_rpnp = ["DESPESAS ORÇAMENTÁRIAS", "Insc.Exerc.Ant. (a)", "Em 31/Dez Ex.Ant. (b)",
                    "Liquidados (c)", "Pagos (d)", "Cancelados (e) / Saldo (f)"]
    def linha_rpnp(nome, dic, bold=False, ind=0):
        return (p(nome, bold=bold, ind=ind), [
            p(_brl(dic.get('inscritos_ant',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('em_dez_ant',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('liquidados',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('pagos',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('cancelados',0)), bold=bold, align=TA_RIGHT)])

    rows_rpnp, bold_rpnp = [], set()
    idx = 0
    rows_rpnp.append(linha_rpnp('DESPESAS CORRENTES', rpnp['DESPESAS CORRENTES'], bold=True)); bold_rpnp.add(idx); idx+=1
    for nome, _, cat in RP_GRUPOS:
        if cat == 'corrente':
            rows_rpnp.append(linha_rpnp(nome, rpnp[nome], ind=8)); idx+=1
    rows_rpnp.append(linha_rpnp('DESPESAS DE CAPITAL', rpnp['DESPESAS DE CAPITAL'], bold=True)); bold_rpnp.add(idx); idx+=1
    for nome, _, cat in RP_GRUPOS:
        if cat == 'capital':
            rows_rpnp.append(linha_rpnp(nome, rpnp[nome], ind=8)); idx+=1
    rows_rpnp.append(linha_rpnp('TOTAL', rpnp['TOTAL'], bold=True)); bold_rpnp.add(idx); idx+=1
    col_w_rp = [60*mm, 42*mm, 42*mm, 42*mm, 42*mm, 42*mm]
    story.append(fmt_table(headers_rpnp, rows_rpnp, col_w_rp, bold_rows=bold_rpnp))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("<b>PROCESSADOS</b>", st("h2", fontName='Helvetica-Bold', fontSize=9)))
    story.append(Spacer(1, 1*mm))
    headers_rpp = ["DESPESAS ORÇAMENTÁRIAS", "Insc.Exerc.Ant. (a)", "Em 31/Dez Ex.Ant. (b)",
                   "Pagos (c)", "Cancelados (d) / Saldo (e)"]
    def linha_rpp(nome, dic, bold=False, ind=0):
        return (p(nome, bold=bold, ind=ind), [
            p(_brl(dic.get('inscritos_ant',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('em_dez_ant',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('pagos',0)), bold=bold, align=TA_RIGHT),
            p(_brl(dic.get('cancelados',0)), bold=bold, align=TA_RIGHT)])

    rows_rpp, bold_rpp = [], set()
    idx = 0
    rows_rpp.append(linha_rpp('DESPESAS CORRENTES', rpp['DESPESAS CORRENTES'], bold=True)); bold_rpp.add(idx); idx+=1
    for nome, _, cat in RP_GRUPOS:
        if cat == 'corrente':
            rows_rpp.append(linha_rpp(nome, rpp[nome], ind=8)); idx+=1
    rows_rpp.append(linha_rpp('DESPESAS DE CAPITAL', rpp['DESPESAS DE CAPITAL'], bold=True)); bold_rpp.add(idx); idx+=1
    for nome, _, cat in RP_GRUPOS:
        if cat == 'capital':
            rows_rpp.append(linha_rpp(nome, rpp[nome], ind=8)); idx+=1
    rows_rpp.append(linha_rpp('TOTAL', rpp['TOTAL'], bold=True)); bold_rpp.add(idx); idx+=1
    col_w_rpp = [70*mm, 50*mm, 50*mm, 50*mm, 50*mm]
    story.append(fmt_table(headers_rpp, rows_rpp, col_w_rpp, bold_rows=bold_rpp))

    # ── PAGINA 5: AUDITORIA ───────────────────────────────────────────────
    if achados:
        story.append(PageBreak())
        story.append(hdr_table("Governo do Distrito Federal",
            "Balanco Orcamentario — Auditoria de Integridade",
            f"Exercicio {ano} | Mes: {mes:02d}-{MESES[mes]} | {ug_label} | "
            f"Posicao: {datetime.now():%d/%m/%Y %H:%M}"))
        story.append(Spacer(1, 3*mm))
        headers_aud = ["Status", "Controle", "Detalhe"]
        rows_aud = []
        for status, titulo, detalhe in achados:
            rows_aud.append((p(status, bold=True), [p(titulo, bold=True), p(detalhe)]))
        col_w_aud = [25*mm, 90*mm, 155*mm]
        tab_aud = fmt_table(headers_aud, rows_aud, col_w_aud)
        # destaca linhas de ERRO em vermelho claro (mesma semantica usada
        # em todos os outros anexos: vermelho = erro, nao confundir com a
        # paleta azul de identidade visual do relatorio)
        VERM_RL = colors.HexColor("#FAD7DA")
        ts_extra = []
        for i, (status, _, _) in enumerate(achados):
            if status == 'ERRO':
                ts_extra.append(('BACKGROUND', (0, i+1), (-1, i+1), VERM_RL))
        if ts_extra:
            tab_aud.setStyle(TableStyle(ts_extra))
        story.append(tab_aud)

    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        f"Emitido por: Python/Oracle  |  Emitido em: {datetime.now():%d/%m/%Y %H:%M:%S}",
        ParagraphStyle("rod", fontName='Helvetica', fontSize=7,
                       textColor=colors.HexColor("#888888"), alignment=TA_RIGHT)))

    doc.build(story)
    print(f"  PDF salvo: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────
def diagnostico(conn, ano):
    print("\n" + "="*64)
    print("  DIAGNOSTICO - ESTRUTURA DO BANCO PARA BALANCO ORCAMENTARIO")
    print("="*64)

    descobrir_anos_notaempenho(conn, ano)

    print(f"\n[0] Verificando MIL{ano}.BALANCOGERAL (saldo anterior de Restos a "
          f"Pagar 531/532) -- existencia, colunas e amostra de INCATEGORIA "
          f"para as contas 531100000/531200000/532100000/532200000:")
    try:
        q = f"""SELECT COCONTACONTABIL, INCATEGORIA, COUNT(*) QTD,
                    SUM(VADEBITO - VACREDITO) SALDO
                FROM MIL{ano}.BALANCOGERAL
                WHERE COCONTACONTABIL IN (531100000,531200000,532100000,532200000)
                GROUP BY COCONTACONTABIL, INCATEGORIA ORDER BY 1,2"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM registro em MIL{ano}.BALANCOGERAL para essas contas.")
        for _, r in df.iterrows():
            print(f"    Conta={int(r['COCONTACONTABIL']):>10}  INCATEGORIA={r['INCATEGORIA']}  "
                  f"qtd={int(r['QTD']):>6}  saldo={float(r['SALDO'] or 0):>20,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[1] Contas de receita (521 - faixa de contas-filha 521100000-"
          f"521299999) por INMES em MIL{ano}:")
    try:
        q = f"""SELECT o.INMES, COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)) SALDO
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL BETWEEN 521100000 AND 521299999
                GROUP BY o.INMES ORDER BY o.INMES"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    INMES={int(r['INMES']):>3}  qtd={int(r['QTD']):>9}  saldo={float(r['SALDO'] or 0):>22,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2] Contas de dotacao/execucao da despesa (522/622) em MIL{ano}:")
    try:
        q = f"""SELECT o.COCONTACONTABIL, COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)) SALDO
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL IN (522110000,522120000,522150000,522190000,
                                            622130000,622130300,622130400,622130700,622920104)
                GROUP BY o.COCONTACONTABIL ORDER BY 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    Conta={int(r['COCONTACONTABIL']):>10}  qtd={int(r['QTD']):>8}  saldo={float(r['SALDO'] or 0):>22,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2b] TODAS as contas na faixa 622130000-622139999 em MIL{ano} (granularidade real):")
    try:
        q = f"""SELECT o.COCONTACONTABIL, COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)) SALDO_SC
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL BETWEEN 622130000 AND 622139999
                  AND o.INMES BETWEEN 1 AND 12
                GROUP BY o.COCONTACONTABIL ORDER BY 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print("    NENHUM registro encontrado nessa faixa de conta.")
        for _, r in df.iterrows():
            print(f"    Conta={int(r['COCONTACONTABIL']):>10}  qtd={int(r['QTD']):>8}  saldo_SC={float(r['SALDO_SC'] or 0):>22,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2c] Verificando se o JOIN com NOTAEMPENHO esta casando "
          f"(sem filtro de GND) para a conta 622130000-622139999 em MIL{ano}:")
    try:
        q = f"""SELECT COUNT(*) QTD_TOTAL,
                    SUM(CASE WHEN EXISTS (
                        SELECT 1 FROM {ne_union(ano)} ne
                        WHERE ne.NUNE = SUBSTR(o.COCONTACORRENTE,1,11)
                          AND ne.COUG = o.COUGCONTAB AND ne.COGESTAO = o.COGESTAOCONTAB
                    ) THEN 1 ELSE 0 END) QTD_COM_MATCH
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL BETWEEN 622130000 AND 622139999
                  AND o.INMES BETWEEN 1 AND 12"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    Total de lancamentos: {int(r['QTD_TOTAL'] or 0)}  |  "
                  f"Com match na NOTAEMPENHO: {int(r['QTD_COM_MATCH'] or 0)}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2d] Amostra de COCONTACORRENTE/COUGCONTAB/COGESTAOCONTAB na conta "
          f"622130000-622139999 (10 linhas) em MIL{ano}:")
    try:
        q = f"""SELECT o.COCONTACONTABIL, o.COCONTACORRENTE, o.COUGCONTAB, o.COGESTAOCONTAB,
                    SUBSTR(o.COCONTACORRENTE,1,11) AS NUNE_DERIVADO
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL BETWEEN 622130000 AND 622139999
                  AND o.INMES BETWEEN 1 AND 12
                  AND ROWNUM <= 10"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    conta={r['COCONTACONTABIL']}  ccorrente={r['COCONTACORRENTE']}  "
                  f"ugcontab={r['COUGCONTAB']}  gestaocontab={r['COGESTAOCONTAB']}  "
                  f"nune_derivado={r['NUNE_DERIVADO']}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[3] Distribuicao de GND (2o digito de CONATUREZA) em MIL{ano}.NOTAEMPENHO:")
    try:
        q = f"""SELECT SUBSTR(CONATUREZA,2,1) GND, COUNT(*) QTD
                FROM MIL{ano}.NOTAEMPENHO GROUP BY SUBSTR(CONATUREZA,2,1) ORDER BY 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    GND={r['GND']}  qtd={int(r['QTD']):>8}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2e] Distribuicao por INMES das contas 622130000-622139999 e 622920104 "
          f"em MIL{ano} (para checar se o saldo esta em outro INMES, ex. 13/15):")
    try:
        q = f"""SELECT o.COCONTACONTABIL, o.INMES, COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)) SALDO_SC
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL BETWEEN 622130000 AND 622139999
                   OR o.COCONTACONTABIL = 622920104
                GROUP BY o.COCONTACONTABIL, o.INMES ORDER BY 1,2"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    Conta={int(r['COCONTACONTABIL']):>10}  INMES={int(r['INMES']):>3}  "
                  f"qtd={int(r['QTD']):>8}  saldo_SC={float(r['SALDO_SC'] or 0):>20,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2f] Colunas completas de MIL{ano}.LANCAMENTOCONTABIL (procurando "
          f"campos de GND/Funcao/Subfuncao/Natureza ja prontos na propria tabela, "
          f"sem precisar decompor COCONTACORRENTE):")
    try:
        q = f"""SELECT column_name, data_type FROM all_tab_columns
                WHERE owner='MIL{ano}' AND table_name='LANCAMENTOCONTABIL'
                ORDER BY column_id"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    {r['COLUMN_NAME']:<30} {r['DATA_TYPE']}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2g] Procurando tabelas com 'CELULA' ou 'NATUREZA' no nome em MIL{ano} "
          f"(possivel tabela de celula orcamentaria para decodificar COCONTACORRENTE):")
    try:
        q = f"""SELECT table_name FROM all_tables
                WHERE owner='MIL{ano}'
                  AND (table_name LIKE '%CELULA%' OR table_name LIKE '%NATUREZA%'
                       OR table_name LIKE '%CONTACORRENTE%' OR table_name LIKE '%EMPENHO%')
                ORDER BY table_name"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    {r['TABLE_NAME']}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2h] Colunas de MIL{ano}.NOTAEMPENHO (para localizar campo de "
          f"GND/CONATUREZA e eventual campo que ligue ao lancamento contabil "
          f"de credito empenhado/liquidado, fora do NUNE classico):")
    try:
        q = f"""SELECT column_name, data_type FROM all_tab_columns
                WHERE owner='MIL{ano}' AND table_name='NOTAEMPENHO'
                ORDER BY column_id"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    {r['COLUMN_NAME']:<30} {r['DATA_TYPE']}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2i] Amostra de uma linha completa (todas as colunas) da conta "
          f"622130100 em MIL{ano}, para inspecionar visualmente todos os campos "
          f"disponiveis no lancamento (pode haver coluna de natureza/GND direta):")
    try:
        q = f"""SELECT * FROM MIL{ano}.LANCAMENTOCONTABIL
                WHERE COCONTACONTABIL = 622130100 AND ROWNUM <= 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for col in df.columns:
            print(f"    {col:<30} = {df.iloc[0][col]}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2j] Decompondo a 'Celula da Despesa com ND Detalhado' (tipo 20): "
          f"buscando a posicao exata onde COUGCONTAB e COGESTAOCONTAB aparecem "
          f"dentro da string de 40 caracteres da COCONTACORRENTE, usando uma "
          f"amostra de MIL{ano}, conta 622130100:")
    try:
        q = f"""SELECT COCONTACORRENTE, COUGCONTAB, COGESTAOCONTAB
                FROM MIL{ano}.LANCAMENTOCONTABIL
                WHERE COCONTACONTABIL = 622130100 AND ROWNUM <= 15"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, row in df.iterrows():
            cc = str(row['COCONTACORRENTE']).strip()
            ug = str(row['COUGCONTAB']).strip()
            ges = str(row['COGESTAOCONTAB']).strip()
            # zero-pad UG a 6 digitos e Gestao a 5 (padrao SIAFI/SIGGO)
            ug6 = ug.zfill(6)
            ges5 = ges.zfill(5)
            pos_ug  = cc.find(ug6)
            pos_ges = cc.find(ges5)
            print(f"    ccorrente={cc}  (len={len(cc)})")
            print(f"      ugcontab={ug} (zfill6={ug6})  pos_na_string={pos_ug}")
            print(f"      gestaocontab={ges} (zfill5={ges5})  pos_na_string={pos_ges}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2k] Cruzando COCONTACORRENTE da conta 622130100 com a Funcao/"
          f"Subfuncao/Programa/Acao do orcamento (tabela de detalhamento da "
          f"despesa) via NUNE classico, para tentar achar QUALQUER tabela que "
          f"contenha esse mesmo codigo de 40 digitos como chave (procurando "
          f"colunas chamadas COCONTACORRENTE ou similar em outras tabelas):")
    try:
        q = f"""SELECT DISTINCT table_name, column_name FROM all_tab_columns
                WHERE owner='MIL{ano}'
                  AND (column_name LIKE '%CONTACORRENTE%' OR column_name LIKE '%CELULA%')
                ORDER BY table_name"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    {r['TABLE_NAME']:<30} {r['COLUMN_NAME']}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2l] Amostra de COCONTACORRENTE da conta 522120100 (Credito Adicional "
          f"Suplementar, tipo '13 - Celula Orcamentaria da Despesa', SEM Natureza "
          f"da Despesa) em MIL{ano}, para confirmar tamanho/layout (esperado: "
          f"Esfera(1)+UO(5)+ProgTrabalho(17)+Fonte(9) = 32 caracteres):")
    try:
        q = f"""SELECT COCONTACORRENTE, COUGCONTAB, COGESTAOCONTAB
                FROM MIL{ano}.LANCAMENTOCONTABIL
                WHERE COCONTACONTABIL = 522120100 AND ROWNUM <= 10"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print("    Nenhum lancamento encontrado na conta 522120100.")
        for _, row in df.iterrows():
            cc = str(row['COCONTACORRENTE']).strip()
            print(f"    ccorrente={cc}  (len={len(cc)})  ugcontab={row['COUGCONTAB']}  "
                  f"gestaocontab={row['COGESTAOCONTAB']}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2m] A celula tipo '13' (Creditos Adicionais) parece ser apenas "
          f"Esfera+UO+ProgTrabalho+Fonte, SEM a Natureza da Despesa -- ou seja, "
          f"sem GND embutido. Procurando tabela de detalhamento/lei orcamentaria "
          f"que ligue essa celula a uma Natureza da Despesa em MIL{ano} "
          f"(nomes tipicos: LOA, DETALHAMENTODESPESA, CREDITODESPESA, DOTACAO):")
    try:
        q = f"""SELECT table_name FROM all_tables
                WHERE owner='MIL{ano}'
                  AND (table_name LIKE '%DOTACAO%' OR table_name LIKE '%CREDITO%'
                       OR table_name LIKE '%DETALHAMENTO%' OR table_name LIKE '%LOA%'
                       OR table_name LIKE '%PROGRAMATRABALHO%' OR table_name LIKE '%PROGTRABALHO%')
                ORDER BY table_name"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    {r['TABLE_NAME']}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2n] Confirmando tamanho da COCONTACORRENTE nas contas de Restos a "
          f"Pagar (531/532, tipo '16 - Numero do Empenho', esperado 11 "
          f"caracteres = NUNE puro) em MIL{ano}:")
    try:
        q = f"""SELECT COCONTACONTABIL, COCONTACORRENTE, COUGCONTAB, COGESTAOCONTAB
                FROM MIL{ano}.LANCAMENTOCONTABIL
                WHERE COCONTACONTABIL IN (531100000,531200000,532100000,532200000)
                  AND ROWNUM <= 8"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print("    Nenhum lancamento encontrado nessas contas.")
        for _, row in df.iterrows():
            cc = str(row['COCONTACORRENTE']).strip()
            print(f"    conta={row['COCONTACONTABIL']}  ccorrente={cc}  (len={len(cc)})  "
                  f"ugcontab={row['COUGCONTAB']}  gestaocontab={row['COGESTAOCONTAB']}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2o] Distribuicao por INMES das contas de Restos a Pagar (531/532) "
          f"em MIL{ano} -- para ver se o saldo de abertura esta mesmo em INMES=0 "
          f"ou se precisa vir de outro schema/mes:")
    try:
        q = f"""SELECT COCONTACONTABIL, INMES, COUNT(*) QTD,
                    SUM(DECODE(INDEBITOCREDITO,'D',VALANCAMENTO,'C',-VALANCAMENTO,0)) SALDO_SD
                FROM MIL{ano}.LANCAMENTOCONTABIL
                WHERE COCONTACONTABIL IN (531100000,531200000,532100000,532200000)
                GROUP BY COCONTACONTABIL, INMES ORDER BY 1,2"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM lancamento em 531/532 dentro de MIL{ano} (qualquer INMES).")
        for _, r in df.iterrows():
            print(f"    Conta={int(r['COCONTACONTABIL']):>10}  INMES={int(r['INMES']):>3}  "
                  f"qtd={int(r['QTD']):>6}  saldo_SD={float(r['SALDO_SD'] or 0):>20,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2p] Verificando se o JOIN com NOTAEMPENHO casa para as contas "
          f"631/632 (RP Liquidados/Pagos/Cancelados) em MIL{ano} (sem filtro "
          f"de GND, para isolar problema de JOIN vs filtro):")
    try:
        q = f"""SELECT COUNT(*) QTD_TOTAL,
                    SUM(CASE WHEN EXISTS (
                        SELECT 1 FROM {ne_union(ano)} ne
                        WHERE ne.NUNE = SUBSTR(o.COCONTACORRENTE,1,11)
                          AND ne.COUG = o.COUGCONTAB AND ne.COGESTAO = o.COGESTAOCONTAB
                    ) THEN 1 ELSE 0 END) QTD_COM_MATCH
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL IN (631300000,631400000,631810000,631820000,631900000)
                  AND o.INMES BETWEEN 1 AND 12"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    Total: {int(r['QTD_TOTAL'] or 0)}  |  Com match: {int(r['QTD_COM_MATCH'] or 0)}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2q] Mesma verificacao para contas 632 (RP Processados Pagos/"
          f"Cancelados) em MIL{ano}:")
    try:
        q = f"""SELECT COUNT(*) QTD_TOTAL,
                    SUM(CASE WHEN EXISTS (
                        SELECT 1 FROM {ne_union(ano)} ne
                        WHERE ne.NUNE = SUBSTR(o.COCONTACORRENTE,1,11)
                          AND ne.COUG = o.COUGCONTAB AND ne.COGESTAO = o.COGESTAOCONTAB
                    ) THEN 1 ELSE 0 END) QTD_COM_MATCH
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL IN (632210100,632210200,632210300,632210400,632900000)
                  AND o.INMES BETWEEN 1 AND 12"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    Total: {int(r['QTD_TOTAL'] or 0)}  |  Com match: {int(r['QTD_COM_MATCH'] or 0)}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2r] Distribuicao por INMES das contas 631/632 em MIL{ano}, com "
          f"contagem e saldo (para ver se ha algum INMES com dados, mesmo que "
          f"o JOIN nao case):")
    try:
        q = f"""SELECT COCONTACONTABIL, INMES, COUNT(*) QTD,
                    SUM(DECODE(INDEBITOCREDITO,'C',VALANCAMENTO,'D',-VALANCAMENTO,0)) SALDO_SC
                FROM MIL{ano}.LANCAMENTOCONTABIL
                WHERE COCONTACONTABIL IN (631300000,631400000,631810000,631820000,631900000,
                                          632210100,632210200,632210300,632210400,632900000)
                GROUP BY COCONTACONTABIL, INMES ORDER BY 1,2"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM lancamento em 631/632 dentro de MIL{ano}.")
        for _, r in df.iterrows():
            print(f"    Conta={int(r['COCONTACONTABIL']):>10}  INMES={int(r['INMES']):>3}  "
                  f"qtd={int(r['QTD']):>6}  saldo_SC={float(r['SALDO_SC'] or 0):>20,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2s] Foco nas contas 531100000/531200000 (Inscritos/Em 31-Dez) em "
          f"MIL{ano}: testando o JOIN com NOTAEMPENHO (range expandido) e o "
          f"INMES real, SEM filtro de GND, para isolar a causa do zero:")
    try:
        q = f"""SELECT o.COCONTACONTABIL, o.INMES, COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)) SALDO_SD,
                    SUM(CASE WHEN EXISTS (
                        SELECT 1 FROM {ne_union(ano)} ne
                        WHERE ne.NUNE = SUBSTR(o.COCONTACORRENTE,1,11)
                          AND ne.COUG = o.COUGCONTAB AND ne.COGESTAO = o.COGESTAOCONTAB
                    ) THEN 1 ELSE 0 END) QTD_COM_MATCH
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL IN (531100000,531200000)
                GROUP BY o.COCONTACONTABIL, o.INMES ORDER BY 1,2"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM lancamento em 531100000/531200000 dentro de MIL{ano}.")
        for _, r in df.iterrows():
            print(f"    Conta={int(r['COCONTACONTABIL']):>10}  INMES={int(r['INMES']):>3}  "
                  f"qtd={int(r['QTD']):>6}  saldo_SD={float(r['SALDO_SD'] or 0):>18,.2f}  "
                  f"com_match={int(r['QTD_COM_MATCH'] or 0)}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2t] Amostra de 5 linhas completas (todas as colunas) da conta "
          f"531100000 em MIL{ano}, para inspecao visual (NUNE, COCONTACORRENTE, "
          f"INMES, UG/Gestao reais):")
    try:
        q = f"""SELECT * FROM MIL{ano}.LANCAMENTOCONTABIL
                WHERE COCONTACONTABIL = 531100000 AND ROWNUM <= 5"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM lancamento na conta 531100000 em MIL{ano}.")
        for idx, row in df.iterrows():
            print(f"    --- linha {idx} ---")
            for col in df.columns:
                print(f"      {col:<25} = {row[col]}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2u] Investigando Amortizacao de Emprestimos (Receita de Capital): "
          f"buscando lancamentos nas contas 521100000-521299999/621200000/621300000-"
          f"621399999 em MIL{ano} cujo COCONTACORRENTE comece com '230' ou '830' "
          f"(prefixo esperado pela planilha de equacoes), e tambem mostrando TODOS "
          f"os prefixos de 3 digitos distintos encontrados nessas contas, para ver "
          f"se '230'/'830' realmente existe no banco ou se o codigo real e outro:")
    try:
        q = f"""SELECT SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,3) AS PREFIXO3,
                    COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)) SALDO_SD
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL BETWEEN 521100000 AND 521299999
                GROUP BY SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,3)
                ORDER BY 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM lancamento em 521100000-521299999 em MIL{ano}.")
        for _, r in df.iterrows():
            marca = "  <-- esperado p/ Amortizacao" if r['PREFIXO3'] in ('230','830') else ""
            print(f"    prefixo3={r['PREFIXO3']:<6} qtd={int(r['QTD']):>6}  "
                  f"saldo_SD={float(r['SALDO_SD'] or 0):>18,.2f}{marca}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2v] Mesma verificacao para a conta 621200000 (Receita Realizada, "
          f"todas as naturezas) em MIL{ano}, prefixo de 3 digitos, restrito a "
          f"prefixos comecando com '2' ou '8' (categoria de capital esperada "
          f"para amortizacao/operacoes de credito):")
    try:
        q = f"""SELECT SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,3) AS PREFIXO3,
                    COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)) SALDO_SC
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL = 621200000
                  AND (SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,1) IN ('2','8'))
                GROUP BY SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,3)
                ORDER BY 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM lancamento com prefixo 2x/8x em 621200000 em MIL{ano}.")
        for _, r in df.iterrows():
            marca = "  <-- esperado p/ Amortizacao" if r['PREFIXO3'] in ('230','830') else ""
            print(f"    prefixo3={r['PREFIXO3']:<6} qtd={int(r['QTD']):>6}  "
                  f"saldo_SC={float(r['SALDO_SC'] or 0):>18,.2f}{marca}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2w] AUDITORIA da linha 'OPERAÇÕES DE CRÉDITO/REFINANCIAMENTO (IV)':")
    print(f"     verificando se existe NATUREZA DE RECEITA com rubrica '1' "
          f"(padrao MCASP de refinanciamento, ex. 211.12.00, 811.12.00) nas "
          f"contas 521100000-521299999 e 621200000 em MIL{ano}. Se aparecer "
          f"algo aqui, a linha IV NAO deve mais ficar zerada -- ajustar o "
          f"script (separar esse valor de RECEITAS DE CAPITAL (II) para a "
          f"linha IV, sem alterar o Subtotal III+IV).")
    try:
        q = f"""SELECT SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,8) AS NATREC8,
                    COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)) SALDO_SD
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL BETWEEN 521100000 AND 521299999
                  AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,3) IN ('211','212','811','812')
                  AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),4,1) = '1'
                GROUP BY SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,8)
                ORDER BY 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    ✔ OK: nenhuma natureza com rubrica '1' encontrada -- "
                  f"linha IV permanece corretamente zerada.")
        else:
            print(f"    ✘ ATENCAO: encontrada(s) natureza(s) com rubrica '1' "
                  f"(possivel Refinanciamento) -- REVISAR linha IV:")
            for _, r in df.iterrows():
                print(f"      natrec8={r['NATREC8']}  qtd={int(r['QTD'])}  "
                      f"saldo_SD={float(r['SALDO_SD'] or 0):,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print("="*64)
    print("  Envie [2u] a [2w] para validar Amortizacao e a linha IV.")
    print("="*64 + "\n")


def main():
    p = argparse.ArgumentParser(description='Balanco Orcamentario GDF (Anexo 12)')
    p.add_argument('--mes', type=int, default=5, help='Mes de referencia 1-12')
    p.add_argument('--ano', type=int, default=2026)
    p.add_argument('--ug', type=str, default=None)
    p.add_argument('--saida', type=str, default=None)
    p.add_argument('--formato', choices=['ambos','excel','pdf'], default='excel')
    p.add_argument('--diag', action='store_true')
    a = p.parse_args()

    if a.diag:
        print("\nConectando para diagnostico...")
        conn = conectar_oracle()
        diagnostico(conn, a.ano)
        conn.close()
        return

    if not 1 <= a.mes <= 12:
        print("ERRO: --mes deve estar entre 1 e 12."); sys.exit(1)

    ug_label = f'UG: {a.ug}' if a.ug else 'Consolidado'
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome = (f'Balanco_Orcamentario_{a.ano}_{a.mes:02d}'
            + (f'_UG{a.ug}' if a.ug else '_Consolidado')
            + f'_{timestamp}')
    base = (Path(a.saida).with_suffix('') if a.saida else OUTPUT_DIR / nome)
    saida_xlsx = base.with_suffix('.xlsx')
    saida_pdf  = base.with_suffix('.pdf')

    print(f"\n{'='*60}")
    print(f"  BALANCO ORCAMENTARIO - {MESES[a.mes]}/{a.ano} - {ug_label}")
    print(f"  Posicao acumulada de Janeiro ate {MESES[a.mes]}")
    print(f"{'='*60}")

    print("\n[1/4] Conectando ao Oracle...")
    conn = conectar_oracle()
    descobrir_anos_notaempenho(conn, a.ano)

    print("\n[2/4] Buscando dados (Receitas, Despesas, Creditos, Restos a Pagar)...")
    receitas_raw  = buscar_receitas(conn, a.mes, a.ano, a.ug)
    opcred_raw    = buscar_op_credito_refinanciamento(conn, a.mes, a.ano, a.ug)
    saldos_raw    = buscar_saldos_exercicios_anteriores(conn, a.mes, a.ano, a.ug)
    despesas_raw  = buscar_despesas(conn, a.mes, a.ano, a.ug)
    creditos_raw  = buscar_creditos_adicionais(conn, a.mes, a.ano, a.ug)
    rpnp_raw      = buscar_rp_nao_processados(conn, a.mes, a.ano, a.ug)
    rpp_raw       = buscar_rp_processados(conn, a.mes, a.ano, a.ug)
    achado_defasagem = auditoria_defasagem_balancogeral(conn, a.ano, a.mes, a.ug)
    saldo_521920500, saldo_521920500_ant = buscar_saldo_521920500(conn, a.mes, a.ano)
    conn.close()

    print("\n[3/4] Calculando totais derivados...")
    t = calcular_tudo(receitas_raw, opcred_raw, saldos_raw, despesas_raw,
                       creditos_raw, rpnp_raw, rpp_raw)
    achados = auditoria_integridade(t, saldo_521920500=saldo_521920500,
                                     saldo_521920500_ant=saldo_521920500_ant,
                                     mes=a.mes, ano=a.ano)
    achados.append(achado_defasagem)

    print("\n[4/4] Gerando arquivo(s) de saida...")
    if a.formato in ('excel', 'ambos'):
        gerar_excel(t, a.mes, a.ano, ug_label, saida_xlsx, achados=achados)
    if a.formato in ('pdf', 'ambos'):
        gerar_pdf(t, a.mes, a.ano, ug_label, saida_pdf, achados=achados)

    print(f"\n{'─'*60}")
    print("  CONFERENCIA RAPIDA:")
    print(f"{'─'*60}")
    rec = t['receitas']; des = t['despesas']
    print(f"    {'RECEITAS CORRENTES (I)':<35}: {rec['RECEITAS CORRENTES (I)']['realizada']:>20,.2f}")
    print(f"    {'RECEITAS DE CAPITAL (II)':<35}: {rec['RECEITAS DE CAPITAL (II)']['realizada']:>20,.2f}")
    print(f"    {'TOTAL (VII) Realizada':<35}: {rec['TOTAL (VII)']['realizada']:>20,.2f}")
    print(f"    {'DESPESAS CORRENTES (VIII) Empenhada':<35}: {des['DESPESAS CORRENTES (VIII)']['empenhada']:>20,.2f}")
    print(f"    {'DESPESAS DE CAPITAL (IX) Empenhada':<35}: {des['DESPESAS DE CAPITAL (IX)']['empenhada']:>20,.2f}")
    print(f"    {'TOTAL (XV) Empenhada':<35}: {des['TOTAL (XV)']['empenhada']:>20,.2f}")
    print(f"{'─'*60}")

    imprimir_auditoria(achados)
    print(f"\n  Concluido em {datetime.now():%H:%M:%S}\n")


if __name__ == '__main__':
    main()

