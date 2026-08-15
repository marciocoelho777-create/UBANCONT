# -*- coding: utf-8 -*-
"""
=============================================================================
  DEMONSTRACAO DAS MUTACOES DO PATRIMONIO LIQUIDO - GDF  (PSIAG550 / DMPL)
  SQL fixo, equacoes embutidas conforme a "Lista Equacoes de Balanco -
  Tipo de Balanco: 08 - DMPL" (arquivo
  Relatório_Equação_de_Balanço_DMPL.xlsx), seguindo o MESMO padrao de
  extracao Oracle ja validado e em producao nos projetos:
    - dvp.py  (Demonstracao das Variacoes Patrimoniais)
    - dfc.py  (Demonstracao dos Fluxos de Caixa, Anexo 15)
    - bf.py   (Balanco Financeiro, Anexo 13)
    - bo.py   (Balanco Orcamentario, Anexo 12)

  Layout de saida (linhas/colunas) espelha o modelo oficial ListaDMPL
  (Governo do Distrito Federal, "DEMONSTRAÇÃO DAS MUTAÇÕES DO PATRIMÔNIO
  LÍQUIDO", PSIAG550, Versão 1) -- 10 linhas (Especificacao) x 9 colunas
  patrimoniais.

  COLUNAS (Especificacao) E SUAS CONTAS-MAE (conforme a Lista de Equacoes):
    1. Pat. Social / Capital Social ............ 231XXXXXX
    2. Adiantamento p/ Futuro Aum. Capital (AFAC) 232XXXXXX
    3. Reserva de Capital ........................ 233XXXXXX
    4. Ajustes de Avaliação Patrimonial .......... 23411XXXX
    5. Reservas de Lucros ........................ 235XXXXXX
    6. Demais Reservas ........................... 236XXXXXX
    7. Resultados Acumulados ..................... 237XXXXXX (+ 4XXXXXXXX/
       3XXXXXXXX nas linhas Resultado do Exercício e Saldos Finais)
    8. Ações / Cotas em Tesouraria ............... sem conta mapeada (= 0)
    9. TOTAL = soma das colunas 1 a 8 (cada linha)

  LINHAS (Especificacao) E FONTE DE DADOS:
  A Lista de Equacoes rotula "Mês: 0" ou "Mês: 13" para cada linha, mas
  esses códigos NÃO viram filtros literais "INMES = 0" / "INMES = 13" em
  LANCAMENTOCONTABIL -- essa interpretação literal foi testada em produção
  e devolveu ZERO em tudo (LANCAMENTOCONTABIL não carrega o saldo de
  abertura/fechamento das contas de Patrimônio Líquido nesses INMES
  exatos). A leitura correta, por analogia direta com os padrões já
  validados em bf.py (saldo de caixa) e dvp.py (VPA/VPD), é:

    * Linhas de SALDO (Saldos Iniciais / Saldos Finais, Tipo Movimento
      "SC" nas contas de Patrimônio Líquido 231/232/233/235/236/237/
      23411) -> vêm da VIEW VSALDOCONTABIL, exatamente como bf.py busca
      o saldo de caixa (ANT_*/SEG_*):
        Saldos Iniciais  -> MIL{ano}.VSALDOCONTABIL, INMES = 0
        Saldos Finais    -> MIL{ano}.VSALDOCONTABIL, INMES BETWEEN 0 AND :mes
      Fórmula: VACREDITO - VACREDITO (natureza credora do PL).

    * Linhas de MOVIMENTO (Ajustes de Exerc. Anteriores, Aumento de
      Capital, Ajuste de Avaliação Patrimonial, Constituição/Reversão de
      Reservas -- Tipo Movimento "MC"/"MD") e o Resultado do Exercício
      (SC 4XXXXXXXX / SD 3XXXXXXXX) -> vêm de LANCAMENTOCONTABIL com
      INMES BETWEEN 1 AND :mes (acumulado janeiro..mês de referência),
      o MESMO range usado em dvp.py para VPA/VPD e em bf.py/bo.py/dfc.py
      para os itens de movimento -- "Mês: 13" na Lista de Equações é o
      código operacional para "acumulado do exercício até o encerramento
      do período corrente", não um INMES literal.

  Esse desenho garante a identidade Saldos Iniciais + Variações = Saldos
  Finais (cada lado fechando como bf.py faz para caixa: ANT_* + movimento
  do período = SEG_*), e garante que o Resultado do Exercício da DMPL
  bata com o Resultado Patrimonial do Período (III) da DVP (mesma fonte,
  mesmo range de INMES) -- conferido pela validação cruzada X6 no
  mestre.py.

  TIPOS DE MOVIMENTO (coluna "Tipo Movimento" da Lista):
    SC = Saldo Credor  = C - D   (saldo acumulado, natureza credora)
    SD = Saldo Devedor = D - C   (saldo acumulado, natureza devedora)
    MC = Movimento a Crédito = só lançamentos C  (no mês exato da regra)
    MD = Movimento a Débito  = só lançamentos D  (no mês exato da regra)
  (mesma convenção SC/SD/MC/MD usada em bf.py/bo.py/dfc.py/dvp.py)

  IDENTIDADE DE FECHAMENTO (auditoria):
    Saldos Iniciais
    + Ajustes de Exercícios Anteriores
    + Aumento de Capital
    + Resgate/Reemissão Ações e Cotas
    + Juros sobre Capital Próprio
    + Resultado do Exercício
    + Ajuste de Avaliação Patrimonial
    + Constituição/Reversão de Reservas
    + Dividendos a Distribuir
    = Saldos Finais
  (por coluna, e também na coluna TOTAL)

  Dependencias:  pip install oracledb openpyxl pandas reportlab
  Uso:
      python dmpl.py --mes 5 --ano 2026
      python dmpl.py --mes 5 --ano 2026 --ug 130101
      python dmpl.py --mes 5 --ano 2026 --formato pdf
      python dmpl.py --mes 5 --ano 2026 --diag
=============================================================================
"""
import argparse, sys
from datetime import datetime
from pathlib import Path

import oracledb
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIGURACAO  (ajuste conforme ambiente - mesmos valores do BF/BO/DFC/DVP)
# ─────────────────────────────────────────────────────────────────────────────
DB_USER     = "usefp07"
DB_PASSWORD = "mar2c"
DB_HOST     = "10.69.1.118"
DB_PORT     = 1521
DB_SERVICE  = "oraprd06"

INSTANT_CLIENT_DIR = r"C:\balanço 2026 gemini arquivos\instantclient_23_9"
OUTPUT_DIR         = Path(r"C:\balanço 2026 gemini arquivos")

MESES = {1:'Janeiro',2:'Fevereiro',3:'Marco',4:'Abril',5:'Maio',
         6:'Junho',7:'Julho',8:'Agosto',9:'Setembro',
         10:'Outubro',11:'Novembro',12:'Dezembro'}

# ─────────────────────────────────────────────────────────────────────────────
#  COLUNAS PATRIMONIAIS DA DMPL  (codigo_coluna, chave_python, descricao,
#  faixa_de_conta_mae) -- conforme a coluna "Coluna" da Lista de Equacoes.
#  Coluna 8 (Acoes/Cotas em Tesouraria) nao tem conta mapeada no plano de
#  contas do ente -- mantida por completude do layout oficial (sempre 0).
# ─────────────────────────────────────────────────────────────────────────────
COLUNAS = [
    ("1", "PAT_SOCIAL_CAPITAL",   "Pat. Social / Capital Social",                "231XXXXXX"),
    ("2", "AFAC",                 "Adiantamento Para Futuro Aumento de Capital (AFAC)", "232XXXXXX"),
    ("3", "RESERVA_CAPITAL",      "Reserva de Capital",                          "233XXXXXX"),
    ("4", "AJUSTE_AVAL_PATRIM",   "Ajustes de Avaliação Patrimonial",            "23411XXXX"),
    ("5", "RESERVAS_LUCROS",      "Reservas de Lucros",                          "235XXXXXX"),
    ("6", "DEMAIS_RESERVAS",      "Demais Reservas",                             "236XXXXXX"),
    ("7", "RESULTADOS_ACUM",      "Resultados Acumulados",                       "237XXXXXX"),
    ("8", "ACOES_TESOURARIA",     "Ações / Cotas em Tesouraria",                 None),
]
COLUNA_CHAVES = [k for _, k, _, _ in COLUNAS]

# ─────────────────────────────────────────────────────────────────────────────
#  LINHAS (ESPECIFICACAO) DA DMPL  -- fonte: Lista Equacoes de Balanco,
#  Tipo de Balanco 08 - DMPL. Cada linha e uma lista de "componentes":
#  (coluna_chave, mascara_conta_ou_None, inmes_fixo, tipo_mov)
#  tipo_mov: 'SC'=C-D ; 'SD'=D-C ; 'MC'=so C ; 'MD'=so D (sinal negativo
#  ja embutido na regra '-' da Lista, aplicado no proprio componente).
# ─────────────────────────────────────────────────────────────────────────────

def _faixa(mascara):
    """Converte mascara tipo '231XXXXXX' em (ini, fim) numerico, igual ao
    padrao usado em dvp.py."""
    n_x = mascara.count('X')
    prefixo = mascara.replace('X', '')
    ini = int(prefixo) * (10 ** n_x)
    fim = ini + (10 ** n_x) - 1
    return ini, fim


# (codigo_item, nome_item, componentes)
#   componentes = [(coluna_chave, mascara, fonte, tipo_mov, sinal), ...]
#   fonte: 'SALDO_INI' = VSALDOCONTABIL INMES=0 (saldo de abertura)
#          'SALDO_FIM' = VSALDOCONTABIL INMES BETWEEN 0 AND :mes (saldo atual)
#          'LANC'      = LANCAMENTOCONTABIL INMES BETWEEN 1 AND :mes (movimento
#                         acumulado do exercicio atual)
#   tipo_mov: 'SC'=C-D ; 'SD'=D-C ; 'MC'=so C ; 'MD'=so D
#   sinal: +1 ou -1 (coluna "Operação" da Lista)
LINHAS_DMPL = [
    ("1.01.01", "SALDOS_INICIAIS", [
        ("PAT_SOCIAL_CAPITAL", "231XXXXXX", "SALDO_INI", "SC", +1),
        ("AFAC",               "232XXXXXX", "SALDO_INI", "SC", +1),
        ("RESERVA_CAPITAL",    "233XXXXXX", "SALDO_INI", "SC", +1),
        ("RESERVAS_LUCROS",    "235XXXXXX", "SALDO_INI", "SC", +1),
        ("DEMAIS_RESERVAS",    "236XXXXXX", "SALDO_INI", "SC", +1),
        ("RESULTADOS_ACUM",    "237XXXXXX", "SALDO_INI", "SC", +1),
        ("AJUSTE_AVAL_PATRIM", "23411XXXX", "SALDO_INI", "SC", +1),
    ]),
    ("1.01.02", "AJUSTES_EXERC_ANTERIORES", [
        ("RESULTADOS_ACUM",    "237XXXXXX", "LANC", "MC", +1),
        ("RESULTADOS_ACUM",    "237XXXXXX", "LANC", "MD", -1),
    ]),
    ("1.01.03", "AUMENTO_CAPITAL", [
        ("PAT_SOCIAL_CAPITAL", "231XXXXXX", "LANC", "MC", +1),
        ("PAT_SOCIAL_CAPITAL", "231XXXXXX", "LANC", "MD", -1),
        ("AFAC",               "232XXXXXX", "LANC", "MC", +1),
        ("AFAC",               "232XXXXXX", "LANC", "MD", -1),
    ]),
    ("1.01.04", "RESGATE_REEMISSAO_ACOES", [
        # Sem conta mapeada na Lista de Equacoes -- sempre 0.
    ]),
    ("1.01.05", "JUROS_CAPITAL_PROPRIO", [
        # Sem conta mapeada na Lista de Equacoes -- sempre 0.
    ]),
    ("1.01.06", "RESULTADO_EXERCICIO", [
        ("RESULTADOS_ACUM",    "4XXXXXXXX", "LANC", "SC", +1),
        ("RESULTADOS_ACUM",    "3XXXXXXXX", "LANC", "SD", -1),
    ]),
    ("1.01.07", "AJUSTE_AVALIACAO_PATRIMONIAL", [
        ("AJUSTE_AVAL_PATRIM", "23411XXXX", "LANC", "MC", +1),
        ("AJUSTE_AVAL_PATRIM", "23411XXXX", "LANC", "MD", -1),
    ]),
    ("1.01.08", "CONSTITUICAO_REVERSAO_RESERVAS", [
        ("RESERVA_CAPITAL",    "233XXXXXX", "LANC", "MC", +1),
        ("RESERVA_CAPITAL",    "233XXXXXX", "LANC", "MD", -1),
        ("RESERVAS_LUCROS",    "235XXXXXX", "LANC", "MC", +1),
        ("RESERVAS_LUCROS",    "235XXXXXX", "LANC", "MD", -1),
        ("DEMAIS_RESERVAS",    "236XXXXXX", "LANC", "MC", +1),
        ("DEMAIS_RESERVAS",    "236XXXXXX", "LANC", "MD", -1),
    ]),
    ("1.01.09", "DIVIDENDOS_DISTRIBUIR", [
        # Sem conta mapeada na Lista de Equacoes -- sempre 0.
    ]),
    ("1.01.10", "SALDOS_FINAIS", [
        ("PAT_SOCIAL_CAPITAL", "231XXXXXX", "SALDO_FIM", "SC", +1),
        ("AFAC",               "232XXXXXX", "SALDO_FIM", "SC", +1),
        ("RESERVA_CAPITAL",    "233XXXXXX", "SALDO_FIM", "SC", +1),
        ("RESERVAS_LUCROS",    "235XXXXXX", "SALDO_FIM", "SC", +1),
        ("DEMAIS_RESERVAS",    "236XXXXXX", "SALDO_FIM", "SC", +1),
        ("RESULTADOS_ACUM",    "237XXXXXX", "SALDO_FIM", "SC", +1),
        ("RESULTADOS_ACUM",    "4XXXXXXXX", "LANC",      "SC", +1),
        ("RESULTADOS_ACUM",    "3XXXXXXXX", "LANC",      "SD", -1),
        ("AJUSTE_AVAL_PATRIM", "23411XXXX", "SALDO_FIM", "SC", +1),
    ]),
]

# Linhas que entram na identidade de fechamento (tudo exceto Saldos
# Iniciais e Saldos Finais, que sao os dois extremos da equacao).
LINHAS_VARIACAO = [chave for cod, chave, _ in LINHAS_DMPL
                   if chave not in ('SALDOS_INICIAIS', 'SALDOS_FINAIS')]

# ─────────────────────────────────────────────────────────────────────────────
#  GERADOR DE SQL
#  Uma unica consulta busca TODOS os componentes de TODAS as linhas, cada
#  um como uma coluna agregada SUM(CASE WHEN INMES=x AND conta...). Os
#  componentes sao depois somados/subtraidos em Python (calcular()) para
#  montar cada celula (linha x coluna) da DMPL.
# ─────────────────────────────────────────────────────────────────────────────
def _alias_componente(linha_chave, idx):
    """Gera um alias SQL legivel e seguro para o Oracle (limite de 30
    bytes em identificadores no Oracle 11g/12c). Usar o nome completo da
    linha (ex.: CONSTITUICAO_REVERSAO_RESERVAS_5, 32 chars) excede esse
    limite e causa ORA-00972 -- por isso o alias usa o CODIGO do item da
    Lista de Equacoes (ex.: '1.01.08' -> 'I010108') em vez do nome da
    linha. Isso mantem o alias rastreavel a olho nu direto na Lista de
    Equacoes/planilha oficial (olhando o SQL ou a aba 'Dados Brutos SQL'
    do Excel, sem precisar abrir o codigo), e fica bem abaixo do limite:
    'I010108_5' = 9 chars, contra os 32 de antes."""
    cod = _COD_POR_LINHA[linha_chave]
    cod_compacto = 'I' + cod.replace('.', '')  # '1.01.08' -> 'I010108'
    return f"{cod_compacto}_{idx}"


# Mapa linha_chave -> codigo do item (ex.: 'SALDOS_INICIAIS' -> '1.01.01'),
# usado por _alias_componente para montar o alias SQL legivel.
_COD_POR_LINHA = {linha_chave: cod for cod, linha_chave, _ in LINHAS_DMPL}


def _clausula_lanc(alias, mascara, tipo_mov):
    """Monta o SUM(CASE...) de um componente de MOVIMENTO (fonte 'LANC'),
    via LANCAMENTOCONTABIL, acumulado INMES BETWEEN 1 AND :mes -- mesmo
    padrao de bf.py/bo.py/dfc.py/dvp.py para itens de movimento/resultado."""
    ini, fim = _faixa(mascara)
    if ini == fim:
        cond_conta = f"o.COCONTACONTABIL = {ini}"
    else:
        cond_conta = f"o.COCONTACONTABIL BETWEEN {ini} AND {fim}"

    if tipo_mov == "SC":
        decode = "DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)"
        cond_extra = ""
    elif tipo_mov == "SD":
        decode = "DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)"
        cond_extra = ""
    elif tipo_mov == "MC":
        decode = "o.VALANCAMENTO"
        cond_extra = " AND o.INDEBITOCREDITO = 'C'"
    else:  # MD
        decode = "o.VALANCAMENTO"
        cond_extra = " AND o.INDEBITOCREDITO = 'D'"

    return (f"    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes AND {cond_conta}{cond_extra}\n"
            f"         THEN {decode} ELSE 0 END)\n"
            f"    AS {alias}")


def _clausula_saldo(alias, mascara, inmes_clause):
    """Monta o SUM(CASE...) de um componente de SALDO (fonte 'SALDO_INI'
    ou 'SALDO_FIM'), via VSALDOCONTABIL -- mesmo padrao usado em bf.py
    para o saldo de caixa (ANT_*/SEG_*). Contas de Patrimonio Liquido
    tem natureza credora (Tipo Movimento 'SC' na Lista de Equacoes para
    estas linhas), por isso a formula fixa e VACREDITO - VADEBITO."""
    ini, fim = _faixa(mascara)
    if ini == fim:
        cond_conta = f"v.COCONTACONTABIL = {ini}"
    else:
        cond_conta = f"v.COCONTACONTABIL BETWEEN {ini} AND {fim}"
    return (f"    SUM(CASE WHEN {inmes_clause} AND {cond_conta}\n"
            f"         THEN (v.VACREDITO - v.VADEBITO) ELSE 0 END)\n"
            f"    AS {alias}")


def montar_sql_lanc():
    """SQL unico (LANCAMENTOCONTABIL) para todos os componentes de
    MOVIMENTO/RESULTADO (fonte 'LANC'): Ajustes Exerc.Anteriores, Aumento
    de Capital, Resultado do Exercicio, Ajuste de Avaliacao Patrimonial,
    Constituicao/Reversao de Reservas, e a parte 4xxx/3xxx de Saldos
    Finais. Acumulado INMES BETWEEN 1 AND :mes (janeiro..mes corrente)."""
    partes = []
    for cod, linha_chave, componentes in LINHAS_DMPL:
        for idx, (col_chave, mascara, fonte, tipo_mov, sinal) in enumerate(componentes):
            if fonte != "LANC":
                continue
            alias = _alias_componente(linha_chave, idx)
            partes.append(_clausula_lanc(alias, mascara, tipo_mov))
    if not partes:
        return None
    corpo = ",\n\n".join(partes)
    return f"""
SELECT
{corpo}
FROM {{schema}}.LANCAMENTOCONTABIL o
WHERE (o.COCONTACONTABIL BETWEEN 200000000 AND 299999999
       OR o.COCONTACONTABIL BETWEEN 300000000 AND 499999999)
  {{filtro_ug}}
"""


def montar_sql_saldo(fonte_filtro, inmes_clause):
    """SQL unico (VSALDOCONTABIL) para os componentes de SALDO (fonte
    'SALDO_INI' ou 'SALDO_FIM'), contas de Patrimonio Liquido (231/232/
    233/235/236/237/23411). inmes_clause: 'v.INMES = 0' (saldo de
    abertura) ou 'v.INMES BETWEEN 0 AND :mes' (saldo ate o mes corrente),
    mesmo padrao de bf.py para o saldo de caixa."""
    partes = []
    for cod, linha_chave, componentes in LINHAS_DMPL:
        for idx, (col_chave, mascara, fonte, tipo_mov, sinal) in enumerate(componentes):
            if fonte != fonte_filtro:
                continue
            alias = _alias_componente(linha_chave, idx)
            partes.append(_clausula_saldo(alias, mascara, inmes_clause))
    if not partes:
        return None
    corpo = ",\n\n".join(partes)
    return f"""
SELECT
{corpo}
FROM {{schema}}.VSALDOCONTABIL v
WHERE v.COCONTACONTABIL BETWEEN 200000000 AND 299999999
  {{filtro_ug}}
"""

SQL_LANC       = montar_sql_lanc()
SQL_SALDO_INI  = montar_sql_saldo("SALDO_INI", "v.INMES = 0")
SQL_SALDO_FIM  = montar_sql_saldo("SALDO_FIM", "v.INMES BETWEEN 0 AND :mes")

# ─────────────────────────────────────────────────────────────────────────────
#  CONEXAO  (identica ao BF/BO/DFC/DVP)
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


def _executar(conn, sql, schema, coug, mes=None, alias_tabela='o'):
    """Executa o SQL via cursor nativo (compativel Oracle 11g thick mode),
    igual ao padrao do dvp.py/bf.py. alias_tabela indica se o filtro de UG
    deve usar 'o.COUG' (LANCAMENTOCONTABIL) ou 'v.COUG' (VSALDOCONTABIL).
    mes, se informado, substitui o placeholder ':mes' (acumulado jan..mes
    ou saldo ate o mes corrente)."""
    filtro = f"AND {alias_tabela}.COUG = {coug}" if coug else ""
    sql_fmt = sql.format(schema=schema, filtro_ug=filtro)
    if mes is not None:
        sql_fmt = sql_fmt.replace(':mes', str(mes))
    cur = conn.cursor()
    cur.execute(sql_fmt)
    cols = [d[0].upper() for d in cur.description]
    row = cur.fetchone()
    cur.close()
    return {k: float(v or 0) for k, v in zip(cols, row)}


def buscar_dados(conn, mes, ano, coug):
    """Busca todos os componentes brutos no schema MIL{ano} (Exercicio
    Atual), combinando tres consultas conforme a fonte de cada linha da
    Lista de Equacoes (ver nota no cabecalho do modulo):
      - SQL_LANC:      LANCAMENTOCONTABIL, INMES BETWEEN 1 AND mes
                       (linhas de movimento + Resultado do Exercicio)
      - SQL_SALDO_INI: VSALDOCONTABIL, INMES = 0
                       (Saldos Iniciais, contas de Patrimonio Liquido)
      - SQL_SALDO_FIM: VSALDOCONTABIL, INMES BETWEEN 0 AND mes
                       (Saldos Finais, contas de Patrimonio Liquido)"""
    print(f"  [DMPL] MIL{ano}  (LANCAMENTOCONTABIL INMES 1..{mes} p/ movimento; "
          f"VSALDOCONTABIL INMES=0 e 0..{mes} p/ saldos de Patrim. Liquido)...")
    schema = f"MIL{ano}"
    brutos = {}
    if SQL_LANC is not None:
        brutos.update(_executar(conn, SQL_LANC, schema, coug, mes=mes, alias_tabela='o'))
    if SQL_SALDO_INI is not None:
        brutos.update(_executar(conn, SQL_SALDO_INI, schema, coug, mes=mes, alias_tabela='v'))
    if SQL_SALDO_FIM is not None:
        brutos.update(_executar(conn, SQL_SALDO_FIM, schema, coug, mes=mes, alias_tabela='v'))
    return brutos


# ─────────────────────────────────────────────────────────────────────────────
#  TOTAIS DERIVADOS
# ─────────────────────────────────────────────────────────────────────────────
def calcular(brutos):
    """Recebe o dict bruto (componentes agregados do SQL) e devolve a
    matriz completa da DMPL: t[(linha_chave, coluna_chave)] = valor,
    alem da coluna TOTAL por linha."""
    t = {}

    for cod, linha_chave, componentes in LINHAS_DMPL:
        # Inicializa todas as colunas da linha em 0.0 (garante celulas
        # vazias do layout oficial, ex.: Resgate/Reemissao = tudo 0).
        for col_chave in COLUNA_CHAVES:
            t[(linha_chave, col_chave)] = 0.0

        for idx, (col_chave, mascara, fonte, tipo_mov, sinal) in enumerate(componentes):
            alias = _alias_componente(linha_chave, idx)
            valor = brutos.get(alias, 0.0)
            t[(linha_chave, col_chave)] += sinal * valor

        # Coluna 9 (TOTAL da linha) = soma das colunas 1 a 8.
        t[(linha_chave, "TOTAL")] = sum(t[(linha_chave, ck)] for ck in COLUNA_CHAVES)

    return t


# ─────────────────────────────────────────────────────────────────────────────
#  AUDITORIA DE INTEGRIDADE
# ─────────────────────────────────────────────────────────────────────────────
def auditoria_integridade(t):
    """
    Retorna lista de tuplas (status, titulo, detalhe) onde status e
    'OK', 'ERRO' ou 'ALERTA'.
    """
    achados = []
    colunas_check = COLUNA_CHAVES + ["TOTAL"]

    # ── Controle 1: Identidade de fechamento por coluna ───────────────────
    #    Saldos Iniciais + soma(linhas de variacao) = Saldos Finais
    for col_chave in colunas_check:
        label_col = next((d for c, k, d, _ in COLUNAS if k == col_chave),
                          'TOTAL') if col_chave != 'TOTAL' else 'TOTAL'
        ini = t.get(('SALDOS_INICIAIS', col_chave), 0.0)
        fim = t.get(('SALDOS_FINAIS', col_chave), 0.0)
        soma_var = sum(t.get((lk, col_chave), 0.0) for lk in LINHAS_VARIACAO)
        calc = ini + soma_var
        dif = calc - fim
        if abs(dif) < 1.00:
            achados.append(('OK', f'Fechamento DMPL — {label_col}',
                            f'Saldos Iniciais {ini:,.2f} + Variações {soma_var:,.2f} '
                            f'= {calc:,.2f}  =  Saldos Finais {fim:,.2f}'))
        else:
            achados.append(('ERRO', f'Fechamento DMPL — {label_col}',
                            f'Diferença: {dif:,.2f}  |  Saldos Iniciais {ini:,.2f} + '
                            f'Variações {soma_var:,.2f} = {calc:,.2f}  vs  '
                            f'Saldos Finais {fim:,.2f}  -- identidade NÃO fecha.'))

    # ── Controle 2: Coluna TOTAL = soma das colunas 1..8, em cada linha ───
    linhas_ok = True
    for cod, linha_chave, _ in LINHAS_DMPL:
        soma_cols = sum(t.get((linha_chave, ck), 0.0) for ck in COLUNA_CHAVES)
        total_reportado = t.get((linha_chave, "TOTAL"), 0.0)
        dif = soma_cols - total_reportado
        if abs(dif) >= 1.00:
            achados.append(('ERRO', f'TOTAL da linha {linha_chave}',
                            f'Diferença: {dif:,.2f}  |  soma das colunas '
                            f'{soma_cols:,.2f}  vs  TOTAL reportado '
                            f'{total_reportado:,.2f}'))
            linhas_ok = False
    if linhas_ok:
        achados.append(('OK', 'TOTAL de cada linha = soma das colunas 1..8',
                        'Todas as linhas conferem.'))

    # ── Controle 3: Resultado do Exercício (DMPL) deve ser coerente ───────
    #    Mesma grandeza que o Resultado Patrimonial do Período da DVP
    #    (VPA Total - VPD Total, classes 4 e 3) -- calculado aqui via
    #    LANCAMENTOCONTABIL, INMES BETWEEN 1 AND mes (mesmo range usado
    #    pela DVP). Registrado como informativo; o cruzamento real com
    #    o valor do dvp.py (se disponivel) e feito no mestre.py (X6).
    resultado_dmpl = t.get(('RESULTADO_EXERCICIO', 'TOTAL'), 0.0)
    achados.append(('INFO', 'Resultado do Exercício (DMPL)',
                    f'{resultado_dmpl:,.2f} — deve igualar o Resultado '
                    f'Patrimonial do Período (III) apurado na DVP, quando '
                    f'ambos cobrirem o mesmo período de referência.'))

    return achados


def imprimir_auditoria(achados):
    icones = {'OK': '✔', 'ERRO': '✘', 'ALERTA': '⚠', 'INFO': 'ℹ'}
    print(f"\n{'─'*60}")
    print("  AUDITORIA DE INTEGRIDADE — DMPL")
    print(f"{'─'*60}")
    for status, titulo, detalhe in achados:
        ic = icones.get(status, '?')
        print(f"  {ic} [{status:<6}] {titulo}")
        print(f"            {detalhe}")
    n_erro = sum(1 for s, _, _ in achados if s == 'ERRO')
    n_alerta = sum(1 for s, _, _ in achados if s == 'ALERTA')
    print(f"{'─'*60}")
    if n_erro:
        print(f"  RESULTADO: {n_erro} pendencia(s) de ERRO sinalizada(s) - ver detalhes acima.")
    elif n_alerta:
        print(f"  RESULTADO: Sem erros, mas {n_alerta} alerta(s) - revisar.")
    else:
        print(f"  RESULTADO: TODOS OS CONTROLES PASSARAM.")
    print(f"{'─'*60}")

# ─────────────────────────────────────────────────────────────────────────────
#  ESTRUTURA DE EXIBICAO  (espelha o ListaDMPL.pdf -- 10 linhas x 9 colunas)
# ─────────────────────────────────────────────────────────────────────────────
ESTRUTURA_DMPL = [
    ("SALDOS_INICIAIS",               "Saldos Iniciais",                     True),
    ("AJUSTES_EXERC_ANTERIORES",      "Ajustes de exercícios anteriores",    False),
    ("AUMENTO_CAPITAL",               "Aumento de capital",                  False),
    ("RESGATE_REEMISSAO_ACOES",       "Resgate/Reemissão Ações e Cotas",     False),
    ("JUROS_CAPITAL_PROPRIO",         "Juros sobre capital próprio",         False),
    ("RESULTADO_EXERCICIO",           "Resultado do exercício",              False),
    ("AJUSTE_AVALIACAO_PATRIMONIAL",  "Ajuste de avaliação patrimonial",     False),
    ("CONSTITUICAO_REVERSAO_RESERVAS","Constituição / Reversão de reservas", False),
    ("DIVIDENDOS_DISTRIBUIR",         "Dividendos a distribuir",             False),
    ("SALDOS_FINAIS",                 "Saldos finais",                       True),
]

CABECALHO_COLUNAS = [d for _, _, d, _ in COLUNAS] + ["TOTAL"]
TODAS_COL_CHAVES = COLUNA_CHAVES + ["TOTAL"]

# ─────────────────────────────────────────────────────────────────────────────
#  EXCEL
# ─────────────────────────────────────────────────────────────────────────────
FMT_BRL = '#,##0.00'
def _fill(h): return PatternFill('solid', fgColor=h)
def _side(s='thin'): return Side(border_style=s, color='B8CCE4')
B_THIN = Border(left=_side(), right=_side(), top=_side(), bottom=_side())
F_HDR  = _fill('BDD7EE')
F_GRAY = _fill('DCE6F1'); F_YELL = _fill('9DC3E6'); F_WHIT = _fill('F5F9FD')


def cel(ws, r, c, v='', bold=False, sz=9, bg=None, ha='left', brd=None,
        fmt=None, wrap=False):
    x = ws.cell(row=r, column=c, value=v)
    x.font = Font(name='Arial', bold=bold, size=sz, color='000000')
    x.alignment = Alignment(horizontal=ha, vertical='center', wrap_text=wrap)
    if bg: x.fill = bg
    if brd: x.border = brd
    if fmt and v not in ('', None): x.number_format = fmt
    return x


def gerar_excel(t, mes, ano, ug_label, output_path, achados=None):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = 'DMPL'
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = 9

    ncol = 1 + len(CABECALHO_COLUNAS)  # coluna A = Especificacao
    ws.column_dimensions['A'].width = 32
    for i in range(2, ncol + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 17

    last_col_letter = openpyxl.utils.get_column_letter(ncol)
    ws.merge_cells(f'A1:{last_col_letter}3')
    c = ws['A1']
    c.value = ('GOVERNO DO DISTRITO FEDERAL\n'
               'Demonstração das Mutações do Patrimônio Líquido\nVersão : 1')
    c.font = Font(name='Arial', bold=True, size=12, color='2E5C8A')
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 45

    ws.merge_cells(f'A4:{last_col_letter}4')
    cel(ws, 4, 1,
        f'Mês de Referência   {mes:02d} - {MESES[mes]}          {ug_label}   '
        f'(Saldos Iniciais = INMES 0; demais linhas = INMES 13/encerramento)',
        bold=True, sz=9)
    ws.row_dimensions[5].height = 4

    H = 6
    ws.row_dimensions[H].height = 40
    cel(ws, H, 1, 'Especificação', bold=True, bg=F_HDR, brd=B_THIN, ha='center', wrap=True)
    for j, desc in enumerate(CABECALHO_COLUNAS, start=2):
        cel(ws, H, j, desc, bold=True, bg=F_HDR, brd=B_THIN, ha='center', wrap=True)

    R0 = 7
    for i, (linha_chave, desc, bold) in enumerate(ESTRUTURA_DMPL):
        r = R0 + i
        bg = F_GRAY if bold else F_WHIT
        cel(ws, r, 1, desc, bold=bold, bg=bg, brd=B_THIN, ha='left')
        for j, col_chave in enumerate(TODAS_COL_CHAVES, start=2):
            valor = t.get((linha_chave, col_chave), 0.0)
            cel(ws, r, j, valor, bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)

    FT = R0 + len(ESTRUTURA_DMPL) + 2
    cel(ws, FT, 1, 'Página: 1', sz=8)
    cel(ws, FT, ncol, f'Emitido em: {datetime.now():%d/%m/%Y}', sz=8, ha='right')

    # ── Aba de dados brutos ─────────────────────────────────────────────────
    ws2 = wb.create_sheet('Dados Brutos SQL')
    j = 1
    for (linha_chave, col_chave), v in t.items():
        ws2.cell(row=1, column=j, value=f'{linha_chave}|{col_chave}').font = \
            Font(bold=True, name='Arial', size=8)
        c2 = ws2.cell(row=2, column=j, value=v)
        c2.number_format = FMT_BRL
        j += 1

    if achados:
        ws3 = wb.create_sheet('Auditoria Integridade')
        ws3.column_dimensions['A'].width = 10
        ws3.column_dimensions['B'].width = 45
        ws3.column_dimensions['C'].width = 90
        hdr_bg = _fill('2E5C8A')
        cel(ws3, 1, 1, 'Status', bold=True, bg=hdr_bg, ha='center', brd=B_THIN)
        ws3['A1'].font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        cel(ws3, 1, 2, 'Controle', bold=True, bg=hdr_bg, ha='center', brd=B_THIN)
        ws3['B1'].font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        cel(ws3, 1, 3, 'Detalhe', bold=True, bg=hdr_bg, ha='center', brd=B_THIN)
        ws3['C1'].font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        cores = {'ERRO': _fill('FAD7DA'), 'ALERTA': _fill('FCF0CE'),
                 'INFO': _fill('DDEBF7')}
        for i, (status, titulo, detalhe) in enumerate(achados, start=2):
            bg = cores.get(status, _fill('D9F0DD'))
            cel(ws3, i, 1, status, bold=True, bg=bg, brd=B_THIN, ha='center')
            cel(ws3, i, 2, titulo, bg=bg, brd=B_THIN, ha='left', wrap=True)
            cel(ws3, i, 3, detalhe, bg=bg, brd=B_THIN, ha='left', wrap=True)

    wb.save(output_path)
    print(f"  Excel salvo: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  PDF
# ─────────────────────────────────────────────────────────────────────────────
def _brl(v):
    s = f'{v:,.2f}'.replace(',', '§').replace('.', ',').replace('§', '.')
    return s


def gerar_pdf(t, mes, ano, ug_label, output_path, achados=None):
    if not REPORTLAB_OK:
        print("  AVISO: reportlab nao instalado -- PDF nao gerado "
              "(pip install reportlab).")
        return

    doc = SimpleDocTemplate(str(output_path), pagesize=landscape(A4),
                             leftMargin=10*mm, rightMargin=10*mm,
                             topMargin=10*mm, bottomMargin=10*mm)
    styles = getSampleStyleSheet()
    elems = []

    st_t = ParagraphStyle('t', parent=styles['Normal'], fontName='Helvetica-Bold',
                           fontSize=12, textColor=colors.HexColor('#2E5C8A'), leading=14)
    st_s = ParagraphStyle('s', parent=styles['Normal'], fontSize=8, alignment=TA_RIGHT)
    st_r = ParagraphStyle('r', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9)

    cab = Table([[Paragraph('GOVERNO DO DISTRITO FEDERAL<br/>'
                             'Demonstração das Mutações do Patrimônio Líquido<br/>'
                             'Versão 1', st_t),
                  Paragraph(f'Exercício {ano}<br/>PSIAG550<br/>'
                            f'Posição em: {datetime.now():%d/%m/%Y às %H:%M:%S}', st_s)]],
                colWidths=[200*mm, 70*mm])
    cab.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    elems.append(cab)
    elems.append(Paragraph(
        f'Mês de Referência: {mes:02d} - {MESES[mes]} &nbsp; {ug_label}', st_r))
    elems.append(Spacer(1, 4*mm))

    st_cell_desc = ParagraphStyle('cd', fontName='Helvetica', fontSize=6.5, leading=7.5)
    st_cell_desc_b = ParagraphStyle('cdb', fontName='Helvetica-Bold', fontSize=6.5, leading=7.5)
    st_cell_val = ParagraphStyle('cv', fontName='Helvetica', fontSize=6.5,
                                  alignment=TA_RIGHT, leading=7.5)
    st_cell_val_b = ParagraphStyle('cvb', fontName='Helvetica-Bold', fontSize=6.5,
                                    alignment=TA_RIGHT, leading=7.5)
    st_hdr = ParagraphStyle('hdr', fontName='Helvetica-Bold', fontSize=7,
                             alignment=TA_CENTER, textColor=colors.black, leading=8)

    header_row = [Paragraph('Especificação', st_hdr)] + \
                 [Paragraph(d, st_hdr) for d in CABECALHO_COLUNAS]
    dados = [header_row]
    for linha_chave, desc, bold in ESTRUTURA_DMPL:
        sd, sv = (st_cell_desc_b, st_cell_val_b) if bold else (st_cell_desc, st_cell_val)
        linha = [Paragraph(desc, sd)]
        for col_chave in TODAS_COL_CHAVES:
            v = t.get((linha_chave, col_chave), 0.0)
            linha.append(Paragraph(_brl(v), sv))
        dados.append(linha)

    W_DESC = 38*mm
    n_val_cols = len(TODAS_COL_CHAVES)
    W_VAL = (270*mm - W_DESC) / n_val_cols
    tab = Table(dados, colWidths=[W_DESC] + [W_VAL]*n_val_cols, repeatRows=1)
    ts = [('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
          ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
          ('TOPPADDING', (0, 0), (-1, -1), 2),
          ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
          ('LEFTPADDING', (0, 0), (-1, -1), 3),
          ('RIGHTPADDING', (0, 0), (-1, -1), 3),
          ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#BDD7EE'))]
    for i, (linha_chave, desc, bold) in enumerate(ESTRUTURA_DMPL, start=1):
        if bold:
            ts.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#EBF3FB')))
    tab.setStyle(TableStyle(ts))
    elems.append(tab)

    # ── Auditoria de Integridade (rodape) ───────────────────────────────────
    if achados:
        elems.append(Spacer(1, 5*mm))
        st_aud_tit = ParagraphStyle('audt', parent=styles['Normal'],
                       fontName='Helvetica-Bold', fontSize=10,
                       textColor=colors.HexColor('#2E5C8A'))
        elems.append(Paragraph('AUDITORIA DE INTEGRIDADE', st_aud_tit))
        elems.append(Spacer(1, 1*mm))

        st_cell_b = ParagraphStyle('cb', parent=styles['Normal'],
                      fontName='Helvetica-Bold', fontSize=7.5)
        st_cell_n = ParagraphStyle('cn', parent=styles['Normal'], fontSize=7.5)

        aud_rows = [['Status', 'Controle', 'Detalhe']]
        for status, titulo, detalhe in achados:
            aud_rows.append([
                Paragraph(status, st_cell_b),
                Paragraph(titulo, st_cell_b),
                Paragraph(detalhe, st_cell_n),
            ])
        aud_tab = Table(aud_rows, colWidths=[20*mm, 75*mm, 178*mm], repeatRows=1)
        aud_ts = [
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E5C8A')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        for i, (status, _, _) in enumerate(achados, start=1):
            cor = {'ERRO': '#FAD7DA', 'ALERTA': '#FCF0CE',
                   'INFO': '#E8F1FA'}.get(status, '#D9F0DD')
            aud_ts.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor(cor)))
            aud_ts.append(('ALIGN', (0, i), (0, i), 'CENTER'))
        aud_tab.setStyle(TableStyle(aud_ts))
        elems.append(aud_tab)

        n_erro = sum(1 for s, _, _ in achados if s == 'ERRO')
        n_alerta = sum(1 for s, _, _ in achados if s == 'ALERTA')
        cor_res = '#B00000' if n_erro else ('#9C6500' if n_alerta else '#1F5C2E')
        st_res = ParagraphStyle('res', parent=styles['Normal'],
                   fontName='Helvetica-Bold', fontSize=8.5,
                   textColor=colors.HexColor(cor_res))
        elems.append(Spacer(1, 2*mm))
        if n_erro:
            elems.append(Paragraph(
                f'RESULTADO: {n_erro} pendência(s) de ERRO sinalizada(s) - ver detalhes acima.',
                st_res))
        elif n_alerta:
            elems.append(Paragraph(
                f'RESULTADO: Sem erros, mas {n_alerta} alerta(s) - revisar.', st_res))
        else:
            elems.append(Paragraph('RESULTADO: TODOS OS CONTROLES PASSARAM.', st_res))

    doc.build(elems)
    print(f"  PDF salvo:   {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  DIAGNOSTICO  (mesmo padrao BF/DFC/DVP, adaptado p/ contas 2/3/4 da DMPL)
# ─────────────────────────────────────────────────────────────────────────────
def diagnostico(conn, ano, mes):
    print("\n" + "="*64)
    print("  DIAGNOSTICO - ESTRUTURA DO BANCO PARA DMPL")
    print("="*64)

    print(f"\n[1] VSALDOCONTABIL existe e tem dados p/ contas 23X (Patrim. "
          f"Liquido) em MIL{ano}? (fonte usada p/ Saldos Iniciais/Finais)")
    try:
        q = f"""SELECT v.INMES, COUNT(*) QTD,
                    SUM(v.VACREDITO - v.VADEBITO) SALDO
                FROM MIL{ano}.VSALDOCONTABIL v
                WHERE v.COCONTACONTABIL BETWEEN 230000000 AND 239999999
                GROUP BY v.INMES ORDER BY v.INMES"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM registro em VSALDOCONTABIL p/ contas 23XXXXXXX "
                  f"-- view pode nao existir, ou coluna/owner diferente.")
        for _, r in df.iterrows():
            print(f"    INMES={int(r['INMES']):>3}  qtd={int(r['QTD']):>9}  "
                  f"saldo(SC)={float(r['SALDO'] or 0):>22,.2f}")
        print(f"    >>> Esperado: INMES=0 com saldo grande (Saldos Iniciais) e "
              f"INMES={mes} (ou acumulado 0..{mes}) tambem com saldo grande.")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2] LANCAMENTOCONTABIL, INMES BETWEEN 1 AND {mes} -- contas "
          f"23X/3X/4X em MIL{ano} (fonte usada p/ movimento/Resultado):")
    try:
        q = f"""SELECT o.COCONTACONTABIL/1000000 AS CONTAMAE, COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)) SALDO
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.INMES BETWEEN 1 AND {mes}
                  AND (o.COCONTACONTABIL BETWEEN 230000000 AND 239999999
                       OR o.COCONTACONTABIL BETWEEN 300000000 AND 499999999)
                GROUP BY o.COCONTACONTABIL/1000000 ORDER BY 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM registro em LANCAMENTOCONTABIL (INMES 1..{mes}) "
                  f"para contas 23X/3X/4X.")
        for _, r in df.iterrows():
            print(f"    ContaMae={int(r['CONTAMAE']):>5}  qtd={int(r['QTD']):>9}  "
                  f"saldo(SC)={float(r['SALDO'] or 0):>22,.2f}")
        print(f"    >>> Contas 4xx (VPA) e 3xx (VPD) devem ter saldo "
              f"significativo aqui (mesma base usada pela DVP).")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[3] Comparativo: LANCAMENTOCONTABIL com INMES=0 / INMES=13 "
          f"LITERAL para contas 23X em MIL{ano} (interpretacao ANTIGA, "
          f"que retornou zero em producao -- mantido aqui so para "
          f"referencia/diagnostico, NAO e mais usado pelo modulo):")
    try:
        q = f"""SELECT o.INMES, COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)) SALDO
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL BETWEEN 230000000 AND 239999999
                  AND o.INMES IN (0, 13)
                GROUP BY o.INMES ORDER BY o.INMES"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM registro (confirma por que a interpretacao "
                  f"literal INMES=0/13 em LANCAMENTOCONTABIL devolvia zero).")
        for _, r in df.iterrows():
            print(f"    INMES={int(r['INMES']):>3}  qtd={int(r['QTD']):>9}  "
                  f"saldo(SC)={float(r['SALDO'] or 0):>22,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print("="*64 + "\n")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(
        description='Demonstração das Mutações do Patrimônio Líquido - GDF (DMPL)')
    p.add_argument('--mes', type=int, default=5,
                   help='Mês de referência 1-12 (informativo no cabeçalho; '
                        'as equações usam INMES=0/13 fixos por linha)')
    p.add_argument('--ano', type=int, default=2026)
    p.add_argument('--ug', type=str, default=None,
                   help='Código da Unidade Gestora (omitir = Consolidado)')
    p.add_argument('--saida', type=str, default=None,
                   help='Caminho base dos arquivos de saída (sem extensão)')
    p.add_argument('--formato', choices=['ambos', 'excel', 'pdf'], default='ambos')
    p.add_argument('--diag', action='store_true',
                   help='Executa diagnóstico do banco e encerra')
    a = p.parse_args()

    if not 1 <= a.mes <= 12:
        print("ERRO: --mes deve estar entre 1 e 12."); sys.exit(1)

    ug_label = f'UG: {a.ug}' if a.ug else 'Consolidado'
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome = (f'DMPL_GDF_{a.ano}_{a.mes:02d}'
            + (f'_UG{a.ug}' if a.ug else '_Consolidado')
            + f'_{timestamp}')
    base = (Path(a.saida).with_suffix('') if a.saida else OUTPUT_DIR / nome)
    saida_xlsx = base.with_suffix('.xlsx')
    saida_pdf = base.with_suffix('.pdf')

    print(f"\n{'='*60}")
    print(f"  DEMONSTRAÇÃO DAS MUTAÇÕES DO PATRIMÔNIO LÍQUIDO — DMPL")
    print(f"  {MESES[a.mes]}/{a.ano}  |  {ug_label}")
    print(f"{'='*60}")

    print("\n[1/4] Conectando ao Oracle...")
    conn = conectar_oracle()

    if a.diag:
        diagnostico(conn, a.ano, a.mes)
        conn.close()
        return

    print("\n[2/4] Buscando dados (INMES=0 e INMES=13)...")
    brutos = buscar_dados(conn, a.mes, a.ano, a.ug)
    conn.close()

    print("\n[3/4] Calculando matriz DMPL (10 linhas x 9 colunas)...")
    t = calcular(brutos)
    achados = auditoria_integridade(t)

    print("\n[4/4] Gerando arquivos...")
    if a.formato in ('ambos', 'excel'):
        gerar_excel(t, a.mes, a.ano, ug_label, saida_xlsx, achados=achados)
    if a.formato in ('ambos', 'pdf'):
        gerar_pdf(t, a.mes, a.ano, ug_label, saida_pdf, achados=achados)

    print(f"\n{'─'*60}")
    print("  CONFERÊNCIA RÁPIDA:")
    print(f"{'─'*60}")
    for cod, linha_chave, _ in LINHAS_DMPL:
        total = t.get((linha_chave, 'TOTAL'), 0.0)
        print(f"    {cod:<10} {linha_chave:<32}: {total:>20,.2f}")
    print(f"{'─'*60}")

    imprimir_auditoria(achados)

    print(f"\n  Concluído em {datetime.now():%H:%M:%S}\n")


if __name__ == '__main__':
    main()