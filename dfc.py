# -*- coding: utf-8 -*-
"""
=============================================================================
  DEMONSTRACAO DOS FLUXOS DE CAIXA - GDF  (PSIAG550 / Anexo 15)
  SQL fixo, equacoes embutidas conforme a Lista de Equacoes de Balanco
  Tipo 07 - Fluxo de Caixa.

  LOGICA DE MES (INMES):
    0  = saldo de abertura  (Caixa Inicial)
    1  = movimento acumulado jan..mes  (despesas e aplicacoes)
   13  = movimento acumulado + encerramento (receitas e transferencias)
  TIPOS DE MOVIMENTO:
    SC = Saldo Credor  =  C - D
    SD = Saldo Devedor =  D - C
    MC = Movimento a Credito (so creditos)
    MD = Movimento a Debito  (so debitos)

  Dependencias:  pip install oracledb openpyxl pandas reportlab
  Uso:
      python dfc.py --mes 5 --ano 2026
      python dfc.py --mes 5 --ano 2026 --ug 130101
      python dfc.py --mes 5 --ano 2026 --formato pdf
      python dfc.py --mes 5 --ano 2026 --diag
=============================================================================
"""
import argparse, sys
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
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, PageBreak)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ALCANCE DO NOTAEMPENHO — ampliado em 14/08/2026
#
# A busca de natureza de despesa passou de MIL2023..MIL{ano} para
# MIL2020..MIL{ano} (7 schemas, 14 blocos EXISTS). Pagamentos de empenhos
# emitidos entre 2020 e 2022 antes nao encontravam contrapartida e ficavam
# sem classificacao por natureza.
#
# EFEITO ESPERADO: as linhas classificadas por natureza (Pessoal, Juros,
# Transferencias Concedidas, Aquisicao de Ativo etc.) so podem CRESCER.
# Em ago/2026 o Pessoal ja estava R$ 1.111.652,78 ACIMA do PSIAG550, entao
# esta mudanca afasta ainda mais desse valor. Se a divergencia aumentar,
# e sinal de que o oficial usa alcance MENOR, nao maior — e a ampliacao
# deve ser revertida.
#
# CUSTO: se o parametro ano for <= 2025, o schema aparece duas vezes no
# UNION ALL (na lista fixa e em MIL{ano}). O EXISTS nao duplica valor,
# mas a consulta fica mais cara. Cada bloco agora varre 7 tabelas.

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIGURACAO  (ajuste conforme ambiente)
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
#  SQL PRINCIPAL  (MIL{ano}.LANCAMENTOCONTABIL)
#  Cada coluna corresponde a um item da DFC (Equacoes de Balanco Tipo 07).
#  Convencao de sinal:
#    SC positivo  -> DECODE(INDC,'C',val,'D',-val)
#    SD positivo  -> DECODE(INDC,'D',val,'C',-val)
#    MC           -> apenas creditos
#    MD           -> apenas debitos
#  Natureza da receita/despesa via SUBSTR(o.COCONTACORRENTE, 1, N)
#  como filtros adicionais quando a equacao especifica.
# ─────────────────────────────────────────────────────────────────────────────
SQL_DFC = """
-- ══════════════════════════════════════════════════════════════════════════
--  DEMONSTRACAO DOS FLUXOS DE CAIXA — SQL PRINCIPAL
--
--  ESTRUTURA DE JOINS:
--    o  = MIL{ano}.LANCAMENTOCONTABIL   (movimentos contábeis)
--    ne = MIL{ano}.NOTAEMPENHO          (natureza da despesa via CONATUREZA)
--         JOIN: SUBSTR(o.COCONTACORRENTE,1,11) = ne.NUNE
--               AND o.COUG = ne.COUG AND o.COGESTAO = ne.COGESTAO
--
--  RECEITAS  -> filtro por SUBSTR(o.COCONTACORRENTE,1,N) = prefixo categoria
--  DESPESAS  -> filtro por SUBSTR(ne.CONATUREZA,1,N)     = prefixo natureza
--
--  Contas de despesa orçamentária executada (SC = C-D):
--    622920104  Despesa Orçamentária Paga
--    6322XXXXX  Despesa Restos a Pagar Paga (RP)
--    631400000  Pagamento RPNP
--    631820000  Pagamento RPP
-- ══════════════════════════════════════════════════════════════════════════
SELECT
-- ──────────────────────────────────────────────────────────────────────────
--  RECEITAS OPERACIONAIS (conta 621200000 SC / 6213XXXXX SD)
--  Filtro de categoria: SUBSTR(o.COCONTACORRENTE, 1, N)
-- ──────────────────────────────────────────────────────────────────────────

-- 1.01.01  IMPOSTOS, TAXAS E CONTRIBUICOES (prefixo 11, 71)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('11','71')
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('11','71')
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS IMPOSTOS,

-- 1.01.02  RECEITA DE CONTRIBUICOES (prefixo 12, 72)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('12','72')
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('12','72')
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS CONTRIBUICOES,

-- 1.01.03  RECEITA PATRIMONIAL (prefixo 13,73 exceto remuneracao disponib.)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('13','73')
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,7)
                  NOT IN ('1321001','1321002','1321003','1321004',
                           '7321001','7321002','7321003','7321004')
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('13','73')
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,7)
                  NOT IN ('1321001','1321002','1321003','1321004',
                           '7321001','7321002','7321003','7321004')
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS PATRIMONIAL,

-- 1.01.04  RECEITA AGROPECUARIA (prefixo 14,74,94)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('14','74','94')
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('14','74','94')
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS AGROPECUARIA,

-- 1.01.05  RECEITA INDUSTRIAL (prefixo 15,75,95)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('15','75','95')
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('15','75','95')
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS INDUSTRIAL,

-- 1.01.06  RECEITA DE SERVICOS (prefixo 16,76)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('16','76')
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('16','76')
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS SERVICOS,

-- 1.01.07  REMUNERACAO DAS DISPONIBILIDADES (1321001..7321004, 2930, 8930)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND (SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,7)
                   IN ('1321001','1321002','1321003','1321004',
                       '7321001','7321002','7321003','7321004')
                OR SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,4) IN ('2930','8930'))
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND (SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,7)
                   IN ('1321001','1321002','1321003','1321004',
                       '7321001','7321002','7321003','7321004')
                OR SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,4) IN ('2930','8930'))
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS REMUNERACAO_DISP,

-- 1.01.08  TRANSFERENCIAS RECEBIDAS (prefixo 17,24,77,97)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('17','24','77','97')
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('17','24','77','97')
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS TRANSF_RECEBIDAS,

-- 1.01.09  OUTRAS RECEITAS / INGRESSOS OPERACIONAIS
--          prefixo 19,79 + codigos especificos + contas 451XXXXXX + depositos
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND (SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('19','79','29','89')
                OR SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,4) IN ('2920','8920')
                OR TO_CHAR(o.COCONTACORRENTE)
                   IN ('29400011','29900011','89400011','89900011','8212041','29999901'))
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND (SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('19','79','29','89')
                OR SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,4) IN ('2920','8920')
                OR TO_CHAR(o.COCONTACORRENTE)
                   IN ('29400011','29900011','89400011','89900011','8212041','29999901'))
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (451120100,451120200,451120300,451120400,
                                        451120900,451121300,451129900,
                                        451220101,451220104,451220109,451220199)
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 451300000 AND 451399999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes AND o.INDEBITOCREDITO = 'C'
              AND (o.COCONTACONTABIL BETWEEN 218810300 AND 218810399
                OR o.COCONTACONTABIL BETWEEN 218810400 AND 218810499
                OR o.COCONTACONTABIL BETWEEN 218810800 AND 218810899
                OR o.COCONTACONTABIL BETWEEN 218817000 AND 218817099
                OR o.COCONTACONTABIL BETWEEN 218820400 AND 218820499
                OR o.COCONTACONTABIL BETWEEN 218827000 AND 218827099
                OR o.COCONTACONTABIL BETWEEN 218837000 AND 218837099
                OR o.COCONTACONTABIL IN (218829900,218924500,218910105,218910108,218920199))
              AND o.COCONTACONTABIL NOT IN (218810304,218810405)
         THEN o.VALANCAMENTO ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND (o.COCONTACONTABIL BETWEEN 218815000 AND 218815099
                OR o.COCONTACONTABIL BETWEEN 218825000 AND 218825099)
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes AND o.INDEBITOCREDITO = 'C'
              AND o.COCONTACONTABIL BETWEEN 113500000 AND 113599999
              AND o.COCONTACONTABIL NOT IN (113510200,113510300,113510500,113510802)
         THEN o.VALANCAMENTO ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes AND o.INDEBITOCREDITO = 'C'
              AND (o.COCONTACONTABIL BETWEEN 113810600 AND 113810699
                OR o.COCONTACONTABIL BETWEEN 113811700 AND 113811799
                OR o.COCONTACONTABIL BETWEEN 113819100 AND 113819199
                OR o.COCONTACONTABIL BETWEEN 113821700 AND 113821799
                OR o.COCONTACONTABIL BETWEEN 113829100 AND 113829199
                OR o.COCONTACONTABIL IN (113819700,113829700))
         THEN o.VALANCAMENTO ELSE 0 END)
    AS OUTRAS_REC_OPERAC,

-- ──────────────────────────────────────────────────────────────────────────
--  DESEMBOLSOS OPERACIONAIS
--  JOIN com NOTAEMPENHO para obter CONATUREZA (natureza da despesa)
--  SUBSTR(o.COCONTACORRENTE,1,11) = ne.NUNE (numero do empenho)
--  Contas: 622920104, 6322XXXXX, 631400000, 631820000
-- ──────────────────────────────────────────────────────────────────────────

-- 1.02.01  PESSOAL E DEMAIS DESPESAS
--   NatDesp: 3190XXXX, 3390XXXX, 3371XXXX, 4471XXXX, 317170XX, 318084XX
--   (CORRIGIDO: 3120/3191/3391/4591/4491 NAO pertencem aqui -- sao
--    TRANSFERENCIAS CONCEDIDAS conforme planilha oficial de equacoes)
--   JOIN multi-schema (2023/2025/{ano}) via COUGCONTAB/COGESTAOCONTAB
--   (Restos a Pagar referenciam empenhos de exercicios anteriores; a UG
--    pagadora COUG pode diferir da UG contabil COUGCONTAB do empenho)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (622920104,631400000,631820000)
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  (SUBSTR(ne.CONATUREZA,1,4)
                          IN ('3190','3390','4471','3371')
                       OR SUBSTR(ne.CONATUREZA,1,6)
                          IN ('317170','318084'))
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 632200000 AND 632299999
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  (SUBSTR(ne.CONATUREZA,1,4)
                          IN ('3190','3390','4471','3371')
                       OR SUBSTR(ne.CONATUREZA,1,6)
                          IN ('317170','318084'))
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS PESSOAL_DESPESAS,

-- 1.02.02  JUROS E ENCARGOS DA DIVIDA
--   NatDesp: 329021XX..329026XX, 329092XX, 328084XX
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (622920104,631400000,631820000)
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6)
                         IN ('329021','329022','329023','329024',
                             '329025','329026','329092','328084')
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 632200000 AND 632299999
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6)
                         IN ('329021','329022','329023','329024',
                             '329025','329026','329092','328084')
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS JUROS_ENCARGOS,

-- 1.02.03  TRANSFERENCIAS CONCEDIDAS
--   NatDesp (17 ativas na Lista de Equacoes, item 1.02.03):
--            3120XXXX, 3191XXXX, 3320XXXX, 3340XXXX, 3350XXXX, 3360XXXX,
--            336783XX, 3370XXXX, 3380XXXX, 3391XXXX, 4420XXXX, 4440XXXX,
--            4450XXXX, 4470XXXX, 4480XXXX, 4491XXXX, 4591XXXX
--
--   ATENCAO: a versao anterior deste comentario listava tambem 449093XX,
--   que NAO esta na tupla do SQL abaixo e NAO deve estar: a Lista aloca
--   449093XX no item 1.04.03 (Outros Desembolsos de Investimento).
--   Conferido em 14/08/2026 contra a Lista filtrando Status = Ativo:
--   a tupla do SQL bate exatamente com as 17 naturezas vigentes.
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (622920104,631400000,631820000)
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  (SUBSTR(ne.CONATUREZA,1,4)
                          IN ('3120','3191','3320','3340','3350','3360',
                              '3370','3380','3391','4420','4440','4450',
                              '4470','4480','4491','4591')
                       OR SUBSTR(ne.CONATUREZA,1,6)
                          IN ('336783'))
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 632200000 AND 632299999
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  (SUBSTR(ne.CONATUREZA,1,4)
                          IN ('3120','3191','3320','3340','3350','3360',
                              '3370','3380','3391','4420','4440','4450',
                              '4470','4480','4491','4591')
                       OR SUBSTR(ne.CONATUREZA,1,6)
                          IN ('336783'))
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS TRANSF_CONCEDIDAS,

-- 1.02.04  OUTROS DESEMBOLSOS OPERACIONAIS
--   Contas 351XXXXXX (SD) + depositos restituiveis (MD) + ajustes 237XXXXXX
--   + contas 113XXXXXX aplicacoes (MD)
--   (estas contas NAO usam empenho — filtro direto por conta contabil)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (351120100,351120200,351120300,351120400,
                                        351120900,351129900,351220101,351220104,
                                        351220109,351220199)
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 351300000 AND 351399999
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes AND o.INDEBITOCREDITO = 'D'
              AND (o.COCONTACONTABIL BETWEEN 218810300 AND 218810399
                OR o.COCONTACONTABIL BETWEEN 218810400 AND 218810499
                OR o.COCONTACONTABIL BETWEEN 218810800 AND 218810899
                OR o.COCONTACONTABIL BETWEEN 218817000 AND 218817099
                OR o.COCONTACONTABIL BETWEEN 218820400 AND 218820499
                OR o.COCONTACONTABIL BETWEEN 218827000 AND 218827099
                OR o.COCONTACONTABIL BETWEEN 218837000 AND 218837099
                OR o.COCONTACONTABIL IN (218829900,218924500,218910105,218910108,218920199))
              AND o.COCONTACONTABIL NOT IN (218810304,218810405)
         THEN o.VALANCAMENTO ELSE 0 END)
  -- Item 1.02.04. 5o digito = esfera: 1 consolidado, 2 intra, 3 inter.
  -- As duas inter (237130301, 237230301) faltavam ate 14/08/2026.
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (237110301,237120301,237210301,237220301,237130301,237230301)
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes AND o.INDEBITOCREDITO = 'D'
              AND o.COCONTACONTABIL BETWEEN 113500000 AND 113599999
              AND o.COCONTACONTABIL NOT IN (113510200,113510300,113510500,113510802)
         THEN o.VALANCAMENTO ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes AND o.INDEBITOCREDITO = 'D'
              AND (o.COCONTACONTABIL BETWEEN 113810600 AND 113810699
                OR o.COCONTACONTABIL BETWEEN 113811700 AND 113811799
                OR o.COCONTACONTABIL BETWEEN 113819100 AND 113819199
                OR o.COCONTACONTABIL BETWEEN 113821700 AND 113821799
                OR o.COCONTACONTABIL BETWEEN 113829100 AND 113829199
                OR o.COCONTACONTABIL IN (113819700,113829700))
         THEN o.VALANCAMENTO ELSE 0 END)
    AS OUTROS_DESEMB_OPERAC,

-- ──────────────────────────────────────────────────────────────────────────
--  INVESTIMENTO — INGRESSOS
-- ──────────────────────────────────────────────────────────────────────────

-- 1.03.01  ALIENACAO DE BENS (prefixo 22, 82)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('22','82')
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('22','82')
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS ALIENACAO_BENS,

-- 1.03.02  AMORTIZACAO DE EMPRESTIMOS (prefixo 23, 83)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('23','83')
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) IN ('23','83')
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS AMORTIZ_EMPREST,

-- 1.03.03  OUTROS INGRESSOS DE INVESTIMENTO
--   CONFIRMADO por diagnostico: a equacao oficial usa SOMENTE a conta
--   821192501 (MD). As contas 122310103/104 e 1141XXXXX/1144XXXXX NAO
--   pertencem a este item -- foram incluidas por engano numa versao
--   anterior e inflavam o resultado em ~6,59 bi.
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 821192501
              AND o.INDEBITOCREDITO = 'D'
         THEN o.VALANCAMENTO ELSE 0 END)
    AS OUTROS_ING_INVEST,

-- ──────────────────────────────────────────────────────────────────────────
--  INVESTIMENTO — DESEMBOLSOS (JOIN NOTAEMPENHO para natureza)
-- ──────────────────────────────────────────────────────────────────────────

-- 1.04.01  AQUISICAO DE ATIVO NAO CIRCULANTE
--   NatDesp: 449051XX, 449052XX, 449061XX, 459061XX..459065XX
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (622920104,631400000,631820000)
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6)
                         IN ('449051','449052','449061',
                             '459061','459063','459064','459065')
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 632200000 AND 632299999
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6)
                         IN ('449051','449052','449061',
                             '459061','459063','459064','459065')
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS AQUIS_ATIVO_NC,

-- 1.04.02  CONCESSAO DE EMPRESTIMOS E FINANCIAMENTOS (NatDesp: 459066XX)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (622920104,631400000,631820000)
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6) = '459066'
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 632200000 AND 632299999
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6) = '459066'
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS CONCESSAO_EMPREST,

-- 1.04.03  OUTROS DESEMBOLSOS DE INVESTIMENTO
--   NatDesp: 449020,449030,449035,449037,449039,449040,449065,449092,449093,
--            446041,456783,459062,459067,459092,459093,455041,455042
--   + aplicacoes financeiras (MD 1141XXXXX/1144XXXXX/122310103/4 + MC 821192501)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (622920104,631400000,631820000)
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6)
                         IN ('449020','449030','449035','449037','449039',
                             '449040','449065','449092','449093','446041',
                             '456783','459062','459067','459092','459093',
                             '455041','455042')
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 632200000 AND 632299999
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6)
                         IN ('449020','449030','449035','449037','449039',
                             '449040','449065','449092','449093','446041',
                             '456783','459062','459067','459092','459093',
                             '455041','455042')
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 821192501
              AND o.INDEBITOCREDITO = 'C'
         THEN o.VALANCAMENTO ELSE 0 END)
    AS OUTROS_DESEMB_INVEST,

-- ──────────────────────────────────────────────────────────────────────────
--  FINANCIAMENTO
-- ──────────────────────────────────────────────────────────────────────────

-- 1.05.01  OPERACOES DE CREDITO (prefixo 21)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL = 621200000
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) = '21'
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
              AND SUBSTR(TO_CHAR(o.COCONTACORRENTE),1,2) = '21'
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS OPERACOES_CREDITO,

-- 1.06.02  OUTROS DESEMBOLSOS DE FINANCIAMENTO
--   NatDesp: 469071XX..469074XX
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (622920104,631400000,631820000)
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6)
                         IN ('469071','469072','469073','469074')
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL BETWEEN 632200000 AND 632299999
              AND EXISTS (
                  SELECT 1 FROM (
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2020.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2021.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2022.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2023.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2024.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL2025.NOTAEMPENHO
                      UNION ALL
                      SELECT NUNE, COUG, COGESTAO, CONATUREZA FROM MIL{ano}.NOTAEMPENHO
                  ) ne
                  WHERE  ne.NUNE     = SUBSTR(o.COCONTACORRENTE,1,11)
                    AND  ne.COUG     = o.COUGCONTAB
                    AND  ne.COGESTAO = o.COGESTAOCONTAB
                    AND  SUBSTR(ne.CONATUREZA,1,6)
                         IN ('469071','469072','469073','469074')
              )
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)
    AS OUTROS_DESEMB_FINANC,

-- ──────────────────────────────────────────────────────────────────────────
--  CAIXA E EQUIVALENTES (111XXXXXX)
--  NOTA: Caixa Inicial/Final NAO entram aqui -- sao buscados em consulta
--  separada contra a VIEW SALDOCONTABIL (ver funcao buscar_caixa()),
--  pois misturar essa fonte com os SUM(CASE...) de LANCAMENTOCONTABIL
--  no mesmo SELECT causa ORA-00937 no Oracle 11g.
-- ──────────────────────────────────────────────────────────────────────────
    0 AS CAIXA_INICIAL_PLACEHOLDER

FROM MIL{ano}.LANCAMENTOCONTABIL o
WHERE o.COCONTACONTABIL BETWEEN 100000000 AND 899999999
  {filtro_ug}
"""

# ─────────────────────────────────────────────────────────────────────────────
#  ESTRUTURA DE EXIBICAO DA DFC  (Anexo 15)
# ─────────────────────────────────────────────────────────────────────────────
#  (descricao, chave_sql, nivel, tipo)
#  tipo: 'titulo' | 'subtitulo' | 'item' | 'resultado' | 'total'
ESTRUTURA_DFC = [
    ("FLUXO DE CAIXA DAS ATIVIDADES OPERACIONAIS", None,                  1, "titulo"),
    ("Ingressos",                                   None,                  2, "subtitulo_calc"),
    ("  Impostos, Taxas e Contribuicoes de Melhoria","IMPOSTOS",           3, "item"),
    ("  Receita de Contribuicoes",                   "CONTRIBUICOES",      3, "item"),
    ("  Receita Patrimonial",                        "PATRIMONIAL",        3, "item"),
    ("  Receita Agropecuaria",                       "AGROPECUARIA",       3, "item"),
    ("  Receita Industrial",                         "INDUSTRIAL",         3, "item"),
    ("  Receita de Servicos",                        "SERVICOS",           3, "item"),
    ("  Remuneracao das Disponibilidades",           "REMUNERACAO_DISP",   3, "item"),
    ("  Transferencias Recebidas",                   "TRANSF_RECEBIDAS",   3, "item"),
    ("  Outras Receitas/Ingressos Operacionais",     "OUTRAS_REC_OPERAC",  3, "item"),
    ("Desembolsos",                                  None,                  2, "subtitulo_calc"),
    ("  Pessoal e Demais Despesas",                  "PESSOAL_DESPESAS",   3, "item"),
    ("  Juros e Encargos da Divida",                 "JUROS_ENCARGOS",     3, "item"),
    ("  Transferencias Concedidas",                  "TRANSF_CONCEDIDAS",  3, "item"),
    ("  Outros Desembolsos Operacionais",            "OUTROS_DESEMB_OPERAC",3,"item"),
    ("Fluxo de Caixa Liq. das Atividades Operacionais (I)", "FCO",         2, "resultado"),

    ("FLUXO DE CAIXA DAS ATIVIDADES DE INVESTIMENTO", None,               1, "titulo"),
    ("Ingressos",                                   None,                  2, "subtitulo_calc"),
    ("  Alienacao de Bens",                          "ALIENACAO_BENS",     3, "item"),
    ("  Amortizacao de Emprestimos e Financ.",       "AMORTIZ_EMPREST",    3, "item"),
    ("  Outros Ingressos de Investimento",           "OUTROS_ING_INVEST",  3, "item"),
    ("Desembolsos",                                  None,                  2, "subtitulo_calc"),
    ("  Aquisicao de Ativo Nao Circulante",          "AQUIS_ATIVO_NC",     3, "item"),
    ("  Concessao de Emprestimos e Financiamentos",  "CONCESSAO_EMPREST",  3, "item"),
    ("  Outros Desembolsos de Investimento",         "OUTROS_DESEMB_INVEST",3,"item"),
    ("Fluxo de Caixa Liq. das Ativ. de Investimento (II)", "FCI",          2, "resultado"),

    ("FLUXO DE CAIXA DAS ATIVIDADES DE FINANCIAMENTO", None,              1, "titulo"),
    ("Ingressos",                                   None,                  2, "subtitulo_calc"),
    ("  Operacoes de Credito",                       "OPERACOES_CREDITO",  3, "item"),
    ("Desembolsos",                                  None,                  2, "subtitulo_calc"),
    ("  Outros Desembolsos de Financiamento",        "OUTROS_DESEMB_FINANC",3,"item"),
    ("Fluxo de Caixa Liq. das Ativ. de Financiamento (III)", "FCF",        2, "resultado"),

    ("GERACAO LIQUIDA DE CAIXA E EQUIVALENTE (I+II+III)", "GERACAO_LIQ",  1, "total"),
    ("  Caixa e Equivalente de Caixa Inicial",       "CAIXA_INICIAL",      2, "item"),
    ("  Caixa e Equivalente de Caixa Final",         "CAIXA_FINAL",        2, "item"),
]

# Chaves para calcular subtotais derivados
INGRESSOS_OPERAC  = ["IMPOSTOS","CONTRIBUICOES","PATRIMONIAL","AGROPECUARIA",
                     "INDUSTRIAL","SERVICOS","REMUNERACAO_DISP",
                     "TRANSF_RECEBIDAS","OUTRAS_REC_OPERAC"]
DESEMB_OPERAC     = ["PESSOAL_DESPESAS","JUROS_ENCARGOS",
                     "TRANSF_CONCEDIDAS","OUTROS_DESEMB_OPERAC"]
INGRESSOS_INVEST  = ["ALIENACAO_BENS","AMORTIZ_EMPREST","OUTROS_ING_INVEST"]
DESEMB_INVEST     = ["AQUIS_ATIVO_NC","CONCESSAO_EMPREST","OUTROS_DESEMB_INVEST"]
INGRESSOS_FINANC  = ["OPERACOES_CREDITO"]
DESEMB_FINANC     = ["OUTROS_DESEMB_FINANC"]

# ─────────────────────────────────────────────────────────────────────────────
#  CONEXAO
# ─────────────────────────────────────────────────────────────────────────────
def conectar_oracle():
    print(f"  Inicializando Oracle Client: {INSTANT_CLIENT_DIR}")
    try:
        oracledb.init_oracle_client(lib_dir=INSTANT_CLIENT_DIR)
    except Exception as e:
        if "already been initialized" not in str(e): raise
    dsn = f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}"
    conn = oracledb.connect(user=DB_USER, password=DB_PASSWORD, dsn=dsn)
    print(f"  Conectado! Oracle {conn.version}")
    return conn

def buscar_caixa(conn, mes, ano, coug):
    """
    Busca Caixa Inicial e Caixa Final.

    Caixa Inicial: VIEW SALDOCONTABIL, INMES=0 (saldo de abertura do
    exercicio, ja encerrado -- fonte estavel).

    Caixa Final: NAO usa SALDOCONTABIL para o movimento do mes corrente.
    Descoberto via auditoria cruzada (mestre.py, Regra 4a/4b e X1b) que a
    VIEW e recalculada em lote (ciclo de consolidacao periodico) enquanto
    o mes corrente ainda esta aberto -- duas execucoes do mesmo script,
    ~1h48 de intervalo, devolveram Caixa Final diferindo em R$10,8 mi sem
    nenhuma mudanca de codigo. Prova: 3 leituras seguidas da view (a poucos
    segundos de distancia) ficam estaveis, mas o valor diverge do oficial
    (BP impresso) por R$12,17 mi distribuidos em ~18 contas 111XXXXXX, sem
    padrao localizavel -- nao e conta faltando, e a fonte inteira.

    A formula do BF (abertura da view + movimento via LANCAMENTOCONTABIL,
    a tabela transacional bruta) bateu EXATA (diferenca zero) contra o
    mesmo oficial. Reproduzida aqui para os dois anexos convergirem
    (elimina os erros cruzados X1b/4a/4b da auditoria).

    Separado do SQL principal porque misturar SUM() com os SUM(CASE...)
    de LANCAMENTOCONTABIL no mesmo SELECT causa ORA-00937 no Oracle 11g
    (conflito de agregacao em subquery escalar).
    """
    filtro_v = f"AND v.COUG = {coug}" if coug else ""
    filtro_o = f"AND o.COUG = {coug}" if coug else ""

    cur = conn.cursor()

    cur.execute(f"""
        SELECT SUM(v.VADEBITO - v.VACREDITO)
        FROM   MIL{ano}.SALDOCONTABIL v
        WHERE  v.INMES = 0
          AND  v.COCONTACONTABIL BETWEEN 111000000 AND 111999999
          {filtro_v}
    """)
    caixa_inicial = float(cur.fetchone()[0] or 0)

    cur.execute(f"""
        SELECT SUM(DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,
                                            'C',-o.VALANCAMENTO,0))
        FROM   MIL{ano}.LANCAMENTOCONTABIL o
        WHERE  o.INMES BETWEEN 1 AND {mes}
          AND  o.COCONTACONTABIL BETWEEN 111000000 AND 111999999
          {filtro_o}
    """)
    movimento = float(cur.fetchone()[0] or 0)
    caixa_final = caixa_inicial + movimento

    cur.close()
    return caixa_inicial, caixa_final


def buscar_dados(conn, mes, ano, coug):
    """
    Executa o SQL da DFC usando cursor nativo (compativel Oracle 11g thick mode).
    Filtro de natureza de receita/despesa via SUBSTR(COCONTACORRENTE,...).
    Caixa Inicial/Final sao buscados separadamente (ver buscar_caixa).
    """
    filtro = f"AND o.COUG = {coug}" if coug else ""
    sql = SQL_DFC.format(ano=ano, filtro_ug=filtro)

    # Substituir bind variable :mes por literal (evita warning pandas/oracledb)
    sql_exec = sql.replace(':mes', str(mes))

    print(f"  Executando SQL DFC em MIL{ano}...")
    cur = conn.cursor()
    cur.execute(sql_exec)
    cols = [d[0].upper() for d in cur.description]
    row  = cur.fetchone()
    cur.close()
    resultado = {k: float(v or 0) for k, v in zip(cols, row)}

    # Remove o placeholder e busca caixa real via SALDOCONTABIL
    resultado.pop('CAIXA_INICIAL_PLACEHOLDER', None)
    print(f"  Buscando Caixa Inicial/Final via SALDOCONTABIL...")
    caixa_inicial, caixa_final = buscar_caixa(conn, mes, ano, coug)
    resultado['CAIXA_INICIAL'] = caixa_inicial
    resultado['CAIXA_FINAL']   = caixa_final

    return resultado


# ─────────────────────────────────────────────────────────────────────────────
#  CALCULO DOS TOTAIS DERIVADOS
# ─────────────────────────────────────────────────────────────────────────────
def calcular(t):
    t = t.copy()
    # REMUNERACAO_DISP e exibida separada no relatorio mas o oficial
    # a inclui em 'Outras Receitas Operacionais'. Incorporamos ela em
    # OUTRAS_REC_OPERAC antes de somar ING_OPERAC para que a linha
    # separada exista (para exibicao) e o total I tambem fecha.
    t['OUTRAS_REC_OPERAC'] = t.get('OUTRAS_REC_OPERAC', 0) + t.get('REMUNERACAO_DISP', 0)
    t['REMUNERACAO_DISP']  = 0  # ja contabilizado em OUTRAS_REC_OPERAC
    t['ING_OPERAC']  = sum(t.get(k, 0) for k in INGRESSOS_OPERAC)
    t['DESEMB_OPER'] = sum(t.get(k, 0) for k in DESEMB_OPERAC)
    t['FCO']         = t['ING_OPERAC'] - t['DESEMB_OPER']

    t['ING_INVEST']  = sum(t.get(k, 0) for k in INGRESSOS_INVEST)
    t['DESEMB_INV']  = sum(t.get(k, 0) for k in DESEMB_INVEST)
    t['FCI']         = t['ING_INVEST'] - t['DESEMB_INV']

    t['ING_FINANC']  = sum(t.get(k, 0) for k in INGRESSOS_FINANC)
    t['DESEMB_FIN']  = sum(t.get(k, 0) for k in DESEMB_FINANC)
    t['FCF']         = t['ING_FINANC'] - t['DESEMB_FIN']

    t['GERACAO_LIQ'] = t['FCO'] + t['FCI'] + t['FCF']
    return t

def resolver_valor(t, desc, chave, tipo):
    """Retorna o valor correto para cada linha da estrutura."""
    if tipo == "titulo":     return None
    if tipo == "total":      return t.get(chave, 0)
    if tipo == "resultado":  return t.get(chave, 0)
    if tipo == "item":       return t.get(chave, 0)
    if tipo == "subtitulo_calc":
        # Calcula automaticamente pelo contexto
        if   "OPERAC" in desc.upper() or "INGRESSOS" in desc.upper():
            # primeiro subtitulo_calc apos OPERACIONAL = ingressos
            pass
        return None  # sera calculado na renderizacao
    return None

# ─────────────────────────────────────────────────────────────────────────────
#  RENDERIZACAO DA TABELA  (com subtotais calculados em linha)
# ─────────────────────────────────────────────────────────────────────────────
def linhas_para_exibir(t):
    """
    Retorna lista de (descricao, valor, tipo) pronta para renderizar.
    Calcula os subtotais de ingressos/desembolsos in-loco.
    """
    linhas = []
    # estado para calcular subtotais
    ctx_ing = None  # chave do total de ingressos corrente
    ctx_des = None  # chave do total de desembolsos corrente

    # Mapeamento de contexto para subtotais
    ctx_map = {
        "FLUXO DE CAIXA DAS ATIVIDADES OPERACIONAIS": ("ING_OPERAC","DESEMB_OPER"),
        "FLUXO DE CAIXA DAS ATIVIDADES DE INVESTIMENTO": ("ING_INVEST","DESEMB_INV"),
        "FLUXO DE CAIXA DAS ATIVIDADES DE FINANCIAMENTO": ("ING_FINANC","DESEMB_FIN"),
    }
    ctx_ing_key = None
    ctx_des_key = None
    flag_next_sub = None  # 'ing' ou 'des'

    for desc, chave, nivel, tipo in ESTRUTURA_DFC:
        desc_strip = desc.strip()

        if tipo == "titulo":
            # muda contexto
            ctx = ctx_map.get(desc_strip)
            if ctx:
                ctx_ing_key, ctx_des_key = ctx
                flag_next_sub = "ing"
            linhas.append((desc_strip, None, "titulo"))

        elif tipo == "subtitulo_calc":
            if flag_next_sub == "ing":
                linhas.append(("Ingressos", t.get(ctx_ing_key, 0), "subtitulo"))
                flag_next_sub = "des"
            else:
                linhas.append(("Desembolsos", t.get(ctx_des_key, 0), "subtitulo"))
                flag_next_sub = None

        elif tipo in ("item","resultado","total"):
            v = t.get(chave, 0) if chave else None
            linhas.append((desc_strip, v, tipo))

    return linhas

# ─────────────────────────────────────────────────────────────────────────────
#  EXCEL
# ─────────────────────────────────────────────────────────────────────────────
FMT_BRL = '#,##0.00'
VERDE_ESC  = "2E5C8A"   # cabecalho/total principal (era 2E75B6, agora alinhado a AZUL_TITU)
VERDE_MED  = "9DC3E6"   # subtotal de destaque (era 4472C4)
VERDE_CLR  = "DCE6F1"   # subtotal leve (era D9E2F3)
AZUL_TITU  = "2E5C8A"
CINZA_CLR  = "F5F9FD"   # linha normal (era EBF3FB)
CINZA_MED  = "DCE6F1"

def _fill(h): return PatternFill('solid', fgColor=h)
def _side(s='thin'): return Side(border_style=s, color='B8CCE4')
B_THIN = Border(left=_side(), right=_side(), top=_side(), bottom=_side())
B_MED  = Border(left=_side('medium'), right=_side('medium'),
                top=_side('medium'), bottom=_side('medium'))

def cel(ws, r, c, v='', bold=False, sz=9, bg=None, fg='000000',
        ha='left', brd=None, fmt=None, ind=0):
    x = ws.cell(row=r, column=c, value=v)
    x.font = Font(name='Arial', bold=bold, size=sz, color=fg)
    x.alignment = Alignment(horizontal=ha, vertical='center', indent=ind)
    if bg: x.fill = bg
    if brd: x.border = brd
    if fmt and v not in ('', None): x.number_format = fmt
    return x

def gerar_excel(t, mes, ano, ug_label, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'DFC Anexo 15'
    ws.sheet_view.showGridLines = False
    ws.page_setup.paperSize = 9  # A4

    # Larguras
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 5

    # ── Cabecalho ────────────────────────────────────────────────────────────
    ws.merge_cells('A1:D1')
    c = ws['A1']
    c.value = 'GOVERNO DO DISTRITO FEDERAL'
    c.font = Font(name='Arial', bold=True, size=13, color='FFFFFF')
    c.fill = _fill(VERDE_ESC)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 22

    ws.merge_cells('A2:D2')
    c = ws['A2']
    c.value = 'Demonstracao dos Fluxos de Caixa — Anexo 15'
    c.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    c.fill = _fill(VERDE_MED)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    ws.merge_cells('A3:D3')
    c = ws['A3']
    c.value = (f'Exercicio {ano}  |  Mes de Referencia: {mes:02d} - {MESES[mes]}'
               f'  |  {ug_label}  |  PSIAG550  |  Posicao: {datetime.now():%d/%m/%Y %H:%M}')
    c.font = Font(name='Arial', italic=True, size=8, color='444444')
    c.fill = _fill(CINZA_CLR)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[3].height = 12

    # ── Cabecalho da tabela ───────────────────────────────────────────────────
    R = 5
    ws.row_dimensions[R].height = 28
    cel(ws, R, 1, 'Cod.', bold=True, sz=9, bg=_fill(AZUL_TITU), fg='FFFFFF',
        ha='center', brd=B_THIN)
    cel(ws, R, 2, 'Descricao', bold=True, sz=9, bg=_fill(AZUL_TITU), fg='FFFFFF',
        ha='center', brd=B_THIN)
    cel(ws, R, 3, f'Exercicio Atual\n{ano} (jan-{MESES[mes]})', bold=True, sz=9,
        bg=_fill(AZUL_TITU), fg='FFFFFF', ha='center', brd=B_THIN)
    ws[f'C{R}'].alignment = Alignment(horizontal='center', vertical='center',
                                       wrap_text=True)
    cel(ws, R, 4, 'Nota', bold=True, sz=9, bg=_fill(AZUL_TITU), fg='FFFFFF',
        ha='center', brd=B_THIN)

    # ── Linhas de dados ───────────────────────────────────────────────────────
    linhas = linhas_para_exibir(t)
    cod_map = {
        "FLUXO DE CAIXA DAS ATIVIDADES OPERACIONAIS":  "1.00",
        "Ingressos":  "",
        "Desembolsos": "",
        "Fluxo de Caixa Liq. das Atividades Operacionais (I)": "I",
        "FLUXO DE CAIXA DAS ATIVIDADES DE INVESTIMENTO": "2.00",
        "Fluxo de Caixa Liq. das Ativ. de Investimento (II)": "II",
        "FLUXO DE CAIXA DAS ATIVIDADES DE FINANCIAMENTO": "3.00",
        "Fluxo de Caixa Liq. das Ativ. de Financiamento (III)": "III",
        "GERACAO LIQUIDA DE CAIXA E EQUIVALENTE (I+II+III)": "=",
    }

    row = R + 1
    for desc, valor, tipo in linhas:
        ws.row_dimensions[row].height = 14

        if tipo == "titulo":
            bg = _fill(VERDE_ESC); fg = 'FFFFFF'; bold = True; ind = 0
        elif tipo == "subtitulo":
            bg = _fill(VERDE_CLR); fg = '000000'; bold = True; ind = 1
        elif tipo == "resultado":
            bg = _fill(VERDE_MED); fg = 'FFFFFF'; bold = True; ind = 1
        elif tipo == "total":
            bg = _fill(AZUL_TITU); fg = 'FFFFFF'; bold = True; ind = 0
        else:  # item
            bg = _fill(CINZA_CLR) if row % 2 == 0 else _fill('FAFCFE')
            fg = '000000'; bold = False; ind = 2

        cod = cod_map.get(desc, "")
        cel(ws, row, 1, cod, bold=bold, bg=bg, fg=fg, ha='center', brd=B_THIN, sz=8)
        cel(ws, row, 2, desc, bold=bold, bg=bg, fg=fg, ha='left', brd=B_THIN,
            sz=9, ind=ind)
        cel(ws, row, 3, valor, bold=bold, bg=bg, fg=fg, ha='right', brd=B_THIN,
            sz=9, fmt=FMT_BRL)
        cel(ws, row, 4, '', bg=bg, brd=B_THIN)
        row += 1

    # Linha de diferenca de conferencia (ajuste/gap da identidade I+II+III)
    ws.row_dimensions[row + 1].height = 13
    cel(ws, row + 1, 2, 'Ajuste (Geracao Liq. - Var. Caixa):', sz=8, ha='right')
    dif = t.get('GERACAO_LIQ', 0) - (t.get('CAIXA_FINAL', 0) - t.get('CAIXA_INICIAL', 0))
    cel(ws, row + 1, 3, dif, sz=8, ha='right', fmt=FMT_BRL)

    # Rodape
    ws.row_dimensions[row + 3].height = 12
    cel(ws, row + 3, 2, f'Emitido por: Python/Oracle  |  {datetime.now():%d/%m/%Y %H:%M:%S}',
        sz=7, ha='right')

    # Aba de dados brutos
    ws2 = wb.create_sheet('Dados SQL Brutos')
    for j, (k, v) in enumerate(t.items(), 1):
        ws2.cell(row=1, column=j, value=k).font = Font(bold=True, name='Arial', size=9)
        c2 = ws2.cell(row=2, column=j, value=v)
        if isinstance(v, float): c2.number_format = FMT_BRL
        ws2.column_dimensions[ws2.cell(row=1, column=j).column_letter].width = 22

    wb.save(output_path)
    print(f"  Excel salvo: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  PDF
# ─────────────────────────────────────────────────────────────────────────────
def _brl(v):
    if v is None: return ""
    neg = v < 0
    s = f"{abs(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    return f"({s})" if neg else s

def gerar_pdf(t, mes, ano, ug_label, output_path):
    if not REPORTLAB_OK:
        print("  PDF ignorado: instale reportlab -> pip install reportlab")
        return

    doc = SimpleDocTemplate(str(output_path), pagesize=A4,
          leftMargin=15*mm, rightMargin=15*mm,
          topMargin=15*mm, bottomMargin=15*mm,
          title=f"DFC GDF {ano} - {MESES[mes]}",
          author="SIAC/SIGGO - GDF")

    styles = getSampleStyleSheet()
    VERDE_RL  = colors.HexColor("#2E5C8A")
    VMED_RL   = colors.HexColor("#9DC3E6")
    VCLR_RL   = colors.HexColor("#DCE6F1")
    AZUL_RL   = colors.HexColor("#2E5C8A")
    CINZA_RL  = colors.HexColor("#F5F9FD")

    def st_h(name, **kw):
        kw.setdefault('fontName', 'Helvetica')
        return ParagraphStyle(name, **kw)
    st_tit   = st_h("tit",  fontName='Helvetica-Bold', fontSize=13,
                    textColor=colors.white, alignment=TA_CENTER)
    st_sub   = st_h("sub",  fontName='Helvetica-Bold', fontSize=10,
                    textColor=colors.white, alignment=TA_CENTER)
    st_meta  = st_h("meta", fontSize=8, textColor=colors.HexColor("#444444"),
                    alignment=TA_CENTER)
    st_bold8 = st_h("b8",   fontName='Helvetica-Bold', fontSize=8)
    st_norm8 = st_h("n8",   fontSize=8)
    st_r8    = st_h("r8",   fontSize=8, alignment=TA_RIGHT)

    story = []

    def hdr_table(title, sub, meta):
        data = [[Paragraph(title, st_tit)],
                [Paragraph(sub,   st_sub)],
                [Paragraph(meta,  st_meta)]]
        tb = Table(data, colWidths=[180*mm])
        tb.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(0,0), VERDE_RL),
            ('BACKGROUND', (0,1),(0,1), VMED_RL),
            ('BACKGROUND', (0,2),(0,2), colors.HexColor("#F5F5F5")),
            ('TOPPADDING',    (0,0),(-1,-1), 4),
            ('BOTTOMPADDING', (0,0),(-1,-1), 4),
            ('LEFTPADDING',   (0,0),(-1,-1), 6),
        ]))
        return tb

    story.append(hdr_table(
        "Governo do Distrito Federal",
        "Demonstracao dos Fluxos de Caixa — Anexo 15",
        (f"Exercicio {ano}  |  Mes: {mes:02d} - {MESES[mes]}  |  "
         f"{ug_label}  |  PSIAG550  |  "
         f"Posicao: {datetime.now():%d/%m/%Y %H:%M}")
    ))
    story.append(Spacer(1, 4*mm))

    # Monta tabela de dados
    W_DESC = 130*mm
    W_VAL  = 45*mm
    W_NOTA = 10*mm

    header = [
        Paragraph("<b>Descricao</b>",
                  ParagraphStyle("hd", fontName='Helvetica-Bold', fontSize=8,
                                 textColor=colors.white, alignment=TA_CENTER)),
        Paragraph("<b>Exercicio Atual</b>",
                  ParagraphStyle("hv", fontName='Helvetica-Bold', fontSize=8,
                                 textColor=colors.white, alignment=TA_CENTER)),
        Paragraph("<b>Nota</b>",
                  ParagraphStyle("hn", fontName='Helvetica-Bold', fontSize=8,
                                 textColor=colors.white, alignment=TA_CENTER)),
    ]
    rows = [header]
    tipos_list = []

    linhas = linhas_para_exibir(t)
    for desc, valor, tipo in linhas:
        fn = 'Helvetica-Bold' if tipo in ('titulo','subtitulo','resultado','total') else 'Helvetica'
        ind = {
            "titulo": 0, "subtitulo": 6, "resultado": 6,
            "total": 0, "item": 14
        }.get(tipo, 0)
        p_desc = Paragraph(desc, ParagraphStyle("d", fontName=fn, fontSize=8,
                           leftIndent=ind, alignment=TA_LEFT))
        p_val  = Paragraph(_brl(valor),
                           ParagraphStyle("v", fontName=fn, fontSize=8,
                                          alignment=TA_RIGHT))
        rows.append([p_desc, p_val, ""])
        tipos_list.append(tipo)

    col_widths = [W_DESC, W_VAL, W_NOTA]
    tab = Table(rows, colWidths=col_widths, repeatRows=1)

    ts = [
        ('BACKGROUND', (0,0), (-1,0), AZUL_RL),
        ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING',   (0,0), (-1,-1), 3),
        ('RIGHTPADDING',  (0,0), (-1,-1), 3),
    ]
    for i, tipo in enumerate(tipos_list):
        r = i + 1
        if tipo == "titulo":
            ts += [('BACKGROUND', (0,r), (-1,r), VERDE_RL),
                   ('TEXTCOLOR',  (0,r), (-1,r), colors.white),
                   ('FONTNAME',   (0,r), (-1,r), 'Helvetica-Bold')]
        elif tipo == "subtitulo":
            ts += [('BACKGROUND', (0,r), (-1,r), VCLR_RL),
                   ('FONTNAME',   (0,r), (-1,r), 'Helvetica-Bold')]
        elif tipo == "resultado":
            ts += [('BACKGROUND', (0,r), (-1,r), VMED_RL),
                   ('TEXTCOLOR',  (0,r), (-1,r), colors.white),
                   ('FONTNAME',   (0,r), (-1,r), 'Helvetica-Bold'),
                   ('LINEABOVE',  (0,r), (-1,r), 0.8, VERDE_RL)]
        elif tipo == "total":
            ts += [('BACKGROUND', (0,r), (-1,r), AZUL_RL),
                   ('TEXTCOLOR',  (0,r), (-1,r), colors.white),
                   ('FONTNAME',   (0,r), (-1,r), 'Helvetica-Bold'),
                   ('LINEABOVE',  (0,r), (-1,r), 1.5, AZUL_RL)]
        else:
            bg = CINZA_RL if i % 2 == 0 else colors.white
            ts.append(('BACKGROUND', (0,r), (-1,r), bg))

    tab.setStyle(TableStyle(ts))
    story.append(tab)
    story.append(Spacer(1, 3*mm))

    geracao_pdf = t.get('GERACAO_LIQ', 0)
    varcaixa_pdf = t.get('CAIXA_FINAL', 0) - t.get('CAIXA_INICIAL', 0)
    gap_pdf = geracao_pdf - varcaixa_pdf
    story.append(Paragraph(
        f"<b>Conferencia (I+II+III):</b>  Geracao Liq. {_brl(geracao_pdf)} "
        f"vs Var. Caixa {_brl(varcaixa_pdf)}  |  Ajuste: {_brl(gap_pdf)}",
        ParagraphStyle("conf", fontName='Helvetica', fontSize=8,
                       textColor=colors.HexColor("#555555"))))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        f"Emitido por: Python/Oracle  |  Emitido em: {datetime.now():%d/%m/%Y %H:%M:%S}",
        ParagraphStyle("rod", fontName='Helvetica', fontSize=7,
                       textColor=colors.HexColor("#888888"), alignment=TA_RIGHT)))

    doc.build(story)
    print(f"  PDF salvo: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  DIAGNOSTICO  (identica ao BF)
# ─────────────────────────────────────────────────────────────────────────────
def diagnostico(conn, ano):
    print("\n" + "="*64)
    print("  DIAGNOSTICO - ESTRUTURA DO BANCO PARA DFC")
    print("="*64)

    print(f"\n[1] Contas de caixa (111XXXXXX) por INMES em MIL{ano}:")
    try:
        q = f"""SELECT o.INMES, COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)) SALDO
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL BETWEEN 111000000 AND 111999999
                GROUP BY o.INMES ORDER BY o.INMES"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    INMES={int(r['INMES']):>3}  qtd={int(r['QTD']):>9}  saldo={float(r['SALDO'] or 0):>22,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2] Contas de receita (621XXXXXX) - top NatRec em MIL{ano}:")
    try:
        q = f"""SELECT TRUNC(o.CONATURREC/10000) PREFIXO, COUNT(*) QTD,
                    SUM(DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)) VALOR
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.COCONTACONTABIL = 621200000
                  AND o.INMES BETWEEN 1 AND 13
                GROUP BY TRUNC(o.CONATURREC/10000)
                ORDER BY 3 DESC"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    NatRec_prefix={int(r['PREFIXO']):>4}  qtd={int(r['QTD']):>8}  valor={float(r['VALOR'] or 0):>22,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[3] Colunas de MIL{ano}.LANCAMENTOCONTABIL:")
    try:
        q = f"""SELECT column_name, data_type FROM all_tab_columns
                WHERE owner='MIL{ano}' AND table_name='LANCAMENTOCONTABIL'
                ORDER BY column_id"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    {r['COLUMN_NAME']:<25} {r['DATA_TYPE']}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print("="*64 + "\n")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(description='Demonstracao dos Fluxos de Caixa - GDF (Anexo 15)')
    p.add_argument('--mes',     type=int, default=5,
                   help='Mes de referencia 1-12 (posicao acumulada jan..mes)')
    p.add_argument('--ano',     type=int, default=2026)
    p.add_argument('--ug',      type=str, default=None,
                   help='Codigo da Unidade Gestora (omitir = Consolidado)')
    p.add_argument('--saida',   type=str, default=None,
                   help='Caminho base dos arquivos de saida (sem extensao)')
    p.add_argument('--formato', choices=['ambos','excel','pdf'], default='ambos')
    p.add_argument('--diag',    action='store_true',
                   help='Executa diagnostico do banco e encerra')
    a = p.parse_args()

    if not 1 <= a.mes <= 12:
        print("ERRO: --mes deve estar entre 1 e 12."); sys.exit(1)

    ug_label = f'UG: {a.ug}' if a.ug else 'Consolidado'
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome = (f'DFC_GDF_{a.ano}_{a.mes:02d}'
            + (f'_UG{a.ug}' if a.ug else '_Consolidado')
            + f'_{timestamp}')
    base = (Path(a.saida).with_suffix('') if a.saida else OUTPUT_DIR / nome)
    saida_xlsx = base.with_suffix('.xlsx')
    saida_pdf  = base.with_suffix('.pdf')

    print(f"\n{'='*60}")
    print(f"  DEMONSTRACAO DOS FLUXOS DE CAIXA — Anexo 15")
    print(f"  {MESES[a.mes]}/{a.ano}  |  {ug_label}")
    print(f"  Posicao acumulada: Janeiro a {MESES[a.mes]}")
    print(f"{'='*60}")

    print("\n[1/4] Conectando ao Oracle...")
    conn = conectar_oracle()

    if a.diag:
        diagnostico(conn, a.ano)
        conn.close()
        return

    print("\n[2/4] Buscando dados DFC...")
    t_raw = buscar_dados(conn, a.mes, a.ano, a.ug)
    conn.close()

    print("\n[3/4] Calculando totais derivados...")
    t = calcular(t_raw)

    print("\n[4/4] Gerando arquivos...")
    if a.formato in ('ambos', 'excel'):
        gerar_excel(t, a.mes, a.ano, ug_label, saida_xlsx)
    if a.formato in ('ambos', 'pdf'):
        gerar_pdf(t, a.mes, a.ano, ug_label, saida_pdf)

    print(f"\n{'─'*60}")
    print("  CONFERENCIA RAPIDA:")
    print(f"{'─'*60}")
    items_conf = [
        ("IMPOSTOS/TAXAS/CONTRIB",   "IMPOSTOS"),
        ("RECEITA CONTRIBUICOES",    "CONTRIBUICOES"),
        ("RECEITA PATRIMONIAL",      "PATRIMONIAL"),
        ("RECEITA AGROPECUARIA",     "AGROPECUARIA"),
        ("RECEITA INDUSTRIAL",       "INDUSTRIAL"),
        ("RECEITA DE SERVICOS",      "SERVICOS"),
        ("REMUNERACAO DISPON.",      "REMUNERACAO_DISP"),
        ("TRANSF. RECEBIDAS",        "TRANSF_RECEBIDAS"),
        ("OUTRAS REC. OPERAC.",      "OUTRAS_REC_OPERAC"),
        ("PESSOAL E DESPESAS",       "PESSOAL_DESPESAS"),
        ("JUROS E ENCARGOS",         "JUROS_ENCARGOS"),
        ("TRANSF. CONCEDIDAS",       "TRANSF_CONCEDIDAS"),
        ("OUTROS DESEMB. OPERAC.",   "OUTROS_DESEMB_OPERAC"),
        ("FCO (I)",                  "FCO"),
        ("ALIENACAO DE BENS",        "ALIENACAO_BENS"),
        ("AMORTIZ. EMPRESTIMOS",     "AMORTIZ_EMPREST"),
        ("OUTROS ING. INVEST.",      "OUTROS_ING_INVEST"),
        ("AQUIS. ATIVO NAO CIRC.",   "AQUIS_ATIVO_NC"),
        ("CONCESSAO EMPREST.",       "CONCESSAO_EMPREST"),
        ("OUTROS DESEMB. INVEST.",   "OUTROS_DESEMB_INVEST"),
        ("FCI (II)",                 "FCI"),
        ("OPERACOES DE CREDITO",     "OPERACOES_CREDITO"),
        ("OUTROS DESEMB. FINANC.",   "OUTROS_DESEMB_FINANC"),
        ("FCF (III)",                "FCF"),
        ("GERACAO LIQUIDA",          "GERACAO_LIQ"),
        ("CAIXA INICIAL",            "CAIXA_INICIAL"),
        ("CAIXA FINAL",              "CAIXA_FINAL"),
    ]
    for lbl, key in items_conf:
        print(f"    {lbl:<28}: {t.get(key, 0):>22,.2f}")
    print(f"{'─'*60}")
    # Identidade fundamental da DFC: I+II+III deve ser igual a Caixa Final - Inicial
    geracao   = t.get('GERACAO_LIQ', 0)
    var_caixa = t.get('CAIXA_FINAL', 0) - t.get('CAIXA_INICIAL', 0)
    gap       = geracao - var_caixa
    print(f"    {'Geracao Liq. (I+II+III)':<28}: {geracao:>22,.2f}")
    print(f"    {'Var. Caixa (Final-Inicial)':<28}: {var_caixa:>22,.2f}")
    print(f"    {'AJUSTE / GAP':<28}: {gap:>22,.2f}")
    print(f"{'─'*60}")
    print(f"    Nota: o PSIAG550 oficial tambem apresenta um gap nesta")
    print(f"    identidade (ajuste de conciliacao / OFSS). O gap so deve")
    print(f"    refletir esse ajuste oficial + eventuais residuos em aberto.")
    print(f"{'─'*60}")
    print(f"\n  Concluido em {datetime.now():%H:%M:%S}\n")

if __name__ == '__main__':
    main()