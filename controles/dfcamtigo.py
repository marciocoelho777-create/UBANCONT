# -*- coding: utf-8 -*-
"""
Controles do Demonstrativo de Fluxos de Caixa (Anexo 15).
Usa o SQL_DFC do dfc.py original com as correções:
  - 449093 removido de TRANSF_CONCEDIDAS (pertence a OUTROS_DESEMB_INVEST)
  - REMUNERACAO_DISP absorvida em OUTRAS_REC_OPERAC
Caixa via SALDOCONTABIL 111XXXXXX (igual ao BP).
"""
from __future__ import annotations
from decimal import Decimal
from . import (Achado, query_one, D,
               achado_ok, achado_erro, checa_gap)

# REMUNERACAO_DISP removida desta lista: seu valor é absorvido em
# OUTRAS_REC_OPERAC antes do cálculo de I (evita dupla contagem).
INGRESSOS_OPERAC = ["IMPOSTOS","CONTRIBUICOES","PATRIMONIAL","AGROPECUARIA",
                    "INDUSTRIAL","SERVICOS",
                    "TRANSF_RECEBIDAS","OUTRAS_REC_OPERAC"]
DESEMB_OPERAC    = ["PESSOAL_DESPESAS","JUROS_ENCARGOS",
                    "TRANSF_CONCEDIDAS","OUTROS_DESEMB_OPERAC"]
INGRESSOS_INVEST = ["ALIENACAO_BENS","AMORTIZ_EMPREST","OUTROS_ING_INVEST"]
DESEMB_INVEST    = ["AQUIS_ATIVO_NC","CONCESSAO_EMPREST","OUTROS_DESEMB_INVEST"]
INGRESSOS_FINANC = ["OPERACOES_CREDITO"]
DESEMB_FINANC    = ["OUTROS_DESEMB_FINANC"]

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
--   NatDesp: 3120XXXX, 3191XXXX, 3320XXXX, 3340XXXX, 3350XXXX, 3360XXXX,
--            336783XX, 3370XXXX, 3380XXXX, 3391XXXX, 4420XXXX, 4440XXXX,
--            4450XXXX, 4470XXXX, 4480XXXX, 449093XX, 4491XXXX, 4591XXXX
--   (CORRIGIDO conforme planilha oficial de equacoes de balanco)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (622920104,631400000,631820000)
              AND EXISTS (
                  SELECT 1 FROM (
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
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND :mes
              AND o.COCONTACONTABIL IN (237110301,237120301,237210301,237220301)
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

SQL_CAIXA = """
SELECT
    SUM(CASE WHEN v.INMES = 0
              AND v.COCONTACONTABIL BETWEEN 111000000 AND 111999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END) AS CAIXA_INI,
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 111000000 AND 111999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END) AS CAIXA_FIN
FROM MIL{ano}.SALDOCONTABIL v
"""


def extrair(conn, mes, ano):
    """Executa SQL_DFC + SQL_CAIXA."""
    # Substituir :mes por literal (mesmo padrão do dfc.py original que evita
    # problemas com Oracle 11g thick mode e múltiplas ocorrências do bind)
    sql = SQL_DFC.format(ano=ano, filtro_ug="").replace(":mes", str(mes))
    cur = conn.cursor()
    cur.execute(sql)
    cols = [d[0].upper() for d in cur.description]
    row  = cur.fetchone()
    cur.close()
    if row is None:
        t = {c: D(0) for c in cols}
    else:
        t = {}
        for c, v in zip(cols, row):
            try:
                t[c] = D(str(v)) if v is not None else D(0)
            except Exception:
                t[c] = D(0)
    t.update(query_one(conn, SQL_CAIXA.format(mes=mes, ano=ano)))
    return t


def auditar(conn, mes, ano):
    t = extrair(conn, mes, ano)
    achados = []

    # REMUNERACAO_DISP: zeramos para nao somar duas vezes (ja esta em OUTRAS_REC_OPERAC)
    t["REMUNERACAO_DISP"] = D(0)


    I   = (sum(t.get(k, D(0)) for k in INGRESSOS_OPERAC)
           - sum(t.get(k, D(0)) for k in DESEMB_OPERAC))
    II  = (sum(t.get(k, D(0)) for k in INGRESSOS_INVEST)
           - sum(t.get(k, D(0)) for k in DESEMB_INVEST))
    III = (sum(t.get(k, D(0)) for k in INGRESSOS_FINANC)
           - sum(t.get(k, D(0)) for k in DESEMB_FINANC))
    ger = I + II + III
    var = t.get("CAIXA_FIN", D(0)) - t.get("CAIXA_INI", D(0))
    gap = ger - var

    t["FCO"] = I; t["FCI"] = II; t["FCF"] = III
    t["GERACAO"] = ger; t["VAR_CAIXA"] = var; t["GAP_X2"] = gap

    if checa_gap(gap, Decimal("50.00")):
        achados.append(achado_ok("DFC","DFC-01",
            "Geracao Liquida (I+II+III) = Variacao do Caixa",
            f"Geracao {ger:,.2f}  =  Variacao {var:,.2f}"))
    else:
        achados.append(achado_erro("DFC","DFC-01",
            "Geracao Liquida != Variacao do Caixa",
            f"Gap: {gap:,.2f}  |  Geracao {ger:,.2f}  vs  Var.Caixa {var:,.2f}",
            valor=gap))

    for nome, val, cod in [("Caixa Inicial", t.get("CAIXA_INI",D(0)), "DFC-02a"),
                            ("Caixa Final",   t.get("CAIXA_FIN",D(0)), "DFC-02b")]:
        if val < 0:
            achados.append(achado_erro("DFC", cod, f"{nome} negativo", f"{val:,.2f}"))
        else:
            achados.append(achado_ok("DFC", cod, f"{nome} positivo", f"{val:,.2f}"))

    return achados, t
