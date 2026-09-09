# -*- coding: utf-8 -*-
"""
Controles do Balanço Orçamentário (Anexo 12).
Contas baseadas no bo.py original (GDF/SIGGO):
  Dotação Inicial   : 522110000..522119999 (SD)
  Dotação Atualizada: 522110000..522129999 (SD) - 522150000..522159999 (SC) + 522190000..522199999 (SD)
  Empenhada         : 622130000..622139999 (SC)
  Liquidada         : 622130300, 622130400, 622130700 (SC)
  Paga              : 622920104 (SC)
  Pgtos RPÑP        : 631400000 + 631820000 (SC)
  Pgtos RPP         : 6322XXXXX (SC)

CORREÇÃO 12/08/2026 — reclassificação de 6322XXXXX
--------------------------------------------------
A versão anterior somava 632200000..632299999 dentro de DESPESA_PAGA, espelhando
o mesmo erro do bf.py. A Lista de Equações de Balanço Financeiro põe 6322XXXXX no
item 2.04.02 (Pagamentos Extraorçamentários de Restos a Pagar Processados) e a
Despesa Orçamentária (item 2.01.xx) usa apenas 622920104 no exercício corrente.

Medido em jul/2026:
    622920104 sozinho    21.699.922.160,81
    6322XXXXX          +  2.283.574.211,13
                       = 23.983.496.371,94   (o valor antes reportado como Paga)

Com a Paga inflada, ela superava a Liquidada (23.617.140.264,77) em
366.356.107,17 e o BO-01b acusava todo mês. Reclassificada, a Paga fica
R$ 1,92 bi ABAIXO da Liquidada — sobra de liquidado ainda não pago, que é o
comportamento normal. O alerta não era estrutural nem RPP pago diretamente.

A mesma remoção foi feita no bf.py; sem ela o cruzamento R8 quebra.
O cruzamentos.py não muda: lê pelas chaves DESPESA_PAGA/PAG_RPNP/PAG_RPP.
"""
from __future__ import annotations
from . import (Achado, query_one, query_all, D,
               achado_ok, achado_erro, achado_alerta, checa_gap)

SQL_BO = """
SELECT
    -- Receita Realizada (SC 621200000 + 612000000, INMES 1..mes)
    -- 612000000 = Receita de Capital (mesmo que BF para que R7 feche)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL IN (621200000, 612000000)
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0) ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0) ELSE 0 END)
    AS RECEITA_REALIZADA,

    -- Dotação Inicial (SD 522110000..522119999, INMES 1..mes)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 522110000 AND 522119999
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0) ELSE 0 END)
    AS DOT_INICIAL,

    -- Dotação Atualizada (522110..22129 SD - 522150..22159 SC + 522190..22199 SD)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 522110000 AND 522129999
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0) ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 522150000 AND 522159999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0) ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 522190000 AND 522199999
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0) ELSE 0 END)
    AS DOT_ATUALIZADA,

    -- Empenhada (SC 622130000..622139999)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 622130000 AND 622139999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0) ELSE 0 END)
    AS DESP_EMPENHADA,

    -- Liquidada (SC 622130300/400/700)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL IN (622130300,622130400,622130700)
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0) ELSE 0 END)
    AS DESP_LIQUIDADA,

    -- Paga (SC 622920104) — 6322XXXXX NÃO entra aqui: item 2.04.02 da Lista
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL = 622920104
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0) ELSE 0 END)
    AS DESPESA_PAGA,

    -- Pagamentos RPÑP (SC 631400000 + 631820000) — item 2.04.01
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL IN (631400000,631820000)
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0) ELSE 0 END)
    AS PAG_RPNP,

    -- Pagamentos RPP (SC 6322XXXXX) — item 2.04.02
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 632200000 AND 632299999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0) ELSE 0 END)
    AS PAG_RPP

FROM MIL{ano}.LANCAMENTOCONTABIL o
"""


# ─────────────────────────────────────────────────────────────────────────────
#  BO-03 — classificação orçamentária ausente na BALANCOGERAL
# ─────────────────────────────────────────────────────────────────────────────
# Os demonstrativos oficiais (PSIAG550/560) são montados a partir da tabela
# MIL{ano}.BALANCOGERAL, que já traz a classificação resolvida em colunas
# próprias (CONATUREZA, COUO, COFONTE, COFUNCAO...). O diagnóstico e os
# demonstrativos do mestre leem LANCAMENTOCONTABIL e resolvem a natureza
# buscando a nota de empenho — caminho diferente, e mais robusto.
#
# Quando um registro entra na BALANCOGERAL com CONATUREZA = 0 e COUO = 0, o
# valor existe na contabilidade mas não pertence a nenhum item do Balanço
# Orçamentário: os itens 2.01.01.0X são filtrados por GND 1..9, e sem natureza
# o valor não cai em nenhum deles. Some da abertura por GND e por UO.
#
# Casos medidos em 2026 (MIL2026, INMES 1..8):
#     mês 5   622130300  Despesa Liquidada       129.929,04
#     mês 6   622130300  Despesa Liquidada     1.922.061,96
#     mês 8   622920104  Despesa Paga          1.134.760,68
#                                             ──────────────
#                                               3.186.751,68
#
# O caso do mês 8 (documentos 2026NL05656 e 2026OB47766, UG 200101, NE
# 2026NE00023) explica exatamente o gap pelo qual o próprio PSIAG550 não
# fecha: I+II+III − variação do caixa = 1.134.760,68. A NOTAEMPENHO estava
# correta (CONATUREZA = '339039'); o zero está só na BALANCOGERAL.
#
# ATENÇÃO ao montar a query: as contas 621XXXXXX são de RECEITA e têm
# CONATUREZA = 0 por construção. Incluí-las gera centenas de falsos positivos.
#
# FAIXA DE CONTAS — extraída da Lista de Equações do Balanço Orçamentário
# (14/08/2026), filtrando as linhas com GND preenchido. São 34 contas ativas,
# todas com GND 1..6 ou 1..7+9. Um registro com CONATUREZA/COUO zerados em
# qualquer uma delas não casa com nenhum item e desaparece do demonstrativo:
#   Dotação        52211XXXX, 52212XXXX, 52215XXXX, 52219XXXX, 522120100,
#                  522120201/202/301, 522150100/200/300
#   Restos a Pagar 531100000, 531200000, 532100000, 532200000
#   Empenh/Liquid  62213XXXX, 622130300/400/500/600/700
#   Paga           622920104
#   Pgtos RP       631300000, 631400000, 631810000, 631820000, 631900000,
#                  632210100/200/300/400, 632900000
# A faixa anterior (622000000..632999999) deixava de fora dotação e restos a
# pagar — um zero ali sumiria da Dotação Atualizada sem ninguém perceber.
# A 622130300 (coluna 4, Liquidada) está entre elas: confirma que os
# R$ 2.051.991,00 de maio e junho também somem do Balanço Orçamentário.
SQL_BO03 = """
SELECT b.INMES, b.COCONTACONTABIL,
       SUM(b.VACREDITO - b.VADEBITO) AS VLR,
       COUNT(*)                      AS QTD
FROM   MIL{ano}.BALANCOGERAL b
WHERE  b.INMES BETWEEN 1 AND {mes}
  AND  ( b.COCONTACONTABIL BETWEEN 522110000 AND 522199999   -- dotacao
      OR b.COCONTACONTABIL BETWEEN 531100000 AND 532299999   -- restos a pagar
      OR b.COCONTACONTABIL BETWEEN 622130000 AND 622139999   -- empenhada/liquidada
      OR b.COCONTACONTABIL = 622920104                       -- paga
      OR b.COCONTACONTABIL BETWEEN 631300000 AND 631999999   -- pgtos RP
      OR b.COCONTACONTABIL BETWEEN 632210000 AND 632999999 ) -- pgtos RP
  AND (b.CONATUREZA IS NULL OR b.CONATUREZA = 0
    OR b.COUO       IS NULL OR b.COUO       = 0)
GROUP  BY b.INMES, b.COCONTACONTABIL
HAVING ABS(SUM(b.VACREDITO - b.VADEBITO)) > 0.01
ORDER  BY 3 DESC
"""


def extrair(conn, mes, ano):
    return query_one(conn, SQL_BO.format(mes=mes, ano=ano))


def auditar(conn, mes, ano):
    t = extrair(conn, mes, ano)
    achados = []

    # BO-01: Empenhada >= Liquidada e Empenhada >= Paga
    liq_ok = t["DESP_LIQUIDADA"] <= t["DESP_EMPENHADA"] + D("0.02")
    pag_ok = t["DESPESA_PAGA"] <= t["DESP_EMPENHADA"] + D("0.02")
    if liq_ok and pag_ok:
        achados.append(achado_ok("BO","BO-01","Empenhada ≥ Liquidada e Empenhada ≥ Paga",
            f"Empenhada {t['DESP_EMPENHADA']:,.2f}  "
            f"Liquidada {t['DESP_LIQUIDADA']:,.2f}  "
            f"Paga {t['DESPESA_PAGA']:,.2f}", valor=t["DESP_EMPENHADA"]))
    else:
        excesso = max(t["DESP_LIQUIDADA"] - t["DESP_EMPENHADA"],
                      t["DESPESA_PAGA"] - t["DESP_EMPENHADA"])
        achados.append(achado_erro("BO","BO-01","Empenhada < Liquidada ou Empenhada < Paga",
            f"Empenhada {t['DESP_EMPENHADA']:,.2f}  "
            f"Liquidada {t['DESP_LIQUIDADA']:,.2f}  "
            f"Paga {t['DESPESA_PAGA']:,.2f}", valor=excesso))

    # BO-01b: Paga > Liquidada.
    # Com 6322XXXXX classificado corretamente (item 2.04.02), a Despesa Paga do
    # exercício corrente é sempre <= Liquidada: paga-se o que foi liquidado, e a
    # sobra de liquidado não pago é normal. Se este controle disparar, NÃO é
    # estrutural nem "RPP pago diretamente" — essa explicação foi testada e
    # descartada em 12/08/2026. Investigar como erro de classificação de conta.
    if t["DESPESA_PAGA"] > t["DESP_LIQUIDADA"] + D("0.02"):
        dif = t["DESPESA_PAGA"] - t["DESP_LIQUIDADA"]
        achados.append(achado_alerta("BO","BO-01b","Paga > Liquidada (verificar)",
            f"Paga {t['DESPESA_PAGA']:,.2f}  >  Liquidada {t['DESP_LIQUIDADA']:,.2f}  "
            f"Diferença: {dif:,.2f}  — verificar se alguma conta de pagamento "
            f"extraorçamentário (RP) entrou na Despesa Orçamentária Paga",
            valor=dif))

    # BO-02: Empenhada <= Dotação Atualizada
    gap = t["DESP_EMPENHADA"] - t["DOT_ATUALIZADA"]
    if gap <= D("0.02"):
        achados.append(achado_ok("BO","BO-02","Despesa Empenhada ≤ Dotação Atualizada",
            f"Empenhada {t['DESP_EMPENHADA']:,.2f}  ≤  Atualizada {t['DOT_ATUALIZADA']:,.2f}",
            valor=t["DESP_EMPENHADA"]))
    else:
        achados.append(achado_erro("BO","BO-02","Despesa Empenhada > Dotação Atualizada",
            f"Excesso: {gap:,.2f}", valor=gap))

    # ── BO-03: classificação orçamentária ausente na BALANCOGERAL ─────────
    # ERRO e não ALERTA: é despesa real que não aparece no Balanço
    # Orçamentário publicado. Cada dia sem correção é um dia a mais de
    # divergência entre a contabilidade e o demonstrativo.
    try:
        linhas = query_all(conn, SQL_BO03.format(mes=mes, ano=ano))
    except Exception as e:
        linhas = None
        achados.append(achado_alerta("BO", "BO-03",
            "Não foi possível verificar a BALANCOGERAL",
            f"{type(e).__name__}: {e}  — sem este controle, registros com "
            f"classificação zerada passam despercebidos até o fechamento"))

    if linhas is not None:
        total = sum(l["VLR"] for l in linhas)
        t["BG_SEM_CLASSIFICACAO"] = total
        if not linhas:
            achados.append(achado_ok("BO", "BO-03",
                "BALANCOGERAL: toda despesa tem classificação orçamentária",
                "Nenhum registro de despesa com CONATUREZA ou COUO zerados",
                valor=D(0)))
        else:
            det = "  |  ".join(
                f"mês {int(l['INMES'])} conta {int(l['COCONTACONTABIL'])}: "
                f"{l['VLR']:,.2f} ({int(l['QTD'])} reg.)" for l in linhas[:8])
            achados.append(achado_erro("BO", "BO-03",
                "Despesa sem classificação orçamentária na BALANCOGERAL",
                f"Total {total:,.2f} em {len(linhas)} combinação(ões) "
                f"mês/conta — este valor existe na contabilidade mas NÃO entra "
                f"no Balanço Orçamentário (itens 2.01.01.0X filtram por GND "
                f"1..9) nem na abertura por UO.  {det}", valor=total))

    # BO-04: Previsão Adicional a Lançar (C18 / 521920500) ──────────────────
    # Conta de trânsito onde créditos adicionais ficam registrados antes de
    # serem reclassificados às dotações definitivas. Saldo != 0 indica que
    # há dotação aprovada que ainda não entrou nas contas de despesa normais
    # e explica o eventual gap no equilíbrio orçamentário (Previsão Atualizada
    # + Superávit Financeiro ≠ Dotação Atualizada). Deve zerar ao fim do mês.
    # Qualquer saldo != 0 já é ALERTA (não espera o mês fechar nem exige
    # saldo de mês anterior) — pedido explícito do usuário em 09/09/2026.
    SQL_C18 = f"""
        SELECT
            NVL(SUM(CASE WHEN INMES <= {mes} THEN VADEBITO - VACREDITO ELSE 0 END), 0) AS SALDO_521,
            NVL(SUM(CASE WHEN INMES <  {mes} THEN VADEBITO - VACREDITO ELSE 0 END), 0) AS SALDO_521_ANT
        FROM   MIL{ano}.SALDOCONTABIL
        WHERE  COCONTACONTABIL = 521920500
    """
    try:
        row_c18 = query_one(conn, SQL_C18)
        saldo_521 = D(str(row_c18['SALDO_521']))
        saldo_521_ant = D(str(row_c18['SALDO_521_ANT']))
        t['SALDO_521920500'] = saldo_521
        if abs(saldo_521) < D('1.00'):
            achados.append(achado_ok("BO", "BO-04",
                "Previsão Adicional a Lançar (521920500) zerada",
                f"Saldo INMES ≤ {mes}: {saldo_521:,.2f} — nenhuma previsão pendente",
                valor=saldo_521))
        else:
            achados.append(achado_alerta("BO", "BO-04",
                "Previsão Adicional a Lançar pendente (C18)",
                f"Saldo 521920500 = {saldo_521:,.2f} (dos quais {saldo_521_ant:,.2f} "
                f"já vem de mês(es) anterior(es) já encerrado(s)) — dotação aprovada "
                f"ainda não reclassificada às contas definitivas; explica o gap do "
                f"equilíbrio orçamentário. Veja rotina C18.",
                valor=abs(saldo_521)))
    except Exception as e:
        achados.append(achado_alerta("BO", "BO-04",
            "Não foi possível verificar 521920500 (C18)",
            f"{type(e).__name__}: {e}"))

    return achados, t
