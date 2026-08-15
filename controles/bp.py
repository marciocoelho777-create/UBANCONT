# -*- coding: utf-8 -*-
"""
Controles do Balanço Patrimonial (Anexo 14).

Contas:
  Ativo Circulante   : 111..119XXXXXX
  Ativo Não Circ.    : 121..129XXXXXX
  Passivo Circulante : 211..219XXXXXX
  Passivo Não Circ.  : 221..229XXXXXX
  PL                 : 231..237XXXXXX + Resultado do Exercício (calculado)
  Caixa              : 111XXXXXX

CORREÇÃO 13/08/2026 — BP-02 era tautológico
-------------------------------------------
    ativo = ATIVO_CIRC + ATIVO_NCIRC
    gap2  = ativo - (ATIVO_CIRC + ATIVO_NCIRC)   # ≡ 0 por construção

O controle passava sempre, com qualquer dado. Agora o BP-02 compara a soma
das duas faixas declaradas contra o total da classe 1 inteira: se existir
conta de ativo fora de 111..119 e 121..129, ela aparece. O mesmo teste de
cobertura foi acrescentado para a classe 2 no BP-04 — é ali que uma conta
como a 238 (Ações em Tesouraria), fora da faixa 231..237 do PL, apareceria.
"""
from __future__ import annotations
from . import (Achado, query_one, D,
               achado_ok, achado_erro, achado_alerta, achado_info, checa_gap)

SQL_BP = """
SELECT
    -- ATIVO CIRCULANTE (saldo final = INMES 0..mes)
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 111000000 AND 119999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)   AS ATIVO_CIRC,

    -- ATIVO NÃO CIRCULANTE
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 121000000 AND 129999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)   AS ATIVO_NCIRC,

    -- Classe 1 inteira — para o teste de cobertura do BP-02
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 100000000 AND 199999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)   AS CLASSE1_TOTAL,

    -- PASSIVO CIRCULANTE
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 211000000 AND 219999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)   AS PASSIVO_CIRC,

    -- PASSIVO NÃO CIRCULANTE
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 221000000 AND 229999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)   AS PASSIVO_NCIRC,

    -- PATRIMÔNIO LÍQUIDO (exceto Resultado do Exercício — calculado à parte)
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 231000000 AND 237999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)   AS PL_SALDO,

    -- Classe 2 inteira — para o teste de cobertura do BP-04
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 200000000 AND 299999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)   AS CLASSE2_TOTAL,

    -- Resultado do Exercício: 4XXXXXXXX(SC) - 3XXXXXXXX(SD)
    SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 400000000 AND 499999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)
  - SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 300000000 AND 399999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)   AS RESULTADO_EXERCICIO,

    -- Caixa e Equivalentes (111XXXXXX)
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 111000000 AND 111999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)   AS CAIXA

FROM MIL{ano}.VSALDOCONTABIL v
"""


def extrair(conn, mes: int, ano: int) -> dict:
    return query_one(conn, SQL_BP.format(mes=mes, ano=ano))


def auditar(conn, mes: int, ano: int) -> tuple[list[Achado], dict]:
    t = extrair(conn, mes, ano)
    achados = []

    ativo  = t["ATIVO_CIRC"] + t["ATIVO_NCIRC"]
    pl     = t["PL_SALDO"] + t["RESULTADO_EXERCICIO"]
    passivo_pl = t["PASSIVO_CIRC"] + t["PASSIVO_NCIRC"] + pl
    t["ATIVO_TOTAL"]        = ativo
    t["PATRIMONIO_LIQUIDO"] = pl
    t["PASSIVO_PL_TOTAL"]   = passivo_pl

    # ── BP-01: Ativo = Passivo + PL ────────────────────────────────────────
    gap = ativo - passivo_pl
    if checa_gap(gap):
        achados.append(achado_ok("BP", "BP-01",
            "Ativo = Passivo + PL",
            f"Ativo {ativo:,.2f}  =  Passivo+PL {passivo_pl:,.2f}",
            valor=ativo))
    else:
        achados.append(achado_erro("BP", "BP-01",
            "Ativo ≠ Passivo + PL",
            f"Diferença: {gap:,.2f}  |  "
            f"Ativo {ativo:,.2f}  Passivo+PL {passivo_pl:,.2f}",
            valor=gap))

    # ── BP-02: cobertura das faixas do ATIVO ───────────────────────────────
    # Circ + Não Circ tem que esgotar a classe 1. Se houver conta de ativo
    # fora de 111..119 e 121..129, a diferença aparece aqui.
    gap2 = ativo - t["CLASSE1_TOTAL"]
    if checa_gap(gap2):
        achados.append(achado_ok("BP", "BP-02",
            "Ativo Circ + Não Circ esgota a classe 1",
            f"Circ {t['ATIVO_CIRC']:,.2f}  +  NCirc {t['ATIVO_NCIRC']:,.2f}  "
            f"=  {ativo:,.2f}  =  classe 1 {t['CLASSE1_TOTAL']:,.2f}",
            valor=ativo))
    else:
        achados.append(achado_erro("BP", "BP-02",
            "Há conta de ativo fora das faixas 111..119 / 121..129",
            f"Diferença: {gap2:,.2f}  |  faixas {ativo:,.2f}  "
            f"vs classe 1 completa {t['CLASSE1_TOTAL']:,.2f}", valor=gap2))

    # ── BP-03: Caixa positivo ──────────────────────────────────────────────
    if t["CAIXA"] < 0:
        achados.append(achado_erro("BP", "BP-03",
            "Caixa negativo no BP",
            f"Caixa: {t['CAIXA']:,.2f}", valor=t["CAIXA"]))
    else:
        achados.append(achado_ok("BP", "BP-03",
            "Caixa positivo",
            f"Caixa: {t['CAIXA']:,.2f}", valor=t["CAIXA"]))

    # ── BP-04: cobertura das faixas do PASSIVO + PL ────────────────────────
    # Passivo Circ + Não Circ + PL(231..237) tem que esgotar a classe 2.
    # É aqui que a 238 (Ações/Cotas em Tesouraria), fora da faixa do PL usada
    # tanto no bp.py quanto no dmpl.py, apareceria se tivesse saldo.
    soma_c2 = t["PASSIVO_CIRC"] + t["PASSIVO_NCIRC"] + t["PL_SALDO"]
    gap4 = soma_c2 - t["CLASSE2_TOTAL"]
    if checa_gap(gap4):
        achados.append(achado_ok("BP", "BP-04",
            "Passivo + PL esgota a classe 2",
            f"Passivo {t['PASSIVO_CIRC'] + t['PASSIVO_NCIRC']:,.2f}  +  "
            f"PL(231..237) {t['PL_SALDO']:,.2f}  =  classe 2 "
            f"{t['CLASSE2_TOTAL']:,.2f}", valor=t["CLASSE2_TOTAL"]))
    else:
        achados.append(achado_erro("BP", "BP-04",
            "Há conta de passivo/PL fora das faixas declaradas",
            f"Diferença: {gap4:,.2f}  |  faixas {soma_c2:,.2f}  "
            f"vs classe 2 completa {t['CLASSE2_TOTAL']:,.2f}  "
            f"— verificar 238 (Ações/Cotas em Tesouraria) e 234 fora de 23411",
            valor=gap4))

    return achados, t
