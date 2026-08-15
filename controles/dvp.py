# -*- coding: utf-8 -*-
"""
Controles da Demonstração das Variações Patrimoniais (Anexo 15).

Classes contábeis:
  VPA = classe 4 (crédito → SC) → VACREDITO - VADEBITO
  VPD = classe 3 (débito  → SD) → VADEBITO - VACREDITO
  Resultado Patrimonial = VPA − VPD

SOBRE O GRUPO 38 (Custo de Merc./Prod./Serviços)
------------------------------------------------
A versão anterior deste arquivo trazia dois comentários que se contradiziam:
o docstring afirmava que o grupo 38 havia sido incluído e que isso corrigira
uma divergência de R$ 50.340,00; o comentário do SQL afirmava que o grupo NÃO
era somado. Na prática nenhum dos dois descrevia o código: a faixa
300000000..399999999 contém 380000000..389999999, então o grupo 38 sempre
esteve incluído. Comentário removido em vez de reescrito — o range fala por si.

CORREÇÃO 13/08/2026 — DVP-01 não testava nada
---------------------------------------------
O controle era um achado_ok incondicional (sem nenhum `if`) logo após a linha
`resultado = vpa - vpd`. Ou seja: reportava uma atribuição como controle
aprovado, e passaria com qualquer dado, inclusive corrompido.

Agora o DVP-01 confronta o resultado calculado com a conta de encerramento
891XXXXXX, que o extrair() já trazia e que ninguém usava. Em meses
intermediários essa conta costuma estar zerada (o encerramento ocorre no
fechamento do exercício); nesse caso o controle vira INFO explicitando que a
validação do resultado fica a cargo dos cruzamentos X6 (DMPL) e R1a (BP),
em vez de fingir que verificou algo.
"""
from __future__ import annotations
from . import (Achado, query_one, D,
               achado_ok, achado_erro, achado_alerta, achado_info, checa_gap)

SQL_DVP = """
SELECT
    -- VPA Total (classe 4 inteira, SC = crédito - débito)
    SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 400000000 AND 499999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)   AS TOTAL_VPA,

    -- VPD Total (classe 3 inteira, SD = débito - crédito)
    SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 300000000 AND 399999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)   AS TOTAL_VPD,

    -- Itens VPA (para conferir subtotais)
    SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 410000000 AND 419999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)   AS VPA_IMPOSTOS,
    SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 420000000 AND 429999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)   AS VPA_CONTRIBUICOES,
    SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 450000000 AND 459999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)   AS VPA_TRANSF_RECEBIDAS,

    -- Itens VPD (para conferir subtotais)
    SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 310000000 AND 319999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)   AS VPD_PESSOAL,
    SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 350000000 AND 359999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)   AS VPD_TRANSF_CONCEDIDAS,

    -- Conta de encerramento — fonte independente do resultado (DVP-01)
    SUM(CASE WHEN v.INMES BETWEEN 1 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 891000000 AND 891999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)   AS RESULTADO_ENCERRAMENTO

FROM MIL{ano}.VSALDOCONTABIL v
"""


def extrair(conn, mes: int, ano: int) -> dict:
    return query_one(conn, SQL_DVP.format(mes=mes, ano=ano))


def auditar(conn, mes: int, ano: int) -> tuple[list[Achado], dict]:
    t = extrair(conn, mes, ano)
    achados = []

    vpa = t["TOTAL_VPA"]
    vpd = t["TOTAL_VPD"]
    resultado = vpa - vpd
    t["RESULTADO_PATRIMONIAL"] = resultado
    enc = t.get("RESULTADO_ENCERRAMENTO", D(0))

    # ── DVP-01: resultado calculado × conta de encerramento ────────────────
    if checa_gap(enc):
        achados.append(achado_info("DVP", "DVP-01",
            "Resultado Patrimonial (encerramento ainda não lançado)",
            f"VPA {vpa:,.2f}  −  VPD {vpd:,.2f}  =  {resultado:,.2f}  — a conta "
            f"891XXXXXX está zerada no período, então não há fonte "
            f"independente para confrontar; a validação do resultado fica com "
            f"os cruzamentos X6 (DMPL) e R1a (BP)", valor=resultado))
    else:
        gap = resultado - enc
        if checa_gap(gap):
            achados.append(achado_ok("DVP", "DVP-01",
                "Resultado Patrimonial = conta de encerramento",
                f"VPA − VPD {resultado:,.2f}  =  891XXXXXX {enc:,.2f}",
                valor=resultado))
        else:
            achados.append(achado_erro("DVP", "DVP-01",
                "Resultado Patrimonial ≠ conta de encerramento",
                f"Diferença: {gap:,.2f}  |  VPA − VPD {resultado:,.2f}  "
                f"vs 891XXXXXX {enc:,.2f}", valor=gap))

    # ── DVP-02: VPA e VPD com sinal esperado ───────────────────────────────
    for nome, val, cod in [("VPA Total", vpa, "DVP-02a"),
                           ("VPD Total", vpd, "DVP-02b")]:
        if val < 0:
            achados.append(achado_alerta("DVP", cod,
                f"{nome} negativo — verificar",
                f"{nome}: {val:,.2f}", valor=val))
        else:
            achados.append(achado_ok("DVP", cod,
                f"{nome} positivo", f"{val:,.2f}", valor=val))

    # ── DVP-03: cobertura dos subtotais extraídos ──────────────────────────
    # Os subtotais não esgotam as classes 3 e 4 — são amostras. Informativo,
    # para dar noção de quanto do total está sendo olhado item a item.
    sub_vpa = (t["VPA_IMPOSTOS"] + t["VPA_CONTRIBUICOES"]
               + t["VPA_TRANSF_RECEBIDAS"])
    sub_vpd = t["VPD_PESSOAL"] + t["VPD_TRANSF_CONCEDIDAS"]
    achados.append(achado_info("DVP", "DVP-03",
        "Cobertura dos subtotais detalhados",
        f"VPA: subtotais {sub_vpa:,.2f} de {vpa:,.2f} | "
        f"VPD: subtotais {sub_vpd:,.2f} de {vpd:,.2f} — o restante não é "
        f"aberto por grupo neste diagnóstico", valor=sub_vpa + sub_vpd))

    return achados, t
