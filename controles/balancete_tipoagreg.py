# -*- coding: utf-8 -*-
"""
Controles de integridade do Balancete Contábil.

Regra fundamental das partidas dobradas aplicada ao balancete consolidado:
    DEVEDORES (classes 1+3+5+7) = CREDORES (classes 2+4+6+8+9)

Corresponde à REGRA 11 — IMPBALANCETE do sistema de rotinas contábeis.

Classes do plano de contas PCASP:
  1 — Ativo
  2 — Passivo e Patrimônio Líquido
  3 — Variações Patrimoniais Diminutivas (VPD)
  4 — Variações Patrimoniais Aumentativas (VPA)
  5 — Controles da Aprovação do Planejamento e Orçamento
  6 — Controles da Execução do Planejamento e Orçamento
  7 — Controles Devedores
  8 — Controles Credores
  9 — Resultado do Exercício (encerramento)
"""
from __future__ import annotations
from . import (Achado, query_one, D,
               achado_ok, achado_erro, achado_info, checa_gap)

# Cópia de balancete.py — filtro por tipo de agregação de gestão. NÃO mexer no original.
def _qf(conn, sql, cogestao_list, alias, **fmt):
    s = sql.format(**fmt) if fmt else sql
    if cogestao_list:
        ids = ",".join(str(int(c)) for c in cogestao_list)
        kw = "AND" if "WHERE" in s.upper() else "WHERE"
        s += f"\n  {kw} {alias}.COGESTAO IN ({ids})"
    return query_one(conn, s)

SQL_BALANCETE = """
SELECT
    -- DEVEDORES (classes de saldo devedor: 1, 3, 5, 7)
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 100000000 AND 199999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)  AS CL1_ATIVO,

    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 300000000 AND 399999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)  AS CL3_VPD,

    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 500000000 AND 599999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)  AS CL5_CONTROR_APROV,

    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 700000000 AND 799999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END)  AS CL7_CONTRDEV,

    -- CREDORES (classes de saldo credor: 2, 4, 6, 8, 9)
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 200000000 AND 299999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)  AS CL2_PASSIVO_PL,

    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 400000000 AND 499999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)  AS CL4_VPA,

    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 600000000 AND 699999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)  AS CL6_CONTROR_EXEC,

    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 800000000 AND 899999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)  AS CL8_CONTRCRED,

    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 900000000 AND 999999999
         THEN v.VACREDITO - v.VADEBITO ELSE 0 END)  AS CL9_RESULTADO

FROM MIL{ano}.SALDOCONTABIL v
"""


def extrair(conn, mes: int, ano: int, cogestao_list=None) -> dict:
    t = _qf(conn, SQL_BALANCETE, cogestao_list, "v", mes=mes, ano=ano)
    return {k: (D(0) if v is None else v) for k, v in t.items()}


def auditar(conn, mes: int, ano: int, cogestao_list=None) -> tuple[list[Achado], dict]:
    t = extrair(conn, mes, ano, cogestao_list)
    achados = []

    devedores = t["CL1_ATIVO"] + t["CL3_VPD"] + t["CL5_CONTROR_APROV"] + t["CL7_CONTRDEV"]
    credores  = t["CL2_PASSIVO_PL"] + t["CL4_VPA"] + t["CL6_CONTROR_EXEC"] + t["CL8_CONTRCRED"] + t["CL9_RESULTADO"]
    t["DEVEDORES_TOTAL"] = devedores
    t["CREDORES_TOTAL"]  = credores

    gap = devedores - credores

    detalhe = (
        f"Devedores {devedores:,.2f}  =  "
        f"Cl.1 {t['CL1_ATIVO']:,.2f}  +  Cl.3 {t['CL3_VPD']:,.2f}  +  "
        f"Cl.5 {t['CL5_CONTROR_APROV']:,.2f}  +  Cl.7 {t['CL7_CONTRDEV']:,.2f}  |  "
        f"Credores {credores:,.2f}  =  "
        f"Cl.2 {t['CL2_PASSIVO_PL']:,.2f}  +  Cl.4 {t['CL4_VPA']:,.2f}  +  "
        f"Cl.6 {t['CL6_CONTROR_EXEC']:,.2f}  +  Cl.8 {t['CL8_CONTRCRED']:,.2f}  +  "
        f"Cl.9 {t['CL9_RESULTADO']:,.2f}"
    )

    # ── BAL-01: Devedores (1+3+5+7) = Credores (2+4+6+8+9) ──────────────────
    if checa_gap(gap):
        achados.append(achado_ok("BAL", "BAL-01",
            "Balancete: Devedores (1+3+5+7) = Credores (2+4+6+8+9)",
            detalhe, valor=devedores))
    else:
        achados.append(achado_erro("BAL", "BAL-01",
            "Balancete: Devedores ≠ Credores",
            f"Diferença: {gap:,.2f}  |  {detalhe}", valor=gap))

    # ── BAL-02: saldos com sinal esperado por classe ───────────────────────────
    for cod_cl, nome, val in [
        ("Cl.1 Ativo",              "CL1_ATIVO",        t["CL1_ATIVO"]),
        ("Cl.2 Passivo/PL",         "CL2_PASSIVO_PL",   t["CL2_PASSIVO_PL"]),
        ("Cl.3 VPD",                "CL3_VPD",          t["CL3_VPD"]),
        ("Cl.4 VPA",                "CL4_VPA",          t["CL4_VPA"]),
        ("Cl.5 Controle Aprovação", "CL5_CONTROR_APROV",t["CL5_CONTROR_APROV"]),
        ("Cl.6 Controle Execução",  "CL6_CONTROR_EXEC", t["CL6_CONTROR_EXEC"]),
        ("Cl.7 Controles Dev.",     "CL7_CONTRDEV",     t["CL7_CONTRDEV"]),
        ("Cl.8 Controles Cred.",    "CL8_CONTRCRED",    t["CL8_CONTRCRED"]),
        ("Cl.9 Resultado",          "CL9_RESULTADO",    t["CL9_RESULTADO"]),
    ]:
        if val < 0:
            achados.append(achado_info("BAL", "BAL-02",
                f"{cod_cl} com saldo negativo",
                f"{cod_cl}: {val:,.2f} — verificar se esperado para o período",
                valor=val))

    return achados, t
