# -*- coding: utf-8 -*-
"""
Controles da Demonstração das Variações Patrimoniais (Anexo 15).

Classes contábeis:
  VPA = classe 4 (crédito → SC) → SC = DECODE(INDEBITOCREDITO,'C',VAL,'D',-VAL,0)
  VPD = classe 3 (débito  → SD) → SD = DECODE(INDEBITOCREDITO,'D',VAL,'C',-VAL,0)
  Resultado Patrimonial = VPA − VPD

MIGRAÇÃO 21/08/2026 — VSALDOCONTABIL → LANCAMENTOCONTABIL
----------------------------------------------------------
A versão anterior usava VSALDOCONTABIL (somando INMES 1..mes). O problema:
VSALDOCONTABIL é uma view que pode ser refrescada por batch intraday; se o
diagnóstico rodar enquanto o refresh está em andamento, o DVP captura um
snapshot mais antigo do que os módulos que executam depois (BP, DMPL). Isso
gerava erros espúrios no X6 (DMPL≠DVP) e no R1a (BP≠DVP) durante runs
intraday, que desapareciam no run noturno quando o refresh já havia terminado.

Para classes 3 e 4 (contas de resultado) o saldo de abertura é sempre zero,
então LANCAMENTOCONTABIL(INMES 1..mes) já dá o acumulado correto — exatamente
o mesmo padrão que o DMPL usa para calcular o Resultado do Exercício. Essa
fonte é sempre atualizada em tempo real pelos lançamentos e elimina o lag.

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

CORREÇÃO 16/08/2026 — DVP-01 comparava totais incompatíveis
-------------------------------------------------------------
Investigado em 16/08/2026 (execução 8/2026): a 891XXXXXX não fica
necessariamente zerada em mês intermediário — algumas UGs (autarquias/fundos
que encerram resultado mensalmente, ex.: UG 130101/conta 891240100) lançam
nela mês a mês, enquanto a administração direta só encerra no fechamento do
exercício. O código anterior comparava esse saldo PARCIAL de uma única UG
contra o VPA−VPD CONSOLIDADO de todo o GDF — comparação de grandezas
incompatíveis, que gerou um "erro" de R$ 14,3 bilhões sem nenhum problema
contábil real (confirmado cruzando com BALANCOGERAL, fonte oficial dos
demonstrativos: os dados batem).

Tentativa de correção por UG (comparar 891 da UG × VPA−VPD da MESMA UG) foi
descartada: testado com os dados reais, a UG 130101 sozinha tem VPA−VPD de
R$ 17,3 bilhões (ela concentra lançamentos de consolidação entre fundos),
contra R$ 256 mil na 891 — o mesmo tipo de falso positivo, só que menor.
Ou seja, um lançamento parcial em 891 não significa "esta UG encerrou o
resultado do período"; é outra coisa (ajuste pontual), e não há como
distinguir isso do encerramento real sem conhecer a natureza do lançamento.

Correção definitiva: 891XXXXXX só é fonte independente do resultado no
FECHAMENTO DO EXERCÍCIO (mês 12, quando o encerramento anual realmente
ocorre). Em qualquer mês intermediário — mesmo com lançamentos parciais na
891 — o DVP-01 vira INFO (a validação do resultado fica com os cruzamentos
X6/DMPL e R1a/BP, que comparam fontes independentes de verdade). Só em
dezembro o controle vira OK/ERRO de fato.
"""
from __future__ import annotations
from . import (Achado, query_one, query_all, D,
               achado_ok, achado_erro, achado_alerta, achado_info, checa_gap)

SQL_DVP = """
SELECT
    -- VPA Total (classe 4 inteira, SC = crédito - débito)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 400000000 AND 499999999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS TOTAL_VPA,

    -- VPD Total (classe 3 inteira, SD = débito - crédito)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 300000000 AND 399999999
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS TOTAL_VPD,

    -- Itens VPA (para conferir subtotais)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 410000000 AND 419999999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS VPA_IMPOSTOS,
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 420000000 AND 429999999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS VPA_CONTRIBUICOES,
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 450000000 AND 459999999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS VPA_TRANSF_RECEBIDAS,

    -- Itens VPD (para conferir subtotais)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 310000000 AND 319999999
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS VPD_PESSOAL,
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 350000000 AND 359999999
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS VPD_TRANSF_CONCEDIDAS,

    -- Conta de encerramento — fonte independente do resultado (DVP-01)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 891000000 AND 891999999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS RESULTADO_ENCERRAMENTO

FROM MIL{ano}.LANCAMENTOCONTABIL o
"""

# Detalhe informativo para o DVP-01 (não decide OK/ERRO fora de dezembro,
# ver docstring do módulo) — quais UGs já lançaram algo em 891XXXXXX.
SQL_DVP_UG = """
SELECT o.COUG,
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 400000000 AND 499999999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS VPA,
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 300000000 AND 399999999
         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS VPD,
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 891000000 AND 891999999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END)   AS ENCERRAMENTO
FROM MIL{ano}.LANCAMENTOCONTABIL o
GROUP BY o.COUG
HAVING SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}
              AND o.COCONTACONTABIL BETWEEN 891000000 AND 891999999
         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
         ELSE 0 END) != 0
"""


def extrair(conn, mes: int, ano: int) -> dict:
    return query_one(conn, SQL_DVP.format(mes=mes, ano=ano))


def extrair_por_ug(conn, mes: int, ano: int) -> list[dict]:
    """UGs que já lançaram algo em 891XXXXXX (encerramento parcial)."""
    return query_all(conn, SQL_DVP_UG.format(mes=mes, ano=ano))


def auditar(conn, mes: int, ano: int) -> tuple[list[Achado], dict]:
    t = extrair(conn, mes, ano)
    achados = []

    vpa = t["TOTAL_VPA"]
    vpd = t["TOTAL_VPD"]
    resultado = vpa - vpd
    t["RESULTADO_PATRIMONIAL"] = resultado
    enc = t.get("RESULTADO_ENCERRAMENTO", D(0))
    t["RESULTADO_ENCERRAMENTO"] = enc

    # ── DVP-01: resultado calculado × conta de encerramento ────────────────
    # 891XXXXXX só é fonte independente do resultado no fechamento do
    # exercício (mês 12). Lançamentos parciais nela em mês intermediário
    # não significam encerramento (nem do GDF, nem de uma UG isolada — ver
    # docstring), então fora de dezembro o controle é sempre INFO.
    ugs = extrair_por_ug(conn, mes, ano)
    t["DVP01_UGS_COM_LANCAMENTO_891"] = ugs
    if mes < 12:
        obs_ugs = ""
        if ugs:
            obs_ugs = ("  |  Lançamentos parciais na 891 (não é encerramento, "
                        "não entra na comparação): " +
                        ", ".join(f"UG {u['COUG']} {u['ENCERRAMENTO']:,.2f}"
                                  for u in ugs))
        achados.append(achado_info("DVP", "DVP-01",
            "Resultado Patrimonial (encerramento do exercício só em dezembro)",
            f"VPA {vpa:,.2f}  −  VPD {vpd:,.2f}  =  {resultado:,.2f}  — a 891XXXXXX "
            f"só fecha o exercício em dezembro; a validação do resultado neste "
            f"período fica com os cruzamentos X6 (DMPL) e R1a (BP)"
            f"{obs_ugs}", valor=resultado))
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
