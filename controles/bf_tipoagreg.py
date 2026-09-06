# -*- coding: utf-8 -*-
"""
Controles do Balanço Financeiro (Anexo 13 / PSIAG560) — equação COMPLETA.

Esta versão porta a lógica do bf.py mestre (que fecha Ingressos = Dispêndios)
para o diagnóstico. A versão anterior usava uma equação simplificada que
omitia o Saldo do Exercício Anterior e o Saldo p/ o Exercício Seguinte, o que
tornava o BF-01 um alerta permanente de ~R$ 2,9 bi sem significado.

LÓGICA DO CAMPO "MÊS" DA LISTA DE EQUAÇÕES
------------------------------------------
A coluna "Mês" da Lista NÃO é mês do calendário — é seletor de origem:

    Mês 0   -> saldo de abertura            : SALDOCONTABIL INMES = 0
    Mês 1   -> em conta de SALDO (2.05.xx)  : VSALDO INMES=0 + LANC INMES 1..mes
    Mês 1   -> em conta de MOVIMENTO        : LANCAMENTOCONTABIL INMES 1..mes
    Mês 13  -> movimento acumulado          : LANCAMENTOCONTABIL INMES 1..mes
    Mês 15  -> encerramento                 : LANCAMENTOCONTABIL INMES = 15

O híbrido do item 2.05.xx é necessário porque o LANCAMENTOCONTABIL não tem
INMES=0 para as contas 111XXXXXX: o saldo de abertura existe apenas no
SALDOCONTABIL. As contas 113XXXXXX seguem a mesma regra (confirmado no
mestre: a diferença observada era exatamente o saldo de abertura).

ATENÇÃO ao parâmetro --mes: com mes <= 12 o range 1..mes nunca alcança
INMES 13/14/15, então os termos de encerramento zeram sozinhos. Com mes >= 13
o range passaria a incluir lançamentos de encerramento AO MESMO TEMPO que os
termos explícitos de INMES=15, contando duas vezes. Por isso o guard em
extrair(): meses 13/14 são rejeitados.

SINAL DA 113829700
------------------
É SC no Saldo Anterior (item 1.05.05) e SD no Saldo Seguinte (item 2.05.05).
A assimetria parece erro mas está assim na Lista de Equações; replicada aqui.

DESVIOS CONHECIDOS DO MESTRE EM RELAÇÃO À LISTA (herdados nesta porta)
---------------------------------------------------------------------
Mantidos para que o resultado seja idêntico ao do demonstrativo publicado.
Revisar com quem mantém a Lista antes de mudar:
  1. Resgates/Aplicações (1.03.01/2.03.01): a Lista mapeia 1141XXXXX,
     1144XXXXX, 122310103, 122310104 e 821192501; o mestre usa só a
     conta de controle 821192501.
  2. A 113511200 é movimento na Lista (1.03.02 Desbloqueios / 2.03.02
     Bloqueios) mas entra no saldo VALTRANS no mestre.
  3. O item 2.03.07 (Ajuste para Perdas, conta 361711501) não existe no mestre.
  4. Os itens 2.03.05 e 2.04.04 compartilham as mesmas contas sem
     qualificador que os distinga; o mestre soma as do 2.04.04 uma vez só.

CORREÇÃO 14/08/2026 — as contas "inter" faltavam no item 2.04.04
----------------------------------------------------------------
O quinto dígito dessas contas indica a esfera: 1 = consolidado, 2 = intra,
3 = inter. São portanto três variantes de duas bases:

    237110301 / 237120301 / 237130301   (base 2371x0301)
    237210301 / 237220301 / 237230301   (base 2372x0301)

O mestre somava consolidado e intra das duas bases e omitia as duas INTER
(237130301 e 237230301). Esta porta herdara a omissão; as seis agora entram,
conforme a Lista de Equações.

ATENÇÃO ao efeito no BF-01: AJUSTES é dispêndio, então incluir as duas
aumenta os Dispêndios. Se o controle fechava em R$ 0,00 com quatro contas,
ele passa a acusar diferença igual ao valor das duas novas. Isso NÃO é
regressão — é o fim de uma compensação: o zero anterior dependia de omitir
também o item 2.03.05, que usa as mesmas seis contas e que nem o mestre nem
esta porta implementam. A relação entre 2.03.05 e 2.04.04 segue em aberto
com quem mantém a Lista.
"""
from __future__ import annotations
from . import (Achado, query_one, D,
               achado_ok, achado_erro, achado_alerta, achado_info, checa_gap)

# Cópia de bf.py — filtro por tipo de agregação de gestão. NÃO mexer no original.
# campo: LANCAMENTOCONTABIL usa COGESTAOCONTAB (gestão contabilizante, = PSIAG);
#        SALDOCONTABIL usa COGESTAO (não tem COGESTAOCONTAB).
def _qf(conn, sql, cogestao_list, alias, campo="COGESTAO", **fmt):
    s = sql.format(**fmt) if fmt else sql
    if cogestao_list:
        ids = ",".join(str(int(c)) for c in cogestao_list)
        kw = "AND" if "WHERE" in s.upper() else "WHERE"
        s += f"\n  {kw} {alias}.{campo} IN ({ids})"
    return query_one(conn, s)

# ── Contas de Depósitos Restituíveis (itens 1.04.03 / 2.04.03) ──────────────
# Idêntica ao mestre, inclusive as duas exclusões.
_DEPREST = """( o.COCONTACONTABIL BETWEEN 218810400 AND 218810499
             OR o.COCONTACONTABIL BETWEEN 218817000 AND 218817099
             OR o.COCONTACONTABIL BETWEEN 218820400 AND 218820499
             OR o.COCONTACONTABIL BETWEEN 218827000 AND 218827099
             OR o.COCONTACONTABIL BETWEEN 218837000 AND 218837099
             OR o.COCONTACONTABIL BETWEEN 218810300 AND 218810399
             OR o.COCONTACONTABIL BETWEEN 218810800 AND 218810899
             OR o.COCONTACONTABIL IN (214120200,218829900,218910105,
                                      218910108,218920199,218924500) )
          AND o.COCONTACONTABIL NOT IN (218810304,218810405)"""

_SC = "DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)"
_SD = "DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)"

# ── Movimento do exercício: LANCAMENTOCONTABIL INMES 1..mes ─────────────────
SQL_MOV = f"""
SELECT
    -- 1.01 Receita Orçamentária
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL IN (621200000,612000000)
         THEN {_SC} ELSE 0 END)
  - SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL BETWEEN 621300000 AND 621399999
         THEN {_SD} ELSE 0 END)
    AS RECEITA,

    -- 1.02 Transferências Financeiras Recebidas
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL IN (451120100,451120200,451120300,451120400,
                                        451120900,451121300,451129900,
                                        451220101,451220104,451220109,451220199)
         THEN {_SC} ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL BETWEEN 451300000 AND 451399999
         THEN {_SC} ELSE 0 END)
    AS TRANSF_RECEB,

    -- 1.03.01 Resgates de Investimentos (conta de controle 821192501, MD)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL = 821192501 AND o.INDEBITOCREDITO = 'D'
         THEN o.VALANCAMENTO ELSE 0 END)
    AS RESGATES,

    -- 1.04.01 / 1.04.02 Inscrição de RPÑP e RPP
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL = 531700000
         THEN {_SD} ELSE 0 END)
    AS INSCRICAO_RPNP,
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL = 532700000
         THEN {_SD} ELSE 0 END)
    AS INSCRICAO_RPP,

    -- 1.04.03 Depósitos Restituíveis — ingresso (MC)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}} AND o.INDEBITOCREDITO = 'C'
              AND {_DEPREST}
         THEN o.VALANCAMENTO ELSE 0 END)
    AS DEPREST_ING,

    -- 1.04.04 Outros Recebimentos (MC - MD nas 2188150XX / 2188250XX)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND ( o.COCONTACONTABIL BETWEEN 218815000 AND 218815099
                 OR o.COCONTACONTABIL BETWEEN 218825000 AND 218825099)
         THEN {_SC} ELSE 0 END)
    AS OUTROS_RECEB,

    -- 2.01 Despesa Orçamentária (parte do exercício corrente)
    -- 6322XXXXX NÃO entra aqui: é o item 2.04.02.
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL = 622920104
         THEN {_SC} ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL IN (622920105,622920106,622920107)
         THEN {_SC} ELSE 0 END)
    AS DESPESA_MOV,

    -- 2.02 Transferências Financeiras Concedidas
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL IN (351120100,351120200,351120300,351120400,
                                        351120900,351129900,351220101,351220104,
                                        351220109,351220199)
         THEN {_SD} ELSE 0 END)
  + SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL BETWEEN 351300000 AND 351399999
         THEN {_SD} ELSE 0 END)
    AS TRANSF_CONC,

    -- 2.03.01 Transferências p/ Investimentos e Aplicações (821192501, MC)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL = 821192501 AND o.INDEBITOCREDITO = 'C'
         THEN o.VALANCAMENTO ELSE 0 END)
    AS APLICACOES,

    -- 2.04.01 Pagamentos de RPÑP
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL IN (631400000,631820000)
         THEN {_SC} ELSE 0 END)
    AS PAG_RPNP,

    -- 2.04.02 Pagamentos de RPP
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL BETWEEN 632200000 AND 632299999
         THEN {_SC} ELSE 0 END)
    AS PAG_RPP,

    -- 2.04.03 Depósitos Restituíveis — dispêndio (MD)
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}} AND o.INDEBITOCREDITO = 'D'
              AND {_DEPREST}
         THEN o.VALANCAMENTO ELSE 0 END)
    AS DEPREST_DISP,

    -- 2.04.04 Outros Pagamentos Extraorçamentários
    SUM(CASE WHEN o.INMES BETWEEN 1 AND {{mes}}
              AND o.COCONTACONTABIL IN (237110301,237120301,237210301,237220301,
                                        237130301,237230301)
         THEN {_SD} ELSE 0 END)
    AS AJUSTES

FROM MIL{{ano}}.LANCAMENTOCONTABIL o
WHERE o.COCONTACONTABIL BETWEEN 100000000 AND 899999999
"""

# ── Encerramento: LANCAMENTOCONTABIL INMES = 15 ─────────────────────────────
# Zerado em meses intermediários. Mantido para que o mesmo código sirva ao
# fechamento anual sem alteração.
SQL_ENC = f"""
SELECT
    - SUM(CASE WHEN o.INMES = 15 AND o.COCONTACONTABIL = 622920104
           THEN {_SC} ELSE 0 END)
    + SUM(CASE WHEN o.INMES = 15
                AND o.COCONTACONTABIL BETWEEN 622130000 AND 622139999
           THEN {_SC} ELSE 0 END)
    - SUM(CASE WHEN o.INMES = 15
                AND o.COCONTACONTABIL IN (622920105,622920106,622920107)
           THEN {_SC} ELSE 0 END)
    AS DESPESA_ENC
FROM MIL{{ano}}.LANCAMENTOCONTABIL o
WHERE o.COCONTACONTABIL BETWEEN 600000000 AND 699999999
"""

# ── Grupos de saldo (itens 1.05.xx e 2.05.xx) ───────────────────────────────
_G_CAIXA = ("""{a}.COCONTACONTABIL BETWEEN 111000000 AND 111999999
             AND {a}.COCONTACONTABIL NOT BETWEEN 111110600 AND 111110699
             AND {a}.COCONTACONTABIL NOT BETWEEN 111300000 AND 111399999""")
_G_RPPS  = "{a}.COCONTACONTABIL BETWEEN 111110600 AND 111110699"
_G_DEPR  = "{a}.COCONTACONTABIL BETWEEN 111300000 AND 111399999"
_G_VCOMP = ("""({a}.COCONTACONTABIL BETWEEN 113819100 AND 113819199
              OR {a}.COCONTACONTABIL BETWEEN 113811700 AND 113811799
              OR {a}.COCONTACONTABIL BETWEEN 113829100 AND 113829199
              OR {a}.COCONTACONTABIL BETWEEN 113821700 AND 113821799)""")
_G_VTRAN = ("""({a}.COCONTACONTABIL BETWEEN 113810600 AND 113810699
              OR {a}.COCONTACONTABIL IN (113819700,113511200))""")
_G_VTCR  = "{a}.COCONTACONTABIL = 113829700"

_GRUPOS = [("CAIXA", _G_CAIXA), ("RPPS", _G_RPPS), ("DEPR", _G_DEPR),
           ("VCOMP", _G_VCOMP), ("VTRAN", _G_VTRAN), ("VTCR", _G_VTCR)]


def _sql_saldo_ini(ano: int) -> str:
    """Saldo de abertura por grupo — SALDOCONTABIL INMES = 0 (Mês 0)."""
    partes = []
    for chave, cond in _GRUPOS:
        c = cond.format(a="v")
        partes.append(f"    SUM(CASE WHEN v.INMES = 0 AND {c}\n"
                      f"         THEN v.VADEBITO - v.VACREDITO ELSE 0 END) AS INI_SD_{chave}")
    # A 113829700 entra como SC no Saldo Anterior (item 1.05.05 da Lista).
    c = _G_VTCR.format(a="v")
    partes.append(f"    SUM(CASE WHEN v.INMES = 0 AND {c}\n"
                  f"         THEN v.VACREDITO - v.VADEBITO ELSE 0 END) AS INI_SC_VTCR")
    return ("SELECT\n" + ",\n\n".join(partes)
            + f"\nFROM MIL{ano}.SALDOCONTABIL v")


def _sql_saldo_mov(mes: int, ano: int) -> str:
    """Movimento do período por grupo — LANCAMENTOCONTABIL INMES 1..mes.
    Somado ao saldo de abertura, forma o Saldo p/ o Exercício Seguinte."""
    partes = []
    for chave, cond in _GRUPOS:
        c = cond.format(a="o")
        partes.append(f"    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes} AND {c}\n"
                      f"         THEN {_SD} ELSE 0 END) AS MOV_SD_{chave}")
    return ("SELECT\n" + ",\n\n".join(partes)
            + f"\nFROM MIL{ano}.LANCAMENTOCONTABIL o"
            + "\nWHERE o.COCONTACONTABIL BETWEEN 111000000 AND 113999999")


# Caixa para os cruzamentos X1a/X1b/R4a — definição do BP, mantida intacta.
SQL_CAIXA = """
SELECT
    SUM(CASE WHEN v.INMES = 0
              AND v.COCONTACONTABIL BETWEEN 111000000 AND 111999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END) AS CAIXA_INICIAL,
    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes}
              AND v.COCONTACONTABIL BETWEEN 111000000 AND 111999999
         THEN v.VADEBITO - v.VACREDITO ELSE 0 END) AS CAIXA_FINAL
FROM MIL{ano}.SALDOCONTABIL v
"""


def extrair(conn, mes, ano, cogestao_list=None):
    if mes in (13, 14):
        raise ValueError(
            f"--mes {mes} nao suportado: o range INMES 1..{mes} passaria a "
            "incluir lancamentos de encerramento que ja sao somados a parte "
            "pelo termo INMES=15, contando duas vezes. Use mes<=12.")

    # LANCAMENTOCONTABIL → COGESTAOCONTAB (gestão contabilizante, = critério do PSIAG)
    # SALDOCONTABIL      → COGESTAO       (não tem COGESTAOCONTAB)
    t = _qf(conn, SQL_MOV, cogestao_list, "o", campo="COGESTAOCONTAB", mes=mes, ano=ano)
    t.update(_qf(conn, SQL_ENC, cogestao_list, "o", campo="COGESTAOCONTAB", ano=ano))
    t.update(_qf(conn, SQL_CAIXA, cogestao_list, "v", mes=mes, ano=ano))
    ini = _qf(conn, _sql_saldo_ini(ano), cogestao_list, "v")
    mov = _qf(conn, _sql_saldo_mov(mes, ano), cogestao_list, "o", campo="COGESTAOCONTAB")
    t.update(ini)
    t.update(mov)
    # SUM() Oracle retorna NULL quando nenhuma linha passou pelo filtro de COGESTAO.
    # Substitui todos os None por D(0) para não quebrar a aritmética subsequente.
    t = {k: (D(0) if v is None else v) for k, v in t.items()}

    # 1.05 Saldo do Exercício Anterior (a 113829700 entra como SC aqui)
    t["SALDO_ANT_TOT"] = (t["INI_SD_CAIXA"] + t["INI_SD_RPPS"]
                          + t["INI_SD_DEPR"] + t["INI_SD_VCOMP"]
                          + t["INI_SD_VTRAN"] + t["INI_SC_VTCR"])

    # 2.05 Saldo p/ o Exercício Seguinte = abertura + movimento do período.
    t["SALDO_SEG_TOT"] = sum(
        (t[f"INI_SD_{k}"] + t[f"MOV_SD_{k}"]) for k, _ in _GRUPOS)

    t["DESPESA_PAGA"] = t["DESPESA_MOV"] + t["DESPESA_ENC"]

    # Nomes esperados pelo cruzamentos.py. NÃO remover: o R7 lê
    # RECEITA_REALIZADA e o .get() dele devolve zero em silêncio se a chave
    # sumir, transformando um rename num "erro" de R$ 27,9 bi.
    t["RECEITA_REALIZADA"] = t["RECEITA"]
    t["TRANSF_RECEBIDAS"]  = t["TRANSF_RECEB"]
    t["TRANSF_CONCEDIDAS"] = t["TRANSF_CONC"]

    t["INGRESSOS"] = (t["RECEITA"] + t["TRANSF_RECEB"] + t["RESGATES"]
                      + t["INSCRICAO_RPNP"] + t["INSCRICAO_RPP"]
                      + t["DEPREST_ING"] + t["OUTROS_RECEB"]
                      + t["SALDO_ANT_TOT"])

    t["DISPENDIOS"] = (t["DESPESA_PAGA"] + t["TRANSF_CONC"] + t["APLICACOES"]
                       + t["PAG_RPNP"] + t["PAG_RPP"]
                       + t["DEPREST_DISP"] + t["AJUSTES"]
                       + t["SALDO_SEG_TOT"])
    return t


def auditar(conn, mes, ano, cogestao_list=None):
    t = extrair(conn, mes, ano, cogestao_list)
    achados = []

    # ── BF-01: equação completa. Agora DEVE fechar; divergência é ERRO. ─────
    gap = t["INGRESSOS"] - t["DISPENDIOS"]
    if checa_gap(gap, D("1.00")):
        achados.append(achado_ok("BF","BF-01","Ingressos = Dispêndios",
            f"Ingressos {t['INGRESSOS']:,.2f}  =  Dispêndios {t['DISPENDIOS']:,.2f}",
            valor=t["INGRESSOS"]))
    else:
        achados.append(achado_erro("BF","BF-01","Ingressos ≠ Dispêndios",
            f"Diferença: {gap:,.2f}  |  Ingressos {t['INGRESSOS']:,.2f}  "
            f"Dispêndios {t['DISPENDIOS']:,.2f}", valor=gap))

    # ── BF-02: Transferências Recebidas = Concedidas ───────────────────────
    gap2 = t["TRANSF_RECEB"] - t["TRANSF_CONC"]
    if checa_gap(gap2):
        achados.append(achado_ok("BF","BF-02","Transf. Recebidas = Transf. Concedidas",
            f"Recebidas {t['TRANSF_RECEB']:,.2f}  =  Concedidas {t['TRANSF_CONC']:,.2f}",
            valor=t["TRANSF_RECEB"]))
    else:
        achados.append(achado_alerta("BF","BF-02","Transf. Recebidas ≠ Transf. Concedidas",
            f"Diferença: {gap2:,.2f}", valor=gap2))

    # ── BF-03: Caixa Final positivo ────────────────────────────────────────
    if t["CAIXA_FINAL"] < 0:
        achados.append(achado_erro("BF","BF-03","Caixa Final negativo",
            f"Caixa Final: {t['CAIXA_FINAL']:,.2f}", valor=t["CAIXA_FINAL"]))
    else:
        achados.append(achado_ok("BF","BF-03","Caixa Final positivo",
            f"Caixa Final: {t['CAIXA_FINAL']:,.2f}", valor=t["CAIXA_FINAL"]))

    # ── BF-04 (novo): caixa 111XXXXXX pelas duas origens ───────────────────
    # CAIXA_FINAL usa SALDOCONTABIL acumulado (INMES 0..mes), definição do BP.
    # O Saldo Seguinte usa VSALDO(0) + LANCAMENTO(1..mes), definição do BF.
    # Restringindo à faixa 111XXXXXX, as duas têm que dar o mesmo número — é a
    # conciliação razão × view de saldos, o mesmo teste do DMPL-01.
    seg_111 = sum((t[f"INI_SD_{k}"] + t[f"MOV_SD_{k}"])
                  for k in ("CAIXA", "RPPS", "DEPR"))
    gap4 = seg_111 - t["CAIXA_FINAL"]
    if checa_gap(gap4, D("1.00")):
        achados.append(achado_ok("BF","BF-04","Caixa 111XXXXXX: razão = view de saldos",
            f"BF (abertura+lançamentos) {seg_111:,.2f}  =  "
            f"SALDOCONTABIL acumulado {t['CAIXA_FINAL']:,.2f}", valor=seg_111))
    else:
        achados.append(achado_erro("BF","BF-04","Caixa 111XXXXXX: razão ≠ view de saldos",
            f"Diferença: {gap4:,.2f}  |  BF (abertura+lançamentos) {seg_111:,.2f}  "
            f"vs SALDOCONTABIL acumulado {t['CAIXA_FINAL']:,.2f}", valor=gap4))

    # ── BF-05 (informativo): composição, p/ conferência contra o mestre ────
    achados.append(achado_info("BF","BF-05","Composição do Balanço Financeiro",
        f"Receita {t['RECEITA']:,.2f} | Transf.Receb {t['TRANSF_RECEB']:,.2f} | "
        f"Resgates {t['RESGATES']:,.2f} | Dep.Restit.Ing {t['DEPREST_ING']:,.2f} | "
        f"Outros Receb {t['OUTROS_RECEB']:,.2f} | Saldo Ant {t['SALDO_ANT_TOT']:,.2f} "
        f"|| Despesa {t['DESPESA_PAGA']:,.2f} | Transf.Conc {t['TRANSF_CONC']:,.2f} | "
        f"Aplicações {t['APLICACOES']:,.2f} | RPÑP {t['PAG_RPNP']:,.2f} | "
        f"RPP {t['PAG_RPP']:,.2f} | Dep.Restit.Disp {t['DEPREST_DISP']:,.2f} | "
        f"Ajustes {t['AJUSTES']:,.2f} | Saldo Seg {t['SALDO_SEG_TOT']:,.2f}",
        valor=t["INGRESSOS"]))

    return achados, t
