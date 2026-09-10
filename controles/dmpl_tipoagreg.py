# -*- coding: utf-8 -*-
"""
Controles da Demonstração das Mutações do Patrimônio Líquido (DMPL).

Estratégia: extrai apenas os totais necessários para os cruzamentos com os
demais demonstrativos, MAS preservando a natureza do controle do dmpl.py
mestre: a identidade de fechamento cruza DUAS fontes distintas.

    SALDO_INI(23x)  [SALDOCONTABIL, INMES = 0]
  + MOVIMENTO(23x)  [LANCAMENTOCONTABIL, INMES 1..mes]
  = SALDO_FIM(23x)  [SALDOCONTABIL, INMES 0..mes]

Isso concilia o razão contra a view de saldos. A versão anterior deste
arquivo tirava os três termos do SALDOCONTABIL, o que tornava a
identidade uma TAUTOLOGIA (INMES 0..mes é a união disjunta de INMES 0 com
INMES 1..mes) — o controle passava sempre, inclusive com dados corrompidos.

O Resultado do Exercício aparece nos dois lados da equação e se cancela;
ele NÃO é validado aqui, e sim pelo cruzamento X6 (DMPL = DVP).

Faixas de conta idênticas às do dmpl.py mestre (COLUNAS): 231, 232, 233,
23411, 235, 236, 237. Atenção: 234 fora de 23411 fica de fora de
propósito, acompanhando o mestre.
"""
from __future__ import annotations
from . import (Achado, query_one, D,
               achado_ok, achado_erro, achado_alerta, achado_info, checa_gap)

# Cópia de dmpl.py — filtro por tipo de agregação de gestão. NÃO mexer no original.
def _qf(conn, sql, cogestao_list, alias, campo="COGESTAO", **fmt):
    s = sql.format(**fmt) if fmt else sql
    if cogestao_list:
        ids = ",".join(str(int(c)) for c in cogestao_list)
        kw = "AND" if "WHERE" in s.upper() else "WHERE"
        s += f"\n  {kw} {alias}.{campo} IN ({ids})"
    return query_one(conn, s)

# (chave, máscara) — espelha COLUNAS do dmpl.py mestre.
# Coluna 8 (Ações/Cotas em Tesouraria) não tem conta mapeada no plano de
# contas do ente; mantida ausente, igual ao mestre.
COLUNAS_PL = [
    ("PAT_SOCIAL_CAPITAL", "231XXXXXX", "Pat. Social / Capital Social"),
    ("AFAC",               "232XXXXXX", "AFAC"),
    ("RESERVA_CAPITAL",    "233XXXXXX", "Reserva de Capital"),
    ("AJUSTE_AVAL_PATRIM", "23411XXXX", "Ajustes de Avaliação Patrimonial"),
    ("RESERVAS_LUCROS",    "235XXXXXX", "Reservas de Lucros"),
    ("DEMAIS_RESERVAS",    "236XXXXXX", "Demais Reservas"),
    ("RESULTADOS_ACUM",    "237XXXXXX", "Resultados Acumulados"),
]

TOLERANCIA = D("1.00")


def _faixa(mascara: str) -> tuple[int, int]:
    """'231XXXXXX' -> (231000000, 231999999). Mesmo padrão do mestre."""
    n_x = mascara.count("X")
    ini = int(mascara.replace("X", "")) * (10 ** n_x)
    return ini, ini + (10 ** n_x) - 1


def _cond(alias_tab: str, mascara: str) -> str:
    ini, fim = _faixa(mascara)
    if ini == fim:
        return f"{alias_tab}.COCONTACONTABIL = {ini}"
    return f"{alias_tab}.COCONTACONTABIL BETWEEN {ini} AND {fim}"


def _montar_sql_saldo(mes: int) -> str:
    """Saldo inicial e saldo final por coluna, via SALDOCONTABIL."""
    partes = []
    for chave, mascara, _ in COLUNAS_PL:
        cond = _cond("v", mascara)
        partes.append(
            f"    SUM(CASE WHEN v.INMES = 0 AND {cond}\n"
            f"         THEN v.VACREDITO - v.VADEBITO ELSE 0 END) AS INI_{chave}")
        partes.append(
            f"    SUM(CASE WHEN v.INMES BETWEEN 0 AND {mes} AND {cond}\n"
            f"         THEN v.VACREDITO - v.VADEBITO ELSE 0 END) AS FIM_{chave}")
    corpo = ",\n\n".join(partes)
    return (f"SELECT\n{corpo}\n"
            f"FROM MIL{{ano}}.SALDOCONTABIL v\n"
            f"WHERE v.COCONTACONTABIL BETWEEN 200000000 AND 299999999")


def _montar_sql_lanc(mes: int) -> str:
    """Movimento do período por coluna + Resultado do Exercício,
    via LANCAMENTOCONTABIL (INMES 1..mes) — mesma fonte do mestre."""
    sc = "DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)"
    sd = "DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)"
    partes = []
    for chave, mascara, _ in COLUNAS_PL:
        partes.append(
            f"    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes} AND {_cond('o', mascara)}\n"
            f"         THEN {sc} ELSE 0 END) AS VAR_{chave}")
    # Resultado do Exercício = VPA - VPD (classes 4 e 3), igual ao mestre.
    partes.append(
        f"    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}\n"
        f"              AND o.COCONTACONTABIL BETWEEN 400000000 AND 499999999\n"
        f"         THEN {sc} ELSE 0 END) AS RES_VPA")
    partes.append(
        f"    SUM(CASE WHEN o.INMES BETWEEN 1 AND {mes}\n"
        f"              AND o.COCONTACONTABIL BETWEEN 300000000 AND 399999999\n"
        f"         THEN {sd} ELSE 0 END) AS RES_VPD")
    corpo = ",\n\n".join(partes)
    return (f"SELECT\n{corpo}\n"
            f"FROM MIL{{ano}}.LANCAMENTOCONTABIL o\n"
            f"WHERE (o.COCONTACONTABIL BETWEEN 200000000 AND 299999999\n"
            f"       OR o.COCONTACONTABIL BETWEEN 300000000 AND 499999999)")


def extrair(conn, mes: int, ano: int, cogestao_list=None) -> dict:
    t = {}
    # SALDOCONTABIL → COGESTAO (default); LANCAMENTOCONTABIL → COGESTAOCONTAB (= critério PSIAG)
    t.update(_qf(conn, _montar_sql_saldo(mes), cogestao_list, "v", ano=ano))
    t.update(_qf(conn, _montar_sql_lanc(mes), cogestao_list, "o", campo="COGESTAOCONTAB", ano=ano))
    t = {k: (D(0) if v is None else v) for k, v in t.items()}

    pl_ini = sum(t[f"INI_{c}"] for c, _, _ in COLUNAS_PL)
    pl_fim = sum(t[f"FIM_{c}"] for c, _, _ in COLUNAS_PL)
    var_pl = sum(t[f"VAR_{c}"] for c, _, _ in COLUNAS_PL)
    resultado = t["RES_VPA"] - t["RES_VPD"]

    t["PL_SALDO_INI"] = pl_ini
    t["PL_SALDO_FIM"] = pl_fim
    t["VARIACAO_PL_PERIOD"] = var_pl
    t["RESULTADO_EXERCICIO"] = resultado

    # PL_FINAL para cruzamento com o BP (R2): em meses intermediários o
    # resultado do período ainda não foi encerrado contra a 237, então
    # precisa ser somado. Mesma composição da linha SALDOS_FINAIS do mestre.
    t["PL_FINAL"] = pl_fim + resultado
    return t


def auditar(conn, mes: int, ano: int, cogestao_list=None) -> tuple[list[Achado], dict]:
    t = extrair(conn, mes, ano, cogestao_list)
    achados: list[Achado] = []

    pl_ini = t["PL_SALDO_INI"]
    pl_fim = t["PL_SALDO_FIM"]
    var_pl = t["VARIACAO_PL_PERIOD"]
    resultado = t["RESULTADO_EXERCICIO"]

    # ── DMPL-01: conciliação razão × saldos (agregado) ────────────────────
    calc = pl_ini + var_pl
    gap = calc - pl_fim
    if checa_gap(gap, TOLERANCIA):
        achados.append(achado_ok("DMPL", "DMPL-01",
            "Fechamento DMPL: Saldo Ini + Movimento (razão) = Saldo Fin",
            f"Ini {pl_ini:,.2f}  +  Mov. {var_pl:,.2f}  =  {calc:,.2f}  "
            f"=  Fin {pl_fim:,.2f}"))
    else:
        # Localiza a(s) coluna(s) responsável(is) — o mestre faz isso sempre;
        # aqui só quando quebra, para manter o diagnóstico enxuto.
        culpadas = []
        for chave, _, desc in COLUNAS_PL:
            d = t[f"INI_{chave}"] + t[f"VAR_{chave}"] - t[f"FIM_{chave}"]
            if not checa_gap(d, TOLERANCIA):
                culpadas.append(f"{desc}: {d:,.2f}")
        detalhe_col = "  |  ".join(culpadas) if culpadas else "sem coluna isolada"
        achados.append(achado_erro("DMPL", "DMPL-01",
            "Fechamento DMPL não fecha (razão ≠ saldos)",
            f"Diferença: {gap:,.2f}  |  Calculado {calc:,.2f}  vs  "
            f"Saldo Final {pl_fim:,.2f}  ||  {detalhe_col}",
            valor=gap))

    # ── DMPL-02: mutações do PL não decorrentes do resultado ──────────────
    # var_pl é movimento em 23x; o resultado ainda não foi encerrado para a
    # 237 em mês intermediário, então este valor deveria ter origem
    # identificável (ajustes de exercícios anteriores, reservas, avaliação
    # patrimonial). Informativo: materialidade é julgamento do contador.
    achados.append(achado_info("DMPL", "DMPL-02",
        "Mutações do PL no período (exceto Resultado do Exercício)",
        f"{var_pl:,.2f}  — identificar origem: ajustes de exerc. anteriores, "
        f"aumento de capital, reservas, avaliação patrimonial"))

    # ── DMPL-03: Resultado do Exercício (cruzado com a DVP via X6) ────────
    achados.append(achado_info("DMPL", "DMPL-03",
        "Resultado do Exercício (DMPL)",
        f"{resultado:,.2f}  — deve igualar o Resultado Patrimonial da DVP "
        f"(controlado pelo cruzamento X6)"))

    return achados, t
