# -*- coding: utf-8 -*-
"""
Motor de regras do relatorio_integridade.py.

Reconstrução das regras originais (Regras_de_Integridade.txt, 13 regras) a
partir de leitura direta dos PDFs oficiais -- os módulos originais que
implementavam essas regras (motor_regras.py antigo) não existiam mais no
projeto e não foram encontrados no histórico do git; foram refeitas do
zero a partir da especificação em texto.

Acrescenta um segundo bloco de regras baseadas no MCASP (Manual de
Contabilidade Aplicada ao Setor Público, STN) e na LRF (Lei de
Responsabilidade Fiscal, LC 101/2000), fora do escopo das 13 originais:
Regra de Ouro, sequência Empenhado≥Liquidado≥Pago, estrutura do BP,
coerência do resultado orçamentário e do DFC.

NÃO reproduz as regras 18/27/31 (continuidade entre exercícios/meses) nem
a 32 (Notas Explicativas) do relatório antigo -- ambas exigiriam ler o PDF
do mês anterior e/ou um anexo de notas que não está publicado nos PDFs
mensais desta pasta. Ver TODO ao final do arquivo.
"""
from __future__ import annotations
from dataclasses import dataclass
from leitor_pdf_oficial import valores, pos, Demonstrativos, PdfNaoEncontrado

TOLERANCIA = 1.00
TOLERANCIA_CONSOLIDADO = 5.00


@dataclass
class Achado:
    regra: str
    titulo: str
    base_normativa: str
    status: str          # OK, DIVERGENTE, ALERTA, INFORMATIVO, NAO_VERIF
    valor_a: float | None = None
    rotulo_a: str = ""
    valor_b: float | None = None
    rotulo_b: str = ""
    observacao: str = ""

    @property
    def diferenca(self):
        if self.valor_a is None or self.valor_b is None:
            return None
        return self.valor_a - self.valor_b


def _cmp(regra, titulo, base, a, rot_a, b, rot_b, obs="", tol=TOLERANCIA):
    dif = a - b
    status = "OK" if abs(dif) < tol else "DIVERGENTE"
    return Achado(regra, titulo, base, status, a, rot_a, b, rot_b, obs)


def _safe(fn, *a, **kw):
    """Roda `fn`; se faltar um rótulo no PDF (layout mudou, item não
    existe naquele mês), devolve um achado NAO_VERIF em vez de quebrar o
    relatório inteiro."""
    try:
        return fn(*a, **kw)
    except (ValueError, PdfNaoEncontrado) as e:
        return [Achado(kw.get("regra", "?"), kw.get("titulo", fn.__name__),
                       "", "NAO_VERIF", observacao=str(e))]


# ─────────────────────────────────────────────────────────────────────────────
#  BLOCO I — Regras originais (Regras_de_Integridade.txt), reconstruídas
# ─────────────────────────────────────────────────────────────────────────────
def regra_01_resultado_patrimonial(d: Demonstrativos):
    dvp, bp, dmpl = d.t("DVP"), d.t("BP"), d.t("DMPL")
    vals_dvp, _ = valores(dvp, "RESULTADO PATRIMONIAL DO PERÍODO", 2)
    res_dvp = vals_dvp[0]
    (res_bp,), _ = valores(bp, "Resultado do Exercício", 1)
    out = [
        _cmp("1", "Resultado Patrimonial: DVP = BP", "MCASP Pte.II, DVP x BP",
             res_dvp, "DVP - Resultado Patrimonial (III)", res_bp, "BP - Resultado do Exercício"),
    ]
    return out


def regra_01b_dmpl_resultado(d: Demonstrativos):
    dvp, dmpl = d.t("DVP"), d.t("DMPL")
    vals_dvp, _ = valores(dvp, "RESULTADO PATRIMONIAL DO PERÍODO", 2)
    res_dvp = vals_dvp[0]
    vals, _ = valores(dmpl, "Resultado do exercício", 9)
    res_dmpl_total = vals[-1]  # coluna TOTAL (9ª)
    return [_cmp("1b", "Resultado do Exercício: DVP = DMPL (coluna TOTAL)",
                 "MCASP Pte.II, DVP x DMPL",
                 res_dvp, "DVP - Resultado Patrimonial (III)",
                 res_dmpl_total, "DMPL - Resultado do exercício (TOTAL)")]


def regra_02_saldo_final_dmpl_pl_bp(d: Demonstrativos):
    dmpl, bp = d.t("DMPL"), d.t("BP")
    vals, _ = valores(dmpl, "Saldos finais", 9)
    saldo_final_dmpl = vals[-1]
    # "PATRIMÔNIO LÍQUIDO" é substring de "PASSIVO E PATRIMÔNIO LÍQUIDO"
    # (que aparece antes, com o total do lado do passivo) -- ancora a
    # busca depois de "PASSIVO NÃO CIRCULANTE" para pegar a linha certa.
    i_pl = pos(bp, "PASSIVO NÃO CIRCULANTE")
    (pl_bp,), _ = valores(bp, "PATRIMÔNIO LÍQUIDO", 1, inicio=i_pl)
    return [_cmp("2", "Saldo Final da DMPL = Patrimônio Líquido do BP",
                 "MCASP Pte.II, DMPL x BP",
                 saldo_final_dmpl, "DMPL - Saldos finais (TOTAL)",
                 pl_bp, "BP - Patrimônio Líquido",
                 tol=TOLERANCIA_CONSOLIDADO)]


DMPL_COLUNAS = [
    ("Patrimônio Social e Capital Social", 0),
    ("Adiantamento para Futuro Aumento de Capital", 1),
    ("Reservas de Capital", 2),
    ("Ajustes de Avaliação Patrimonial", 3),
    ("Reservas de Lucros", 4),
    ("Demais Reservas", 5),
    ("Resultado Acumulado", 6),
]


def regra_03_colunas_dmpl_linhas_pl(d: Demonstrativos):
    dmpl, bp = d.t("DMPL"), d.t("BP")
    vals, _ = valores(dmpl, "Saldos finais", 9)
    achados = []
    for rotulo_bp, idx_dmpl in DMPL_COLUNAS:
        try:
            (v_bp,), _ = valores(bp, rotulo_bp, 1)
        except ValueError as e:
            achados.append(Achado("3", f"Coluna DMPL x linha BP — {rotulo_bp}",
                                  "MCASP Pte.II, DMPL x BP", "NAO_VERIF",
                                  observacao=str(e)))
            continue
        v_dmpl = vals[idx_dmpl]
        achados.append(_cmp("3", f"Coluna DMPL x linha BP — {rotulo_bp}",
                            "MCASP Pte.II, DMPL x BP",
                            v_dmpl, f"DMPL - {rotulo_bp}", v_bp, f"BP - {rotulo_bp}"))
    return achados


def regra_04_caixa_bp_dfc_bf(d: Demonstrativos):
    bp, dfc, bf = d.t("BP"), d.t("DFC"), d.t("BF")
    (caixa_bp,), _ = valores(bp, "Caixa e Equivalentes de Caixa", 1)
    (caixa_dfc,), _ = valores(dfc, "Caixa e Equivalente de Caixa Final", 1)
    i_seg = pos(bf, "SALDO PARA O EXERCÍCIO SEGUINTE")
    (caixa_bf_ex_rpps,), _ = valores(bf, "Caixa e Equivalentes de Caixa (Exceto RPPS)", 1, inicio=i_seg)
    (caixa_bf_rpps,), _ = valores(bf, "CAIXA E EQUIVALENTES DE CAIXA RPPS", 1, inicio=i_seg)
    (dep_rest_bf,), _ = valores(bf, "Depósitos Restituíveis e Valores Vinculados", 1, inicio=i_seg)
    # A conciliação só fecha exato incluindo Depósitos Restituíveis e
    # Valores Vinculados (confirmado empiricamente em 18/08/2026, mês
    # 07/2026: sem esse componente sobrava uma diferença de R$ 8.457.361,41
    # -- exatamente o valor dessa linha do BF).
    caixa_bf_tot = caixa_bf_ex_rpps + caixa_bf_rpps + dep_rest_bf
    return [
        _cmp("4a", "Caixa e Equiv. de Caixa: BP = DFC (Caixa Final)",
             "MCASP Pte.II, BP x DFC", caixa_bp, "BP - Caixa e Equiv. de Caixa",
             caixa_dfc, "DFC - Caixa e Equivalente de Caixa Final"),
        _cmp("4b", "Caixa e Equiv. de Caixa: BP = BF (Caixa+RPPS+Dep.Restituíveis, saldo seguinte)",
             "MCASP Pte.II, BP x BF", caixa_bp, "BP - Caixa e Equiv. de Caixa",
             caixa_bf_tot, "BF - Caixa (Exceto RPPS) + Caixa RPPS + Dep. Restituíveis"),
    ]


def regra_05_ativo_anexo_bp(d: Demonstrativos):
    bp = d.t("BP")
    i = pos(bp, "QUADRO DOS ATIVOS E PASSIVOS FINANCEIROS E PERMANENTES")
    (ativo_i,), _ = valores(bp, "ATIVO ( I )", 1, inicio=i)
    (ativo_fin,), _ = valores(bp, "Ativo Financeiro", 1, inicio=i)
    (ativo_perm,), _ = valores(bp, "Ativo Permanente", 1, inicio=i)
    (ativo_total,), _ = valores(bp, "ATIVO ", 1)
    return [_cmp("5a", "Ativo do Anexo (Financeiro+Permanente) = Ativo Total do BP",
                 "MCASP Pte.II, Anexo do BP",
                 ativo_fin + ativo_perm, "Anexo BP - Ativo Financeiro + Permanente",
                 ativo_total, "BP - ATIVO (corpo do balanço)")]


def regra_06_superavit_financeiro(d: Demonstrativos):
    bp = d.t("BP")
    i = pos(bp, "QUADRO DOS ATIVOS E PASSIVOS FINANCEIROS E PERMANENTES")
    (ativo_fin,), _ = valores(bp, "Ativo Financeiro", 1, inicio=i)
    (passivo_fin,), _ = valores(bp, "Passivo Financeiro", 1, inicio=i)
    (saldo_patr,), _ = valores(bp, "SALDO PATRIMONIAL (III) = (I - II)", 1, inicio=i)
    calc = ativo_fin - passivo_fin
    return [Achado("6", "Superávit Financeiro = Ativo Financeiro − Passivo Financeiro",
                   "Art. 43 §2º LRF; Anexo do BP", "INFORMATIVO", calc,
                   "Anexo BP - Ativo Financeiro − Passivo Financeiro",
                   observacao=f"Só é o Superávit Financeiro oficial (para "
                              f"abertura de crédito adicional) no fechamento "
                              f"do exercício (dezembro), por fonte de "
                              f"recursos — aqui é o total consolidado, "
                              f"informativo. Saldo Patrimonial (III) do "
                              f"anexo, para referência: {saldo_patr:,.2f}")]


def regra_07_receita_bf_bo(d: Demonstrativos):
    bf, bo = d.t("BF"), d.t("BO")
    vals_bf, _ = valores(bf, "RECEITA ORÇAMENTÁRIA", 2)
    rec_bf = vals_bf[0]
    i = pos(bo, "SUBTOTAL DAS RECEITAS (III) = (I + II)")
    vals, _ = valores(bo, "SUBTOTAL DAS RECEITAS (III) = (I + II)", 4, inicio=i)
    rec_bo = vals[2]  # coluna (c) Receitas Realizadas
    return [_cmp("7", "Receita Orçamentária: BF = Receitas Realizadas do BO",
                 "MCASP Pte.II, BF x BO",
                 rec_bf, "BF - Receita Orçamentária", rec_bo, "BO - Receitas Realizadas")]


def regra_08_despesa_bf_bo(d: Demonstrativos):
    bf, bo = d.t("BF"), d.t("BO")
    vals_bf, _ = valores(bf, "DESPESA ORÇAMENTÁRIA", 2)
    desp_bf = vals_bf[0]
    i = pos(bo, "SUBTOTAL DAS DESPESAS")
    vals, _ = valores(bo, "SUBTOTAL DAS DESPESAS", 6, inicio=i)
    desp_bo_paga = vals[4]  # coluna (i) Despesas Pagas
    return [_cmp("8", "Despesa Orçamentária: BF = Despesas Pagas do BO",
                 "MCASP Pte.II, BF x BO",
                 desp_bf, "BF - Despesa Orçamentária", desp_bo_paga, "BO - Despesas Pagas")]


def regra_10_pagamentos_rp(d: Demonstrativos):
    bf, bo = d.t("BF"), d.t("BO")
    vals_a, _ = valores(bf, "Pagamentos de RPÑP", 2)
    pag_rpnp_bf = vals_a[0]
    vals_b, _ = valores(bf, "Pagamentos de RPP", 2)
    pag_rpp_bf = vals_b[0]
    i1 = pos(bo, "EXECUÇÃO DE RESTOS A PAGAR NÃO PROCESSADOS")
    vals_np, _ = valores(bo, "TOTAL", 6, inicio=i1)
    pag_rpnp_bo = vals_np[3]  # coluna (d) Pagos
    i2 = pos(bo, "EXECUÇÃO DE RESTOS A PAGAR PROCESSADOS")
    vals_p, _ = valores(bo, "TOTAL", 5, inicio=i2)
    pag_rpp_bo = vals_p[2]  # coluna (c) Pagos
    return [
        _cmp("10a", "Pagamento de RPÑP (não processados): BF = BO",
             "MCASP Pte.II, BF x BO (Anexo RP)",
             pag_rpnp_bf, "BF - Pagamentos de RPÑP", pag_rpnp_bo, "BO - RP Não Proc. Pagos"),
        _cmp("10b", "Pagamento de RPP (processados): BF = BO",
             "MCASP Pte.II, BF x BO (Anexo RP)",
             pag_rpp_bf, "BF - Pagamentos de RPP", pag_rpp_bo, "BO - RP Processados Pagos"),
    ]


ITENS_RECEITA_BO_DFC = [
    ("Impostos, Taxas e Contribuições de Melhoria", "IMPOSTOS, TAXAS E CONTRIBUIÇÕES DE MELHORIA"),
    ("Receita de Contribuições", "RECEITA CONTRIBUIÇÕES"),
    ("Receita Patrimonial", "RECEITA PATRIMONIAL"),
    ("Receita Agropecuária", "RECEITA AGROPECUÁRIA"),
    ("Receita Industrial", "RECEITA INDUSTRIAL"),
    ("Receita de Serviços", "RECEITA DE SERVIÇOS"),
]


def regra_11_receitas_bo_dfc(d: Demonstrativos):
    bo, dfc = d.t("BO"), d.t("DFC")
    achados = []
    for rot_dfc, rot_bo in ITENS_RECEITA_BO_DFC:
        try:
            (v_dfc,), _ = valores(dfc, rot_dfc, 1)
            vals_bo, _ = valores(bo, rot_bo, 3)
            v_bo = vals_bo[2]
        except ValueError as e:
            achados.append(Achado("11", f"Receitas: BO x DFC — {rot_dfc}",
                                  "MCASP Pte.II, BO x DFC", "NAO_VERIF", observacao=str(e)))
            continue
        achados.append(_cmp("11", f"Receitas: BO x DFC — {rot_dfc}",
                            "MCASP Pte.II, BO x DFC",
                            v_bo, f"BO - {rot_bo} (Realizadas)", v_dfc, f"DFC - {rot_dfc}"))
    return achados


def regra_12a_superavit_orcamentario(d: Demonstrativos):
    bo = d.t("BO")
    i2 = pos(bo, "SUBTOTAL DAS RECEITAS (III) = (I + II)")
    vals_rec, _ = valores(bo, "SUBTOTAL DAS RECEITAS (III) = (I + II)", 4, inicio=i2)
    delta_prev = vals_rec[1] - vals_rec[0]  # Atualizada - Inicial
    i3 = pos(bo, "SUBTOTAL DAS DESPESAS")
    vals_desp, _ = valores(bo, "SUBTOTAL DAS DESPESAS", 6, inicio=i3)
    delta_dot = vals_desp[1] - vals_desp[0]  # Atualizada - Inicial
    calc = delta_dot - delta_prev
    try:
        (superavit_xiii,), _ = valores(bo, "SUPERAVIT ( XIII )", 1)
    except ValueError:
        superavit_xiii = 0.0
    return [Achado("12a", "Superávit Orçamentário = ΔDotação(Despesa) − ΔPrevisão(Receita)",
                   "Art. 4º LRF; MCASP Pte.II Anexo 12", "INFORMATIVO", calc,
                   "Calculado (ΔDotação − ΔPrevisão)", superavit_xiii,
                   "BO - SUPERÁVIT (XIII) reportado",
                   observacao="Divergência grande é esperada quando o ajuste é "
                              "alocado via Saldos de Exercícios Anteriores / "
                              "Superávit Financeiro, não na linha SUPERÁVIT "
                              "(XIII) — ver regra original nº12a.")]


def regra_12b_variacao_dotacao(d: Demonstrativos):
    bo = d.t("BO")
    i3 = pos(bo, "SUBTOTAL DAS DESPESAS")
    vals_desp, _ = valores(bo, "SUBTOTAL DAS DESPESAS", 6, inicio=i3)
    var_dotacao = vals_desp[1] - vals_desp[0]
    # Ancora depois da última linha de dados ("RESERVA DE CONTIGÊNCIA",
    # grafia oficial sem o 'n') -- a palavra "TOTAL" aparece antes disso
    # dentro do próprio cabeçalho da coluna ("TOTAL DE ALTERAÇÕES"), o que
    # pegava um valor errado se buscado logo após o título do quadro.
    i4 = pos(bo, "RESERVA DE CONTIGÊNCIA")
    vals_creditos, _ = valores(bo, "TOTAL", 8, inicio=i4)
    total_creditos = vals_creditos[-1]  # coluna (h) = Total de Alterações
    return [_cmp("12b", "Variação da Dotação (Atualizada−Inicial) = Total de Créditos Adicionais",
                 "MCASP Pte.II Anexo 12",
                 var_dotacao, "BO - Dotação Atualizada − Inicial",
                 total_creditos, "BO - Total de Alterações (Créditos Adicionais)",
                 tol=TOLERANCIA_CONSOLIDADO)]


# ─────────────────────────────────────────────────────────────────────────────
#  BLOCO II — MCASP / LRF (não estavam nas 13 regras originais)
# ─────────────────────────────────────────────────────────────────────────────
def regra_14_regra_de_ouro(d: Demonstrativos):
    """Regra de Ouro (CF/88 art. 167, III): operações de crédito não podem
    superar as despesas de capital, salvo créditos suplementares/especiais
    com finalidade precisa autorizados por maioria absoluta do legislativo."""
    bo = d.t("BO")
    vals_credito, _ = valores(bo, "OPERAÇÕES DE CRÉDITO", 3)
    op_credito = vals_credito[2]
    i3 = pos(bo, "DESPESAS DE CAPITAL (IX)")
    vals_cap, _ = valores(bo, "DESPESAS DE CAPITAL (IX)", 6, inicio=i3)
    despesa_capital_empenhada = vals_cap[2]  # coluna (g) Empenhadas
    dif = op_credito - despesa_capital_empenhada
    status = "OK" if dif <= TOLERANCIA_CONSOLIDADO else "DIVERGENTE"
    return [Achado("R.OURO", "Regra de Ouro: Operações de Crédito ≤ Despesas de Capital",
                   "CF/88 art. 167, III", status,
                   op_credito, "BO - Operações de Crédito (Receita Realizada)",
                   despesa_capital_empenhada, "BO - Despesas de Capital (Empenhadas)",
                   observacao="Operações de crédito não podem superar despesas "
                              "de capital, salvo créditos suplementares/especiais "
                              "com finalidade precisa aprovados por maioria "
                              "absoluta do Legislativo (ressalva constitucional "
                              "não verificável automaticamente aqui).")]


def regra_15_empenhado_liquidado_pago(d: Demonstrativos):
    bo = d.t("BO")
    blocos = [
        ("DESPESAS CORRENTES (VIII)", "Despesas Correntes"),
        ("DESPESAS DE CAPITAL (IX)", "Despesas de Capital"),
        ("SUBTOTAL DAS DESPESAS", "Subtotal das Despesas"),
    ]
    achados = []
    for rot, nome in blocos:
        try:
            i = pos(bo, rot)
            vals, _ = valores(bo, rot, 6, inicio=i)
        except ValueError as e:
            achados.append(Achado("15", f"Empenhado ≥ Liquidado ≥ Pago — {nome}",
                                  "MCASP Pte.II, Anexo 12", "NAO_VERIF", observacao=str(e)))
            continue
        emp, liq, pago = vals[2], vals[3], vals[4]
        ok = (emp + TOLERANCIA >= liq) and (liq + TOLERANCIA >= pago)
        status = "OK" if ok else "DIVERGENTE"
        achados.append(Achado("15", f"Empenhado ≥ Liquidado ≥ Pago — {nome}",
                              "Execução orçamentária (Lei 4.320/64 arts. 58-65)",
                              status, emp, f"{nome} - Empenhada", liq, f"{nome} - Liquidada",
                              observacao=f"Paga: {pago:,.2f}. "
                                        f"{'Sequência respeitada.' if ok else 'Sequência violada — investigar.'}"))
    return achados


def regra_16_resultado_orcamentario(d: Demonstrativos):
    bo = d.t("BO")
    i2 = pos(bo, "SUBTOTAL DAS RECEITAS (III) = (I + II)")
    vals_rec, _ = valores(bo, "SUBTOTAL DAS RECEITAS (III) = (I + II)", 4, inicio=i2)
    receita_realizada = vals_rec[2]
    i3 = pos(bo, "SUBTOTAL DAS DESPESAS")
    vals_desp, _ = valores(bo, "SUBTOTAL DAS DESPESAS", 6, inicio=i3)
    despesa_empenhada = vals_desp[2]
    resultado = receita_realizada - despesa_empenhada
    sinal = "superavitário" if resultado >= 0 else "deficitário"
    return [Achado("16", "Resultado Orçamentário = Receitas Realizadas − Despesas Empenhadas",
                   "MCASP Pte.II, Anexo 12", "INFORMATIVO", receita_realizada,
                   "BO - Receitas Realizadas", despesa_empenhada, "BO - Despesas Empenhadas",
                   observacao=f"Resultado {sinal}: {resultado:,.2f}")]


def regra_17_dfc_geracao_liquida(d: Demonstrativos):
    dfc = d.t("DFC")
    (op,), _ = valores(dfc, "Fluxo de Caixa Líq. das Atividades Operacionais(I)", 1)
    (inv,), _ = valores(dfc, "Fluxo de Caixa Líq. das Ativ. de Investimento (II)", 1)
    (fin,), _ = valores(dfc, "Fluxo de Caixa Líq.das Ativ. Financiamento (III)", 1)
    (geracao,), _ = valores(dfc, "GERAÇAO LÍQUIDA DE CAIXA E EQUIVALENTE (I+II+III)", 1)
    (caixa_ini,), _ = valores(dfc, "Caixa e Equivalente de Caixa Inicial", 1)
    (caixa_fim,), _ = valores(dfc, "Caixa e Equivalente de Caixa Final", 1)
    calc_soma = op + inv + fin
    calc_var = caixa_fim - caixa_ini
    return [
        _cmp("17a", "DFC: Geração Líquida = Operacional + Investimento + Financiamento",
             "MCASP Pte.II, Anexo 15", calc_soma, "DFC - Soma dos 3 fluxos",
             geracao, "DFC - Geração Líquida reportada"),
        _cmp("17b", "DFC: Geração Líquida = Caixa Final − Caixa Inicial",
             "MCASP Pte.II, Anexo 15", geracao, "DFC - Geração Líquida reportada",
             calc_var, "DFC - Caixa Final − Caixa Inicial"),
    ]


def regra_18_estrutura_bp(d: Demonstrativos):
    bp = d.t("BP")
    (ativo,), _ = valores(bp, "ATIVO ", 1)
    (ativo_circ,), _ = valores(bp, "ATIVO CIRCULANTE", 1)
    (ativo_ncirc,), _ = valores(bp, "ATIVO NÃO CIRCULANTE", 1)
    (passivo_pl,), _ = valores(bp, "PASSIVO E PATRIMÔNIO LÍQUIDO", 1)
    (passivo_circ,), _ = valores(bp, "PASSIVO CIRCULANTE", 1)
    i_pncirc = pos(bp, "PASSIVO NÃO CIRCULANTE")
    (passivo_ncirc,), _ = valores(bp, "PASSIVO NÃO CIRCULANTE", 1)
    (pl,), _ = valores(bp, "PATRIMÔNIO LÍQUIDO", 1, inicio=i_pncirc)
    return [
        _cmp("18a", "BP: Ativo Circulante + Não Circulante = Ativo Total",
             "MCASP Pte.II, Anexo 1 do BP", ativo_circ + ativo_ncirc,
             "Soma AC+ANC", ativo, "BP - ATIVO"),
        _cmp("18b", "BP: Ativo = Passivo + Patrimônio Líquido",
             "MCASP Pte.II, Anexo 1 do BP (identidade contábil fundamental)",
             ativo, "BP - ATIVO", passivo_pl, "BP - PASSIVO E PATRIMÔNIO LÍQUIDO"),
        _cmp("18c", "BP: Passivo Circulante + Não Circulante + PL = Passivo e PL Total",
             "MCASP Pte.II, Anexo 1 do BP", passivo_circ + passivo_ncirc + pl,
             "Soma PC+PNC+PL", passivo_pl, "BP - PASSIVO E PATRIMÔNIO LÍQUIDO"),
    ]


TODAS_AS_REGRAS = [
    regra_01_resultado_patrimonial,
    regra_01b_dmpl_resultado,
    regra_02_saldo_final_dmpl_pl_bp,
    regra_03_colunas_dmpl_linhas_pl,
    regra_04_caixa_bp_dfc_bf,
    regra_05_ativo_anexo_bp,
    regra_06_superavit_financeiro,
    regra_07_receita_bf_bo,
    regra_08_despesa_bf_bo,
    regra_10_pagamentos_rp,
    regra_11_receitas_bo_dfc,
    regra_12a_superavit_orcamentario,
    regra_12b_variacao_dotacao,
    regra_14_regra_de_ouro,
    regra_15_empenhado_liquidado_pago,
    regra_16_resultado_orcamentario,
    regra_17_dfc_geracao_liquida,
    regra_18_estrutura_bp,
]


def run_all_rules(d: Demonstrativos) -> list[Achado]:
    achados = []
    for fn in TODAS_AS_REGRAS:
        try:
            achados.extend(fn(d))
        except (ValueError, PdfNaoEncontrado) as e:
            achados.append(Achado(fn.__name__, fn.__doc__ or fn.__name__,
                                  "", "NAO_VERIF", observacao=str(e)))
    return achados


# TODO (fora do escopo desta rodada — exigem PDF do mês anterior ou anexo
# de Notas Explicativas, não presentes na pasta mensal):
#   - Regra 9 (Inscrição de RPNP/RPP: BF = BO) -- só faz sentido no
#     fechamento de dezembro (inscrição de RP só ocorre no encerramento).
#   - Regras 18/27/31 do relatório antigo (continuidade: saldo final de um
#     mês = saldo inicial do mês seguinte) -- exige ler 2 meses e comparar.
#   - Regra 32 (achados extraídos das Notas Explicativas) -- as Notas
#     Explicativas não aparecem nos PDFs mensais desta pasta, só no anexo
#     de encerramento anual.
