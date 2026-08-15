# -*- coding: utf-8 -*-
"""
Validações cruzadas entre os demonstrativos (regras X1..X13).

Recebe os dicionários de totais já extraídos por cada módulo e roda
as comparações, sem tocar no banco de dados.
"""
from __future__ import annotations
from decimal import Decimal
from . import (Achado, D, achado_ok, achado_erro,
               achado_alerta, achado_info, checa_gap)


def auditar(t_bf: dict, t_bo: dict, t_dfc: dict,
            t_dvp: dict, t_dmpl: dict, t_bp: dict) -> list[Achado]:
    """
    Valida as regras cruzadas entre os seis demonstrativos.
    Qualquer módulo ausente (None) tem seus controles pulados com INFO.
    """
    achados = []

    def _v(t: dict, chave: str, cod: str):
        """Lê uma chave obrigatória do dicionário do módulo.

        Antes isto era t.get(chave, D(0)). Uma chave ausente virava zero em
        silêncio e o cruzamento reportava um erro contábil convincente: em
        13/08/2026 um rename de RECEITA_REALIZADA para RECEITA fez o R7
        acusar diferença de R$ 27,9 bi, que era o valor inteiro do outro
        lado. Erro de programação não pode se disfarçar de achado contábil.
        """
        if chave not in t:
            raise KeyError(
                f"{cod}: chave '{chave}' ausente no dicionário do módulo — "
                f"o extrair() do módulo mudou de nome ou não a preenche. "
                f"Chaves disponíveis: {sorted(t)[:12]}...")
        return t[chave]

    def skip(cod, titulo, motivo):
        achados.append(achado_info("X", cod,
            titulo, f"Pulado — {motivo} não extraído."))

    # ── X1a. Caixa Inicial: BF = DFC ────────────────────────────────────────
    if t_bf and t_dfc:
        bf_ci  = _v(t_bf, "CAIXA_INICIAL", "X1a")
        dfc_ci = _v(t_dfc, "CAIXA_INI", "X1a")
        gap = bf_ci - dfc_ci
        if checa_gap(gap):
            achados.append(achado_ok("X", "X1a",
                "Caixa Inicial: BF = DFC",
                f"BF {bf_ci:,.2f}  =  DFC {dfc_ci:,.2f}", valor=bf_ci))
        else:
            achados.append(achado_erro("X", "X1a",
                "Caixa Inicial: BF ≠ DFC",
                f"Diferença: {gap:,.2f}  |  BF {bf_ci:,.2f}  DFC {dfc_ci:,.2f}",
                valor=gap))
    else:
        skip("X1a", "Caixa Inicial BF=DFC", "BF ou DFC")

    # ── X1b. Caixa Final: BF = DFC ──────────────────────────────────────────
    if t_bf and t_dfc:
        bf_cf  = _v(t_bf, "CAIXA_FINAL", "X1b")
        dfc_cf = _v(t_dfc, "CAIXA_FIN", "X1b")
        gap = bf_cf - dfc_cf
        if checa_gap(gap):
            achados.append(achado_ok("X", "X1b",
                "Caixa Final: BF = DFC",
                f"BF {bf_cf:,.2f}  =  DFC {dfc_cf:,.2f}", valor=bf_cf))
        else:
            achados.append(achado_erro("X", "X1b",
                "Caixa Final: BF ≠ DFC",
                f"Diferença: {gap:,.2f}  |  BF {bf_cf:,.2f}  DFC {dfc_cf:,.2f}",
                valor=gap))
    else:
        skip("X1b", "Caixa Final BF=DFC", "BF ou DFC")

    # ── X2. Geração Líquida DFC = Variação do Caixa ─────────────────────────
    if t_dfc:
        gap = _v(t_dfc, "GAP_X2", "X2")
        if checa_gap(gap, Decimal("0.05")):
            achados.append(achado_ok("X", "X2",
                "DFC: Geração Líquida (I+II+III) = Variação do Caixa",
                f"Gap: {gap:,.2f}", valor=gap))
        else:
            achados.append(achado_erro("X", "X2",
                "DFC: Geração Líquida ≠ Variação do Caixa",
                f"Gap: {gap:,.2f}  |  "
                f"Geração {t_dfc.get('GERACAO',D(0)):,.2f}  "
                f"Var.Caixa {t_dfc.get('VAR_CAIXA',D(0)):,.2f}",
                valor=gap))
    else:
        skip("X2", "DFC: I+II+III = Var.Caixa", "DFC")

    # ── X6. Resultado do Exercício: DMPL = DVP ──────────────────────────────
    if t_dmpl and t_dvp:
        res_dmpl = _v(t_dmpl, "RESULTADO_EXERCICIO", "X6")
        res_dvp  = _v(t_dvp, "RESULTADO_PATRIMONIAL", "X6")
        gap = res_dmpl - res_dvp
        if checa_gap(gap):
            achados.append(achado_ok("X", "X6",
                "Resultado do Exercício: DMPL = DVP",
                f"DMPL {res_dmpl:,.2f}  =  DVP {res_dvp:,.2f}", valor=res_dmpl))
        else:
            achados.append(achado_erro("X", "X6",
                "Resultado do Exercício: DMPL ≠ DVP",
                f"Diferença: {gap:,.2f}  |  "
                f"DMPL {res_dmpl:,.2f}  DVP {res_dvp:,.2f}",
                valor=gap))
    else:
        skip("X6", "Resultado do Exercício DMPL=DVP", "DMPL ou DVP")

    # ── Regra 1a. Resultado do Exercício: BP = DVP ───────────────────────────
    if t_bp and t_dvp:
        res_bp  = _v(t_bp, "RESULTADO_EXERCICIO", "R1a")
        res_dvp = _v(t_dvp, "RESULTADO_PATRIMONIAL", "R1a")
        gap = res_bp - res_dvp
        if checa_gap(gap):
            achados.append(achado_ok("X", "R1a",
                "Resultado do Exercício: BP = DVP(III)",
                f"BP {res_bp:,.2f}  =  DVP {res_dvp:,.2f}", valor=res_bp))
        else:
            achados.append(achado_erro("X", "R1a",
                "Resultado do Exercício: BP ≠ DVP(III)",
                f"Diferença: {gap:,.2f}",
                valor=gap))
    else:
        skip("R1a", "Resultado BP=DVP", "BP ou DVP")

    # ── Regra 2. Saldo Final DMPL = PL do BP ────────────────────────────────
    if t_dmpl and t_bp:
        pl_dmpl = _v(t_dmpl, "PL_FINAL", "R2")
        pl_bp   = _v(t_bp, "PATRIMONIO_LIQUIDO", "R2")
        gap = pl_dmpl - pl_bp
        if checa_gap(gap):
            achados.append(achado_ok("X", "R2",
                "Saldo Final DMPL = PL do BP",
                f"DMPL {pl_dmpl:,.2f}  =  BP {pl_bp:,.2f}", valor=pl_dmpl))
        else:
            achados.append(achado_erro("X", "R2",
                "Saldo Final DMPL ≠ PL do BP",
                f"Diferença: {gap:,.2f}",
                valor=gap))
    else:
        skip("R2", "Saldo Final DMPL=PL BP", "DMPL ou BP")

    # ── Regra 4a. Caixa: BP = DFC ───────────────────────────────────────────
    if t_bp and t_dfc:
        caixa_bp  = _v(t_bp, "CAIXA", "R4a")
        caixa_dfc = _v(t_dfc, "CAIXA_FIN", "R4a")
        gap = caixa_bp - caixa_dfc
        if checa_gap(gap):
            achados.append(achado_ok("X", "R4a",
                "Caixa: BP = DFC (Caixa Final)",
                f"BP {caixa_bp:,.2f}  =  DFC {caixa_dfc:,.2f}", valor=caixa_bp))
        else:
            achados.append(achado_erro("X", "R4a",
                "Caixa: BP ≠ DFC",
                f"Diferença: {gap:,.2f}",
                valor=gap))
    else:
        skip("R4a", "Caixa BP=DFC", "BP ou DFC")

    # ── Regra 7. Receita Orçamentária: BF = BO ──────────────────────────────
    if t_bf and t_bo:
        rec_bf = _v(t_bf, "RECEITA_REALIZADA", "R7")
        rec_bo = _v(t_bo, "RECEITA_REALIZADA", "R7")
        gap = rec_bf - rec_bo
        if checa_gap(gap):
            achados.append(achado_ok("X", "R7",
                "Receita Realizada: BF = BO",
                f"BF {rec_bf:,.2f}  =  BO {rec_bo:,.2f}", valor=rec_bf))
        else:
            achados.append(achado_erro("X", "R7",
                "Receita Realizada: BF ≠ BO",
                f"Diferença: {gap:,.2f}",
                valor=gap))
    else:
        skip("R7", "Receita BF=BO", "BF ou BO")

    # ── Regra 8. Despesa Paga: BF = BO ──────────────────────────────────────
    if t_bf and t_bo:
        pag_bf = _v(t_bf, "DESPESA_PAGA", "R8")
        pag_bo = _v(t_bo, "DESPESA_PAGA", "R8")
        gap = pag_bf - pag_bo
        if checa_gap(gap):
            achados.append(achado_ok("X", "R8",
                "Despesa Paga: BF = BO",
                f"BF {pag_bf:,.2f}  =  BO {pag_bo:,.2f}", valor=pag_bf))
        else:
            achados.append(achado_erro("X", "R8",
                "Despesa Paga: BF ≠ BO",
                f"Diferença: {gap:,.2f}",
                valor=gap))
    else:
        skip("R8", "Despesa Paga BF=BO", "BF ou BO")

    # ── Regra 10a. Pgtos RPNP: BF = BO ─────────────────────────────────────
    if t_bf and t_bo:
        rpnp_bf = _v(t_bf, "PAG_RPNP", "R10a")
        rpnp_bo = _v(t_bo, "PAG_RPNP", "R10a")
        gap = rpnp_bf - rpnp_bo
        if checa_gap(gap):
            achados.append(achado_ok("X", "R10a",
                "Pgtos RPNP: BF = BO",
                f"BF {rpnp_bf:,.2f}  =  BO {rpnp_bo:,.2f}", valor=rpnp_bf))
        else:
            achados.append(achado_erro("X", "R10a",
                "Pgtos RPNP: BF ≠ BO",
                f"Diferença: {gap:,.2f}",
                valor=gap))
    else:
        skip("R10a", "Pgtos RPNP BF=BO", "BF ou BO")

    # ── Regra 10b. Pgtos RPP: BF = BO ───────────────────────────────────────
    if t_bf and t_bo:
        rpp_bf = _v(t_bf, "PAG_RPP", "R10b")
        rpp_bo = _v(t_bo, "PAG_RPP", "R10b")
        gap = rpp_bf - rpp_bo
        if checa_gap(gap):
            achados.append(achado_ok("X", "R10b",
                "Pgtos RPP: BF = BO",
                f"BF {rpp_bf:,.2f}  =  BO {rpp_bo:,.2f}", valor=rpp_bf))
        else:
            achados.append(achado_erro("X", "R10b",
                "Pgtos RPP: BF ≠ BO",
                f"Diferença: {gap:,.2f}",
                valor=gap))
    else:
        skip("R10b", "Pgtos RPP BF=BO", "BF ou BO")

    return achados
