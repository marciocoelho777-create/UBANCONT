#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
  MONITOR VS PDF — Compara totais do banco Oracle com os PDFs oficiais
  (13 - DEMONSTRATIVOS CONTÁBEIS) e reporta APENAS divergências.

  Uso:
      python monitor_vs_pdf.py --mes 8 --ano 2026
      python monitor_vs_pdf.py --mes 9 --ano 2026

  Tolerância: R$ 1,00 (arredondamento normal de centavos).
================================================================================
"""
from __future__ import annotations
import argparse, json, sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import oracledb

RAIZ      = Path(__file__).parent
DIR_DADOS = RAIZ / "painel" / "dados" / "vs_pdf"
TOLERANCIA = Decimal("1.00")

# ─────────────────────────────────────────────────────────────────────────────
#  Credenciais
# ─────────────────────────────────────────────────────────────────────────────
DB_USER = DB_PASSWORD = DB_HOST = DB_SERVICE = ""
DB_PORT = "1521"
INSTANT_CLIENT_DIR = ""

try:
    import importlib.util as _ilu
    _cfg = RAIZ / "config_local.py"
    if _cfg.exists():
        _spec = _ilu.spec_from_file_location("config_local", _cfg)
        _mod  = _ilu.module_from_spec(_spec); _spec.loader.exec_module(_mod)
        for _k in ("DB_USER","DB_PASSWORD","DB_HOST","DB_PORT","DB_SERVICE",
                   "INSTANT_CLIENT_DIR"):
            if hasattr(_mod, _k): globals()[_k] = getattr(_mod, _k)
        print("  [config] credenciais carregadas de config_local.py")
except Exception:
    pass


def _conectar():
    if INSTANT_CLIENT_DIR:
        try:
            oracledb.init_oracle_client(lib_dir=str(INSTANT_CLIENT_DIR))
        except Exception as e:
            if "already been initialized" not in str(e): raise
    conn = oracledb.connect(user=DB_USER, password=DB_PASSWORD,
                            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}")
    print(f"  Conectado! Oracle {conn.version}")
    return conn


# ─────────────────────────────────────────────────────────────────────────────
#  Utilitários de formatação
# ─────────────────────────────────────────────────────────────────────────────
def _brl(v) -> str:
    if v is None: return "—"
    return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _f(v) -> float:
    if v is None: return 0.0
    return float(v)


def _cmp(demo, rotulo, db_val, pdf_val, pdf_path: Path | None = None):
    """Compara dois valores; devolve dict de divergência ou None se OK."""
    dif = abs(_f(db_val) - _f(pdf_val))
    if dif <= float(TOLERANCIA):
        return None
    return {
        "demo":        demo,
        "rotulo":      rotulo,
        "valor_banco": _f(db_val),
        "valor_pdf":   _f(pdf_val),
        "diferenca":   _f(db_val) - _f(pdf_val),
        "pdf_arquivo": str(pdf_path) if pdf_path else None,
        "pdf_data":    (datetime.fromtimestamp(pdf_path.stat().st_mtime)
                        .strftime("%d/%m/%Y %H:%M") if pdf_path else None),
    }


def _err(demo, rotulo, motivo, pdf_path: Path | None = None):
    return {
        "demo":        demo,
        "rotulo":      rotulo,
        "valor_banco": None,
        "valor_pdf":   None,
        "diferenca":   None,
        "erro":        motivo,
        "pdf_arquivo": str(pdf_path) if pdf_path else None,
        "pdf_data":    (datetime.fromtimestamp(pdf_path.stat().st_mtime)
                        .strftime("%d/%m/%Y %H:%M") if pdf_path else None),
    }


# ─────────────────────────────────────────────────────────────────────────────
#  Extração de totais do banco
# ─────────────────────────────────────────────────────────────────────────────
def _totais_banco(conn, mes, ano) -> dict:
    """Roda os módulos BO, BF, BP, DVP e devolve dict com todos os totais."""
    from controles import bo as _bo, bf as _bf, bp as _bp, dvp as _dvp
    totais = {}
    for nome, fn in [("bo", _bo.auditar), ("bf", _bf.auditar),
                     ("bp", _bp.auditar), ("dvp", _dvp.auditar)]:
        try:
            _, t = fn(conn, mes, ano)
            totais[nome] = {k: _f(v) for k, v in t.items()
                            if not isinstance(v, (list, dict))}
        except Exception as e:
            print(f"  [aviso] módulo {nome.upper()} falhou: {e}")
            totais[nome] = {}
    return totais


# ─────────────────────────────────────────────────────────────────────────────
#  Extração de totais dos PDFs
# ─────────────────────────────────────────────────────────────────────────────
def _comparar(conn, mes, ano) -> list[dict]:
    from leitor_pdf_oficial import carregar, PdfNaoEncontrado, _localizar_pdf
    from leitor_pdf_oficial import valores as pdf_valores, pos as pdf_pos

    print("  Extraindo totais do banco Oracle…")
    t = _totais_banco(conn, mes, ano)
    bo, bf, bp, dvp = t.get("bo",{}), t.get("bf",{}), t.get("bp",{}), t.get("dvp",{})

    print("  Lendo PDFs oficiais…")
    try:
        demos = carregar(mes, ano)
    except PdfNaoEncontrado as e:
        print(f"  [erro] {e}")
        return [_err("?", "Carregamento dos PDFs", str(e))]

    achados = []

    # ── Caminho dos PDFs para exibir no painel ──────────────────────────────
    def _pdf_path(tipo):
        try:
            return _localizar_pdf(tipo, mes, ano)
        except Exception:
            return None

    # ── BO — Balanço Orçamentário ───────────────────────────────────────────
    bo_path = _pdf_path("BO")
    try:
        txt_bo = demos.t("BO")
        i_rec = pdf_pos(txt_bo, "SUBTOTAL DAS RECEITAS (III) = (I + II)")
        vals_rec, _ = pdf_valores(txt_bo, "SUBTOTAL DAS RECEITAS (III) = (I + II)", 4,
                                   inicio=i_rec)
        rec_pdf = vals_rec[2]
        d = _cmp("BO", "Receita Realizada",
                 bo.get("RECEITA_REALIZADA"), rec_pdf, bo_path)
        if d: achados.append(d)
    except Exception as e:
        achados.append(_err("BO", "Receita Realizada", str(e), bo_path))

    try:
        txt_bo = demos.t("BO")
        i_desp = pdf_pos(txt_bo, "SUBTOTAL DAS DESPESAS")
        vals_desp, _ = pdf_valores(txt_bo, "SUBTOTAL DAS DESPESAS", 6, inicio=i_desp)
        for campo, col, rotulo in [
            ("DESP_EMPENHADA", 2, "Despesa Empenhada"),
            ("DESP_LIQUIDADA", 3, "Despesa Liquidada"),
            ("DESPESA_PAGA",   4, "Despesa Paga"),
        ]:
            d = _cmp("BO", rotulo, bo.get(campo), vals_desp[col], bo_path)
            if d: achados.append(d)
    except Exception as e:
        achados.append(_err("BO", "Despesas (Empenhada/Liquidada/Paga)", str(e), bo_path))

    # ── BF — Balanço Financeiro ─────────────────────────────────────────────
    bf_path = _pdf_path("BF")
    try:
        txt_bf = demos.t("BF")
        vals_rec_bf, _ = pdf_valores(txt_bf, "RECEITA ORÇAMENTÁRIA", 2)
        rec_orcam_pdf = vals_rec_bf[0]
        d = _cmp("BF", "Receita Orçamentária",
                 bo.get("RECEITA_REALIZADA"), rec_orcam_pdf, bf_path)
        if d: achados.append(d)
    except Exception as e:
        achados.append(_err("BF", "Receita Orçamentária", str(e), bf_path))

    try:
        txt_bf = demos.t("BF")
        # Mesmo fallback de regra_08_despesa_bf_bo em motor_regras.py: o "D"
        # inicial cai na linha de corte esq/dirt no PDF novo por mês fechado
        # (10/09/2026 em diante) e some da extração.
        try:
            vals_desp_bf, _ = pdf_valores(txt_bf, "DESPESA ORÇAMENTÁRIA", 2)
        except ValueError:
            vals_desp_bf, _ = pdf_valores(txt_bf, "ESPESA ORÇAMENTÁRIA", 2)
        desp_orcam_pdf = vals_desp_bf[0]
        d = _cmp("BF", "Despesa Orçamentária (Paga)",
                 bo.get("DESPESA_PAGA"), desp_orcam_pdf, bf_path)
        if d: achados.append(d)
    except Exception as e:
        achados.append(_err("BF", "Despesa Orçamentária", str(e), bf_path))

    try:
        txt_bf = demos.t("BF")
        # Mesmo fallback de regra_04_caixa_bp_dfc_bf em motor_regras.py: no
        # PDF novo por mês fechado (10/09/2026 em diante) "SALDO PARA O
        # EXERCÍCIO SEGUINTE" cai na linha de corte esq/dirt (usa só
        # "EXERCÍCIO SEGUINTE"), e os sufixos "RPPS)"/"RPPS"/"Vinculados"
        # das 3 linhas seguintes somem da extração (fallback truncado).
        i_seg = pdf_pos(txt_bf, "EXERCÍCIO SEGUINTE")
        try:
            vals_caixa, _ = pdf_valores(txt_bf,
                                         "Caixa e Equivalentes de Caixa (Exceto RPPS)", 1,
                                         inicio=i_seg)
        except ValueError:
            vals_caixa, _ = pdf_valores(txt_bf,
                                         "Caixa e Equivalentes de Caixa (Exceto", 1,
                                         inicio=i_seg)
        try:
            vals_rpps, _ = pdf_valores(txt_bf,
                                        "CAIXA E EQUIVALENTES DE CAIXA RPPS", 1,
                                        inicio=i_seg)
        except ValueError:
            vals_rpps, _ = pdf_valores(txt_bf,
                                        "CAIXA E EQUIVALENTES DE CAIXA", 1,
                                        inicio=i_seg)
        try:
            vals_dep, _ = pdf_valores(txt_bf,
                                       "Depósitos Restituíveis e Valores Vinculados", 1,
                                       inicio=i_seg)
        except ValueError:
            vals_dep, _ = pdf_valores(txt_bf,
                                       "Depósitos Restituíveis e Valores", 1,
                                       inicio=i_seg)
        caixa_pdf_tot = vals_caixa[0] + vals_rpps[0] + vals_dep[0]
        d = _cmp("BF", "Caixa Final (Exceto RPPS + RPPS + Depósitos Restituíveis)",
                 bf.get("CAIXA_FINAL"), caixa_pdf_tot, bf_path)
        if d: achados.append(d)
    except Exception as e:
        achados.append(_err("BF", "Caixa Final", str(e), bf_path))

    # ── BP — Balanço Patrimonial ────────────────────────────────────────────
    bp_path = _pdf_path("BP")
    try:
        txt_bp = demos.t("BP")
        vals_ativo, _ = pdf_valores(txt_bp, "ATIVO ", 1)
        d = _cmp("BP", "Ativo Total", bp.get("ATIVO_TOTAL"), vals_ativo[0], bp_path)
        if d: achados.append(d)
    except Exception as e:
        achados.append(_err("BP", "Ativo Total", str(e), bp_path))

    try:
        txt_bp = demos.t("BP")
        i_pl = pdf_pos(txt_bp, "PASSIVO NÃO CIRCULANTE")
        vals_pl, _ = pdf_valores(txt_bp, "PATRIMÔNIO LÍQUIDO", 1, inicio=i_pl)
        d = _cmp("BP", "Patrimônio Líquido",
                 bp.get("PATRIMONIO_LIQUIDO"), vals_pl[0], bp_path)
        if d: achados.append(d)
    except Exception as e:
        achados.append(_err("BP", "Patrimônio Líquido", str(e), bp_path))

    # ── DVP — Demonstração das Variações Patrimoniais ───────────────────────
    dvp_path = _pdf_path("DVP")
    try:
        txt_dvp = demos.t("DVP")
        vals_res, _ = pdf_valores(txt_dvp, "RESULTADO PATRIMONIAL DO PERÍODO", 2)
        res_pdf = vals_res[0]
        d = _cmp("DVP", "Resultado Patrimonial do Período",
                 dvp.get("RESULTADO_PATRIMONIAL"), res_pdf, dvp_path)
        if d: achados.append(d)
    except Exception as e:
        achados.append(_err("DVP", "Resultado Patrimonial", str(e), dvp_path))

    # ── DFC — Demonstração dos Fluxos de Caixa ─────────────────────────────
    dfc_path = _pdf_path("DFC")
    try:
        txt_dfc = demos.t("DFC")
        vals_caixa_dfc, _ = pdf_valores(txt_dfc, "Caixa e Equivalente de Caixa Final", 1)
        d = _cmp("DFC", "Caixa e Equivalente de Caixa Final",
                 bf.get("CAIXA_FINAL"), vals_caixa_dfc[0], dfc_path)
        if d: achados.append(d)
    except Exception as e:
        achados.append(_err("DFC", "Caixa Final", str(e), dfc_path))

    try:
        txt_dfc = demos.t("DFC")
        (geracao_pdf,), _ = pdf_valores(txt_dfc,
                                         "GERAÇAO LÍQUIDA DE CAIXA E EQUIVALENTE (I+II+III)", 1)
        # DB: soma I+II+III — extrair do módulo dfc se disponível
        # (dfc não está nos totais_banco; usamos diferença caixa_final − caixa_inicial)
        # Não comparamos aqui sem o módulo DFC — apenas quando dfc rodou
        # Deixa como NAO_VERIF se não tiver os dados
    except Exception:
        pass  # GERAÇÃO LÍQUIDA não comparável sem rodar controles/dfc

    return achados


# ─────────────────────────────────────────────────────────────────────────────
#  Markdown
# ─────────────────────────────────────────────────────────────────────────────
def _markdown(doc: dict) -> str:
    L = [
        f"# Monitor vs PDF — {doc['mes']:02d}/{doc['ano']}",
        f"Gerado em {doc['gerado_em']}",
        "",
    ]
    achados = doc.get("achados", [])
    divs = [a for a in achados if a.get("diferenca") is not None]
    errs = [a for a in achados if a.get("diferenca") is None]

    if not divs and not errs:
        L.append("✅ **Nenhuma divergência entre banco e PDFs oficiais.**")
        return "\n".join(L)

    if divs:
        L += [f"## Divergências ({len(divs)})", ""]
        L += ["| Demo | Campo | Banco | PDF | Diferença | PDF arquivo | Data PDF |",
              "|------|-------|------:|----:|---------:|-------------|----------|"]
        for a in divs:
            L.append(f"| {a['demo']} | {a['rotulo']} | {_brl(a['valor_banco'])} "
                     f"| {_brl(a['valor_pdf'])} | {_brl(a['diferenca'])} "
                     f"| {Path(a['pdf_arquivo']).name if a.get('pdf_arquivo') else '—'} "
                     f"| {a.get('pdf_data','—')} |")

    if errs:
        L += ["", f"## Erros de extração ({len(errs)})", ""]
        for a in errs:
            L.append(f"- **[{a['demo']}] {a['rotulo']}**: {a.get('erro','?')}")

    return "\n".join(L)


# ─────────────────────────────────────────────────────────────────────────────
#  Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser()
    p.add_argument("--mes", type=int, required=True)
    p.add_argument("--ano", type=int, required=True)
    a = p.parse_args()

    print(f"\n== Monitor vs PDF — {a.mes:02d}/{a.ano} ==")
    conn = _conectar()
    try:
        achados = _comparar(conn, a.mes, a.ano)
    finally:
        conn.close(); print("  Conexão Oracle encerrada.")

    divs = [x for x in achados if x.get("diferenca") is not None]
    errs = [x for x in achados if x.get("diferenca") is None]

    if divs:
        print(f"\n  DIVERGÊNCIAS — {len(divs)} campo(s) diferem do PDF oficial:")
        for d in divs:
            print(f"  !! [{d['demo']}] {d['rotulo']}")
            print(f"        Banco: {_brl(d['valor_banco'])}   "
                  f"PDF: {_brl(d['valor_pdf'])}   "
                  f"Diferença: {_brl(d['diferenca'])}")
            if d.get("pdf_data"):
                print(f"        PDF impresso em: {d['pdf_data']}")
    else:
        print("\n  ✅ Banco e PDFs oficiais conferem (sem divergências).")

    if errs:
        print(f"\n  ERROS DE EXTRAÇÃO — {len(errs)} item(ns) não puderam ser verificados:")
        for e in errs:
            print(f"  ?? [{e['demo']}] {e['rotulo']}: {e.get('erro','?')[:80]}")

    gerado_em = datetime.now().isoformat(timespec="seconds")
    doc = {
        "gerado_em": gerado_em,
        "mes":    a.mes,
        "ano":    a.ano,
        "achados": achados,
        "n_diverg": len(divs),
        "n_erros":  len(errs),
    }

    DIR_DADOS.mkdir(parents=True, exist_ok=True)
    ts  = gerado_em.replace("-","").replace(":","").replace("T","_")
    dest = DIR_DADOS / f"{a.ano}-{a.mes:02d}_{ts}.json"
    dest.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n  JSON: {dest}")

    (RAIZ / "MONITOR_VS_PDF.md").write_text(_markdown(doc), encoding="utf-8")

    # Injeta no painel_classificacao.html
    from saida import html_out
    html_out.gerar_vs_pdf(doc, RAIZ / "painel" / "painel_classificacao.html")

    sys.exit(1 if divs else 0)


if __name__ == "__main__":
    main()
