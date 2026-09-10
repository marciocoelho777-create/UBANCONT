# -*- coding: utf-8 -*-
"""
painel.py — monta o README.md do repositório a partir do histórico em
painel/dados/*.json.

O GitHub renderiza o README na página inicial do repositório, e num
repositório privado só colaboradores convidados enxergam. Não usa GitHub
Pages de propósito: no plano gratuito, o Pages de um repositório privado
publica o site na web aberta.

Uso:  python painel.py
"""
from __future__ import annotations
import glob
import json
import os
from datetime import datetime

DIR_DADOS = os.path.join("painel", "dados")
SAIDA = "README.md"
N_HISTORICO = 14          # execuções mostradas na trilha de status
ICONE = {"OK": "✅", "ALERTA": "⚠️", "ERRO": "❌", "INFO": "ℹ️"}

# Rótulos dos indicadores. As chaves do JSON são ASCII de propósito (nomes de
# campo), mas o painel é lido por gente.
ROTULOS = {
    "caixa_final":         "Caixa Final",
    "receita_realizada":   "Receita Realizada",
    "despesa_paga":        "Despesa Orçamentária Paga",
    "ingressos":           "Ingressos",
    "dispendios":          "Dispêndios",
    "despesa_empenhada":   "Despesa Empenhada",
    "despesa_liquidada":   "Despesa Liquidada",
    "ativo_total":         "Ativo Total",
    "patrimonio_liquido":  "Patrimônio Líquido",
    "resultado_exercicio": "Resultado do Exercício",
}


def _plural(n: int, singular: str, plural: str) -> str:
    return f"{n} {singular if n == 1 else plural}"


def carregar() -> list[dict]:
    docs = []
    for caminho in glob.glob(os.path.join(DIR_DADOS, "*.json")):
        try:
            with open(caminho, encoding="utf-8") as f:
                d = json.load(f)
            d["_arquivo"] = os.path.basename(caminho)
            docs.append(d)
        except (json.JSONDecodeError, OSError) as e:
            print(f"  [aviso] ignorando {caminho}: {e}")
    docs.sort(key=lambda d: d.get("gerado_em", ""))
    return docs


def _brl(v) -> str:
    if v is None:
        return "—"
    s = f"{abs(v):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return ("−" if v < 0 else "") + "R$ " + s


def _dt(iso: str) -> str:
    try:
        return datetime.fromisoformat(iso).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return iso


def montar(docs: list[dict]) -> str:
    if not docs:
        return "# Diagnóstico Contábil GDF\n\nSem execuções registradas ainda.\n"

    ult = docs[-1]
    t = ult["totais"]
    L = []

    # ── Cabeçalho ──────────────────────────────────────────────────────────
    if t.get("ERRO", 0):
        estado = f"❌ **{t['ERRO']} erro(s)**"
    elif t.get("ALERTA", 0):
        estado = f"⚠️ **{t['ALERTA']} alerta(s)**"
    else:
        estado = "✅ **Todos os controles passaram**"

    L.append("# Diagnóstico Contábil GDF")
    L.append("")
    L.append(f"{estado} — competência **{ult['mes']:02d}/{ult['ano']}** · "
             f"escopo **{ult.get('escopo','?')}** · "
             f"execução de {_dt(ult['gerado_em'])}")
    L.append("")
    L.append(f"`{t.get('ERRO',0)} erro` · `{t.get('ALERTA',0)} alerta` · "
             f"`{t.get('OK',0)} ok` · `{t.get('INFO',0)} info` · "
             f"**{t.get('TOTAL',0)} controles**")
    L.append("")

    # ── Indicadores ────────────────────────────────────────────────────────
    ind = ult.get("indicadores") or {}
    if ind:
        L.append("## Indicadores")
        L.append("")
        L.append("| Indicador | Valor |")
        L.append("|---|---:|")
        for k, v in ind.items():
            L.append(f"| {ROTULOS.get(k, k.replace('_',' ').title())} | {_brl(v)} |")
        L.append("")

    # ── Situação atual ─────────────────────────────────────────────────────
    # A coluna Diferença só aparece se algum controle trouxer valor. Uma
    # coluna inteira de trações parece falha de preenchimento — e hoje é: os
    # achado_ok não propagam `valor`, então num dia sem erro ela fica vazia.
    tem_valor = any(a["valor"] is not None for a in ult["achados"])

    L.append("## Controles — última execução")
    L.append("")
    if tem_valor:
        L.append("| | Módulo | Código | Controle | Diferença |")
        L.append("|:-:|---|---|---|---:|")
    else:
        L.append("| | Módulo | Código | Controle |")
        L.append("|:-:|---|---|---|")
    for a in ult["achados"]:
        linha = (f"| {ICONE.get(a['status'],'·')} | {a['modulo']} | "
                 f"`{a['codigo']}` | {a['titulo']} |")
        if tem_valor:
            linha += f" {_brl(a['valor']) if a['valor'] is not None else '—'} |"
        L.append(linha)
    L.append("")

    # ── Detalhe do que não passou ──────────────────────────────────────────
    problemas = [a for a in ult["achados"] if a["status"] in ("ERRO", "ALERTA")]
    if problemas:
        L.append("## O que não passou")
        L.append("")
        for a in problemas:
            L.append(f"**{ICONE[a['status']]} `{a['codigo']}` {a['titulo']}**  ")
            L.append(f"{a['detalhe']}")
            L.append("")

    # ── Trilha de status ───────────────────────────────────────────────────
    recentes = docs[-N_HISTORICO:]
    codigos = []
    for d in recentes:
        for a in d["achados"]:
            if a["codigo"] not in codigos:
                codigos.append(a["codigo"])

    L.append(f"## Histórico ({_plural(len(recentes), 'execução', 'execuções')})")
    L.append("")
    L.append("Uma coluna por execução, da mais antiga à mais recente. "
             "Célula vazia = o controle não existia naquela execução.")
    L.append("")
    L.append("| Controle | " + " | ".join(
        datetime.fromisoformat(d["gerado_em"]).strftime("%d/%m")
        if "gerado_em" in d else "?" for d in recentes) + " |")
    L.append("|---" * (len(recentes) + 1) + "|")
    for cod in codigos:
        celulas = []
        for d in recentes:
            achado = next((a for a in d["achados"] if a["codigo"] == cod), None)
            celulas.append(ICONE.get(achado["status"], "·") if achado else "")
        L.append(f"| `{cod}` | " + " | ".join(celulas) + " |")
    L.append("")

    # ── Rodapé ─────────────────────────────────────────────────────────────
    L.append("---")
    L.append("")
    L.append(f"Gerado por `painel.py` em {datetime.now().strftime('%d/%m/%Y %H:%M')} · "
             f"{_plural(len(docs), 'execução', 'execuções')} no histórico · "
             f"dados em `painel/dados/`.")
    L.append("")
    L.append("> Repositório privado. Não habilitar GitHub Pages: no plano "
             "gratuito, o Pages de um repositório privado publica o conteúdo "
             "na web aberta.")
    L.append("")
    return "\n".join(L)


def main():
    docs = carregar()
    texto = montar(docs)
    with open(SAIDA, "w", encoding="utf-8") as f:
        f.write(texto)
    print(f"  {SAIDA} atualizado — {len(docs)} execução(ões) no histórico.")


if __name__ == "__main__":
    main()
