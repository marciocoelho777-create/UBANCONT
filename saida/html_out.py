# -*- coding: utf-8 -*-
"""
saida/html_out.py — atualiza os painéis HTML com dados do run atual.

Substitui o bloco entre os marcadores
  /* AUTO-DATA-START */
  ...
  /* AUTO-DATA-END */
preservando todo o restante do arquivo (CSS, estrutura, JS de renderização).
"""
from __future__ import annotations
import json
import re
from pathlib import Path

RAIZ  = Path(__file__).parent.parent
PAINEL = RAIZ / "painel"

_PAT = re.compile(r'/\* AUTO-DATA-START \*/.*?/\* AUTO-DATA-END \*/', re.DOTALL)


def _update(html_path: Path, js_block: str) -> None:
    html = html_path.read_text(encoding="utf-8")
    novo = f"/* AUTO-DATA-START */\n{js_block}\n/* AUTO-DATA-END */"
    html2 = _PAT.sub(novo, html)
    if html2 == html:
        print(f"  [html] marcadores AUTO-DATA não encontrados em {html_path.name}")
        return
    html_path.write_text(html2, encoding="utf-8")
    print(f"  HTML:  {html_path}")


def _ler_historico(pasta: Path, n: int = 6, tipo: str = "diag") -> list[dict]:
    """Últimas n execuções de uma pasta de JSONs (mais recente primeiro)."""
    if not pasta.exists():
        return []
    arquivos = sorted(pasta.glob("*.json"), reverse=True)[:n]
    hist: list[dict] = []
    for arq in arquivos:
        try:
            d = json.loads(arq.read_text(encoding="utf-8"))
            if tipo == "classif":
                hist.append({
                    "gerado_em": d.get("gerado_em", ""),
                    "n": len(d.get("achados", [])),
                    "por_dem": {
                        k: v.get("valor", 0)
                        for k, v in d.get("por_demonstrativo", {}).items()
                    },
                })
            else:
                erros = [
                    a["codigo"]
                    for a in d.get("achados", [])
                    if a.get("status") == "ERRO"
                ]
                hist.append({
                    "gerado_em": d.get("gerado_em", ""),
                    "totais": d.get("totais", {}),
                    "erros": erros,
                })
        except Exception:
            pass
    return hist


def _ler_historico_full(pasta: Path, n: int = 24) -> list[dict]:
    """Últimas n execuções completas (achados + indicadores) para o seletor de datas."""
    if not pasta.exists():
        return []
    arquivos = sorted(pasta.glob("*.json"), reverse=True)[:n]
    result: list[dict] = []
    for arq in arquivos:
        try:
            d = json.loads(arq.read_text(encoding="utf-8"))
            result.append({
                "gerado_em":   d.get("gerado_em", ""),
                "mes":         d.get("mes", 0),
                "ano":         d.get("ano", 0),
                "escopo":      d.get("escopo", "Consolidado"),
                "totais":      d.get("totais", {}),
                "indicadores": {
                    k: (float(v) if v is not None else None)
                    for k, v in (d.get("indicadores") or {}).items()
                },
                "achados": [
                    {"s": a.get("status", ""), "m": a.get("modulo", ""),
                     "c": a.get("codigo", ""),  "t": a.get("titulo", ""),
                     "d": a.get("detalhe", "")}
                    for a in d.get("achados", [])
                ],
            })
        except Exception:
            pass
    return result


def gerar_diag(doc: dict, destino: Path | None = None) -> None:
    """Atualiza painel_diag.html com os dados do run atual."""
    destino = destino or (PAINEL / "painel_diag.html")
    if not destino.exists():
        print(f"  [html] {destino.name} não encontrado — painel não atualizado")
        return

    meta = {
        "gerado_em":  doc["gerado_em"],
        "mes":        doc["mes"],
        "ano":        doc["ano"],
        "escopo":     doc.get("escopo", "Consolidado"),
        "totais":     doc["totais"],
        "indicadores": {
            k: (float(v) if v is not None else None)
            for k, v in (doc.get("indicadores") or {}).items()
        },
        "historico":  _ler_historico(PAINEL / "dados", n=6),
    }
    achados = [
        {
            "s": a["status"],
            "m": a["modulo"],
            "c": a["codigo"],
            "t": a["titulo"],
            "d": a.get("detalhe", ""),
        }
        for a in doc.get("achados", [])
    ]
    hist_full = _ler_historico_full(PAINEL / "dados", n=24)
    js = (
        f"const META = {json.dumps(meta, ensure_ascii=False, indent=2)};\n"
        f"const ACHADOS = {json.dumps(achados, ensure_ascii=False, indent=2)};\n"
        f"var HISTORICO_FULL = {json.dumps(hist_full, ensure_ascii=False, indent=2)};"
    )
    _update(destino, js)


def gerar_classificacao(doc: dict, destino: Path | None = None) -> None:
    """Atualiza painel_classificacao.html com os dados do run atual."""
    destino = destino or (PAINEL / "painel_classificacao.html")
    if not destino.exists():
        print(f"  [html] {destino.name} não encontrado — painel não atualizado")
        return

    cdata = {
        "gerado_em":         doc.get("gerado_em", ""),
        "mes":               doc.get("mes"),
        "ano":               doc.get("ano"),
        "por_demonstrativo": doc.get("por_demonstrativo", {}),
        "achados":           doc.get("achados", []),
        "historico":         _ler_historico(PAINEL / "dados" / "classificacao",
                                            n=6, tipo="classif"),
    }
    js = f"const CDATA = {json.dumps(cdata, ensure_ascii=False, indent=2)};"
    _update(destino, js)
