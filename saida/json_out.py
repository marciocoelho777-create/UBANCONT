# -*- coding: utf-8 -*-
"""
saida/json_out.py — grava o resultado do diagnóstico em JSON.

Um arquivo por execução, em painel/dados/. O painel.py lê todos e monta o
histórico. JSON e não Excel porque o Excel é saída de apresentação: extrair
dado dele de volta é frágil e quebra a cada mudança de layout.

Uso no diag.py:

    from saida import json_out
    ...
    if args.json:
        json_out.gravar(achados, mes=args.mes, ano=args.ano,
                        escopo=escopo, totais=totais, extras=extras)
"""
from __future__ import annotations
import json
import os
from datetime import datetime
from decimal import Decimal

# Diretório do painel, relativo à raiz do projeto.
DIR_DADOS = os.path.join("painel", "dados")


def _campo(obj, *nomes, padrao=None):
    """Lê o primeiro atributo existente. O Achado pode ter nomes diferentes
    conforme a versão do controles/__init__.py; isto evita acoplamento."""
    for n in nomes:
        if isinstance(obj, dict) and n in obj:
            return obj[n]
        if hasattr(obj, n):
            return getattr(obj, n)
    return padrao


def _num(v):
    """Decimal -> float para o JSON. None permanece None."""
    if v is None:
        return None
    if isinstance(v, Decimal):
        return float(v)
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def serializar(achados, mes: int, ano: int, escopo: str = "Consolidado",
               extras: dict | None = None) -> dict:
    itens = []
    for a in achados:
        itens.append({
            "status":  str(_campo(a, "status", "nivel", "tipo", padrao="?")).upper(),
            "modulo":  _campo(a, "modulo", "grupo", padrao=""),
            "codigo":  _campo(a, "codigo", "cod", padrao=""),
            "titulo":  _campo(a, "titulo", "nome", padrao=""),
            "detalhe": _campo(a, "detalhe", "mensagem", "msg", padrao=""),
            "valor":   _num(_campo(a, "valor")),
        })

    totais = {"ERRO": 0, "ALERTA": 0, "OK": 0, "INFO": 0}
    for i in itens:
        totais[i["status"]] = totais.get(i["status"], 0) + 1
    # O total é a soma de TODOS os status, inclusive INFO — o rodapé do Excel
    # contava 29 controles e quebrava só em 2 ALERTA + 25 OK, que não fechava.
    totais["TOTAL"] = len(itens)

    return {
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
        "mes": mes,
        "ano": ano,
        "escopo": escopo,
        "totais": totais,
        "indicadores": {k: _num(v) for k, v in (extras or {}).items()},
        "achados": itens,
    }


def gravar(achados, mes: int, ano: int, escopo: str = "Consolidado",
           extras: dict | None = None, dir_dados: str = DIR_DADOS) -> str:
    doc = serializar(achados, mes, ano, escopo, extras)
    os.makedirs(dir_dados, exist_ok=True)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome = f"{ano}-{mes:02d}_{carimbo}.json"
    caminho = os.path.join(dir_dados, nome)
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    return caminho
