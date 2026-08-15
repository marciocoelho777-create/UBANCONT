# -*- coding: utf-8 -*-
"""Formatação dos achados para o terminal."""
from __future__ import annotations
from datetime import datetime
from controles import Achado

# Cores ANSI (desativadas automaticamente se não for TTY)
import sys
_COLOR = sys.stdout.isatty()
def _c(code, txt): return f"\033[{code}m{txt}\033[0m" if _COLOR else txt
VERDE   = lambda t: _c("32", t)
AMARELO = lambda t: _c("33", t)
VERMELHO= lambda t: _c("31", t)
CIANO   = lambda t: _c("36", t)
NEGRITO = lambda t: _c("1",  t)

ICONES = {
    "OK":     "✔",
    "ERRO":   "✘",
    "ALERTA": "⚠",
    "INFO":   "ℹ",
}
CORES_STATUS = {
    "OK":     VERDE,
    "ERRO":   VERMELHO,
    "ALERTA": AMARELO,
    "INFO":   CIANO,
}

W = 70  # largura da caixa


def contar(achados: list[Achado]) -> dict:
    """Contagem única, usada pelo console, pelo Excel e pelo JSON.
    Inclui INFO: um total que não soma com as parcelas exibidas é a primeira
    coisa que alguém aponta num relatório de auditoria."""
    c = {"ERRO": 0, "ALERTA": 0, "OK": 0, "INFO": 0}
    for a in achados:
        c[a.status] = c.get(a.status, 0) + 1
    c["TOTAL"] = len(achados)
    return c


def cabecalho(mes: int, ano: int, coug: str = "Consolidado"):
    nomes = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",
             6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",
             11:"Novembro",12:"Dezembro"}
    print()
    print(NEGRITO("═" * W))
    print(NEGRITO(f"  DIAGNÓSTICO CONTÁBIL GDF — {nomes.get(mes,mes)}/{ano}  |  {coug}"))
    print(NEGRITO(f"  {datetime.now():%d/%m/%Y %H:%M:%S}"))
    print(NEGRITO("═" * W))


def secao(modulo: str, titulo: str):
    print()
    print(NEGRITO(f"{'─'*W}"))
    print(NEGRITO(f"  {modulo} — {titulo}"))
    print(NEGRITO(f"{'─'*W}"))


def linha_achado(a: Achado):
    icone = ICONES.get(a.status, "?")
    cor   = CORES_STATUS.get(a.status, lambda x: x)
    cod   = f"[{a.codigo}]"
    print(f"  {cor(icone)} {cor(cod):<10} {a.titulo}")
    if a.detalhe:
        print(f"             {a.detalhe}")


def resumo(achados: list[Achado]):
    n = contar(achados)
    print()
    print("─" * W)
    if n["ERRO"] == 0 and n["ALERTA"] == 0:
        txt = f"  RESULTADO: TODOS OS {n['OK']} CONTROLES PASSARAM."
        if n["INFO"]:
            txt += f"  ({n['INFO']} informativo(s) · {n['TOTAL']} no total)"
        print(VERDE(txt))
    else:
        partes = []
        if n["ERRO"]:   partes.append(VERMELHO(f"{n['ERRO']} erro(s)"))
        if n["ALERTA"]: partes.append(AMARELO(f"{n['ALERTA']} alerta(s)"))
        if n["OK"]:     partes.append(VERDE(f"{n['OK']} OK"))
        if n["INFO"]:   partes.append(CIANO(f"{n['INFO']} info"))
        print(f"  RESULTADO: {' / '.join(partes)}  ({n['TOTAL']} controles)")
    print("─" * W)


def imprimir_modulo(modulo: str, titulo: str, achados: list[Achado]):
    secao(modulo, titulo)
    for a in achados:
        linha_achado(a)
    n = contar(achados)
    if n["ERRO"] == 0 and n["ALERTA"] == 0:
        print(VERDE("  RESULTADO: OK."))
    else:
        partes = []
        if n["ERRO"]:   partes.append(VERMELHO(f"{n['ERRO']} erro(s)"))
        if n["ALERTA"]: partes.append(AMARELO(f"{n['ALERTA']} alerta(s)"))
        print(f"  RESULTADO: {' / '.join(partes)}")


def rodape(achados: list[Achado]):
    print()
    print(NEGRITO("═" * W))
    resumo(achados)
    print(NEGRITO("═" * W))
    print()
