# -*- coding: utf-8 -*-
"""
diag_contabil.controles
=======================
Tipos e utilitários compartilhados entre os módulos de controle.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Optional
import oracledb


# ─────────────────────────────────────────────────────────────────────────────
#  Tipo central: Achado
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class Achado:
    """Resultado de um controle individual."""
    status:  str        # "OK" | "ERRO" | "ALERTA" | "INFO"
    modulo:  str        # "BF" | "BO" | "DFC" | "DVP" | "DMPL" | "BP" | "X"
    codigo:  str        # ex: "BF-01", "X2", "DFC-03"
    titulo:  str        # descrição curta (≤ 60 chars)
    detalhe: str = ""   # valores, diferenças, contexto
    valor:   Optional[Decimal] = None
    # `valor` é a GRANDEZA SOB TESTE, não necessariamente a diferença:
    #   - em ERRO/ALERTA, é o gap apurado (o que falta fechar);
    #   - em OK/INFO, é o montante que o controle verificou.
    # Num controle OK a diferença é zero por definição, então gravar o gap
    # daria uma coluna de zeros. Gravando a grandeza, a planilha fica
    # ordenável por materialidade mesmo num dia sem nenhum erro.

    @property
    def ok(self) -> bool:
        return self.status == "OK"

    @property
    def critico(self) -> bool:
        return self.status == "ERRO"


# ─────────────────────────────────────────────────────────────────────────────
#  Helper de query
# ─────────────────────────────────────────────────────────────────────────────
def query_one(conn: oracledb.Connection, sql: str,
              params: Optional[dict] = None) -> dict:
    """Roda um SELECT que retorna exatamente uma linha e devolve como dict."""
    cur = conn.cursor()
    cur.execute(sql, params or {})
    cols = [d[0].upper() for d in cur.description]
    row  = cur.fetchone()
    cur.close()
    if row is None:
        return {c: Decimal(0) for c in cols}
    return {c: (Decimal(str(v)) if isinstance(v, (int, float)) else v)
            for c, v in zip(cols, row)}


def query_all(conn: oracledb.Connection, sql: str,
              params: Optional[dict] = None) -> list[dict]:
    """Roda um SELECT que pode retornar várias linhas."""
    cur = conn.cursor()
    cur.execute(sql, params or {})
    cols = [d[0].upper() for d in cur.description]
    rows = cur.fetchall()
    cur.close()
    return [{c: (Decimal(str(v)) if isinstance(v, (int, float)) else v)
             for c, v in zip(cols, row)} for row in rows]


def D(v) -> Decimal:
    """Converte para Decimal, tratando None como zero."""
    if v is None:
        return Decimal(0)
    return Decimal(str(v))


def achado_ok(modulo, codigo, titulo, detalhe="", valor=None) -> Achado:
    return Achado("OK", modulo, codigo, titulo, detalhe, valor)


def achado_erro(modulo, codigo, titulo, detalhe="", valor=None) -> Achado:
    return Achado("ERRO", modulo, codigo, titulo, detalhe, valor)


def achado_alerta(modulo, codigo, titulo, detalhe="", valor=None) -> Achado:
    return Achado("ALERTA", modulo, codigo, titulo, detalhe, valor)


def achado_info(modulo, codigo, titulo, detalhe="", valor=None) -> Achado:
    return Achado("INFO", modulo, codigo, titulo, detalhe, valor)


def checa_gap(valor: Decimal, tolerancia: Decimal = Decimal("0.02")) -> bool:
    """Retorna True se o valor estiver dentro da tolerância (fecha)."""
    return abs(valor) <= tolerancia
