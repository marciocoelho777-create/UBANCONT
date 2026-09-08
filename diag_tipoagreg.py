#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
  DIAGNÓSTICO CONTÁBIL GDF — POR TIPO DE AGREGAÇÃO DE GESTÃO
  Versão separada de diag.py — NÃO modifica nem importa o original.

  Uso:
      python diag_tipoagreg.py --mes 8 --ano 2026 --tipo 1
      python diag_tipoagreg.py --mes 8 --ano 2026 --tipo 3 --excel
      python diag_tipoagreg.py --mes 8 --ano 2026 --tipo 0   (todas as gestões)

  Tipos de agregação (MIL{ano}.TIPOAGREGACAO):
      0  TODAS AS GESTÕES
      1  DIRETA
      2  DIRETA + FUNDOS
      3  AUTARQUIAS
      4  FUNDAÇÃO
      5  EMPRESA PUBLICA
      6  ECONOMIA MISTA
      7  FUNDOS
      8  DIRETA + REPASSE
      9  FUNDOS DA INDIRETA
     10  DIRETA + FUNDOS + REPASSE

  Credenciais: edite config_local.py.
=============================================================================
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import oracledb

# ─────────────────────────────────────────────────────────────────────────────
#  Configuração (sobreposta por config_local.py)
# ─────────────────────────────────────────────────────────────────────────────
DB_USER = ""; DB_PASSWORD = ""; DB_HOST = ""; DB_PORT = "1521"; DB_SERVICE = ""
INSTANT_CLIENT_DIR = ""
OUTPUT_DIR  = Path(__file__).parent / "saidas"
PAINEL_DIR  = Path(__file__).parent / "painel" / "dados" / "tipoagreg"

TIPOS_NOCTURNOS = [1, 3, 4, 5, 6, 7, 9]

try:
    import importlib.util as _ilu, pathlib as _pl
    _cfg = _pl.Path(__file__).parent / "config_local.py"
    if _cfg.exists():
        _spec = _ilu.spec_from_file_location("config_local", _cfg)
        _mod  = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_mod)
        for _k in ("DB_USER","DB_PASSWORD","DB_HOST","DB_PORT","DB_SERVICE","INSTANT_CLIENT_DIR","OUTPUT_DIR"):
            if hasattr(_mod, _k): globals()[_k] = getattr(_mod, _k)
        print("  [config] credenciais carregadas de config_local.py")
except Exception:
    pass


def conectar() -> oracledb.Connection:
    if INSTANT_CLIENT_DIR:
        try: oracledb.init_oracle_client(lib_dir=str(INSTANT_CLIENT_DIR))
        except Exception as e:
            if "already been initialized" not in str(e): raise
    faltando = [n for n, v in [("DB_USER",DB_USER),("DB_PASSWORD",DB_PASSWORD),
                                ("DB_HOST",DB_HOST),("DB_SERVICE",DB_SERVICE)] if not v]
    if faltando:
        sys.exit(f"ERRO: credenciais ausentes em config_local.py: {faltando}")
    conn = oracledb.connect(user=DB_USER, password=DB_PASSWORD,
                            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}")
    print(f"  Conectado! Oracle {conn.version}")
    return conn


def resolver_cogestao(conn, cotipo: int, ano: int) -> tuple[list[int], str]:
    """
    Retorna (lista_de_cogestao, nome_do_tipo).
    cotipo=0 → lista vazia (= sem filtro = todas as gestões).
    """
    cur = conn.cursor()
    cur.execute(f"SELECT NOAGREGACAO FROM MIL{ano}.TIPOAGREGACAO WHERE COTIPO = :t",
                {"t": cotipo})
    row = cur.fetchone()
    nome = row[0].strip() if row else f"Tipo {cotipo}"

    if cotipo == 0:
        return [], nome   # sem filtro

    cur.execute(f"""
        SELECT g.COGESTAO
        FROM MIL{ano}.GESTAO g
        JOIN MIL{ano}.TIPOAGREGACAOADM ta ON TRIM(ta.INTIPOADM) = TRIM(g.INTIPOADM)
        WHERE ta.COTIPO = :t
    """, {"t": cotipo})
    cogestao_list = [r[0] for r in cur.fetchall()]
    return cogestao_list, nome


# ─────────────────────────────────────────────────────────────────────────────
#  Módulos _tipoagreg (cópias separadas com suporte a cogestao_list)
# ─────────────────────────────────────────────────────────────────────────────
from controles import bf_tipoagreg        as _bf
from controles import bo_tipoagreg        as _bo
from controles import dvp_tipoagreg       as _dvp
from controles import dmpl_tipoagreg      as _dmpl
from controles import bp_tipoagreg        as _bp
from controles import balancete_tipoagreg as _bal
from controles import cruzamentos_tipoagreg as _x

MODULOS = {
    "bf":   (_bf.auditar,  "Balanço Financeiro"),
    "bo":   (_bo.auditar,  "Balanço Orçamentário"),
    "dvp":  (_dvp.auditar, "Variações Patrimoniais"),
    "dmpl": (_dmpl.auditar,"Mutações do Patrimônio Líquido"),
    "bp":   (_bp.auditar,  "Balanço Patrimonial"),
    "bal":  (_bal.auditar, "Balancete Contábil"),
}
TODOS = list(MODULOS.keys()) + ["x"]


def main():
    p = argparse.ArgumentParser(
        description="Diagnóstico de integridade por tipo de agregação de gestão")
    p.add_argument("--mes",  type=int, required=True)
    p.add_argument("--ano",  type=int, required=True)
    p.add_argument("--tipo", type=int, default=None,
                   help="Código COTIPO de TIPOAGREGACAO (0=todas, 1=Direta…; obrigatório sem --todos)")
    p.add_argument("--modulos", type=str, default="all",
                   help='Módulos: "all" ou lista separada por vírgula (ex: bf,bp,x)')
    p.add_argument("--skip",   type=str, default="")
    p.add_argument("--excel",  action="store_true")
    p.add_argument("--json",   action="store_true",
                   help="Grava JSON em painel/dados/tipoagreg/{tipo}/ e atualiza painel_diag_tipoagreg.html")
    p.add_argument("--todos",  action="store_true",
                   help=f"Roda todos os tipos nocturnos {TIPOS_NOCTURNOS} numa só conexão (ignora --tipo)")
    a = p.parse_args()

    skip = {s.strip().lower() for s in a.skip.split(",") if s.strip()}
    if a.modulos.lower() == "all":
        rodar = [m for m in TODOS if m not in skip]
    else:
        rodar = [m.strip().lower() for m in a.modulos.split(",")
                 if m.strip() and m.strip().lower() not in skip]

    if not a.todos and a.tipo is None:
        p.error("--tipo é obrigatório quando --todos não é especificado")
    tipos_a_rodar = TIPOS_NOCTURNOS if a.todos else [a.tipo]
    conn = conectar()

    # Pré-resolve cogestao para todos os tipos numa só ida ao banco
    mapa = {}
    for t in tipos_a_rodar:
        cgl, nome = resolver_cogestao(conn, t, a.ano)
        mapa[t] = {"nome": nome, "cogestao_list": cgl}

    from saida.console import imprimir_modulo, rodape
    resultados_json = {}   # tipo -> doc serializado (para --json)
    n_err_global = 0

    try:
        for tipo in tipos_a_rodar:
            nome_tipo    = mapa[tipo]["nome"]
            cogestao_list = mapa[tipo]["cogestao_list"]
            filtro_desc  = f"Tipo {tipo}: {nome_tipo}" if tipo != 0 else "Todas as gestões (sem filtro)"

            print(f"\n{'='*60}")
            print(f"  DIAGNÓSTICO CONTÁBIL — {filtro_desc}")
            print(f"  Competência: {a.mes:02d}/{a.ano}")
            if cogestao_list:
                print(f"  Gestões ({len(cogestao_list)} COGESTAO): "
                      + ", ".join(str(c) for c in sorted(cogestao_list)[:10])
                      + ("…" if len(cogestao_list) > 10 else ""))
            print('='*60)

            todos_achados = []
            totais = {}

            for mod in rodar:
                if mod == "x":
                    continue
                if mod not in MODULOS:
                    print(f"  [aviso] módulo '{mod}' desconhecido, ignorado.")
                    continue
                fn, titulo = MODULOS[mod]
                print(f"\n[{mod.upper()}] Extraindo dados...")
                achados, t = fn(conn, a.mes, a.ano, cogestao_list or None)
                totais[mod] = t
                imprimir_modulo(mod.upper(), titulo, achados)
                todos_achados.extend(achados)

            if "x" in rodar:
                print("\n[X] Rodando validações cruzadas...")
                achados_x = _x.auditar(
                    t_bf=totais.get("bf"), t_bo=totais.get("bo"),
                    t_dfc=None, t_dvp=totais.get("dvp"),
                    t_dmpl=totais.get("dmpl"), t_bp=totais.get("bp"),
                )
                imprimir_modulo("X", "Validações Cruzadas entre Demonstrativos", achados_x)
                todos_achados.extend(achados_x)

            # Resumo por tipo
            n_err = sum(1 for x in todos_achados if x.status == "ERRO")
            n_alr = sum(1 for x in todos_achados if x.status == "ALERTA")
            n_ok  = sum(1 for x in todos_achados if x.status == "OK")
            n_inf = sum(1 for x in todos_achados if x.status == "INFO")
            n_err_global += n_err
            print(f"\n=== RESUMO — {nome_tipo} ===")
            for ach in todos_achados:
                if ach.status in ("ERRO", "ALERTA"):
                    print(f"  [{ach.status:<6}] {ach.codigo:<8} {ach.titulo}")
            print(f"  TOTAL: {n_err} ERRO(S), {n_alr} ALERTA(S), {n_ok} OK, "
                  f"{n_inf} INFO  ({len(todos_achados)} controles)")

            if a.json:
                from saida import json_out
                escopo = nome_tipo
                doc = json_out.serializar(todos_achados, mes=a.mes, ano=a.ano,
                                          escopo=escopo)
                doc["tipo"]      = tipo
                doc["nome_tipo"] = nome_tipo
                sub = str(PAINEL_DIR / str(tipo))
                caminho = json_out.gravar(todos_achados, mes=a.mes, ano=a.ano,
                                          escopo=escopo, dir_dados=sub)
                print(f"  JSON ({nome_tipo}): {caminho}")
                resultados_json[tipo] = doc

            if a.excel and not a.todos:
                from saida.excel import gerar
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                nome = f"Diagnostico_tipo{tipo}_{a.ano}_{a.mes:02d}_{ts}.xlsx"
                path = OUTPUT_DIR / nome
                gerar(todos_achados, a.mes, a.ano, path)
                print(f"  Excel salvo: {path}")

            rodape(todos_achados)

    finally:
        conn.close()
        print("\n  Conexão Oracle encerrada.")

    # Atualiza painel HTML com todos os tipos rodados
    if a.json and resultados_json:
        from saida import html_out
        html_out.gerar_diag_tipoagreg(resultados_json, a.mes, a.ano)

    sys.exit(1 if n_err_global else 0)


if __name__ == "__main__":
    main()
