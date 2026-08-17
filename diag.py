#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
  DIAGNÓSTICO CONTÁBIL GDF
  Verifica integridade e validações cruzadas dos demonstrativos contábeis
  sem gerar os relatórios completos (BF, BO, DFC, DVP, DMPL, BP).

  Uso:
      python diag.py --mes 7 --ano 2026
      python diag.py --mes 7 --ano 2026 --excel
      python diag.py --mes 7 --ano 2026 --json
      python diag.py --mes 7 --ano 2026 --modulos bf,dfc,x
      python diag.py --mes 7 --ano 2026 --ug 160101

  Credenciais: edite config_local.py (nunca commite esse arquivo).
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
#  Configuração  (sobreposta por config_local.py se existir)
# ─────────────────────────────────────────────────────────────────────────────
DB_USER    = ""
DB_PASSWORD= ""
DB_HOST    = ""
DB_PORT    = "1521"
DB_SERVICE = ""
INSTANT_CLIENT_DIR = ""
OUTPUT_DIR = Path(__file__).parent / "saidas"

# Dados do painel. Fica junto do script, não em OUTPUT_DIR: o Excel é saída
# local (e está no .gitignore), enquanto os JSON são versionados no repositório.
PAINEL_DIR = Path(__file__).parent / "painel" / "dados"

try:
    import importlib.util as _ilu, pathlib as _pl
    _cfg = _pl.Path(__file__).parent / "config_local.py"
    if _cfg.exists():
        _spec = _ilu.spec_from_file_location("config_local", _cfg)
        _mod  = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_mod)
        for _k in ("DB_USER","DB_PASSWORD","DB_HOST","DB_PORT","DB_SERVICE",
                   "INSTANT_CLIENT_DIR","OUTPUT_DIR"):
            if hasattr(_mod, _k):
                globals()[_k] = getattr(_mod, _k)
        print("  [config] credenciais carregadas de config_local.py")
except Exception:
    pass


# ─────────────────────────────────────────────────────────────────────────────
#  Conexão Oracle
# ─────────────────────────────────────────────────────────────────────────────
def conectar() -> oracledb.Connection:
    if INSTANT_CLIENT_DIR:
        try:
            oracledb.init_oracle_client(lib_dir=str(INSTANT_CLIENT_DIR))
        except Exception as e:
            if "already been initialized" not in str(e):
                raise
    faltando = [n for n, v in [
        ("DB_USER", DB_USER), ("DB_PASSWORD", DB_PASSWORD),
        ("DB_HOST", DB_HOST), ("DB_SERVICE", DB_SERVICE)] if not v]
    if faltando:
        sys.exit(f"ERRO: credenciais ausentes em config_local.py: {faltando}")
    dsn  = f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}"
    conn = oracledb.connect(user=DB_USER, password=DB_PASSWORD, dsn=dsn)
    print(f"  Conectado! Oracle {conn.version}")
    return conn


# ─────────────────────────────────────────────────────────────────────────────
#  Mapa de módulos disponíveis
# ─────────────────────────────────────────────────────────────────────────────
from controles import bf   as _bf
from controles import bo   as _bo
from controles import dfc  as _dfc
from controles import dvp  as _dvp
from controles import dmpl as _dmpl
from controles import bp   as _bp
from controles import cruzamentos as _x

# Registrar módulos — adicione novos aqui conforme forem criados
MODULOS = {
    "bf":   (_bf.auditar,   "Balanço Financeiro"),
    "bo":   (_bo.auditar,   "Balanço Orçamentário"),
    "dfc":  (_dfc.auditar,  "Fluxos de Caixa"),
    "dvp":  (_dvp.auditar,  "Variações Patrimoniais"),
    "dmpl": (_dmpl.auditar, "Mutações do Patrimônio Líquido"),
    "bp":   (_bp.auditar,   "Balanço Patrimonial"),
}
TODOS = list(MODULOS.keys()) + ["x"]

# Indicadores exportados para o painel: (rótulo, módulo, chave).
# Colhidos com tolerância — se o módulo não rodou (--modulos bf) ou a chave
# não existe, o indicador simplesmente não aparece, em vez de derrubar o run.
INDICADORES = [
    ("caixa_final",        "bf", "CAIXA_FINAL"),
    ("receita_realizada",  "bf", "RECEITA_REALIZADA"),
    ("despesa_paga",       "bf", "DESPESA_PAGA"),
    ("ingressos",          "bf", "INGRESSOS"),
    ("dispendios",         "bf", "DISPENDIOS"),
    ("despesa_empenhada",  "bo", "DESP_EMPENHADA"),
    ("despesa_liquidada",  "bo", "DESP_LIQUIDADA"),
    ("ativo_total",        "bp", "ATIVO_TOTAL"),
    ("patrimonio_liquido", "bp", "PATRIMONIO_LIQUIDO"),
    ("resultado_exercicio","dvp","RESULTADO_PATRIMONIAL"),
]


def coletar_indicadores(totais: dict) -> dict:
    ind = {}
    for rotulo, mod, chave in INDICADORES:
        t = totais.get(mod)
        if t and chave in t:
            ind[rotulo] = t[chave]
    return ind


# ─────────────────────────────────────────────────────────────────────────────
#  Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(
        description="Diagnóstico de integridade dos demonstrativos contábeis GDF")
    p.add_argument("--mes",  type=int, required=True, help="Mês de referência (1-12)")
    p.add_argument("--ano",  type=int, required=True, help="Ano de referência")
    p.add_argument("--ug",   type=str, default=None,  help="Código da UG (omitir = consolidado)")
    p.add_argument("--modulos", type=str, default="all",
                   help='Módulos a rodar: "all" ou lista separada por vírgula (ex: bf,dfc,x)')
    p.add_argument("--excel", action="store_true",
                   help="Gera arquivo Excel com os resultados")
    p.add_argument("--json", action="store_true",
                   help="Grava o resultado em painel/dados/ para alimentar o painel")
    p.add_argument("--skip", type=str, default="",
                   help='Módulos a pular (ex: dvp,dmpl)')
    a = p.parse_args()

    # Resolver lista de módulos
    skip = {s.strip().lower() for s in a.skip.split(",") if s.strip()}
    if a.modulos.lower() == "all":
        rodar = [m for m in TODOS if m not in skip]
    else:
        rodar = [m.strip().lower() for m in a.modulos.split(",")
                 if m.strip() and m.strip().lower() not in skip]

    label = a.ug or "Consolidado"
    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")

    from saida.console import cabecalho, imprimir_modulo, rodape
    cabecalho(a.mes, a.ano, label)

    conn = conectar()
    todos_achados = []
    totais = {}

    try:
        for mod in rodar:
            if mod == "x":
                continue  # cruzamentos rodam depois de todos os módulos
            if mod not in MODULOS:
                print(f"  [aviso] módulo '{mod}' desconhecido, ignorado.")
                continue
            fn, titulo = MODULOS[mod]
            print(f"\n[{mod.upper()}] Extraindo dados...")
            achados, t = fn(conn, a.mes, a.ano)
            totais[mod] = t
            imprimir_modulo(mod.upper(), titulo, achados)
            todos_achados.extend(achados)

        # Cruzamentos (sempre no final, se solicitado)
        if "x" in rodar:
            print("\n[X] Rodando validações cruzadas...")
            achados_x = _x.auditar(
                t_bf   = totais.get("bf"),
                t_bo   = totais.get("bo"),
                t_dfc  = totais.get("dfc"),
                t_dvp  = totais.get("dvp"),
                t_dmpl = totais.get("dmpl"),
                t_bp   = totais.get("bp"),
            )
            from saida.console import imprimir_modulo
            imprimir_modulo("X", "Validações Cruzadas entre Demonstrativos", achados_x)
            todos_achados.extend(achados_x)

    finally:
        conn.close()
        print("\n  Conexão Oracle encerrada.")

    # Resumo compacto — facilita leitura mesmo sem cores do terminal
    print("\n=== RESUMO COMPACTO ===")
    for ach in todos_achados:
        if ach.status in ("ERRO", "ALERTA"):
            print(f"  [{ach.status:<6}] {ach.codigo:<8} {ach.titulo}")
            if ach.detalhe:
                print(f"             {ach.detalhe[:110]}")
    n_err = sum(1 for x in todos_achados if x.status == "ERRO")
    n_alr = sum(1 for x in todos_achados if x.status == "ALERTA")
    n_ok  = sum(1 for x in todos_achados if x.status == "OK")
    n_inf = sum(1 for x in todos_achados if x.status == "INFO")
    # O total é a contagem de TODOS os achados, inclusive INFO — senão a linha
    # não fecha com o número de controles executados.
    print(f"  TOTAL: {n_err} ERRO(S), {n_alr} ALERTA(S), {n_ok} OK, "
          f"{n_inf} INFO  ({len(todos_achados)} controles)")
    print("=" * 45 + "\n")
    rodape(todos_achados)

    # Excel opcional
    if a.excel:
        from saida.excel import gerar
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        nome = f"Diagnostico_{a.ano}_{a.mes:02d}_{ts}.xlsx"
        path = OUTPUT_DIR / nome
        gerar(todos_achados, a.mes, a.ano, path)
        print(f"  Excel salvo: {path}")

    # JSON para o painel — gravado ANTES do sys.exit, senão um run com erro
    # contábil (exit 1) nunca chegaria a publicar, que é justamente o dia em
    # que o painel mais importa.
    if a.json:
        from saida import json_out, html_out
        doc = json_out.serializar(
            todos_achados,
            mes=a.mes, ano=a.ano, escopo=label,
            extras=coletar_indicadores(totais),
        )
        caminho = json_out.gravar(
            todos_achados,
            mes=a.mes, ano=a.ano, escopo=label,
            extras=coletar_indicadores(totais),
            dir_dados=str(PAINEL_DIR),
        )
        print(f"  JSON salvo: {caminho}")
        html_out.gerar_diag(doc, PAINEL_DIR.parent / "painel_diag.html")

    # Exit code: 1 se houver erros
    sys.exit(1 if n_err else 0)


if __name__ == "__main__":
    main()
