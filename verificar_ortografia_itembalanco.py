#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
  VERIFICADOR DE ORTOGRAFIA — MIL{ano}.ITEMBALANCO
=============================================================================

Consulta o campo NOITEMBALANCO (e outros textuais) da lista de equações de
balanço e reporta:

  1. Erros ortográficos em português (pyspellchecker, pt-BR)
  2. Capitalização inconsistente (primeira letra minúscula)
  3. Espaços duplicados / espaços à esquerda ou direita
  4. Palavras isoladas suspeitas (muito curtas e não reconhecidas)

Uso:
    python verificar_ortografia_itembalanco.py --ano 2026
    python verificar_ortografia_itembalanco.py --ano 2026 --demo BO
    python verificar_ortografia_itembalanco.py --ano 2026 --mostrar-todos
=============================================================================
"""
from __future__ import annotations
import argparse, json, re, sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import oracledb

RAIZ = Path(__file__).parent

TIPO_BALANCO = {
    1: "Balanço Financeiro",
    2: "Balanço Patrimonial",
    3: "Variação Patrimonial",
    4: "Balanço Orçamentário",
    5: "DRE",
    6: "Balanço Patrimonial-Empresa",
    7: "DFC",
    8: "DMPL",
}
SIGLA_DEMO = {
    1: "BF", 2: "BP", 3: "DVP", 4: "BO", 5: "DRE", 6: "BP-Emp", 7: "DFC", 8: "DMPL",
}

# ---------------------------------------------------------------------------
# Whitelist — termos técnicos do setor público que NÃO são erros ortográficos
# ---------------------------------------------------------------------------
WHITELIST = {w.lower() for w in [
    # Siglas e acrônimos
    "RPPS", "RPNP", "SIAFEM", "SIGGO", "SEFAZ", "GDF", "PGDF", "SES", "SE",
    "SEC", "SED", "SSP", "ADASA", "BRB", "CAESB", "NOVACAP", "CEB", "METRÔ",
    "DER", "DETRAN", "AGEFIS", "SEAP", "SEMA", "SEMAD", "SEAGRI", "SEDHAB",
    "SINESP", "SETUR", "SOHB", "SLU", "DF", "UO", "UG", "OB", "NS", "NE", "PE",
    "RP", "PL", "BP", "BF", "BO", "DFC", "DVP", "DRE", "DMPL", "FUNPRESP",
    "IPCA", "IGP", "SELIC", "TJLP", "PTAX", "BID", "BIRD", "BNDES", "CEF",
    "STN", "SOF", "SIOP", "SIAFI", "TCE", "TCU", "MPDFT", "TJDFT", "PCDF",
    "CBMDF", "PMDF", "DODF", "IOF", "ISS", "ICMS", "IPI", "IR", "CSLL", "PIS",
    "COFINS", "FGTS", "INSS", "PASEP", "FUNREBOM", "DETRAF",
    # Termos contábeis e jurídicos
    "ATIVO", "PASSIVO", "PATRIMONIAL", "ORÇAMENTÁRIO", "FINANCEIRO",
    "BALANCETE", "BALANÇO", "RECEITA", "DESPESA", "SUPERÁVIT", "DÉFICIT",
    "LIQUIDADO", "EMPENHADO", "PAGO", "CANCELADO", "INSCRITO", "PROCESSADO",
    "INTRAORÇAMENTÁRIA", "INTRAORÇAMENTÁRIAS", "EXTRAORÇAMENTÁRIA",
    "EXTRAORÇAMENTÁRIAS", "EXTRAORÇAMENTÁRIO", "EXTRAORÇAMENTÁRIOS",
    "COMPENSATÓRIO", "COMPENSATÓRIA", "PECUNIÁRIA", "PECUNIÁRIO",
    "TRIBUTÁRIO", "TRIBUTÁRIA", "FIDEJUSSÓRIA", "FIDEJUSSÓRIO",
    "CAUTELAR", "VINCULADO", "VINCULADA", "CONSIGNAÇÃO", "DEDUÇÃO",
    "TRANSFERÊNCIA", "REPASSE", "SUBVENÇÃO", "AUXÍLIO", "CONTRIBUIÇÃO",
    "AMORTIZAÇÃO", "INVERSÃO", "PERMUTA", "RESTOS", "PAGAR", "EXERCÍCIOS",
    "ANTERIORES", "FUTUROS", "CRÉDITO", "DÉBITO", "CONTRAPARTIDA",
    "REGULARIZAÇÃO", "AJUSTE", "RECLASSIFICAÇÃO", "INCORPORAÇÃO",
    "DESINCORPORAÇÃO", "DEPRECIAÇÃO", "EXAUSTÃO",
    "INTEGRALIZAÇÃO", "PREVIDENCIÁRIO", "PREVIDENCIÁRIA",
    "MOBILIÁRIA", "MOBILIÁRIO", "IMOBILIÁRIA", "IMOBILIÁRIO",
    "RESSARCIMENTO", "RESSARCIMENTOS",
    # Variantes pt-BR corretas (dicionário usa pt-PT e sugere formas erradas)
    "CONVÊNIO", "CONVÊNIOS", "ECONÔMICO", "ECONÔMICA", "ECONÔMICAS", "ECONÔMICOS",
    "PATRIMÔNIO", "INDENIZAÇÃO", "INDENIZAÇÕES", "INDENIZAR",
    "ORÇAMENTO", "ORÇAMENTÁRIA", "ORÇAMENTÁRIO", "ORÇAMENTÁRIAS", "ORÇAMENTÁRIOS",
    # Preposições e artigos que aparecem sozinhos em nomes
    "DE", "DA", "DO", "DAS", "DOS", "EM", "NA", "NO", "NAS", "NOS",
    "E", "A", "O", "AS", "OS", "À", "AO", "ÀS", "AOS", "POR", "PELO",
    "PELA", "PELOS", "PELAS", "COM", "SEM", "ATÉ", "APÓS", "PARA",
    # Abreviações intencionais (limite de campo no SIGGO)
    "VINC", "INDEP", "EXEC", "INVEST", "SERV", "PÚBL", "PUBL", "PERM",
    "REF", "TRANSP", "IDENT", "VLR", "INCORP", "MOB", "ADM", "PREV",
    "DESENV", "TRANSF", "CONTRIB", "ARREC", "AMORTZ", "LIQUIDEZ",
    # Abreviações com ponto (aparecem sem ponto após tokenização)
    "LDO", "LOA", "PPA", "PLC", "EAP", "NF", "NFS",
    "DOC", "TED", "STF", "STJ", "SRF", "RFB", "MF", "ME", "MGI",
    # Termos específicos GDF
    "FEPECS", "FHDF", "HRAS", "HMIB", "HRC", "HRG", "HRT", "HBDF",
    "IGESDF", "IDHAB", "TERRACAP", "CODHAB", "ESPORTE",
]}

# ---------------------------------------------------------------------------
# Credenciais Oracle (mesmo padrão do monitor_classificacao.py)
# ---------------------------------------------------------------------------
DB_USER = DB_PASSWORD = DB_HOST = DB_SERVICE = ""
DB_PORT = "1521"
INSTANT_CLIENT_DIR = ""
try:
    import importlib.util as _ilu
    _cfg = RAIZ / "config_local.py"
    if _cfg.exists():
        _spec = _ilu.spec_from_file_location("config_local", _cfg)
        _mod = _ilu.module_from_spec(_spec); _spec.loader.exec_module(_mod)
        for _k in ("DB_USER", "DB_PASSWORD", "DB_HOST", "DB_PORT",
                   "DB_SERVICE", "INSTANT_CLIENT_DIR"):
            if hasattr(_mod, _k):
                globals()[_k] = getattr(_mod, _k)
        print("  [config] credenciais carregadas de config_local.py")
except Exception:
    pass


def conectar():
    if INSTANT_CLIENT_DIR:
        try:
            oracledb.init_oracle_client(lib_dir=str(INSTANT_CLIENT_DIR))
        except Exception as e:
            if "already been initialized" not in str(e):
                raise
    falta = [n for n, v in [("DB_USER", DB_USER), ("DB_PASSWORD", DB_PASSWORD),
                              ("DB_HOST", DB_HOST), ("DB_SERVICE", DB_SERVICE)] if not v]
    if falta:
        sys.exit(f"ERRO: credenciais ausentes em config_local.py: {falta}")
    c = oracledb.connect(user=DB_USER, password=DB_PASSWORD,
                          dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}")
    print(f"  Conectado! Oracle {c.version}")
    return c


SQL_ITENS = """
SELECT COITEMBALANCO, COITEMBALANCO2, COCONTACONTABIL, NOITEMBALANCO,
       INTIPOBALANCO, INSTATUS, NULINHA, NUCOLUNA, ULTALTERACAO,
       INEXCNIVELAGRE, INEXCGERAL, INOPERANDO, INEXCLUI,
       COGESTAO, COFONTE, CORECEITA, COUO, COFUNCAO, CODESPESA,
       COCATEGORIA, COTIPO, INTIPOMOVIMENTO, NUUSUARIO, INMES
FROM   MIL{ano}.ITEMBALANCO
ORDER  BY INTIPOBALANCO, COITEMBALANCO
"""


def carregar_itens(conn, ano):
    cur = conn.cursor()
    cur.execute(SQL_ITENS.format(ano=ano))
    cols = [d[0] for d in cur.description]
    linhas = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close()
    print(f"  ITEMBALANCO: {len(linhas)} itens carregados")
    return linhas


# ---------------------------------------------------------------------------
# Verificações
# ---------------------------------------------------------------------------

def _palavras(texto):
    """Retorna lista de palavras do texto (apenas letras e hífen)."""
    return re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ]+(?:-[A-Za-zÀ-ÖØ-öø-ÿ]+)*", texto or "")


def verificar_capitalizacao(nome):
    """Primeira letra do nome deve ser maiúscula."""
    if not nome:
        return None
    nome = nome.strip()
    if nome and nome[0].islower():
        return f'Primeira letra minúscula: "{nome}"'
    return None


def verificar_espacos(nome):
    """Detecta espaços duplos ou espaços nas bordas."""
    if not nome:
        return None
    problemas = []
    if nome != nome.strip():
        problemas.append("espaço(s) na borda")
    if "  " in nome:
        problemas.append("espaço duplo")
    return (", ".join(problemas) + f': "{nome}"') if problemas else None


def verificar_ortografia(nome, speller):
    """Retorna lista de palavras desconhecidas (possíveis erros ortográficos)."""
    if not nome or speller is None:
        return []
    palavras = _palavras(nome)
    erros = []
    for p in palavras:
        pu = p.upper()
        pl = p.lower()
        if pu in WHITELIST or pl in WHITELIST:
            continue
        if len(p) <= 2:
            continue
        if p.isupper() and len(p) <= 6:  # sigla curta
            continue
        if re.fullmatch(r"\d+", p):
            continue
        # pyspellchecker trabalha com minúsculas
        desconhecidas = speller.unknown([pl])
        if desconhecidas:
            sugestoes = speller.candidates(pl) or set()
            sugestoes_str = ", ".join(sorted(sugestoes)[:3]) if sugestoes else "—"
            erros.append({"palavra": p, "sugestoes": sugestoes_str})
    return erros


def analisar(linhas, speller, demo_filtro=None):
    """Analisa cada item e retorna lista de achados, deduplucando por (demo, coitembalanco, nome)."""
    # chave: (sigla_demo, coitembalanco, noitembalanco) -> achado
    vistos = {}
    for r in linhas:
        tipo = r.get("INTIPOBALANCO")
        sigla = SIGLA_DEMO.get(tipo, f"tipo{tipo}")
        if demo_filtro and sigla not in demo_filtro:
            continue

        nome = (r.get("NOITEMBALANCO") or "").strip()
        item_id = r.get("COITEMBALANCO")
        conta = r.get("COCONTACONTABIL") or "—"
        status = r.get("INSTATUS")

        chave = (sigla, item_id, nome)
        if chave in vistos:
            # já analisado — apenas acumula contas distintas
            vistos[chave]["contas"].add(str(conta).strip())
            continue

        problemas = []

        cap = verificar_capitalizacao(nome)
        if cap:
            problemas.append({"tipo": "capitalização", "detalhe": cap})

        esp = verificar_espacos(nome)
        if esp:
            problemas.append({"tipo": "espaços", "detalhe": esp})

        if not nome:
            problemas.append({"tipo": "nome vazio", "detalhe": "NOITEMBALANCO vazio/nulo"})

        erros_ort = verificar_ortografia(nome, speller)
        for e in erros_ort:
            problemas.append({
                "tipo": "ortografia",
                "detalhe": f'"{e["palavra"]}" — sugestões: {e["sugestoes"]}',
            })

        if problemas:
            vistos[chave] = {
                "demo": sigla,
                "demo_nome": TIPO_BALANCO.get(tipo, f"tipo{tipo}"),
                "coitembalanco": item_id,
                "contas": {str(conta).strip()},
                "status": status,
                "noitembalanco": nome,
                "problemas": problemas,
            }

    # Converte contas de set para lista ordenada
    achados = []
    for a in vistos.values():
        a["contas"] = sorted(a["contas"])
        achados.append(a)
    return achados


# ---------------------------------------------------------------------------
# Saída
# ---------------------------------------------------------------------------

def _emoji_tipo(tipo):
    return {"capitalização": "🔡", "espaços": "⎵", "nome vazio": "❌",
            "ortografia": "📝"}.get(tipo, "⚠️")


def imprimir(achados, mostrar_todos=False):
    if not achados:
        print("\n  ✅ Nenhum problema encontrado.")
        return
    por_demo = defaultdict(list)
    for a in achados:
        por_demo[a["demo"]].append(a)
    for demo, itens in sorted(por_demo.items()):
        print(f"\n  ── {demo} ({len(itens)} nome(s) único(s) com problema) ──")
        for a in sorted(itens, key=lambda x: str(x["coitembalanco"])):
            st = " [INATIVO]" if a["status"] != 0 else ""
            contas_str = ", ".join(a["contas"]) if len(a["contas"]) <= 4 else \
                         ", ".join(a["contas"][:4]) + f" +{len(a['contas'])-4}"
            print(f"    [{a['coitembalanco']}] {a['noitembalanco']!r}{st}")
            print(f"         contas: {contas_str}")
            for p in a["problemas"]:
                print(f"      {_emoji_tipo(p['tipo'])} {p['tipo'].upper()}: {p['detalhe']}")


def gerar_markdown(achados, ano, speller_disponivel):
    L = [f"# Verificação de Ortografia — MIL{ano}.ITEMBALANCO", "",
         f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ""]

    if not speller_disponivel:
        L += ["> ⚠️ **pyspellchecker não instalado** — verificação ortográfica desativada.",
              "> Apenas capitalização e espaços foram verificados.", ""]

    if not achados:
        L += ["✅ Nenhum problema encontrado.", ""]
        return "\n".join(L)

    por_demo = defaultdict(list)
    for a in achados:
        por_demo[a["demo"]].append(a)

    # Resumo
    total = len(achados)
    L += [f"## Resumo — {total} item(s) com problema", "",
          "| Demonstrativo | Itens | Tipos de problema |",
          "|---|---:|---|"]
    for demo, itens in sorted(por_demo.items()):
        tipos = sorted({p["tipo"] for a in itens for p in a["problemas"]})
        L.append(f"| {demo} | {len(itens)} | {', '.join(tipos)} |")
    L += [""]

    # Detalhe por demonstrativo
    for demo, itens in sorted(por_demo.items()):
        L += [f"## {demo} — {TIPO_BALANCO.get(next(a['demo_nome'] for a in itens if a['demo'] == demo), demo)}", ""]
        for a in sorted(itens, key=lambda x: str(x["coitembalanco"])):
            st = " *(inativo)*" if a["status"] != 0 else ""
            contas = a.get("contas", [])
            contas_str = ", ".join(f"`{c}`" for c in contas[:4])
            if len(contas) > 4:
                contas_str += f" +{len(contas)-4}"
            L += [f"### Item {a['coitembalanco']}{st}",
                  f"> {a['noitembalanco'] or '*(vazio)*'}",
                  f"> contas: {contas_str or '—'}", ""]
            for p in a["problemas"]:
                L.append(f"- {_emoji_tipo(p['tipo'])} **{p['tipo'].upper()}**: {p['detalhe']}")
            L += [""]
    return "\n".join(L)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--ano", type=int, required=True)
    p.add_argument("--demo", nargs="+",
                   help="Filtrar por sigla de demonstrativo: BF BO BP DVP DFC DRE DMPL")
    p.add_argument("--mostrar-todos", action="store_true",
                   help="Mostrar todos os itens (inclusive sem problema) no terminal")
    a = p.parse_args()

    print(f"\n== Verificação de Ortografia ITEMBALANCO — {a.ano} ==")

    # Spell checker (opcional — não falha se não estiver instalado)
    speller = None
    try:
        from spellchecker import SpellChecker
        speller = SpellChecker(language="pt")
        print("  pyspellchecker: ativo (pt-BR)")
    except ImportError:
        print("  pyspellchecker: NÃO instalado — só capitalização/espaços serão verificados")
        print("  Para instalar: pip install pyspellchecker")

    conn = conectar()
    try:
        linhas = carregar_itens(conn, a.ano)
    finally:
        conn.close()
        print("  Conexão Oracle encerrada.")

    achados = analisar(linhas, speller, demo_filtro=a.demo)
    imprimir(achados, mostrar_todos=a.mostrar_todos)

    md = gerar_markdown(achados, a.ano, speller is not None)
    dest_md = RAIZ / "ORTOGRAFIA_ITEMBALANCO.md"
    dest_md.write_text(md, encoding="utf-8")

    dest_json = RAIZ / f"ortografia_itembalanco_{a.ano}.json"
    dest_json.write_text(
        json.dumps({"gerado_em": datetime.now().isoformat(timespec="seconds"),
                    "ano": a.ano, "total_itens": len(linhas),
                    "achados": achados},
                   ensure_ascii=False, indent=2),
        encoding="utf-8")

    print(f"\n  Relatório MD : {dest_md}")
    print(f"  JSON         : {dest_json}")
    print(f"\n  Total de itens analisados : {len(linhas)}")
    print(f"  Itens com problema        : {len(achados)}")
    sys.exit(1 if achados else 0)


if __name__ == "__main__":
    main()
