#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
  MONITOR GND — classificação orçamentária ausente na BALANCOGERAL
=============================================================================

O QUE ISTO RESOLVE
------------------
Os demonstrativos oficiais (PSIAG550/560) são montados sobre a tabela
MIL{ano}.BALANCOGERAL, que traz a classificação já resolvida em colunas
próprias. Quando um registro entra ali com CONATUREZA = 0 e COUO = 0, o valor
existe na contabilidade mas não casa com nenhum item do demonstrativo — os
itens filtram por GND 1..9 — e simplesmente desaparece do relatório publicado.

Medido em MIL2026, INMES 1..8:
    mês 5   622130300  Liquidada      129.929,04
    mês 6   622130300  Liquidada    1.922.061,96   (doc 2026NL12224, UG 170101)
    mês 8   622920104  Paga         1.134.760,68   (docs 2026NL05656 +
                                                    2026OB47766, UG 200101)
                                  ──────────────
                                    3.186.751,68

O caso do mês 8 explica exatamente o gap pelo qual o próprio PSIAG550 não
fecha: I+II+III − variação do caixa = 1.134.760,68. E o Anexo 12 oficial de
agosto confirma: Empenhadas e Liquidadas ficam R$ 2.051.991,00 abaixo da
contabilidade, Pagas R$ 1.134.760,68 abaixo.

Nos dois casos verificados a ORIGEM estava íntegra, por caminhos diferentes:
  - ago: NOTAEMPENHO 2026NE00023 com CONATUREZA = '339039'
  - jun: conta-corrente do lançamento terminando em '33903702'
E os formatos de conta-corrente são uniformes (622920104 sempre 11 chars em
192.068 registros; 622130300 sempre 40 chars em 300.671). Ou seja: não é
cadastro do documento, é o processo que popula a BALANCOGERAL.

O QUE ESTE SCRIPT FAZ
---------------------
1. Determina quais contas EXIGEM GND, lendo a Lista de Equações
2. Varre a BALANCOGERAL procurando registros dessas contas sem classificação
3. Para cada achado, localiza o documento de origem na LANCAMENTOCONTABIL
4. Deduz o GND que DEVERIA ter sido carregado, por duas estratégias
5. Grava JSON em painel/dados/gnd/ e um relatório Markdown para o painel

Uso:
    python monitor_gnd.py --mes 8 --ano 2026
    python monitor_gnd.py --mes 8 --ano 2026 --listas ./listas_equacoes
=============================================================================
"""
from __future__ import annotations
import argparse, glob, json, os, re, sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import oracledb

RAIZ        = Path(__file__).parent
DIR_LISTAS  = RAIZ / "listas_equacoes"      # onde ficam os .xlsx exportados
DIR_DADOS   = RAIZ / "painel" / "dados" / "gnd"
CACHE_CONTAS = RAIZ / "painel" / "contas_com_gnd.json"

# ─────────────────────────────────────────────────────────────────────────────
#  Configuração (mesma do diag.py)
# ─────────────────────────────────────────────────────────────────────────────
DB_USER = DB_PASSWORD = DB_HOST = DB_SERVICE = ""
DB_PORT = "1521"
INSTANT_CLIENT_DIR = ""
try:
    import importlib.util as _ilu
    _cfg = RAIZ / "config_local.py"
    if _cfg.exists():
        _spec = _ilu.spec_from_file_location("config_local", _cfg)
        _mod = _ilu.module_from_spec(_spec); _spec.loader.exec_module(_mod)
        for _k in ("DB_USER","DB_PASSWORD","DB_HOST","DB_PORT","DB_SERVICE",
                   "INSTANT_CLIENT_DIR"):
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
    falta = [n for n, v in [("DB_USER",DB_USER),("DB_PASSWORD",DB_PASSWORD),
                            ("DB_HOST",DB_HOST),("DB_SERVICE",DB_SERVICE)] if not v]
    if falta:
        sys.exit(f"ERRO: credenciais ausentes em config_local.py: {falta}")
    c = oracledb.connect(user=DB_USER, password=DB_PASSWORD,
                         dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}")
    print(f"  Conectado! Oracle {c.version}")
    return c


# ─────────────────────────────────────────────────────────────────────────────
#  ETAPA 1 — contas que exigem GND, a partir da Lista de Equações
# ─────────────────────────────────────────────────────────────────────────────
# A Lista vem de exportações manuais do SIGGO (.xlsx). Um script não reexporta
# sozinho: para atualização diária de verdade, é preciso descobrir a TABELA
# Oracle que guarda a Lista — o PSIAG550 a consome de algum lugar. Encontrada
# essa tabela, troque ler_listas_excel() por uma query e o monitor passa a ser
# integralmente automático. Enquanto isso, o cache em painel/contas_com_gnd.json
# preserva o último conjunto conhecido para o script não parar se o export
# estiver ausente num dia.
def expandir_mascara(m: str):
    """'6322XXXXX' -> (632200000, 632299999); '622920104' -> (622920104,)*2."""
    m = str(m).strip().upper()
    if not re.fullmatch(r"[0-9]+X*", m):
        return None
    nx = m.count("X"); base = m.replace("X", "")
    if not base:
        return None
    ini = int(base) * (10 ** nx)
    return (ini, ini + 10 ** nx - 1)


def ler_listas_excel(dir_listas: Path):
    """Devolve as faixas de conta cujos itens filtram por GND (status Ativo)."""
    import pandas as pd
    arqs = sorted(glob.glob(str(dir_listas / "*.xlsx")))
    if not arqs:
        return None
    faixas, detalhe = [], {}
    for a in arqs:
        try:
            df = pd.read_excel(a, header=5)
        except Exception as e:
            print(f"  [aviso] {os.path.basename(a)}: {type(e).__name__}: {e}")
            continue
        cols = {str(c).strip(): c for c in df.columns}
        if "GND" not in cols or "Conta Contábil" not in cols:
            continue                      # lista sem GND: BF, BP, DVP, DMPL
        d = df[df[cols["Conta Contábil"]].notna() & df[cols["GND"]].notna()]
        if "Status" in cols:
            d = d[d[cols["Status"]].astype(str).str.strip() == "Ativo"]
        for _, r in d.iterrows():
            f = expandir_mascara(r[cols["Conta Contábil"]])
            if not f:
                continue
            faixas.append(f)
            k = str(r[cols["Conta Contábil"]]).strip()
            gnd = str(r[cols["GND"]]).replace(".0", "")
            detalhe.setdefault(k, set()).add(gnd)
    if not faixas:
        return None
    faixas = sorted(set(faixas))
    return {"faixas": faixas,
            "gnd_por_mascara": {k: sorted(v) for k, v in detalhe.items()},
            "origem": "excel", "arquivos": [os.path.basename(a) for a in arqs],
            "lido_em": datetime.now().isoformat(timespec="seconds")}


# ─────────────────────────────────────────────────────────────────────────────
#  ETAPA 1b — leitura direta da Lista de Equações no Oracle (preferencial)
# ─────────────────────────────────────────────────────────────────────────────
# A Lista mora em MIL{ano}.ITEMBALANCO. Mapeamento confirmado em 15/08/2026
# comparando as contagens da tabela com as das exportações (cada export traz
# uma linha a mais, de rodapé):
#
#   INTIPOBALANCO  linhas  demonstrativo
#         1          3768  Balanço Financeiro
#         2           162  Balanço Patrimonial
#         3           148  Variações Patrimoniais
#         4           876  Balanço Orçamentário
#         5           161  DRE            (sem alteração desde 2015)
#         6            44  BP-Empresa     (sem alteração desde 2015)
#         7           873  Fluxo de Caixa
#         8            35  DMPL
#
# ATENÇÃO: o DECODE usado no relatório oficial cobre só 1..6 e devolve string
# vazia para 7 e 8 — ou seja, DFC e DMPL aparecem sem nome. Não é erro desta
# consulta, é o DECODE que está desatualizado.
#
# COLUNAS RELEVANTES
#   COCATEGORIA      = GND  (a 622920104 gera 8 itens com 1,2,3,4,5,6,7,9)
#   NUCOLUNA         = "Coluna" do export (5 = Paga, no BO)
#   INSTATUS         = 0 Ativo, 1 Inativo
#   INTIPOMOVIMENTO  = SC | SD | MC | MD
#   INOPERANDO       = + | -
#   COCONTACONTABIL  = máscara, ex.: '6322XXXXX'
#   ULTALTERACAO     = permite detectar mudança na Lista sem rele-la inteira
TIPO_BALANCO = {1:"BF", 2:"BP", 3:"DVP", 4:"BO", 5:"DRE",
                6:"BP-Empresa", 7:"DFC", 8:"DMPL"}

SQL_ITEMBALANCO = """
SELECT i.INTIPOBALANCO, i.COITEMBALANCO, i.NOITEMBALANCO,
       i.COCONTACONTABIL, i.COCATEGORIA, i.NUCOLUNA,
       i.INTIPOMOVIMENTO, i.INOPERANDO, i.INMES, i.ULTALTERACAO
FROM   MIL{ano}.ITEMBALANCO i
WHERE  i.INSTATUS = 0
  AND  i.COCATEGORIA IS NOT NULL
  AND  i.COCATEGORIA <> 0
  AND  i.COCONTACONTABIL IS NOT NULL
"""


def ler_listas_oracle(conn, ano):
    """Contas cujos itens filtram por GND (COCATEGORIA), direto do Oracle."""
    cur = conn.cursor()
    cur.execute(SQL_ITEMBALANCO.format(ano=ano))
    cols = [d[0] for d in cur.description]
    linhas = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close()
    if not linhas:
        return None
    faixas, detalhe = [], {}
    for r in linhas:
        f = expandir_mascara(r["COCONTACONTABIL"])
        if not f:
            continue
        faixas.append(f)
        k = str(r["COCONTACONTABIL"]).strip()
        detalhe.setdefault(k, set()).add(str(r["COCATEGORIA"]))
    if not faixas:
        return None
    return {"faixas": sorted(set(faixas)),
            "gnd_por_mascara": {k: sorted(v) for k, v in detalhe.items()},
            "origem": "oracle:ITEMBALANCO",
            "itens": len({(r["INTIPOBALANCO"], r["COITEMBALANCO"]) for r in linhas}),
            "demonstrativos": sorted({TIPO_BALANCO.get(r["INTIPOBALANCO"],
                                      f"tipo{r['INTIPOBALANCO']}") for r in linhas}),
            "ult_alteracao": max(str(r["ULTALTERACAO"]) for r in linhas
                                 if r["ULTALTERACAO"]),
            "lido_em": datetime.now().isoformat(timespec="seconds")}


def obter_contas_gnd(dir_listas: Path, conn=None, ano=None):
    """Ordem de preferência: Oracle (sempre atual) > Excel > cache."""
    novo = None
    if conn is not None and ano is not None:
        try:
            novo = ler_listas_oracle(conn, ano)
            if novo:
                print(f"  Lista de Equações via ITEMBALANCO: "
                      f"{len(novo['faixas'])} faixas com GND, "
                      f"{novo['itens']} itens, demonstrativos "
                      f"{', '.join(novo['demonstrativos'])} "
                      f"(últ. alteração {novo['ult_alteracao']})")
        except Exception as e:
            print(f"  [aviso] ITEMBALANCO indisponível ({type(e).__name__}: {e}) "
                  f"— tentando as planilhas")
    if not novo:
        novo = ler_listas_excel(dir_listas)
    if novo:
        # A chave 'arquivos' só existe no caminho Excel; imprimir sem checar
        # quebrava com KeyError quando a fonte era o Oracle. Cada caminho já
        # imprimiu seu próprio resumo acima.
        if novo.get("origem") == "excel":
            print(f"  Lista de Equações via planilhas: "
                  f"{len(novo['faixas'])} faixas com GND "
                  f"({len(novo.get('arquivos', []))} arquivo(s))")
        # O cache serve de rede de segurança para o dia em que o ITEMBALANCO
        # estiver indisponível e não houver export na pasta.
        try:
            CACHE_CONTAS.parent.mkdir(parents=True, exist_ok=True)
            CACHE_CONTAS.write_text(json.dumps(novo, ensure_ascii=False, indent=2),
                                    encoding="utf-8")
        except Exception as e:
            print(f"  [aviso] não foi possível gravar o cache: {type(e).__name__}: {e}")
        return novo
    if CACHE_CONTAS.exists():
        c = json.loads(CACHE_CONTAS.read_text(encoding="utf-8"))
        c["faixas"] = [tuple(f) for f in c["faixas"]]
        print(f"  [aviso] sem ITEMBALANCO e sem .xlsx — usando cache de "
              f"{c.get('lido_em','?')} ({len(c['faixas'])} faixas)")
        return c
    sys.exit(f"ERRO: ITEMBALANCO indisponível, nenhuma planilha em "
             f"{dir_listas} e nenhum cache em {CACHE_CONTAS}")


# ─────────────────────────────────────────────────────────────────────────────
#  ETAPA 2 — varredura da BALANCOGERAL
# ─────────────────────────────────────────────────────────────────────────────
def _pred_faixas(faixas, col="b.COCONTACONTABIL"):
    return " OR ".join(f"{col} BETWEEN {a} AND {z}" for a, z in faixas)


def varrer(conn, mes, ano, faixas):
    """Registros cujo GND não foi carregado na BALANCOGERAL.

    ARMADILHA DE NOMENCLATURA — LEIA ANTES DE MEXER
    ------------------------------------------------
    Nestas duas tabelas o campo chamado "categoria" É O GND, não a categoria
    econômica. Confirmado com a área em 15/08/2026: o nome está errado nas
    duas pontas e induz ao erro.

        ITEMBALANCO.COCATEGORIA   = GND do filtro do item
        BALANCOGERAL.INCATEGORIA  = GND do registro

    Nas demais tabelas "categoria" tem o significado usual. Se você vier
    consertar isto achando que faltou tratar categoria econômica: não faltou.

    O CRITÉRIO
    ----------
    O que faz o valor sumir do Balanço Orçamentário é INCATEGORIA = 0: sem
    GND, o registro não casa com nenhum item (todos filtram GND 1..9).

    Este monitor já usou CONATUREZA = 0 OR COUO = 0 como critério, o que
    trouxe 7 ocorrências em 15/08/2026 — 4 delas da 951210500, que tem
    INCATEGORIA = 3 preenchida e portanto aloca normalmente. Eram falsos
    positivos gerados pelo critério errado.

    Natureza e UO seguem sendo lidas, mas como informação: natureza é a
    origem de onde se deduz o GND esperado, e UO zerada afeta a abertura por
    unidade, não a alocação por grupo.
    """
    sql = f"""
    SELECT b.INMES, b.COCONTACONTABIL,
           NVL(b.INCATEGORIA,0) AS GND,
           NVL(b.CONATUREZA,0)  AS NATUREZA,
           NVL(b.COUO,0)        AS UO,
           NVL(b.COFONTE,0)     AS FONTE,
           SUM(b.VACREDITO - b.VADEBITO) AS VLR, COUNT(*) AS QTD
    FROM   MIL{ano}.BALANCOGERAL b
    WHERE  b.INMES BETWEEN 1 AND {mes}
      AND  ({_pred_faixas(faixas)})
      AND  (b.INCATEGORIA IS NULL OR b.INCATEGORIA = 0)
    GROUP  BY b.INMES, b.COCONTACONTABIL, NVL(b.INCATEGORIA,0),
              NVL(b.CONATUREZA,0), NVL(b.COUO,0), NVL(b.COFONTE,0)
    HAVING ABS(SUM(b.VACREDITO - b.VADEBITO)) > 0.01
    ORDER  BY 1, 2
    """
    cur = conn.cursor(); cur.execute(sql)
    cols = [d[0] for d in cur.description]
    linhas = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close()
    for l in linhas:
        l["afeta_gnd"] = True          # o filtro já é o GND ausente
        faltando = ["GND"]
        if int(l["NATUREZA"] or 0) == 0:
            faltando.append("natureza")
        if int(l["UO"] or 0) == 0:
            faltando.append("UO")
        l["classe"] = "+".join(faltando)
    return linhas


# ─────────────────────────────────────────────────────────────────────────────
#  ETAPA 3/4 — documento de origem e GND que deveria constar
# ─────────────────────────────────────────────────────────────────────────────
# Duas estratégias, porque as contas guardam a natureza em lugares diferentes:
#   (a) conta-corrente longo (40 chars), natureza nos últimos 8 dígitos
#       ex.: 622130300 -> ...000510000000033903702  ->  natureza 33903702
#   (b) conta-corrente curto (só a NE), natureza vem da NOTAEMPENHO
#       ex.: 622920104 -> '2026NE00023' -> NOTAEMPENHO.CONATUREZA = '339039'
# GND = 2º dígito da natureza (3.3.90.39 -> categoria 3, GRUPO 3).
# COMO LOCALIZAR O DOCUMENTO
# --------------------------
# A BALANCOGERAL consolida; a LANCAMENTOCONTABIL detalha. Um registro pode
# corresponder a VÁRIOS lançamentos, em VÁRIOS documentos. Medido para o caso
# de ago/2026 (R$ 1.134.760,68 na 622920104, UG 200101): a NE 2026NE00023
# agrega 2026NL05656, 2026NS00014, 2026NS00015 e 2026OB47766 — quatro
# documentos de três tipos diferentes.
#
# Por isso houve duas tentativas fracassadas antes desta:
#   1) casar VALANCAMENTO com o valor da BALANCOGERAL -> nao acha nada, o
#      valor esta partido (1.098.360,98 + 36.399,70);
#   2) agrupar por NUDOCUMENTO -> tambem nao, os lançamentos estao em
#      documentos distintos.
# A CHAVE DE AGRUPAMENTO DEPENDE DA CONTA — cada uma guarda coisa diferente
# no conta-corrente, o mesmo padrao que ja usamos para achar a natureza:
#
#   curto  (<=15 chars, ex. 622920104) -> '2026NE00023'          -> agrupa por NE
#   longo  (40 chars,   ex. 622130300) -> '...0510000000033903702'
#                                       -> agrupa pelo conta-corrente inteiro
#
# Agrupar tudo por NE fazia o caso de junho (622130300) REGREDIR: ele era
# localizado quando o agrupamento era por NUDOCUMENTO e parou de ser quando
# troquei para NE, porque nessa conta os 11 primeiros chars nao sao a nota de
# empenho e sim o inicio da classificacao orcamentaria. Troquei o agrupamento
# sem testar contra o caso que ja funcionava — dai a regressao.
#
# O IDOC da BALANCOGERAL nao serve: vem 99999 em todas as linhas.
#
# LISTAGG(DISTINCT ...) nao existe no Oracle 11.2 (ORA-30482) e o LISTAGG
# simples estoura em 4000 chars sem avisar. Usamos MIN/MAX/COUNT(DISTINCT).
SQL_DOC = """
SELECT CASE WHEN LENGTH(TRIM(o.COCONTACORRENTE)) <= 15
            THEN SUBSTR(o.COCONTACORRENTE,1,11)
            ELSE TRIM(o.COCONTACORRENTE) END AS CHAVE,
       o.COUGCONTAB, o.COGESTAOCONTAB,
       MIN(o.NUDOCUMENTO)          AS DOC_MIN,
       MAX(o.NUDOCUMENTO)          AS DOC_MAX,
       COUNT(DISTINCT o.NUDOCUMENTO) AS N_DOCS,
       MIN(o.COCONTACORRENTE)      AS CC,
       SUM(DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)) AS VLR,
       COUNT(*) AS QTD
FROM   MIL{ano}.LANCAMENTOCONTABIL o
WHERE  o.INMES = :mes
  AND  o.COCONTACONTABIL = :conta
GROUP  BY CASE WHEN LENGTH(TRIM(o.COCONTACORRENTE)) <= 15
               THEN SUBSTR(o.COCONTACORRENTE,1,11)
               ELSE TRIM(o.COCONTACORRENTE) END,
          o.COUGCONTAB, o.COGESTAOCONTAB
HAVING ABS(SUM(DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,
                      'D',-o.VALANCAMENTO,0)) - :vlr) < 0.005
"""

SQL_NE = """
SELECT CONATUREZA
FROM   MIL{ano}.NOTAEMPENHO
WHERE  NUNE = :nune AND COUG = :coug
"""


def gnd_da_natureza(nat):
    """GND é o 2º dígito da natureza (33903973 -> 3)."""
    s = re.sub(r"\D", "", str(nat or ""))
    return s[1] if len(s) >= 2 else None


def investigar(conn, ano, mes, conta, vlr):
    """Localiza o documento e deduz o GND que deveria ter sido carregado."""
    cur = conn.cursor()
    cur.execute(SQL_DOC.format(ano=ano), mes=mes, conta=int(conta), vlr=float(vlr))
    cols = [d[0] for d in cur.description]
    docs = [dict(zip(cols, r)) for r in cur.fetchall()]

    achados = []
    for d in docs:
        cc = (d["CC"] or "").strip()
        nat = origem = None
        if len(cc) >= 8 and cc[-8:].isdigit():           # estratégia (a)
            nat, origem = cc[-8:], "conta-corrente"
        else:                                            # estratégia (b)
            m = re.search(r"\d{4}NE\d+", cc)
            if m:
                try:
                    cur.execute(SQL_NE.format(ano=ano), nune=m.group(0),
                                coug=str(d["COUGCONTAB"]))
                    row = cur.fetchone()
                    if row:
                        nat, origem = row[0], f"NOTAEMPENHO {m.group(0)}"
                except Exception as e:
                    origem = f"falha ao ler NOTAEMPENHO: {type(e).__name__}"
        docs_txt = (d["DOC_MIN"] if int(d["N_DOCS"]) == 1
                    else f"{d['DOC_MIN']} .. {d['DOC_MAX']} "
                         f"({int(d['N_DOCS'])} documentos)")
        achados.append({
            "documento": docs_txt, "chave": (d["CHAVE"] or "").strip(),
            "n_documentos": int(d["N_DOCS"]), "ug": d["COUGCONTAB"],
            "gestao": d["COGESTAOCONTAB"], "conta_corrente": cc,
            "valor": float(d["VLR"]), "lancamentos": int(d["QTD"]),
            "natureza_origem": str(nat) if nat else None,
            "fonte_natureza": origem,
            "gnd_esperado": gnd_da_natureza(nat),
        })
    cur.close()
    return achados


# ─────────────────────────────────────────────────────────────────────────────
#  Saída
# ─────────────────────────────────────────────────────────────────────────────
def _brl(v):
    """Formata em padrão brasileiro. Aplicar SÓ ao número: rodar o replace
    sobre a frase inteira troca o ponto final por vírgula."""
    if v is None:
        return "—"
    s = f"{abs(v):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return ("−" if v < 0 else "") + "R$ " + s


def markdown(doc):
    L = ["# Monitor GND — classificação ausente na BALANCOGERAL", ""]
    t = doc["total"]
    doc["ocorrencias"] = doc.get("ocorrencias", [])
    if not [x for x in doc["ocorrencias"] if x.get("afeta_gnd", True)]:
        L += [f"✅ Nenhuma ocorrência até {doc['mes']:02d}/{doc['ano']}.", ""]
        return "\n".join(L)
    rel = [x for x in doc["ocorrencias"] if x.get("afeta_gnd", True)]
    L += [f"❌ **{len(rel)} ocorrência(s)** somando "
          f"**{_brl(t)}** até {doc['mes']:02d}/{doc['ano']}.", "",
          (f"_Outras {sum(1 for x in doc['ocorrencias'] if not x.get('afeta_gnd', True))} "
           f"ocorrência(s) entraram apenas por UO zerada e não afetam a "
           f"alocação por GND — ver JSON._" if any(
               not x.get("afeta_gnd", True) for x in doc["ocorrencias"]) else ""), "",
          "Estes valores existem na contabilidade e **não entram** no "
          "demonstrativo publicado: todos os itens filtram por GND 1..9 e o "
          "registro entrou na BALANCOGERAL com o GND zerado "
          "(campo `INCATEGORIA`, que apesar do nome é o GND).", "",
          "| Mês | Conta | Valor | Documento | UG | Natureza de origem | GND esperado |",
          "|---|---|---:|---|---|---|:-:|"]
    for o in [x for x in doc["ocorrencias"] if x.get("afeta_gnd", True)]:
        if o.get("gnd_esperado_bg"):
            L.append(f"| {o['mes']} | {o['conta']} | {_brl(o['valor'])} | "
                     f"— | — | {o['natureza_bg']} (BALANCOGERAL) | "
                     f"**{o['gnd_esperado_bg']}** |")
            continue
        if not o["documentos"]:
            L.append(f"| {o['mes']} | {o['conta']} | {_brl(o['valor'])} | "
                     f"_natureza zerada, doc. não localizado_ | — | — | — |")
            continue
        for d in o["documentos"]:
            L.append(f"| {o['mes']} | {o['conta']} | {_brl(d['valor'])} | "
                     f"`{d['documento']}` | {d['ug']} | "
                     f"{d['natureza_origem'] or '—'} "
                     f"({d['fonte_natureza'] or '—'}) | "
                     f"**{d['gnd_esperado'] or '?'}** |")
    L += ["", "A coluna **GND esperado** é o 2º dígito da natureza encontrada "
              "na origem — é o grupo que deveria ter sido gravado na "
              "BALANCOGERAL e que o PSIAG550 usaria para alocar o valor.", ""]
    return "\n".join(L)


def main():
    p = argparse.ArgumentParser(description="Monitor de GND ausente na BALANCOGERAL")
    p.add_argument("--mes", type=int, required=True)
    p.add_argument("--ano", type=int, required=True)
    p.add_argument("--listas", type=str, default=str(DIR_LISTAS))
    a = p.parse_args()

    print(f"\n== Monitor GND — {a.mes:02d}/{a.ano} ==")
    conn = conectar()
    try:
        contas = obter_contas_gnd(Path(a.listas), conn, a.ano)
        linhas = varrer(conn, a.mes, a.ano, contas["faixas"])
        print(f"  BALANCOGERAL: {len(linhas)} ocorrência(s) sem classificação")
        ocorrencias = []
        for l in linhas:
            # ORDEM DE BUSCA DO GND ESPERADO
            # 1) CONATUREZA da própria BALANCOGERAL: se a natureza está lá e
            #    só o GND ficou zerado, o grupo é o 2º dígito dela e não há
            #    por que consultar mais nada.
            # 2) Só quando a natureza também veio zerada (os 3 casos de 2026)
            #    é preciso descer ao documento de origem.
            nat_bg = int(l["NATUREZA"] or 0)
            gnd_bg = gnd_da_natureza(nat_bg) if nat_bg else None
            docs = [] if gnd_bg else investigar(
                conn, a.ano, int(l["INMES"]), int(l["COCONTACONTABIL"]), l["VLR"])
            ocorrencias.append({
                "mes": int(l["INMES"]), "conta": int(l["COCONTACONTABIL"]),
                "valor": float(l["VLR"]), "registros": int(l["QTD"]),
                "classe": l["classe"], "afeta_gnd": bool(l["afeta_gnd"]),
                "fonte": int(l["FONTE"] or 0),
                "gnd_gravado": int(l["GND"] or 0),
                "natureza_bg": nat_bg, "uo": int(l["UO"] or 0),
                "gnd_esperado_bg": gnd_bg,
                "documentos": docs,
            })
            print(f"  !! mes {int(l['INMES']):>2}  conta {int(l['COCONTACONTABIL'])}  "
                  f"{float(l['VLR']):>18,.2f}  ({int(l['QTD'])} reg.)  "
                  f"sem: {l['classe']}")
            if gnd_bg:
                print(f"        -> GND esperado={gnd_bg} (natureza {nat_bg} "
                      f"da propria BALANCOGERAL; so o GND ficou zerado)")
            elif not docs:
                print("        -> natureza tambem zerada e documento nao localizado")
            for d in docs:
                print(f"        -> {d['documento']}  UG {d['ug']}  "
                      f"({d['lancamentos']} lanc.)  nat={d['natureza_origem']}  "
                      f"GND esperado={d['gnd_esperado']}  ({d['fonte_natureza']})")
        print(f"\n  {len(ocorrencias)} ocorrencia(s) sem GND na BALANCOGERAL "
              f"-- todas somem do Balanco Orcamentario.")
    finally:
        conn.close()
        print("  Conexão Oracle encerrada.")

    doc = {"gerado_em": datetime.now().isoformat(timespec="seconds"),
           "mes": a.mes, "ano": a.ano,
           "faixas_com_gnd": len(contas["faixas"]),
           # guardadas para auditoria: foi ampliando as faixas do ITEMBALANCO
           # que o total passou de 3 para 7 ocorrencias, incluindo contas de
           # compensacao (9xxxxxxxx) onde a ausencia de classificacao pode ser
           # normal. Sem esta lista nao da para saber qual mascara pegou o que.
           "mascaras_gnd": contas.get("gnd_por_mascara", {}),
           "origem_listas": contas.get("origem"),
           "listas_lidas_em": contas.get("lido_em"),
           # Dois totais: somar tudo junto produz numero sem significado —
           # a 951210500 entra com sinais opostos em meses diferentes e o
           # agregado deu -14,4 mi, que nao representa nada.
           "total_gnd": sum(o["valor"] for o in ocorrencias if o["afeta_gnd"]),
           "total_outros": sum(o["valor"] for o in ocorrencias if not o["afeta_gnd"]),
           "total": sum(o["valor"] for o in ocorrencias if o["afeta_gnd"]),
           "ocorrencias": ocorrencias}

    DIR_DADOS.mkdir(parents=True, exist_ok=True)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = DIR_DADOS / f"{a.ano}-{a.mes:02d}_{carimbo}.json"
    dest.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    md = RAIZ / "MONITOR_GND.md"
    md.write_text(markdown(doc), encoding="utf-8")
    print(f"  JSON: {dest}\n  Markdown: {md}")
    print(f"  TOTAL que some do Balanco Orcamentario: {_brl(doc['total_gnd'])}")
    if doc["total_outros"]:
        print(f"  (outros, apenas UO zerada: {_brl(doc['total_outros'])})")
    sys.exit(1 if any(o["afeta_gnd"] for o in ocorrencias) else 0)


if __name__ == "__main__":
    main()
