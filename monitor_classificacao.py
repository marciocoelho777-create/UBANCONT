#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
  MONITOR DE CLASSIFICAÇÃO — valores que somem dos demonstrativos oficiais
=============================================================================

Sucessor do monitor_gnd.py, que só olhava GND e por isso só enxergava o
Balanço Orçamentário. Este generaliza: para CADA conta, levanta o que CADA
demonstrativo exige, e reporta de quais anexos o valor desaparece.

POR QUE UM VALOR SOME
---------------------
Os demonstrativos oficiais (PSIAG550/560) são montados sobre a BALANCOGERAL.
Cada item da Lista de Equações (MIL{ano}.ITEMBALANCO) casa registros por
igualdade de campos. Se o registro tem zerado um campo que o item exige, ele
não casa com aquele item — e se não casar com NENHUM item daquele
demonstrativo, o valor simplesmente não aparece no relatório publicado.

O QUE CADA DEMONSTRATIVO EXIGE (medido em 15/08/2026, conta 622920104)
    BF   FONTE     — 330 valores distintos em 662 itens
    BO   GND       — 8 valores em 8 itens
    DFC  FUNCAO (29) + NATUREZA DE DESPESA (64) em 140 itens
    BP   nenhum filtro
Por isso o registro de ago/2026, com COFONTE=100000000 preenchido mas GND e
natureza zerados, PERMANECEU no Balanço Financeiro e SUMIU do Orçamentário e
do Fluxo de Caixa. Confere com os oficiais de 13/08:
    BF Despesa Orçamentária   24.221.316.966,30
    BO Despesas Pagas         24.220.182.205,62
                              ─────────────────
                               1.134.760,68     <- o registro sem GND

ARMADILHA DE NOMENCLATURA
-------------------------
O campo chamado "categoria" nestas duas tabelas É O GND, não a categoria
econômica: ITEMBALANCO.COCATEGORIA e BALANCOGERAL.INCATEGORIA. Confirmado
com a área em 15/08/2026. Nas demais tabelas o nome tem o sentido usual.

COMO A REGRA É AVALIADA
-----------------------
Um item captura o registro quando TODOS os campos que ele exige batem por
igualdade. O valor sobrevive num demonstrativo se AO MENOS UM item o
capturar. Isso importa: no BF existe o item "Recursos Não Vinculados", que
não filtra fonte — então nem todo registro sem fonte some do BF. Testar
apenas "campo zerado" daria falso positivo ali.

Uso:
    python monitor_classificacao.py --mes 8 --ano 2026
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

RAIZ      = Path(__file__).parent
DIR_DADOS = RAIZ / "painel" / "dados" / "classificacao"

TIPO_BALANCO = {1: "BF", 2: "BP", 3: "DVP", 4: "BO", 5: "DRE",
                6: "BP-Empresa", 7: "DFC", 8: "DMPL"}

# ---------------------------------------------------------------------------
# Ortografia — whitelist de termos técnicos do setor público
# ---------------------------------------------------------------------------
_ORT_WHITELIST = {w.lower() for w in [
    "RPPS","RPNP","SIAFEM","SIGGO","SEFAZ","GDF","PGDF","SES","SE","SEC","SED",
    "SSP","ADASA","BRB","CAESB","NOVACAP","CEB","METRÔ","DER","DETRAN","AGEFIS",
    "SEAP","SEMA","SEMAD","SEAGRI","SEDHAB","SINESP","SETUR","SOHB","SLU",
    "DF","UO","UG","OB","NS","NE","PE","RP","PL","BP","BF","BO","DFC","DVP",
    "DRE","DMPL","FUNPRESP","IPCA","IGP","SELIC","TJLP","PTAX","BID","BIRD",
    "BNDES","CEF","STN","SOF","SIOP","SIAFI","TCE","TCU","MPDFT","TJDFT",
    "PCDF","CBMDF","PMDF","DODF","IOF","ISS","ICMS","IPI","IR","CSLL","PIS",
    "COFINS","FGTS","INSS","PASEP","FUNREBOM","DETRAF",
    "ATIVO","PASSIVO","PATRIMONIAL","ORÇAMENTÁRIO","FINANCEIRO","BALANCETE",
    "BALANÇO","RECEITA","DESPESA","SUPERÁVIT","DÉFICIT","LIQUIDADO","EMPENHADO",
    "PAGO","CANCELADO","INSCRITO","PROCESSADO",
    "INTRAORÇAMENTÁRIA","INTRAORÇAMENTÁRIAS","EXTRAORÇAMENTÁRIA",
    "EXTRAORÇAMENTÁRIAS","EXTRAORÇAMENTÁRIO","EXTRAORÇAMENTÁRIOS",
    "COMPENSATÓRIO","COMPENSATÓRIA","PECUNIÁRIA","PECUNIÁRIO",
    "TRIBUTÁRIO","TRIBUTÁRIA","FIDEJUSSÓRIA","FIDEJUSSÓRIO","CAUTELAR",
    "VINCULADO","VINCULADA","CONSIGNAÇÃO","DEDUÇÃO","TRANSFERÊNCIA","REPASSE",
    "SUBVENÇÃO","AUXÍLIO","CONTRIBUIÇÃO","AMORTIZAÇÃO","INVERSÃO","PERMUTA",
    "RESTOS","PAGAR","EXERCÍCIOS","ANTERIORES","FUTUROS","CRÉDITO","DÉBITO",
    "CONTRAPARTIDA","REGULARIZAÇÃO","AJUSTE","RECLASSIFICAÇÃO","INCORPORAÇÃO",
    "DESINCORPORAÇÃO","DEPRECIAÇÃO","EXAUSTÃO","INTEGRALIZAÇÃO",
    "PREVIDENCIÁRIO","PREVIDENCIÁRIA","MOBILIÁRIA","MOBILIÁRIO",
    "IMOBILIÁRIA","IMOBILIÁRIO","RESSARCIMENTO","RESSARCIMENTOS",
    "CONVÊNIO","CONVÊNIOS","ECONÔMICO","ECONÔMICA","ECONÔMICAS","ECONÔMICOS",
    "PATRIMÔNIO","INDENIZAÇÃO","INDENIZAÇÕES","ORÇAMENTO","ORÇAMENTÁRIA",
    "ORÇAMENTÁRIO","ORÇAMENTÁRIAS","ORÇAMENTÁRIOS",
    "DE","DA","DO","DAS","DOS","EM","NA","NO","NAS","NOS","E","A","O","AS",
    "OS","À","AO","ÀS","AOS","POR","PELO","PELA","PELOS","PELAS","COM","SEM",
    "ATÉ","APÓS","PARA",
    "VINC","INDEP","EXEC","INVEST","SERV","PÚBL","PUBL","PERM","REF","TRANSP",
    "IDENT","VLR","INCORP","MOB","ADM","PREV","DESENV","TRANSF","CONTRIB",
    "ARREC","AMORTZ","LDO","LOA","PPA","PLC","EAP","NF","NFS","DOC","TED",
    "STF","STJ","SRF","RFB","MF","ME","MGI","LIQ","LÍQ","ATIV","FINANC",
    "DEP","INSTR","INSTRUM","ASSIST","PREVID","TRAB","OBRIG","APLIC","PAT",
    "FEPECS","FHDF","HRAS","HMIB","HRC","HRG","HRT","HBDF","IGESDF","IDHAB",
    "TERRACAP","CODHAB","ESPORTE","SUPERÁVITS","DÉFICITS","CONGÊNERES",
    "CONVENIADOS","CONVENIADAS","INTRAGOVERNAMENTAIS","MULTIGOVERNAMENTAIS",
    "REEMISSÃO","PREMIAÇÕES","AUMENTATIVAS","AUMENTATIVA",
]}


def _ort_palavras(texto):
    return re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ]+(?:-[A-Za-zÀ-ÖØ-öø-ÿ]+)*", texto or "")


def _ort_check(nome, speller):
    """Retorna lista de problemas encontrados num nome de item."""
    if not nome:
        return [{"tipo": "nome vazio", "detalhe": "NOITEMBALANCO vazio/nulo", "sugestoes": []}]
    problemas = []
    nome_s = nome.strip()
    if nome_s and nome_s[0].islower():
        problemas.append({"tipo": "capitalização", "detalhe": "primeira letra minúscula", "sugestoes": []})
    if nome != nome_s:
        problemas.append({"tipo": "espaços", "detalhe": "espaço(s) na borda", "sugestoes": []})
    if "  " in nome:
        problemas.append({"tipo": "espaços", "detalhe": "espaço duplo", "sugestoes": []})
    if speller:
        for p in _ort_palavras(nome):
            pu, pl = p.upper(), p.lower()
            if pu in _ORT_WHITELIST or pl in _ORT_WHITELIST:
                continue
            if len(p) <= 2 or (p.isupper() and len(p) <= 6):
                continue
            if speller.unknown([pl]):
                sugs = sorted(speller.candidates(pl) or set())[:3]
                problemas.append({"tipo": "ortografia", "detalhe": f'"{p}"', "sugestoes": sugs})
    return problemas


def varrer_ortografia(linhas, speller):
    """Verifica ortografia dos nomes ativos do ITEMBALANCO (já carregados)."""
    vistos = set()
    achados = []
    for r in linhas:
        nome = (r.get("NOITEMBALANCO") or "").strip()
        tipo = r.get("INTIPOBALANCO")
        item_id = r.get("COITEMBALANCO")
        chave = (tipo, item_id, nome)
        if chave in vistos:
            continue
        vistos.add(chave)
        problemas = _ort_check(nome, speller)
        if problemas:
            achados.append({
                "demo":          TIPO_BALANCO.get(tipo, f"tipo{tipo}"),
                "coitembalanco": item_id,
                "noitembalanco": nome,
                "problemas":     problemas,
            })
    return achados


# ---------------------------------------------------------------------------
# Mudanças no ITEMBALANCO — detecção por snapshot diário
# ---------------------------------------------------------------------------
_SNAP_DIR = DIR_DADOS / "snaps_ib"


def _salvar_snapshot_ib(linhas, ano, gerado_em):
    """Persiste snapshot do dia e apaga mais de 40 dias antigos."""
    _SNAP_DIR.mkdir(parents=True, exist_ok=True)
    arquivos = sorted(_SNAP_DIR.glob("*.json"), reverse=True)
    for arq in arquivos[39:]:
        arq.unlink(missing_ok=True)

    itens = {}
    for r in linhas:
        k = f"{r['INTIPOBALANCO']}|{r['COITEMBALANCO']}|{r.get('COCONTACONTABIL','')}"
        itens[k] = {
            "demo": TIPO_BALANCO.get(r["INTIPOBALANCO"], f"tipo{r['INTIPOBALANCO']}"),
            "tipo": r["INTIPOBALANCO"],
            "cod":  r["COITEMBALANCO"],
            "conta": r.get("COCONTACONTABIL", ""),
            "nome": (r.get("NOITEMBALANCO") or "").strip(),
        }

    dest = _SNAP_DIR / f"{gerado_em[:10].replace('-', '')}.json"
    dest.write_text(json.dumps({"gerado_em": gerado_em, "ano": ano, "itens": itens},
                               ensure_ascii=False, separators=(",", ":")),
                    encoding="utf-8")
    return itens


def varrer_mudancas_ib(linhas, ano, gerado_em):
    """Compara ITEMBALANCO atual com snapshot de até 15 dias atrás.

    Detecta: itens novos, itens removidos, contas adicionadas/removidas,
    nomes alterados.  Retorna lista de dicts prontos para o painel.
    """
    from datetime import date, timedelta

    itens_atual = _salvar_snapshot_ib(linhas, ano, gerado_em)

    if not _SNAP_DIR.exists():
        return []

    hoje_str   = gerado_em[:10].replace("-", "")
    limite_str = (date.today() - timedelta(days=15)).strftime("%Y%m%d")

    # Snaps anteriores ao de hoje, ordenados do mais antigo ao mais recente
    candidatos = sorted(
        [a for a in _SNAP_DIR.glob("*.json") if a.stem < hoje_str],
    )
    if not candidatos:
        return []

    # Preferir o mais antigo dentro da janela de 15 dias; senão, o mais recente disponível
    dentro = [a for a in candidatos if a.stem >= limite_str]
    ref_arq = dentro[0] if dentro else candidatos[-1]

    try:
        itens_ref = json.loads(ref_arq.read_text(encoding="utf-8")).get("itens", {})
    except Exception:
        return []

    ref_date = f"{ref_arq.stem[:4]}-{ref_arq.stem[4:6]}-{ref_arq.stem[6:]}"

    mudancas = []

    # Novidades (item/conta no atual que não estava na referência)
    for k, v in itens_atual.items():
        if k not in itens_ref:
            item_chave = f"{v['tipo']}|{v['cod']}|"
            item_existia = any(kk.startswith(item_chave) for kk in itens_ref)
            mudancas.append({
                "tipo_mud":      "conta_adicionada" if item_existia else "item_novo",
                "demo":          v["demo"],
                "coitembalanco": v["cod"],
                "noitembalanco": v["nome"],
                "conta":         v["conta"],
                "desde":         ref_date,
            })

    # Remoções (item/conta na referência que sumiu no atual)
    for k, v in itens_ref.items():
        if k not in itens_atual:
            item_chave = f"{v['tipo']}|{v['cod']}|"
            item_existe = any(kk.startswith(item_chave) for kk in itens_atual)
            mudancas.append({
                "tipo_mud":      "conta_removida" if item_existe else "item_removido",
                "demo":          v["demo"],
                "coitembalanco": v["cod"],
                "noitembalanco": v["nome"],
                "conta":         v["conta"],
                "desde":         ref_date,
            })

    # Nomes alterados (mesma chave, nome diferente)
    for k in itens_atual:
        if k in itens_ref:
            nome_novo = itens_atual[k]["nome"]
            nome_ref  = itens_ref[k]["nome"]
            if nome_novo != nome_ref:
                mudancas.append({
                    "tipo_mud":      "nome_alterado",
                    "demo":          itens_atual[k]["demo"],
                    "coitembalanco": itens_atual[k]["cod"],
                    "noitembalanco": nome_novo,
                    "nome_anterior": nome_ref,
                    "conta":         itens_atual[k]["conta"],
                    "desde":         ref_date,
                })

    return mudancas

# Mapeamento natureza → GND esperado: (nat_ini, nat_fim, gnd, descricao)
# Baseado no prefixo de 2 dígitos da natureza de despesa (6 dígitos no SIGGO).
GND_NATUREZA = [
    (310000, 319999, 1, "Pessoal e Encargos Sociais"),
    (320000, 329999, 2, "Juros e Encargos da Dívida"),
    (330000, 399999, 3, "Outras Despesas Correntes"),
    (440000, 449999, 4, "Investimentos"),
    (450000, 459999, 5, "Inversões Financeiras"),
    (460000, 469999, 6, "Amortização da Dívida"),
]

def _gnd_por_nat(nat):
    """Retorna (gnd_int, descricao) pelo prefixo da natureza, ou None."""
    if not nat:
        return None
    for ini, fim, gnd, desc in GND_NATUREZA:
        if ini <= nat <= fim:
            return (gnd, desc)
    return None

# ITEMBALANCO (o que o item exige)  ->  BALANCOGERAL (o que o registro tem).
# COGESTAO e CORECEITA existem no ITEMBALANCO mas NÃO na BALANCOGERAL, então
# não são verificáveis aqui — itens que dependam só deles entram na lista de
# "não verificado" em vez de gerar conclusão errada.
CAMPOS = {
    "COCATEGORIA": "INCATEGORIA",   # GND (apesar do nome)
    "CODESPESA":   "CONATUREZA",    # natureza da despesa
    "COFONTE":     "COFONTE",
    "COUO":        "COUO",
    "COFUNCAO":    "COFUNCAO",
}
NAO_VERIFICAVEIS = ("COGESTAO", "CORECEITA")

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


def expandir_mascara(m):
    """'6322XXXXX' -> (632200000, 632299999); '622920104' -> (622920104,)*2."""
    m = str(m or "").strip().upper()
    if not re.fullmatch(r"[0-9]+X*", m):
        return None
    nx = m.count("X"); base = m.replace("X", "")
    if not base:
        return None
    ini = int(base) * (10 ** nx)
    return (ini, ini + 10 ** nx - 1)


def _norm(v):
    """Normaliza para comparação por igualdade. 0/None/'' viram None."""
    if v is None:
        return None
    s = str(v).strip()
    if s in ("", "-"):
        return None
    try:
        n = int(float(s))
        return None if n == 0 else n
    except ValueError:
        return s


SQL_ITENS = """
SELECT i.INTIPOBALANCO, i.COITEMBALANCO, i.NOITEMBALANCO, i.COCONTACONTABIL,
       i.COCATEGORIA, i.CODESPESA, i.COFONTE, i.COUO, i.COFUNCAO,
       i.COGESTAO, i.CORECEITA
FROM   MIL{ano}.ITEMBALANCO i
WHERE  i.INSTATUS = 0
  AND  i.COCONTACONTABIL IS NOT NULL
"""


def carregar_itens(conn, ano):
    """Para cada (demonstrativo, faixa de conta), a lista de exigências."""
    cur = conn.cursor(); cur.execute(SQL_ITENS.format(ano=ano))
    cols = [d[0] for d in cur.description]
    linhas = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close()

    regras = defaultdict(list)      # (dem, ini, fim) -> [ {campo_bg: valor} ]
    nao_verif = defaultdict(set)
    for r in linhas:
        f = expandir_mascara(r["COCONTACONTABIL"])
        if not f:
            continue
        dem = TIPO_BALANCO.get(r["INTIPOBALANCO"], f"tipo{r['INTIPOBALANCO']}")
        exig = {}
        for ib, bg in CAMPOS.items():
            v = _norm(r.get(ib))
            if v is not None:
                exig[bg] = v
        for nv in NAO_VERIFICAVEIS:
            if _norm(r.get(nv)) is not None:
                nao_verif[dem].add(nv)
        regras[(dem, f[0], f[1])].append(exig)
    print(f"  ITEMBALANCO: {len(linhas)} linhas ativas, "
          f"{len(regras)} combinações demonstrativo/faixa")
    for dem, campos in sorted(nao_verif.items()):
        print(f"  [aviso] {dem} usa {', '.join(sorted(campos))}, que não "
              f"existe(m) na BALANCOGERAL — não verificável")
    return regras, linhas


def demonstrativos_da_conta(regras, conta):
    """Regras aplicáveis a uma conta real, agrupadas por demonstrativo."""
    out = defaultdict(list)
    for (dem, ini, fim), itens in regras.items():
        if ini <= conta <= fim:
            out[dem].extend(itens)
    return out


def capturado(item_exig, registro):
    """O item captura o registro? Todos os campos exigidos batem por igualdade."""
    return all(registro.get(campo) == valor for campo, valor in item_exig.items())


SQL_BG = """
SELECT b.INMES, b.COCONTACONTABIL,
       b.INCATEGORIA, b.CONATUREZA, b.COFONTE, b.COUO, b.COFUNCAO,
       SUM(b.VACREDITO - b.VADEBITO) AS VLR, COUNT(*) AS QTD
FROM   MIL{ano}.BALANCOGERAL b
WHERE  b.INMES BETWEEN 1 AND {mes}
  AND  ({faixas})
GROUP  BY b.INMES, b.COCONTACONTABIL, b.INCATEGORIA, b.CONATUREZA,
          b.COFONTE, b.COUO, b.COFUNCAO
HAVING ABS(SUM(b.VACREDITO - b.VADEBITO)) > 0.01
"""


def _campos_obrigatorios(regras, conta):
    """Para cada demonstrativo, campos que TODOS os itens daquela conta exigem.

    A lógica: um campo é obrigatório num demonstrativo se ele aparece em TODOS
    os itens daquela conta — sem exceção. Se existe ao menos um item sem aquele
    campo, um registro sem o campo ainda pode casar com esse item, logo o campo
    não é obrigatório para esse demonstrativo.

    Exemplo confirmado em 15/08/2026:
      BO: todos os 8 itens da 622920104 exigem INCATEGORIA (GND 1..9).
          Logo INCATEGORIA é obrigatório. Zero = some do BO. ✔
      DFC: 140 itens exigem COFUNCAO+CONATUREZA, mas não toda despesa pertence
          ao DFC. Um item genérico sem esses filtros poderia existir, e mesmo
          sem ele, "não pertencer ao DFC" é correto, não erro.
          → DFC excluído: não há como distinguir "classificação ausente" de
            "simplesmente não pertence ao DFC" sem conhecer o escopo esperado.
      BF: há um item sem COFONTE (Recursos Não Vinculados), logo COFONTE não é
          obrigatório. Zero pode ser correto para esse item específico.
          → BF excluído pela mesma razão.

    Resultado prático: apenas o BO tem campos verdadeiramente obrigatórios
    neste conjunto de contas. O monitor reporta somente o que é demonstrável.
    """
    obrig = {}
    # Demonstrativos excluídos do monitor com justificativa documentada.
    EXCLUIDOS = {
        "DFC":  ("O DFC captura categorias específicas de fluxo, não toda "
                 "despesa de uma conta. Um registro sem COFUNCAO pode "
                 "legitimamente não pertencer ao DFC. Medido: 37.252 falsos "
                 "positivos em 15/08/2026, R$ 52,5 bi inexistentes."),
        "DVP":  "Sem campo universalmente obrigatório nas contas monitoradas.",
        "DMPL": "Sem campo universalmente obrigatório nas contas monitoradas.",
        "BP":   "O BP não filtra por nenhum campo — captura todos os registros.",
    }
    for (dem, ini, fim), itens in regras.items():
        if dem in EXCLUIDOS:
            continue
        if not (ini <= conta <= fim):
            continue
        # Campos que aparecem em TODOS os itens deste demonstrativo
        todos = None
        for exig in itens:
            campos_item = set(exig.keys())
            todos = campos_item if todos is None else todos & campos_item
        if todos:
            obrig.setdefault(dem, set()).update(todos)
    return obrig


def varrer(conn, mes, ano, regras):
    """Registros onde um campo OBRIGATÓRIO está zerado — e portanto o valor
    não pode casar com NENHUM item daquele demonstrativo.

    Só inclui demonstrativos onde existe ao menos um campo verdadeiramente
    obrigatório para a conta (ver _campos_obrigatorios). Em 15/08/2026 com
    as contas da Lista de Equações do BO, isso restringe o monitor ao BO:
    BF tem item sem COFONTE e DFC não tem campo universalmente obrigatório.

    Por que não o DFC:
    37.252 falsos positivos em 15/08/2026 — a varredura anterior testava
    "campo zerado" sem verificar se o campo era realmente obrigatório para
    aquela conta naquele demonstrativo. A maioria das despesas não pertence
    ao DFC, e isso é correto, não erro. Incluir o DFC exigiria saber QUAIS
    registros deveriam pertencer a ele, que é informação que não temos.
    """
    # Restringe às faixas do BO — único demonstrativo onde "campo obrigatório
    # zerado = valor perdido" é demonstrável sem ambiguidade. As 494 combinações
    # do ITEMBALANCO incluem contas de caixa, receita e investimento; varrer
    # todas trouxe 173.745 agregados e DFC falso-positivo em 16/08/2026.
    # A expansão para outros demonstrativos exige evidência caso a caso.
    faixas = sorted({(ini, fim) for (dem, ini, fim) in regras if dem == "BO"})
    if not faixas:
        print("  [aviso] nenhuma faixa BO encontrada — verifique o ITEMBALANCO")
        return []
    pred = " OR ".join(f"b.COCONTACONTABIL BETWEEN {a} AND {z}" for a, z in faixas)

    cur = conn.cursor()
    cur.execute(SQL_BG.format(ano=ano, mes=mes, faixas=pred))
    cols = [d[0] for d in cur.description]
    linhas = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close()
    print(f"  BALANCOGERAL: {len(linhas)} agregados analisados")

    achados = []
    for l in linhas:
        conta = int(l["COCONTACONTABIL"])
        reg = {bg: _norm(l[bg]) for bg in CAMPOS.values()}
        obrig = _campos_obrigatorios(regras, conta)

        perdidos, mantidos, zerados_tot = [], [], set()
        for dem, campos_obrig in obrig.items():
            zerados = [c for c in campos_obrig if reg.get(c) is None]
            if zerados:
                perdidos.append(dem)
                zerados_tot.update(zerados)
            else:
                mantidos.append(dem)

        if perdidos:
            # Demonstrativos onde o campo não é obrigatório: o registro
            # pode ou não aparecer — não sabemos, não reportamos.
            todos_dems = set(d for (d, ini, fim) in regras
                           if ini <= conta <= fim)
            incertos = sorted(todos_dems - set(perdidos) - set(mantidos))
            achados.append({
                "mes": int(l["INMES"]), "conta": conta,
                "valor": float(l["VLR"]), "registros": int(l["QTD"]),
                "campos": {k: v for k, v in reg.items()},
                "zerados": sorted(zerados_tot),
                "some_de": sorted(perdidos),
                "permanece_em": sorted(mantidos),
                "incerto_em": incertos,
            })
    return achados


SQL_BG_GND = """
SELECT b.INMES, b.COCONTACONTABIL,
       b.INCATEGORIA, b.CONATUREZA, b.COFONTE, b.COUO, b.COFUNCAO,
       SUM(b.VACREDITO - b.VADEBITO) AS VLR, COUNT(*) AS QTD
FROM   MIL{ano}.BALANCOGERAL b
WHERE  b.INMES BETWEEN 1 AND {mes}
  AND  ({faixas})
  AND  b.CONATUREZA > 0
  AND  b.INCATEGORIA > 0
GROUP  BY b.INMES, b.COCONTACONTABIL, b.INCATEGORIA, b.CONATUREZA,
          b.COFONTE, b.COUO, b.COFUNCAO
HAVING ABS(SUM(b.VACREDITO - b.VADEBITO)) > 0.01
"""


def varrer_gnd_errado(conn, mes, ano, regras):
    """Registros onde INCATEGORIA (GND) diverge do GND implícito pela natureza.

    Detecta classificações incorretas onde o valor APARECE no BO mas na linha
    errada — diferente do varrer() que detecta valores que SOMEM do BO.
    Exemplo: natureza 319011 (GND 1 Pessoal) com INCATEGORIA=3 (Outras DC).
    """
    faixas = sorted({(ini, fim) for (dem, ini, fim) in regras if dem == "BO"})
    if not faixas:
        return []
    pred = " OR ".join(f"b.COCONTACONTABIL BETWEEN {a} AND {z}" for a, z in faixas)

    cur = conn.cursor()
    cur.execute(SQL_BG_GND.format(ano=ano, mes=mes, faixas=pred))
    cols = [d[0] for d in cur.description]
    linhas = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close()
    print(f"  GND-cruzado: {len(linhas)} agregados analisados")

    achados = []
    for l in linhas:
        nat = _norm(l["CONATUREZA"])
        inc = _norm(l["INCATEGORIA"])
        if nat is None or inc is None:
            continue
        esp = _gnd_por_nat(nat)
        if esp is None:
            continue
        gnd_esp, desc_esp = esp
        if inc != gnd_esp:
            achados.append({
                "mes":              int(l["INMES"]),
                "conta":            int(l["COCONTACONTABIL"]),
                "valor":            float(l["VLR"]),
                "registros":        int(l["QTD"]),
                "incategoria":      inc,
                "gnd_esperado":     gnd_esp,
                "gnd_desc":         desc_esp,
                "conatureza":       nat,
                "cofonte":          _norm(l["COFONTE"]),
                "couo":             _norm(l["COUO"]),
                "cofuncao":         _norm(l["COFUNCAO"]),
            })
    return sorted(achados, key=lambda x: -abs(x["valor"]))


def _brl(v):
    if v is None:
        return "—"
    s = f"{abs(v):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return ("−" if v < 0 else "") + "R$ " + s


def markdown(doc):
    L = ["# Monitor de Classificação — valores ausentes dos demonstrativos", ""]
    achados    = doc.get("achados", [])
    gnd_errado = doc.get("gnd_errado", [])

    if not achados:
        L += [f"✅ Nenhum valor perdido até {doc['mes']:02d}/{doc['ano']}.", ""]
    else:
        L += [f"❌ **{len(achados)} agregado(s)** com classificação "
              f"incompleta até {doc['mes']:02d}/{doc['ano']}.", "",
              "## Impacto por demonstrativo", "",
              "| Demonstrativo | Agregados | Valor que não aparece |",
              "|---|---:|---:|"]
        for dem, t in sorted(doc["por_demonstrativo"].items(),
                             key=lambda x: -abs(x[1]["valor"])):
            L.append(f"| **{dem}** | {t['n']} | {_brl(t['valor'])} |")
        L += ["", "## Detalhe — campo zerado (valor some do demonstrativo)", "",
              "| Mês | Conta | Valor | Campos zerados | Some de | Permanece em |",
              "|---|---|---:|---|---|---|"]
        for a in sorted(achados, key=lambda x: -abs(x["valor"])):
            L.append(f"| {a['mes']} | {a['conta']} | {_brl(a['valor'])} | "
                     f"{', '.join(a['zerados']) or '—'} | "
                     f"**{', '.join(a['some_de'])}** | "
                     f"{', '.join(a['permanece_em']) or '—'} |")
        L += ["", "Um item captura o registro quando todos os campos que ele exige "
                  "batem por igualdade; o valor sobrevive num demonstrativo se ao "
                  "menos um item o capturar. `INCATEGORIA` é o GND, apesar do nome.", ""]

    if not gnd_errado:
        L += ["", f"✅ Nenhum GND incorreto (INCATEGORIA ≠ natureza) até "
                  f"{doc['mes']:02d}/{doc['ano']}.", ""]
    else:
        total_gnd = sum(abs(x["valor"]) for x in gnd_errado)
        L += ["", f"⚠️ **{len(gnd_errado)} registro(s)** com GND incorreto "
                  f"(INCATEGORIA ≠ natureza) — valor em linha errada do BO: "
                  f"{_brl(total_gnd)}", "",
              "## GND errado — valor em linha incorreta do BO", "",
              "| Mês | Conta | UO | Natureza | GND atual | GND esperado | Valor |",
              "|---|---|---|---|---|---|---:|"]
        for a in gnd_errado:
            L.append(f"| {a['mes']} | {a['conta']} | {a['couo'] or '—'} | "
                     f"{a['conatureza']} | **{a['incategoria']}** | "
                     f"{a['gnd_esperado']} ({a['gnd_desc']}) | {_brl(a['valor'])} |")
        L += [""]

    return "\n".join(L)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--mes", type=int, required=True)
    p.add_argument("--ano", type=int, required=True)
    a = p.parse_args()

    print(f"\n== Monitor de Classificação — {a.mes:02d}/{a.ano} ==")
    conn = conectar()
    try:
        regras, linhas_ib = carregar_itens(conn, a.ano)
        achados    = varrer(conn, a.mes, a.ano, regras)
        gnd_errado = varrer_gnd_errado(conn, a.mes, a.ano, regras)
    finally:
        conn.close(); print("  Conexão Oracle encerrada.")

    try:
        from spellchecker import SpellChecker
        speller = SpellChecker(language="pt")
    except Exception:
        speller = None
        print("  [ortografia] pyspellchecker não disponível — verificação desativada")
    gerado_em   = datetime.now().isoformat(timespec="seconds")
    mudancas_ib = varrer_mudancas_ib(linhas_ib, a.ano, gerado_em)
    if mudancas_ib:
        print(f"\n  MUDANÇAS ITEMBALANCO — {len(mudancas_ib)} mudança(s) desde referência:")
        for m in mudancas_ib[:10]:
            print(f"  >> [{m['demo']}] {m['tipo_mud']}  item {m['coitembalanco']}  "
                  f"conta {m['conta']}  \"{m['noitembalanco'][:50]}\"")
        if len(mudancas_ib) > 10:
            print(f"  ... e mais {len(mudancas_ib) - 10} mudança(s)")
    else:
        print("\n  ITEMBALANCO: nenhuma mudança detectada nos últimos 15 dias.")

    ort_achados = varrer_ortografia(linhas_ib, speller)
    if ort_achados:
        print(f"\n  ORTOGRAFIA — {len(ort_achados)} item(ns) com problema:")
        for x in ort_achados[:10]:
            tipos = ", ".join(pr["tipo"] for pr in x["problemas"])
            print(f"  !! [{x['demo']}] {x['coitembalanco']}  {tipos}  \"{x['noitembalanco'][:60]}\"")
        if len(ort_achados) > 10:
            print(f"  ... e mais {len(ort_achados) - 10} item(ns)")
    else:
        print("\n  Ortografia: nenhum problema encontrado.")

    por_dem = defaultdict(lambda: {"n": 0, "valor": 0.0})
    for x in achados:
        for dem in x["some_de"]:
            por_dem[dem]["n"] += 1
            por_dem[dem]["valor"] += x["valor"]

    # --- campo zerado ---
    for x in sorted(achados, key=lambda y: -abs(y["valor"])):
        print(f"  !! mes {x['mes']:>2}  conta {x['conta']}  {x['valor']:>18,.2f}  "
              f"zerado: {','.join(x['zerados']) or '-'}")
        print(f"        some de: {', '.join(x['some_de'])}"
              f"   permanece em: {', '.join(x['permanece_em']) or '-'}")
    print()
    for dem, t in sorted(por_dem.items(), key=lambda x: -abs(x[1]["valor"])):
        print(f"  {dem:<12} {t['n']:>3} agregado(s)   {_brl(t['valor'])}")

    # --- GND errado ---
    if gnd_errado:
        print(f"\n  GND ERRADO — {len(gnd_errado)} registro(s) com INCATEGORIA ≠ natureza:")
        for x in gnd_errado:
            print(f"  !! mes {x['mes']:>2}  conta {x['conta']}  UO {x['couo']}  "
                  f"nat {x['conatureza']}  "
                  f"GND atual={x['incategoria']} esperado={x['gnd_esperado']} "
                  f"({x['gnd_desc']})  {x['valor']:>14,.2f}")
    else:
        print("\n  GND-cruzado: nenhuma divergência INCATEGORIA × natureza.")

    doc = {"gerado_em": gerado_em,
           "mes": a.mes, "ano": a.ano,
           "por_demonstrativo": dict(por_dem),
           "achados":     achados,
           "gnd_errado":  gnd_errado,
           "mudancas_ib": mudancas_ib,
           "ortografia":  ort_achados}
    DIR_DADOS.mkdir(parents=True, exist_ok=True)
    ts = gerado_em.replace("-","").replace(":","").replace("T","_")
    dest = DIR_DADOS / f"{a.ano}-{a.mes:02d}_{ts}.json"
    dest.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    (RAIZ / "MONITOR_CLASSIFICACAO.md").write_text(markdown(doc), encoding="utf-8")
    from saida import html_out
    html_out.gerar_classificacao(doc, RAIZ / "painel" / "painel_classificacao.html")
    print(f"\n  JSON: {dest}")
    sys.exit(1 if (achados or gnd_errado) else 0)


if __name__ == "__main__":
    main()
