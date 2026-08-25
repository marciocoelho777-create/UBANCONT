# -*- coding: utf-8 -*-
"""
=============================================================================
  DEMONSTRAÇÃO DAS VARIAÇÕES PATRIMONIAIS - GDF  (PSIAG550 / Anexo 15 / DVP)
  SQL fixo, equações embutidas conforme a Lista de Equações de Balanço,
  Tipo de Balanço 3 - DVP (MIL{ano}.ITEMBALANCO, INTIPOBALANCO=3),
  seguindo o MESMO padrão de extração Oracle já validado e em produção nos
  projetos irmãos: dmpl.py, dfc.py, bf.py, bo.py, bp.py.

  ORIGEM DAS CATEGORIAS (extraídas de ITEMBALANCO em 17/08/2026, não
  chutadas -- consultar MIL{ano}.ITEMBALANCO WHERE INTIPOBALANCO=3 AND
  INSTATUS=0 para conferir):

  VARIAÇÕES PATRIMONIAIS AUMENTATIVAS (classe 4, SC = crédito - débito):
    41XXXXXXX  Impostos, Taxas e Contribuições de Melhoria
    42XXXXXXX  Contribuições
    43XXXXXXX  Exploração e Venda de Bens, Serviços e Direitos
    44XXXXXXX  Variações Patrimoniais Aumentativas Financeiras
    45XXXXXXX  Transferências e Delegações Recebidas
    46XXXXXXX  Valorização e Ganhos c/ Ativos e Desinc. de Passivos
    49XXXXXXX  Outras Variações Patrimoniais Aumentativas
    4XXXXXXXX  TOTAL (I) -- soma das 7 linhas acima

  VARIAÇÕES PATRIMONIAIS DIMINUTIVAS (classe 3, SD = débito - crédito):
    31XXXXXXX  Pessoal e Encargos
    32XXXXXXX  Benefícios Previdenciários e Assistenciais
    33XXXXXXX  Uso de Bens, Serviços e Consumo de Capital Fixo
    34XXXXXXX  Variações Patrimoniais Diminutivas Financeiras
    35XXXXXXX  Transferências e Delegações Concedidas
    36XXXXXXX  Desvalorização e Perda de Ativos e Incorporação de Passivos
    37XXXXXXX  Tributárias
    39XXXXXXX  Outras Variações Patrimoniais Diminutivas
    3XXXXXXXX  TOTAL (II) -- soma das 8 linhas acima

  RESULTADO PATRIMONIAL DO PERÍODO (III) = (I - II)

  GRUPO 38 (Custo de Mercadorias/Produtos Vendidos e Serviços Prestados):
  até 24/08/2026 existia como item na Lista de Equações (COITEMBALANCO
  20800000000) mas SEM máscara de conta associada (COCONTACONTABIL em
  branco) -- não tinha categoria própria de exibição, apesar de seus
  valores (quando existiam) já fazerem parte da faixa 3XXXXXXXX do TOTAL
  oficial. Confirmado em 17/08/2026 (mês 07/2026): a soma das 8 categorias
  de VPD nomeadas dava 108.458.651.150,25, mas o TOTAL oficial (I) era
  108.458.701.490,25 -- gap de R$ 50.340,00, IDÊNTICO no PDF oficial da
  pasta 13. Em 25/08/2026 a Lista de Equações passou a trazer a conta
  381110000 associada ao item 2.08.00.00.00.00 ("CUSTO DAS MERC. E PROD.
  VENDIDOS..."), então o grupo 38 ganhou categoria própria (ver
  CATEGORIAS_VPD abaixo) e o gap de R$ 50.340,00 fecha. TOTAL_VPA/TOTAL_VPD
  continuam vindo SEMPRE da faixa direta da classe inteira
  (300000000-399999999 / 400000000-499999999), nunca da soma das
  categorias nomeadas -- mesmo padrão já validado em controles/dvp.py; a
  categoria nova só melhora a abertura exibida, não muda o total.

  VARIAÇÕES PATRIMONIAIS QUALITATIVAS (Anexo da DVP, contas de controle 95/96):
    96131XXXX  Incorporação de Ativo               (SC)
    96133XXXX  Desincorporação de Passivo           (SC)
    95133XXXX  Incorporação de Passivo              (SD)
    95131XXXX  Desincorporação de Ativo             (SD)

  FONTE E ACUMULAÇÃO (mesmo padrão já validado em controles/dvp.py, cujos
  totais batem exatos com o oficial -- ver comparação 17/08/2026, mês
  07/2026 fechado: VPA/VPD/Resultado idênticos ao PSIAG550):
    Exercício Atual    -> MIL{ano}.VSALDOCONTABIL,   INMES BETWEEN 1 AND :mes
    Exercício Anterior -> MIL{ano-1}.VSALDOCONTABIL, INMES BETWEEN 1 AND 13
    (mesma regra "Exercício Anterior = INMES 1..13 do schema MIL{ano-1}"
    documentada e validada em bp.py)

  Dependências:  pip install oracledb openpyxl pandas reportlab
  Uso:
      python dvp.py --mes 7 --ano 2026
      python dvp.py --mes 7 --ano 2026 --ug 130101
      python dvp.py --mes 7 --ano 2026 --formato pdf
=============================================================================
"""
import argparse, sys
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import oracledb
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIGURACAO  (mesmos valores do BF/BO/BP/DFC/DMPL)
# ─────────────────────────────────────────────────────────────────────────────
DB_USER = DB_PASSWORD = DB_HOST = DB_SERVICE = ""
DB_PORT = 1521
INSTANT_CLIENT_DIR = ""
try:
    import importlib.util as _ilu, pathlib as _pl
    _cfg = _pl.Path(__file__).parent / "config_local.py"
    if _cfg.exists():
        _spec = _ilu.spec_from_file_location("config_local", _cfg)
        _mod  = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_mod)
        for _k in ("DB_USER","DB_PASSWORD","DB_HOST","DB_PORT",
                   "DB_SERVICE","INSTANT_CLIENT_DIR"):
            if hasattr(_mod, _k):
                globals()[_k] = getattr(_mod, _k)
except Exception:
    pass
OUTPUT_DIR = Path(r"C:\balanço 2026 gemini arquivos")

MESES = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',
         6:'Junho',7:'Julho',8:'Agosto',9:'Setembro',
         10:'Outubro',11:'Novembro',12:'Dezembro'}

# ─────────────────────────────────────────────────────────────────────────────
#  CATEGORIAS  (chave, descrição, máscara de conta, tipo de movimento)
#  Fonte: MIL2026.ITEMBALANCO, INTIPOBALANCO=3, INSTATUS=0 (consultado
#  17/08/2026 -- ver docstring do módulo).
# ─────────────────────────────────────────────────────────────────────────────
CATEGORIAS_VPA = [
    ("IMPOSTOS",        "Impostos, Taxas e Contribuições de Melhoria",          "41XXXXXXX"),
    ("CONTRIBUICOES",   "Contribuições",                                        "42XXXXXXX"),
    ("EXPLOR_VENDA",    "Exploração e Venda de Bens, Serviços e Direitos",       "43XXXXXXX"),
    ("VPA_FINANCEIRAS", "Variações Patrimoniais Aumentativas Financeiras",       "44XXXXXXX"),
    ("TRANSF_RECEB",    "Transferências e Delegações Recebidas",                 "45XXXXXXX"),
    ("VALORIZ_GANHOS",  "Valorização e Ganhos c/ Ativos e Desinc. de Passivos",  "46XXXXXXX"),
    ("OUTRAS_VPA",      "Outras Variações Patrimoniais Aumentativas",            "49XXXXXXX"),
]
CATEGORIAS_VPD = [
    ("PESSOAL",         "Pessoal e Encargos",                                      "31XXXXXXX"),
    ("BENEF_PREV",      "Benefícios Previdenciários e Assistenciais",              "32XXXXXXX"),
    ("USO_BENS",        "Uso de Bens, Serviços e Consumo de Capital Fixo",         "33XXXXXXX"),
    ("VPD_FINANCEIRAS", "Variações Patrimoniais Diminutivas Financeiras",          "34XXXXXXX"),
    ("TRANSF_CONC",     "Transferências e Delegações Concedidas",                  "35XXXXXXX"),
    ("DESVALOR_PERDA",  "Desvalorização e Perda de Ativos e Incorporação de Passivos", "36XXXXXXX"),
    ("TRIBUTARIAS",     "Tributárias",                                              "37XXXXXXX"),
    ("CUSTO_MERC_PROD", "Custo de Mercadorias, Prod. Vendidos e Serviços Prestados", "381110000"),
    ("OUTRAS_VPD",      "Outras Variações Patrimoniais Diminutivas",               "39XXXXXXX"),
]
QUALITATIVAS = [
    ("INCORP_ATIVO",    "Incorporação de Ativo",        "96131XXXX", "SC"),
    ("DESINCORP_PASSIVO","Desincorporação de Passivo",   "96133XXXX", "SC"),
    ("INCORP_PASSIVO",  "Incorporação de Passivo",       "95133XXXX", "SD"),
    ("DESINCORP_ATIVO", "Desincorporação de Ativo",      "95131XXXX", "SD"),
]

# Conta de encerramento -- fonte independente do Resultado Patrimonial,
# mesma regra já validada em controles/dvp.py: só é comparável no
# fechamento do exercício (INMES 13..15); fora disso o "erro" é falso
# positivo (grandezas incompatíveis -- ver docstring de controles/dvp.py).
CONTA_ENCERRAMENTO = "891XXXXXX"


def _faixa(mascara):
    """'41XXXXXXX' -> (410000000, 419999999). Mesmo padrão de dmpl.py."""
    n_x = mascara.count('X')
    prefixo = mascara.replace('X', '')
    ini = int(prefixo) * (10 ** n_x)
    fim = ini + (10 ** n_x) - 1
    return ini, fim


def _clausula(alias, mascara, tipo_mov, alias_tabela='v'):
    ini, fim = _faixa(mascara)
    cond_conta = f"{alias_tabela}.COCONTACONTABIL BETWEEN {ini} AND {fim}"
    if tipo_mov == "SC":
        formula = f"{alias_tabela}.VACREDITO - {alias_tabela}.VADEBITO"
    else:  # SD
        formula = f"{alias_tabela}.VADEBITO - {alias_tabela}.VACREDITO"
    return (f"    SUM(CASE WHEN {alias_tabela}.INMES BETWEEN :mi AND :mf "
            f"AND {cond_conta}\n"
            f"         THEN {formula} ELSE 0 END) AS {alias}")


def montar_sql():
    """SQL único (VSALDOCONTABIL) para todas as categorias de VPA, VPD e
    Qualitativas, mais a conta de encerramento (891) e sua movimentação
    avulsa no período -- mesmo padrão já validado em controles/dvp.py."""
    partes = []
    for chave, _, mascara in CATEGORIAS_VPA:
        partes.append(_clausula(chave, mascara, "SC"))
    for chave, _, mascara in CATEGORIAS_VPD:
        partes.append(_clausula(chave, mascara, "SD"))
    for chave, _, mascara, tipo_mov in QUALITATIVAS:
        partes.append(_clausula(chave, mascara, tipo_mov))

    # TOTAL direto da classe inteira (300000000-399999999 / 400000000-
    # 499999999) -- NÃO a soma das categorias nomeadas, por causa do grupo
    # 38 sem categoria própria (ver docstring do módulo).
    partes.append(
        "    SUM(CASE WHEN v.INMES BETWEEN :mi AND :mf\n"
        "         AND v.COCONTACONTABIL BETWEEN 400000000 AND 499999999\n"
        "         THEN v.VACREDITO - v.VADEBITO ELSE 0 END) AS TOTAL_VPA")
    partes.append(
        "    SUM(CASE WHEN v.INMES BETWEEN :mi AND :mf\n"
        "         AND v.COCONTACONTABIL BETWEEN 300000000 AND 399999999\n"
        "         THEN v.VADEBITO - v.VACREDITO ELSE 0 END) AS TOTAL_VPD")

    ini_enc, fim_enc = _faixa(CONTA_ENCERRAMENTO)
    partes.append(
        f"    SUM(CASE WHEN v.INMES BETWEEN 13 AND 15\n"
        f"         AND v.COCONTACONTABIL BETWEEN {ini_enc} AND {fim_enc}\n"
        f"         THEN v.VACREDITO - v.VADEBITO ELSE 0 END) AS RESULTADO_ENCERRAMENTO")
    partes.append(
        f"    SUM(CASE WHEN v.INMES BETWEEN :mi AND :mf\n"
        f"         AND v.COCONTACONTABIL BETWEEN {ini_enc} AND {fim_enc}\n"
        f"         THEN v.VACREDITO - v.VADEBITO ELSE 0 END) AS MOV_891_PERIODO")

    corpo = ",\n\n".join(partes)
    return f"""
SELECT
{corpo}
FROM {{schema}}.VSALDOCONTABIL v
WHERE (v.COCONTACONTABIL BETWEEN 300000000 AND 499999999
       OR v.COCONTACONTABIL BETWEEN 891000000 AND 891999999
       OR v.COCONTACONTABIL BETWEEN 951310000 AND 951339999
       OR v.COCONTACONTABIL BETWEEN 961310000 AND 961339999)
  {{filtro_ug}}
"""


SQL_DVP = montar_sql()


# ─────────────────────────────────────────────────────────────────────────────
#  CONEXAO  (idêntica ao BF/BO/BP/DFC/DMPL)
# ─────────────────────────────────────────────────────────────────────────────
def conectar_oracle():
    print(f"  Inicializando Oracle Client: {INSTANT_CLIENT_DIR}")
    try:
        oracledb.init_oracle_client(lib_dir=INSTANT_CLIENT_DIR)
    except Exception as e:
        if "already been initialized" not in str(e):
            raise
    dsn = f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}"
    conn = oracledb.connect(user=DB_USER, password=DB_PASSWORD, dsn=dsn)
    print(f"  Conectado! Oracle {conn.version}")
    return conn


def _buscar(conn, schema, mi, mf, coug):
    filtro = f"AND v.COUG = {coug}" if coug else ""
    sql = SQL_DVP.format(schema=schema, filtro_ug=filtro)
    sql = sql.replace(':mi', str(mi)).replace(':mf', str(mf))
    cur = conn.cursor()
    cur.execute(sql)
    cols = [d[0].upper() for d in cur.description]
    row = cur.fetchone()
    cur.close()
    return {k: float(v or 0) for k, v in zip(cols, row)}


def buscar_dados(conn, mes, ano, coug):
    """Exercício Atual: MIL{ano}, INMES 1..mes.
    Exercício Anterior: MIL{ano-1}, INMES 1..13 (ano fechado completo --
    mesma regra validada em bp.py)."""
    print(f"  [DVP] MIL{ano} (Exercício Atual, INMES 1..{mes})...")
    atual = _buscar(conn, f"MIL{ano}", 1, mes, coug)
    print(f"  [DVP] MIL{ano-1} (Exercício Anterior, INMES 1..13)...")
    try:
        anterior = _buscar(conn, f"MIL{ano-1}", 1, 13, coug)
    except Exception as e:
        print(f"    AVISO: não foi possível ler o exercício anterior: {e}")
        anterior = {}
    return atual, anterior


# ─────────────────────────────────────────────────────────────────────────────
#  TOTAIS DERIVADOS
# ─────────────────────────────────────────────────────────────────────────────
def calcular(brutos):
    t = dict(brutos)
    # TOTAL_VPA/TOTAL_VPD já vêm prontos do SQL (faixa direta da classe
    # inteira) -- não recalcular como soma das categorias nomeadas, por
    # causa do grupo 38 sem categoria própria (ver docstring do módulo).
    t.setdefault("TOTAL_VPA", 0.0)
    t.setdefault("TOTAL_VPD", 0.0)
    t["SOMA_CATEGORIAS_VPA"] = sum(t.get(chave, 0.0) for chave, _, _ in CATEGORIAS_VPA)
    t["SOMA_CATEGORIAS_VPD"] = sum(t.get(chave, 0.0) for chave, _, _ in CATEGORIAS_VPD)
    t["RESULTADO_PATRIMONIAL"] = t["TOTAL_VPA"] - t["TOTAL_VPD"]
    t["TOTAL_QUALITATIVAS"] = sum(t.get(chave, 0.0)
                                  for chave, _, _, _ in QUALITATIVAS)
    return t


# ─────────────────────────────────────────────────────────────────────────────
#  AUDITORIA DE INTEGRIDADE
# ─────────────────────────────────────────────────────────────────────────────
def auditoria_integridade(t, mes):
    achados = []

    for nome, chaves, total_chave, soma_chave in [
        ("VPA", CATEGORIAS_VPA, "TOTAL_VPA", "SOMA_CATEGORIAS_VPA"),
        ("VPD", CATEGORIAS_VPD, "TOTAL_VPD", "SOMA_CATEGORIAS_VPD"),
    ]:
        soma = t.get(soma_chave, 0.0)
        total = t.get(total_chave, 0.0)
        dif = total - soma
        if abs(dif) < 1.00:
            achados.append(('OK', f'Soma das categorias = TOTAL {nome}',
                            f'Soma das {len(chaves)} categorias {soma:,.2f}  =  '
                            f'TOTAL {nome} {total:,.2f}'))
        else:
            # Diferença esperada = grupo 38 (Custo de Mercadorias/Produtos/
            # Serviços), item sem categoria própria na Lista de Equações,
            # mas presente na faixa da classe -- o próprio PDF oficial tem
            # o mesmo gap (ver docstring do módulo). Não é erro de dado.
            achados.append(('INFO', f'Soma das categorias nomeadas < TOTAL {nome}',
                            f'Diferença: {dif:,.2f}  |  soma das {len(chaves)} '
                            f'categorias nomeadas {soma:,.2f}  vs  TOTAL {nome} '
                            f'(faixa direta da classe) {total:,.2f}  — esperado: '
                            f'é o grupo 38 (Custo de Mercadorias/Produtos/Serviços), '
                            f'que não tem categoria própria na Lista de Equações. '
                            f'O PDF oficial tem o mesmo gap.'))

    for nome, chave in [("VPA Total", "TOTAL_VPA"), ("VPD Total", "TOTAL_VPD")]:
        v = t.get(chave, 0.0)
        if v < 0:
            achados.append(('ALERTA', f'{nome} negativo — verificar', f'{v:,.2f}'))
        else:
            achados.append(('OK', f'{nome} positivo', f'{v:,.2f}'))

    # Resultado Patrimonial x conta de encerramento (891) -- mesma lógica
    # validada em controles/dvp.py: só é comparável em INMES 13..15.
    resultado = t.get("RESULTADO_PATRIMONIAL", 0.0)
    enc = t.get("RESULTADO_ENCERRAMENTO", 0.0)
    if abs(enc) < 1.00:
        mov = t.get("MOV_891_PERIODO", 0.0)
        extra = ("" if abs(mov) < 1.00 else
                 f"  (há {mov:,.2f} de movimentação avulsa na 891 no período)")
        achados.append(('INFO', 'Resultado Patrimonial (exercício ainda não encerrado)',
                        f'VPA − VPD = {resultado:,.2f}  — a 891XXXXXX só recebe o '
                        f'encerramento em INMES 13..15{extra}'))
    else:
        dif = resultado - enc
        if abs(dif) < 1.00:
            achados.append(('OK', 'Resultado Patrimonial = conta de encerramento',
                            f'{resultado:,.2f}  =  891XXXXXX {enc:,.2f}'))
        else:
            achados.append(('ERRO', 'Resultado Patrimonial ≠ conta de encerramento',
                            f'Diferença: {dif:,.2f}  |  {resultado:,.2f}  vs  '
                            f'891XXXXXX {enc:,.2f}'))

    return achados


def imprimir_auditoria(achados):
    icones = {'OK': '✔', 'ERRO': '✘', 'ALERTA': '⚠', 'INFO': 'ℹ'}
    print(f"\n{'─'*60}")
    print("  AUDITORIA DE INTEGRIDADE — DVP")
    print(f"{'─'*60}")
    for status, titulo, detalhe in achados:
        ic = icones.get(status, '?')
        print(f"  {ic} [{status:<6}] {titulo}")
        print(f"            {detalhe}")
    n_erro = sum(1 for s, _, _ in achados if s == 'ERRO')
    n_alerta = sum(1 for s, _, _ in achados if s == 'ALERTA')
    print(f"{'─'*60}")
    if n_erro:
        print(f"  RESULTADO: {n_erro} pendencia(s) de ERRO sinalizada(s) - ver detalhes acima.")
    elif n_alerta:
        print(f"  RESULTADO: Sem erros, mas {n_alerta} alerta(s) - revisar.")
    else:
        print(f"  RESULTADO: TODOS OS CONTROLES PASSARAM.")
    print(f"{'─'*60}")


# ─────────────────────────────────────────────────────────────────────────────
#  EXCEL
# ─────────────────────────────────────────────────────────────────────────────
FMT_BRL = '#,##0.00'
def _fill(h): return PatternFill('solid', fgColor=h)
def _side(s='thin'): return Side(border_style=s, color='B8CCE4')
B_THIN = Border(left=_side(), right=_side(), top=_side(), bottom=_side())
F_HDR  = _fill('BDD7EE')
F_GRAY = _fill('DCE6F1'); F_WHIT = _fill('F5F9FD')


def cel(ws, r, c, v='', bold=False, sz=9, bg=None, ha='left', brd=None,
        fmt=None, wrap=False):
    x = ws.cell(row=r, column=c, value=v)
    x.font = Font(name='Arial', bold=bold, size=sz, color='000000')
    x.alignment = Alignment(horizontal=ha, vertical='center', wrap_text=wrap)
    if bg: x.fill = bg
    if brd: x.border = brd
    if fmt and v not in ('', None): x.number_format = fmt
    return x


def gerar_excel(t_atual, t_ant, mes, ano, ug_label, output_path, achados=None):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = 'DVP'
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = 9

    for col, w in [('A', 42), ('B', 20), ('C', 20)]:
        ws.column_dimensions[col].width = w

    ws.merge_cells('A1:C3')
    c = ws['A1']
    c.value = ('GOVERNO DO DISTRITO FEDERAL\n'
               'Demonstração das Variações Patrimoniais\nVersão : 1')
    c.font = Font(name='Arial', bold=True, size=12, color='2E5C8A')
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 45

    ws.merge_cells('A4:C4')
    cel(ws, 4, 1, f'Mês de Referência   {mes:02d} - {MESES[mes]}          {ug_label}',
        bold=True, sz=9)

    r = 6
    cel(ws, r, 1, 'VARIAÇÕES PATRIMONIAIS AUMENTATIVAS', bold=True, bg=F_HDR, brd=B_THIN)
    cel(ws, r, 2, 'Exercício Atual', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    cel(ws, r, 3, 'Exercício Anterior', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    r += 1
    for chave, desc, _ in CATEGORIAS_VPA:
        cel(ws, r, 1, desc, brd=B_THIN, bg=F_WHIT)
        cel(ws, r, 2, t_atual.get(chave, 0.0), brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_WHIT)
        cel(ws, r, 3, t_ant.get(chave, 0.0), brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_WHIT)
        r += 1
    cel(ws, r, 1, 'TOTAL DAS VARIAÇÕES PATRIMONIAIS AUMENTATIVAS (I)', bold=True, brd=B_THIN, bg=F_GRAY)
    cel(ws, r, 2, t_atual.get('TOTAL_VPA', 0.0), bold=True, brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_GRAY)
    cel(ws, r, 3, t_ant.get('TOTAL_VPA', 0.0), bold=True, brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_GRAY)
    r += 2

    cel(ws, r, 1, 'VARIAÇÕES PATRIMONIAIS DIMINUTIVAS E RESULT. PATR.', bold=True, bg=F_HDR, brd=B_THIN)
    cel(ws, r, 2, 'Exercício Atual', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    cel(ws, r, 3, 'Exercício Anterior', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    r += 1
    for chave, desc, _ in CATEGORIAS_VPD:
        cel(ws, r, 1, desc, brd=B_THIN, bg=F_WHIT)
        cel(ws, r, 2, t_atual.get(chave, 0.0), brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_WHIT)
        cel(ws, r, 3, t_ant.get(chave, 0.0), brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_WHIT)
        r += 1
    cel(ws, r, 1, 'TOTAL DAS VARIAÇÕES PATRIMONIAIS DIMINUTIVAS (II)', bold=True, brd=B_THIN, bg=F_GRAY)
    cel(ws, r, 2, t_atual.get('TOTAL_VPD', 0.0), bold=True, brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_GRAY)
    cel(ws, r, 3, t_ant.get('TOTAL_VPD', 0.0), bold=True, brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_GRAY)
    r += 1
    cel(ws, r, 1, 'RESULTADO PATRIMONIAL DO PERÍODO (III) = (I - II)', bold=True, brd=B_THIN, bg=F_HDR)
    cel(ws, r, 2, t_atual.get('RESULTADO_PATRIMONIAL', 0.0), bold=True, brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_HDR)
    cel(ws, r, 3, t_ant.get('RESULTADO_PATRIMONIAL', 0.0), bold=True, brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_HDR)
    r += 3

    cel(ws, r, 1, 'VARIAÇÕES PATRIMONIAIS QUALITATIVAS', bold=True, bg=F_HDR, brd=B_THIN)
    cel(ws, r, 2, 'Exercício Atual', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    cel(ws, r, 3, 'Exercício Anterior', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    r += 1
    for chave, desc, _, _ in QUALITATIVAS:
        cel(ws, r, 1, desc, brd=B_THIN, bg=F_WHIT)
        cel(ws, r, 2, t_atual.get(chave, 0.0), brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_WHIT)
        cel(ws, r, 3, t_ant.get(chave, 0.0), brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_WHIT)
        r += 1
    cel(ws, r, 1, 'TOTAL', bold=True, brd=B_THIN, bg=F_GRAY)
    cel(ws, r, 2, t_atual.get('TOTAL_QUALITATIVAS', 0.0), bold=True, brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_GRAY)
    cel(ws, r, 3, t_ant.get('TOTAL_QUALITATIVAS', 0.0), bold=True, brd=B_THIN, ha='right', fmt=FMT_BRL, bg=F_GRAY)
    r += 3

    cel(ws, r, 1, f'Emitido em: {datetime.now():%d/%m/%Y %H:%M:%S}', sz=8)

    if achados:
        ws3 = wb.create_sheet('Auditoria Integridade')
        ws3.column_dimensions['A'].width = 10
        ws3.column_dimensions['B'].width = 45
        ws3.column_dimensions['C'].width = 90
        hdr_bg = _fill('2E5C8A')
        cel(ws3, 1, 1, 'Status', bold=True, bg=hdr_bg, ha='center', brd=B_THIN)
        ws3['A1'].font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        cel(ws3, 1, 2, 'Controle', bold=True, bg=hdr_bg, ha='center', brd=B_THIN)
        ws3['B1'].font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        cel(ws3, 1, 3, 'Detalhe', bold=True, bg=hdr_bg, ha='center', brd=B_THIN)
        ws3['C1'].font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        cores = {'ERRO': _fill('FAD7DA'), 'ALERTA': _fill('FCF0CE'),
                 'INFO': _fill('DDEBF7')}
        for i, (status, titulo, detalhe) in enumerate(achados, start=2):
            bg = cores.get(status, _fill('D9F0DD'))
            cel(ws3, i, 1, status, bold=True, bg=bg, brd=B_THIN, ha='center')
            cel(ws3, i, 2, titulo, bg=bg, brd=B_THIN, ha='left', wrap=True)
            cel(ws3, i, 3, detalhe, bg=bg, brd=B_THIN, ha='left', wrap=True)

    wb.save(output_path)
    print(f"  Excel salvo: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  PDF
# ─────────────────────────────────────────────────────────────────────────────
def _brl(v):
    return f'{v:,.2f}'.replace(',', '§').replace('.', ',').replace('§', '.')


def gerar_pdf(t_atual, t_ant, mes, ano, ug_label, output_path, achados=None):
    if not REPORTLAB_OK:
        print("  AVISO: reportlab nao instalado -- PDF nao gerado "
              "(pip install reportlab).")
        return

    doc = SimpleDocTemplate(str(output_path), pagesize=landscape(A4),
                             leftMargin=10*mm, rightMargin=10*mm,
                             topMargin=10*mm, bottomMargin=10*mm)
    styles = getSampleStyleSheet()
    elems = []

    st_t = ParagraphStyle('t', parent=styles['Normal'], fontName='Helvetica-Bold',
                           fontSize=12, textColor=colors.HexColor('#2E5C8A'), leading=14)
    st_s = ParagraphStyle('s', parent=styles['Normal'], fontSize=8, alignment=TA_RIGHT)
    st_r = ParagraphStyle('r', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9)
    st_hdr = ParagraphStyle('hdr', fontName='Helvetica-Bold', fontSize=8,
                             alignment=TA_CENTER, leading=9)
    st_desc = ParagraphStyle('d', fontName='Helvetica', fontSize=8, leading=9)
    st_desc_b = ParagraphStyle('db', fontName='Helvetica-Bold', fontSize=8, leading=9)
    st_val = ParagraphStyle('v', fontName='Helvetica', fontSize=8, alignment=TA_RIGHT, leading=9)
    st_val_b = ParagraphStyle('vb', fontName='Helvetica-Bold', fontSize=8, alignment=TA_RIGHT, leading=9)

    cab = Table([[Paragraph('GOVERNO DO DISTRITO FEDERAL<br/>'
                             'Demonstração das Variações Patrimoniais<br/>Versão 1', st_t),
                  Paragraph(f'Exercício {ano}<br/>PSIAG550<br/>'
                            f'Posição em: {datetime.now():%d/%m/%Y às %H:%M:%S}', st_s)]],
                colWidths=[200*mm, 70*mm])
    cab.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    elems.append(cab)
    elems.append(Paragraph(f'Mês de Referência: {mes:02d} - {MESES[mes]} &nbsp; {ug_label}', st_r))
    elems.append(Spacer(1, 4*mm))

    def bloco(titulo, categorias, total_chave, total_label):
        dados = [[Paragraph(titulo, st_hdr), Paragraph('Exercício Atual', st_hdr),
                  Paragraph('Exercício Anterior', st_hdr)]]
        for chave, desc, _ in categorias:
            dados.append([Paragraph(desc, st_desc),
                          Paragraph(_brl(t_atual.get(chave, 0.0)), st_val),
                          Paragraph(_brl(t_ant.get(chave, 0.0)), st_val)])
        dados.append([Paragraph(total_label, st_desc_b),
                      Paragraph(_brl(t_atual.get(total_chave, 0.0)), st_val_b),
                      Paragraph(_brl(t_ant.get(total_chave, 0.0)), st_val_b)])
        tab = Table(dados, colWidths=[150*mm, 60*mm, 60*mm], repeatRows=1)
        ts = [('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
              ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
              ('TOPPADDING', (0, 0), (-1, -1), 3),
              ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
              ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#BDD7EE')),
              ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#DCE6F1'))]
        tab.setStyle(TableStyle(ts))
        return tab

    elems.append(bloco('VARIAÇÕES PATRIMONIAIS AUMENTATIVAS', CATEGORIAS_VPA,
                        'TOTAL_VPA', 'TOTAL DAS VARIAÇÕES PATRIMONIAIS AUMENTATIVAS (I)'))
    elems.append(Spacer(1, 4*mm))
    elems.append(bloco('VARIAÇÕES PATRIMONIAIS DIMINUTIVAS', CATEGORIAS_VPD,
                        'TOTAL_VPD', 'TOTAL DAS VARIAÇÕES PATRIMONIAIS DIMINUTIVAS (II)'))
    elems.append(Spacer(1, 2*mm))

    res_tab = Table([[Paragraph('RESULTADO PATRIMONIAL DO PERÍODO (III) = (I - II)', st_desc_b),
                      Paragraph(_brl(t_atual.get('RESULTADO_PATRIMONIAL', 0.0)), st_val_b),
                      Paragraph(_brl(t_ant.get('RESULTADO_PATRIMONIAL', 0.0)), st_val_b)]],
                     colWidths=[150*mm, 60*mm, 60*mm])
    res_tab.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
                                  ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#BDD7EE')),
                                  ('TOPPADDING', (0, 0), (-1, -1), 4),
                                  ('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))
    elems.append(res_tab)
    elems.append(Spacer(1, 4*mm))

    dados_q = [[Paragraph('VARIAÇÕES PATRIMONIAIS QUALITATIVAS', st_hdr),
                Paragraph('Exercício Atual', st_hdr), Paragraph('Exercício Anterior', st_hdr)]]
    for chave, desc, _, _ in QUALITATIVAS:
        dados_q.append([Paragraph(desc, st_desc),
                        Paragraph(_brl(t_atual.get(chave, 0.0)), st_val),
                        Paragraph(_brl(t_ant.get(chave, 0.0)), st_val)])
    dados_q.append([Paragraph('TOTAL', st_desc_b),
                    Paragraph(_brl(t_atual.get('TOTAL_QUALITATIVAS', 0.0)), st_val_b),
                    Paragraph(_brl(t_ant.get('TOTAL_QUALITATIVAS', 0.0)), st_val_b)])
    tab_q = Table(dados_q, colWidths=[150*mm, 60*mm, 60*mm], repeatRows=1)
    tab_q.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
                                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                ('TOPPADDING', (0, 0), (-1, -1), 3),
                                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#BDD7EE')),
                                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#DCE6F1'))]))
    elems.append(tab_q)

    if achados:
        elems.append(Spacer(1, 5*mm))
        st_aud_tit = ParagraphStyle('audt', parent=styles['Normal'],
                       fontName='Helvetica-Bold', fontSize=10,
                       textColor=colors.HexColor('#2E5C8A'))
        elems.append(Paragraph('AUDITORIA DE INTEGRIDADE', st_aud_tit))
        elems.append(Spacer(1, 1*mm))
        st_cell_b = ParagraphStyle('cb', parent=styles['Normal'],
                      fontName='Helvetica-Bold', fontSize=7.5)
        st_cell_n = ParagraphStyle('cn', parent=styles['Normal'], fontSize=7.5)
        aud_rows = [['Status', 'Controle', 'Detalhe']]
        for status, titulo, detalhe in achados:
            aud_rows.append([Paragraph(status, st_cell_b), Paragraph(titulo, st_cell_b),
                             Paragraph(detalhe, st_cell_n)])
        aud_tab = Table(aud_rows, colWidths=[20*mm, 75*mm, 175*mm], repeatRows=1)
        aud_ts = [
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E5C8A')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        for i, (status, _, _) in enumerate(achados, start=1):
            cor = {'ERRO': '#FAD7DA', 'ALERTA': '#FCF0CE',
                   'INFO': '#E8F1FA'}.get(status, '#D9F0DD')
            aud_ts.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor(cor)))
            aud_ts.append(('ALIGN', (0, i), (0, i), 'CENTER'))
        aud_tab.setStyle(TableStyle(aud_ts))
        elems.append(aud_tab)

        n_erro = sum(1 for s, _, _ in achados if s == 'ERRO')
        n_alerta = sum(1 for s, _, _ in achados if s == 'ALERTA')
        cor_res = '#B00000' if n_erro else ('#9C6500' if n_alerta else '#1F5C2E')
        st_res = ParagraphStyle('res', parent=styles['Normal'],
                   fontName='Helvetica-Bold', fontSize=8.5,
                   textColor=colors.HexColor(cor_res))
        elems.append(Spacer(1, 2*mm))
        if n_erro:
            elems.append(Paragraph(
                f'RESULTADO: {n_erro} pendência(s) de ERRO sinalizada(s) - ver detalhes acima.', st_res))
        elif n_alerta:
            elems.append(Paragraph(f'RESULTADO: Sem erros, mas {n_alerta} alerta(s) - revisar.', st_res))
        else:
            elems.append(Paragraph('RESULTADO: TODOS OS CONTROLES PASSARAM.', st_res))

    doc.build(elems)
    print(f"  PDF salvo:   {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(
        description='Demonstração das Variações Patrimoniais - GDF (DVP)')
    p.add_argument('--mes', type=int, default=5, help='Mês de referência 1-12')
    p.add_argument('--ano', type=int, default=2026)
    p.add_argument('--ug', type=str, default=None,
                   help='Código da Unidade Gestora (omitir = Consolidado)')
    p.add_argument('--saida', type=str, default=None,
                   help='Caminho base dos arquivos de saída (sem extensão)')
    p.add_argument('--formato', choices=['ambos', 'excel', 'pdf'], default='ambos')
    a = p.parse_args()

    if not 1 <= a.mes <= 12:
        print("ERRO: --mes deve estar entre 1 e 12."); sys.exit(1)

    ug_label = f'UG: {a.ug}' if a.ug else 'Consolidado'
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome = (f'DVP_GDF_{a.ano}_{a.mes:02d}'
            + (f'_UG{a.ug}' if a.ug else '_Consolidado')
            + f'_{timestamp}')
    base = (Path(a.saida).with_suffix('') if a.saida else OUTPUT_DIR / nome)
    saida_xlsx = base.with_suffix('.xlsx')
    saida_pdf = base.with_suffix('.pdf')

    print(f"\n{'='*60}")
    print(f"  DEMONSTRAÇÃO DAS VARIAÇÕES PATRIMONIAIS — DVP")
    print(f"  {MESES[a.mes]}/{a.ano}  |  {ug_label}")
    print(f"{'='*60}")

    print("\n[1/4] Conectando ao Oracle...")
    conn = conectar_oracle()

    print("\n[2/4] Buscando dados (Exercício Atual e Anterior)...")
    brutos_atual, brutos_ant = buscar_dados(conn, a.mes, a.ano, a.ug)
    conn.close()

    print("\n[3/4] Calculando VPA, VPD e Resultado Patrimonial...")
    t_atual = calcular(brutos_atual)
    t_ant = calcular(brutos_ant) if brutos_ant else calcular({})
    achados = auditoria_integridade(t_atual, a.mes)

    print("\n[4/4] Gerando arquivos...")
    if a.formato in ('ambos', 'excel'):
        gerar_excel(t_atual, t_ant, a.mes, a.ano, ug_label, saida_xlsx, achados=achados)
    if a.formato in ('ambos', 'pdf'):
        gerar_pdf(t_atual, t_ant, a.mes, a.ano, ug_label, saida_pdf, achados=achados)

    print(f"\n{'─'*60}")
    print("  CONFERÊNCIA RÁPIDA:")
    print(f"{'─'*60}")
    print(f"    VPA TOTAL (I)              : {t_atual['TOTAL_VPA']:>20,.2f}")
    print(f"    VPD TOTAL (II)             : {t_atual['TOTAL_VPD']:>20,.2f}")
    print(f"    RESULTADO PATRIMONIAL (III): {t_atual['RESULTADO_PATRIMONIAL']:>20,.2f}")
    print(f"    QUALITATIVAS (TOTAL)       : {t_atual['TOTAL_QUALITATIVAS']:>20,.2f}")
    print(f"{'─'*60}")

    imprimir_auditoria(achados)

    print(f"\n  Concluído em {datetime.now():%H:%M:%S}\n")


if __name__ == '__main__':
    main()
