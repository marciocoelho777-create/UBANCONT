#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera Controles_Rotina_{ano}_{mes:02d}_{ts}.xlsx
com uma aba por controle de rotina da pasta rotina/.

Apresentação similar aos relatórios Oracle Discoverer originais:
  - título + período no topo
  - coluna DIFERENÇA realçada em roxo
  - linhas com diferença ≠ 0 em vermelho claro
  - linhas sem diferença em verde claro
  - painéis fixos nas quatro colunas de UG
  - auto-filtro no cabeçalho
  - totais no rodapé

Uso:
    python rotina\\gerar_controles_rotina.py --mes 8 --ano 2026
    python rotina\\gerar_controles_rotina.py --mes 8 --ano 2026 --controle 02
"""
from __future__ import annotations
import argparse
import sys
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import oracledb
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── Localização ───────────────────────────────────────────────────────────────
ROTINA_DIR = Path(__file__).parent
OUTPUT_DIR  = ROTINA_DIR.parent

# ── Credenciais (sobreposta por config_local.py) ──────────────────────────────
DB_USER = DB_PASSWORD = DB_HOST = DB_SERVICE = INSTANT_CLIENT_DIR = ""
DB_PORT = "1521"

try:
    import importlib.util as _ilu
    _cfg = OUTPUT_DIR / "config_local.py"
    if _cfg.exists():
        _spec = _ilu.spec_from_file_location("config_local", _cfg)
        _mod  = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_mod)
        for _k in ("DB_USER", "DB_PASSWORD", "DB_HOST", "DB_PORT",
                   "DB_SERVICE", "INSTANT_CLIENT_DIR"):
            if hasattr(_mod, _k):
                globals()[_k] = getattr(_mod, _k)
except Exception:
    pass


# ── Definição dos controles ───────────────────────────────────────────────────
_UG = {0: "Gestão", 1: "Gestão Contáb.", 2: "UG Contáb.", 3: "UG"}

CONTROLES = [
    {
        "numero":    "01",
        "titulo":    "Controle 1 — Contas 21 F Principal do exercício × 827110101",
        "aba":       "C01",
        "sql":       "01-Controle 1 - Contas 21 F Principal do exercício x 827110101 LANÇAMENTO.sql",
        "tipo":      "ug",
        "diff_cols": [5],        # C_84
        "col_labels": {
            **_UG,
            4:   "2152301XX",   # C_10
            5:   "DIFERENÇA",   # C_84
            6:   "218910300",   # C_94
            7:   "2111101XX",   # C_12
            8:   "2111104XX",   # C_42
            9:   "2111105XX",   # C_50
            10:  "2111106XX",   # C_82
            11:  "2111107XX",   # C_83
            12:  "2112101XX",   # C_11
            13:  "2112104XX",   # C_13
            14:  "2112105XX",   # C_14
            15:  "2112106XX",   # C_15
            16:  "2112107XX",   # C_16
            17:  "2112305XX",   # C_17
            18:  "2113101XX",   # C_31
            19:  "2113103XX",   # C_32
            20:  "2113104XX",   # C_33
            21:  "2113105XX",   # C_34
            22:  "2114101XX",   # C_35
            23:  "2114105XX",   # C_36
            24:  "2114106XX",   # C_37
            25:  "2114107XX",   # C_38
            26:  "2114108XX",   # C_39
            27:  "2114109XX",   # C_40
            28:  "2114112XX",   # C_41
            29:  "2114201XX",   # C_43
            30:  "2114202XX",   # C_44
            31:  "2114203XX",   # C_45
            32:  "2114301XX",   # C_46
            33:  "2114303XX",   # C_47
            34:  "2114305XX",   # C_48
            35:  "2114306XX",   # C_49
            36:  "2114307XX",   # C_51
            37:  "2114308XX",   # C_52
            38:  "2121102XX",   # C_85
            39:  "2121302XX",   # C_86
            40:  "2121304XX",   # C_87
            41:  "2122102XX",   # C_88
            42:  "2122103XX",   # C_89
            43:  "2122302XX",   # C_90
            44:  "2122303XX",   # C_91
            45:  "2123101XX",   # C_92
            46:  "2123301XX",   # C_93
            47:  "2124101XX",   # C_95
            48:  "2124102XX",   # C_96
            49:  "2124301XX",   # C_97
            50:  "2124302XX",   # C_98
            51:  "21251XXXX",   # C_99
            52:  "21253XXXX",   # C_100
            53:  "21261XXXX",   # C_101
            54:  "21263XXXX",   # C_102
            55:  "2131101XX",   # C_103
            56:  "2131103XX",   # C_104
            57:  "2131105XX",   # C_105
            58:  "2131106XX",   # C_106
            59:  "2131107XX",   # C_107
            60:  "2131108XX",   # C_108
            61:  "2131109XX",   # C_109
            62:  "2131110XX",   # C_110
            63:  "2131201XX",   # C_111
            64:  "2131203XX",   # C_112
            65:  "2131205XX",   # C_113
            66:  "2132101XX",   # C_114
            67:  "2132102XX",   # C_115
            68:  "2141301XX",   # C_116
            69:  "2141302XX",   # C_117
            70:  "2141305XX",   # C_118
            71:  "2141309XX",   # C_119
            72:  "2141310XX",   # C_120
            73:  "2141311XX",   # C_121
            74:  "2141312XX",   # C_122
            75:  "2141399XX",   # C_123
            76:  "2142201XX",   # C_1
            77:  "2142202XX",   # C_2
            78:  "2142203XX",   # C_3
            79:  "2142206XX",   # C_4
            80:  "2142299XX",   # C_5
            81:  "2143201XX",   # C_6
            82:  "2143202XX",   # C_7
            83:  "2143502XX",   # C_8
            84:  "2152101XX",   # C_9
            85:  "2153101XX",   # C_18
            86:  "2153102XX",   # C_19
            87:  "2153103XX",   # C_20
            88:  "2153104XX",   # C_21
            89:  "2153105XX",   # C_22
            90:  "2153106XX",   # C_23
            91:  "2153199XX",   # C_24
            92:  "2153301XX",   # C_25
            93:  "2153303XX",   # C_26
            94:  "2153304XX",   # C_27
            95:  "2153305XX",   # C_28
            96:  "2153306XX",   # C_29
            97:  "2153399XX",   # C_30
            98:  "218910101",   # C_53
            99:  "218910102",   # C_54
            100: "218910104",   # C_55
            101: "218910199",   # C_56
            102: "218910200",   # C_57
            103: "218910500",   # C_58
            104: "218910700",   # C_59
            105: "218910800",   # C_60
            106: "218910900",   # C_61
            107: "218911000",   # C_62
            108: "218911100",   # C_63
            109: "218911200",   # C_64
            110: "218911400",   # C_65
            111: "218911800",   # C_66
            112: "218911900",   # C_67
            113: "218912300",   # C_68
            114: "218913000",   # C_69
            115: "218913301",   # C_70
            116: "218913401",   # C_71
            117: "218913402",   # C_72
            118: "218913403",   # C_73
            119: "218920102",   # C_75
            120: "218920400",   # C_76
            121: "218920500",   # C_77
            122: "218920700",   # C_78
            123: "218930500",   # C_80
            124: "827110101",   # C_81
            125: "218919600",   # C_74
            126: "218929600",   # C_79
        },
    },
    {
        "numero":    "02",
        "titulo":    "Controle 2 — Contas 21 F Ret do exercício × 827110201/827110203/63181",
        "aba":       "C02",
        "sql":       "02-Controle 2 - Contas 21 F Ret do exercício x 827110201,827110203,63181 LANÇAMENTO.sql",
        "tipo":      "ug",
        "diff_cols": [4],
        "col_labels": {**_UG, 4: "DIFERENÇA", 5: "2188XXXXX",
                       6: "827110203", 7: "827110201", 8: "631810000"},
    },
    {
        "numero":    "03",
        "titulo":    "Controle 3 — 21xxx98xx = 63211xxxx",
        "aba":       "C03",
        "sql":       "03-Controle 3 - 21xxx98xx = 63211xxxx LANÇAMENTO.sql",
        "tipo":      "ug",
        "diff_cols": [4, 7],
        "col_labels": {**_UG,
                       4: "DIFERENÇA (827110301+303 − 632110300)",
                       5: "827110303", 6: "827110301",
                       7: "DIFERENÇA (632110100+300 − 21xxx)",
                       8: "Σ Contas 21", 9: "632110300", 10: "632110100"},
    },
    {
        "numero":    "04",
        "titulo":    "Controle 4 — Contas 21 sem NE × 827110401",
        "aba":       "C04",
        "sql":       "04-Controle 4 - Contas 21 sem NE x 827110401 LANÇAMENTO.sql",
        "tipo":      "ug",
        "diff_cols": [4],
        "col_labels": {**_UG, 4: "DIFERENÇA", 5: "Σ Contas 21", 6: "827110401"},
    },
    {
        "numero":    "05",
        "titulo":    "Controle 5 — Contas 21 em liquidação × 827110196",
        "aba":       "C05",
        "sql":       "05-Controle 5 - Contas 21 em liquidação  x 827110196 LANÇAMENTO.sql",
        "tipo":      "ug",
        "diff_cols": [4],
        "col_labels": {**_UG, 4: "DIFERENÇA",
                       5: "827110196", 6: "218929600",
                       7: "218919600", 8: "218919896",
                       9: "218929896", 10: "631200000"},
    },
    {
        "numero":    "06",
        "titulo":    "Controle 6 — Contas 21 RPNP liquidado × 6313",
        "aba":       "C06",
        "sql":       "06-Controle 6 - Contas 21 RPNP liquidado x 6313 LANÇAMENTO.sql",
        "tipo":      "ug",
        "diff_cols": [4],
        "col_labels": {**_UG, 4: "DIFERENÇA",
                       5: "631300000", 6: "218924002", 7: "218914002"},
    },
    {
        "numero":    "08",
        "titulo":    "Controle 8 — Balancete INTRA",
        "aba":       "C08",
        "sql":       "08-Controle 8 - Balancete INTRA.sql",
        "tipo":      "ug",
        "diff_cols": [4],
        "col_labels": {**_UG, 4: "DIFERENÇA (CL1+2+3+4)",
                       5: "Cl.4 VPA", 6: "Cl.3 VPD",
                       7: "Cl.2 Passivo/PL", 8: "Cl.1 Ativo"},
    },
    {
        "numero":    "09",
        "titulo":    "Controle 9 — Balancete Não INTRA",
        "aba":       "C09",
        "sql":       "09-Controle 9 - Balancete Não INTRA.sql",
        "tipo":      "ug",
        "diff_cols": [4],
        "col_labels": {**_UG, 4: "DIFERENÇA (CL1+2+3+4)",
                       5: "Cl.4 VPA", 6: "Cl.3 VPD",
                       7: "Cl.2 Passivo/PL", 8: "Cl.1 Ativo"},
    },
    {
        "numero":    "10",
        "titulo":    "Controle 10 — Balanço Financeiro",
        "aba":       "C10",
        "sql":       "10-Balanço-BF.sql",
        "tipo":      "ug",
        "diff_cols": [4],
        "col_labels": {**_UG,
                       4:  "DIFERENÇA (BF)",  # C_21
                       5:  "1138117XX",        # C_20
                       6:  "1138106XX",        # C_19
                       7:  "1141XXXXX",        # C_18
                       8:  "111XXXXXX",        # C_17
                       9:  "3513XXXXX (D)",    # C_16
                       10: "237110301",        # C_15
                       11: "2188104XX",        # C_14
                       12: "6322XXXXX",        # C_13
                       13: "631400000",        # C_12
                       14: "3513XXXXX (C)",    # C_11
                       15: "351220101",        # C_10
                       16: "351120100",        # C_9
                       17: "6213XXXXX (D)",    # C_8
                       18: "2188150XX",        # C_7
                       19: "2188104XX",        # C_6
                       20: "532700000",        # C_5
                       21: "531700000",        # C_4
                       22: "4513XXXXX",        # C_3
                       23: "451220101",        # C_2
                       24: "451120100",        # C_22
                       25: "6213XXXXX (C)",    # C_1
                       },
    },
    {
        "numero":    "13",
        "titulo":    "Controle 13 — Balanço Patrimonial",
        "aba":       "C13",
        "sql":       "13-Balanço-BP.sql",
        "tipo":      "ug",
        "diff_cols": [4],
        "col_labels": {**_UG,
                       4:  "DIFERENÇA (BP)",  # C_27
                       5:  "119XXXXXX",        # C_32
                       6:  "215XXXXXX",        # C_10
                       7:  "237XXXXXX",        # C_26
                       8:  "4XXXXXXXX",        # C_25
                       9:  "235XXXXXX",        # C_23
                       10: "233XXXXXX",        # C_21
                       11: "121XXXXXX",        # C_2
                       12: "227XXXXXX",        # C_17
                       13: "217XXXXXX",        # C_11
                       14: "221XXXXXX",        # C_13
                       15: "211XXXXXX",        # C_6
                       16: "224XXXXXX",        # C_16
                       17: "214XXXXXX",        # C_9
                       18: "114XXXXXX",        # C_30
                       19: "122XXXXXX",        # C_3
                       20: "124XXXXXX",        # C_5
                       21: "123XXXXXX",        # C_4
                       22: "223XXXXXX",        # C_15
                       23: "213XXXXXX",        # C_8
                       24: "115XXXXXX",        # C_31
                       25: "222XXXXXX",        # C_14
                       26: "212XXXXXX",        # C_7
                       27: "236XXXXXX",        # C_24
                       28: "228XXXXXX",        # C_18
                       29: "218XXXXXX",        # C_12
                       30: "113XXXXXX",        # C_29
                       31: "112XXXXXX",        # C_28
                       32: "231XXXXXX",        # C_19
                       33: "111XXXXXX",        # C_1
                       34: "234XXXXXX",        # C_22
                       35: "232XXXXXX",        # C_20
                       },
    },
    {
        "numero":    "14",
        "titulo":    "Controle 14 — 72119XXXX (Par Devedor/Credor)",
        "aba":       "C14",
        "sql":       "14-72119XXXX.sql",
        "tipo":      "ug",
        "diff_cols": [4, 7, 10, 13],
        "col_labels": {**_UG,
                       4:  "DIFERENÇA 1", 5:  "821190100", 6:  "721190100",
                       7:  "DIFERENÇA 2", 8:  "821190200", 9:  "721190200",
                       10: "DIFERENÇA 3", 11: "821190300", 12: "721190300",
                       13: "DIFERENÇA 4", 14: "821190400", 15: "721190400"},
    },
    {
        "numero":    "15",
        "titulo":    "Controle 15 — Receita Negativa Saldo",
        "aba":       "C15",
        "sql":       "15-Receita Negativa Saldo.sql",
        "tipo":      "detalhe",  # só negativos retornados são ocorrências
        "diff_cols": [2],
        "col_labels": {0: "Conta Corrente (8 dígitos)",
                       1: "Conta Contábil",
                       2: "Saldo",
                       3: "UG"},
        # UG (3) vira seletor; UG+Conta Corrente+Conta Contábil ficam na linha
        "detalhe_row_id_cols": [3, 0, 1],
    },
    {
        "numero":    "16",
        "titulo":    "Controle 16 — 5221904XX",
        "aba":       "C16",
        "sql":       "16-5221904XX.sql",
        "tipo":      "ug",
        "diff_cols": [4],
        "col_labels": {**_UG, 4: "TOTAL (52219040X)",
                       5: "522190401", 6: "522190409"},
    },
    {
        "numero":    "17",
        "titulo":    "Controle 17 — Inversão de Saldo",
        "aba":       "C17",
        "sql":       "17-Inversão de Saldo.sql",
        "tipo":      "detalhe",  # só saldos invertidos (< 0) do mês corrente
        "diff_cols": [3],
        "col_labels": {0: "Ind. Inversão Saldo",
                       1: "Conta Contábil",
                       2: "Conta Corrente",
                       3: "Saldo",
                       4: "UG"},
        # UG (4) vira seletor; UG + Conta Contábil + Conta Corrente ficam na linha
        "detalhe_row_id_cols": [4, 1, 2],
    },
    {
        "numero":    "18",
        "titulo":    "Controle 18 — Previsão Adicional a Lançar",
        "aba":       "C18",
        "sql":       "18-Previsão Adicional a Lançar.sql",
        "tipo":      "detalhe",
        "diff_cols": [3],
        "col_labels": {0: "Nome Conta Contábil",
                       1: "Conta Contábil",
                       2: "Gestão",
                       3: "Saldo",
                       4: "UG"},
        # UG (4) e Gestão (2) ficam na linha como identificadores E viram seletores
        "detalhe_row_id_cols": [4, 2],
        # Drill-down: UG → Gestão → Conta (col 0=nome, col 1=código, col 3=saldo)
        "hierarquia": {
            "grupo_cols":      [4, 2],
            "grupo_labels":    ["UG", "Gestão"],
            "saldo_col":       3,
            "detalhe_cols":    [0, 1],
            "conta_filter_col": 1,   # Conta Contábil — vira seletor de filtro
        },
    },
]


# ── Estilos ───────────────────────────────────────────────────────────────────
def _lado(s="thin"):
    return Side(border_style=s, color="B8CCE4")

BORDA      = Border(left=_lado(), right=_lado(), top=_lado(), bottom=_lado())
HDR_FILL   = PatternFill("solid", fgColor="2E5C8A")   # azul escuro — cabeçalho geral
DIFF_FILL  = PatternFill("solid", fgColor="7030A0")   # roxo — coluna diferença
UG_FILL    = PatternFill("solid", fgColor="D9E8F5")   # azul claro — colunas UG
OK_FILL    = PatternFill("solid", fgColor="D9F0DD")   # verde — diferença ≈ 0
ERR_FILL   = PatternFill("solid", fgColor="FAD7DA")   # vermelho — diferença ≠ 0
TOT_FILL   = PatternFill("solid", fgColor="E2EFDA")   # verde pálido — totais
DET_FILL   = PatternFill("solid", fgColor="FFF2CC")   # amarelo — linhas detalhe
TOL        = 0.02                                      # tolerância de equilíbrio


def _cel(ws, r, c, val="", bold=False, fill=None, ha="left",
         fmt=None, wrap=False, size=9):
    cell = ws.cell(row=r, column=c, value=val)
    cor   = "FFFFFF" if fill in (HDR_FILL, DIFF_FILL) else "000000"
    cell.font      = Font(name="Arial", bold=bold, size=size, color=cor)
    cell.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
    cell.border    = BORDA
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    return cell


# ── Conexão ───────────────────────────────────────────────────────────────────
def conectar() -> oracledb.Connection:
    if INSTANT_CLIENT_DIR:
        try:
            oracledb.init_oracle_client(lib_dir=str(INSTANT_CLIENT_DIR))
        except Exception as e:
            if "already been initialized" not in str(e):
                raise
    faltando = [n for n, v in [("DB_USER", DB_USER), ("DB_PASSWORD", DB_PASSWORD),
                                ("DB_HOST", DB_HOST), ("DB_SERVICE", DB_SERVICE)] if not v]
    if faltando:
        sys.exit(f"ERRO: credenciais ausentes: {faltando}")
    conn = oracledb.connect(user=DB_USER, password=DB_PASSWORD,
                            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}")
    print(f"  Conectado! Oracle {conn.version}")
    return conn


# ── Execução SQL ──────────────────────────────────────────────────────────────
def executar(conn, sql_path: Path, ano: int, mes: int = 12) -> tuple[list[str], list[tuple]]:
    sql = sql_path.read_text(encoding="utf-8").strip().rstrip(";")
    sql = sql.replace(f"MIL{ano}", f"MIL{ano}")  # mantém flexível para outros anos
    sql = sql.replace("{MES}", str(mes))           # substitui mês dinâmico
    cur = conn.cursor()
    cur.execute(sql)
    cols = [d[0] for d in cur.description]
    rows = cur.fetchall()
    cur.close()
    return cols, rows


# ── Escrita de aba: tipo "ug" (agrupado por UG) ───────────────────────────────
def _escrever_aba_ug(ws, ctrl: dict, col_aliases: list[str],
                     rows: list[tuple], mes: int, ano: int) -> None:
    nc          = len(col_aliases)
    diff_set    = set(ctrl["diff_cols"])
    col_labels  = ctrl.get("col_labels", {})
    agora       = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    ws.sheet_view.showGridLines = False

    # ── Linha 1: título ───────────────────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(row=1, column=1, value=ctrl["titulo"])
    c.font      = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    c.fill      = HDR_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # ── Linha 2: período ──────────────────────────────────────────────────────
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nc)
    c2 = ws.cell(row=2, column=1,
                 value=f"Período: {mes:02d}/{ano}   |   Gerado em: {agora}"
                       f"   |   {len(rows)} UG(s) encontrada(s)")
    c2.font      = Font(name="Arial", size=9, italic=True, color="444444")
    c2.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 14

    # ── Linha 3: cabeçalho ────────────────────────────────────────────────────
    for i, alias in enumerate(col_aliases):
        label = col_labels.get(i, alias)
        fill  = DIFF_FILL if i in diff_set else HDR_FILL
        _cel(ws, 3, i + 1, label, bold=True, fill=fill, ha="center", wrap=True)
    ws.row_dimensions[3].height = 36

    # Congela as 4 colunas UG + as 3 primeiras linhas
    ws.freeze_panes = ws.cell(row=4, column=5)
    ws.auto_filter.ref = f"A3:{get_column_letter(nc)}3"

    # ── Linhas de dados ───────────────────────────────────────────────────────
    for r_idx, row in enumerate(rows, start=4):
        # verifica se há alguma diferença relevante
        any_diff = any(
            row[dc] is not None and abs(float(row[dc])) > TOL
            for dc in ctrl["diff_cols"]
            if dc < len(row)
        )

        for i, val in enumerate(row):
            is_diff = i in diff_set
            is_ug   = i < 4

            if is_ug:
                _cel(ws, r_idx, i + 1, val, fill=UG_FILL)
            elif val is None or val == "":
                _cel(ws, r_idx, i + 1, 0.0, fill=ERR_FILL if any_diff else None,
                     ha="right", fmt='#,##0.00')
            else:
                try:
                    fval = float(val)
                    if is_diff:
                        dfill = ERR_FILL if abs(fval) > TOL else OK_FILL
                    elif any_diff:
                        dfill = ERR_FILL
                    else:
                        dfill = None
                    _cel(ws, r_idx, i + 1, fval, fill=dfill,
                         ha="right", fmt='#,##0.00')
                except (TypeError, ValueError):
                    _cel(ws, r_idx, i + 1, val,
                         fill=ERR_FILL if any_diff else None)

    # ── Linha de totais ───────────────────────────────────────────────────────
    r_tot = len(rows) + 4
    ws.merge_cells(start_row=r_tot, start_column=1,
                   end_row=r_tot, end_column=4)
    _cel(ws, r_tot, 1, f"TOTAL  ({len(rows)} linhas)",
         bold=True, fill=TOT_FILL, ha="right")
    for i in range(4, nc):
        try:
            total = sum(
                float(row[i]) for row in rows
                if row[i] is not None
            )
            fill = DIFF_FILL if i in diff_set else TOT_FILL
            _cel(ws, r_tot, i + 1, total, bold=True,
                 fill=fill, ha="right", fmt='#,##0.00')
        except (TypeError, ValueError):
            _cel(ws, r_tot, i + 1, "", fill=TOT_FILL)

    # ── Larguras das colunas ──────────────────────────────────────────────────
    _ajustar_colunas(ws, nc, col_aliases, col_labels, rows)


# ── Escrita de aba: tipo "detalhe" ────────────────────────────────────────────
def _escrever_aba_detalhe(ws, ctrl: dict, col_aliases: list[str],
                          rows: list[tuple], mes: int, ano: int) -> None:
    nc         = len(col_aliases)
    col_labels = ctrl.get("col_labels", {})
    agora      = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    ws.sheet_view.showGridLines = False

    # título
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(row=1, column=1, value=ctrl["titulo"])
    c.font      = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    c.fill      = HDR_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # período + contagem
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nc)
    aviso = "⚠ Todas as linhas retornadas são ocorrências" if rows else "✔ Sem ocorrências"
    c2 = ws.cell(row=2, column=1,
                 value=f"Período: {mes:02d}/{ano}   |   Gerado em: {agora}"
                       f"   |   {len(rows)} ocorrência(s)   |   {aviso}")
    c2.font      = Font(name="Arial", size=9, italic=True,
                        color="CC0000" if rows else "006600")
    c2.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 14

    # cabeçalho
    for i, alias in enumerate(col_aliases):
        label = col_labels.get(i, alias)
        _cel(ws, 3, i + 1, label, bold=True, fill=HDR_FILL, ha="center", wrap=True)
    ws.row_dimensions[3].height = 30

    ws.freeze_panes = ws.cell(row=4, column=1)
    ws.auto_filter.ref = f"A3:{get_column_letter(nc)}3"

    # dados
    diff_set = set(ctrl["diff_cols"])
    for r_idx, row in enumerate(rows, start=4):
        for i, val in enumerate(row):
            is_diff = i in diff_set
            if val is None:
                _cel(ws, r_idx, i + 1, "", fill=DET_FILL)
            else:
                try:
                    fval = float(val)
                    dfill = ERR_FILL if (is_diff and abs(fval) > TOL) else DET_FILL
                    _cel(ws, r_idx, i + 1, fval, fill=dfill,
                         ha="right", fmt='#,##0.00')
                except (TypeError, ValueError):
                    _cel(ws, r_idx, i + 1, val, fill=DET_FILL)

    _ajustar_colunas(ws, nc, col_aliases, col_labels, rows)


# ── Ajuste de larguras ────────────────────────────────────────────────────────
def _ajustar_colunas(ws, nc: int, col_aliases: list[str],
                     col_labels: dict, rows: list[tuple]) -> None:
    for i in range(nc):
        label  = col_labels.get(i, col_aliases[i])
        # calcula largura baseada no rótulo e nos dados
        max_w  = max(len(str(label)), 10)
        for row in rows[:200]:          # amostra as primeiras 200 linhas
            v = row[i]
            if v is not None:
                try:
                    float(v)
                    max_w = max(max_w, 16)
                except (TypeError, ValueError):
                    max_w = max(max_w, len(str(v)))
        col_letter = get_column_letter(i + 1)
        ws.column_dimensions[col_letter].width = min(max_w + 2, 40)


# ── Aba "Resumo" ──────────────────────────────────────────────────────────────
def _escrever_resumo(ws, resultados: list[dict], mes: int, ano: int) -> None:
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:E1")
    c = ws.cell(row=1, column=1,
                value=f"RESUMO DOS CONTROLES DE ROTINA — {mes:02d}/{ano}")
    c.font      = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill      = HDR_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:E2")
    c2 = ws.cell(row=2, column=1, value=f"Gerado em: {agora}")
    c2.font      = Font(name="Arial", size=9, italic=True)
    c2.alignment = Alignment(horizontal="left", vertical="center")

    hdrs = ["Controle", "Título", "Ocorrências", "Status", "Obs."]
    for i, h in enumerate(hdrs, start=1):
        _cel(ws, 3, i, h, bold=True, fill=HDR_FILL, ha="center")

    for r_idx, res in enumerate(resultados, start=4):
        n    = res["n_ocorrencias"]
        tipo = res["tipo"]

        if tipo == "detalhe":
            status = "⚠ OCORRÊNCIAS" if n > 0 else "✔ OK"
            sfill  = ERR_FILL if n > 0 else OK_FILL
            obs    = f"{n} linha(s) retornada(s)"
        else:
            n_dif = res.get("n_diff", 0)
            status = f"⚠ {n_dif} UG(s) c/ diferença" if n_dif > 0 else "✔ OK"
            sfill  = ERR_FILL if n_dif > 0 else OK_FILL
            obs    = f"{n} UG(s) total"

        _cel(ws, r_idx, 1, res["numero"], bold=True, fill=UG_FILL, ha="center")
        _cel(ws, r_idx, 2, res["titulo"])
        _cel(ws, r_idx, 3, n, ha="center", fmt="0")
        _cel(ws, r_idx, 4, status, bold=True, fill=sfill,
             ha="center")
        _cel(ws, r_idx, 5, obs)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 30


# ── Painel HTML ───────────────────────────────────────────────────────────────
PAINEL_DIR   = Path(__file__).parent.parent / "painel"
MAX_ROWS_HTML = 2000          # salvaguarda para controles com muitas linhas

_MESES_PT = ["","Janeiro","Fevereiro","Março","Abril","Maio","Junho",
              "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]


def _he(s) -> str:
    return (str(s).replace("&","&amp;").replace("<","&lt;")
                  .replace(">","&gt;").replace('"','&quot;'))


def _brl(v) -> str:
    if v is None:
        return ""
    try:
        n = float(v)
        neg = n < 0
        s = f"{abs(n):,.2f}".replace(",","\x00").replace(".",",").replace("\x00",".")
        return ("−" if neg else "") + s
    except Exception:
        return _he(v)


def _find_col(labels: dict, targets: set) -> "int | None":
    for k, v in labels.items():
        if v in targets:
            return k
    return None


def _html_ctrl_hierarquia(r: dict, rows: list, hier: dict) -> str:
    """Tabela drill-down 3 níveis: L1=grupo_cols[0], L2=grupo_cols[1], L3=detalhe."""
    from collections import OrderedDict
    gc      = hier["grupo_cols"]    # e.g. [4, 2]
    gl      = hier["grupo_labels"]  # e.g. ["UG","Gestão"]
    sc      = hier["saldo_col"]     # e.g. 3
    dc      = hier["detalhe_cols"]  # e.g. [0, 1]
    numero  = r["numero"]
    col_labels = r.get("col_labels", {})

    cfc = hier.get("conta_filter_col")  # column used for conta selector/filter

    def _v(row, i):
        return row[i] if i < len(row) and row[i] is not None else None

    def _saldo(row):
        try: return float(_v(row, sc) or 0)
        except: return 0.0

    # pré-calcula contas por grupo (para data-contas em L1/L2)
    conta_per_l0: dict = {}
    conta_per_l1: dict = {}
    if cfc is not None:
        for row in rows:
            k0c = str(_v(row, gc[0]) or "").strip()
            k1c = str(_v(row, gc[1]) or "").strip()
            ckv = str(_v(row, cfc) or "").strip()
            if ckv:
                conta_per_l0.setdefault(k0c, set()).add(ckv)
                conta_per_l1.setdefault(f"{k0c}|{k1c}", set()).add(ckv)

    # agrupa em OrderedDict: L0 → L1 → [rows]
    groups: dict = OrderedDict()
    for row in rows:
        k0 = str(_v(row, gc[0]) or "").strip()
        k1 = str(_v(row, gc[1]) or "").strip()
        groups.setdefault(k0, OrderedDict()).setdefault(k1, []).append(row)

    saldo_lbl = _he(col_labels.get(sc, "Saldo"))
    trs = []
    for k0, sub in groups.items():
        tot0 = sum(sum(_saldo(rw) for rw in v) for v in sub.values())
        ek0  = _he(k0)
        l0_contas = "|".join(sorted(conta_per_l0.get(k0, set())))
        d_contas_l0 = f' data-contas="{_he(l0_contas)}"' if cfc is not None else ""
        trs.append(
            f'<tr class="hier-l1" data-key="{ek0}" data-ug="{ek0}"{d_contas_l0}>'
            f'<td class="hier-cell">'
            f'<button class="hier-btn" data-key="{ek0}" onclick="toggleHier(this)">&#9654;</button>'
            f'<span class="hier-lbl">{gl[0]} {_he(k0)}</span></td>'
            f'<td class="num hier-total">{_brl(tot0)}</td></tr>'
        )
        for k1, leaf_rows in sub.items():
            tot1 = sum(_saldo(rw) for rw in leaf_rows)
            ek1  = _he(f"{k0}|{k1}")
            ek1_disp = _he(k1)
            l1_contas = "|".join(sorted(conta_per_l1.get(f"{k0}|{k1}", set())))
            d_contas_l1 = f' data-contas="{_he(l1_contas)}"' if cfc is not None else ""
            trs.append(
                f'<tr class="hier-l2" data-parent="{ek0}" data-key="{ek1}"'
                f' data-ug="{ek0}" data-gest="{_he(k1)}"{d_contas_l1} style="display:none">'
                f'<td class="hier-cell">'
                f'<button class="hier-btn" data-key="{ek1}" onclick="toggleHier(this)">&#9654;</button>'
                f'<span class="hier-lbl">{gl[1]} {ek1_disp}</span></td>'
                f'<td class="num hier-total">{_brl(tot1)}</td></tr>'
            )
            for rw in leaf_rows:
                parts = []
                for ci in dc:
                    v = _v(rw, ci)
                    if v is not None:
                        parts.append(_he(str(v).strip()))
                desc = " &mdash; ".join(p for p in parts if p)
                s    = _saldo(rw)
                conta_val = _he(str(_v(rw, cfc) or "").strip()) if cfc is not None else ""
                d_conta = f' data-conta="{conta_val}"' if cfc is not None else ""
                trs.append(
                    f'<tr class="hier-l3" data-parent="{ek1}"'
                    f' data-ug="{ek0}" data-gest="{_he(k1)}"{d_conta} style="display:none">'
                    f'<td class="hier-cell">{desc}</td>'
                    f'<td class="num">{_brl(s)}</td></tr>'
                )

    grand_total = sum(
        _saldo(rw)
        for sub in groups.values()
        for leaf_rows in sub.values()
        for rw in leaf_rows
    )
    tfoot = (f'<tfoot><tr class="hier-total-row">'
             f'<td class="hier-cell">TOTAL</td>'
             f'<td class="num hier-total">{_brl(grand_total)}</td>'
             f'</tr></tfoot>')

    tbody = "".join(trs)
    thead = (f'<thead><tr>'
             f'<th>Descri&#231;&#227;o</th>'
             f'<th class="num diff-hd">{saldo_lbl}</th>'
             f'</tr></thead>')
    return (f'<div class="tbl-wrap">'
            f'<table class="tbl tbl-hier" data-ctrl="{numero}">'
            f'{thead}<tbody>{tbody}</tbody>{tfoot}'
            f'</table></div>')


def _html_ctrl_section(r: dict) -> str:
    numero     = r["numero"]
    titulo     = _he(r["titulo"])
    chip       = r.get("chip", "inf")
    chip_cls   = {"ok": "c-ok", "err": "c-err", "inf": "c-inf"}.get(chip, "c-inf")
    chip_txt   = {"ok": "OK", "err": "COM DIFERENÇA", "inf": "OCORRÊNCIAS"}.get(chip, "INFO")
    tipo       = r["tipo"]
    n_ocorr    = r.get("n_ocorrencias", 0)
    n_diff     = r.get("n_diff", 0)
    rows       = r.get("rows", [])
    col_labels = r.get("col_labels", {})
    col_aliases= r.get("col_aliases", [])
    diff_set   = set(r.get("diff_cols", []))

    # Para "ug": col 0=Gestão → seletor, col 1=Gestão Contáb. → linha,
    #            col 2=UG Contáb. → linha,    col 3=UG → seletor
    # Para "detalhe": ug_col/gest_col → seletor; detalhe_row_id_cols → linha E seletor
    if tipo == "ug":
        sel_ug_col   = 3   # COUG
        sel_gest_col = 0   # COGESTAO
        skip_cols    = {0, 3}
        row_id_set   = {1, 2}     # Gestão Contáb. e UG Contáb.
        stats = (f"{n_ocorr} UG(s) &nbsp;|&nbsp; {n_diff} c/ diferença"
                 if n_diff else f"{n_ocorr} UG(s) &nbsp;|&nbsp; OK")
    else:
        sel_ug_col   = r.get("ug_col")
        sel_gest_col = r.get("gest_col")
        # detalhe_row_id_cols: aparecem na tabela (ug-cell) E viram seletores (não vão p/ skip)
        detalhe_row_id = set(r.get("detalhe_row_id_cols", []))
        row_id_set   = detalhe_row_id
        skip_cols    = set()
        if sel_ug_col   is not None and sel_ug_col   not in detalhe_row_id: skip_cols.add(sel_ug_col)
        if sel_gest_col is not None and sel_gest_col not in detalhe_row_id: skip_cols.add(sel_gest_col)
        stats = f"{n_ocorr} ocorrência(s)"

    # Rótulos que indicam código contábil — não formatar como BRL
    _CODE_LABELS = {"conta contábil", "conta corrente", "conta corrente (8 dígitos)",
                    "nome conta contábil", "ind. inversão saldo"}

    # ── valores únicos para os seletores deste controle ──────────────────────
    uniq_ugs : list = []
    uniq_gest: list = []
    if sel_ug_col is not None:
        seen: set = set()
        for row in rows:
            v = row[sel_ug_col] if len(row) > sel_ug_col and row[sel_ug_col] is not None else None
            if v is not None:
                s = str(v).strip()
                if s not in seen: seen.add(s); uniq_ugs.append(s)
        uniq_ugs.sort()
    if sel_gest_col is not None:
        seen = set()
        for row in rows:
            v = row[sel_gest_col] if len(row) > sel_gest_col and row[sel_gest_col] is not None else None
            if v is not None:
                s = str(v).strip()
                if s not in seen: seen.add(s); uniq_gest.append(s)
        uniq_gest.sort()

    # seletor extra de conta para controles hierárquicos com conta_filter_col
    uniq_contas_hier: list = []
    _hier_for_sel = r.get("hierarquia")
    if _hier_for_sel:
        _cfc = _hier_for_sel.get("conta_filter_col")
        if _cfc is not None:
            seen = set()
            for row in rows:
                v = row[_cfc] if len(row) > _cfc and row[_cfc] is not None else None
                if v is not None:
                    s = str(v).strip()
                    if s not in seen: seen.add(s); uniq_contas_hier.append(s)
            uniq_contas_hier.sort()

    # ── barra de filtros do controle ──────────────────────────────────────────
    filter_html = ""
    if uniq_ugs or uniq_gest or uniq_contas_hier:
        parts = []
        if uniq_ugs:
            opts = "\n".join(f'<option value="{_he(u)}">{_he(u)}</option>' for u in uniq_ugs)
            parts.append(
                f'<label for="sel-ug-{numero}">UG</label>'
                f'<select id="sel-ug-{numero}" onchange="filtrarCtrl(\'{numero}\')">'
                f'<option value="">&#8212; todas as UGs &#8212;</option>{opts}</select>')
        if uniq_gest:
            opts = "\n".join(f'<option value="{_he(g)}">{_he(g)}</option>' for g in uniq_gest)
            parts.append(
                f'<label for="sel-gest-{numero}">Gest&#227;o</label>'
                f'<select id="sel-gest-{numero}" onchange="filtrarCtrl(\'{numero}\')">'
                f'<option value="">&#8212; todas as gest&#245;es &#8212;</option>{opts}</select>')
        if uniq_contas_hier:
            _cfc_lbl = _he(col_labels.get(_hier_for_sel.get("conta_filter_col", -1), "Conta Cont&#225;bil"))
            opts = "\n".join(f'<option value="{_he(c)}">{_he(c)}</option>' for c in uniq_contas_hier)
            parts.append(
                f'<label for="sel-conta-{numero}">{_cfc_lbl}</label>'
                f'<select id="sel-conta-{numero}" onchange="filtrarCtrl(\'{numero}\')">'
                f'<option value="">&#8212; todas &#8212;</option>{opts}</select>')
        parts.append(f'<button onclick="limparCtrl(\'{numero}\')">Limpar</button>')
        parts.append(f'<span class="ctrl-filter-info" id="info-{numero}"></span>')
        filter_html = f'<div class="ctrl-filter">{"".join(parts)}</div>'

    # ── modo hierárquico (drill-down) quando configurado ─────────────────────
    hier = r.get("hierarquia")
    if hier and rows:
        body = _html_ctrl_hierarquia(r, rows, hier)
    elif not rows:
        body = '<p class="ok-msg">Sem ocorrências</p>'
    else:
        # ── ordem das colunas na tabela (modo flat) ───────────────────────────
        if tipo == "ug":
            col_order = [2, 1] + [i for i in range(4, len(col_aliases))]
        else:
            row_ids_ordered = list(r.get("detalhe_row_id_cols", []))
            other_cols = [i for i in range(len(col_aliases))
                          if i not in skip_cols and i not in row_id_set]
            col_order = row_ids_ordered + other_cols

        text_col_set = {i for i, alias in enumerate(col_aliases)
                        if col_labels.get(i, alias).lower() in _CODE_LABELS}

        ths = []
        for i in col_order:
            label = col_labels.get(i, col_aliases[i] if i < len(col_aliases) else str(i))
            cls   = ' class="diff-hd"' if i in diff_set else ""
            ths.append(f"<th{cls}>{_he(label)}</th>")

        display_rows = rows[:MAX_ROWS_HTML]
        trs = []
        for row in display_rows:
            ug_v = (str(row[sel_ug_col]).strip()
                    if sel_ug_col is not None and len(row) > sel_ug_col and row[sel_ug_col] is not None
                    else "")
            gs_v = (str(row[sel_gest_col]).strip()
                    if sel_gest_col is not None and len(row) > sel_gest_col and row[sel_gest_col] is not None
                    else "")

            if tipo == "ug":
                has_diff = any(
                    len(row) > dc and row[dc] is not None and abs(float(row[dc])) > TOL
                    for dc in diff_set if dc < len(row)
                )
                row_cls = "has-diff" if has_diff else "ok-row"
            else:
                row_cls = ""

            tr = f'<tr data-ug="{_he(ug_v)}" data-gest="{_he(gs_v)}"'
            if row_cls:
                tr += f' class="{row_cls}"'
            tr += ">"
            for i in col_order:
                val        = row[i] if i < len(row) else None
                is_row_id  = (i in row_id_set)
                is_diff    = i in diff_set
                is_txt_col = (i in text_col_set)
                if is_row_id:
                    tr += f'<td class="ug-cell">{_he(val) if val is not None else ""}</td>'
                elif val is None:
                    tr += f'<td class="{"num diff-cell" if is_diff else "num"}"></td>'
                elif is_txt_col:
                    tr += f"<td>{_he(val)}</td>"
                else:
                    try:
                        float(val)
                        tr += f'<td class="{"num diff-cell" if is_diff else "num"}">{_brl(val)}</td>'
                    except (TypeError, ValueError):
                        tr += f"<td>{_he(val)}</td>"
            tr += "</tr>"
            trs.append(tr)

        trunc = (f'<p class="trunc-note">&#9888; Exibindo primeiras {MAX_ROWS_HTML:,}'
                 f' de {len(rows):,} linhas. Consulte o Excel para o total.</p>'
                 if len(rows) > MAX_ROWS_HTML else "")
        body = (f'<div class="tbl-wrap"><table class="tbl" data-ctrl="{numero}">'
                f'<thead><tr>{"".join(ths)}</tr></thead>'
                f'<tbody>{"".join(trs)}</tbody>'
                f'</table></div>{trunc}')

    return (f'<details class="ctrl" id="ctrl-{numero}">\n'
            f'  <summary><span class="ctrl-num">C{numero}</span>'
            f'    {titulo}'
            f'    <span class="chip {chip_cls}">{chip_txt}</span></summary>\n'
            f'  <div class="ctrl-body">'
            f'    {filter_html}'
            f'    <span class="ctrl-stats">{stats}</span>'
            f'    {body}'
            f'  </div>\n</details>')


def _gerar_html(resultados_html: list, mes: int, ano: int,
                meses_disponiveis: list | None = None,
                nome_atual: str | None = None) -> str:
    agora    = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    mes_nome = _MESES_PT[mes] if 1 <= mes <= 12 else str(mes)

    n_total = len(resultados_html)
    n_err   = sum(1 for r in resultados_html if r.get("chip") == "err")
    n_ok    = sum(1 for r in resultados_html if r.get("chip") == "ok")
    n_inf   = sum(1 for r in resultados_html if r.get("chip") == "inf")

    nav_items = "".join(f'<a href="#ctrl-{r["numero"]}">C{r["numero"]}</a>' for r in resultados_html)
    secoes    = "\n".join(_html_ctrl_section(r) for r in resultados_html)

    # ── seletor de histórico: uma opção por execução com data/hora ────────────
    # meses_disponiveis: lista de (ano, mes, fname, lbl) mais recente primeiro
    _sel_fname = nome_atual or f"rotina_controles_{ano}_{mes:02d}.html"
    if meses_disponiveis and len(meses_disponiveis) > 1:
        opts_hist = "\n".join(
            f'<option value="{_he(fname)}"{"selected" if fname == _sel_fname else ""}>'
            f'{_he(lbl)}</option>'
            for a, m, fname, lbl in meses_disponiveis
        )
        run_wrap = (
            f'<div class="run-wrap">'
            f'<span class="run-lbl">Compet&#234;ncia</span>'
            f'<select id="run-select" onchange="window.location.href=this.value">'
            f'{opts_hist}</select>'
            f'</div>'
        )
    else:
        run_wrap = ""

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Controles de Rotina GDF {mes:02d}/{ano}</title>
<style>
:root{{
  --brand:#1A4A8F;--pg:#F1F4FB;--s1:#FFF;--s2:#E8EDF7;--bd:#C8D5ED;
  --t1:#0D1829;--t2:#4C5C7A;--t3:#8A9BBD;
  --err:#BE1C1C;--err-bg:#FEF2F2;--err-bd:#FECACA;
  --ok:#156030;--ok-bg:#F0FDF4;--ok-bd:#BBF7D0;
  --inf:#1547A0;--inf-bg:#EFF6FF;--inf-bd:#BFDBFE;
  --diff:#7030A0;
  --fn:'Segoe UI',system-ui,-apple-system,BlinkMacSystemFont,sans-serif;
  --fm:'Cascadia Code','SF Mono',Consolas,'Courier New',monospace;
  --r:8px;
}}
@media(prefers-color-scheme:dark){{:root:not([data-theme="light"]){{
  --brand:#4889D8;--pg:#080C18;--s1:#101827;--s2:#182035;--bd:#1D2C48;
  --t1:#D4DFF5;--t2:#6A7EA6;--t3:#3D5070;
  --err:#F87171;--err-bg:#170404;--err-bd:#7F1D1D;
  --ok:#4ADE80;--ok-bg:#031309;--ok-bd:#14532D;
  --inf:#93C5FD;--inf-bg:#0C1A35;--inf-bd:#1E3A5F;
  --diff:#C084FC;
}}}}
:root[data-theme="dark"]{{
  --brand:#4889D8;--pg:#080C18;--s1:#101827;--s2:#182035;--bd:#1D2C48;
  --t1:#D4DFF5;--t2:#6A7EA6;--t3:#3D5070;
  --err:#F87171;--err-bg:#170404;--err-bd:#7F1D1D;
  --ok:#4ADE80;--ok-bg:#031309;--ok-bd:#14532D;
  --inf:#93C5FD;--inf-bg:#0C1A35;--inf-bd:#1E3A5F;
  --diff:#C084FC;
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:var(--fn);font-size:14px;line-height:1.5;color:var(--t1);background:var(--pg)}}
.hd{{background:var(--brand);color:#fff;padding:14px 24px;display:flex;align-items:center;
    justify-content:space-between;gap:16px;position:sticky;top:0;z-index:10}}
.hd-title{{font-size:15px;font-weight:700;letter-spacing:-.01em}}
.hd-sub{{font-size:12px;opacity:.72;margin-top:2px}}
.hd-ts{{font-size:11px;opacity:.55;margin-top:3px;font-family:var(--fm)}}
.hd-badge{{font-size:11px;font-weight:600;background:rgba(255,255,255,.18);
          border:1px solid rgba(255,255,255,.28);border-radius:100px;padding:3px 11px}}
.wrap{{max-width:1200px;margin:0 auto;padding:24px 20px;display:flex;flex-direction:column;gap:14px}}
.ctrl-filter{{display:flex;flex-wrap:wrap;align-items:center;gap:8px;
             padding:8px 0 4px;border-bottom:1px solid var(--bd);margin-bottom:2px}}
.ctrl-filter label{{font-size:11px;font-weight:600;color:var(--t2);white-space:nowrap}}
.ctrl-filter select{{font-size:11px;font-family:var(--fm);border:1px solid var(--bd);
                    border-radius:5px;padding:3px 7px;background:var(--s2);color:var(--t1);
                    min-width:120px;max-width:200px}}
.ctrl-filter button{{font-size:10px;padding:3px 10px;border:1px solid var(--bd);
                    border-radius:5px;background:var(--s2);color:var(--t2);cursor:pointer}}
.ctrl-filter button:hover{{background:var(--bd)}}
.ctrl-filter-info{{font-size:10px;color:var(--t3);margin-left:4px}}
.nav{{display:flex;flex-wrap:wrap;gap:6px;padding:10px 14px;background:var(--s1);
     border:1px solid var(--bd);border-radius:var(--r)}}
.nav a{{font-size:11px;font-weight:600;color:var(--brand);text-decoration:none;
       padding:3px 9px;border-radius:100px;border:1px solid var(--bd)}}
.nav a:hover{{border-color:var(--brand);background:var(--s2)}}
.kpi-row{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}}
.kpi{{background:var(--s1);border:1px solid var(--bd);border-radius:var(--r);
     padding:13px 15px;position:relative;overflow:hidden}}
.kpi::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;
             background:var(--kpi-stripe,var(--bd))}}
.kpi-lbl{{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--t3);margin-bottom:5px}}
.kpi-val{{font-size:26px;font-weight:700;letter-spacing:-.025em;color:var(--kpi-color,var(--t1));
         font-variant-numeric:tabular-nums}}
.kpi-sub{{font-size:11px;color:var(--t3);margin-top:3px}}
details.ctrl{{background:var(--s1);border:1px solid var(--bd);border-radius:var(--r);overflow:hidden}}
details.ctrl summary{{padding:11px 18px;font-size:13.5px;font-weight:600;cursor:pointer;
                     list-style:none;display:flex;align-items:center;gap:10px;background:var(--s2);
                     user-select:none}}
details.ctrl summary::-webkit-details-marker{{display:none}}
details.ctrl summary::after{{content:'+';font-size:16px;font-weight:400;color:var(--t3);margin-left:auto}}
details.ctrl[open] summary::after{{content:'−'}}
.ctrl-num{{font-family:var(--fm);font-size:11px;font-weight:600;color:var(--t3);
          background:var(--pg);border:1px solid var(--bd);border-radius:4px;
          padding:1px 7px;letter-spacing:.02em;flex-shrink:0}}
.ctrl-body{{padding:14px 18px;display:flex;flex-direction:column;gap:10px;border-top:1px solid var(--bd)}}
.ctrl-stats{{font-size:12px;color:var(--t2)}}
.chip{{font-size:10px;font-weight:700;padding:2px 9px;border-radius:100px;border:1px solid;white-space:nowrap;flex-shrink:0}}
.c-ok{{background:var(--ok-bg);color:var(--ok);border-color:var(--ok-bd)}}
.c-err{{background:var(--err-bg);color:var(--err);border-color:var(--err-bd)}}
.c-inf{{background:var(--inf-bg);color:var(--inf);border-color:var(--inf-bd)}}
.tbl-wrap{{overflow-x:auto;border:1px solid var(--bd);border-radius:6px;max-height:480px;overflow-y:auto}}
table.tbl{{width:100%;border-collapse:collapse;font-size:11.5px;font-family:var(--fm)}}
table.tbl th,table.tbl td{{padding:4px 9px;border-bottom:1px solid var(--bd);white-space:nowrap}}
table.tbl th{{background:var(--s2);font-weight:700;font-family:var(--fn);position:sticky;top:0;z-index:1}}
table.tbl th.diff-hd{{background:var(--diff);color:#fff}}
table.tbl td.num{{text-align:right;font-variant-numeric:tabular-nums}}
table.tbl td.ug-cell{{background:rgba(217,232,245,.4);font-weight:600}}
table.tbl tr.has-diff td{{background:var(--err-bg)}}
table.tbl tr.has-diff td.ug-cell{{background:#fde8ea}}
table.tbl tr.has-diff td.diff-cell{{color:var(--err);font-weight:700}}
table.tbl tr.ok-row td.diff-cell{{color:var(--ok);font-weight:600}}
table.tbl td.diff-cell{{font-weight:600}}
.ok-msg{{color:var(--ok);font-size:13px;font-weight:600;display:flex;align-items:center;gap:7px}}
.ok-msg::before{{content:'✓';font-size:15px}}
.trunc-note{{font-size:11px;color:var(--inf);font-style:italic;padding:6px 2px}}
footer{{font-size:11px;color:var(--t3);text-align:center;padding:14px;margin-top:10px}}
@media(max-width:680px){{.kpi-row{{grid-template-columns:repeat(2,1fr)}}}}
.run-wrap{{display:flex;align-items:center;gap:10px;padding:9px 14px;background:var(--s1);
          border:1px solid var(--bd);border-radius:var(--r)}}
.run-lbl{{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--t3);
         white-space:nowrap;flex-shrink:0}}
#run-select{{flex:1;padding:5px 9px;border:1px solid var(--bd);border-radius:6px;
            background:var(--s2);color:var(--t1);font-size:12px;font-family:var(--fn);
            cursor:pointer;max-width:320px}}
#run-select:focus{{outline:2px solid var(--brand);outline-offset:1px}}
/* ── Tabela hierárquica (drill-down) ── */
.tbl-hier tr.hier-l1 td{{background:var(--s2);font-weight:700;font-size:12px}}
.tbl-hier tr.hier-l2 td{{background:var(--s1);font-size:11.5px}}
.tbl-hier tr.hier-l3 td{{background:var(--pg);font-size:11px;color:var(--t2)}}
.hier-cell{{white-space:nowrap}}
.tbl-hier tr.hier-l2 .hier-cell{{padding-left:22px}}
.tbl-hier tr.hier-l3 .hier-cell{{padding-left:42px}}
.hier-btn{{background:none;border:none;cursor:pointer;font-size:10px;color:var(--brand);
          padding:0 5px 0 0;line-height:1;transition:transform .15s;display:inline-block;
          vertical-align:middle}}
.hier-btn.exp{{transform:rotate(90deg)}}
.hier-lbl{{vertical-align:middle}}
.hier-total{{font-variant-numeric:tabular-nums;font-weight:700}}
.tbl-hier tfoot tr.hier-total-row td{{
  background:var(--brand);color:#fff;font-weight:700;font-size:12px;
  position:sticky;bottom:0;z-index:1;border-top:2px solid var(--brand)}}
</style>
</head>
<body>
<div class="hd">
  <div>
    <div class="hd-title">Controles de Rotina &#8212; GDF</div>
    <div class="hd-sub">Dados de competência {mes:02d}/{ano} ({mes_nome}/{ano})</div>
    <div class="hd-ts">Executado em {agora}</div>
  </div>
  <div class="hd-badge">{mes:02d}/{ano}</div>
</div>
<div class="wrap">
  {run_wrap}
  <div class="kpi-row">
    <div class="kpi" style="--kpi-stripe:var(--brand)">
      <div class="kpi-lbl">Controles</div><div class="kpi-val">{n_total}</div>
      <div class="kpi-sub">executados em {mes:02d}/{ano}</div>
    </div>
    <div class="kpi" style="--kpi-stripe:var(--err);--kpi-color:var(--err)">
      <div class="kpi-lbl">Com Diferença</div><div class="kpi-val">{n_err}</div>
    </div>
    <div class="kpi" style="--kpi-stripe:var(--ok);--kpi-color:var(--ok)">
      <div class="kpi-lbl">OK</div><div class="kpi-val">{n_ok}</div>
    </div>
    <div class="kpi" style="--kpi-stripe:var(--inf);--kpi-color:var(--inf)">
      <div class="kpi-lbl">Ocorrências</div><div class="kpi-val">{n_inf}</div>
    </div>
  </div>
  <nav class="nav" aria-label="Controles">{nav_items}</nav>
  {secoes}
</div>
<footer>Gerado em {agora} &mdash; Controles de Rotina GDF {mes:02d}/{ano} &mdash; dados de MIL{ano}.LANCAMENTOCONTABIL</footer>
<script>
/* ── toggle drill-down ── */
function toggleHier(btn){{
  var key=btn.dataset.key;
  var tbl=btn.closest('table');
  var exp=btn.classList.toggle('exp');
  tbl.querySelectorAll('tr[data-parent="'+key+'"]').forEach(function(tr){{
    tr.style.display=exp?'':'none';
    if(!exp){{
      var ck=tr.dataset.key;
      if(ck){{
        tbl.querySelectorAll('tr[data-parent="'+ck+'"]').forEach(function(g){{g.style.display='none';}});
        var cb=tr.querySelector('.hier-btn');
        if(cb)cb.classList.remove('exp');
      }}
    }}
  }});
}}
/* ── filtros por controle ── */
function filtrarCtrl(num){{
  var ugEl=document.getElementById('sel-ug-'+num);
  var gsEl=document.getElementById('sel-gest-'+num);
  var ctEl=document.getElementById('sel-conta-'+num);
  var ug=ugEl?ugEl.value:'';
  var gs=gsEl?gsEl.value:'';
  var conta=ctEl?ctEl.value:'';
  var tbl=document.querySelector('table[data-ctrl="'+num+'"]');
  if(!tbl)return;
  var isHier=tbl.classList.contains('tbl-hier');
  var info=document.getElementById('info-'+num);
  if(isHier){{
    /* colapsa tudo e filtra L1 por UG e/ou Conta Contábil */
    tbl.querySelectorAll('tbody tr').forEach(function(tr){{
      if(tr.classList.contains('hier-l1')){{
        var okUg=!ug||tr.dataset.ug===ug;
        var okCt=!conta||(tr.dataset.contas||'').split('|').indexOf(conta)>=0;
        tr.style.display=(okUg&&okCt)?'':'none';
      }} else {{
        tr.style.display='none';
        var cb=tr.querySelector('.hier-btn');
        if(cb)cb.classList.remove('exp');
      }}
    }});
    if(info){{
      var vis=tbl.querySelectorAll('tr.hier-l1').length - tbl.querySelectorAll('tr.hier-l1[style*="none"]').length;
      var tot=tbl.querySelectorAll('tr.hier-l1').length;
      info.textContent=(ug||conta)?vis+' de '+tot+' UG(s)':'';
    }}
  }} else {{
    var vis=0,tot=0;
    tbl.querySelectorAll('tbody tr').forEach(function(tr){{
      var m=(!ug||tr.dataset.ug===ug)&&(!gs||tr.dataset.gest===gs);
      tr.style.display=m?'':'none'; tot++;if(m)vis++;
    }});
    if(info)info.textContent=(ug||gs)?vis+' de '+tot+' linha(s)':'';
  }}
}}
function limparCtrl(num){{
  var ugEl=document.getElementById('sel-ug-'+num);
  var gsEl=document.getElementById('sel-gest-'+num);
  var ctEl=document.getElementById('sel-conta-'+num);
  if(ugEl)ugEl.value='';
  if(gsEl)gsEl.value='';
  if(ctEl)ctEl.value='';
  var tbl=document.querySelector('table[data-ctrl="'+num+'"]');
  if(!tbl)return;
  if(tbl.classList.contains('tbl-hier')){{
    tbl.querySelectorAll('tbody tr').forEach(function(tr){{
      tr.style.display=tr.classList.contains('hier-l1')?'':'none';
      var cb=tr.querySelector('.hier-btn');
      if(cb)cb.classList.remove('exp');
    }});
  }} else {{
    tbl.querySelectorAll('tbody tr').forEach(function(tr){{tr.style.display='';}});
  }}
  var info=document.getElementById('info-'+num);
  if(info)info.textContent='';
}}
</script>
</body>
</html>"""


# ── Principal ─────────────────────────────────────────────────────────────────
def main() -> None:
    p = argparse.ArgumentParser(
        description="Gera Excel dos Controles de Rotina (Discoverer-like)")
    p.add_argument("--mes",      type=int, required=True, help="Mês (1-12)")
    p.add_argument("--ano",      type=int, required=True, help="Ano (ex.: 2026)")
    p.add_argument("--controle", type=str, default=None,
                   help="Roda só este controle (ex.: 02). Omitir = todos.")
    args = p.parse_args()

    controles_exec = CONTROLES
    if args.controle:
        controles_exec = [c for c in CONTROLES if c["numero"] == args.controle]
        if not controles_exec:
            sys.exit(f"ERRO: controle '{args.controle}' não encontrado.")

    conn = conectar()
    try:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arq = f"Controles_Rotina_{args.ano}_{args.mes:02d}_{ts}.xlsx"
        caminho  = OUTPUT_DIR / nome_arq

        wb = openpyxl.Workbook()
        wb.remove(wb.active)          # remove aba vazia padrão

        resultados      = []
        resultados_html = []          # coleta dados para o painel HTML

        for ctrl in controles_exec:
            sql_path = ROTINA_DIR / ctrl["sql"]
            if not sql_path.exists():
                print(f"  [AVISO] SQL não encontrado: {sql_path.name} — pulando.")
                continue

            print(f"  Executando {ctrl['numero']} — {ctrl['titulo']}...")
            try:
                col_aliases, rows = executar(conn, sql_path, args.ano, args.mes)
            except Exception as exc:
                print(f"    ERRO ao executar: {exc}")
                resultados.append({
                    "numero": ctrl["numero"], "titulo": ctrl["titulo"],
                    "tipo": ctrl["tipo"], "n_ocorrencias": 0, "n_diff": 0,
                })
                continue

            ws = wb.create_sheet(title=ctrl["aba"])

            col_labels = ctrl.get("col_labels", {})

            if ctrl["tipo"] == "ug":
                _escrever_aba_ug(ws, ctrl, col_aliases, rows,
                                 args.mes, args.ano)
                n_diff = sum(
                    1 for row in rows
                    if any(
                        row[dc] is not None and abs(float(row[dc])) > TOL
                        for dc in ctrl["diff_cols"] if dc < len(row)
                    )
                )
                resultados.append({
                    "numero": ctrl["numero"], "titulo": ctrl["titulo"],
                    "tipo": ctrl["tipo"], "n_ocorrencias": len(rows),
                    "n_diff": n_diff,
                })
                chip   = "err" if n_diff else "ok"
                status = f"{n_diff} UG(s) c/ diferença" if n_diff else "OK"
                resultados_html.append({
                    "numero":      ctrl["numero"],
                    "titulo":      ctrl["titulo"],
                    "tipo":        "ug",
                    "n_ocorrencias": len(rows),
                    "n_diff":      n_diff,
                    "chip":        chip,
                    "col_labels":  col_labels,
                    "diff_cols":   ctrl.get("diff_cols", []),
                    "col_aliases": col_aliases,
                    "rows":        rows,
                })
            else:
                _escrever_aba_detalhe(ws, ctrl, col_aliases, rows,
                                      args.mes, args.ano)
                resultados.append({
                    "numero": ctrl["numero"], "titulo": ctrl["titulo"],
                    "tipo": ctrl["tipo"], "n_ocorrencias": len(rows),
                })
                chip   = "inf" if rows else "ok"
                status = f"{len(rows)} ocorrência(s)" if rows else "OK"
                # detecta colunas UG e Gestão pelo rótulo
                ug_col   = _find_col(col_labels, {"UG"})
                gest_col = _find_col(col_labels, {"Gestão", "Gestão Contáb.", "Gestão Contáb"})
                resultados_html.append({
                    "numero":      ctrl["numero"],
                    "titulo":      ctrl["titulo"],
                    "tipo":        "detalhe",
                    "n_ocorrencias": len(rows),
                    "chip":        chip,
                    "col_labels":  col_labels,
                    "diff_cols":   ctrl.get("diff_cols", []),
                    "col_aliases": col_aliases,
                    "rows":        rows,
                    "ug_col":      ug_col,
                    "gest_col":    gest_col,
                    "detalhe_row_id_cols": ctrl.get("detalhe_row_id_cols", []),
                    "hierarquia":  ctrl.get("hierarquia"),
                })

            print(f"    → {len(rows)} linha(s)  |  {status}")

        # aba Resumo sempre na frente
        ws_res = wb.create_sheet(title="Resumo", index=0)
        _escrever_resumo(ws_res, resultados, args.mes, args.ano)

        wb.save(caminho)
        print(f"\n  ✔ Arquivo salvo: {caminho}")

        # ── Painel HTML ────────────────────────────────────────────────────────
        if resultados_html:
            PAINEL_DIR.mkdir(parents=True, exist_ok=True)

            # nome com timestamp (arquivo arquivado) e nome "latest"
            _ts_str   = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_ts   = f"rotina_controles_{args.ano}_{args.mes:02d}_{_ts_str}.html"
            nome_html = f"rotina_controles_{args.ano}_{args.mes:02d}.html"

            # descobre execuções arquivadas (rotina_controles_AAAA_MM_YYYYMMDD_HHmmss.html)
            import re as _re
            _pat_ts = _re.compile(
                r"rotina_controles_(\d{4})_(\d{2})_(\d{8})_(\d{6})\.html")
            _meses: list = []
            for _f in PAINEL_DIR.glob("rotina_controles_*_*_*.html"):
                _m = _pat_ts.match(_f.name)
                if _m:
                    _a, _mo = int(_m.group(1)), int(_m.group(2))
                    _d, _t  = _m.group(3), _m.group(4)
                    _mn = _MESES_PT[_mo] if 1 <= _mo <= 12 else str(_mo)
                    _lbl = f"{_mn}/{_a} — {_d[6:8]}/{_d[4:6]}/{_d[:4]} {_t[:2]}:{_t[2:4]}"
                    _meses.append((_a, _mo, _d + _t, _f.name, _lbl))

            # inclui a execução atual (ainda não salva)
            _d0, _t0 = _ts_str[:8], _ts_str[9:]
            _mn0 = _MESES_PT[args.mes] if 1 <= args.mes <= 12 else str(args.mes)
            _lbl0 = f"{_mn0}/{args.ano} — {_d0[6:8]}/{_d0[4:6]}/{_d0[:4]} {_t0[:2]}:{_t0[2:4]}"
            _meses.append((args.ano, args.mes, _d0 + _t0, nome_ts, _lbl0))
            _meses.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
            meses_disp = [(a, mo, fname, lbl) for a, mo, _, fname, lbl in _meses]

            html = _gerar_html(resultados_html, args.mes, args.ano,
                               meses_disp, nome_atual=nome_ts)

            # salva arquivo arquivado (com timestamp)
            (PAINEL_DIR / nome_ts).write_text(html, encoding="utf-8")
            # atualiza "latest" do mês
            (PAINEL_DIR / nome_html).write_text(html, encoding="utf-8")
            print(f"  ✔ Painel HTML salvo: {PAINEL_DIR / nome_ts}")
            print(f"  ✔ Latest atualizado: {PAINEL_DIR / nome_html}")

    finally:
        conn.close()


if __name__ == "__main__":
    main()
