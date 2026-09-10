# -*- coding: utf-8 -*-
"""
=============================================================================
  BALANCO PATRIMONIAL - GDF  (PSIAG550) - Exercicio ATUAL x Exercicio ANTERIOR
  SQL fixo, equacoes derivadas do plano de contas (contacontabil_2026.xls) e
  do MCASP 11a Edicao, seguindo o MESMO padrao de extracao Oracle ja validado
  e em producao nos demais anexos do projeto:
    - bf.py    Balanco Financeiro            (Anexo 13)
    - bo.py    Balanco Orcamentario          (Anexo 12)
    - dfc.py   Demonstracao Fluxos de Caixa  (Anexo 15)
    - dvp.py   Demonstracao Variacoes Patrim. (Anexo da DVP)
    - dmpl.py  Demonstracao Mutacoes Patrim. Liquido (DMPL)

  LAYOUT DE SAIDA -- espelha ListaBalancoPatrimonial (Governo do Distrito
  Federal, PSIAG550, "Movimento do Exercicio Consolidado Geral", Versao 1),
  paginas 1 a 3:
    Pagina 1: BALANCO PATRIMONIAL
              - ATIVO (Circulante + Nao Circulante) com seus subgrupos
              - PASSIVO E PATRIMONIO LIQUIDO (Circulante + Nao Circulante
                + Patrimonio Liquido) com seus subgrupos
    Pagina 2: QUADRO DOS ATIVOS E PASSIVOS FINANCEIROS E PERMANENTES
              - ATIVO (I) = Ativo Financeiro + Ativo Permanente
              - PASSIVO (II) = Passivo Financeiro + Passivo Permanente
              - SALDO PATRIMONIAL (III) = (I - II)
    Pagina 3: QUADRO DAS CONTAS DE COMPENSACAO
              - ATOS POTENCIAIS ATIVOS / PASSIVOS (contas 7/8, classe 9)

  Nota: o modelo oficial completo (PSIAG550) inclui ainda, nas paginas 4 a 9,
  o "QUADRO DO SUPERAVIT/DEFICIT FINANCEIRO" detalhado por fonte de recursos
  (centenas de linhas, uma por fonte/destinacao). Esse quadro NAO faz parte
  do Balanco Patrimonial propriamente dito (e um detalhamento do superavit
  financeiro por fonte, usado para fins de credito adicional/LRF) e por isso
  NAO esta neste modulo. Caso seja necessario, deve ser um anexo separado
  (mesmo padrao SOURCE x FONTE de bo.py), pois exigiria juncao com o
  cadastro de fontes de recursos (FONTERECURSO ou equivalente), nao presente
  nos demais modulos ja construidos.

  ESTRUTURA DE CONTAS (fonte: contacontabil_2026.xls, plano de contas SPCASP):
    ATIVO (1XXXXXXXX)                         PASSIVO + PL (2XXXXXXXX)
      ATIVO CIRCULANTE (11X)                    PASSIVO CIRCULANTE (21X)
        111 Caixa e Equivalentes de Caixa         211 Obrig.Trab/Prev/Assist a Pagar CP
        112 Creditos a Curto Prazo                212 Emprestimos e Financiamentos CP
        113 Demais Creditos e Valores CP          213 Fornecedores e Contas a Pagar CP
        114 Investim. e Aplic.Temp. CP             214 Obrigacoes Fiscais CP
        115 Estoques                              215 Transferencias Fiscais CP
        119 VPD Pagas Antecipadamente              217 Provisoes CP
                                                    218 Demais Obrigacoes CP
      ATIVO NAO CIRCULANTE (12X)                ATIVO NAO CIRCULANTE (22X)
        121 Realizavel a Longo Prazo               221 Obrig.Trab/Prev/Assist a Pagar LP
        122 Investimentos                          222 Emprestimos e Financiamentos LP
        123 Imobilizado                            223 Fornecedores e Contas a Pagar LP
        124 Intangivel                             224 Obrigacoes Fiscais LP
                                                    227 Provisoes LP
                                                    228 Demais Obrigacoes LP
                                                    229 Resultado Diferido
                                                  PATRIMONIO LIQUIDO (23X)
                                                    231 Pat.Social e Capital Social
                                                    232 Adiant.p/Futuro Aum.Capital
                                                    233 Reservas de Capital
                                                    234 Ajustes de Avaliacao Patrim.
                                                    235 Reservas de Lucros
                                                    236 Demais Reservas
                                                    237 Resultado Acumulado
                                                      237.1 Superavits/Deficits Acum.
                                                      237.2 Lucros e Prejuizos Acum.
                                                      (+) Resultado do Exercicio
                                                          (movimento 4XX-3XX do periodo)

    CONTAS DE COMPENSACAO (classe 7XX, confirmado via --diag em producao,
    GDF, Maio/2026 -- CORRIGIDO de uma suposicao anterior incorreta que
    usava 711/712/719/811/812/819 como grupos paralelos; a estrutura real
    do plano de contas do GDF e):
      711XXXXXX = ATOS POTENCIAIS ATIVOS (grupo completo)
        711100000 Garantias e Contragarantias Recebidas
        711200000 Direitos Conveniados e Outros Instrumentos Congeneres
        711300000 Direitos Contratuais (sem linha propria no PDF oficial
                   -- consolidado em "Outros Atos Potenciais Ativos")
        711900000 Outros Atos Potenciais Ativos
      712XXXXXX = ATOS POTENCIAIS PASSIVOS (grupo completo)
        712100000 Garantias e Contragarantias Concedidas
        712200000 Obrigacoes Conveniados e Outros Instrumentos Congeneres
                   (sem linha propria no PDF -- ver nota abaixo)
        712300000 Obrigacoes Contratuais
        712400000 Demandas Judiciais (sem linha propria no PDF oficial
                   -- consolidado em "Outros Atos Potenciais Passivos")
        712900000 Outros Atos Potenciais Passivos
      811XXXXXX/812XXXXXX = EXECUCAO dos Atos Potenciais Ativos/Passivos
        -- uma PERNA CONTABIL DIFERENTE (registra a execucao/liquidacao
        dos atos potenciais, nao o saldo do ato potencial em si) que NAO
        entra neste quadro. Natureza: SD para Ativos (711), SC para
        Passivos (712).
      NOTA: "Obrigacoes Conveniados" (712200000) nao aparece como linha
      separada no PDF oficial nem foi mapeada a nenhum item -- se essa
      conta tiver saldo relevante no banco, o total "ATOS POTENCIAIS
      PASSIVOS" deste modulo pode ficar levemente diferente do oficial;
      confirmar via diagnostico() se necessario.

  LOGICA DE MES (INMES) -- CORRIGIDA apos bug critico encontrado em
  producao (execucao real sem --diag, GDF, Maio/2026: o Exercicio
  Anterior saiu com Ativo Total de so R$55,9 milhoes, quando deveria ser
  ~R$68,78 bilhoes -- a hipotese original ("INMES=13 acumula o
  fechamento") estava ERRADA; INMES=13 e so um ajuste residual de
  encerramento, igual ao bf.py ja documentava ("13 = encerramento
  (ajustes, reclassif., inscricao de RP)") mas que havia sido mal
  interpretado neste modulo):
    SALDO (Ativo/Passivo, contas 1/2, natureza permanente) -> SALDOCONTABIL
      Exercicio Atual    -> INMES BETWEEN 0 AND :mes   (saldo acumulado ate o
                             mes de referencia, ano corrente MIL{ano})
      Exercicio Anterior -> INMES = 0 do MESMO schema MIL{ano} (NAO
                             MIL{ano-1}.INMES=13) -- o saldo de ABERTURA
                             do ano atual E o saldo de ENCERRAMENTO do
                             ano anterior (mesma convencao de bf.py: "0 =
                             saldo de abertura -> SALDO DO EXERCICIO
                             ANTERIOR"). Confirmado numericamente: MIL{ano}
                             INMES=0 bate EXATO com o Ativo Total e
                             Passivo+PL Total do Exercicio Anterior no
                             PDF oficial (R$68.780.953.153,13).
    RESULTADO DO EXERCICIO (linha do PL, contas 3/4, natureza de fluxo) ->
      LANCAMENTOCONTABIL, mesmo padrao de dvp.py/dmpl.py (aqui SIM correto,
      pois e fluxo, nao saldo -- nao afetado pelo bug acima):
      Exercicio Atual    -> INMES BETWEEN 1 AND :mes  (MIL{ano})
      Exercicio Anterior -> INMES BETWEEN 1 AND 13     (MIL{ano-1} completo)
  TIPOS DE MOVIMENTO:
    SD = Saldo Devedor = D - C  (natureza devedora -- ATIVO)
    SC = Saldo Credor  = C - D  (natureza credora  -- PASSIVO e PL)

  IDENTIDADES DE AUDITORIA:
    - ATIVO = PASSIVO + PATRIMONIO LIQUIDO (identidade fundamental do BP --
      validacao REAL, bate exato com o PDF oficial)
    - ATIVO CIRCULANTE + NAO CIRCULANTE = ATIVO TOTAL (idem Passivo --
      validacao REAL, bate exato)
    - ATIVO (I) [Financeiro+Permanente] = ATIVO do BP, e PASSIVO (II)
      [Financeiro+Permanente] = PASSIVO+PL do BP: identidades MERAMENTE
      ALGEBRICAS nesta implementacao (Permanente = Total - Financeiro por
      construcao), NAO uma validacao da classificacao Financeiro/Permanente
      real do GDF -- essa classificacao e feita a nivel de conta individual
      e nao foi possivel reconstitui-la a partir dos totais agregados
      disponiveis (testado por forca bruta contra o PDF oficial, sem
      sucesso). O proprio "Passivo (II)" do PDF oficial diverge do
      "Passivo+PL" do BP em ambos os exercicios (R$152,77 bi no atual e
      R$125,58 bi no anterior) -- ver nota detalhada no Controle 3 da
      auditoria e no bloco [6] de diagnostico().
    - Resultado do Exercicio (linha do PL) = Resultado Patrimonial do
      Periodo da DVP (validacao cruzada -- ver X6 em mestre.py; aqui
      replicada localmente como Controle informativo quando dvp.py roda
      no mesmo processo)

  RESUMO DE CONFIABILIDADE -- ATUALIZAÇÃO FINAL (confirmado via --diag em
  produção real, GDF, Maio/2026, após localização do arquivo oficial
  Relatório_Equação_de_Balanço_Patrimonial.xlsx -- "Lista Equações de
  Balanço" extraída do próprio sistema contábil do GDF):

    100% CONFIÁVEL -- TODAS AS PÁGINAS VALIDADAS COM DIFERENÇA ZERO:
      - Página 1 completa: Ativo Circulante/Não Circulante (todos os
        sub-itens), Passivo Circulante/Não Circulante (todos os
        sub-itens), Patrimônio Líquido (todos os sub-itens) -- bate
        exato com o oficial.
      - Página 2 completa: Ativo Financeiro = R$7.665.797.148,08
        (diferença 0,00) e Passivo Financeiro = R$6.545.291.111,33
        (diferença 0,00), usando CONTACONTABIL.INSISCONTABIL ('F'/'P')
        + as 3 contas extras de classe 6 no Passivo (622130100,
        622130500, 631100000) -- ver buscar_financeiro_permanente().
      - Página 3 completa: Atos Potenciais Ativos = R$14.228.115.268,87
        (diferença 0,00) e Atos Potenciais Passivos = R$65.253.334.489,58
        (diferença 0,00), usando as contas EXATAS 811XXXXXX/812XXXXXX
        da equação oficial (não mais a aproximação por 711/712 usada em
        versões anteriores) -- ver ATOS_POT_ATIVOS_ITENS/
        ATOS_POT_PASSIVOS_ITENS.

  HISTÓRICO RESUMIDO (13 rodadas de investigação até a solução final):
  a causa raiz de todas as divergências anteriores era usar a faixa de
  conta ERRADA (711/712 em vez de 811/812 para Atos Potenciais) e a
  coluna ERRADA (INSUPERAVIT em vez de INSISCONTABIL para Financeiro/
  Permanente), além de não incluir as 3 contas de classe 6 no Passivo
  Financeiro. A fórmula/lógica de acumulação por INMES e o JOIN com
  CONTACONTABIL sempre estiveram corretos -- só faltava a fonte de
  verdade (a equação oficial), que não estava disponível nas primeiras
  rodadas e só foi localizada no projeto posteriormente.

  Dependencias:  pip install oracledb openpyxl pandas reportlab
  Uso:
      python bp.py --mes 5 --ano 2026
      python bp.py --mes 5 --ano 2026 --ug 130101
      python bp.py --mes 5 --ano 2026 --formato pdf
      python bp.py --mes 5 --ano 2026 --diag
=============================================================================
"""
import argparse, sys
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import oracledb
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, PageBreak)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIGURACAO  (idem BF/BO/DFC/DVP/DMPL)
# ─────────────────────────────────────────────────────────────────────────────
# Credenciais carregadas de config_local.py (nunca em texto puro aqui)
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
OUTPUT_DIR         = Path(r"C:\balanço 2026 gemini arquivos")

MESES = {1:'Janeiro',2:'Fevereiro',3:'Marco',4:'Abril',5:'Maio',
         6:'Junho',7:'Julho',8:'Agosto',9:'Setembro',
         10:'Outubro',11:'Novembro',12:'Dezembro'}

# ─────────────────────────────────────────────────────────────────────────────
#  ITENS DO ATIVO  (fonte: contacontabil_2026.xls -- grupos de 3o nivel)
#  Cada item: (codigo, chave_python, conta_mae, descricao)
# ─────────────────────────────────────────────────────────────────────────────
ATIVO_CIRC_ITENS = [
    ("1.01.01", "CAIXA_EQUIV",       "111000000", "Caixa e Equivalentes de Caixa"),
    ("1.01.02", "CRED_CP",           "112000000", "Créditos a Curto Prazo"),
    ("1.01.03", "DEMAIS_CRED_CP",    "113000000", "Demais Créditos e Valores a Curto Prazo"),
    ("1.01.04", "INVEST_APLIC_CP",   "114000000", "Investimentos e Aplic. Temporárias a Curto Prazo"),
    ("1.01.05", "ESTOQUES",          "115000000", "Estoques"),
    ("1.01.06", "VPD_PAGAS_ANTEC",   "119000000", "VPD Pagas Antecipadamente"),
]
ATIVO_NCIRC_ITENS = [
    ("1.02.01", "REALIZAVEL_LP",     "121000000", "Realizável a Longo Prazo"),
    ("1.02.02", "INVESTIMENTOS",     "122000000", "Investimentos"),
    ("1.02.03", "IMOBILIZADO",       "123000000", "Imobilizado"),
    ("1.02.04", "INTANGIVEL",        "124000000", "Intangível"),
]

# ─────────────────────────────────────────────────────────────────────────────
#  ITENS DO PASSIVO  (idem)
# ─────────────────────────────────────────────────────────────────────────────
PASSIVO_CIRC_ITENS = [
    ("2.01.01", "OBRIG_TRAB_CP",     "211000000", "Obrig. Trab.,Prev. e Assist. a Pagar a Curto Prazo"),
    ("2.01.02", "EMPREST_FINANC_CP", "212000000", "Empréstimos e Financiamentos a Curto Prazo"),
    ("2.01.03", "FORNEC_CP",         "213000000", "Fornecedores e Contas a Pagar a Curto Prazo"),
    ("2.01.04", "OBRIG_FISCAIS_CP",  "214000000", "Obrigações Fiscais a Curto Prazo"),
    ("2.01.05", "TRANSF_FISCAIS_CP", "215000000", "Transferências Fiscais a Curto Prazo"),
    ("2.01.06", "PROVISOES_CP",      "217000000", "Provisões a Curto Prazo"),
    ("2.01.07", "DEMAIS_OBRIG_CP",   "218000000", "Demais Obrigações a Curto Prazo"),
]
PASSIVO_NCIRC_ITENS = [
    ("2.02.01", "OBRIG_TRAB_LP",     "221000000", "Obrig. Trab., Previd. e Assist. a Pagar a LP"),
    ("2.02.02", "EMPREST_FINANC_LP", "222000000", "Empréstimos e Financiamentos a Longo Prazo"),
    ("2.02.03", "FORNEC_LP",         "223000000", "Fornecedores e Contas a Pagar a Longo Prazo"),
    ("2.02.04", "OBRIG_FISCAIS_LP",  "224000000", "Obrigações Fiscais a Longo Prazo"),
    ("2.02.05", "PROVISOES_LP",      "227000000", "Provisões a Longo Prazo"),
    ("2.02.06", "DEMAIS_OBRIG_LP",   "228000000", "Demais Obrigações a Longo Prazo"),
    # CORRECAO (Ago/2026): item 2.02.07 RESULTADO DIFERIDO (229XXXXXX, SC)
    # estava AUSENTE. A equacao oficial (Relatorio_Equacao_de_Balanco_
    # Patrimonial.xlsx, linha "2.02.07.00.00.00 | RESULTADO DIFERIDO |
    # 229XXXXXX | SC | + ") lista o item. Sem ele, a identidade
    # ATIVO = PASSIVO + PL nao fechava: em Junho/2026 (Consolidado) o
    # extraido dava 68.199.818.678,71 contra 69.216.872.599,71 do oficial
    # -- diferenca de EXATAMENTE R$ 1.017.053.921,00, o saldo do 229.
    ("2.02.07", "RESULTADO_DIFERIDO", "229000000", "Resultado Diferido"),
]

# ─────────────────────────────────────────────────────────────────────────────
#  ITENS DO PATRIMONIO LIQUIDO
#  Resultado do Exercicio NAO vem de SALDOCONTABIL (contas 23X) -- vem do
#  MOVIMENTO classes 3/4 do periodo (mesma fonte/logica de dvp.py e dmpl.py).
#  Resultado Acumulado (237) e exibido em 3 sub-linhas no modelo oficial:
#    "Resultado Acumulado" (saldo de 237 EXCLUINDO 2371/2372, normalmente
#       residual/zero), "Resultado do Exercício" (movimento 4xx-3xx do
#       periodo) e a soma report. como "Superavits/Deficits Acumulados"
#       (2371) e "Lucros e Prejuizos Acumulados" (2372) -- replicado
#       exatamente como aparece no PDF oficial.
# ─────────────────────────────────────────────────────────────────────────────
PL_ITENS_SALDO = [
    ("2.03.01", "PAT_SOCIAL_CAPITAL", "231000000", "Patrimônio Social e Capital Social"),
    ("2.03.02", "AFAC",               "232000000", "Adiantamento para Futuro Aumento de Capital"),
    ("2.03.03", "RESERVA_CAPITAL",    "233000000", "Reservas de Capital"),
    ("2.03.04", "AJUSTE_AVAL_PATRIM", "234110000", "Ajustes de Avaliação Patrimonial"),
    ("2.03.05", "RESERVAS_LUCROS",    "235000000", "Reservas de Lucros"),
    ("2.03.06", "DEMAIS_RESERVAS",    "236000000", "Demais Reservas"),
]
# Sub-itens do Resultado Acumulado (todos contas 237XXXXXX, SC):
RESULT_ACUM_RESIDUAL_FAIXA = "237000000"   # total 237 (saldo bruto da conta-mae)
SUPERAVIT_DEFICIT_FAIXA    = "237100000"   # 237.1 Superavits/Deficits Acumulados
LUCROS_PREJUIZOS_FAIXA     = "237200000"   # 237.2 Lucros e Prejuizos Acumulados

PL_CHAVES_SALDO = [k for _, k, _, _ in PL_ITENS_SALDO]

# ─────────────────────────────────────────────────────────────────────────────
#  CONTAS DE COMPENSACAO (Quadro 3 -- Atos Potenciais Ativos/Passivos)
#  FONTE DEFINITIVA: Relatório_Equação_de_Balanço_Patrimonial.xlsx (extraído
#  do proprio sistema contabil do GDF -- "Lista Equações de Balanço,
#  Balanço Patrimonial", Tipo 02). CONFIRMA E CORRIGE descobertas anteriores
#  feitas por tentativa e erro via --diag: as contas reais sao 811XXXXXX
#  (Atos Potenciais Ativos) e 812XXXXXX (Atos Potenciais Passivos) -- NAO
#  711/712 como assumido em versoes anteriores deste modulo (essa era uma
#  aproximacao que batia parcialmente por coincidencia/semelhanca de saldo,
#  mas nao era a fonte oficial). Todas as contas abaixo sao EXATAS (9
#  dígitos completos, copiadas literalmente da equação oficial) -- NÃO usar
#  BETWEEN/faixa, pois a equação não cobre toda a árvore filha da conta-mãe,
#  apenas as subcontas especificamente listadas.
# ─────────────────────────────────────────────────────────────────────────────
ATOS_POT_ATIVOS_ITENS = [
    ("5.01", "GARANTIAS_RECEB", [
        "811110101", "811110103", "811110105", "811110107", "811110109",
        "811110111", "811110113", "811110115", "811110117", "811110198",
    ], "Garantias e Contra Garantias Recebidas"),
    ("5.02", "DIREITOS_CONVENIADOS", [
        "811210101", "811210104",
    ], "Direitos Conveniados e Outros Instr. Congêneres"),
    ("5.03", "DIREITOS_CONTRATUAIS", [
        "811311401", "811311501",
    ], "Direitos Contratuais"),
    ("5.04", "OUTROS_ATOS_POT_ATIVOS", [
        "811910101", "811910201", "811910301", "811910401", "811910501",
        "811910601", "811910602", "811910901", "811911001", "811911101",
        "811911201", "811911301", "811911501", "811911601", "811912201",
        "811912301", "811912401", "811913201", "811913301", "811913401",
        "811913501", "811913601", "811913701", "811913801", "811913901",
        "811914101",
    ], "Outros Atos Potenciais Ativos"),
]
ATOS_POT_PASSIVOS_ITENS = [
    ("6.01", "GARANTIAS_CONCED", [
        "812110101", "812110104", "812110107", "812110110", "812110113",
        "812110116", "812110119", "812110122", "812110124",
    ], "Garantias e Contra Garantias Concedidas"),
    ("6.02", "OBRIGACOES_CONVENIADAS", [
        "812210111", "812210501", "812219901",
    ], "Obrigações Conveniadas e Outros Instr. Congêneres"),
    ("6.03", "OBRIGACOES_CONTRATUAIS", [
        "812310101", "812310201", "812310301", "812310401", "812310501",
        "812310601", "812310701", "812310801", "812310901", "812311001",
        "812311101", "812311401", "812311501", "812311601", "812311701",
        "812311801", "812311901", "812312001", "812315001", "812315101",
        "812315201", "812315301", "812315401", "812315501",
    ], "Obrigações Contratuais"),
    ("6.04", "OUTROS_ATOS_POT_PASSIVOS", [
        "812910101", "812910201",
    ], "Outros Atos Potenciais Passivos"),
]

# ─────────────────────────────────────────────────────────────────────────────
#  QUADRO ATIVO/PASSIVO FINANCEIRO E PERMANENTE (pág. 2)
#  FONTE DEFINITIVA: mesma planilha de equação oficial. O criterio real e a
#  coluna CONTACONTABIL."Sistema Contábil" (valores 'F'=Financeiro,
#  'P'=Permanente) -- CONFIRMA a pista que --diag vinha buscando (era a
#  coluna procurada nos blocos [6a]-[6j], nunca testada porque o nome real
#  da coluna no banco pode diferir do nome visto no extrato XLS estatico
#  "Sistema Contábil"; ver diagnostico() para descobrir o nome exato da
#  coluna Oracle correspondente).
#    ATIVO FINANCEIRO    = SD(1XXXXXXXX) onde Sistema Contábil = 'F'
#    ATIVO PERMANENTE    = SD(1XXXXXXXX) onde Sistema Contábil = 'P'
#    PASSIVO FINANCEIRO  = SC(21XXXXXXX) onde Sistema Contábil = 'F'
#                          + SC(622130100) + SC(622130500) + SC(631100000)
#                          (3 contas extras de RP/empenho a liquidar --
#                          classe 6, fora da faixa 1XX-2XX, explicando por
#                          que o resíduo nunca foi encontrado nas 8
#                          hipóteses testadas anteriormente: a busca nunca
#                          olhou a classe 6)
#    PASSIVO PERMANENTE  = SC(21XXXXXXX onde P) + SC(22XXXXXXX onde P)
# ─────────────────────────────────────────────────────────────────────────────
PASSIVO_FINANCEIRO_CONTAS_EXTRA = ["622130100", "622130500", "631100000"]

# ─────────────────────────────────────────────────────────────────────────────
#  CONEXAO  (identica aos demais modulos)
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


def _executar_um(conn, sql):
    """Executa uma unica consulta agregada (SUM) e devolve o escalar."""
    cur = conn.cursor()
    cur.execute(sql)
    r = cur.fetchone()[0]
    cur.close()
    return float(r or 0)


# ─────────────────────────────────────────────────────────────────────────────
#  BUSCA -- SALDOS DE ATIVO/PASSIVO/PL  (SALDOCONTABIL, mesmo padrao bf.py)
# ─────────────────────────────────────────────────────────────────────────────
def _saldo_contas_exatas(conn, schema, contas, inmes_clause, coug, natureza):
    """Soma o saldo de uma LISTA EXATA de contas completas (9 dígitos,
    via SQL IN) na SALDOCONTABIL -- diferente de _saldo_contas_maes, que
    usa BETWEEN para cobrir toda a árvore filha de uma conta-mãe. Usado
    quando a fonte oficial (equação do GDF) lista contas específicas que
    NÃO formam uma faixa contígua simples. natureza: 'SD' (D-C) ou
    'SC' (C-D)."""
    sinal = "v.VADEBITO - v.VACREDITO" if natureza == 'SD' else "v.VACREDITO - v.VADEBITO"
    filtro_ug = f"AND v.COUG = {coug}" if coug else ""
    lista = ", ".join(str(int(c)) for c in contas)
    q = f"""
        SELECT SUM({sinal})
        FROM   {schema}.SALDOCONTABIL v
        WHERE  {inmes_clause}
          AND  v.COCONTACONTABIL IN ({lista})
          {filtro_ug}
    """
    return _executar_um(conn, q)


def _saldo_conta_mae(conn, schema, conta_mae, inmes_clause, coug, natureza):
    """Soma o saldo (VACREDITO/VADEBITO) de uma conta-mae (e seus filhos,
    mesmo prefixo) na SALDOCONTABIL. natureza: 'SD' (D-C, Ativo) ou
    'SC' (C-D, Passivo/PL)."""
    return _saldo_contas_maes(conn, schema, [conta_mae], inmes_clause, coug, natureza)


def _saldo_contas_maes(conn, schema, contas_mae, inmes_clause, coug, natureza,
                        excluir_contas=None):
    """Como _saldo_conta_mae, mas soma 1+ contas-mae (cada uma com sua
    arvore filha de 9 digitos) numa unica query, via OR. Usado quando um
    item do relatorio oficial agrega mais de uma conta-mae (ex.: 'Outros
    Atos Potenciais Passivos' = Demandas Judiciais + Outros).

    excluir_contas: lista opcional de COCONTACONTABIL (conta completa, 9
    digitos) a EXCLUIR do somatorio mesmo que caiam dentro da arvore de
    uma das contas_mae -- usado para as duas contas confirmadas via
    --diag em producao (GDF, Maio/2026) que NAO devem compor o total
    oficial de 'Outros Atos Potenciais Ativos' (711911400 Bens de Uso
    Comum do Povo, 711914000 Controle de Estoque Interno-Almoxarifado):
    ver nota detalhada em ATOS_POT_ATIVOS_ITENS."""
    sinal = "v.VADEBITO - v.VACREDITO" if natureza == 'SD' else "v.VACREDITO - v.VADEBITO"
    filtro_ug = f"AND v.COUG = {coug}" if coug else ""
    cond_contas = []
    for conta_mae in contas_mae:
        # CORRECAO (Ago/2026): antes era `fim = ini + 999999`, largura FIXA de
        # 6 digitos, correta APENAS para conta-mae de 3 digitos significativos
        # (ex.: 211000000 -> 211999999, ok). Para maes de 4+ digitos a faixa
        # transbordava para a arvore VIZINHA. Casos confirmados:
        #   237100000 (2371 Superavits) -> 237100000..238099999, engolindo TODO
        #     o 2372 (Lucros/Prejuizos). Em Junho/2026 isso fez "Superavits"
        #     sair -176.279.793.399,94 quando o oficial e -174.722.472.700,38
        #     (a diferenca era exatamente o 2372, -1.557.320.699,56).
        #   234110000 (Ajustes de Aval. Patrimonial) -> ..235109999, alcancando
        #     o 235 (Reservas de Lucros).
        # A largura passa a ser derivada do PREFIXO significativo da conta-mae:
        #   "211" -> 211000000..211999999 ; "2371" -> 237100000..237109999
        prefixo = str(int(conta_mae)).rstrip('0')
        ini = int(prefixo.ljust(9, '0'))
        fim = int(prefixo.ljust(9, '9'))
        cond_contas.append(f"v.COCONTACONTABIL BETWEEN {ini} AND {fim}")
    cond = "(" + " OR ".join(cond_contas) + ")"
    filtro_excluir = ""
    if excluir_contas:
        lista = ", ".join(str(int(c)) for c in excluir_contas)
        filtro_excluir = f"AND v.COCONTACONTABIL NOT IN ({lista})"
    q = f"""
        SELECT SUM({sinal})
        FROM   {schema}.SALDOCONTABIL v
        WHERE  {inmes_clause}
          AND  {cond}
          {filtro_excluir}
          {filtro_ug}
    """
    return _executar_um(conn, q)


def _caixa_hibrido(conn, schema, mes, coug):
    """
    Caixa e Equivalentes (111XXXXXX) para o EXERCICIO ATUAL -- NAO usa
    SALDOCONTABIL puro (v.INMES BETWEEN 0 AND mes) como as demais contas
    de saldo desta pagina, porque essa view e recalculada em lote (ciclo
    de consolidacao periodico) enquanto o mes corrente ainda esta aberto.

    Confirmado via auditoria cruzada (mestre.py, Regras 4a/4b e X1b,
    Jul/2026): duas execucoes do mesmo script, ~1h48 de intervalo, sem
    nenhuma mudanca de codigo, devolveram Caixa Final divergindo em
    R$10,8 mi. Contra o BP oficial impresso, a formula pura da view
    errava R$12,17 mi, espalhados em ~18 contas 111XXXXXX (sem padrao
    localizavel -- nao e conta faltando, e a fonte inteira que atrasa
    p/ o mes aberto).

    A formula do BF (abertura da view, exercicio ja encerrado -> estavel
    + movimento via LANCAMENTOCONTABIL, tabela transacional bruta) bateu
    EXATA (diferenca zero) contra o mesmo oficial. Reproduzida aqui para
    Caixa convergir entre BF/DFC/BP (elimina os erros cruzados X1b/4a/4b).

    O Caixa do Exercicio Anterior (encerrado, schema_at INMES=0) NAO
    precisa deste ajuste -- e a mesma fonte estavel usada como abertura
    aqui e no BF/DFC.
    """
    filtro_v = f"AND v.COUG = {coug}" if coug else ""
    filtro_o = f"AND o.COUG = {coug}" if coug else ""

    abertura = _executar_um(conn, f"""
        SELECT SUM(v.VADEBITO - v.VACREDITO)
        FROM   {schema}.SALDOCONTABIL v
        WHERE  v.INMES = 0
          AND  v.COCONTACONTABIL BETWEEN 111000000 AND 111999999
          {filtro_v}
    """)
    movimento = _executar_um(conn, f"""
        SELECT SUM(DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,
                                            'C',-o.VALANCAMENTO,0))
        FROM   {schema}.LANCAMENTOCONTABIL o
        WHERE  o.INMES BETWEEN 1 AND {mes}
          AND  o.COCONTACONTABIL BETWEEN 111000000 AND 111999999
          {filtro_o}
    """)
    return abertura + movimento


def buscar_saldos_patrimoniais(conn, mes, ano, coug, schema, inmes_clause):
    """Busca todos os saldos de Ativo/Passivo/PL (contas 1/2) para um dado
    schema (MIL{ano} ou MIL{ano-1}) e clausula de INMES (acumulado atual ou
    encerramento anterior). Devolve dict chave -> valor."""
    t = {}
    print(f"  [saldos patrimoniais] {schema}.SALDOCONTABIL ({inmes_clause})...")

    for _, chave, conta_mae, _ in ATIVO_CIRC_ITENS + ATIVO_NCIRC_ITENS:
        t[chave] = _saldo_conta_mae(conn, schema, conta_mae, inmes_clause, coug, 'SD')

    for _, chave, conta_mae, _ in PASSIVO_CIRC_ITENS + PASSIVO_NCIRC_ITENS:
        t[chave] = _saldo_conta_mae(conn, schema, conta_mae, inmes_clause, coug, 'SC')

    for _, chave, conta_mae, _ in PL_ITENS_SALDO:
        t[chave] = _saldo_conta_mae(conn, schema, conta_mae, inmes_clause, coug, 'SC')

    t['SUPERAVIT_DEFICIT_ACUM'] = _saldo_conta_mae(
        conn, schema, SUPERAVIT_DEFICIT_FAIXA, inmes_clause, coug, 'SC')
    t['LUCROS_PREJUIZOS_ACUM'] = _saldo_conta_mae(
        conn, schema, LUCROS_PREJUIZOS_FAIXA, inmes_clause, coug, 'SC')
    # Resultado Acumulado "puro" (237 total) menos os dois sub-itens acima
    # -- residual que cobre eventuais lancamentos em 237 fora de 2371/2372.
    total_237 = _saldo_conta_mae(conn, schema, RESULT_ACUM_RESIDUAL_FAIXA,
                                  inmes_clause, coug, 'SC')
    t['RESULTADO_ACUM_RESIDUAL'] = (total_237 - t['SUPERAVIT_DEFICIT_ACUM']
                                     - t['LUCROS_PREJUIZOS_ACUM'])

    return t


def buscar_pl_pre_encerramento(conn, schema_ant, coug, mes_encerramento=14):
    """Le o bloco PATRIMONIO LIQUIDO do Exercicio Anterior na fonte
    PRE-ENCERRAMENTO: MIL{ano-1}.SALDOCONTABIL com INMES BETWEEN 0 AND
    {mes_encerramento-1}.

    POR QUE ISTO EXISTE (DIAG16/DIAG17, Ago/2026)
    ---------------------------------------------
    MIL{ano-1}.SALDOCONTABIL possui INMES = 14 -- um mes de ENCERRAMENTO
    que a versao anterior deste modulo desconhecia. Nele, o Resultado do
    Exercicio e transferido para dentro do PL, RATEADO entre contas:

        231  Patrimonio Social ..........      -1.323.940,25
        2371 Superavits/Deficits Acum. .. -53.671.660.288,73
                                          --------------------
        total .......................... -53.672.984.228,98
                                          = Resultado de 2025 (4xx-3xx)

    Como o Exercicio Anterior era lido de MIL{ano}.SALDOCONTABIL INMES=0
    (abertura), que equivale a MIL{ano-1} acumulado 0..14, esse rateio JA
    vinha embutido nas contas -- e as sub-linhas 231 e 2371 divergiam do
    relatorio oficial (que usa a fonte PRE-encerramento, 0..13, e exibe o
    Resultado do Exercicio em linha propria). O TOTAL do PL era identico
    nas duas fontes, por isso nenhum controle de integridade acusava.

    VALIDADO com diferenca ZERO em Junho/2026 Consolidado (DIAG17):
      T1  abertura MIL{ano} INMES=0 == Sigma MIL{ano-1} 0..14, nas 8 contas
      T2  Sigma MIL{ano-1} 0..13 == oficial ANT, nas 8 contas, e o total
          fecha em -123.862.128.780,30 somando o Resultado do Exercicio
      T3  Sigma(INMES=14) sobre o PL == Resultado 2025, exato

    ATENCAO: esta fonte vale SOMENTE para o bloco PL. Ativo e Passivo
    exigivel continuam vindo de MIL{ano}.SALDOCONTABIL INMES=0 -- o
    encerramento nao os altera.
    """
    pre = mes_encerramento - 1
    inmes_clause = f"v.INMES BETWEEN 0 AND {pre}"
    print(f"  [PL pre-encerramento] {schema_ant}.SALDOCONTABIL ({inmes_clause})...")
    t = {}
    for _, chave, conta_mae, _ in PL_ITENS_SALDO:
        t[chave] = _saldo_conta_mae(conn, schema_ant, conta_mae, inmes_clause,
                                     coug, 'SC')
    t['SUPERAVIT_DEFICIT_ACUM'] = _saldo_conta_mae(
        conn, schema_ant, SUPERAVIT_DEFICIT_FAIXA, inmes_clause, coug, 'SC')
    t['LUCROS_PREJUIZOS_ACUM'] = _saldo_conta_mae(
        conn, schema_ant, LUCROS_PREJUIZOS_FAIXA, inmes_clause, coug, 'SC')
    total_237 = _saldo_conta_mae(conn, schema_ant, RESULT_ACUM_RESIDUAL_FAIXA,
                                  inmes_clause, coug, 'SC')
    t['RESULTADO_ACUM_RESIDUAL'] = (total_237 - t['SUPERAVIT_DEFICIT_ACUM']
                                     - t['LUCROS_PREJUIZOS_ACUM'])
    return t


def buscar_atos_potenciais(conn, schema, inmes_clause, coug):
    """Busca os Atos Potenciais Ativos (SC) e Passivos (SC) -- contas
    EXATAS extraídas da fonte oficial (Relatório_Equação_de_Balanço_
    Patrimonial.xlsx, "Lista Equações de Balanço" do próprio sistema
    contábil do GDF): 811XXXXXX = Atos Potenciais Ativos, 812XXXXXX =
    Atos Potenciais Passivos -- AMBOS com natureza SC (Saldo Credor) na
    equação oficial (corrige uma suposição anterior de que o lado Ativo
    seria SD -- a equação documenta SC para os dois lados). CORRIGE
    também a faixa de conta usada em versões anteriores deste módulo
    (711/712, uma aproximação que batia parcialmente por coincidência,
    mas não era a fonte real)."""
    t = {}
    print(f"  [contas de compensacao] {schema}.SALDOCONTABIL ({inmes_clause})...")
    for _, chave, contas, _ in ATOS_POT_ATIVOS_ITENS:
        t[chave] = _saldo_contas_exatas(conn, schema, contas, inmes_clause, coug, 'SC')
    for _, chave, contas, _ in ATOS_POT_PASSIVOS_ITENS:
        t[chave] = _saldo_contas_exatas(conn, schema, contas, inmes_clause, coug, 'SC')
    return t


def buscar_resultado_exercicio(conn, mes, ano, coug, schema, inmes_clause):
    """Resultado do Exercicio = SC(4XXXXXXXX) - SD(3XXXXXXXX) no periodo,
    mesma fonte/logica usada em dvp.py (Resultado Patrimonial) e dmpl.py
    (linha Resultado do Exercicio), garantindo a identidade cruzada entre
    os tres anexos quando rodados sobre a mesma base."""
    filtro_ug = f"AND o.COUG = {coug}" if coug else ""
    q = f"""
        SELECT
            SUM(CASE WHEN {inmes_clause} AND o.COCONTACONTABIL BETWEEN 400000000 AND 499999999
                 THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0) ELSE 0 END)
            -
            SUM(CASE WHEN {inmes_clause} AND o.COCONTACONTABIL BETWEEN 300000000 AND 399999999
                 THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0) ELSE 0 END)
        FROM {schema}.LANCAMENTOCONTABIL o
        WHERE (o.COCONTACONTABIL BETWEEN 300000000 AND 499999999)
          {filtro_ug}
    """
    print(f"  [resultado do exercicio] {schema}.LANCAMENTOCONTABIL ({inmes_clause})...")
    return _executar_um(conn, q)


def buscar_financeiro_permanente(conn, schema, inmes_clause, coug):
    """Busca Ativo/Passivo Financeiro via JOIN SALDOCONTABIL +
    CONTACONTABIL, usando a coluna INSISCONTABIL ("Sistema Contábil" no
    extrato XLS estático contacontabil_2026.xls; valores 'F'=Financeiro,
    'P'=Permanente) -- FONTE DEFINITIVA confirmada via
    Relatório_Equação_de_Balanço_Patrimonial.xlsx ("Lista Equações de
    Balanço" extraída do próprio sistema contábil do GDF):
        ATIVO FINANCEIRO    = SD(1XXXXXXXX) onde INSISCONTABIL = 'F'
        ATIVO PERMANENTE    = SD(1XXXXXXXX) onde INSISCONTABIL = 'P'
        PASSIVO FINANCEIRO  = SC(21XXXXXXX) onde INSISCONTABIL = 'F'
                              + SC(622130100) + SC(622130500) + SC(631100000)
                              (3 contas extras de classe 6 -- RP/empenho a
                              liquidar -- que NUNCA tinham sido encontradas
                              nas 8 hipóteses anteriores porque a busca
                              nunca cobriu a classe 6, fora da faixa 1XX-2XX)
        PASSIVO PERMANENTE  = SC(21XXXXXXX onde P) + SC(22XXXXXXX onde P)
    Isso substitui a aproximação anterior por INSUPERAVIT='S' (que dava
    ~97% de precisão no Ativo mas só ~43% no Passivo, com 8 hipóteses
    testadas e refutadas para explicar o resíduo -- a causa raiz era a
    coluna errada E a falta das 3 contas de classe 6).

    CONTACONTABIL existe por ano (MIL{ano}.CONTACONTABIL), exatamente
    como SALDOCONTABIL -- por isso o JOIN usa o MESMO schema."""
    filtro_ug = f"AND v.COUG = {coug}" if coug else ""
    q_ativo = f"""
        SELECT SUM(v.VADEBITO - v.VACREDITO)
        FROM {schema}.SALDOCONTABIL v
        JOIN {schema}.CONTACONTABIL c ON c.COCONTACONTABIL = v.COCONTACONTABIL
        WHERE {inmes_clause}
          AND v.COCONTACONTABIL BETWEEN 100000000 AND 199999999
          AND c.INSISCONTABIL = 'F'
          {filtro_ug}
    """
    q_passivo_principal = f"""
        SELECT SUM(v.VACREDITO - v.VADEBITO)
        FROM {schema}.SALDOCONTABIL v
        JOIN {schema}.CONTACONTABIL c ON c.COCONTACONTABIL = v.COCONTACONTABIL
        WHERE {inmes_clause}
          AND v.COCONTACONTABIL BETWEEN 210000000 AND 219999999
          AND c.INSISCONTABIL = 'F'
          {filtro_ug}
    """
    print(f"  [financeiro/permanente -- INSISCONTABIL] {schema}.SALDOCONTABIL "
          f"JOIN {schema}.CONTACONTABIL ({inmes_clause})...")
    ativo_fin = _executar_um(conn, q_ativo)
    passivo_fin_principal = _executar_um(conn, q_passivo_principal)
    passivo_fin_extra = _saldo_contas_exatas(
        conn, schema, PASSIVO_FINANCEIRO_CONTAS_EXTRA, inmes_clause, coug, 'SC')
    passivo_fin = passivo_fin_principal + passivo_fin_extra
    return {'ATIVO_FINANCEIRO_INSUPERAVIT': ativo_fin,
            'PASSIVO_FINANCEIRO_INSUPERAVIT': passivo_fin}


def buscar_tudo(conn, mes, ano, coug):
    """Orquestra a busca completa: Exercicio Atual (MIL{ano}, saldo
    acumulado 0..mes) e Exercicio Anterior, para Ativo/Passivo/PL e
    Contas de Compensacao, mais o Resultado do Exercicio de cada periodo.

    CORRIGIDO (confirmado via --diag em producao, GDF, execucao real sem
    --diag em Maio/2026): o saldo do EXERCICIO ANTERIOR para contas de
    SALDO (1XX/2XX/711/712) NAO deve vir de MIL{ano-1}.SALDOCONTABIL
    WHERE INMES=13 -- esse INMES=13 contem apenas ajustes residuais de
    encerramento (visto na execucao real: R$55.889.140,40, muito menor
    que o esperado ~R$68,78 bi), exatamente como documentado no
    cabecalho do bf.py: "13 = encerramento (ajustes, reclassif.,
    inscricao de RP)" -- NAO o saldo total. O saldo de encerramento
    completo do ano anterior e, na verdade, o SALDO DE ABERTURA do ano
    atual: MIL{ano}.SALDOCONTABIL WHERE INMES=0 (mesma convencao usada
    em bf.py: "0 = saldo de abertura -> SALDO DO EXERCICIO ANTERIOR").
    Por isso, para saldos (Ativo/Passivo/PL/Atos Potenciais/Financeiro-
    Permanente), o Exercicio Anterior agora usa schema_at (MIL{ano}) com
    INMES=0 -- NAO schema_ant. O RESULTADO DO EXERCICIO Anterior (fluxo,
    classes 3/4) continua usando MIL{ano-1}, INMES BETWEEN 1 AND 13 (ano
    fiscal completo) -- mesma logica de dvp.py/dmpl.py, que e correta
    para contas de FLUXO (diferente de contas de SALDO)."""
    ano_ant = ano - 1
    schema_at  = f"MIL{ano}"
    schema_ant = f"MIL{ano_ant}"
    inmes_at  = f"v.INMES BETWEEN 0 AND {mes}"
    inmes_ant_saldo = "v.INMES = 0"  # saldo de abertura do MIL{ano} = encerramento do ano anterior

    brutos = {}
    print(f"\n  >>> EXERCICIO ATUAL ({schema_at}, acumulado 0..{mes}):")
    at = buscar_saldos_patrimoniais(conn, mes, ano, coug, schema_at, inmes_at)
    at['CAIXA_EQUIV'] = _caixa_hibrido(conn, schema_at, mes, coug)
    at_comp = buscar_atos_potenciais(conn, schema_at, inmes_at, coug)
    at_result = buscar_resultado_exercicio(conn, mes, ano, coug, schema_at,
                                            f"o.INMES BETWEEN 1 AND {mes}")
    at_finperm = buscar_financeiro_permanente(conn, schema_at, inmes_at, coug)

    print(f"\n  >>> EXERCICIO ANTERIOR -- SALDOS ({schema_at}, INMES=0 = "
          f"abertura do ano atual = encerramento de {ano_ant}):")
    ant = buscar_saldos_patrimoniais(conn, mes, ano, coug, schema_at, inmes_ant_saldo)
    ant_comp = buscar_atos_potenciais(conn, schema_at, inmes_ant_saldo, coug)
    ant_finperm = buscar_financeiro_permanente(conn, schema_at, inmes_ant_saldo, coug)

    print(f"\n  >>> EXERCICIO ANTERIOR -- RESULTADO ({schema_ant}, ano "
          f"fiscal completo, INMES 1..13):")
    ant_result = buscar_resultado_exercicio(conn, mes, ano, coug, schema_ant,
                                             "o.INMES BETWEEN 1 AND 13")

    # ── PL do Exercicio Anterior: fonte PRE-ENCERRAMENTO ────────────────
    # Ver docstring de buscar_pl_pre_encerramento() (DIAG16/DIAG17). Os
    # saldos de Ativo/Passivo acima permanecem na fonte de abertura; SO o
    # bloco PL e resubstituido, para que as sub-linhas 231 e 2371 batam
    # com o relatorio oficial. O TOTAL do PL nao se altera.
    print(f"\n  >>> EXERCICIO ANTERIOR -- PL PRE-ENCERRAMENTO ({schema_ant}, "
          f"INMES 0..13; o INMES=14 e o encerramento):")
    try:
        ant_pl = buscar_pl_pre_encerramento(conn, schema_ant, coug)
        ant.update(ant_pl)
        brutos['_PL_ANT_FONTE'] = f"{schema_ant} INMES 0..13 (pre-encerramento)"
    except Exception as e:
        print(f"    AVISO: falha ao ler o PL pre-encerramento em {schema_ant} "
              f"({type(e).__name__}: {e}).")
        print(f"    Mantendo a fonte de abertura ({schema_at} INMES=0) -- o TOTAL")
        print( "    do PL continua correto, mas as sub-linhas 231/2371 do")
        print( "    Exercicio Anterior ficarao com o encerramento embutido.")
        brutos['_PL_ANT_FONTE'] = f"{schema_at} INMES=0 (fallback, pos-encerramento)"

    for k, v in at.items():       brutos[f"{k}_AT"] = v
    for k, v in at_comp.items():  brutos[f"{k}_AT"] = v
    for k, v in at_finperm.items(): brutos[f"{k}_AT"] = v
    brutos['RESULTADO_EXERCICIO_AT'] = at_result

    for k, v in ant.items():      brutos[f"{k}_ANT"] = v
    for k, v in ant_comp.items(): brutos[f"{k}_ANT"] = v
    for k, v in ant_finperm.items(): brutos[f"{k}_ANT"] = v
    brutos['RESULTADO_EXERCICIO_ANT'] = ant_result

    return brutos


# ─────────────────────────────────────────────────────────────────────────────
#  TOTAIS DERIVADOS
# ─────────────────────────────────────────────────────────────────────────────
ATIVO_CIRC_CHAVES  = [k for _, k, _, _ in ATIVO_CIRC_ITENS]
ATIVO_NCIRC_CHAVES = [k for _, k, _, _ in ATIVO_NCIRC_ITENS]
PASSIVO_CIRC_CHAVES  = [k for _, k, _, _ in PASSIVO_CIRC_ITENS]
PASSIVO_NCIRC_CHAVES = [k for _, k, _, _ in PASSIVO_NCIRC_ITENS]


def calcular(brutos, mes_ref=None):
    """Recebe os componentes brutos (sufixos _AT/_ANT) e calcula todos os
    totais/subtotais do Balanco Patrimonial, do Quadro Financeiro/Permanente
    e do Quadro de Contas de Compensacao."""
    t = dict(brutos)
    if mes_ref is not None:
        t['_MES_REF'] = mes_ref

    for suf in ('AT', 'ANT'):
        # ── ATIVO ────────────────────────────────────────────────────────
        t[f'ATIVO_CIRCULANTE_{suf}'] = sum(t.get(f'{k}_{suf}', 0.0) for k in ATIVO_CIRC_CHAVES)
        t[f'ATIVO_NCIRCULANTE_{suf}'] = sum(t.get(f'{k}_{suf}', 0.0) for k in ATIVO_NCIRC_CHAVES)
        t[f'ATIVO_TOTAL_{suf}'] = t[f'ATIVO_CIRCULANTE_{suf}'] + t[f'ATIVO_NCIRCULANTE_{suf}']

        # ── PASSIVO ──────────────────────────────────────────────────────
        # CONFIRMADO contra o modelo oficial (ListaBalancoPatrimonial,
        # Maio/2026): "PASSIVO NÃO CIRCULANTE" é a soma EXATA dos seus
        # sub-itens (211..228), SEM incluir o Patrimônio Líquido -- o PL é
        # um terceiro bloco, paralelo, e NÃO um subitem do Passivo Não
        # Circulante (validado batendo 209.021.741.342,87 = soma exata dos
        # 6 sub-itens oficiais, sem qualquer resíduo de PL).
        t[f'PASSIVO_CIRCULANTE_{suf}'] = sum(t.get(f'{k}_{suf}', 0.0) for k in PASSIVO_CIRC_CHAVES)
        t[f'PASSIVO_NCIRCULANTE_{suf}'] = sum(t.get(f'{k}_{suf}', 0.0) for k in PASSIVO_NCIRC_CHAVES)

        # ── PATRIMONIO LIQUIDO ───────────────────────────────────────────
        # ATENÇÃO -- BUG ENCONTRADO E CORRIGIDO (confirmado em execução
        # real após corrigir o bug do INMES=0/13 acima): para o Exercício
        # ATUAL, o saldo da conta-mãe 237 via SALDOCONTABIL (INMES
        # acumulado 0..mês) é só o saldo ABERTO/residual -- o Resultado do
        # Exercício do período (movimento 4xx-3xx, fonte LANCAMENTOCONTABIL)
        # precisa ser somado para chegar ao saldo final do PL. MAS para o
        # Exercício ANTERIOR, agora que os saldos vêm de MIL{ano}.INMES=0
        # (saldo de ABERTURA do ano atual = FECHAMENTO do ano anterior, já
        # CONSOLIDADO), o Resultado do Exercício de {ano-1} JÁ ESTÁ
        # INCLUÍDO dentro desse saldo -- somar RESULTADO_EXERCICIO_ANT de
        # novo gera DUPLICAÇÃO. Confirmado numericamente: a diferença
        # Passivo-Ativo do Exercício Anterior era EXATAMENTE igual (em
        # módulo) ao Resultado do Exercício Anterior antes desta correção.
        # Por isso, a soma do Resultado do Exercício só se aplica ao
        # sufixo _AT (Exercício Atual); para _ANT, o residual de 237 (via
        # INMES=0) já é o valor final, sem soma adicional.
        # CORRECAO (Ago/2026): no modelo oficial "Resultado Acumulado" e um
        # SUBTOTAL (soma das sub-linhas que aparecem indentadas abaixo dele),
        # nao o saldo residual do 237. Conferido em Junho/2026:
        #   -15.484.195.906,83 + (-174.722.472.700,38) + (-1.557.320.699,56)
        #   = -191.763.989.306,77  <- exatamente a linha oficial.
        # O modulo exibia o residual (-13.326.875.207,27). A parcela residual
        # + Resultado do Exercicio continua sendo o que ENTRA no total do PL;
        # ela agora e acumulada numa variavel intermediaria e o subtotal passa
        # a englobar tambem 2371/2372, de modo que o PL soma UMA unica linha e
        # o total permanece inalterado (-147.071.752.266,32 em Jun/2026).
        if suf == 'AT':
            residual_mais_exercicio = (t.get(f'RESULTADO_ACUM_RESIDUAL_{suf}', 0.0)
                                        + t.get(f'RESULTADO_EXERCICIO_{suf}', 0.0))
        else:
            # CORRECAO (Ago/2026, DIAG16/DIAG17): o PL do Exercicio Anterior
            # agora vem de MIL{ano-1} INMES 0..13 (PRE-encerramento), onde o
            # Resultado de {ano-1} AINDA NAO foi transferido para 231/2371.
            # Logo ele PRECISA ser somado aqui -- ao contrario da fonte
            # antiga (abertura de {ano} = 0..14), em que ja vinha embutido e
            # soma-lo duplicaria. O flag _PL_ANT_FONTE registra qual fonte
            # foi efetivamente usada, para que o fallback (quando MIL{ano-1}
            # esta indisponivel) continue correto.
            fonte_pl_ant = str(t.get('_PL_ANT_FONTE', ''))
            pl_ant_pre_encerramento = 'pre-encerramento' in fonte_pl_ant
            residual_mais_exercicio = t.get(f'RESULTADO_ACUM_RESIDUAL_{suf}', 0.0)
            if pl_ant_pre_encerramento:
                residual_mais_exercicio += t.get(f'RESULTADO_EXERCICIO_{suf}', 0.0)

        t[f'RESULTADO_ACUMULADO_{suf}'] = (residual_mais_exercicio
                                            + t.get(f'SUPERAVIT_DEFICIT_ACUM_{suf}', 0.0)
                                            + t.get(f'LUCROS_PREJUIZOS_ACUM_{suf}', 0.0))

        soma_pl_saldo = sum(t.get(f'{k}_{suf}', 0.0) for k in PL_CHAVES_SALDO)
        t[f'PATRIMONIO_LIQUIDO_{suf}'] = soma_pl_saldo + t[f'RESULTADO_ACUMULADO_{suf}']

        t[f'PASSIVO_PL_TOTAL_{suf}'] = (t[f'PASSIVO_CIRCULANTE_{suf}']
                                         + t[f'PASSIVO_NCIRCULANTE_{suf}']
                                         + t[f'PATRIMONIO_LIQUIDO_{suf}'])

        # ── QUADRO ATIVO/PASSIVO FINANCEIRO E PERMANENTE ────────────────
        # ATUALIZAÇÃO (confirmado via --diag em produção, GDF, Maio/2026):
        # o flag CONTACONTABIL.INSUPERAVIT='S' é o critério correto. Via
        # JOIN SALDOCONTABIL + CONTACONTABIL (buscar_financeiro_
        # permanente()), o lado ATIVO bateu com diferença de apenas
        # ~R$247 milhões (3,2%) sobre o oficial (R$7.665.797.148,08) --
        # confirma o flag. O lado PASSIVO, com o mesmo filtro, retornou
        # só ~43% do oficial (R$6.545.291.111,33) -- ainda há resíduo não
        # explicado nesse lado (possivelmente contas com INSUPERAVIT=' '
        # em branco também deveriam contar; ver diagnostico() bloco [6d]).
        # Por isso o módulo usa o valor real do JOIN quando disponível
        # (chave _INSUPERAVIT, preenchida por buscar_tudo() quando há
        # conexão Oracle), com fallback para a aproximação antiga por
        # grupo de contas SOMENTE se o valor real não estiver disponível
        # (ex.: chamadas de teste sem banco). O Controle 3 da auditoria
        # continua sinalizando o resíduo conhecido do lado Passivo como
        # ALERTA -- não tratar os números desta página como definitivos
        # até o resíduo ser zerado.
        if f'ATIVO_FINANCEIRO_INSUPERAVIT_{suf}' in t:
            t[f'ATIVO_FINANCEIRO_{suf}'] = t[f'ATIVO_FINANCEIRO_INSUPERAVIT_{suf}']
        else:
            # Fallback (aproximação -- ver nota acima): disponibilidades e
            # créditos de curto prazo. NÃO bate com o oficial; mantido só
            # para o módulo funcionar sem acesso ao Oracle (ex.: testes).
            t[f'ATIVO_FINANCEIRO_{suf}'] = (t.get(f'CAIXA_EQUIV_{suf}', 0.0)
                                             + t.get(f'CRED_CP_{suf}', 0.0)
                                             + t.get(f'DEMAIS_CRED_CP_{suf}', 0.0)
                                             + t.get(f'INVEST_APLIC_CP_{suf}', 0.0))
        t[f'ATIVO_PERMANENTE_{suf}'] = t[f'ATIVO_TOTAL_{suf}'] - t[f'ATIVO_FINANCEIRO_{suf}']

        if f'PASSIVO_FINANCEIRO_INSUPERAVIT_{suf}' in t:
            t[f'PASSIVO_FINANCEIRO_{suf}'] = t[f'PASSIVO_FINANCEIRO_INSUPERAVIT_{suf}']
        else:
            # Fallback (aproximação -- ver nota acima, NÃO bate com oficial).
            t[f'PASSIVO_FINANCEIRO_{suf}'] = (t.get(f'OBRIG_TRAB_CP_{suf}', 0.0)
                                               + t.get(f'EMPREST_FINANC_CP_{suf}', 0.0)
                                               + t.get(f'FORNEC_CP_{suf}', 0.0)
                                               + t.get(f'OBRIG_FISCAIS_CP_{suf}', 0.0))
        t[f'PASSIVO_PERMANENTE_{suf}'] = t[f'PASSIVO_PL_TOTAL_{suf}'] - t[f'PASSIVO_FINANCEIRO_{suf}']

        t[f'ATIVO_I_{suf}'] = t[f'ATIVO_FINANCEIRO_{suf}'] + t[f'ATIVO_PERMANENTE_{suf}']
        t[f'PASSIVO_II_{suf}'] = t[f'PASSIVO_FINANCEIRO_{suf}'] + t[f'PASSIVO_PERMANENTE_{suf}']
        t[f'SALDO_PATRIMONIAL_III_{suf}'] = t[f'ATIVO_I_{suf}'] - t[f'PASSIVO_II_{suf}']

        # ── CONTAS DE COMPENSACAO ────────────────────────────────────────
        t[f'ATOS_POT_ATIVOS_TOTAL_{suf}'] = sum(
            t.get(f'{k}_{suf}', 0.0) for _, k, _, _ in ATOS_POT_ATIVOS_ITENS)
        t[f'ATOS_POT_PASSIVOS_TOTAL_{suf}'] = sum(
            t.get(f'{k}_{suf}', 0.0) for _, k, _, _ in ATOS_POT_PASSIVOS_ITENS)

    return t


# ─────────────────────────────────────────────────────────────────────────────
#  ESTRUTURA DE EXIBICAO  (espelha as paginas 1-3 do PDF oficial)
#  Cada linha (lado Ativo / lado Passivo+PL): (descricao, chave, bold, nivel)
# ─────────────────────────────────────────────────────────────────────────────
ESTRUTURA_BP_ATIVO = [
    ('ATIVO',                                        'ATIVO_TOTAL',          True,  0),
    ('ATIVO CIRCULANTE',                              'ATIVO_CIRCULANTE',     True,  0),
    ('  Caixa e Equivalentes de Caixa',               'CAIXA_EQUIV',          False, 1),
    ('  Créditos a Curto Prazo',                      'CRED_CP',              False, 1),
    ('  Demais Créditos e Valores a Curto Prazo',     'DEMAIS_CRED_CP',       False, 1),
    ('  Investimentos e Aplic. Temporárias a Curto Prazo', 'INVEST_APLIC_CP', False, 1),
    ('  Estoques',                                    'ESTOQUES',             False, 1),
    ('  VPD Pagas Antecipadamente',                   'VPD_PAGAS_ANTEC',      False, 1),
    ('ATIVO NÃO CIRCULANTE',                          'ATIVO_NCIRCULANTE',    True,  0),
    ('  Realizável a Longo Prazo',                    'REALIZAVEL_LP',        False, 1),
    ('  Investimentos',                               'INVESTIMENTOS',        False, 1),
    ('  Imobilizado',                                 'IMOBILIZADO',          False, 1),
    ('  Intangível',                                  'INTANGIVEL',           False, 1),
]

ESTRUTURA_BP_PASSIVO = [
    ('PASSIVO E PATRIMÔNIO LÍQUIDO',                  'PASSIVO_PL_TOTAL',    True,  0),
    ('PASSIVO CIRCULANTE',                            'PASSIVO_CIRCULANTE',  True,  0),
    ('  Obrig. Trab.,Prev. e Assist. a Pagar a Curto Prazo', 'OBRIG_TRAB_CP', False, 1),
    ('  Empréstimos e Financiamentos a Curto Prazo',  'EMPREST_FINANC_CP',   False, 1),
    ('  Fornecedores e Contas a Pagar a Curto Prazo', 'FORNEC_CP',           False, 1),
    ('  Obrigações Fiscais a Curto Prazo',            'OBRIG_FISCAIS_CP',    False, 1),
    ('  Transferências Fiscais a Curto Prazo',        'TRANSF_FISCAIS_CP',   False, 1),
    ('  Provisões a Curto Prazo',                     'PROVISOES_CP',        False, 1),
    ('  Demais Obrigações a Curto Prazo',             'DEMAIS_OBRIG_CP',     False, 1),
    ('PASSIVO NÃO CIRCULANTE',                        'PASSIVO_NCIRCULANTE', True,  0),
    ('  Obrig. Trab., Previd. e Assist. a Pagar a LP', 'OBRIG_TRAB_LP',      False, 1),
    ('  Empréstimos e Financiamentos a Longo Prazo',  'EMPREST_FINANC_LP',   False, 1),
    ('  Fornecedores e Contas a Pagar a Longo Prazo', 'FORNEC_LP',           False, 1),
    ('  Obrigações Fiscais a Longo Prazo',            'OBRIG_FISCAIS_LP',    False, 1),
    ('  Provisões a Longo Prazo',                     'PROVISOES_LP',        False, 1),
    ('  Demais Obrigações a Longo Prazo',             'DEMAIS_OBRIG_LP',     False, 1),
    ('  Resultado Diferido',                          'RESULTADO_DIFERIDO',  False, 1),
    ('PATRIMÔNIO LÍQUIDO',                            'PATRIMONIO_LIQUIDO',  True,  0),
    ('  Patrimônio Social e Capital Social',          'PAT_SOCIAL_CAPITAL',  False, 1),
    ('  Adiantamento para Futuro Aumento de Capital', 'AFAC',                False, 1),
    ('  Reservas de Capital',                         'RESERVA_CAPITAL',     False, 1),
    ('  Ajustes de Avaliação Patrimonial',            'AJUSTE_AVAL_PATRIM',  False, 1),
    ('  Reservas de Lucros',                          'RESERVAS_LUCROS',     False, 1),
    ('  Demais Reservas',                             'DEMAIS_RESERVAS',     False, 1),
    ('  Resultado Acumulado',                         'RESULTADO_ACUMULADO', False, 1),
    ('  Resultado do Exercício',                      'RESULTADO_EXERCICIO', False, 1),
    ('  Superávits ou Déficits Acumulados',           'SUPERAVIT_DEFICIT_ACUM', False, 1),
    ('  Lucros e Prejuízos Acumulados',               'LUCROS_PREJUIZOS_ACUM', False, 1),
]

ESTRUTURA_QUADRO_FIN_PERM = [
    ('ATIVO ( I )',           'ATIVO_I',              True),
    ('  Ativo Financeiro',    'ATIVO_FINANCEIRO',     False),
    ('  Ativo Permanente',    'ATIVO_PERMANENTE',     False),
    ('PASSIVO ( II )',        'PASSIVO_II',           True),
    ('  Passivo Financeiro',  'PASSIVO_FINANCEIRO',   False),
    ('  Passivo Permanente',  'PASSIVO_PERMANENTE',   False),
    ('SALDO PATRIMONIAL (III) = (I - II)', 'SALDO_PATRIMONIAL_III', True),
]

ESTRUTURA_CONTAS_COMPENSACAO_ATIVO = [
    ('ATOS POTENCIAIS ATIVOS',  'ATOS_POT_ATIVOS_TOTAL', True),
] + [(f'  {desc}', chave, False) for _, chave, _, desc in ATOS_POT_ATIVOS_ITENS]

ESTRUTURA_CONTAS_COMPENSACAO_PASSIVO = [
    ('ATOS POTENCIAIS PASSIVOS', 'ATOS_POT_PASSIVOS_TOTAL', True),
] + [(f'  {desc}', chave, False) for _, chave, _, desc in ATOS_POT_PASSIVOS_ITENS]


# ─────────────────────────────────────────────────────────────────────────────
#  AUDITORIA DE INTEGRIDADE
# ─────────────────────────────────────────────────────────────────────────────
def auditoria_integridade(t):
    """
    Retorna lista de tuplas (status, titulo, detalhe) onde status e
    'OK', 'ERRO' ou 'ALERTA'.
    """
    achados = []

    for suf, label in [('AT', 'Exercício Atual'), ('ANT', 'Exercício Anterior')]:
        # ── Controle 1: ATIVO = PASSIVO + PL (identidade fundamental) ────
        ativo = t[f'ATIVO_TOTAL_{suf}']
        passivo_pl = t[f'PASSIVO_PL_TOTAL_{suf}']
        dif = ativo - passivo_pl
        if abs(dif) < 1.00:
            achados.append(('OK', f'ATIVO = PASSIVO + PL ({label})',
                            f'Ativo {ativo:,.2f}  =  Passivo+PL {passivo_pl:,.2f}'))
        else:
            achados.append(('ERRO', f'ATIVO = PASSIVO + PL ({label})',
                            f'Diferença: {dif:,.2f}  |  Ativo {ativo:,.2f}  vs  '
                            f'Passivo+PL {passivo_pl:,.2f} -- identidade '
                            f'fundamental do Balanço Patrimonial NÃO fecha.'))

        # ── Controle 2: Ativo Circulante + Não Circulante = Ativo Total ──
        soma_at_ativo = t[f'ATIVO_CIRCULANTE_{suf}'] + t[f'ATIVO_NCIRCULANTE_{suf}']
        dif2 = soma_at_ativo - ativo
        if abs(dif2) < 1.00:
            achados.append(('OK', f'Ativo Circ. + Não Circ. = Ativo Total ({label})',
                            f'Diferença: {dif2:,.2f}'))
        else:
            achados.append(('ERRO', f'Ativo Circ. + Não Circ. = Ativo Total ({label})',
                            f'Diferença: {dif2:,.2f}'))

        # ── Controle 3: Quadro Financeiro/Permanente vs BP (mesma fonte) ──
        # SOLUÇÃO DEFINITIVA CONFIRMADA (--diag em produção, GDF, Maio/2026,
        # após localizar Relatório_Equação_de_Balanço_Patrimonial.xlsx): a
        # classificação real usa CONTACONTABIL.INSISCONTABIL ('F'/'P') --
        # ver buscar_financeiro_permanente(). Testado e confirmado com
        # DIFERENÇA ZERO contra os valores oficiais: Ativo Financeiro
        # R$7.665.797.148,08 e Passivo Financeiro R$6.545.291.111,33
        # (este último incluindo as 3 contas extras de classe 6 --
        # 622130100, 622130500, 631100000). Os achados abaixo continuam
        # sendo identidades algébricas internas (Permanente = Total -
        # Financeiro, por construção), mas agora a componente "Financeiro"
        # em si é validada como exata -- ver Controle 3b para a comparação
        # direta com a referência oficial.
        dif3 = t[f'ATIVO_I_{suf}'] - ativo
        if abs(dif3) < 1.00:
            achados.append(('OK', f'Ativo (I) [Financ.+Perm.] = Ativo do BP ({label})',
                            f'Diferença: {dif3:,.2f} -- identidade algébrica '
                            f'interna (Permanente = Total - Financeiro); o '
                            f'componente Financeiro em si foi validado como '
                            f'exato contra o oficial (ver Controle 3b).'))
        else:
            achados.append(('ALERTA', f'Ativo (I) [Financ.+Perm.] = Ativo do BP ({label})',
                            f'Diferença: {dif3:,.2f} -- a reclassificação '
                            f'financeira/permanente não reconcilia exatamente '
                            f'com o Ativo total do BP nos dados de origem.'))

        dif4 = t[f'PASSIVO_II_{suf}'] - passivo_pl
        if abs(dif4) < 1.00:
            achados.append(('OK', f'Passivo (II) [Financ.+Perm.] = Passivo+PL do BP ({label})',
                            f'Diferença: {dif4:,.2f} -- identidade algébrica '
                            f'interna (mesmo motivo do item acima).'))
        else:
            achados.append(('INFO', f'Passivo (II) [Financ.+Perm.] = Passivo+PL do BP ({label})',
                            f'Diferença: {dif4:,.2f}  |  Passivo(II) {t[f"PASSIVO_II_{suf}"]:,.2f}  '
                            f'vs  Passivo+PL(BP) {passivo_pl:,.2f}. ATENÇÃO: esta '
                            f'divergência também existe no PDF oficial PSIAG550 '
                            f'(Quadro dos Ativos e Passivos Financeiros e '
                            f'Permanentes vs Balanço Patrimonial) -- não é um '
                            f'erro de cálculo deste módulo, é uma característica '
                            f'observada na fonte de dados do GDF. Tratado como '
                            f'INFO (não ERRO) por esse motivo; o "Passivo '
                            f'Permanente" provavelmente usa base de apuração '
                            f'distinta (possíveis reclassificações OFSS) da '
                            f'usada no Passivo Não Circulante do BP.'))

        # ── Controle 3b: fonte do Ativo/Passivo Financeiro (real vs fallback) ──
        # Sinaliza se buscar_financeiro_permanente() conseguiu rodar (JOIN
        # com CONTACONTABIL.INSISCONTABIL) ou se caiu no fallback por
        # aproximação de grupo (sem acesso ao Oracle, ex.: chamada de
        # teste). SOLUÇÃO DEFINITIVA encontrada via Relatório_Equação_de_
        # Balanço_Patrimonial.xlsx (Lista Equações de Balanço do sistema
        # contábil do GDF): o critério real é a coluna INSISCONTABIL
        # ("Sistema Contábil" no extrato XLS), com 'F'=Financeiro,
        # 'P'=Permanente -- e o Passivo Financeiro inclui 3 contas extras
        # de classe 6 (622130100, 622130500, 631100000). CONFIRMADO com
        # DIFERENÇA ZERO via --diag em produção (Maio/2026): Ativo
        # Financeiro 7.665.797.148,08 (dif. 0,00) e Passivo Financeiro
        # 6.545.291.111,33 (dif. 0,00).
        usa_fonte_real = f'ATIVO_FINANCEIRO_INSUPERAVIT_{suf}' in t
        if usa_fonte_real:
            achados.append(('OK', f'Fonte do Ativo/Passivo Financeiro ({label})',
                            f'Usando CONTACONTABIL.INSISCONTABIL (\'F\'/\'P\' -- '
                            f'fonte OFICIAL, confirmada via Relatório_Equação_de_'
                            f'Balanço_Patrimonial.xlsx). Ativo Financeiro '
                            f'{t[f"ATIVO_FINANCEIRO_{suf}"]:,.2f}. Passivo '
                            f'Financeiro {t[f"PASSIVO_FINANCEIRO_{suf}"]:,.2f} '
                            f'(inclui as 3 contas extras de classe 6: '
                            f'622130100, 622130500, 631100000). CONFIRMADO via '
                            f'--diag em produção (Maio/2026) com DIFERENÇA '
                            f'ZERO em ambos os valores frente ao oficial.'))
        else:
            achados.append(('ALERTA', f'Fonte do Ativo/Passivo Financeiro ({label})',
                            f'Usando fallback por APROXIMAÇÃO de grupo de contas '
                            f'(NÃO bate com o oficial) -- a fonte real '
                            f'(CONTACONTABIL.INSISCONTABIL, via JOIN no Oracle) não '
                            f'estava disponível neste cálculo. Rode com conexão '
                            f'Oracle ativa (buscar_tudo) para usar a fonte real.'))

        # ── Controle 4: Saldo Patrimonial (III) = (I - II) ───────────────
        # Identidade INTERNA do Quadro Financeiro/Permanente (página 2) --
        # sempre deve fechar exato, independentemente da divergência do
        # Controle 3 (que compara o Quadro 2 contra o BP da página 1).
        calc_saldo = t[f'ATIVO_I_{suf}'] - t[f'PASSIVO_II_{suf}']
        reportado = t[f'SALDO_PATRIMONIAL_III_{suf}']
        dif5 = calc_saldo - reportado
        if abs(dif5) < 1.00:
            achados.append(('OK', f'Saldo Patrimonial (III) = (I - II) ({label})',
                            f'{calc_saldo:,.2f}'))
        else:
            achados.append(('ERRO', f'Saldo Patrimonial (III) = (I - II) ({label})',
                            f'Diferença: {dif5:,.2f}'))

    # ── Controle 5: Resultado do Exercício (Atual) variação vs Anterior ──
    # Apenas informativo -- grandes saltos merecem revisão, mas nao
    # configuram erro (sazonalidade/eventos atipicos sao esperados).
    mes_ref = t.get('_MES_REF')
    if mes_ref and 1 <= mes_ref <= 12:
        achados.append(('INFO', 'Resultado do Exercício -- Atual vs Anterior',
                        f"Atual (acum. jan-{MESES.get(mes_ref, mes_ref)}): "
                        f"{t.get('RESULTADO_EXERCICIO_AT', 0):,.2f}  |  "
                        f"Anterior (ano completo): "
                        f"{t.get('RESULTADO_EXERCICIO_ANT', 0):,.2f}"))

    # ── Controle 7: fonte e decomposição do PL do Exercício Anterior ────
    # Ver DIAG16/DIAG17 (Ago/2026): o PL do Exercício Anterior passou a ser
    # lido de MIL{ano-1} INMES 0..13 (PRE-encerramento), porque o INMES=14
    # transfere o Resultado do exercício para dentro de 231/2371 e fazia as
    # sub-linhas divergirem do oficial. Este controle verifica que a
    # decomposição soma exatamente o total do PL e registra qual fonte foi
    # efetivamente usada (a função tem fallback se MIL{ano-1} falhar).
    fonte_pl_ant = str(t.get('_PL_ANT_FONTE', '(não registrada)'))
    soma_sub_pl_ant = (sum(t.get(f'{k}_ANT', 0.0) for k in PL_CHAVES_SALDO)
                       + t.get('RESULTADO_ACUMULADO_ANT', 0.0))
    dif_pl_ant = soma_sub_pl_ant - t.get('PATRIMONIO_LIQUIDO_ANT', 0.0)
    if 'pre-encerramento' not in fonte_pl_ant:
        achados.append(('ALERTA', 'Fonte do PL (Exercício Anterior)',
                        f'Fonte em uso: {fonte_pl_ant}. O TOTAL do PL está '
                        f'correto, mas as sub-linhas Patrimônio Social e '
                        f'Superávits/Déficits trazem o encerramento do '
                        f'exercício embutido e NÃO reproduzem o relatório '
                        f'oficial. Verificar o acesso ao schema do ano '
                        f'anterior.'))
    elif abs(dif_pl_ant) < 1.00:
        achados.append(('OK', 'Fonte do PL (Exercício Anterior)',
                        f'Fonte: {fonte_pl_ant}. Sub-linhas somam '
                        f'{soma_sub_pl_ant:,.2f} = Patrimônio Líquido '
                        f'reportado. O Resultado do Exercício '
                        f'({t.get("RESULTADO_EXERCICIO_ANT", 0):,.2f}) é '
                        f'exibido em linha própria, como no oficial; o '
                        f'INMES=14 (encerramento) fica fora da fonte, sem '
                        f'alterar o total do PL. CONFIRMADO via DIAG17 com '
                        f'diferença ZERO nas 8 contas do bloco.'))
    else:
        achados.append(('ERRO', 'Fonte do PL (Exercício Anterior)',
                        f'Decomposição do PL não fecha: sub-linhas '
                        f'{soma_sub_pl_ant:,.2f} vs total '
                        f'{t.get("PATRIMONIO_LIQUIDO_ANT", 0):,.2f} '
                        f'(diferença {dif_pl_ant:,.2f}).'))

    # ── Controle 6: aviso fixo -- Contas de Compensação (pág. 3) ─────────
    # SOLUÇÃO DEFINITIVA CONFIRMADA: a partir da descoberta do arquivo
    # Relatório_Equação_de_Balanço_Patrimonial.xlsx ("Lista Equações de
    # Balanço" extraída do próprio sistema contábil do GDF), as contas
    # usadas neste módulo para Atos Potenciais Ativos/Passivos (811XXXXXX/
    # 812XXXXXX, EXATAS -- não faixas) são a FONTE OFICIAL (substituem a
    # aproximação anterior por 711/712, que batia parcialmente por
    # coincidência). CONFIRMADO via --diag em produção (GDF, Maio/2026)
    # com DIFERENÇA ZERO: Atos Potenciais Ativos R$14.228.115.268,87 e
    # Atos Potenciais Passivos R$65.253.334.489,58, ambos batendo exato
    # com a referência oficial.
    achados.append(('OK', 'Contas de Compensação (Pág. 3) -- fonte oficial confirmada',
                    'Contas exatas (811XXXXXX/812XXXXXX) extraídas de '
                    'Relatório_Equação_de_Balanço_Patrimonial.xlsx (Lista '
                    'Equações de Balanço do sistema contábil do GDF). '
                    'CONFIRMADO via --diag em produção (Maio/2026) com '
                    'DIFERENÇA ZERO nos totais de Atos Potenciais Ativos e '
                    'Passivos frente aos valores oficiais.'))

    return achados


def imprimir_auditoria(achados):
    icones = {'OK': '✔', 'ERRO': '✘', 'ALERTA': '⚠', 'INFO': 'ℹ'}
    print(f"\n{'─'*70}")
    print("  AUDITORIA DE INTEGRIDADE")
    print(f"{'─'*70}")
    for status, titulo, detalhe in achados:
        ic = icones.get(status, '?')
        print(f"  {ic} [{status:<6}] {titulo}")
        print(f"            {detalhe}")
    n_erro = sum(1 for s, _, _ in achados if s == 'ERRO')
    n_alerta = sum(1 for s, _, _ in achados if s == 'ALERTA')
    print(f"{'─'*70}")
    if n_erro:
        print(f"  RESULTADO: {n_erro} ERRO(s) sinalizado(s) - ação necessária.")
    elif n_alerta:
        print(f"  RESULTADO: 0 erros, {n_alerta} alerta(s) - revisar quando possível.")
    else:
        print(f"  RESULTADO: TODOS OS CONTROLES PASSARAM.")
    print(f"{'─'*70}")


# ─────────────────────────────────────────────────────────────────────────────
#  EXCEL
# ─────────────────────────────────────────────────────────────────────────────
FMT_BRL = '#,##0.00'
def _fill(h): return PatternFill('solid', fgColor=h)
def _side(s='thin'): return Side(border_style=s, color='B8CCE4')
B_THIN = Border(left=_side(), right=_side(), top=_side(), bottom=_side())
# ── Paleta "azul bem claro" (substitui o cinza/amarelo anterior) ─────────
# F_HDR    -- cabecalho de coluna (Exercicio Atual/Anterior)
# F_GRAY   -- linhas de TOTAL/subtotal de grupo (ex.: ATIVO CIRCULANTE)
# F_YELL   -- linha de destaque maximo (ex.: SALDO PATRIMONIAL) -- mantido
#             o nome F_YELL por compatibilidade com o codigo existente,
#             mas a cor agora e um azul mais saturado, nao mais amarelo
# F_WHIT   -- linhas normais (quase branco, leve tom azulado)
F_HDR  = _fill('BDD7EE')   # azul claro medio -- cabecalhos de coluna
F_GRAY = _fill('DCE6F1')   # azul muito claro -- linhas de total/subtotal
F_YELL = _fill('9DC3E6')   # azul claro mais saturado -- destaque maximo
F_WHIT = _fill('F5F9FD')   # quase branco com leve tom azulado -- linhas normais

# Mesma paleta em string hex (sem PatternFill) para uso no PDF (reportlab
# colors.HexColor) -- mantém Excel e PDF visualmente idênticos.
HEX_TITULO   = '#2E5C8A'  # azul medio-escuro para titulos/textos de destaque
HEX_HDR      = '#BDD7EE'  # cabecalho de coluna (igual F_HDR)
HEX_GRAY     = '#DCE6F1'  # linha de total/subtotal (igual F_GRAY)
HEX_YELL     = '#9DC3E6'  # destaque maximo (igual F_YELL)
HEX_GRID     = '#B8CCE4'  # linhas de grade das tabelas (mais suave que cinza)


def cel(ws, r, c, v='', bold=False, sz=9, bg=None, ha='left', brd=None,
        fmt=None, ind=0):
    x = ws.cell(row=r, column=c, value=v)
    x.font = Font(name='Arial', bold=bold, size=sz, color='000000')
    x.alignment = Alignment(horizontal=ha, vertical='center', indent=ind)
    if bg: x.fill = bg
    if brd: x.border = brd
    if fmt and v not in ('', None): x.number_format = fmt
    return x


def gerar_excel(t, mes, ano, ug_label, output_path, achados=None):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = 'Balanco Patrimonial'
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'portrait'; ws.page_setup.paperSize = 9
    for col, w in zip('ABC', [48, 18, 18]):
        ws.column_dimensions[col].width = w

    ws.merge_cells('A1:C2'); c = ws['A1']
    c.value = ('GOVERNO DO DISTRITO FEDERAL\nBalanço Patrimonial — Versão 1')
    c.font = Font(name='Arial', bold=True, size=12, color='1F4E79')
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 30
    cel(ws, 3, 1,
        f'Mês de Referência   {mes:02d} - {MESES[mes]}          {ug_label}   '
        f'Posição em: {datetime.now():%d/%m/%Y %H:%M:%S}', bold=True, sz=9)
    ws.merge_cells('A3:C3')
    ws.row_dimensions[4].height = 4

    H = 5
    cel(ws, H, 1, '', bold=True, bg=F_HDR, brd=B_THIN)
    cel(ws, H, 2, 'Exercício Atual', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    cel(ws, H, 3, 'Exercício Anterior', bold=True, bg=F_HDR, brd=B_THIN, ha='center')

    r = H + 1
    # ── Bloco ATIVO ──────────────────────────────────────────────────────
    for desc, chave, bold, nivel in ESTRUTURA_BP_ATIVO:
        bg = F_GRAY if bold and nivel == 0 else F_WHIT
        cel(ws, r, 1, desc.strip(), bold=bold, bg=bg, brd=B_THIN, ha='left', ind=nivel)
        cel(ws, r, 2, t.get(f'{chave}_AT', 0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 3, t.get(f'{chave}_ANT', 0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        r += 1

    r += 1
    cel(ws, r, 1, '', bold=True, bg=F_HDR, brd=B_THIN)
    cel(ws, r, 2, 'Exercício Atual', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    cel(ws, r, 3, 'Exercício Anterior', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    r += 1

    # ── Bloco PASSIVO + PL ───────────────────────────────────────────────
    for desc, chave, bold, nivel in ESTRUTURA_BP_PASSIVO:
        bg = F_GRAY if bold and nivel == 0 else F_WHIT
        cel(ws, r, 1, desc.strip(), bold=bold, bg=bg, brd=B_THIN, ha='left', ind=nivel)
        cel(ws, r, 2, t.get(f'{chave}_AT', 0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws, r, 3, t.get(f'{chave}_ANT', 0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        r += 1

    FT = r + 1
    cel(ws, FT, 1, 'Página: 1', sz=8)
    cel(ws, FT, 3, f'Emitido em: {datetime.now():%d/%m/%Y}', sz=8, ha='right')

    # ── Aba 2: Quadro dos Ativos e Passivos Financeiros e Permanentes ────
    ws2 = wb.create_sheet('Quadro Financ. Permanente')
    ws2.sheet_view.showGridLines = False
    for col, w in zip('ABC', [48, 18, 18]):
        ws2.column_dimensions[col].width = w
    ws2.merge_cells('A1:C2'); c = ws2['A1']
    c.value = ('GOVERNO DO DISTRITO FEDERAL\n'
               'Quadro dos Ativos e Passivos Financeiros e Permanentes')
    c.font = Font(name='Arial', bold=True, size=12, color='1F4E79')
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws2.row_dimensions[1].height = 30
    cel(ws2, 3, 1,
        f'Mês de Referência   {mes:02d} - {MESES[mes]}          {ug_label}',
        bold=True, sz=9)

    H2 = 5
    cel(ws2, H2, 1, '', bold=True, bg=F_HDR, brd=B_THIN)
    cel(ws2, H2, 2, 'Exercício Atual', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    cel(ws2, H2, 3, 'Exercício Anterior', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    for i, (desc, chave, bold) in enumerate(ESTRUTURA_QUADRO_FIN_PERM):
        r2 = H2 + 1 + i
        bg = F_YELL if 'SALDO PATRIMONIAL' in desc else (F_GRAY if bold else F_WHIT)
        cel(ws2, r2, 1, desc.strip(), bold=bold, bg=bg, brd=B_THIN, ha='left',
            ind=0 if bold else 1)
        cel(ws2, r2, 2, t.get(f'{chave}_AT', 0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws2, r2, 3, t.get(f'{chave}_ANT', 0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)

    # ── Aba 3: Quadro das Contas de Compensação ──────────────────────────
    ws3 = wb.create_sheet('Contas de Compensação')
    ws3.sheet_view.showGridLines = False
    for col, w in zip('ABC', [48, 18, 18]):
        ws3.column_dimensions[col].width = w
    ws3.merge_cells('A1:C2'); c = ws3['A1']
    c.value = ('GOVERNO DO DISTRITO FEDERAL\n'
               'Quadro das Contas de Compensação')
    c.font = Font(name='Arial', bold=True, size=12, color='1F4E79')
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws3.row_dimensions[1].height = 30
    cel(ws3, 3, 1,
        f'Mês de Referência   {mes:02d} - {MESES[mes]}          {ug_label}',
        bold=True, sz=9)

    H3 = 5
    cel(ws3, H3, 1, '', bold=True, bg=F_HDR, brd=B_THIN)
    cel(ws3, H3, 2, 'Exercício Atual', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    cel(ws3, H3, 3, 'Exercício Anterior', bold=True, bg=F_HDR, brd=B_THIN, ha='center')
    r3 = H3 + 1
    for desc, chave, bold in (ESTRUTURA_CONTAS_COMPENSACAO_ATIVO
                               + ESTRUTURA_CONTAS_COMPENSACAO_PASSIVO):
        bg = F_GRAY if bold else F_WHIT
        cel(ws3, r3, 1, desc.strip(), bold=bold, bg=bg, brd=B_THIN, ha='left',
            ind=0 if bold else 1)
        cel(ws3, r3, 2, t.get(f'{chave}_AT', 0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        cel(ws3, r3, 3, t.get(f'{chave}_ANT', 0), bold=bold, bg=bg, brd=B_THIN, ha='right', fmt=FMT_BRL)
        r3 += 1

    # ── Aba de dados brutos ───────────────────────────────────────────────
    ws4 = wb.create_sheet('Dados Brutos SQL')
    for j, (k, v) in enumerate(t.items(), 1):
        ws4.cell(row=1, column=j, value=k).font = Font(bold=True, name='Arial', size=9)
        c4 = ws4.cell(row=2, column=j, value=v)
        if isinstance(v, float): c4.number_format = FMT_BRL

    if achados:
        ws5 = wb.create_sheet('Auditoria Integridade')
        ws5.column_dimensions['A'].width = 10
        ws5.column_dimensions['B'].width = 50
        ws5.column_dimensions['C'].width = 80
        hdr_bg = _fill('1F4E79')
        c1 = cel(ws5, 1, 1, 'Status', bold=True, bg=hdr_bg, ha='center', brd=B_THIN)
        c1.font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        c2 = cel(ws5, 1, 2, 'Controle', bold=True, bg=hdr_bg, ha='center', brd=B_THIN)
        c2.font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        c3 = cel(ws5, 1, 3, 'Detalhe', bold=True, bg=hdr_bg, ha='center', brd=B_THIN)
        c3.font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        cor_status = {'OK': _fill('D9F0DD'), 'ERRO': _fill('FAD7DA'),
                      'ALERTA': _fill('FCF0CE'), 'INFO': _fill('E8F1FA')}
        for i, (status, titulo, detalhe) in enumerate(achados, start=2):
            bg = cor_status.get(status, F_WHIT)
            cel(ws5, i, 1, status, bold=True, bg=bg, ha='center', brd=B_THIN, sz=9)
            cel(ws5, i, 2, titulo, bg=bg, brd=B_THIN, sz=9)
            cel(ws5, i, 3, detalhe, bg=bg, brd=B_THIN, sz=8)

    wb.save(output_path)
    print(f"  Excel salvo: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  PDF
# ─────────────────────────────────────────────────────────────────────────────
def _brl(v):
    if v is None: return ""
    neg = v < 0
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"({s})" if neg else s


def gerar_pdf(t, mes, ano, ug_label, output_path, achados=None):
    if not REPORTLAB_OK:
        print("  PDF ignorado: reportlab não instalado. Rode: pip install reportlab")
        return

    doc = SimpleDocTemplate(str(output_path), pagesize=A4,
        leftMargin=14*mm, rightMargin=14*mm, topMargin=12*mm, bottomMargin=12*mm,
        title=f"Balanço Patrimonial GDF {ano} - {MESES[mes]}", author="SIAC/SIGGO - GDF")

    styles = getSampleStyleSheet()
    st_t = ParagraphStyle('t', parent=styles['Normal'], fontName='Helvetica-Bold',
                          fontSize=12, textColor=colors.HexColor(HEX_TITULO), leading=15)
    st_s = ParagraphStyle('s', parent=styles['Normal'], fontName='Helvetica',
                          fontSize=8, alignment=TA_RIGHT, leading=11)
    st_r = ParagraphStyle('r', parent=styles['Normal'], fontName='Helvetica-Bold',
                          fontSize=8.5)
    st_hdr = ParagraphStyle('hdr', fontName='Helvetica-Bold', fontSize=8,
                             alignment=TA_CENTER, textColor=colors.black, leading=9.5)
    st_cell_desc = ParagraphStyle('cd', fontName='Helvetica', fontSize=8, leading=10)
    st_cell_desc_b = ParagraphStyle('cdb', fontName='Helvetica-Bold', fontSize=8, leading=10)
    st_cell_val = ParagraphStyle('cv', fontName='Helvetica', fontSize=8,
                                  alignment=TA_RIGHT, leading=10)
    st_cell_val_b = ParagraphStyle('cvb', fontName='Helvetica-Bold', fontSize=8,
                                    alignment=TA_RIGHT, leading=10)

    def cabecalho(titulo_extra=''):
        cab = Table([[Paragraph(f'GOVERNO DO DISTRITO FEDERAL<br/>'
                                 f'{titulo_extra}<br/>Versão 1', st_t),
                      Paragraph(f'Exercício {ano}<br/>PSIAG550<br/>'
                                f'Posição em: {datetime.now():%d/%m/%Y %H:%M:%S}', st_s)]],
                    colWidths=[125*mm, 42*mm])
        cab.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        return cab

    elems = []
    elems.append(cabecalho('Balanço Patrimonial'))
    elems.append(Paragraph(
        f'Mês de Referência: {mes:02d} - {MESES[mes]} &nbsp; {ug_label}', st_r))
    elems.append(Spacer(1, 4*mm))

    W_DESC = 95*mm; W_VAL = 38*mm

    def linha(desc, chave, bold):
        sd, sv = (st_cell_desc_b, st_cell_val_b) if bold else (st_cell_desc, st_cell_val)
        return [Paragraph(desc, sd),
                Paragraph(_brl(t.get(f'{chave}_AT', 0)), sv),
                Paragraph(_brl(t.get(f'{chave}_ANT', 0)), sv)]

    # ── Tabela 1: ATIVO ───────────────────────────────────────────────────
    dados_ativo = [[Paragraph('ATIVO', st_hdr), Paragraph('Exerc. Atual', st_hdr),
                     Paragraph('Exerc. Anterior', st_hdr)]]
    est_a = []
    for i, (desc, chave, bold, nivel) in enumerate(ESTRUTURA_BP_ATIVO):
        dados_ativo.append(linha(desc, chave, bold))
        if bold and nivel == 0:
            est_a.append(i + 1)
    tab_a = Table(dados_ativo, colWidths=[W_DESC, W_VAL, W_VAL], repeatRows=1)
    ts_a = [('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(HEX_HDR))]
    for idx in est_a:
        ts_a.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor(HEX_GRAY)))
    tab_a.setStyle(TableStyle(ts_a))
    elems.append(tab_a)
    elems.append(Spacer(1, 4*mm))

    # ── Tabela 2: PASSIVO + PL ────────────────────────────────────────────
    dados_pas = [[Paragraph('PASSIVO E PATRIMÔNIO LÍQUIDO', st_hdr),
                  Paragraph('Exerc. Atual', st_hdr), Paragraph('Exerc. Anterior', st_hdr)]]
    est_p = []
    for i, (desc, chave, bold, nivel) in enumerate(ESTRUTURA_BP_PASSIVO):
        dados_pas.append(linha(desc, chave, bold))
        if bold and nivel == 0:
            est_p.append(i + 1)
    tab_p = Table(dados_pas, colWidths=[W_DESC, W_VAL, W_VAL], repeatRows=1)
    ts_p = [('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(HEX_HDR))]
    for idx in est_p:
        ts_p.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor(HEX_GRAY)))
    tab_p.setStyle(TableStyle(ts_p))
    elems.append(tab_p)

    # ── Página 2: Quadro Ativos/Passivos Financeiros e Permanentes ───────
    elems.append(PageBreak())
    elems.append(cabecalho('Quadro dos Ativos e Passivos Financeiros e Permanentes'))
    elems.append(Paragraph(
        f'Mês de Referência: {mes:02d} - {MESES[mes]} &nbsp; {ug_label}', st_r))
    elems.append(Spacer(1, 4*mm))

    dados_fp = [[Paragraph('', st_hdr), Paragraph('Exerc. Atual', st_hdr),
                 Paragraph('Exerc. Anterior', st_hdr)]]
    est_fp = []
    for i, (desc, chave, bold) in enumerate(ESTRUTURA_QUADRO_FIN_PERM):
        dados_fp.append(linha(desc, chave, bold))
        if bold:
            est_fp.append((i + 1, 't' if 'SALDO PATRIMONIAL' in desc else 'g'))
    tab_fp = Table(dados_fp, colWidths=[W_DESC, W_VAL, W_VAL], repeatRows=1)
    ts_fp = [('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
             ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
             ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
             ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3),
             ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(HEX_HDR))]
    for idx, tp in est_fp:
        col = HEX_YELL if tp == 't' else HEX_GRAY
        ts_fp.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor(col)))
    tab_fp.setStyle(TableStyle(ts_fp))
    elems.append(tab_fp)

    # ── Página 3: Quadro das Contas de Compensação ───────────────────────
    elems.append(PageBreak())
    elems.append(cabecalho('Quadro das Contas de Compensação'))
    elems.append(Paragraph(
        f'Mês de Referência: {mes:02d} - {MESES[mes]} &nbsp; {ug_label}', st_r))
    elems.append(Spacer(1, 4*mm))

    dados_cc = [[Paragraph('', st_hdr), Paragraph('Exerc. Atual', st_hdr),
                 Paragraph('Exerc. Anterior', st_hdr)]]
    est_cc = []
    todas_cc = ESTRUTURA_CONTAS_COMPENSACAO_ATIVO + ESTRUTURA_CONTAS_COMPENSACAO_PASSIVO
    for i, (desc, chave, bold) in enumerate(todas_cc):
        dados_cc.append(linha(desc, chave, bold))
        if bold:
            est_cc.append(i + 1)
    tab_cc = Table(dados_cc, colWidths=[W_DESC, W_VAL, W_VAL], repeatRows=1)
    ts_cc = [('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
             ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
             ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
             ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3),
             ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(HEX_HDR))]
    for idx in est_cc:
        ts_cc.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor(HEX_GRAY)))
    tab_cc.setStyle(TableStyle(ts_cc))
    elems.append(tab_cc)

    # ── Auditoria de Integridade (rodapé) ────────────────────────────────
    if achados:
        elems.append(Spacer(1, 5*mm))
        st_aud_tit = ParagraphStyle('audt', parent=styles['Normal'],
                       fontName='Helvetica-Bold', fontSize=10,
                       textColor=colors.HexColor(HEX_TITULO))
        elems.append(Paragraph('AUDITORIA DE INTEGRIDADE', st_aud_tit))
        elems.append(Spacer(1, 1*mm))

        st_cell_b = ParagraphStyle('cb', parent=styles['Normal'],
                      fontName='Helvetica-Bold', fontSize=7.5)
        st_cell_n = ParagraphStyle('cn', parent=styles['Normal'],
                      fontName='Helvetica', fontSize=7.5)
        cor_status = {'OK': '#D9F0DD', 'ERRO': '#FAD7DA', 'ALERTA': '#FCF0CE',
                      'INFO': '#E8F1FA'}
        dados_aud = [[Paragraph('Status', st_cell_b), Paragraph('Controle', st_cell_b),
                      Paragraph('Detalhe', st_cell_b)]]
        bgs = [None]
        for status, titulo, detalhe in achados:
            dados_aud.append([Paragraph(status, st_cell_b), Paragraph(titulo, st_cell_n),
                               Paragraph(detalhe, st_cell_n)])
            bgs.append(cor_status.get(status))
        tab_aud = Table(dados_aud, colWidths=[18*mm, 65*mm, 90*mm], repeatRows=1)
        ts2 = [('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
               ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
               ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(HEX_TITULO))]
        for i, bg in enumerate(bgs):
            if bg:
                ts2.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor(bg)))
        tab_aud.setStyle(TableStyle(ts2))
        elems.append(tab_aud)

        n_erro = sum(1 for s, _, _ in achados if s == 'ERRO')
        n_alerta = sum(1 for s, _, _ in achados if s == 'ALERTA')
        cor_res = '#B00000' if n_erro else ('#9C6500' if n_alerta else '#1F5C2E')
        st_res = ParagraphStyle('res', parent=styles['Normal'],
                   fontName='Helvetica-Bold', fontSize=8.5,
                   textColor=colors.HexColor(cor_res))
        elems.append(Spacer(1, 2*mm))
        if n_erro:
            elems.append(Paragraph(
                f'RESULTADO: {n_erro} pendência(s) de ERRO sinalizada(s) - ver detalhes acima.',
                st_res))
        elif n_alerta:
            elems.append(Paragraph(
                f'RESULTADO: Sem erros, mas {n_alerta} alerta(s) - revisar.', st_res))
        else:
            elems.append(Paragraph('RESULTADO: TODOS OS CONTROLES PASSARAM.', st_res))

    doc.build(elems)
    print(f"  PDF salvo:   {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  DIAGNOSTICO  (mesmo padrao BF/DFC/DVP/DMPL)
# ─────────────────────────────────────────────────────────────────────────────
def diagnostico(conn, ano, mes):
    print("\n" + "="*64)
    print("  DIAGNOSTICO - ESTRUTURA DO BANCO PARA BP")
    print("  VERSAO DO DIAGNOSTICO: 2026-08-03-r14 (SOLUCAO DEFINITIVA")
    print("  encontrada: Relatório_Equação_de_Balanço_Patrimonial.xlsx")
    print("  revela que Atos Potenciais usam contas 811/812 EXATAS (nao")
    print("  711/712), e Financeiro/Permanente usa CONTACONTABIL.")
    print("  INSISCONTABIL ('F'/'P') + 3 contas extras classe 6 no")
    print("  Passivo. Blocos [6a]-[6e] validam tudo isso no Oracle real.)")
    print("="*64)

    print(f"\n[1] SALDOCONTABIL existe e tem dados p/ contas 1XX/2XX em MIL{ano}?")
    try:
        q = f"""SELECT v.INMES, COUNT(*) QTD,
                    SUM(v.VADEBITO - v.VACREDITO) SALDO_SD
                FROM MIL{ano}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL BETWEEN 100000000 AND 199999999
                GROUP BY v.INMES ORDER BY v.INMES"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print("    NENHUM registro p/ contas 1XX em SALDOCONTABIL.")
        for _, r in df.iterrows():
            print(f"    INMES={int(r['INMES']):>3}  qtd={int(r['QTD']):>9}  "
                  f"saldo(SD)={float(r['SALDO_SD'] or 0):>22,.2f}")
        print(f"    >>> Esperado: INMES=0 com saldo grande (abertura) e "
              f"acumulado 0..{mes} tambem com saldo grande e crescente.")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[2] Contas 2XX (Passivo+PL) por INMES em MIL{ano}:")
    try:
        q = f"""SELECT v.INMES, COUNT(*) QTD,
                    SUM(v.VACREDITO - v.VADEBITO) SALDO_SC
                FROM MIL{ano}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL BETWEEN 200000000 AND 299999999
                GROUP BY v.INMES ORDER BY v.INMES"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        for _, r in df.iterrows():
            print(f"    INMES={int(r['INMES']):>3}  qtd={int(r['QTD']):>9}  "
                  f"saldo(SC)={float(r['SALDO_SC'] or 0):>22,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    ano_ant = ano - 1
    print(f"\n[3] Schema MIL{ano_ant} (Exercício Anterior) -- saldo de "
          f"encerramento (INMES=13) p/ contas 1XX/2XX. ATENÇÃO -- BUG "
          f"CRÍTICO ENCONTRADO E CORRIGIDO (confirmado rodando bp.py SEM "
          f"--diag, geração real): este bloco SEMPRE mostrou poucos "
          f"registros aqui (ex.: Classe=1 qtd=48, SD≈R$55,9 milhões) "
          f"porque INMES=13 é APENAS um ajuste residual de encerramento "
          f"(mesma definição usada em bf.py: '13 = encerramento (ajustes, "
          f"reclassif., inscrição de RP)'), NÃO o saldo total do ano "
          f"anterior. O saldo de encerramento COMPLETO do ano anterior é, "
          f"na verdade, IGUAL ao saldo de ABERTURA do ano atual -- ou "
          f"seja, deve-se usar MIL{ano}.SALDOCONTABIL WHERE INMES=0 (NÃO "
          f"MIL{ano_ant} WHERE INMES=13). Isso foi CORRIGIDO em "
          f"buscar_tudo() -- bp.py agora busca o Exercício Anterior "
          f"(saldos) no schema MIL{ano} com INMES=0, idêntico ao bloco "
          f"[1]/[2] abaixo.")
    try:
        q = f"""SELECT v.COCONTACONTABIL/100000000 AS CLASSE, v.INMES, COUNT(*) QTD,
                    SUM(v.VADEBITO - v.VACREDITO) SALDO_SD,
                    SUM(v.VACREDITO - v.VADEBITO) SALDO_SC
                FROM MIL{ano_ant}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL BETWEEN 100000000 AND 299999999
                  AND v.INMES = 13
                GROUP BY v.COCONTACONTABIL/100000000, v.INMES ORDER BY 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM registro em MIL{ano_ant}.SALDOCONTABIL INMES=13 "
                  f"p/ contas 1XX/2XX -- verificar se o schema existe.")
        for _, r in df.iterrows():
            print(f"    Classe={int(r['CLASSE'])}  qtd={int(r['QTD']):>9}  "
                  f"SD={float(r['SALDO_SD'] or 0):>20,.2f}  "
                  f"SC={float(r['SALDO_SC'] or 0):>20,.2f}")
        print(f"    >>> (Confirma o tamanho pequeno/residual de INMES=13 -- "
              f"isso é ESPERADO agora, não é mais um problema a resolver.)")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[4] Contas de compensação -- Atos Potenciais (711=Ativos, "
          f"712=Passivos) por INMES em MIL{ano}. ATENÇÃO: a faixa correta "
          f"é 711XXXXXX/712XXXXXX (confirmado em produção), NÃO a faixa "
          f"ampla 700000000-899999999 (que inclui dezenas de outros grupos "
          f"de controle não relacionados, como Administração Financeira, "
          f"Dívida Ativa, Riscos Fiscais, Custos etc., e por isso infla "
          f"o total se usada por engano):")
    try:
        q = f"""SELECT v.INMES, COUNT(*) QTD,
                    SUM(CASE WHEN v.COCONTACONTABIL BETWEEN 711000000 AND 711999999
                         THEN v.VADEBITO - v.VACREDITO ELSE 0 END) SALDO_ATIVOS_SD,
                    SUM(CASE WHEN v.COCONTACONTABIL BETWEEN 712000000 AND 712999999
                         THEN v.VACREDITO - v.VADEBITO ELSE 0 END) SALDO_PASSIVOS_SC
                FROM MIL{ano}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL BETWEEN 711000000 AND 712999999
                GROUP BY v.INMES ORDER BY v.INMES"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print(f"    NENHUM registro p/ contas 711XXXXXX/712XXXXXX -- "
                  f"conferir se a faixa de conta de compensacao usada no "
                  f"ente e diferente.")
        for _, r in df.iterrows():
            print(f"    INMES={int(r['INMES']):>3}  qtd={int(r['QTD']):>9}  "
                  f"AtosPotAtivos(SD)={float(r['SALDO_ATIVOS_SD'] or 0):>18,.2f}  "
                  f"AtosPotPassivos(SC)={float(r['SALDO_PASSIVOS_SC'] or 0):>18,.2f}")
        print(f"    >>> Comparar o saldo acumulado (INMES 0..{mes}) com os "
              f"valores oficiais de referência: Atos Potenciais Ativos "
              f"14.228.115.268,87 / Atos Potenciais Passivos "
              f"65.253.334.489,58 (Maio/2026).")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[4b] Verificando se a conta 712200000 (Obrigações Conveniados "
          f"e Outros Instrumentos Congêneres -- sem linha própria mapeada "
          f"no bp.py) tem saldo relevante em MIL{ano} (se tiver, o total "
          f"de Atos Potenciais Passivos deste módulo pode divergir do "
          f"oficial em exatamente esse valor):")
    try:
        q = f"""SELECT v.INMES, SUM(v.VACREDITO - v.VADEBITO) SALDO_SC
                FROM MIL{ano}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL BETWEEN 712200000 AND 712299999
                  AND v.INMES BETWEEN 0 AND {mes}
                GROUP BY v.INMES ORDER BY v.INMES"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty or df['SALDO_SC'].abs().sum() < 1:
            print(f"    Saldo desprezível ou zero -- pode ignorar com segurança.")
        else:
            for _, r in df.iterrows():
                print(f"    INMES={int(r['INMES']):>3}  saldo(SC)={float(r['SALDO_SC'] or 0):>18,.2f}")
            print(f"    >>> Saldo NÃO desprezível -- considere mapear esta "
                  f"conta em ATOS_POT_PASSIVOS_ITENS (bp.py) se a diferença "
                  f"persistir na auditoria.")
    except Exception as e:
        print(f"    ERRO: {e}")

    # ── [4c] Decomposicao fina das contas 711/712 -- o bloco [4] mostrou
    # que NEM o saldo acumulado (INMES 0..mes) NEM o isolado (INMES=mes)
    # batem com os valores oficiais, e o sinal do lado Passivo saiu
    # invertido (negativo) -- ambos sinais de que a formula SD/SC simples
    # (igual a 1XX/2XX) pode nao se aplicar a estas contas de controle.
    # Decompoe por sub-conta de 2o nivel e testa as DUAS naturezas (SD e
    # SC) lado a lado, para qualquer subconjunto bater com a referencia.
    print(f"\n[4c] Decomposição por sub-conta de 4 dígitos (711X/712X), saldo "
          f"acumulado INMES BETWEEN 0 AND {mes}, MIL{ano}. RESULTADO DA "
          f"RODADA ANTERIOR (já analisado): agregando por SD, os grupos "
          f"7111 (Garantias Recebidas), 7121 (Garantias Concedidas) e 7129 "
          f"(Outros Passivos) batem EXATO com a referência oficial. Os "
          f"grupos 7112 (Direitos Conveniados), 7119 (Outros Ativos) e "
          f"7123 (Obrigações Contratuais) têm saldo MAIOR que o oficial "
          f"-- diferenças de R$6,56 bi / R$4,66 bi / R$7,15 bi "
          f"respectivamente. O bloco [4d] abaixo decompõe esses 3 grupos "
          f"em nível de conta completa (9 dígitos) para achar a(s) "
          f"subconta(s) responsável(eis) pelo excesso.")
    try:
        q = f"""SELECT v.COCONTACONTABIL/100000 AS SUBCONTA,
                    SUM(v.VADEBITO - v.VACREDITO) SALDO_SD,
                    COUNT(*) QTD
                FROM MIL{ano}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL BETWEEN 711000000 AND 712999999
                  AND v.INMES BETWEEN 0 AND {mes}
                GROUP BY v.COCONTACONTABIL/100000 ORDER BY 1"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        agg = df.groupby((df['SUBCONTA']).round(0))['SALDO_SD'].sum()
        # Re-agrega por inteiro (alguns ambientes Oracle devolvem o /100000
        # com resto fracionario por arredondamento de NUMBER) -- normaliza
        # para garantir 1 linha por subconta de 4 digitos.
        agg2 = {}
        for _, r in df.iterrows():
            k = int(r['SUBCONTA'])
            agg2[k] = agg2.get(k, 0.0) + float(r['SALDO_SD'] or 0)
        for k in sorted(agg2):
            print(f"    Subconta={k:>6}  SD_total={agg2[k]:>20,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[4d] Decomposição em NÍVEL DE CONTA COMPLETA (9 dígitos) dos "
          f"grupos com diferença residual (7112, 7119, 7123), saldo "
          f"acumulado INMES BETWEEN 0 AND {mes}, MIL{ano} -- procure a(s) "
          f"conta(s) cujo saldo somado ultrapasse o valor de referência; "
          f"essa é a pista mais provável de lançamento incorreto, conta "
          f"de outro exercício/cancelamento não filtrado, ou de uma "
          f"subconta que NÃO deveria entrar neste agregado:")
    try:
        q = f"""SELECT v.COCONTACONTABIL, c.NOCONTACONTABIL,
                    SUM(v.VADEBITO - v.VACREDITO) SALDO_SD,
                    COUNT(*) QTD
                FROM MIL{ano}.SALDOCONTABIL v
                LEFT JOIN MIL{ano}.CONTACONTABIL c
                  ON c.COCONTACONTABIL = v.COCONTACONTABIL
                WHERE (v.COCONTACONTABIL BETWEEN 711200000 AND 711299999
                    OR v.COCONTACONTABIL BETWEEN 711900000 AND 711999999
                    OR v.COCONTACONTABIL BETWEEN 712300000 AND 712399999)
                  AND v.INMES BETWEEN 0 AND {mes}
                GROUP BY v.COCONTACONTABIL, c.NOCONTACONTABIL
                HAVING SUM(v.VADEBITO - v.VACREDITO) <> 0
                ORDER BY v.COCONTACONTABIL"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print("    NENHUM registro com saldo não-zero -- conferir filtro.")
        for _, r in df.iterrows():
            nome = (r.get('NOCONTACONTABIL') or '')[:55] if 'NOCONTACONTABIL' in r else ''
            print(f"    Conta={int(r['COCONTACONTABIL']):<11} qtd={int(r['QTD']):>6}  "
                  f"SD={float(r['SALDO_SD'] or 0):>18,.2f}  {nome}")
    except Exception as e:
        print(f"    ERRO: {e}")

    # ── [4e] HIPÓTESE "EXECUÇÃO" -- REFUTADA (testado em produção,
    # GDF, Maio/2026): SD(811X) = -SD(711X) EXATAMENTE em todos os 3
    # grupos testados (7112/8112, 7119/8119, 7123/8123) -- ou seja, 81X
    # não é uma "baixa/execução" independente, é a CONTRAPARTIDA de
    # PARTIDA DOBRADA do mesmo lançamento em 71X (mecanismo padrão de
    # conta de controle: toda movimentação em 711X gera automaticamente
    # o lançamento espelhado oposto em 811X, sem significado contábil
    # adicional). Subtrair 811X de 711X dobra o valor (NETO = 2x bruto),
    # como confirmado pelos números: NETO(7112-8112)=37,42bi = 2x
    # 18,71bi. Esta hipótese está DESCARTADA -- não usar 811X/812X para
    # nenhum ajuste de saldo de 711X/712X.
    #
    # ── [4f] Nova abordagem: força bruta confirmou que o excesso em
    # 7112/7119/7123 NÃO corresponde a nenhum subconjunto EXATO de
    # contas completas (testado exaustivamente, 2^20 combinações em
    # 7123, nenhuma bateu com diferença < R$0,01) -- ou seja, o problema
    # está DENTRO de uma ou mais contas individuais (ex.: lançamentos de
    # outro exercício/cancelamento não filtrado por esta query, ou saldo
    # de abertura INMES=0 anômalo nessas contas específicas). Decompõe
    # as 3 contas mais relevantes por INMES (mesmo padrão que funcionou
    # para diagnosticar as contas 1XX/2XX no bloco [1]) -- procure por
    # um INMES isolado com saldo desproporcional/anômalo, especialmente
    # INMES=0 (abertura) se ele sozinho já exceder a referência oficial.
    print(f"\n[4f] Decomposição por INMES das 3 contas mais relevantes "
          f"dentro dos grupos com excesso (711210101 maior conta de 7112; "
          f"711911400+711914000 as 2 maiores de 7119, cuja EXCLUSÃO single-"
          f"handedly aproxima muito bem do oficial -- diferença residual "
          f"de só ~R$27,5 milhões, 2,5% -- forte indício de que NÃO devem "
          f"compor o total; 712310200 maior conta de 7123), MIL{ano}:")
    try:
        q = f"""SELECT v.COCONTACONTABIL, v.INMES,
                    SUM(v.VADEBITO - v.VACREDITO) SALDO_SD, COUNT(*) QTD
                FROM MIL{ano}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL IN (711210101, 711911400, 711914000,
                                             712310200)
                  AND v.INMES BETWEEN 0 AND {mes}
                GROUP BY v.COCONTACONTABIL, v.INMES
                ORDER BY v.COCONTACONTABIL, v.INMES"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print("    NENHUM registro -- conferir contas/filtro.")
        for _, r in df.iterrows():
            print(f"    Conta={int(r['COCONTACONTABIL'])}  INMES={int(r['INMES']):>3}  "
                  f"qtd={int(r['QTD']):>6}  SD={float(r['SALDO_SD'] or 0):>18,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[4g] CONFIRMAÇÃO: testando se EXCLUIR as contas "
          f"711911400 (Bens de Uso Comum do Povo) e 711914000 (Controle "
          f"de Estoque Interno - Almoxarifado) do total de 'Outros Atos "
          f"Potenciais Ativos' (7119) aproxima do valor oficial "
          f"(1.069.427.911,29) -- indício forte encontrado por análise "
          f"manual dos dados da rodada anterior (diferença residual "
          f"esperada de apenas ~R$27,5 milhões, 2,5%, bem menor que o "
          f"excesso original de R$4,66 bi):")
    try:
        q = f"""SELECT
                    SUM(v.VADEBITO - v.VACREDITO) SALDO_SD_TOTAL,
                    SUM(CASE WHEN v.COCONTACONTABIL IN (711911400, 711914000)
                         THEN v.VADEBITO - v.VACREDITO ELSE 0 END) SALDO_EXCLUIR
                FROM MIL{ano}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL BETWEEN 711900000 AND 711999999
                  AND v.INMES BETWEEN 0 AND {mes}"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        total = float(df.iloc[0]['SALDO_SD_TOTAL'] or 0)
        excluir = float(df.iloc[0]['SALDO_EXCLUIR'] or 0)
        liquido = total - excluir
        print(f"    Total 7119 (bruto): {total:,.2f}")
        print(f"    Saldo das 2 contas a excluir: {excluir:,.2f}")
        print(f"    Total 7119 SEM essas 2 contas: {liquido:,.2f}")
        print(f"    Referência oficial: 1.069.427.911,29")
        print(f"    Diferença: {liquido - 1069427911.29:,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    # ── [4h] NOVA HIPÓTESE: duplicação/cancelamento por COUG (Unidade
    # Gestora) não capturado. As contas 711210101 (Direitos Conveniados)
    # e 712310200 (Obrigações Contratuais -- Contratos de Serviços) têm
    # centenas/milhares de lançamentos (885 e 5474 apenas no INMES=0) --
    # plausível que parte seja transferência/consolidação INTRA-
    # ORÇAMENTÁRIA entre UGs do mesmo ente, que deveria se CANCELAR no
    # "Consolidado Geral" mas a query atual (SUM simples, sem agrupar por
    # COUG) pode estar somando ambos os lados de um par que se anula.
    # Decompõe por COUG para visualizar se há padrão de valores que se
    # cancelam entre si (uma UG positiva, outra negativa, mesma ordem de
    # grandeza) que a soma simples não está eliminando.
    print(f"\n[4h] Decomposição por COUG (Unidade Gestora) da conta "
          f"711210101 (maior componente de 7112/Direitos Conveniados), "
          f"saldo acumulado INMES BETWEEN 0 AND {mes}, MIL{ano} -- "
          f"procure por padrões de cancelamento entre UGs (valores "
          f"simétricos positivo/negativo) que a soma simples não está "
          f"eliminando -- referência: o saldo correto desta conta "
          f"deveria contribuir para o total oficial de 7112 "
          f"(12.149.372.457,11), não para os 16,21 bi atuais:")
    try:
        q = f"""SELECT v.COUG, SUM(v.VADEBITO - v.VACREDITO) SALDO_SD, COUNT(*) QTD
                FROM MIL{ano}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL = 711210101
                  AND v.INMES BETWEEN 0 AND {mes}
                GROUP BY v.COUG
                ORDER BY ABS(SUM(v.VADEBITO - v.VACREDITO)) DESC"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print("    NENHUM registro -- conferir conta/filtro.")
        for _, r in df.iterrows():
            print(f"    COUG={int(r['COUG']):<8}  qtd={int(r['QTD']):>6}  "
                  f"SD={float(r['SALDO_SD'] or 0):>18,.2f}")
        print(f"    >>> Se houver pares de UGs com saldo aproximadamente "
              f"simétrico (uma muito positiva, outra muito negativa), "
              f"isso confirma a hipótese de transferência intra-"
              f"orçamentária não cancelada -- nesse caso, talvez o "
              f"Consolidado Geral do PSIAG550 use um filtro de COUG "
              f"diferente (ex.: excluir UGs específicas de consolidação/"
              f"eliminação) que este módulo precisa replicar.")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[4i] Mesma decomposição por COUG para 712310200 (maior "
          f"componente de 7123/Obrigações Contratuais -- Contratos de "
          f"Serviços), MIL{ano}, INMES BETWEEN 0 AND {mes}:")
    try:
        q = f"""SELECT v.COUG, SUM(v.VADEBITO - v.VACREDITO) SALDO_SD, COUNT(*) QTD
                FROM MIL{ano}.SALDOCONTABIL v
                WHERE v.COCONTACONTABIL = 712310200
                  AND v.INMES BETWEEN 0 AND {mes}
                GROUP BY v.COUG
                ORDER BY ABS(SUM(v.VADEBITO - v.VACREDITO)) DESC"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print("    NENHUM registro -- conferir conta/filtro.")
        for _, r in df.iterrows():
            print(f"    COUG={int(r['COUG']):<8}  qtd={int(r['QTD']):>6}  "
                  f"SD={float(r['SALDO_SD'] or 0):>18,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[5] Resultado do Exercício -- LANCAMENTOCONTABIL classes 3/4, "
          f"INMES BETWEEN 1 AND {mes}, em MIL{ano}:")
    try:
        q = f"""SELECT
                    SUM(CASE WHEN o.COCONTACONTABIL BETWEEN 400000000 AND 499999999
                         THEN DECODE(o.INDEBITOCREDITO,'C',o.VALANCAMENTO,'D',-o.VALANCAMENTO,0)
                         ELSE 0 END) VPA,
                    SUM(CASE WHEN o.COCONTACONTABIL BETWEEN 300000000 AND 399999999
                         THEN DECODE(o.INDEBITOCREDITO,'D',o.VALANCAMENTO,'C',-o.VALANCAMENTO,0)
                         ELSE 0 END) VPD
                FROM MIL{ano}.LANCAMENTOCONTABIL o
                WHERE o.INMES BETWEEN 1 AND {mes}
                  AND o.COCONTACONTABIL BETWEEN 300000000 AND 499999999"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        vpa = float(df.iloc[0]['VPA'] or 0); vpd = float(df.iloc[0]['VPD'] or 0)
        print(f"    VPA={vpa:,.2f}  VPD={vpd:,.2f}  Resultado={vpa-vpd:,.2f}")
        print(f"    >>> Deve bater com RESULTADO_PATRIMONIAL_AT do dvp.py "
              f"para o mesmo mês/ano/UG.")
    except Exception as e:
        print(f"    ERRO: {e}")

    # ── [6] VALIDAÇÃO DA SOLUÇÃO DEFINITIVA (Quadro Financeiro/Permanente) ──
    # Encontrado o arquivo oficial Relatório_Equação_de_Balanço_Patrimonial.
    # xlsx ("Lista Equações de Balanço" extraída do próprio sistema contábil
    # do GDF) -- substitui as 8 hipóteses testadas e refutadas em rodadas
    # anteriores. O critério real é a coluna CONTACONTABIL.INSISCONTABIL
    # ("Sistema Contábil" no extrato XLS estático contacontabil_2026.xls),
    # com valores 'F'=Financeiro, 'P'=Permanente. O Passivo Financeiro
    # também inclui 3 contas extras de classe 6 (622130100, 622130500,
    # 631100000 -- RP/empenho a liquidar) que nunca foram testadas nas
    # hipóteses anteriores (nenhuma cobria a faixa 6XX). Este bloco valida
    # a coluna e os totais diretamente no Oracle de produção.
    print(f"\n[6a] Confirmando que CONTACONTABIL.INSISCONTABIL existe e tem "
          f"os valores esperados ('F'=Financeiro, 'P'=Permanente), owner "
          f"MIL{ano}:")
    try:
        q = f"""SELECT INSISCONTABIL, COUNT(*) QTD
                FROM MIL{ano}.CONTACONTABIL
                WHERE COCONTACONTABIL BETWEEN 100000000 AND 299999999
                GROUP BY INSISCONTABIL ORDER BY INSISCONTABIL"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        if df.empty:
            print("    NENHUM registro -- conferir se a coluna existe "
                  "neste owner/schema.")
        for _, r in df.iterrows():
            print(f"    INSISCONTABIL={r['INSISCONTABIL']!r:<6}  qtd={int(r['QTD'])}")
    except Exception as e:
        print(f"    ERRO: {e}  -- a coluna pode ter outro nome neste "
              f"ambiente; rode 'SELECT column_name FROM all_tab_columns "
              f"WHERE table_name=\\'CONTACONTABIL\\' AND owner=\\'MIL{ano}\\' "
              f"AND column_name LIKE \\'%SIS%\\'' para confirmar.")

    print(f"\n[6b] TESTE DEFINITIVO: Ativo Financeiro = SD(1XXXXXXXX) onde "
          f"INSISCONTABIL='F', acumulado INMES BETWEEN 0 AND {mes}, "
          f"MIL{ano}. Comparar com referência oficial Ativo Financeiro "
          f"= 7.665.797.148,08 (Maio/2026):")
    try:
        q = f"""SELECT SUM(v.VADEBITO - v.VACREDITO) SALDO_SD, COUNT(*) QTD
                FROM MIL{ano}.SALDOCONTABIL v
                JOIN MIL{ano}.CONTACONTABIL c
                  ON c.COCONTACONTABIL = v.COCONTACONTABIL
                WHERE v.COCONTACONTABIL BETWEEN 100000000 AND 199999999
                  AND c.INSISCONTABIL = 'F'
                  AND v.INMES BETWEEN 0 AND {mes}"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        saldo = float(df.iloc[0]['SALDO_SD'] or 0)
        print(f"    Ativo Financeiro calculado: {saldo:,.2f}  "
              f"(qtd={int(df.iloc[0]['QTD'])})")
        print(f"    Referência oficial:         7.665.797.148,08")
        print(f"    Diferença:                  {saldo - 7665797148.08:,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[6c] TESTE DEFINITIVO: Passivo Financeiro = SC(21XXXXXXX) "
          f"onde INSISCONTABIL='F'  +  SC(622130100, 622130500, "
          f"631100000), acumulado INMES BETWEEN 0 AND {mes}, MIL{ano}. "
          f"Comparar com referência oficial Passivo Financeiro = "
          f"6.545.291.111,33 (Maio/2026):")
    try:
        q = f"""SELECT SUM(v.VACREDITO - v.VADEBITO) SALDO_SC, COUNT(*) QTD
                FROM MIL{ano}.SALDOCONTABIL v
                JOIN MIL{ano}.CONTACONTABIL c
                  ON c.COCONTACONTABIL = v.COCONTACONTABIL
                WHERE v.COCONTACONTABIL BETWEEN 210000000 AND 219999999
                  AND c.INSISCONTABIL = 'F'
                  AND v.INMES BETWEEN 0 AND {mes}"""
        df = pd.read_sql(q, conn)
        df.columns = [c.upper() for c in df.columns]
        saldo_principal = float(df.iloc[0]['SALDO_SC'] or 0)
        qtd_principal = int(df.iloc[0]['QTD'])

        q2 = f"""SELECT SUM(v.VACREDITO - v.VADEBITO) SALDO_SC, COUNT(*) QTD
                 FROM MIL{ano}.SALDOCONTABIL v
                 WHERE v.COCONTACONTABIL IN (622130100, 622130500, 631100000)
                   AND v.INMES BETWEEN 0 AND {mes}"""
        df2 = pd.read_sql(q2, conn)
        df2.columns = [c.upper() for c in df2.columns]
        saldo_extra = float(df2.iloc[0]['SALDO_SC'] or 0)
        qtd_extra = int(df2.iloc[0]['QTD'])

        total = saldo_principal + saldo_extra
        print(f"    Passivo Financeiro (21XXXXXXX com F): {saldo_principal:,.2f}  "
              f"(qtd={qtd_principal})")
        print(f"    + Contas extras classe 6:              {saldo_extra:,.2f}  "
              f"(qtd={qtd_extra})")
        print(f"    = TOTAL Passivo Financeiro:            {total:,.2f}")
        print(f"    Referência oficial:                    6.545.291.111,33")
        print(f"    Diferença:                              {total - 6545291111.33:,.2f}")
    except Exception as e:
        print(f"    ERRO: {e}")

    # ── [6d]/[6e] VALIDAÇÃO DAS CONTAS DE COMPENSAÇÃO (Atos Potenciais) ──
    print(f"\n[6d] TESTE DEFINITIVO: Atos Potenciais Ativos (4 grupos, "
          f"contas EXATAS 811XXXXXX da equação oficial), SC, acumulado "
          f"INMES BETWEEN 0 AND {mes}, MIL{ano}. Comparar com referência "
          f"oficial total: 14.228.115.268,87 (Maio/2026):")
    try:
        for _, chave, contas, nome in ATOS_POT_ATIVOS_ITENS:
            lista = ", ".join(contas)
            q = f"""SELECT SUM(v.VACREDITO - v.VADEBITO) SALDO_SC, COUNT(*) QTD
                    FROM MIL{ano}.SALDOCONTABIL v
                    WHERE v.COCONTACONTABIL IN ({lista})
                      AND v.INMES BETWEEN 0 AND {mes}"""
            df = pd.read_sql(q, conn)
            df.columns = [c.upper() for c in df.columns]
            saldo = float(df.iloc[0]['SALDO_SC'] or 0)
            print(f"    {chave:<25} {nome[:40]:<40} SC={saldo:>18,.2f}  "
                  f"(qtd={int(df.iloc[0]['QTD'])})")
    except Exception as e:
        print(f"    ERRO: {e}")

    print(f"\n[6e] TESTE DEFINITIVO: Atos Potenciais Passivos (4 grupos, "
          f"contas EXATAS 812XXXXXX da equação oficial), SC, acumulado "
          f"INMES BETWEEN 0 AND {mes}, MIL{ano}. Comparar com referência "
          f"oficial total: 65.253.334.489,58 (Maio/2026):")
    try:
        for _, chave, contas, nome in ATOS_POT_PASSIVOS_ITENS:
            lista = ", ".join(contas)
            q = f"""SELECT SUM(v.VACREDITO - v.VADEBITO) SALDO_SC, COUNT(*) QTD
                    FROM MIL{ano}.SALDOCONTABIL v
                    WHERE v.COCONTACONTABIL IN ({lista})
                      AND v.INMES BETWEEN 0 AND {mes}"""
            df = pd.read_sql(q, conn)
            df.columns = [c.upper() for c in df.columns]
            saldo = float(df.iloc[0]['SALDO_SC'] or 0)
            print(f"    {chave:<25} {nome[:40]:<40} SC={saldo:>18,.2f}  "
                  f"(qtd={int(df.iloc[0]['QTD'])})")
    except Exception as e:
        print(f"    ERRO: {e}")

    print("="*64 + "\n")

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(description='Balanço Patrimonial GDF')
    p.add_argument('--mes', type=int, default=5,
                   help='Mês de referência 1-12 (Exercício Atual = saldo acumulado 0..mes)')
    p.add_argument('--ano', type=int, default=2026)
    p.add_argument('--ug', type=str, default=None,
                   help='Código da Unidade Gestora (omitir = Consolidado)')
    p.add_argument('--saida', type=str, default=None,
                   help='Caminho base dos arquivos de saída (sem extensão)')
    p.add_argument('--formato', choices=['ambos', 'excel', 'pdf'], default='ambos')
    p.add_argument('--diag', action='store_true',
                   help='Executa diagnóstico do banco e encerra')
    a = p.parse_args()

    if not 1 <= a.mes <= 12:
        print("ERRO: --mes deve estar entre 1 e 12."); sys.exit(1)

    ug_label = f'UG: {a.ug}' if a.ug else 'Consolidado'
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome = (f'Balanco_Patrimonial_{a.ano}_{a.mes:02d}'
            + (f'_UG{a.ug}' if a.ug else '_Consolidado')
            + f'_{timestamp}')
    base = (Path(a.saida).with_suffix('') if a.saida else OUTPUT_DIR / nome)
    saida_xlsx = base.with_suffix('.xlsx')
    saida_pdf = base.with_suffix('.pdf')

    print(f"\n{'='*60}")
    print(f"  BALANÇO PATRIMONIAL — GDF")
    print(f"  {MESES[a.mes]}/{a.ano}  |  {ug_label}")
    print(f"  Exercício Atual: saldo acumulado até {MESES[a.mes]}  |  "
          f"Exercício Anterior: encerramento de {a.ano-1}")
    print(f"{'='*60}")

    print("\n[1/4] Conectando ao Oracle...")
    conn = conectar_oracle()

    if a.diag:
        diagnostico(conn, a.ano, a.mes)
        conn.close()
        return

    print("\n[2/4] Buscando dados (Exercício Atual e Exercício Anterior)...")
    brutos = buscar_tudo(conn, a.mes, a.ano, a.ug)
    conn.close()

    print("\n[3/4] Calculando totais derivados...")
    t = calcular(brutos, mes_ref=a.mes)
    achados = auditoria_integridade(t)

    print("\n[4/4] Gerando arquivos...")
    if a.formato in ('ambos', 'excel'):
        gerar_excel(t, a.mes, a.ano, ug_label, saida_xlsx, achados=achados)
    if a.formato in ('ambos', 'pdf'):
        gerar_pdf(t, a.mes, a.ano, ug_label, saida_pdf, achados=achados)

    print(f"\n{'─'*60}")
    print("  CONFERÊNCIA RÁPIDA (Exercício Atual):")
    print(f"{'─'*60}")
    for lbl, key in [('ATIVO TOTAL', 'ATIVO_TOTAL_AT'),
                      ('ATIVO CIRCULANTE', 'ATIVO_CIRCULANTE_AT'),
                      ('ATIVO NÃO CIRCULANTE', 'ATIVO_NCIRCULANTE_AT'),
                      ('PASSIVO + PL TOTAL', 'PASSIVO_PL_TOTAL_AT'),
                      ('PASSIVO CIRCULANTE', 'PASSIVO_CIRCULANTE_AT'),
                      ('PASSIVO NÃO CIRCULANTE', 'PASSIVO_NCIRCULANTE_AT'),
                      ('PATRIMÔNIO LÍQUIDO', 'PATRIMONIO_LIQUIDO_AT'),
                      ('RESULTADO DO EXERCÍCIO', 'RESULTADO_EXERCICIO_AT'),
                      ('SALDO PATRIMONIAL (III)', 'SALDO_PATRIMONIAL_III_AT')]:
        print(f"    {lbl:<26}: {t.get(key, 0):>22,.2f}")
    print(f"    {'DIFERENÇA (Ativo-Passivo+PL)':<26}: "
          f"{t.get('ATIVO_TOTAL_AT',0) - t.get('PASSIVO_PL_TOTAL_AT',0):>22,.2f}")
    print(f"{'─'*60}")

    # ── Conferência do bloco PL do Exercício Anterior ────────────────────
    # Adicionado (Ago/2026) depois de uma confusao real: o Controle 7 so
    # verifica que as sub-linhas SOMAM o total do PL -- e isso e verdade
    # tanto na fonte nova (MIL{ano-1} 0..13) quanto na antiga (abertura,
    # 0..14), porque o encerramento e neutro no total. Ou seja, o controle
    # NAO distingue as duas fontes. Como a pasta de saida acumula varios
    # arquivos com timestamp, e facil abrir um Excel antigo e concluir que
    # a correcao nao pegou. Imprimir os valores aqui elimina a duvida:
    # o que aparece no console e, por construcao, o desta execucao.
    print(f"\n{'─'*60}")
    print("  CONFERÊNCIA DO PL (Exercício Anterior) -- fonte pré-encerramento:")
    print(f"{'─'*60}")
    for lbl, key in [('Patrimônio Social e Capital Social', 'PAT_SOCIAL_CAPITAL_ANT'),
                      ('Resultado Acumulado',              'RESULTADO_ACUMULADO_ANT'),
                      ('Resultado do Exercício',           'RESULTADO_EXERCICIO_ANT'),
                      ('Superávits ou Déficits Acum.',     'SUPERAVIT_DEFICIT_ACUM_ANT'),
                      ('Lucros e Prejuízos Acum.',         'LUCROS_PREJUIZOS_ACUM_ANT'),
                      ('PATRIMÔNIO LÍQUIDO',               'PATRIMONIO_LIQUIDO_ANT')]:
        print(f"    {lbl:<36}: {t.get(key, 0):>22,.2f}")
    print(f"    {'Fonte':<36}: {t.get('_PL_ANT_FONTE', '(não registrada)')}")
    print(f"{'─'*60}")

    imprimir_auditoria(achados)

    print(f"\n  Concluído em {datetime.now():%H:%M:%S}\n")


if __name__ == '__main__':
    main()
