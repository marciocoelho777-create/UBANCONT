# COMO REPETIR O RELATÓRIO DE INTEGRIDADE TODO MÊS

## Resumo em 3 passos (fluxo real, com Claude extraindo os PDFs)

1. **Anexe os PDFs do mês novo** no projeto/conversa com Claude (mesmo padrão
   de nomes já usado: `ListaBalancoPatrimonial_07.pdf`,
   `ListaBalancoOrcamentario_07.pdf`, `ListaBalancoFinanceiro_07.pdf`,
   `ListaFluxoCaixa_07.pdf`, `ListaVariacaoPatrimonial_07.pdf`,
   `ListaDMPL_07.pdf` — troque "07" pelo número do mês de referência).
2. **Peça a Claude para atualizar a planilha** com algo como: *"Anexei os
   PDFs do mês 07. Atualize a dados_mensais.xlsx e gere os relatórios."*
   Claude lê os PDFs, extrai os valores e roda `atualizar_mes.py` +
   `relatorio_integridade.py` por você.
3. **Baixe os arquivos atualizados**: `dados_mensais.xlsx` (com o histórico
   acumulado), `Relatorio_Integridade.xlsx` e `Relatorio_Integridade.pdf`.

Você não precisa copiar nenhum número manualmente, nem abrir o Excel para
digitar valores. A planilha existe para guardar o histórico e para você (ou
qualquer pessoa da equipe) conseguir auditar exatamente o que foi extraído de
cada PDF — não para preenchimento manual rotineiro.

---

## Como funciona por baixo dos panos

Quando você anexa os PDFs e pede a atualização, Claude:

1. Lê os PDFs do mês novo (igual fez na primeira extração do histórico).
2. Chama `atualizar_mes.py`, passando os valores extraídos — esse script
   adiciona uma coluna nova em cada aba de `dados_mensais.xlsx` **sem apagar
   meses anteriores** (e se o mês já existir, atualiza em vez de duplicar).
3. Roda `relatorio_integridade.py`, que lê a planilha atualizada, aplica as
   32 Regras de Integridade e regrava `Relatorio_Integridade.xlsx` e
   `Relatorio_Integridade.pdf`.

Isso significa que `dados_mensais.xlsx` funciona como um **banco de dados
histórico auditável**: cada número nele tem uma "Onde encontrar no PDF"
correspondente, então se você (ou um auditor) quiser confirmar de onde veio
um valor, basta abrir a planilha e comparar com o PDF original.

---

## Alternativa: preenchimento manual (se preferir não usar o chat)

Se em algum mês você preferir preencher você mesmo (por exemplo, sem acesso
ao Claude naquele momento), ainda é possível editar `dados_mensais.xlsx`
manualmente — o formato continua o mesmo: linhas = campos, colunas = meses,
com a coluna "Onde encontrar no PDF" como guia. Depois, rode:

```
python3 relatorio_integridade.py
```

Os passos detalhados desse caminho manual estão na seção "Passo a passo
detalhado (preenchimento manual)" mais abaixo.

---

## Passo a passo detalhado (preenchimento manual)

Abra `dados_mensais.xlsx`. Ela tem uma aba por demonstrativo oficial:

| Aba                | Demonstrativo                                  |
|--------------------|--------------------------------------------------|
| `BP`               | Balanço Patrimonial                             |
| `Anexo_BP`         | Quadro dos Ativos e Passivos Financ. e Permanentes |
| `BO`               | Balanço Orçamentário                            |
| `BF`               | Balanço Financeiro                              |
| `DFC`              | Demonstração dos Fluxos de Caixa                |
| `DVP`              | Demonstração das Variações Patrimoniais         |
| `DVP_Qualitativas` | Anexo da DVP — Variações Patrimoniais Qualitativas |
| `DMPL`             | Demonstração das Mutações do Patrimônio Líquido |

Em cada aba (exceto DMPL), as **linhas são os campos** e as **colunas são os
meses**. Para adicionar um mês novo:

1. Vá até a primeira coluna vazia à direita.
2. No cabeçalho (linha 1), escreva algo como `Mês 07 (Julho)` — o número do
   mês é o que importa (`07`); o texto entre parênteses é só decorativo.
3. Preencha os valores numéricos de cada linha, usando a coluna **"Onde
   encontrar no PDF"** como guia de onde tirar o número no demonstrativo
   oficial.
4. Repita o mesmo processo nas outras 6 abas simples.
5. Na aba `DMPL` (que tem um layout diferente — 7 linhas de movimentação x 9
   colunas por mês), copie um bloco de mês existente e cole abaixo do
   último bloco, ajustando o cabeçalho `MÊS XX (...)` e preenchendo os
   valores da tabela da DMPL daquele mês.

**Dica:** se algum valor não estiver disponível no PDF daquele mês, **deixe a
célula em branco** — não coloque `0`. O motor de regras trata célula vazia
como "dado não disponível" e marca a verificação correspondente como **NÃO
VERIFICÁVEL** em vez de **DIVERGENTE**, e isso aparece com a explicação certa
no relatório, então não há erro nem confusão.

### 2. Salvar e rodar o script

Salve `dados_mensais.xlsx` (mantenha esse nome e mantenha-o na mesma pasta dos
arquivos `.py`). Depois, no terminal/PowerShell, na pasta do projeto:

```
python3 relatorio_integridade.py
```

O script imprime um resumo no terminal (quantas verificações ficaram OK,
DIVERGENTE, ALERTA, etc.) e gera:

- `Relatorio_Integridade.xlsx` — 4 abas (Resumo, Detalhamento, Divergências e
  Alertas, Legenda).
- `Relatorio_Integridade.pdf` — versão executiva/narrativa, pronta para
  impressão ou envio por e-mail.

### 3. Conferir o relatório

Abra qualquer um dos dois relatórios. O painel inicial já mostra quantas
regras passaram e quantas precisam de atenção. A aba/seção "Divergências e
Alertas" explica cada achado em português, com o valor de cada lado da
comparação e a diferença apurada.

---

## Perguntas frequentes

**"Esqueci de preencher um campo, o script vai quebrar?"**
Não. Campos vazios geram "NÃO VERIFICÁVEL" em vez de erro.

**"Posso reabrir e corrigir um mês antigo?"**
Sim. Anexe de novo o PDF daquele mês e peça a Claude para atualizar — o
`atualizar_mes.py` substitui os valores da coluna existente em vez de
duplicar. Se preferir, também pode editar a célula direto na planilha.

**"E se eu não tiver todos os PDFs de um mês (por exemplo, só o BP saiu
mas o BO ainda não)?"**
Sem problema — anexe o que tiver disponível. Claude preenche só os
demonstrativos enviados; os demais campos daquele mês continuam em branco
até você anexar o restante.

**"Quero adicionar uma regra nova ou mudar uma tolerância."**
Isso já exige editar `motor_regras.py` (a constante `TOL` no topo do arquivo
controla a tolerância padrão em R$). Me chame que eu ajusto para você.

**"O relatório mostrou uma regra que não existia antes (ex.: Regra 22 com
rubricas detalhadas de despesa)."**
Algumas regras (como a 22, que verifica Empenhado ≥ Liquidado ≥ Pago por
rubrica) precisam de campos extras na aba `BO` (eles já estão no template,
nas últimas linhas da aba — procure por "Empenhada", "Liquidada", "Paga" por
categoria de despesa).

**"Quero manter um arquivo só, sem vários .py separados."**
Dá para consolidar tudo em um único arquivo, como fizemos na primeira
versão — me avise se preferir esse formato; a única diferença é que aí
qualquer ajuste futuro nas regras exige editar um arquivo maior em vez de um
arquivo pequeno e específico.

---

## Arquivos do projeto e o que cada um faz

| Arquivo                     | Você edita? | O que faz |
|------------------------------|:-----------:|-----------|
| `dados_mensais.xlsx`        | Não diretamente (Claude atualiza) | Histórico auditável dos valores extraídos dos PDFs |
| `relatorio_integridade.py`  | Não | Script principal — roda tudo |
| `atualizar_mes.py`          | Não (Claude usa) | Escreve os valores extraídos de um mês na planilha, sem apagar histórico |
| `catalogo_campos.py`        | Não (raramente) | Lista de campos e onde achá-los no PDF |
| `leitor_dados_mensais.py`   | Não | Converte a planilha em dados que o motor entende |
| `motor_regras.py`           | Não (raramente) | As 32 Regras de Integridade |
| `gerar_excel.py`            | Não | Monta o relatório em Excel |
| `gerar_pdf.py`              | Não | Monta o relatório em PDF |
| `gerar_template_dados.py`   | Não (uso único) | Cria `dados_mensais.xlsx` do zero, se precisar recriá-lo |
