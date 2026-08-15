# Diagnóstico Contábil GDF

❌ **1 erro(s)** — competência **07/2026** · escopo **Consolidado** · execução de 14/08/2026 13:04

`1 erro` · `0 alerta` · `27 ok` · `4 info` · **32 controles**

## Indicadores

| Indicador | Valor |
|---|---:|
| Caixa Final | R$ 7.438.503.244,54 |
| Receita Realizada | R$ 27.942.220.303,94 |
| Despesa Orçamentária Paga | R$ 21.699.922.160,81 |
| Ingressos | R$ 60.629.833.177,82 |
| Dispêndios | R$ 60.629.833.177,82 |
| Despesa Empenhada | R$ 27.161.590.016,16 |
| Despesa Liquidada | R$ 23.617.140.264,77 |
| Ativo Total | R$ 69.920.399.167,59 |
| Patrimônio Líquido | −R$ 146.662.549.672,12 |
| Resultado do Exercício | −R$ 15.023.520.666,13 |

## Controles — última execução

| | Módulo | Código | Controle | Diferença |
|:-:|---|---|---|---:|
| ✅ | BF | `BF-01` | Ingressos = Dispêndios | R$ 60.629.833.177,82 |
| ✅ | BF | `BF-02` | Transf. Recebidas = Transf. Concedidas | R$ 24.226.641.217,62 |
| ✅ | BF | `BF-03` | Caixa Final positivo | R$ 7.438.503.244,54 |
| ✅ | BF | `BF-04` | Caixa 111XXXXXX: razão = view de saldos | R$ 7.438.503.244,54 |
| ℹ️ | BF | `BF-05` | Composição do Balanço Financeiro | R$ 60.629.833.177,82 |
| ✅ | BO | `BO-01` | Empenhada ≥ Liquidada e Empenhada ≥ Paga | — |
| ✅ | BO | `BO-02` | Despesa Empenhada ≤ Dotação Atualizada | — |
| ✅ | DFC | `DFC-01` | Geracao Liquida (I+II+III) = Variacao do Caixa | R$ 1.978.173.702,80 |
| ✅ | DFC | `DFC-02a` | Caixa Inicial positivo | R$ 5.460.329.541,74 |
| ✅ | DFC | `DFC-02b` | Caixa Final positivo | R$ 7.438.503.244,54 |
| ❌ | DVP | `DVP-01` | Resultado Patrimonial ≠ conta de encerramento | −R$ 15.023.776.682,61 |
| ✅ | DVP | `DVP-02a` | VPA Total positivo | R$ 93.435.180.824,12 |
| ✅ | DVP | `DVP-02b` | VPD Total positivo | R$ 108.458.701.490,25 |
| ℹ️ | DVP | `DVP-03` | Cobertura dos subtotais detalhados | R$ 93.779.171.524,67 |
| ✅ | DMPL | `DMPL-01` | Fechamento DMPL: Saldo Ini + Movimento (razão) = Saldo Fin | — |
| ℹ️ | DMPL | `DMPL-02` | Mutações do PL no período (exceto Resultado do Exercício) | — |
| ℹ️ | DMPL | `DMPL-03` | Resultado do Exercício (DMPL) | — |
| ✅ | BP | `BP-01` | Ativo = Passivo + PL | R$ 69.920.399.167,59 |
| ✅ | BP | `BP-02` | Ativo Circ + Não Circ esgota a classe 1 | R$ 69.920.399.167,59 |
| ✅ | BP | `BP-03` | Caixa positivo | R$ 7.438.503.244,54 |
| ✅ | BP | `BP-04` | Passivo + PL esgota a classe 2 | R$ 84.943.919.833,72 |
| ✅ | X | `X1a` | Caixa Inicial: BF = DFC | R$ 5.460.329.541,74 |
| ✅ | X | `X1b` | Caixa Final: BF = DFC | R$ 7.438.503.244,54 |
| ✅ | X | `X2` | DFC: Geração Líquida (I+II+III) = Variação do Caixa | R$ 0,00 |
| ✅ | X | `X6` | Resultado do Exercício: DMPL = DVP | −R$ 15.023.520.666,13 |
| ✅ | X | `R1a` | Resultado do Exercício: BP = DVP(III) | −R$ 15.023.520.666,13 |
| ✅ | X | `R2` | Saldo Final DMPL = PL do BP | −R$ 146.662.549.672,12 |
| ✅ | X | `R4a` | Caixa: BP = DFC (Caixa Final) | R$ 7.438.503.244,54 |
| ✅ | X | `R7` | Receita Realizada: BF = BO | R$ 27.942.220.303,94 |
| ✅ | X | `R8` | Despesa Paga: BF = BO | R$ 21.699.922.160,81 |
| ✅ | X | `R10a` | Pgtos RPNP: BF = BO | R$ 1.075.676.548,60 |
| ✅ | X | `R10b` | Pgtos RPP: BF = BO | R$ 2.283.574.211,13 |

## O que não passou

**❌ `DVP-01` Resultado Patrimonial ≠ conta de encerramento**  
Diferença: -15,023,776,682.61  |  VPA − VPD -15,023,520,666.13  vs 891XXXXXX 256,016.48

## Histórico (2 execuções)

Uma coluna por execução, da mais antiga à mais recente. Célula vazia = o controle não existia naquela execução.

| Controle | 13/08 | 14/08 |
|---|---|---|
| `BF-01` | ✅ | ✅ |
| `BF-02` | ✅ | ✅ |
| `BF-03` | ✅ | ✅ |
| `BF-04` | ✅ | ✅ |
| `BF-05` | ℹ️ | ℹ️ |
| `BO-01` | ✅ | ✅ |
| `BO-02` | ✅ | ✅ |
| `DFC-01` | ✅ | ✅ |
| `DFC-02a` | ✅ | ✅ |
| `DFC-02b` | ✅ | ✅ |
| `DVP-01` | ❌ | ❌ |
| `DVP-02a` | ✅ | ✅ |
| `DVP-02b` | ✅ | ✅ |
| `DVP-03` | ℹ️ | ℹ️ |
| `DMPL-01` | ✅ | ✅ |
| `DMPL-02` | ℹ️ | ℹ️ |
| `DMPL-03` | ℹ️ | ℹ️ |
| `BP-01` | ✅ | ✅ |
| `BP-02` | ✅ | ✅ |
| `BP-03` | ✅ | ✅ |
| `BP-04` | ✅ | ✅ |
| `X1a` | ✅ | ✅ |
| `X1b` | ✅ | ✅ |
| `X2` | ✅ | ✅ |
| `X6` | ✅ | ✅ |
| `R1a` | ✅ | ✅ |
| `R2` | ✅ | ✅ |
| `R4a` | ✅ | ✅ |
| `R7` | ✅ | ✅ |
| `R8` | ✅ | ✅ |
| `R10a` | ✅ | ✅ |
| `R10b` | ✅ | ✅ |

---

Gerado por `painel.py` em 15/08/2026 11:13 · 2 execuções no histórico · dados em `painel/dados/`.

> Repositório privado. Não habilitar GitHub Pages: no plano gratuito, o Pages de um repositório privado publica o conteúdo na web aberta.
