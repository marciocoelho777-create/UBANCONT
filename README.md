# Diagnóstico Contábil GDF

✅ **Todos os controles passaram** — competência **07/2026** · escopo **Consolidado** · execução de 13/08/2026 16:50

`0 erro` · `0 alerta` · `27 ok` · `3 info` · **30 controles**

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

| | Módulo | Código | Controle |
|:-:|---|---|---|
| ✅ | BF | `BF-01` | Ingressos = Dispêndios |
| ✅ | BF | `BF-02` | Transf. Recebidas = Transf. Concedidas |
| ✅ | BF | `BF-03` | Caixa Final positivo |
| ✅ | BF | `BF-04` | Caixa 111XXXXXX: razão = view de saldos |
| ℹ️ | BF | `BF-05` | Composição do Balanço Financeiro |
| ✅ | BO | `BO-01` | Empenhada ≥ Liquidada e Empenhada ≥ Paga |
| ✅ | BO | `BO-02` | Despesa Empenhada ≤ Dotação Atualizada |
| ✅ | DFC | `DFC-01` | Geracao Liquida (I+II+III) = Variacao do Caixa |
| ✅ | DFC | `DFC-02a` | Caixa Inicial positivo |
| ✅ | DFC | `DFC-02b` | Caixa Final positivo |
| ✅ | DVP | `DVP-01` | Resultado Patrimonial = VPA − VPD |
| ✅ | DVP | `DVP-02a` | VPA Total positivo |
| ✅ | DVP | `DVP-02b` | VPD Total positivo |
| ✅ | DMPL | `DMPL-01` | Fechamento DMPL: Saldo Ini + Movimento (razão) = Saldo Fin |
| ℹ️ | DMPL | `DMPL-02` | Mutações do PL no período (exceto Resultado do Exercício) |
| ℹ️ | DMPL | `DMPL-03` | Resultado do Exercício (DMPL) |
| ✅ | BP | `BP-01` | Ativo = Passivo + PL |
| ✅ | BP | `BP-02` | Ativo Total = Circ + Não Circ |
| ✅ | BP | `BP-03` | Caixa positivo |
| ✅ | X | `X1a` | Caixa Inicial: BF = DFC |
| ✅ | X | `X1b` | Caixa Final: BF = DFC |
| ✅ | X | `X2` | DFC: Geração Líquida (I+II+III) = Variação do Caixa |
| ✅ | X | `X6` | Resultado do Exercício: DMPL = DVP |
| ✅ | X | `R1a` | Resultado do Exercício: BP = DVP(III) |
| ✅ | X | `R2` | Saldo Final DMPL = PL do BP |
| ✅ | X | `R4a` | Caixa: BP = DFC (Caixa Final) |
| ✅ | X | `R7` | Receita Realizada: BF = BO |
| ✅ | X | `R8` | Despesa Paga: BF = BO |
| ✅ | X | `R10a` | Pgtos RPNP: BF = BO |
| ✅ | X | `R10b` | Pgtos RPP: BF = BO |

## Histórico (1 execução)

Uma coluna por execução, da mais antiga à mais recente. Célula vazia = o controle não existia naquela execução.

| Controle | 13/08 |
|---|---|
| `BF-01` | ✅ |
| `BF-02` | ✅ |
| `BF-03` | ✅ |
| `BF-04` | ✅ |
| `BF-05` | ℹ️ |
| `BO-01` | ✅ |
| `BO-02` | ✅ |
| `DFC-01` | ✅ |
| `DFC-02a` | ✅ |
| `DFC-02b` | ✅ |
| `DVP-01` | ✅ |
| `DVP-02a` | ✅ |
| `DVP-02b` | ✅ |
| `DMPL-01` | ✅ |
| `DMPL-02` | ℹ️ |
| `DMPL-03` | ℹ️ |
| `BP-01` | ✅ |
| `BP-02` | ✅ |
| `BP-03` | ✅ |
| `X1a` | ✅ |
| `X1b` | ✅ |
| `X2` | ✅ |
| `X6` | ✅ |
| `R1a` | ✅ |
| `R2` | ✅ |
| `R4a` | ✅ |
| `R7` | ✅ |
| `R8` | ✅ |
| `R10a` | ✅ |
| `R10b` | ✅ |

---

Gerado por `painel.py` em 13/08/2026 17:05 · 1 execução no histórico · dados em `painel/dados/`.

> Repositório privado. Não habilitar GitHub Pages: no plano gratuito, o Pages de um repositório privado publica o conteúdo na web aberta.
