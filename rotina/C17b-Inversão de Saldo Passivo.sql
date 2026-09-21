-- C17b — Inversão de Saldo Classe 2 (Passivo com saldo devedor)
-- Complementa C17 (que só verifica Classe 1 — Ativo).
-- Contas de Passivo têm natureza credora; saldo devedor (VACREDITO < VADEBITO) é anômalo.
-- Possíveis causas: baixa > saldo inscrito, pagamento sem empenho/liquidação,
--   lançamento invertido ou conta usada indevidamente como ativo.
SELECT
    O978241.ININVERSAOSALDO                                  AS "Ind. Inversao Saldo",
    O996998.COCONTACONTABIL                                  AS "Conta Contabil",
    O996998.COCONTACORRENTE                                  AS "Conta Corrente",
    O996998.VACREDITO - O996998.VADEBITO                     AS "Saldo",
    O996998.COUG                                             AS "UG"
FROM MIL2026.VCONTACONTABIL O978241
JOIN MIL2026.SALDOCONTABIL O996998
  ON O978241.COCONTACONTABIL = O996998.COCONTACONTABIL
WHERE O978241.INSALDOCONTABIL = 'C'
  AND O978241.ININVERSAOSALDO  = 'N'
  AND O978241.COCONTACONTABIL  BETWEEN 200000000 AND 299999999
  AND O996998.INMES             = {MES}
  AND O996998.VACREDITO - O996998.VADEBITO < 0
ORDER BY O996998.COUG, O996998.COCONTACONTABIL, O996998.COCONTACORRENTE
