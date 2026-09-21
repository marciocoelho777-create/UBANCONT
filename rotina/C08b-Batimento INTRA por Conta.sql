-- C08b — Batimento INTRA por conta contábil (NET entre todas as UGs)
-- Para cada conta intragovernamental (5.º dígito = 2, classes 1–4),
-- a soma dos movimentos de TODAS as UGs deve ser zero:
--   se UG A debitou a conta X (INTRA), UG B deve ter creditado a mesma conta.
-- Saldo ≠ 0 = uma UG registrou sem a contrapartida da UG parceira — erro real.
-- Ordenado pelo maior desequilíbrio absoluto.
SELECT
    L.COCONTACONTABIL                                                         AS "Conta Contabil",
    NVL(V.NOCONTACONTABIL, '(sem cadastro)')                                  AS "Nome Conta",
    COUNT(DISTINCT CASE
        WHEN DECODE(L.INDEBITOCREDITO,'D',L.VALANCAMENTO,'C',-L.VALANCAMENTO) > 0
        THEN L.COUG END)                                                      AS "UGs saldo +",
    COUNT(DISTINCT CASE
        WHEN DECODE(L.INDEBITOCREDITO,'D',L.VALANCAMENTO,'C',-L.VALANCAMENTO) < 0
        THEN L.COUG END)                                                      AS "UGs saldo -",
    SUM(DECODE(L.INDEBITOCREDITO,'D',L.VALANCAMENTO,'C',-L.VALANCAMENTO))    AS "Saldo Liquido Total"
FROM MIL2026.LANCAMENTOCONTABIL L
LEFT JOIN MIL2026.VCONTACONTABIL V ON V.COCONTACONTABIL = L.COCONTACONTABIL
WHERE SUBSTR(L.COCONTACONTABIL, 5, 1) = '2'
  AND L.COCONTACONTABIL BETWEEN 100000000 AND 499999999
GROUP BY L.COCONTACONTABIL, NVL(V.NOCONTACONTABIL, '(sem cadastro)')
HAVING SUM(DECODE(L.INDEBITOCREDITO,'D',L.VALANCAMENTO,'C',-L.VALANCAMENTO)) <> 0
ORDER BY ABS(SUM(DECODE(L.INDEBITOCREDITO,'D',L.VALANCAMENTO,'C',-L.VALANCAMENTO))) DESC
