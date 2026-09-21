-- C06b — Variantes de 218914XXX / 218924XXX fora do escopo do C06
-- O C06 verifica apenas 218914002 e 218924002.
-- Este controle detecta outras subcontas dessas famílias com saldo ≠ 0
-- que poderiam ampliar o escopo do C06.
SELECT
    L.COCONTACONTABIL                                                          AS "Conta Contabil",
    NVL(V.NOCONTACONTABIL, '(sem cadastro)')                                   AS "Nome Conta",
    SUM(DECODE(L.INDEBITOCREDITO,'D',L.VALANCAMENTO,'C',-L.VALANCAMENTO))     AS "Saldo",
    COUNT(DISTINCT L.COUG)                                                     AS "N. UGs"
FROM MIL2026.LANCAMENTOCONTABIL L
LEFT JOIN MIL2026.VCONTACONTABIL V ON V.COCONTACONTABIL = L.COCONTACONTABIL
WHERE (
      L.COCONTACONTABIL BETWEEN 218914000 AND 218914999
   OR L.COCONTACONTABIL BETWEEN 218924000 AND 218924999
  )
  AND L.COCONTACONTABIL NOT IN (218914002, 218924002)
GROUP BY L.COCONTACONTABIL, NVL(V.NOCONTACONTABIL, '(sem cadastro)')
HAVING SUM(DECODE(L.INDEBITOCREDITO,'D',L.VALANCAMENTO,'C',-L.VALANCAMENTO)) <> 0
ORDER BY L.COCONTACONTABIL
