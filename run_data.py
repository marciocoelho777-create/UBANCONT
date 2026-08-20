# -*- coding: utf-8 -*-
"""
Roda qualquer demonstrativo contabil com filtro de data de lancamento.

O filtro AND o.DALANCAMENTO <= DATE 'AAAA-MM-DD' e injetado em todas as
queries que usam LANCAMENTOCONTABIL (alias o). Queries em VSALDOCONTABIL
nao sao afetadas (a view nao tem campo de data — apenas INMES).

Uso:
    python run_data.py bf  --data 2026-08-19          # Balanco Financeiro
    python run_data.py bo  --data 2026-08-19          # Balanco Orcamentario
    python run_data.py bp  --data 2026-08-19          # Balanco Patrimonial (aviso)
    python run_data.py dfc --data 2026-08-19          # DFC
    python run_data.py dmpl --data 2026-08-19         # DMPL
    python run_data.py dvp  --data 2026-08-19         # DVP (aviso)

    # Com parametros extras repassados ao script original:
    python run_data.py bf --data 2026-08-19 --formato pdf --ug 123456
"""
import argparse, datetime, re, runpy, sys
import pandas as _pd

# Demonstrativos que usam VSALDOCONTABIL (sem campo de data = sem filtro diario)
_SEM_LANCAMENTO = {'bp', 'dvp'}

_SCRIPTS = {
    'bf':   'mestre.py',
    'bo':   'bo.py',
    'bp':   'bp.py',
    'dfc':  'dfc.py',
    'dmpl': 'dmpl.py',
    'dvp':  'dvp.py',
}

_DATA_LIM: datetime.date | None = None
_ORIG_READ_SQL = _pd.read_sql


def _patched_read_sql(sql, *args, **kwargs):
    """Intercepta pd.read_sql e injeta filtro de data em LANCAMENTOCONTABIL."""
    if _DATA_LIM:
        data_str = _DATA_LIM.strftime('%Y-%m-%d')
        sql = re.sub(
            r'(FROM\s+MIL\d+\.LANCAMENTOCONTABIL\s+o\s+WHERE\s+)',
            r"\1o.DALANCAMENTO <= DATE '" + data_str + "' AND ",
            str(sql),
            flags=re.IGNORECASE | re.DOTALL,
        )
    return _ORIG_READ_SQL(sql, *args, **kwargs)


def main():
    global _DATA_LIM

    p = argparse.ArgumentParser(
        description='Demonstrativo contabil com filtro de data de lancamento',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='Parametros nao reconhecidos sao repassados ao script original.',
    )
    p.add_argument('demonstrativo', choices=_SCRIPTS.keys(),
                   help='bf | bo | bp | dfc | dmpl | dvp')
    p.add_argument('--data', metavar='AAAA-MM-DD',
                   help='Data limite para DALANCAMENTO (ex: 2026-08-19)')
    p.add_argument('--mes', type=int,
                   help='Mes de referencia 1-12 (derivado de --data se omitido)')
    p.add_argument('--ano', type=int, default=datetime.date.today().year)
    a, extras = p.parse_known_args()

    mes = a.mes
    if a.data:
        _DATA_LIM = datetime.date.fromisoformat(a.data)
        if mes is None:
            mes = _DATA_LIM.month
        # Aplica o patch antes de carregar o script
        _pd.read_sql = _patched_read_sql
        print(f"\n  [run_data] Filtro: DALANCAMENTO <= {a.data}")
        if a.demonstrativo in _SEM_LANCAMENTO:
            print(f"  [run_data] AVISO: '{a.demonstrativo}' usa VSALDOCONTABIL "
                  f"(sem campo de data) — filtro nao aplicado a esta demonstracao.")
    else:
        if mes is None:
            mes = datetime.date.today().month

    script = _SCRIPTS[a.demonstrativo]
    # Reconstroi sys.argv para o script original
    sys.argv = [script, '--mes', str(mes), '--ano', str(a.ano)] + extras
    print(f"  [run_data] -> {script}  --mes {mes} --ano {a.ano}"
          + (f"  --data {a.data}" if a.data else "")
          + (f"  {' '.join(extras)}" if extras else ""))
    print()

    runpy.run_path(script, run_name='__main__')


if __name__ == '__main__':
    main()
