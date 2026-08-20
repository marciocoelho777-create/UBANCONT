# -*- coding: utf-8 -*-
"""
Publica HTMLs do painel/ no repositorio UBANCONT (origin/main)
e atualiza automaticamente o index.html com links para os novos arquivos.

So arquivos .html sao enviados; scripts Python e outros arquivos ficam
apenas no repositorio local.

Uso:
    python push_painel.py                       # publica todos os HTMLs novos/alterados
    python push_painel.py painel/foo.html ...   # publica arquivos especificos
"""
import re, shutil, subprocess, sys, tempfile
from pathlib import Path

BRANCH_TMP    = "_pub_painel"
REMOTE        = "origin"
REMOTE_BRANCH = "main"
PAINEL_DIR    = Path(__file__).parent / "painel"

MESES = {
    '01': 'Janeiro',  '02': 'Fevereiro', '03': 'Marco',
    '04': 'Abril',    '05': 'Maio',      '06': 'Junho',
    '07': 'Julho',    '08': 'Agosto',    '09': 'Setembro',
    '10': 'Outubro',  '11': 'Novembro',  '12': 'Dezembro',
}

# Nomes fixos para arquivos conhecidos (href → texto do link)
NOMES_FIXOS = {
    'painel_diag.html':          'Painel de Diagnostico (33 controles de integridade)',
    'painel_classificacao.html': 'Monitor de Classificacao',
}


def run(cmd, check=True, capture=False):
    r = subprocess.run(cmd, shell=True, capture_output=capture, text=True,
                       encoding='utf-8', errors='replace')
    if check and r.returncode != 0:
        err = (r.stderr or '').strip()
        raise RuntimeError(f"Falhou ({r.returncode}): {cmd}\n{err}")
    return r


def nome_amigavel(href):
    """Gera texto de link a partir do href do arquivo HTML."""
    nome = Path(href).name
    if nome in NOMES_FIXOS:
        return NOMES_FIXOS[nome]

    # rotina_controles_AAAA_MM.html
    m = re.match(r'rotina_controles_(\d{4})_(\d{2})\.html', nome)
    if m:
        ano, mes = m.group(1), m.group(2)
        return f'Controles de Rotina -- {MESES.get(mes, mes)}/{ano}'

    # fallback: capitaliza e substitui _ por espaco
    return nome.replace('_', ' ').replace('.html', '').title()


def html_alterados():
    """HTMLs locais que nao existem ou diferem do origin/main."""
    run(f'git fetch {REMOTE} {REMOTE_BRANCH} -q')
    r = run(f'git ls-tree -r --name-only {REMOTE}/{REMOTE_BRANCH} -- painel/',
            capture=True, check=False)
    remoto = set(r.stdout.strip().splitlines())

    locais = sorted(
        str(p.relative_to(Path(__file__).parent)).replace('\\', '/')
        for p in PAINEL_DIR.glob('*.html')
    )
    alterados = []
    for arq in locais:
        if arq not in remoto:
            alterados.append(arq)
            continue
        r2 = run(f'git diff {REMOTE}/{REMOTE_BRANCH} HEAD -- {arq}',
                 capture=True, check=False)
        if r2.stdout.strip():
            alterados.append(arq)
    return alterados


def atualizar_index(html_remoto, novos_hrefs):
    """
    Recebe o conteudo atual do index.html (string) e uma lista de hrefs
    novos a adicionar. Retorna o HTML atualizado (string).
    """
    # links ja existentes no index
    existentes = set(re.findall(r'href="([^"]+)"', html_remoto))

    linhas_novas = []
    for href in novos_hrefs:
        if href not in existentes:
            texto = nome_amigavel(href)
            linhas_novas.append(f'  <li><a href="{href}">{texto}</a></li>')

    if not linhas_novas:
        return html_remoto  # nada a acrescentar

    insercao = '\n'.join(linhas_novas)
    # insere antes do </ul>
    return re.sub(r'(</ul>)', insercao + r'\n\1', html_remoto, count=1)


def publicar(arquivos):
    arquivos = [a.replace('\\', '/') for a in arquivos
                if Path(a).suffix.lower() == '.html']

    if not arquivos:
        print('Nenhum HTML para publicar.')
        return

    print(f'Publicando {len(arquivos)} arquivo(s) no UBANCONT:')
    for f in arquivos:
        print(f'  {f}')

    # copia os HTMLs para temp e remove do disco se nao rastreados (evita bloqueio no checkout)
    tmp = Path(tempfile.mkdtemp())
    for arq in arquivos:
        src = Path(arq)
        if src.exists():
            shutil.copy2(src, tmp / src.name)
            if run(f'git ls-files --error-unmatch "{arq}"', check=False).returncode != 0:
                src.unlink()  # arquivo nao rastreado: remove para nao bloquear checkout

    run(f'git fetch {REMOTE} {REMOTE_BRANCH} -q')

    # le index.html atual do remoto
    r_idx = run(f'git show {REMOTE}/{REMOTE_BRANCH}:index.html',
                capture=True, check=False)
    index_atual = r_idx.stdout if r_idx.returncode == 0 else ''

    # calcula hrefs relativos ao root do repo para os novos arquivos
    novos_hrefs = [a for a in arquivos]   # ex: painel/rotina_controles_2026_07.html
    index_novo = atualizar_index(index_atual, novos_hrefs) if index_atual else ''
    index_mudou = index_novo and index_novo != index_atual

    stashed = run('git stash -q', check=False).returncode == 0
    branch_orig = run('git branch --show-current', capture=True).stdout.strip()

    try:
        run(f'git checkout -b {BRANCH_TMP} {REMOTE}/{REMOTE_BRANCH} -q')

        PAINEL_DIR.mkdir(exist_ok=True)
        for arq in arquivos:
            src_tmp = tmp / Path(arq).name
            if src_tmp.exists():
                shutil.copy2(src_tmp, Path(arq))
            run(f'git add "{arq}"')

        if index_mudou:
            Path('index.html').write_text(index_novo, encoding='utf-8')
            run('git add index.html')
            print('  index.html atualizado com novos links.')

        status = run('git diff --cached --name-only', capture=True).stdout.strip()
        if not status:
            print('Sem alteracoes detectadas -- nada a commitar.')
            return

        nomes = ', '.join(Path(f).name for f in arquivos)
        run(f'git commit -q -m "painel: {nomes}"')
        run(f'git push {REMOTE} {BRANCH_TMP}:{REMOTE_BRANCH}')
        print('Push concluido para UBANCONT.')

    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        run(f'git checkout -q {branch_orig}', check=False)
        run(f'git branch -D {BRANCH_TMP}', check=False)
        if stashed:
            run('git stash pop -q', check=False)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        publicar(sys.argv[1:])
    else:
        publicar(html_alterados())
