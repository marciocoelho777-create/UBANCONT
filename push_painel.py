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


def _card_icon(href):
    """Emoji-icone por tipo de painel."""
    nome = Path(href).name
    if 'rotina_controles' in nome:  return '&#9881;'   # ⚙
    if 'diag' in nome:              return '&#128202;'  # 📊
    if 'classificacao' in nome:     return '&#128203;'  # 📋
    if 'conciliacao' in nome:       return '&#9878;'    # ⚖
    return '&#128196;'                                   # 📄

def _card_tags(href):
    """Tags descritivas para o card."""
    nome = Path(href).name
    m = re.match(r'rotina_controles_(\d{4})_(\d{2})\.html', nome)
    if m:
        return [f'{m.group(2)}/{m.group(1)}', 'Controles', 'SIGGO']
    if 'diag' in nome:
        return ['Diagnostico', '33 controles', 'Integridade']
    if 'classificacao' in nome:
        return ['Classificacao', 'Monitor']
    return []

_INDEX_CSS = """
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;font-size:14px;
     line-height:1.5;color:#0d1829;background:#eef1f8}
.hd{background:#0d1f3c;color:#fff;padding:0}
.hd-top{font-size:10.5px;opacity:.55;padding:5px 24px;
        border-bottom:1px solid rgba(255,255,255,.08);letter-spacing:.02em}
.hd-main{padding:12px 24px 14px;display:flex;align-items:center;gap:14px}
.hd-icon{width:38px;height:38px;border-radius:9px;background:rgba(255,255,255,.13);
         display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0}
.hd-title{font-size:16px;font-weight:700}
.hd-sub{font-size:12px;opacity:.6}
.wrap{max-width:1180px;margin:0 auto;padding:30px 20px}
.intro{margin-bottom:8px}
.intro h2{font-size:20px;font-weight:700;color:#0d1829;margin-bottom:4px}
.intro p{font-size:13px;color:#455268;max-width:700px;line-height:1.65;margin-bottom:18px}
.badge{display:inline-block;background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;
       border-radius:100px;font-size:11px;font-weight:600;padding:2px 12px;margin-bottom:20px}
.sec-lbl{font-size:10.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;
         color:#8a96b0;margin-bottom:14px;border-bottom:1px solid #d1d9e8;padding-bottom:8px}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:16px;margin-bottom:32px}
.card{background:#fff;border:1px solid #cdd5e8;border-radius:10px;padding:20px;
      display:flex;flex-direction:column;gap:12px;
      transition:box-shadow .18s,border-color .18s}
.card:hover{box-shadow:0 4px 18px rgba(18,85,204,.12);border-color:#a0b4dc}
.card-icon{width:40px;height:40px;border-radius:10px;background:#eff6ff;color:#1255cc;
           display:flex;align-items:center;justify-content:center;font-size:20px}
.card-title{font-size:14px;font-weight:700;color:#0d1829;line-height:1.35}
.card-desc{font-size:12.5px;color:#455268;line-height:1.6;flex:1}
.card-tags{display:flex;flex-wrap:wrap;gap:5px}
.tag{font-size:10.5px;background:#f4f6fb;color:#455268;border:1px solid #cdd5e8;
     border-radius:100px;padding:2px 9px}
.card-link{display:flex;align-items:center;justify-content:space-between;
           color:#1255cc;font-size:12.5px;font-weight:600;text-decoration:none;
           border-top:1px solid #eef1f8;padding-top:12px;margin-top:4px;
           transition:color .15s}
.card-link:hover{color:#0a3ea0}
footer{font-size:11px;color:#8a96b0;text-align:center;padding:20px}
@media(max-width:600px){.cards{grid-template-columns:1fr}}
"""

def _gerar_index_completo(links_existentes: dict) -> str:
    """
    Gera o index.html completo em estilo card.
    links_existentes: {href: texto} de todos os arquivos presentes no repo.
    """
    # Agrupa
    rotinas, outros = [], []
    for href, texto in sorted(links_existentes.items(), reverse=True):
        nome = Path(href).name
        if 'rotina_controles_' in nome:
            rotinas.append((href, texto))
        else:
            outros.append((href, texto))

    def _cards(lista):
        html = '<div class="cards">'
        for href, texto in lista:
            tags = ''.join(f'<span class="tag">{t}</span>' for t in _card_tags(href))
            html += f"""
<div class="card">
  <div class="card-icon">{_card_icon(href)}</div>
  <div class="card-title">{texto}</div>
  <div class="card-desc">Dados extraidos do SIGGO via Oracle SQL. Atualizado periodicamente.</div>
  {'<div class="card-tags">' + tags + '</div>' if tags else ''}
  <a class="card-link" href="{href}">Abrir painel <span>&#8594;</span></a>
</div>"""
        html += '</div>'
        return html

    secoes = ''
    if rotinas:
        secoes += '<div class="sec-lbl">Controles de Rotina</div>' + _cards(rotinas)
    if outros:
        secoes += '<div class="sec-lbl" style="margin-top:8px">Outros Pain&eacute;is</div>' + _cards(outros)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pain&eacute;is de Controles Cont&aacute;beis — GDF</title>
<style>{_INDEX_CSS}</style>
</head>
<body>
<div class="hd">
  <div class="hd-top">Governo do Distrito Federal &middot; Contadoria Geral do Distrito Federal &middot; SEEC/SEFIN/CONTDF</div>
  <div class="hd-main">
    <div class="hd-icon">&#127963;</div>
    <div>
      <div class="hd-title">Controles Cont&aacute;beis</div>
      <div class="hd-sub">CONTDF &middot; Contadoria Geral do Distrito Federal &middot; SIGGO</div>
    </div>
  </div>
</div>
<div class="wrap">
  <div class="intro">
    <h2>Pain&eacute;is de Controles Cont&aacute;beis <span class="badge">Exerc&iacute;cio 2026</span></h2>
    <p>Pain&eacute;is de controles cont&aacute;beis criados com vistas a subsidiar a atua&ccedil;&atilde;o dos usu&aacute;rios.
       Os dados s&atilde;o extra&iacute;dos do SIGGO com o aux&iacute;lio do Oracle SQL e passam por atualiza&ccedil;&otilde;es
       periodicamente. Cada painel permite filtrar, analisar e exportar os dados de interesse.</p>
    <p style="font-size:12.5px;color:#8a96b0">&#8595; Selecione o controle que deseja consultar</p>
  </div>
  {secoes}
</div>
<footer>CONTDF &middot; Controles Cont&aacute;beis SIGGO</footer>
</body>
</html>"""


def atualizar_index(html_remoto, novos_hrefs):
    """
    Reconstroi o index.html completo como card-grid.
    Preserva todos os links ja existentes no remoto e adiciona os novos.
    """
    # Coleta hrefs existentes (funciona com formato antigo <ul> e novo card-grid)
    hrefs = set(re.findall(r'href="([^"#][^"]*\.html)"', html_remoto))
    hrefs.discard('index.html')

    # Adiciona novos
    for href in novos_hrefs:
        hrefs.add(href)

    if not hrefs:
        return html_remoto

    existentes = {href: nome_amigavel(href) for href in hrefs}
    novo = _gerar_index_completo(existentes)
    return novo if novo != html_remoto else html_remoto


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
