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
:root{
  --navy:#0d1b3e;--navy-mid:#162550;--navy-light:#1e3267;
  --teal:#0090a8;--teal-light:#00b8d4;--teal-pale:#e6f7fb;
  --surface:#fff;--bg:#f2f5f9;--border:#dce3ed;
  --text:#1a2033;--muted:#6b7a99;
  --radius:12px;
  --shadow-card:0 2px 14px rgba(13,27,62,.09);
  --shadow-hover:0 8px 28px rgba(13,27,62,.17);
}
body{font-family:'Segoe UI',system-ui,Arial,sans-serif;background:var(--bg);
     color:var(--text);min-height:100vh;display:flex;flex-direction:column}
.top-bar{background:var(--navy);color:#a8c0d8;font-size:11px;letter-spacing:.4px;
         padding:6px 48px;display:flex;align-items:center;gap:8px}
header{background:linear-gradient(135deg,var(--navy) 0%,var(--navy-light) 100%);
       color:#fff;padding:24px 48px 22px;display:flex;align-items:center;gap:20px;
       box-shadow:0 4px 20px rgba(13,27,62,.35)}
.header-brasao{width:60px;height:60px;flex-shrink:0;background:rgba(255,255,255,.10);
               border:2px solid rgba(255,255,255,.18);border-radius:50%;
               display:flex;align-items:center;justify-content:center;font-size:28px}
.header-text h1{font-size:18px;font-weight:700;letter-spacing:.4px;line-height:1.2}
.header-text .subtitle{font-size:12px;color:#9ab0cc;margin-top:4px;letter-spacing:.2px}
.section-intro{background:var(--surface);border-bottom:1px solid var(--border);
               padding:32px 48px 28px}
.section-intro h2{font-size:20px;font-weight:700;color:var(--navy);
                  display:flex;align-items:center;gap:10px;margin-bottom:8px}
.section-intro h2 .badge-ano{font-size:12px;font-weight:700;background:var(--teal-pale);
  color:var(--teal);border:1px solid var(--teal-light);border-radius:20px;
  padding:2px 12px;letter-spacing:.3px}
.section-intro p{font-size:13px;color:var(--muted);max-width:680px;line-height:1.65}
.instrucao{margin-top:16px;font-size:12.5px;color:var(--navy);font-weight:600;
           letter-spacing:.2px;display:flex;align-items:center;gap:6px}
.instrucao::before{content:'↓';color:var(--teal);font-size:14px}
.cards-area{padding:32px 48px 48px;flex:1}
.group-label{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;
             letter-spacing:1px;margin-bottom:16px;display:flex;align-items:center;gap:10px}
.group-label::after{content:'';flex:1;height:1px;background:var(--border)}
.cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));
            gap:14px;margin-bottom:40px}
.card{background:var(--surface);border:1.5px solid var(--border);border-radius:var(--radius);
      box-shadow:var(--shadow-card);overflow:hidden;text-decoration:none;color:var(--text);
      display:flex;flex-direction:column;transition:transform .18s,box-shadow .18s,border-color .18s;
      position:relative}
.card:hover{transform:translateY(-3px);box-shadow:var(--shadow-hover);border-color:var(--teal)}
.card::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;
             background:linear-gradient(90deg,var(--teal),var(--teal-light));
             opacity:0;transition:opacity .18s}
.card:hover::after{opacity:1}
.card-inner{padding:16px 18px 12px;display:flex;gap:14px;flex:1}
.card-icon{width:40px;height:40px;flex-shrink:0;background:var(--teal-pale);
           border-radius:10px;display:flex;align-items:center;justify-content:center;
           font-size:20px;transition:background .18s}
.card:hover .card-icon{background:#cceef5}
.card-info{flex:1;min-width:0}
.card-info h3{font-size:13px;font-weight:700;color:var(--navy);line-height:1.3;margin-bottom:4px}
.card-info p{font-size:11.5px;color:var(--muted);line-height:1.5}
.card-tags{padding:0 18px 10px;display:flex;flex-wrap:wrap;gap:4px}
.tag{font-size:10px;font-weight:600;padding:2px 7px;border-radius:20px;
     background:var(--bg);border:1px solid var(--border);color:var(--muted);letter-spacing:.2px}
.card-link{border-top:1px solid var(--border);padding:9px 18px;display:flex;
           justify-content:space-between;align-items:center;font-size:11.5px;
           font-weight:600;color:var(--teal);background:var(--bg);transition:background .15s}
.card:hover .card-link{background:var(--teal-pale)}
.card-link .arrow{font-size:15px;transition:transform .18s}
.card:hover .card-link .arrow{transform:translateX(4px)}
footer{background:var(--navy);color:#6b88a8;font-size:11px;padding:14px 48px;
       display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
@media(max-width:640px){
  .top-bar,header,.section-intro,.cards-area,footer{padding-left:18px;padding-right:18px}
  .cards-grid{grid-template-columns:1fr}
}
"""

def _gerar_index_completo(links_existentes: dict) -> str:
    """Gera o index.html completo em estilo card-grid."""
    rotinas, outros = [], []
    for href, texto in sorted(links_existentes.items(), reverse=True):
        nome = Path(href).name
        if 'rotina_controles_' in nome:
            rotinas.append((href, texto))
        else:
            outros.append((href, texto))

    def _cards(lista):
        html = '<div class="cards-grid">'
        for href, texto in lista:
            tags_html = ''.join(f'<span class="tag">{t}</span>' for t in _card_tags(href))
            html += f"""
<a class="card" href="{href}">
  <div class="card-inner">
    <div class="card-icon">{_card_icon(href)}</div>
    <div class="card-info">
      <h3>{texto}</h3>
      <p>Dados extra&iacute;dos do SIGGO via Oracle SQL. Atualizado periodicamente.</p>
    </div>
  </div>
  {'<div class="card-tags">' + tags_html + '</div>' if tags_html else ''}
  <div class="card-link">Abrir painel <span class="arrow">&#8594;</span></div>
</a>"""
        html += '</div>'
        return html

    secoes = ''
    if rotinas:
        secoes += '<div class="group-label">Controles de Rotina</div>' + _cards(rotinas)
    if outros:
        secoes += '<div class="group-label" style="margin-top:8px">Outros Pain&eacute;is</div>' + _cards(outros)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pain&eacute;is de Controles Cont&aacute;beis &mdash; GDF</title>
<style>{_INDEX_CSS}</style>
</head>
<body>
<div class="top-bar">Governo do Distrito Federal &middot; Contadoria Geral do Distrito Federal &middot; SEEC/SEFIN/CONTDF</div>
<header>
  <div class="header-brasao">&#127963;</div>
  <div class="header-text">
    <h1>Controles Cont&aacute;beis</h1>
    <div class="subtitle">CONTDF &middot; Contadoria Geral do Distrito Federal &middot; SIGGO</div>
  </div>
</header>
<div class="section-intro">
  <h2>Pain&eacute;is de Controles Cont&aacute;beis
    <span class="badge-ano">Exerc&iacute;cio 2026</span>
  </h2>
  <p>Pain&eacute;is de controles cont&aacute;beis criados com vistas a subsidiar a atua&ccedil;&atilde;o dos usu&aacute;rios.
     Os dados s&atilde;o extra&iacute;dos do SIGGO com o aux&iacute;lio do Oracle SQL e passam por atualiza&ccedil;&otilde;es
     periodicamente. Cada painel permite filtrar, analisar e exportar os dados de interesse.</p>
  <div class="instrucao">Selecione o controle que deseja consultar</div>
</div>
<div class="cards-area">
  {secoes}
</div>
<footer>
  <span>CONTDF &middot; Controles Cont&aacute;beis SIGGO</span>
</footer>
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
