# run_noturno.ps1 — Execução automática noturna do diagnóstico contábil GDF
# Criado em 21/08/2026. Agendado via Agendador de Tarefas às 02:00 diariamente.

$ErrorActionPreference = "Continue"

$PROJ = "C:\Users\marcio.coelho\Desktop\balanço 2026 gemini arquivos"
$PYTHON = "$PROJ\.venv\Scripts\python.exe"
$LOG = "$PROJ\log_noturno.txt"

Set-Location $PROJ

$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content $LOG "`n===== $ts ====="

# Mês e ano correntes
$mes = (Get-Date).Month
$ano = (Get-Date).Year

Add-Content $LOG "  Rodando mes=$mes ano=$ano"

# 1. Diagnóstico de integridade (gera JSON + painel_diag.html)
& $PYTHON diag.py --mes $mes --ano $ano --json 2>&1 | Tee-Object -Append -FilePath $LOG

# 2. Balancete Contábil (gera xlsx lido pela auditoria consolidada)
& $PYTHON gerar_balancete.py --mes $mes --ano $ano 2>&1 | Tee-Object -Append -FilePath $LOG

# 3. Monitor de classificação (gera painel_classificacao.html)
& $PYTHON monitor_classificacao.py --mes $mes --ano $ano 2>&1 | Tee-Object -Append -FilePath $LOG

# 4. Auditoria consolidada (gera auditoria_consolidada.html)
& $PYTHON auditoria_consolidada.py --mes $mes --ano $ano 2>&1 | Tee-Object -Append -FilePath $LOG

# 5. Controles de rotina (gera Excel + rotina_controles_{ano}_{mes:02d}.html)
& $PYTHON rotina\gerar_controles_rotina.py --mes $mes --ano $ano 2>&1 | Tee-Object -Append -FilePath $LOG

# 6. Commit no git (este repositório)
$label = "{0:d2}/{1}" -f $mes, $ano
$msg = "run noturno $(Get-Date -Format 'yyyy-MM-dd') — $label"
$mesPad = "{0:d2}" -f $mes
git add painel/dados/ painel/painel_diag.html painel/painel_classificacao.html painel/auditoria_consolidada.html "painel/rotina_controles_${ano}_${mesPad}.html" 2>&1 | Tee-Object -Append -FilePath $LOG
git commit -m $msg 2>&1 | Tee-Object -Append -FilePath $LOG

# 7. Publicar painéis no UBANCONT (GitHub Pages)
$UBANCONT = "C:\Users\marcio.coelho\UBANCONT"
if (Test-Path $UBANCONT) {
    $painelSrc = "$PROJ\painel"
    $painelDst = "$UBANCONT\painel"
    Copy-Item "$painelSrc\painel_diag.html"            $painelDst -Force
    Copy-Item "$painelSrc\painel_classificacao.html"   $painelDst -Force
    Copy-Item "$painelSrc\auditoria_consolidada.html"  $painelDst -Force
    Copy-Item "$painelSrc\painel_balancete.html"       $painelDst -Force
    Copy-Item "$painelSrc\rotina_controles_${ano}_${mesPad}.html" $painelDst -Force
    Set-Location $UBANCONT
    git add painel/ 2>&1 | Tee-Object -Append -FilePath $LOG
    git commit -m $msg 2>&1 | Tee-Object -Append -FilePath $LOG
    git push origin main 2>&1 | Tee-Object -Append -FilePath $LOG
    Set-Location $PROJ
}

Add-Content $LOG "  Concluido em $(Get-Date -Format 'HH:mm:ss')"
