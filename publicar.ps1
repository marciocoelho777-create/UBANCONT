# publicar.ps1 — execução diária do diagnóstico e publicação do painel.
#
# Agendar no Windows (Tarefa Agendada, diária às 06:00):
#   Programa:   powershell.exe
#   Argumentos: -ExecutionPolicy Bypass -File "C:\...\publicar.ps1"
#   Iniciar em: C:\Users\marcio.coelho\Desktop\balanço 2026 gemini arquivos
#
# O diag.py roda AQUI, na máquina do GDF, porque o Oracle não é acessível de
# fora da rede. O GitHub recebe só o resultado — nenhum runner na nuvem
# consegue (nem deve) alcançar o banco.

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

# Competência: mês anterior até o dia 10, mês corrente depois disso.
# Ajuste conforme o calendário de fechamento contábil de vocês.
$hoje = Get-Date
if ($hoje.Day -le 10) { $ref = $hoje.AddMonths(-1) } else { $ref = $hoje }
$mes = $ref.Month
$ano = $ref.Year

& ".\.venv\Scripts\Activate.ps1"

Write-Host "== Diagnóstico $mes/$ano =="
python diag.py --mes $mes --ano $ano --json --excel
if ($LASTEXITCODE -ne 0) {
    # Código de saída != 0 pode significar ERRO contábil encontrado, o que é
    # um resultado legítimo e deve ser publicado. Só aborta se não houve JSON.
    Write-Host "diag.py retornou $LASTEXITCODE — verificando se gerou dados..."
}

$novos = Get-ChildItem -Path "painel\dados" -Filter "*.json" -ErrorAction SilentlyContinue |
         Where-Object { $_.LastWriteTime -gt $hoje.Date }
if (-not $novos) {
    Write-Error "Nenhum JSON gerado hoje. Painel não atualizado."
    exit 1
}

python painel.py

git add painel/dados README.md
git commit -m "diagnóstico $($ref.ToString('MM/yyyy')) — $($hoje.ToString('dd/MM HH:mm'))"
git push

Write-Host "Painel publicado."
