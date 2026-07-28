# Briefing diario de um projeto: roda claude -p na pasta do projeto com o prompt local.
# Uso: rodar_briefing.ps1 -Pasta "C:\caminho\do\projeto"
param([Parameter(Mandatory = $true)][string]$Pasta)

Set-Location $Pasta
$log = Join-Path $Pasta "briefing-log.txt"
$env:CLAUDECODE = $null; $env:CLAUDE_CODE_ENTRYPOINT = $null

# ponytail: escrita restrita ao briefing e plano do proprio projeto; leitura livre (emails.json da triagem)
& "$env:USERPROFILE\.local\bin\claude.exe" -p (Get-Content -Raw (Join-Path $Pasta "BRIEFING-PROMPT.md")) `
  --allowedTools "Read,Glob,Grep,Write(BRIEFING-DIA.md),Edit(BRIEFING-DIA.md),Write(PLANO-ACAO*.md),Edit(PLANO-ACAO*.md)" `
  2>&1 | Tee-Object -FilePath $log
