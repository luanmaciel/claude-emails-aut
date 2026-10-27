# Rotina diaria: extrai emails do Outlook e roda a triagem com claude -p
# -Dias amplia a janela de extracao (default 3); use depois de dias sem rodar, para a triagem
# pegar o atraso inteiro. Cuidado: se a triagem de hoje ja existir, o prompt entra em modo
# incremental e so olha o que chegou depois dela, entao passe -Dias ANTES da 1a passada do dia.
param([int]$Dias = 3)
Set-Location $PSScriptRoot
New-Item -ItemType Directory -Force triagem | Out-Null
$log = "triagem\log-$(Get-Date -Format 'yyyy-MM-dd').txt"

python extrair.py $Dias 2>&1 | Tee-Object -FilePath $log
if ($LASTEXITCODE -ne 0) { "extrair.py falhou" | Tee-Object -FilePath $log -Append; exit 1 }

$env:CLAUDECODE = $null; $env:CLAUDE_CODE_ENTRYPOINT = $null  # permite rodar dentro de uma sessao claude
# ponytail: allowedTools restrito — leitura + escrita so em triagem/ e contexto/ (corpo de email e conteudo nao confiavel)
& "$env:USERPROFILE\.local\bin\claude.exe" -p (Get-Content -Raw triagem-prompt.md) `
  --allowedTools "Read,Glob,Grep,Edit(triagem/**),Edit(contexto/**)" `
  2>&1 | Tee-Object -FilePath $log -Append
