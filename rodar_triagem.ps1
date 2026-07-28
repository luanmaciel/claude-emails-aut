# Rotina diaria: extrai emails do Outlook e roda a triagem com claude -p
Set-Location $PSScriptRoot
New-Item -ItemType Directory -Force triagem | Out-Null
$log = "triagem\log-$(Get-Date -Format 'yyyy-MM-dd').txt"

python extrair.py 3 2>&1 | Tee-Object -FilePath $log
if ($LASTEXITCODE -ne 0) { "extrair.py falhou" | Tee-Object -FilePath $log -Append; exit 1 }

$env:CLAUDECODE = $null; $env:CLAUDE_CODE_ENTRYPOINT = $null  # permite rodar dentro de uma sessao claude
# ponytail: allowedTools restrito — leitura + escrita so em triagem/ e contexto/ (corpo de email e conteudo nao confiavel)
& "$env:USERPROFILE\.local\bin\claude.exe" -p (Get-Content -Raw triagem-prompt.md) `
  --allowedTools "Read,Glob,Grep,Write(triagem/**),Write(contexto/**),Edit(triagem/**),Edit(contexto/**)" `
  2>&1 | Tee-Object -FilePath $log -Append

# abre a triagem do dia na tela quando pronta
$md = "triagem\$(Get-Date -Format 'yyyy-MM-dd').md"
if (Test-Path $md) { Start-Process $md }
