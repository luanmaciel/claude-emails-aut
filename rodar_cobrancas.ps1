# Levanta as threads sem resposta e escreve triagem\COBRANCAS-PARADAS.md
# Uso: .\rodar_cobrancas.ps1 [janela_dias] [dias_minimos]   (default: 60 e 4)
param([int]$Janela = 60, [int]$MinDias = 4)
Set-Location $PSScriptRoot
New-Item -ItemType Directory -Force triagem | Out-Null
$log = "triagem\log-cobrancas-$(Get-Date -Format 'yyyy-MM-dd').txt"

python cobrancas.py $Janela $MinDias 2>&1 | Tee-Object -FilePath $log
if ($LASTEXITCODE -ne 0) { "cobrancas.py falhou" | Tee-Object -FilePath $log -Append; exit 1 }

$env:CLAUDECODE = $null; $env:CLAUDE_CODE_ENTRYPOINT = $null  # permite rodar dentro de uma sessao claude
# allowedTools restrito: leitura + escrita so em triagem/ (corpo de email e conteudo nao confiavel)
& "$env:USERPROFILE\.local\bin\claude.exe" -p (Get-Content -Raw cobrancas-prompt.md) `
  --allowedTools "Read,Glob,Grep,Write(triagem/**),Edit(triagem/**)" `
  2>&1 | Tee-Object -FilePath $log -Append

$md = "triagem\COBRANCAS-PARADAS.md"
if (Test-Path $md) { Start-Process $md }
