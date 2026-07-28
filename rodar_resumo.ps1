# Consolida os briefings do dia em triagem\RESUMO-DIA.md e abre na tela
Set-Location $PSScriptRoot
$env:CLAUDECODE = $null; $env:CLAUDE_CODE_ENTRYPOINT = $null
& "$env:USERPROFILE\.local\bin\claude.exe" -p (Get-Content -Raw resumo-prompt.md) `
  --allowedTools "Read,Glob,Grep,Write(triagem/**),Edit(triagem/**)" `
  2>&1 | Tee-Object -FilePath "triagem\log-resumo.txt"
if (Test-Path "triagem\RESUMO-DIA.md") { Start-Process "triagem\RESUMO-DIA.md" }

# backup diario dos dados sensiveis (fora do git) + limpeza de zips com mais de 30 dias
$bkp = "C:\Users\luan_\Claude\backups"
New-Item -ItemType Directory -Force $bkp | Out-Null
Compress-Archive -Path contexto, triagem -DestinationPath "$bkp\emails-aut-$(Get-Date -Format 'yyyy-MM-dd').zip" -Force
Get-ChildItem "$bkp\emails-aut-*.zip" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force
