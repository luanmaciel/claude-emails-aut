# Sessao do dia em Remote Control, janela MINIMIZADA (continua acessivel no PC pela barra de tarefas).
# Idempotente: se a sessao de hoje ja esta viva, nao faz nada (roda de hora em hora como watchdog).
$projeto = "C:\Users\luan_\Claude\claude-emails-aut"
Set-Location $projeto
$pidFile = Join-Path $projeto "sessao_dia.pid"
$hoje = Get-Date -Format "yyyy-MM-dd"

if (Test-Path $pidFile) {
    $partes = (Get-Content $pidFile) -split '\|'
    $oldPid = $partes[0]; $oldData = if ($partes.Count -gt 1) { $partes[1] } else { "" }
    $vivo = Get-Process -Id $oldPid -ErrorAction SilentlyContinue
    if ($vivo -and $oldData -eq $hoje) { exit 0 }   # sessao de hoje ok, nada a fazer
    if ($vivo) { try { Stop-Process -Id $oldPid -Force -ErrorAction Stop } catch {} }
    Remove-Item $pidFile -Force
}

$data = Get-Date -Format "dd/MM/yyyy"
$nome = "$data - Resumo diario de emails e pendencias"
$antes = @(Get-Process claude -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
Start-Process -FilePath "cmd.exe" `
    -ArgumentList ('/c cd /d "' + $projeto + '" && "' + $env:USERPROFILE + '\.local\bin\claude.exe" --remote-control "' + $nome + '"') `
    -WorkingDirectory $projeto -WindowStyle Minimized | Out-Null
Start-Sleep 8
$novo = Get-Process claude -ErrorAction SilentlyContinue | Where-Object { $antes -notcontains $_.Id } |
    Sort-Object StartTime -Descending | Select-Object -First 1
if ($novo) { "$($novo.Id)|$hoje" | Set-Content $pidFile }
