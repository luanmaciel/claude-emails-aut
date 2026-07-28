# Sessao do dia em Remote Control, janela minimizada.
# - Primeira abertura do dia: sessao nova com nome datado.
# - Reaberturas no mesmo dia (queda, fechamento acidental, PC religado): RETOMA a mesma conversa (--continue).
# - Idempotente: se a sessao de hoje esta viva, nao faz nada. Roda no agendamento horario e no logon.
$projeto = "C:\Users\luan_\Claude\claude-emails-aut"
Set-Location $projeto
$pidFile = Join-Path $projeto "sessao_dia.pid"
$hoje = Get-Date -Format "yyyy-MM-dd"

if ((Get-Date).Hour -lt 8) { exit 0 }  # antes das 08h o resumo do dia ainda nao existe

$retomar = $false
if (Test-Path $pidFile) {
    $partes = (Get-Content $pidFile) -split '\|'
    $oldPid = $partes[0]; $oldData = if ($partes.Count -gt 1) { $partes[1] } else { "" }
    $vivo = Get-Process -Id $oldPid -ErrorAction SilentlyContinue
    if ($vivo -and $oldData -eq $hoje) { exit 0 }   # sessao de hoje ok
    if ($vivo) { try { Stop-Process -Id $oldPid -Force -ErrorAction Stop } catch {} }
    if ($oldData -eq $hoje) { $retomar = $true }     # ja houve sessao hoje -> retomar conversa
    Remove-Item $pidFile -Force
}

$data = Get-Date -Format "dd/MM/yyyy"
$nome = "$data - Resumo diario de emails e pendencias"
$claude = "$env:USERPROFILE\.local\bin\claude.exe"
if ($retomar) {
    $cmdClaude = '"' + $claude + '" remote-control --continue'
} else {
    $cmdClaude = '"' + $claude + '" --allowedTools "Read,Glob,Grep" --remote-control "' + $nome + '"'
}

$antes = @(Get-Process claude -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
Start-Process -FilePath "cmd.exe" `
    -ArgumentList ('/c title SESSAO DO DIA (Claude) - NAO FECHAR ^| minimizar apenas && cd /d "' + $projeto + '" && ' + $cmdClaude) `
    -WorkingDirectory $projeto -WindowStyle Minimized | Out-Null
Start-Sleep 8
$novo = Get-Process claude -ErrorAction SilentlyContinue | Where-Object { $antes -notcontains $_.Id } |
    Sort-Object StartTime -Descending | Select-Object -First 1
if ($novo) { "$($novo.Id)|$hoje" | Set-Content $pidFile }
