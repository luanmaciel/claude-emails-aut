# Sessao do dia em Remote Control (janela minimizada, com titulo NAO FECHAR).
# Garantias:
#  - trava anti-corrida (startup + watchdog horario nao criam duplicatas)
#  - detecta a sessao viva pelo processo real (nao por PID salvo, que se perde no reboot)
#  - reabertura no mesmo dia RETOMA a conversa (claude remote-control --continue)
$projeto = "C:\Users\luan_\Claude\claude-emails-aut"
Set-Location $projeto
if ((Get-Date).Hour -lt 8) { exit 0 }
$hoje = Get-Date -Format "yyyy-MM-dd"
$dataFile = Join-Path $projeto "sessao_dia.data"

# trava: se outra instancia deste script estiver rodando, sai
$lockPath = Join-Path $projeto "sessao_dia.lock"
try { $lock = [IO.File]::Open($lockPath, "CreateNew", "Write", "None") } catch { exit 0 }
try {
    $rc = Get-CimInstance Win32_Process -Filter "Name='claude.exe'" |
        Where-Object { $_.CommandLine -match "remote-control" }
    $dataCriada = if (Test-Path $dataFile) { (Get-Content $dataFile).Trim() } else { "" }

    if ($rc -and $dataCriada -eq $hoje) { exit 0 }   # sessao de hoje viva: nada a fazer
    if ($rc) { $rc | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } }

    $claude = "$env:USERPROFILE\.local\bin\claude.exe"
    if ($dataCriada -eq $hoje) {
        # ja houve sessao hoje -> retomar a mesma conversa
        $cmdClaude = '"' + $claude + '" remote-control --continue'
    } else {
        $data = Get-Date -Format "dd/MM/yyyy"
        $cmdClaude = '"' + $claude + '" --allowedTools "Read,Glob,Grep" --remote-control "' + $data + ' - Resumo diario de emails e pendencias"'
    }

    Start-Process -FilePath "cmd.exe" `
        -ArgumentList ('/c title SESSAO DO DIA (Claude) - NAO FECHAR ^| minimizar apenas && cd /d "' + $projeto + '" && ' + $cmdClaude) `
        -WorkingDirectory $projeto -WindowStyle Minimized | Out-Null
    Set-Content $dataFile $hoje
} finally {
    $lock.Close()
    Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
}
