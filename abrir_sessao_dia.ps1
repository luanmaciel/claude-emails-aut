# Abre a sessao do dia em Remote Control (janela oculta, acessivel pelo celular).
# Mata a sessao do dia anterior antes, pra nao acumular processos.
$projeto = "C:\Users\luan_\Claude\claude-emails-aut"
Set-Location $projeto
$pidFile = Join-Path $projeto "sessao_dia.pid"

if (Test-Path $pidFile) {
    $oldPid = Get-Content $pidFile
    try { Stop-Process -Id $oldPid -Force -ErrorAction Stop } catch {}
    Remove-Item $pidFile -Force
}

$data = Get-Date -Format "dd/MM/yyyy"
$nome = "$data - Resumo diario de emails e pendencias"
# cmd /c garante que o claude nasce com o cwd do projeto (o projeto da sessao vem do cwd)
$antes = @(Get-Process claude -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$p = Start-Process -FilePath "cmd.exe" `
    -ArgumentList ('/c cd /d "' + $projeto + '" && "' + $env:USERPROFILE + '\.local\bin\claude.exe" --remote-control "' + $nome + '"') `
    -WorkingDirectory $projeto -WindowStyle Hidden -PassThru
Start-Sleep 8
$novo = Get-Process claude -ErrorAction SilentlyContinue | Where-Object { $antes -notcontains $_.Id } |
    Sort-Object StartTime -Descending | Select-Object -First 1
if ($novo) { $novo.Id | Set-Content $pidFile } else { $p.Id | Set-Content $pidFile }
