# Rotina matinal no logon: triagem -> sessoes por email pendente -> briefings -> resumo (+Notion).
# Incremental: cada etapa so roda se ainda nao saiu hoje, entao logons repetidos
# nao refazem nada e falha no meio retoma de onde parou (no proximo logon ou na 2a passada).
# -Tarde (tarefa das 13h): re-extrai, triagem incremental e abre sessoes so dos emails novos.
# ponytail: 2 passadas com pausa de 5 min cobrem erro transitorio (ex.: API 529)
param([switch]$Tarde)
Set-Location $PSScriptRoot

function FeitoHoje($arquivo) {
  (Test-Path $arquivo) -and ((Get-Item $arquivo).LastWriteTime.Date -eq (Get-Date).Date)
}

if ($Tarde) {
  & "$PSScriptRoot\rodar_triagem.ps1"
  & "$PSScriptRoot\rodar_pendentes.ps1"
  exit
}

$briefings = @(
  'C:\dados\alfa\bibliotecas',
  'C:\Users\usuario\Claude\fabrica',
  'C:\dados\alfa\howwiki'
)

foreach ($passada in 1..2) {
  $md = "triagem\$(Get-Date -Format 'yyyy-MM-dd').md"
  if (-not (FeitoHoje $md)) { & "$PSScriptRoot\rodar_triagem.ps1" }

  # sessoes por email pendente (idempotente por entry_id) — logo apos a triagem
  if (FeitoHoje $md) { & "$PSScriptRoot\rodar_pendentes.ps1" }

  # briefings so depois da triagem do dia — senao leem o emails.json de ontem
  foreach ($p in $briefings) {
    if ((FeitoHoje $md) -and -not (FeitoHoje (Join-Path $p 'BRIEFING-DIA.md'))) { & "$PSScriptRoot\rodar_briefing.ps1" -Pasta $p }
  }
  $briefingsOk = -not ($briefings | Where-Object { -not (FeitoHoje (Join-Path $_ 'BRIEFING-DIA.md')) })

  # resumo so depois dos briefings; na ultima passada sai mesmo com briefing faltando
  if ((FeitoHoje $md) -and ($briefingsOk -or $passada -eq 2) -and -not (FeitoHoje 'triagem\RESUMO-DIA.md')) {
    & "$PSScriptRoot\rodar_resumo.ps1"
  }

  if ((FeitoHoje $md) -and $briefingsOk -and (FeitoHoje 'triagem\RESUMO-DIA.md')) { break }
  if ($passada -eq 1) { Start-Sleep 300 }
}
