# Abre uma sessao claude por email pendente (triagem\pendentes-<hoje>.json), com dossie + rascunho no Outlook.
# Idempotente: triagem\sessoes-<hoje>.json guarda entry_id -> sessao; rodar de novo so cria as que faltam.
Set-Location $PSScriptRoot
$hoje = Get-Date -Format 'yyyy-MM-dd'
$pend = "triagem\pendentes-$hoje.json"
$sess = "triagem\sessoes-$hoje.json"
$log = "triagem\log-pendentes-$hoje.txt"
if (-not (Test-Path $pend)) { "sem $pend" | Tee-Object -FilePath $log -Append; exit 0 }
New-Item -ItemType Directory -Force triagem\pendentes | Out-Null

$itens = Get-Content -Raw -Encoding UTF8 $pend | ConvertFrom-Json
$feitas = @{}
if (Test-Path $sess) {
  (Get-Content -Raw -Encoding UTF8 $sess | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $feitas[$_.Name] = $_.Value }
}
$template = Get-Content -Raw -Encoding UTF8 pendente-prompt.md
$env:CLAUDECODE = $null; $env:CLAUDE_CODE_ENTRYPOINT = $null

function Limpa($s) { ($s -replace '^(re|fw|fwd|enc|res):\s*', '').Trim() }

foreach ($it in $itens) {
  if ($feitas.ContainsKey($it.entry_id)) { continue }
  $id = [guid]::NewGuid().ToString()
  $assunto = Limpa $it.assunto
  $slug = (("$($it.remetente)-$assunto") -replace '[^\w]', '-' -replace '-+', '-').ToLower().Trim('-')
  if ($slug.Length -gt 50) { $slug = $slug.Substring(0, 50).Trim('-') }
  $empresa = if ($it.conta -match 'beta') { 'BETA' } else { 'ALFA' }
  if ($assunto.Length -gt 45) { $assunto = $assunto.Substring(0, 45) }
  $nome = "$(Get-Date -Format 'dd/MM') $empresa - $($it.remetente) - $assunto"
  $prompt = $template.Replace('{{ITEM}}', ($it | ConvertTo-Json -Depth 5)).Replace('{{DATA}}', $hoje).Replace('{{SLUG}}', $slug)

  "== $nome ($id)" | Tee-Object -FilePath $log -Append
  # ponytail: Bash so para os 2 scripts de leitura/rascunho; enviar.py fica fora, entao a sessao automatica nao consegue enviar
  & "$env:USERPROFILE\.local\bin\claude.exe" -p $prompt --session-id $id -n $nome `
    --allowedTools "Read,Glob,Grep,Edit(triagem/pendentes/**),Edit(contexto/**),Bash(python ler_email.py*),Bash(python rascunho.py*)" `
    --disallowedTools "Bash(python C:/Users/usuario/Claude/claude-emails-aut/enviar.py *),Bash(python C:\Users\usuario\Claude\claude-emails-aut\enviar.py *),Bash(python enviar.py*)" `
    2>&1 | Tee-Object -FilePath $log -Append
  if ($LASTEXITCODE -eq 0) {
    $feitas[$it.entry_id] = @{ sessao = $id; nome = $nome; slug = $slug }
    $feitas | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 $sess
  }
}

# painel.py + sync da Mesa de Emails no Notion, sob o lock do poller (um escritor so no mapa)
python mesa_poller.py --sync 2>&1 | Tee-Object -FilePath $log -Append
