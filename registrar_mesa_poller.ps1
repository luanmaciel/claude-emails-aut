# Registra (ou atualiza) a tarefa MesaPoller no Task Scheduler: pythonw mesa_poller.py a cada 2 min,
# no logon e a partir de agora, so com o usuario logado (COM do Outlook), sem nova instancia se a
# anterior ainda roda, limite de 35 min por run (um "faz" grande pode levar 25 min).
# Rodar uma vez (e de novo se mudar algo aqui): powershell -File registrar_mesa_poller.ps1
$py = "$env:LOCALAPPDATA\Programs\Python\Python313\pythonw.exe"
$repo = $PSScriptRoot
$user = "$env:USERDOMAIN\$env:USERNAME"
$action = New-ScheduledTaskAction -Execute $py -Argument "mesa_poller.py" -WorkingDirectory $repo
$rep = (New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 2)).Repetition
$logon = New-ScheduledTaskTrigger -AtLogOn -User $user
$logon.Repetition = $rep
$agora = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 2)
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 35) `
  -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName MesaPoller -Action $action -Trigger $logon, $agora -Settings $settings -Principal $principal -Force | Out-Null
$t = Get-ScheduledTask MesaPoller
"MesaPoller: $($t.State) | $($t.Actions[0].Execute) $($t.Actions[0].Arguments) | cwd $($t.Actions[0].WorkingDirectory)"
"triggers: " + (($t.Triggers | ForEach-Object { $_.CimClass.CimClassName + " a cada " + $_.Repetition.Interval }) -join " | ")
"proximo run: " + (Get-ScheduledTaskInfo MesaPoller).NextRunTime
