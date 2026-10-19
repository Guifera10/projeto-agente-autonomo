#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$pasta = Split-Path -Parent $MyInvocation.MyCommand.Path
$taskName = "DOE Agente Autonomo"
$cmd = Join-Path $pasta "rodar_ciclo.cmd"
$python = "C:\Users\guilherme-medina\AppData\Local\Programs\Python\Python313\python.exe"
$orquestrador = Join-Path $pasta "orquestrador_doe.py"

if (-not (Test-Path -LiteralPath $python)) {
    throw "Python nao encontrado: $python"
}
if (-not (Test-Path -LiteralPath $orquestrador)) {
    throw "Orquestrador nao encontrado: $orquestrador"
}
if (-not (Test-Path -LiteralPath $cmd)) {
    throw "Wrapper nao encontrado: $cmd"
}

New-Item -ItemType Directory -Force -Path (Join-Path $pasta "resultados") | Out-Null

$action = New-ScheduledTaskAction `
    -Execute "C:\Windows\System32\cmd.exe" `
    -Argument "/c `"$cmd`"" `
    -WorkingDirectory $pasta
$inv = [cultureinfo]::InvariantCulture
$inicioManha = [datetime]::ParseExact("2026-09-11 07:00:00", "yyyy-MM-dd HH:mm:ss", $inv)
$inicioNoite = [datetime]::ParseExact("2026-09-11 19:00:00", "yyyy-MM-dd HH:mm:ss", $inv)
$triggerManha = New-ScheduledTaskTrigger -Daily -At "07:00"
$triggerNoite = New-ScheduledTaskTrigger -Daily -At "19:00"
$triggerManha.StartBoundary = $inicioManha.ToString("s")
$triggerNoite.StartBoundary = $inicioNoite.ToString("s")
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 16) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 5)
$principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType Interactive `
    -RunLevel Limited

Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger @($triggerManha, $triggerNoite) `
    -Settings $settings `
    -Principal $principal `
    -Description "Se o DOE do dia ainda nao saiu, consulta a cada 10 min ate gravar resultados\AAAA-MM-DD. Com o resultado pronto, encerra. Extras na sessao das 19:00. Disparo 07:00 e 19:00. Limite 16h. Log em resultados\ultimo_ciclo.log." `
    -Force | Out-Null

$info = Get-ScheduledTask -TaskName $taskName
$state = Get-ScheduledTaskInfo -TaskName $taskName
$act = $info.Actions[0]
Write-Host "Tarefa criada: $($info.TaskName)"
Write-Host "Estado: $($info.State)"
Write-Host "Execute: $($act.Execute)"
Write-Host "Arguments: $($act.Arguments)"
Write-Host "WorkingDirectory: $($act.WorkingDirectory)"
Write-Host "Proxima execucao: $($state.NextRunTime)"
Write-Host "Horarios: todos os dias 07:00 e 19:00"
