# Tarea unica para hoy 9:25 AM
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $scriptDir "recordatorios_smtp.py"
$taskName = "Arainco Recordatorios - Hoy 9:25"

$pythonExe = if (Get-Command py -ErrorAction SilentlyContinue) { "py" } else { "python" }

$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false }

$trigger = New-ScheduledTaskTrigger -Once -At "9:25AM"
$action = New-ScheduledTaskAction -Execute $pythonExe -Argument "`"$scriptPath`"" -WorkingDirectory $scriptDir
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Description "Ejecucion unica hoy 9:25 AM"

Write-Host "Tarea creada: $taskName - Se ejecutara hoy a las 9:25 AM" -ForegroundColor Green
