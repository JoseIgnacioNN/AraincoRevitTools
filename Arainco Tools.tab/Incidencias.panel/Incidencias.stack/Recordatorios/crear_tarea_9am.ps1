# Crea la tarea programada para recordatorios de incidencias
# Ejecutar como Administrador: powershell -ExecutionPolicy Bypass -File crear_tarea_9am.ps1

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $scriptDir "recordatorios_smtp.py"
$taskName = "Arainco Recordatorios Incidencias"

# Buscar Python
$pythonExe = $null
if (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonExe = "py"
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonExe = "python"
} else {
    Write-Host "ERROR: No se encontro Python (py o python)" -ForegroundColor Red
    exit 1
}

$arguments = "`"$scriptPath`""
$workingDir = $scriptDir

# Eliminar tarea existente si existe
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Tarea anterior eliminada." -ForegroundColor Yellow
}

# Crear trigger: Lunes a Viernes a las 9:00 AM
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At "09:00"
$action = New-ScheduledTaskAction -Execute $pythonExe -Argument $arguments -WorkingDirectory $workingDir
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description "Recordatorios de incidencias BIM - Lunes a Viernes 9:00 AM (respeta cadencia por prioridad)"

Write-Host "Tarea creada: $taskName" -ForegroundColor Green
Write-Host "  Horario: Lunes a Viernes, 9:00 AM" -ForegroundColor Gray
Write-Host "  Script: $scriptPath" -ForegroundColor Gray
Write-Host ""
Write-Host "Para verificar: taskschd.msc -> Buscar '$taskName'" -ForegroundColor Cyan
