<#
.SYNOPSIS
  Registra la tarea programada "Tablero Industrial - Actualizar" en Task Scheduler.

.DESCRIPTION
  Crea/actualiza una tarea de Windows que corre el pipeline de actualizacion
  del Tablero Industrial. La tarea:
    - Corre todos los dias a las 06:00 AM
    - Ademas corre al iniciar sesion (por si el PC estaba apagado)
    - Guarda logs en automatizacion/logs/
    - Notifica al usuario en caso de falla (toast Windows + ULTIMO_ERROR.txt)
    - Corre como el usuario actual (NO requiere admin, NO corre cuando estas
      deslogueado salvo que configures lo contrario)

.NOTES
  Corre este script UNA SOLA VEZ para registrar la tarea.
  Requiere permisos de usuario normales (NO admin).

.USAGE
  Desde PowerShell:
    cd "C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos\HIVIMAR\Proyectos\SGI 2.0\Indicadores Tablero\automatizacion"
    powershell -ExecutionPolicy Bypass -File configurar_scheduler.ps1

  Opcional - cambiar la hora:
    powershell -ExecutionPolicy Bypass -File configurar_scheduler.ps1 -HoraDiaria "06:30"

  Para desregistrar:
    powershell -ExecutionPolicy Bypass -File configurar_scheduler.ps1 -Remove
#>

param(
    [string]$HoraDiaria = "06:00",
    [switch]$Remove,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

# Nombre de la tarea y rutas
$TaskName = "Tablero Industrial - Actualizar"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir
$PipelineScript = Join-Path $ScriptDir "update_tablero.py"

# Python del user (sin admin)
$UserName = $env:USERNAME
$PythonExe = "C:\Users\$UserName\AppData\Local\Programs\Python\Python312\python.exe"
if (-not (Test-Path $PythonExe)) {
    Write-Host "ERROR: no se encontro Python en $PythonExe" -ForegroundColor Red
    Write-Host "Edita la variable `$PythonExe en este script para apuntar al Python correcto."
    exit 1
}

if (-not (Test-Path $PipelineScript)) {
    Write-Host "ERROR: no se encontro $PipelineScript" -ForegroundColor Red
    exit 1
}

# ============================================
# OPCION DE REMOCION
# ============================================
if ($Remove) {
    Write-Host "Removiendo tarea '$TaskName'..."
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-Host "Tarea removida exitosamente." -ForegroundColor Green
    } catch {
        Write-Host "La tarea no existe o no se pudo remover: $_" -ForegroundColor Yellow
    }
    exit 0
}

# ============================================
# CREAR/ACTUALIZAR TAREA
# ============================================

Write-Host ""
Write-Host "=== Tablero Industrial - Registrar tarea programada ===" -ForegroundColor Cyan
Write-Host "Task name     : $TaskName"
Write-Host "Python        : $PythonExe"
Write-Host "Pipeline      : $PipelineScript"
Write-Host "Project dir   : $ProjectDir"
Write-Host "Hora diaria   : $HoraDiaria"
Write-Host "Usuario       : $UserName"
Write-Host ""

if ($DryRun) {
    Write-Host "DRY RUN: no se aplicaran cambios." -ForegroundColor Yellow
    exit 0
}

# Argumentos: -u para unbuffered (importante para logs en vivo)
# --backup para que respalde el Excel si algo escribe (hoy por defecto no lo hace)
$Arguments = "-u `"$PipelineScript`" --backup"

$Action = New-ScheduledTaskAction `
    -Execute $PythonExe `
    -Argument $Arguments `
    -WorkingDirectory $ProjectDir

# Trigger 1: diario a la hora indicada
$TriggerDiario = New-ScheduledTaskTrigger -Daily -At $HoraDiaria

# Trigger 2: al iniciar sesion (compensa si la PC estaba apagada a las 6 AM)
$TriggerLogon = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$UserName"
# Retraso de 2 min para que Windows termine de cargar
$TriggerLogon.Delay = 'PT2M'

# Settings: solo con AC (evita correr con bateria baja), permitir arrancar
# tarde si la PC estaba apagada, reiniciar si falla.
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 15) `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

# Correr como el usuario actual, solo cuando el user esta logueado (mas seguro)
$Principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$UserName" `
    -LogonType Interactive `
    -RunLevel Limited

# Descripcion
$Desc = "Actualiza el tablero Hivimar Industrial 8.0: descarga Qlik + Odoo, " +
        "enriquece datos, genera db_output.js y actualiza el HTML. " +
        "Logs en automatizacion/logs/. En caso de falla crea ULTIMO_ERROR.txt."

$Task = New-ScheduledTask `
    -Action $Action `
    -Trigger @($TriggerDiario, $TriggerLogon) `
    -Settings $Settings `
    -Principal $Principal `
    -Description $Desc

# Remover primero si ya existe (para re-registrar limpiamente)
try {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
} catch {}

# Registrar
Register-ScheduledTask -TaskName $TaskName -InputObject $Task | Out-Null
Write-Host "Tarea registrada exitosamente." -ForegroundColor Green

# Info final
Write-Host ""
Write-Host "Proximos triggers:" -ForegroundColor Cyan
$t = Get-ScheduledTask -TaskName $TaskName
$info = Get-ScheduledTaskInfo -TaskName $TaskName
Write-Host "  Ultima ejecucion : $($info.LastRunTime)"
Write-Host "  Proxima ejecucion: $($info.NextRunTime)"
Write-Host "  Estado           : $($t.State)"
Write-Host ""
Write-Host "Para probar manualmente sin esperar:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host ""
Write-Host "Para ver historico:" -ForegroundColor Yellow
Write-Host "  Administrador de Tareas (taskschd.msc) -> Libreria -> buscar '$TaskName'"
Write-Host ""
Write-Host "Los logs quedaran en: $ScriptDir\logs\" -ForegroundColor Yellow
Write-Host ""
