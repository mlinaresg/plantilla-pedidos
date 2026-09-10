<#
  registrar-tarea.ps1
  ---------------------------------------------------------------
  Ejecuta este script UNA SOLA VEZ para dejar registrada la tarea
  que arrancará start-servidor.ps1 automáticamente cada vez que
  inicies sesión en Windows. No hace falta ser administrador.
#>

$ErrorActionPreference = "Stop"

$NombreTarea = "Plantilla pedidos - servidor local"
$RutaScript  = "C:\QGPL\Outlook\plantilla-addin\start-servidor.ps1"

try {
    if (-not (Test-Path $RutaScript)) {
        Write-Host "ERROR: no se encuentra '$RutaScript'. Revisa la ruta antes de continuar." -ForegroundColor Red
        exit 1
    }

    # Si ya existe una tarea con este nombre (por ejemplo, de una prueba
    # anterior), la quitamos primero para no duplicarla.
    $existente = Get-ScheduledTask -TaskName $NombreTarea -ErrorAction SilentlyContinue
    if ($existente) {
        Write-Host "Ya existía una tarea con este nombre; se reemplaza." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $NombreTarea -Confirm:$false
    }

    $accion = New-ScheduledTaskAction -Execute "powershell.exe" `
        -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$RutaScript`""

    $disparador = New-ScheduledTaskTrigger -AtLogOn

    $config = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    Register-ScheduledTask -TaskName $NombreTarea `
        -Action $accion `
        -Trigger $disparador `
        -Settings $config `
        -Description "Arranca el servidor local del complemento de Outlook 'Plantilla reenvío pedidos' al iniciar sesión." `
        | Out-Null

    Write-Host "Tarea '$NombreTarea' registrada correctamente." -ForegroundColor Green
    Write-Host "A partir del próximo login se arrancará sola. Para probarla ahora mismo sin reiniciar sesión, ejecuta:" -ForegroundColor Green
    Write-Host "    Start-ScheduledTask -TaskName `"$NombreTarea`"" -ForegroundColor Cyan
}
catch {
    Write-Host "ERROR al registrar la tarea: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
