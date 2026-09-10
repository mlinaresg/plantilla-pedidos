<#
  start-servidor.ps1
  ---------------------------------------------------------------
  Arranca en segundo plano el servidor local (http-server) que
  sirve el complemento de Outlook "Plantilla reenvío pedidos".
  Pensado para lanzarse solo, al iniciar sesión, vía Programador
  de tareas (ver registrar-tarea.ps1).

  Todo lo que hace queda registrado en servidor.log, dentro de la
  misma carpeta del complemento, para poder revisar qué ha pasado
  si algún día el botón de Outlook no responde.
#>

$ErrorActionPreference = "Stop"

$Carpeta  = "C:\QGPL\Outlook\plantilla-addin"
$CertPath = Join-Path $env:USERPROFILE ".office-addin-dev-certs\localhost.crt"
$KeyPath  = Join-Path $env:USERPROFILE ".office-addin-dev-certs\localhost.key"
$Puerto   = 3000
$LogFile  = Join-Path $Carpeta "servidor.log"

function Escribir-Log {
    param([string]$Mensaje)
    $marca = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try {
        "$marca  $Mensaje" | Out-File -FilePath $LogFile -Append -Encoding UTF8
    } catch {
        # Si ni siquiera se puede escribir el log, no hay mucho más que hacer,
        # pero que al menos no rompa el script por esto.
    }
}

try {
    Escribir-Log "----- Intento de arranque -----"

    # 1) La carpeta del complemento debe existir
    if (-not (Test-Path $Carpeta)) {
        Escribir-Log "ERROR: no existe la carpeta '$Carpeta'. Revisa que la ruta sea correcta."
        exit 1
    }

    # 2) Deben estar los archivos clave del complemento
    foreach ($archivo in @("manifest.xml", "commands.html", "commands.js", "fallback.html")) {
        $ruta = Join-Path $Carpeta $archivo
        if (-not (Test-Path $ruta)) {
            Escribir-Log "ERROR: falta el archivo '$ruta'. ¿Se movió o se borró algo de la carpeta?"
            exit 1
        }
    }

    # 3) Debe existir el certificado local ya generado
    if (-not (Test-Path $CertPath) -or -not (Test-Path $KeyPath)) {
        Escribir-Log "ERROR: no se encuentra el certificado en '$CertPath' / '$KeyPath'. Ejecuta 'npx office-addin-dev-certs install' de nuevo."
        exit 1
    }

    # 4) http-server debe estar instalado y accesible
    $httpServerCmd = Get-Command "http-server" -ErrorAction SilentlyContinue
    if (-not $httpServerCmd) {
        Escribir-Log "ERROR: 'http-server' no está en el PATH. Reinstala con 'npm install -g http-server'."
        exit 1
    }

    # 4b) node.exe debe estar disponible (lo usamos directamente, en vez de
    #     pasar por el .cmd de http-server via cmd.exe, ver nota más abajo).
    $nodeCmd = Get-Command "node" -ErrorAction SilentlyContinue
    if (-not $nodeCmd) {
        Escribir-Log "ERROR: 'node' no está en el PATH."
        exit 1
    }

    # 4c) Localizamos el script real de http-server (el .cmd que "Get-Command"
    #     encuentra es solo un envoltorio de npm; el JS real vive dentro de
    #     node_modules, al lado de ese .cmd).
    $carpetaNpm = Split-Path $httpServerCmd.Source -Parent
    $httpServerJs = Join-Path $carpetaNpm "node_modules\http-server\bin\http-server"
    if (-not (Test-Path $httpServerJs)) {
        Escribir-Log "ERROR: no se encuentra el script de http-server en '$httpServerJs'. Reinstala con 'npm install -g http-server'."
        exit 1
    }

    # 5) Si el puerto ya está escuchando, asumimos que el servidor ya está
    #    arrancado (por ejemplo, de un login anterior) y no hacemos nada.
    $puertoOcupado = Get-NetTCPConnection -LocalPort $Puerto -State Listen -ErrorAction SilentlyContinue
    if ($puertoOcupado) {
        Escribir-Log "El puerto $Puerto ya está en uso; el servidor probablemente ya está arrancado. No se hace nada."
        exit 0
    }

    # 6) Arrancar el servidor, oculto, con su propia salida redirigida a log.
    #    OJO: 'http-server' en Windows es un script .cmd (no un .exe real), y
    #    Start-Process con redirección de salida exige un .exe genuino. En vez
    #    de pasar por cmd.exe (cuyas reglas de comillas con varias rutas entre
    #    comillas son frágiles al ejecutarse en segundo plano), llamamos
    #    directamente a node.exe con el script real como argumento — es
    #    exactamente lo que hace el .cmd por dentro, sin la capa intermedia.
    Set-Location $Carpeta

    Start-Process -FilePath $nodeCmd.Source `
        -ArgumentList @($httpServerJs, ".", "-S", "-C", $CertPath, "-K", $KeyPath, "-p", "$Puerto", "--cors", "-c-1") `
        -WindowStyle Hidden `
        -RedirectStandardOutput (Join-Path $Carpeta "servidor-salida.log") `
        -RedirectStandardError  (Join-Path $Carpeta "servidor-error.log")

    Escribir-Log "Servidor arrancado correctamente en https://localhost:$Puerto"
}
catch {
    Escribir-Log "ERROR INESPERADO: $($_.Exception.Message)"
    exit 1
}
