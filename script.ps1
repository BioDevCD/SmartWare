$version = "1.3.4"

# ════════════════════════════════════════════════════════════════════════════════════
# PROTOCOLO DE DIAGNÓSTICO INTELIGENTE SMARTWARE
# Potencia tu Mundo Digital
# ════════════════════════════════════════════════════════════════════════════════════

# --- VALIDACIÓN DE PRIVILEGIOS CRÍTICOS ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    exit
}

function Write-Log {
    param($Mensaje, $Gravedad = "INFO")

    # 1. DETECCIÓN DE RUTA BASE (Script o EXE)
    $rutaBase = try {
        if ($PSScriptRoot) { $PSScriptRoot }
        else { Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) }
    } catch { $env:TEMP }

    $archivoLog = "SmartWare_Debug.log"
    $linea = "FECHA: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss') | GRAVEDAD: $Gravedad | MENSAJE: $Mensaje"

    # 2. INTENTO DE ESCRITURA CON CASCADA DE SEGURIDAD
    # Definimos las rutas en orden de prioridad
    $rutasPosibles = @(
        (Join-Path $rutaBase $archivoLog),                              # Al lado del programa
        (Join-Path ([Environment]::GetFolderPath("Desktop")) $archivoLog), # Escritorio (puede ser OneDrive)
        (Join-Path $env:USERPROFILE "Desktop\$archivoLog"),            # Escritorio Local forzado
        (Join-Path $env:TEMP $archivoLog)                              # Temp (Último recurso)
    )

    $logEscrito = $false
    foreach ($ruta in $rutasPosibles) {
        if ($logEscrito) { break }
        try {
            # Verificamos si el directorio existe (por si OneDrive movió la ruta)
            $dir = Split-Path -Parent $ruta
            if (-not (Test-Path $dir)) { continue }

            $linea | Out-File -FilePath $ruta -Append -Encoding UTF8 -ErrorAction Stop
            $logEscrito = $true
        } catch {
            continue # Si falla esta ruta, probamos la siguiente
        }
    }
}

[console]::CursorVisible = $false

try {
    $host.UI.RawUI.WindowTitle = "Diagnóstico Inteligente SmartWare"
    $Host.UI.RawUI.BackgroundColor = "Black"
    $Host.UI.RawUI.ForegroundColor = "White"
} catch {
    # Fallback silencioso
}

try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
} catch {
    Write-Log -Mensaje "Error cargando ensamblados visuales: $($_.Exception.Message)" -Gravedad "CRITICA"
    # Aquí podrías decidir si salir o seguir en modo solo texto
}

$ErrorActionPreference = "SilentlyContinue"
Clear-Host

# --- CONFIGURACIÓN DE ENTORNO VISUAL ---
$H = $Host.UI.RawUI
$WinSize = $H.WindowSize
$BufSize = $H.BufferSize

# Definimos un ancho estándar de 100 caracteres y alto de 50
$WinSize.Width = 100
$WinSize.Height = 50
$BufSize.Width = 100
$BufSize.Height = 3000 # El buffer alto permite hacer scroll hacia arriba

# Aplicamos los cambios (con manejo de error por si la pantalla es muy pequeña)
try {
    $H.BufferSize = $BufSize
    $H.WindowSize = $WinSize
} catch {
    # Si falla (ej. resolución muy baja), ignoramos para no detener el script
}

# Limpiamos para empezar de cero con el nuevo tamaño
Clear-Host

function Out-SmartWare {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][Parameter(Position=0)][string]$Texto = "",
        [Parameter(Position=1)][string]$Color = "Gray",
        [switch]$NoNewLine
    )
    # Si mandas algo vacío, es un salto de línea real para sincronizar la consola
    if ($Texto.Length -eq 0 -and -not $NoNewLine) {
        Write-Host ""
    } else {
        Write-Host -Object $Texto -ForegroundColor $Color -NoNewline:$NoNewLine
    }
}

# --- [CONFIGURACIÓN DE ENTORNO SEGURO] ---
$code = @"
using System;
using System.Runtime.InteropServices;

public class ConsoleHelper {
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    const int STD_INPUT_HANDLE = -10;
    const uint ENABLE_QUICK_EDIT_MODE = 0x0040;
    const uint ENABLE_EXTENDED_FLAGS = 0x0080;

    public static void DisableQuickEdit() {
        IntPtr consoleHandle = GetStdHandle(STD_INPUT_HANDLE);
        uint consoleMode;
        if (GetConsoleMode(consoleHandle, out consoleMode)) {
            consoleMode &= ~ENABLE_QUICK_EDIT_MODE;
            consoleMode |= ENABLE_EXTENDED_FLAGS;
            consoleMode |= 0x0001; // ENABLE_PROCESSED_INPUT
            SetConsoleMode(consoleHandle, consoleMode);
        }
    }
}
"@
try {
    Add-Type -TypeDefinition $code -ErrorAction Stop
    [ConsoleHelper]::DisableQuickEdit()
} catch {
    # Si falla, simplemente ignoramos. El script funcionará igual,
    # solo que el usuario podría pausar la consola sin querer al hacer clic.
    Write-Log -Mensaje "No se pudo desactivar QuickEdit (C#): $($_.Exception.Message)" -Gravedad "AVISO"
}

# Bloquear entrada de teclado "fantasma" durante procesos
[console]::InputEncoding = [System.Text.Encoding]::UTF8

function Write-Typewriter($texto, $color = "Gray", $ms = 15, [switch]$NoNewLine) {
    if ([string]::IsNullOrEmpty($texto)) { return }

    $letras = $texto.ToCharArray()
    $limiteEscritura = $letras.Count
    $tienePuntos = $texto.EndsWith("...")

    if ($tienePuntos) { $limiteEscritura -= 3 }

    for ($i = 0; $i -lt $limiteEscritura; $i++) {
        Out-SmartWare -Texto $letras[$i] -Color $color -NoNewLine
        Start-Sleep -Milliseconds $ms
    }

    if ($tienePuntos) {
        $posPuntos = $host.UI.RawUI.CursorPosition
        $puntosAnim = @("    ", ".   ", "..  ", "...")

        for ($j = 0; $j -lt 8; $j++) {
            $host.UI.RawUI.CursorPosition = $posPuntos
            Out-SmartWare -Texto $puntosAnim[$j % 4] -Color $color -NoNewLine
            Start-Sleep -Milliseconds 400
         }
    }

    if (-not $NoNewLine) { Out-SmartWare "" }
}

function Show-Loader($segundos) {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $puntos = @("    ", ".   ", "..  ", "...")
    $i = 0
    Write-Typewriter -NoNewline -Texto " Cargando" -Color "Cyan"
    $pos = $host.UI.RawUI.CursorPosition
    while ($stopwatch.Elapsed.TotalSeconds -lt $segundos) {
        $host.UI.RawUI.CursorPosition = $pos
        Out-SmartWare $puntos[$i % 4] -NoNewline -Color Cyan
        Start-Sleep -Milliseconds 400
        $i++
    }
    $host.UI.RawUI.CursorPosition = $pos
    Out-SmartWare -Texto "    `r" -NoNewline
}

# Aquí comienza el try global del script
try {
$scriptExitoso = $false

Write-Log -Mensaje "SmartWare v$version iniciado correctamente."

#region [ DATOS ADJUNTOS ]
$base64Zip = @"
# [aquí va el bloque base64 del tools.zip que contiene HWMonitor, CrystalDiskInfo, TreeSize Free y AdwCleaner]
"@
#endregion

# [1. ESCUDO DE PROTECCIÓN DE PREVUELO]

Clear-Host

# El primer impacto visual
Show-Loader 2.5

Write-Typewriter -Texto " ════════════════════════════════════════════════════════════════════════════════════" -Color "Cyan"
Write-Typewriter -Texto " SMARTWARE - Potencia tu Mundo Digital" -Color "Cyan"
Write-Typewriter -Texto " DIAGNÓSTICO INTELIGENTE v$version" -Color "Cyan"
Write-Typewriter -Texto " PANEL DE CONTROL" -Color "Cyan"
Write-Typewriter -Texto " ════════════════════════════════════════════════════════════════════════════════════" -Color "Cyan"
Write-Typewriter -Texto " Presiona [Esc] en cualquier momento para salir del diagnóstico" -Color "DarkGray"
Out-SmartWare ""

Write-Typewriter -Texto " [i] Verificando estabilidad del sistema..." -Color "Gray"

# Limpiar cualquier tecla presionada antes de empezar (Evita "escritura fantasma")
while ([console]::KeyAvailable) { $null = [console]::ReadKey($true) }

# --- RECOLECCIÓN DE MÉTRICAS (MODO MUESTREO 2s) ---
$equipoLento = $false

$preOS = Get-CimInstance Win32_OperatingSystem

# 1. SSD Check
$preWinsat = Get-CimInstance Win32_WinSAT -ErrorAction SilentlyContinue
$esSSD = if ($preWinsat) { $preWinsat.DiskScore -ge 6.0 } else { $false }

# [NUEVO] Detección de hilos lógicos para Umbral Dinámico
$precpuInfo = Get-CimInstance Win32_Processor
$numHilosLogicos = if ($precpuInfo.NumberOfLogicalProcessors) {
    $precpuInfo.NumberOfLogicalProcessors
} else {
    # Fallback 1: Variable de entorno del sistema (Muy confiable)
    if ($env:NUMBER_OF_PROCESSORS) {
        [int]$env:NUMBER_OF_PROCESSORS
    } else {
        # Fallback 2: Conteo manual vía CIM
        (Get-CimInstance Win32_Processor).NumberOfLogicalProcessors | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    }
}

# Si después de todo sigue siendo nulo o 0 (caso extremo), usamos un mínimo seguro de 2
if ($null -eq $numHilosLogicos -or $numHilosLogicos -le 0) { $numHilosLogicos = 2 }

$precpuModelo = if ($precpuInfo.Name) { $precpuInfo.Name.Trim() } else { "Requiere software especializado" }
$epocaCPU = "Antiguo" # Valor por defecto (Pre-2019 / Intel < 9th / Ryzen < 3000)
$pregenNum = 0
if ($precpuModelo -match "(i\d|Ryzen \d)[- ](?<num>\d{4,5})") {
    $fullNum = $Matches['num']
    $pregenNum = if ($fullNum.Length -eq 5) { [int]$fullNum.Substring(0, 2) } else { [int]$fullNum.Substring(0, 1) }
}
elseif ($precpuModelo -match "Ultra \d") {
    if ($precpuModelo -match "Ultra (?<gen>\d)") { $pregenNum = [int]$Matches['gen'] }
}

if ($precpuModelo -match "Ultra|Ryzen [789]\d{3}|i\d-1[34]") {
    $epocaCPU = "Moderno" # 2024-2026
}
elseif ($pregenNum -ge 10 -or ($precpuModelo -match "Ryzen" -and $pregenNum -ge 3)) {
    # Intel 10th+ o Ryzen 3000+ se consideran Intermedios/Modernos funcionales
    $epocaCPU = "Intermedio" # 2019-2023
}

# --- CÁLCULO DINÁMICO DE UMBRAL (v1.3.4 REVISADA) ---
$pisoBaseOS = if ($preOS.Version -ge "10.0.22000") { 3000 } else { 2000 }

$multiplicadorHilos = switch ($epocaCPU) {
    "Moderno"    { 500 }
    "Intermedio" { 375 }
    "Antiguo"    { 250 }
}

$umbralHilosDinamico = $pisoBaseOS + ($numHilosLogicos * $multiplicadorHilos)

# --- CAPTURA DE SENSORES (VERSIÓN FINAL BLINDADA) ---

function Exit-SmartWare {
    Write-Typewriter -Texto "`n`n [!] Diagnóstico detenido. Limpiando entorno..." -Color "Yellow"
    exit
}

# Nota: Durante estos intervalos de 1s, el usuario podría presionar Esc
try {
    $muestrasCPU = Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
} catch {
    $muestrasCPU = $null
    Write-Log -Mensaje "Fallo en Get-Counter CPU: $($_.Exception.Message)" -Gravedad "AVISO"
}
if ([console]::KeyAvailable) { if ([console]::ReadKey($true).Key -eq "Escape") { Exit-SmartWare } }

try {
    $muestrasDisk = Get-Counter '\PhysicalDisk(_Total)\Disk Transfers/sec' -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
} catch {
    $muestrasDisk = $null
    Write-Log -Mensaje "Fallo en Get-Counter Disco: $($_.Exception.Message)" -Gravedad "AVISO"
}
if ([console]::KeyAvailable) { if ([console]::ReadKey($true).Key -eq "Escape") { Exit-SmartWare } }

# Procesamiento CPU
if ($null -eq $muestrasCPU) {
    $preCPU = [math]::Round((Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average, 0)
} else {
    $preCPU = [math]::Round(($muestrasCPU.CounterSamples.CookedValue | Measure-Object -Average).Average, 0)
}

# Procesamiento Disco
if ($null -eq $muestrasDisk) {
    $preDiskIO = 50
} else {
    $preDiskIO = [math]::Round(($muestrasDisk.CounterSamples.CookedValue | Measure-Object -Average).Average, 1)
}

if ($null -eq $preCPU) { $preCPU = 15 }
if ($null -eq $preDiskIO) { $preDiskIO = 50 }

# 3. RAM y Otros
$preRAM = [math]::Round((($preOS.TotalVisibleMemorySize - $preOS.FreePhysicalMemory) / $preOS.TotalVisibleMemorySize) * 100, 1)
$preThreads = (Get-Process | Select-Object -ExpandProperty Threads).Count

# --- [NUEVO] ESCUDO DE ALMACENAMIENTO INTEGRADO ---
$letraOS = $env:SystemDrive.Replace(":","")
try {
    $discoOS = Get-Volume -DriveLetter $letraOS -ErrorAction Stop
    $particionOS = Get-Partition -DriveLetter $letraOS -ErrorAction Stop
    $numDiscoFisico = if ($particionOS) { $particionOS.DiskNumber } else { 0 }

    # Salud física (SMART)
    $smartStatus = Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction Stop |
                   Where-Object { $_.InstanceName -match "PHYSICALDRIVE$numDiscoFisico" } | Select-Object -First 1

    $healthStatus = Get-PhysicalDisk -DeviceNumber $numDiscoFisico -ErrorAction Stop
    $fallaReal = ($null -ne $smartStatus -and $smartStatus.PredictFailure) -or ($null -ne $healthStatus -and $healthStatus.HealthStatus -ne 'Healthy')
} catch {
    # Si falla, asumimos valores seguros para que el script no se cierre,
    # pero marcamos que no pudimos verificar la salud física.
    $fallaReal = $false
    Write-Log -Mensaje "No se pudo obtener telemetría de salud de disco: $($_.Exception.Message)"
}

$espacioLibreGB = [math]::Round($discoOS.SizeRemaining / 1GB, 1)
$porcLibre = [math]::Round(($discoOS.SizeRemaining / $discoOS.Size) * 100, 1)

# --- LÓGICA DE DETENCIÓN ---
$umbralDiscoPre = if($esSSD){ 1200 } else { 150 }
$alertas = @()

if ($espacioLibreGB -lt 3 -or $porcLibre -lt 3) {
    $alertas += " - ESPACIO EN C: CRÍTICO (${espacioLibreGB}GB / $porcLibre% libre)"
}
if ($preRAM -ge 96) { $alertas += " - Memoria RAM saturada: $preRAM%" }
if ($preCPU -ge 98) { $alertas += " - Procesador al límite: $preCPU%" }
if ($preDiskIO -ge $umbralDiscoPre) { $alertas += " - Actividad de Disco excesiva: $preDiskIO ops/sec" }
if ($preThreads -ge $umbralHilosDinamico) { $alertas += " - Congestión de Subprocesos: $preThreads (Límite para este CPU: $umbralHilosDinamico)" }
if ($fallaReal) { $alertas += " - FALLA FÍSICA DETECTADA en el disco del sistema" }

if ($alertas.Count -gt 0) {
    $cargaAltaRAM = $preRAM -ge 85
    $cargaAltaCPU = $preCPU -ge 75
    $congestionHilos = $preThreads -ge ($umbralHilosDinamico * 0.8)

    if ($porcLibre -lt 10 -or $cargaAltaRAM -or $cargaAltaCPU -or $congestionHilos) {
        $equipoLento = $true
    }

    $msjEscudo = "¡ALERTA! SISTEMA DE PROTECCIÓN ACTIVADO`n`n"
    $msjEscudo += "Se ha detenido el diagnóstico para evitar una mayor sobrecarga del sistema.`n"
    $msjEscudo += "──────────────────────────────────────────────────`n"
    $msjEscudo += ($alertas -join "`n") + "`n"
    $msjEscudo += "──────────────────────────────────────────────────`n"
    $msjEscudo += "Acciones recomendadas:`n"

    if ($espacioLibreGB -lt 3 -or $porcLibre -lt 3) { $msjEscudo += "- Libere espacio en el disco C: (Se recomienda >20GB libres).`n" }
    if ($preThreads -ge $umbralHilosDinamico) { $msjEscudo += "- Cierre programas pesados y espere o reinicie el equipo.`n" }
    if ($fallaReal) { $msjEscudo += "- FALLA FÍSICA DEL DISCO: Respalde sus datos de inmediato." }

    [System.Windows.Forms.MessageBox]::Show($msjEscudo, "SmartWare - Protección del Sistema", 0, 48)
    exit
}

# --- BLOQUE DE EXTRACCIÓN ---
if ($alertas.Count -eq 0) {
    $rutaTemp = Join-Path $env:TEMP "Herramientas SmartWare"
    if (Test-Path $rutaTemp) { [void](Remove-Item $rutaTemp -Recurse -Force -ErrorAction SilentlyContinue) }
    [void](New-Item -ItemType Directory -Path $rutaTemp -Force)

    try {
        Write-Typewriter -Texto "`n [OK] Sistema estable." -Color "Green"
        Write-Typewriter -Texto "`n [i] INICIANDO MOTOR: Preparando herramientas de SmartWare..." -Color "Cyan"

        $zipPath = Join-Path $rutaTemp "tools.zip"
        $base64Limpio = $base64Zip.Trim() -replace "[\r\n\s]", ""
        try {
            $bytes = [System.Convert]::FromBase64String($base64Limpio)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("El motor de herramientas está corrupto. Por favor, descargue el script nuevamente.", "Error de Integridad", 0, 16)
            exit
        }
        [System.IO.File]::WriteAllBytes($zipPath, $bytes)

        $originalProgress = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Expand-Archive -Path $zipPath -DestinationPath $rutaTemp -Force | Out-Null
        } finally {
            $ProgressPreference = $originalProgress
        }

        [void](Remove-Item $zipPath -Force)

        Write-Typewriter -Texto "`n [OK] Herramientas cargadas correctamente." -Color "Green"
        Write-Typewriter -Texto "`n [!] TODO LISTO: Presiona ENTER para continuar" -Color "Yellow"

        while ($true) {
            if ([console]::KeyAvailable) {
                $key = [console]::ReadKey($true)
                if ($key.Key -eq "Enter") { break }
                if ($key.Key -eq "Escape") { Exit-SmartWare }
            }
            Start-Sleep -Milliseconds 50
        }

        $global:t = $rutaTemp

    } catch {
        $errorMsg = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Error técnico de extracción: $errorMsg", "SmartWare - Error Crítico", 0, 16)
        exit
    }
}

# [2. VALIDACIÓN DE RESPUESTAS]
Clear-Host

# 1. IDENTIDAD VISUAL ---
Out-SmartWare -Texto " ════════════════════════════════════════════════════════════════════════════════════" -Color "Cyan"
Out-SmartWare -Texto " SMARTWARE - Potencia tu Mundo Digital" -Color "Cyan"
Out-SmartWare -Texto " DIAGNÓSTICO INTELIGENTE v$version" -Color "Cyan"
Out-SmartWare -Texto " PANEL DE CONTROL" -Color "Cyan"
Out-SmartWare -Texto " ════════════════════════════════════════════════════════════════════════════════════" -Color "Cyan"
Out-SmartWare -Texto " Presiona [Esc] en cualquier momento para salir del diagnóstico" -Color "DarkGray"
Out-SmartWare ""

Write-Typewriter -Texto " [i] NOTA: Todas las respuestas se deben confirmar con la tecla ENTER 2 veces." -Color "Gray"

function Get-Respuesta($pregunta) {
    try {
        $posEntrada = $host.UI.RawUI.CursorPosition
    } catch {
        $posEntrada = $null
    }
    $preguntaDibujada = $false

    while($true) {
        if (-not $preguntaDibujada) {
            if ($null -ne $posEntrada) { $host.UI.RawUI.CursorPosition = $posEntrada }

            Write-Typewriter -Texto "`r`n$pregunta " -NoNewLine -Color "Gray"
            Out-SmartWare -Texto "S/N > " -Color "Cyan" -NoNewLine
            $preguntaDibujada = $true
        }

        # 2. CAPTURA CRÍTICA
        $posCursorRespuesta = try { $host.UI.RawUI.CursorPosition } catch { $null }
        [console]::CursorVisible = $true

        $resp = ""
        while($true) {
            if (![console]::KeyAvailable) { Start-Sleep -Milliseconds 20; continue }
            $tecla = [console]::ReadKey($true)

            if ($tecla.Key -eq "Escape") { Exit-SmartWare }
            if ($tecla.Key -eq "Enter") {
                if ($resp.Length -gt 0) { break }
                continue
            }
            if ($tecla.Key -eq "Backspace") {
                if ($resp.Length -gt 0) {
                    $resp = $resp.Substring(0, $resp.Length - 1)
                    Out-SmartWare -Texto "`b `b" -NoNewLine
                }
                continue
            }
            $resp += $tecla.KeyChar
            Out-SmartWare -Texto $tecla.KeyChar -NoNewLine
        }

        [console]::CursorVisible = $false

        # VALIDACIÓN
        if ($resp -match "^(s|si|n|no)$") {
            Out-SmartWare -Texto " [Borrar: Corregir]" -Color "DarkGray" -NoNewLine

            while($true) {
                if (![console]::KeyAvailable) { Start-Sleep -Milliseconds 20; continue }
                $confirm = [console]::ReadKey($true)

                if ($confirm.Key -eq "Escape") { Exit-SmartWare }
                if ($confirm.Key -eq "Enter") {
                    Out-SmartWare ""
                    return ($resp -match "^(s|si)$")
                }
                if ($confirm.Key -eq "Backspace") {
                    if ($null -ne $posCursorRespuesta) {
                        $host.UI.RawUI.CursorPosition = $posCursorRespuesta
                        $anchoMaxWin = [math]::Max(0, ($host.UI.RawUI.WindowSize.Width - $posCursorRespuesta.X - 1))
                        $espaciosABorrar = [math]::Min(($resp.Length + 45), $anchoMaxWin)
                        Out-SmartWare -Texto (" " * $espaciosABorrar) -NoNewLine
                        $host.UI.RawUI.CursorPosition = $posCursorRespuesta
                    }
                    break
                }
            }
        } else {
            Out-SmartWare -Texto " [!] Error: Use S o N." -Color "Red"
            Start-Sleep -Milliseconds 800
            if ($null -ne $posCursorRespuesta) {
                $host.UI.RawUI.CursorPosition = $posCursorRespuesta
                Out-SmartWare -Texto (" " * [math]::Max(0, ($host.UI.RawUI.WindowSize.Width - $posCursorRespuesta.X))) -NoNewLine
                $host.UI.RawUI.CursorPosition = $posCursorRespuesta
            }
        }
    }
}

function Write-Progreso($tarea, $segundos) {
    $posInicial = try { $host.UI.RawUI.CursorPosition } catch { $null }

    Write-Typewriter -Texto "$tarea" -Color "Gray" -NoNewLine

    if ($null -ne $posInicial) {
        $posBarraEstructura = $posInicial
        $posBarraEstructura.Y += 2
        $host.UI.RawUI.CursorPosition = $posBarraEstructura
        Out-SmartWare -Texto " [░░░░░░░░░░░░░░░░░░░░] " -Color "DarkGray" -NoNewline
    }

    $porcentajeActual = 0
    $bloquesDibujados = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $msTotales = if ($segundos -le 0) { 1000 } else { $segundos * 1000 }

    while ($porcentajeActual -lt 100) {
        if ([console]::KeyAvailable) {
            if ([console]::ReadKey($true).Key -eq "Escape") { Exit-SmartWare }
        }

        $incremento = Get-Random -Minimum 1 -Maximum 13
        $porcentajeActual += $incremento
        if ($porcentajeActual -gt 100) { $porcentajeActual = 100 }

        $bloquesObjetivo = [math]::Floor($porcentajeActual / 5)

        while ($bloquesDibujados -lt $bloquesObjetivo) {
            $bloquesDibujados++
            if ($null -ne $posInicial) {
                $posLlenado = $posInicial
                $posLlenado.Y += 2
                $posLlenado.X += (1 + $bloquesDibujados)
                $host.UI.RawUI.CursorPosition = $posLlenado
                Out-SmartWare -Texto "█" -Color "Cyan" -NoNewLine
            }
        }

        if ($null -ne $posInicial) {
            $posNum = $posInicial
            $posNum.Y += 2
            $posNum.X += 24
            $host.UI.RawUI.CursorPosition = $posNum
            Out-SmartWare -Texto ($porcentajeActual.ToString().PadLeft(3) + "%") -Color "Cyan" -NoNewLine
        }

        $tiempoTranscurrido = $sw.ElapsedMilliseconds
        $progresoEsperado = ($tiempoTranscurrido / $msTotales) * 100
        if ($porcentajeActual -gt $progresoEsperado) {
            Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 150)
        }
    }

    if ($null -ne $posInicial) {
        $posNum.X += 5
        $host.UI.RawUI.CursorPosition = $posNum
        Out-SmartWare -Texto "- " -Color "DarkGray" -NoNewLine
        Out-SmartWare -Texto "OK" -Color "Green"
    }

    $sw.Stop()
    Out-SmartWare ""
}

# --- NUEVA PREGUNTA DE ASISTENCIA ---
$esTecnico = Get-Respuesta " ¿Eres técnico o usuario con conocimientos avanzados?"

# 2. VERIFICACIÓN DE HERRAMIENTAS
if ($esTecnico) {
    Out-SmartWare ""
    Write-Progreso " Verificando integridad de herramientas de diagnóstico..." 1.5
    $appsClave = @("DiskInfo", "HWMonitor", "TreeSizeFree", "adwcleaner")
    $rutaValida = $false

    if ($null -ne $global:t -and (Test-Path $global:t)) {
        $conteo = 0
        foreach ($app in $appsClave) {
            if (Get-ChildItem -Path $global:t -Filter "$app*.exe" -File -Recurse -Depth 2 -ErrorAction SilentlyContinue) { $conteo++ }
        }
        if ($conteo -ge 3) { $rutaValida = $true }
    }

    if (-not $rutaValida) {
        $conteoLocal = 0
        foreach ($app in $appsClave) {
            if (Get-ChildItem -Path $PSScriptRoot -Filter "$app*.exe" -File -Recurse -Depth 2 -ErrorAction SilentlyContinue) { $conteoLocal++ }
        }
        if ($conteoLocal -ge 3) {
            $global:t = $PSScriptRoot
            $rutaValida = $true
        }
    }

    if (-not $rutaValida) {
        Write-Log -Mensaje "ERROR: No se hallaron herramientas." -Gravedad "ALTA"
        [System.Windows.Forms.MessageBox]::Show("ERROR CRÍTICO: No se localizaron las herramientas.", "SmartWare", 0, 16)
        exit
    }
    Write-Typewriter -Texto " [+] Entorno de herramientas validado correctamente." -Color "Green"
}

# 3. SEGURO DE VIDA SMARTWARE (Punto de Restauración)
Write-Typewriter -Texto "`n [i] El Punto de Restauración permite revertir cambios en la configuración" -Color "Gray"
Write-Typewriter -Texto "     de Windows si fuera necesario. No afecta sus archivos personales." -Color "Gray"
if (Get-Respuesta " ¿Deseas realizar un punto de restauración? (Recomendado)") {
    $omitirPorReciente = $false
    try {
        Write-Typewriter -Texto "`n [i] Revisando si existe un Punto de Restauración reciente. Por favor, espere..." -Color "Gray"
        $ultimoPunto = Get-ComputerRestorePoint -ErrorAction Stop | Sort-Object CreationTime -Descending | Select-Object -First 1

        if ($ultimoPunto) {
            $fechaPunto = [Management.ManagementDateTimeConverter]::ToDateTime($ultimoPunto.CreationTime)
            $limite24h = (Get-Date).AddDays(-1)
            if ($fechaPunto -gt $limite24h) {
                $omitirPorReciente = $true
                Write-Typewriter -Texto "`n [!] Ya existe un punto de restauración reciente ($($fechaPunto.ToString('dd/MM HH:mm')))." -Color "Yellow"
                Write-Typewriter -Texto " [i] Se omitirá la creación para evitar duplicidad." -Color "DarkGray"
            }
        }
    } catch {
        Write-Log -Mensaje "No se pudo consultar puntos de restauración: $($_.Exception.Message)" -Gravedad "AVISO"
    }

    if (-not $omitirPorReciente) {
        try {
            Enable-ComputerRestore -Drive "C:\" -ErrorAction Stop
            Write-Typewriter -Texto "`n [!] Creando punto de restauración. Por favor, espere..." -Color "Yellow"
            Checkpoint-Computer -Description "SmartWare_Pre_Diag_$(Get-Date -Format 'ddMMyy')" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
            Write-Typewriter -Texto " [+] Punto de restauración creado con éxito." -Color "Green"
        } catch {
            Write-Typewriter -Texto "`n [!] Aviso: No se pudo crear el punto (Verificar protección del sistema)." -Color "Red"
            Write-Log -Mensaje "Fallo en Checkpoint-Computer: $($_.Exception.Message)" -Gravedad "ERROR"
        }
    }
} else {
    Write-Typewriter -Texto "`n [i] Creación del punto de restauración omitida por el usuario." -Color "DarkGray"
}

# --- BLOQUE 1: LENTITUD Y SEGURIDAD ---
$equipoLento = Get-Respuesta " ¿Notas lentitud al trabajar o al iniciar el equipo?"
$quiereSeguridad = $false
if ($equipoLento) {
    Write-Typewriter -Texto "`n [!] ADVERTENCIA: Debido a la lentitud, el análisis de seguridad podría saturar el sistema." -Color "Yellow"
    $quiereSeguridad = Get-Respuesta " ¿Deseas ejecutar el análisis de todas formas?"
} else {
    if ($preRAM -gt 80 -or $preCPU -gt 70) {
        Write-Typewriter -Texto "`n [i] NOTA: Se detecta carga alta ($preRAM% RAM / $preCPU% CPU)." -Color "Cyan"
    }
    $quiereSeguridad = Get-Respuesta " ¿Deseas realizar un Análisis de Seguridad? (Recomendado)"
}

if ($quiereSeguridad) {
    try {
        $adwPath = Get-ChildItem -Path $global:t -Filter "adwcleaner*" -Recurse -ErrorAction Stop | Select-Object -ExpandProperty FullName -First 1
        if ($adwPath) {
            Write-Typewriter -Texto "`n [i] Iniciando aplicación de análisis..." -Color "Cyan"
            Start-Process -FilePath $adwPath -ArgumentList "/eula /scan /noreboot" -Wait -PassThru -ErrorAction Stop
            Write-Typewriter -Texto " [+] Análisis completado." -Color "Green"
        }
    } catch {
        Write-Typewriter -Texto "`n [!] No se pudo iniciar el análisis: $($_.Exception.Message)" -Color "Red"
    }
}

# [BLOQUE WIN32 CORREGIDO]
try {
    if (-not ("Win32Functions.Win32ShowWindowAsync" -as [type])) {
        $win32Code = @"
        using System;
        using System.Runtime.InteropServices;
        namespace Win32Functions {
            public class Win32ShowWindowAsync {
                [DllImport("user32.dll")]
                public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
            }
        }
"@
        Add-Type -TypeDefinition $win32Code -ErrorAction Stop
    }
} catch {
    Write-Log -Mensaje "No se pudo inyectar Win32. Usando modo estándar."
}

# --- NORMALIZACIÓN Y CÁLCULO DE CARGA ---
$pesoDisk = [math]::Min(100, ($preDiskIO / $umbralDiscoPre) * 100)
$pesoThreads = [math]::Min(100, ($preThreads / $umbralHilosDinamico) * 100)
$cargaEspacioDisco = 100 - $porcLibre

$cargaSistemaReal = [math]::Min(100, [math]::Round(
    ($preCPU * 0.35) + ($preRAM * 0.25) + ($pesoDisk * 0.20) + ($pesoThreads * 0.10) + ($cargaEspacioDisco * 0.10), 0))

$physicalDisks = Get-PhysicalDisk -ErrorAction SilentlyContinue
if ($null -eq $physicalDisks) {
    # Fallback: Intentar obtener al menos los discos básicos por CIM
    $physicalDisks = Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue
}

$sistemaEnHDD = $false
$contieneHDD = $false

foreach ($disk in $physicalDisks) {
    # --- Lógica de Tipo de Disco ---
    if ($disk.MediaType -eq 'HDD') {
        $contieneHDD = $true
        # Si el HDD que estamos recorriendo coincide con el disco del OS que detectamos en la Secc 1
        if ($disk.DeviceNumber -eq $numDiscoFisico) {
            $sistemaEnHDD = $true
        }
    }
}

if ($sistemaEnHDD -and $pesoDisk -gt 80) {
    $cargaSistemaReal = [math]::Min(100, ($cargaSistemaReal + 20))
}

function Write-TelemetriaDinamica($tarea, $segundos, $valorFinal) {
    $pasos = 20
    $msPorPaso = ($segundos * 1000) / $pasos
    $posInicial = $host.UI.RawUI.CursorPosition

    # 1. TÍTULO (Solo el texto, sin paréntesis todavía)
    Write-Typewriter -Texto "$tarea" -Color "Gray" -NoNewLine

    # 2. ESTRUCTURA BARRA (Línea de abajo)
    $posBarraEstructura = $posInicial
    $posBarraEstructura.Y += 2
    $host.UI.RawUI.CursorPosition = $posBarraEstructura
    Out-SmartWare -Texto " [░░░░░░░░░░░░░░░░░░░░] " -Color "DarkGray" -NoNewLine

    $numActual = 0

    for ($i = 0; $i -le $pasos; $i++) {
        if ([console]::KeyAvailable) {
            if ([console]::ReadKey($true).Key -eq "Escape") { Exit-SmartWare }
        }

        # --- LÓGICA DE NÚMEROS (CARGA REAL) ---
        if ($i -eq 0) { $numDisplay = Get-Random -Minimum 1 -Maximum 10 }
        elseif ($i -lt $pasos) {
            $pasosRestantes = $pasos - $i
            $techoDinamico = [math]::Max($numActual + 1, [math]::Round($valorFinal - ($pasosRestantes * 0.5)))
            $numDisplay = if ($numActual -ge $valorFinal) { $valorFinal } else { Get-Random -Minimum $numActual -Maximum ([math]::Max($numActual + 1, $techoDinamico)) }
        } else { $numDisplay = $valorFinal }
        $numActual = $numDisplay

        # --- 3. DIBUJO DE LA CARGA (Línea 1 - Segmentado por colores) ---
        $posCarga = $posInicial
        $posCarga.Y += 0
        $posCarga.X += ($tarea.Length + 1)
        $host.UI.RawUI.CursorPosition = $posCarga

        # A. Paréntesis de apertura (Gris)
        Out-SmartWare -Texto "(" -Color "DarkGray" -NoNewLine

        # B. Número y % (Cyan)
        # No usamos PadLeft para que quede pegado al (
        Out-SmartWare -Texto "Carga actual: $numDisplay%" -Color "DarkGray" -NoNewLine

        # C. Paréntesis de cierre y espacios de limpieza (Gris)
        Out-SmartWare -Texto ")  " -Color "DarkGray" -NoNewLine

        # --- 4. RELLENO DE LA BARRA (Línea 2) ---
        if ($i -gt 0) {
            $posLlenado = $posInicial
            $posLlenado.Y += 2
            $posLlenado.X += ($i + 1)
            $host.UI.RawUI.CursorPosition = $posLlenado
            Out-SmartWare -Texto "█" -Color "Cyan" -NoNewLine
        }

        # --- 5. DIBUJO DEL % DE PROGRESO (Línea 2) ---
        $progresoAnim = [math]::Round(($i / $pasos) * 100)
        $posProg = $posInicial
        $posProg.Y += 2
        $posProg.X += 24
        $host.UI.RawUI.CursorPosition = $posProg
        Out-SmartWare -Texto ($progresoAnim.ToString().PadLeft(3) + "%") -Color "Cyan" -NoNewLine

        # --- 6. VEREDICTO FINAL ---
        if ($i -eq $pasos) {
            $colorTxt = if($valorFinal -ge 90){ "Red" } elseif($valorFinal -ge 70){ "Yellow" } else { "Green" }
            $msgTxt   = if($valorFinal -ge 90){ "CRÍTICO" } elseif($valorFinal -ge 70){ "ELEVADO" } else { "OK" }

            Out-SmartWare -Texto " - " -Color "DarkGray" -NoNewLine
            Out-SmartWare -Texto $msgTxt -Color $colorTxt
        }

        Start-Sleep -Milliseconds $msPorPaso
    }

    Out-SmartWare "" # Salto de línea de seguridad
}

[System.GC]::Collect()

# [3. SHOW DE DIAGNÓSTICO SEGMENTADO]
Clear-Host
Out-SmartWare -Texto " ════════════════════════════════════════════════════════════════════════════════════" -Color "Cyan"
Out-SmartWare -Texto " SMARTWARE - Potencia tu Mundo Digital" -Color "Cyan"
Out-SmartWare -Texto " DIAGNÓSTICO INTELIGENTE v$version" -Color "Cyan"
Out-SmartWare -Texto " PANEL DE CONTROL" -Color "Cyan"
Out-SmartWare -Texto " ════════════════════════════════════════════════════════════════════════════════════" -Color "Cyan"
Out-SmartWare -Texto " Presiona [Esc] en cualquier momento para salir del diagnóstico" -Color "DarkGray"
Out-SmartWare ""

# --- EJECUCIÓN DEL MOTOR VISUAL ---
Write-Progreso " Iniciando motor de análisis profundo..." 2.0
Write-Progreso " Sincronizando sensores de hardware..." 2.5

if ($null -eq $cargaSistemaReal) { $cargaSistemaReal = 30 } # Valor neutro
if ($cargaSistemaReal -ge 60) { $equipoLento = $true }
if ($null -eq $equipoLento) { $equipoLento = $false }

# Uso de la función dinámica con la carga ponderada calculada en la sección anterior
Write-TelemetriaDinamica " Analizando carga del sistema..." 3.0 $cargaSistemaReal

# [3.1 DETECCIÓN DE ARQUITECTURA]
try {
    $es64 = [Environment]::Is64BitOperatingSystem
    $sufijoArch = if ($es64) { "64" } else { "32" }
} catch {
    # Fallback seguro a 32 bits si la consulta falla
    $es64 = $false
    $sufijoArch = "32"
    Write-Log -Mensaje "No se pudo determinar la arquitectura del OS de forma nativa." -Gravedad "AVISO"
}

# [3.2 PREPARACIÓN DE ENTORNO SEGÚN ESTADO]
if (-not $equipoLento) {
    Write-Typewriter -Texto " [i] Carga óptima. Preparando interfaz de diagnóstico..." -Color "Gray"
    Out-SmartWare ""
    Write-Progreso " Revisando componentes del equipo..." 2.0
} else {
    Out-SmartWare ""
    Write-Typewriter -Texto " [!] Carga elevada detectada. Optimizando recursos para el panel..." -Color "Yellow"
    Out-SmartWare ""
    Write-Progreso " Aislando procesos de alta carga..." 3.0
    Write-Progreso " Revisando componentes del equipo..." 2.0
}

# Efecto visual de transición (Espera rápida)
try {
    for($k=0; $k -lt 20; $k++){
        if([console]::KeyAvailable) {
            $key = [console]::ReadKey($true)
            if($key.Key -eq "Escape"){ Exit-SmartWare }
        }
        Start-Sleep -Milliseconds 100
    }
} catch {
    # Si la consola no permite leer teclas, simplemente esperamos los 2 segundos
    Start-Sleep -Seconds 2
}

# [LIMPIEZA DE MEMORIA POST-SHOW]
try {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} catch {
    # Ignorar si no se puede liberar memoria en este instante
}

# [3.2 LANZAMIENTO DE HERRAMIENTAS Y CONTROL DE VENTANA]

# Solo se ejecuta si se requiere asistencia, no es un equipo crítico/lento y la ruta base existe
if ($esTecnico -and -not $equipoLento -and $global:t) {

    # --- MUESTREO DE DISCO (IOPS) CON MANEJO DE ERRORES ---
    $diskIO = try {
        # Intentamos obtener el contador de rendimiento de disco (Lecturas/Escrituras por segundo)
        # Ajuste: Validación de disponibilidad del contador antes de la ejecución
        $datosContador = Get-Counter '\PhysicalDisk(_Total)\Disk Transfers/sec' -SampleInterval 1 -MaxSamples 2 -ErrorAction Stop

        if ($datosContador -and $datosContador.CounterSamples) {
            $promedioIO = ($datosContador.CounterSamples.CookedValue | Measure-Object -Average).Average
            if ($null -eq $promedioIO) { 50 } else { $promedioIO }
        } else {
            50 # Valor de contingencia si el contador no responde
        }
    } catch {
        # Si el contador de Windows está dañado o no disponible, asignamos precaución media
        Write-Log -Mensaje "Contadores de rendimiento de disco no disponibles. Usando valor base." -Gravedad "BAJA"
        50
    }

    # Validación de seguridad para la variable
    if ($null -eq $diskIO) { $diskIO = 50 }

    # Ajuste de umbral según tecnología: Los HDD sufren mucho más con pocas operaciones
    $umbralDisco = switch ($true) {
    ($esNVMe) { 5000 } # Los NVMe aguantan mucho más proceso antes de sentirse "lentos"
    ($esSSD)  { 1000 }
    Default   { 150 }  # HDD
    }

    # Determinamos si el equipo está bajo estrés (CPU > 80%, RAM > 75% o Disco saturado)
    $usoElevado = ($preRAM -ge 75 -or $preCPU -ge 80 -or $diskIO -ge $umbralDisco)
    $listaApps = @()

    $sufijoArch = if($es64){ "64" } else { "32" }

    # Diccionario de herramientas disponibles en el paquete
    $appsDisp = [ordered]@{
        "1" = @{ File = "DiskInfo$($sufijoArch).exe"; Nom = "DiskInfo (Salud)" }
        "2" = @{ File = "HWMonitor_x$($sufijoArch).exe"; Nom = "HWMonitor (Temperaturas)" }
        "3" = @{ File = "TreeSizeFree.exe"; Nom = "TreeSize (Espacio)" }
    }

    Write-Typewriter -Texto " [i] Diagnóstico base completado. Preparando monitoreo detallado..." -Color "Cyan"
    Out-SmartWare ""

    if ($usoElevado) {
        $detallesCarga = @()
        if ($preRAM -ge 75) { $detallesCarga += "RAM: $preRAM%" }
        if ($preCPU -ge 80) { $detallesCarga += "CPU: $preCPU%" }
        if ($diskIO -ge $umbralDisco) { $detallesCarga += "DISCO: $([math]::Round($diskIO,1)) ops/sec" }

        Write-Typewriter -Texto " [!] CARGA ELEVADA DETECTADA: $($detallesCarga -join " / ")" -Color "Yellow"
        Write-Typewriter -Texto " [i] Seleccione herramientas manualmente para evitar saturación del sistema." -Color "Gray"

        # Bucle de selección manual en caso de estrés
        do {
            # Ajuste: Se agrega un mensaje más descriptivo para el input
            $opc = (Read-Host " S/N (0: Ninguna, 1: DiskInfo, 2: HWMonitor, 3: TreeSize, A: Todas)").ToUpper()
            $valido = $true
            if ($opc -eq "A") {
                $listaApps = $appsDisp.Values
            }
            elseif ($opc -eq "0") {
                $listaApps = @()
            }
            elseif ($appsDisp.ContainsKey($opc)) {
                $listaApps = @($appsDisp[$opc])
            }
            else {
                Write-Typewriter -Texto " [!] Error: Respuesta inválida." -Color "Red"
                $valido = $false
            }
        } while (-not $valido)
    } else {
        # Carga normal: se preparan todas las aplicaciones para apertura automática
        $listaApps = $appsDisp.Values
    }

    # --- DESPLIEGUE DE LAS APLICACIONES SELECCIONADAS ---
    foreach ($appInfo in $listaApps) {
        $fileName = $appInfo.File
        $friendlyName = $appInfo.Nom
        # Ajuste: Manejo robusto de la ruta del proceso
        $nombreProceso = [System.IO.Path]::GetFileNameWithoutExtension($fileName)

        # Solo abrimos si no está ya en ejecución
        if (-not (Get-Process $nombreProceso -ErrorAction SilentlyContinue)) {
            Write-Progreso " Desplegando $friendlyName..." 2

            # Búsqueda recursiva protegida en la carpeta definida en $global:t
            $rutaReal = try {
                Get-ChildItem -Path $global:t -Filter $fileName -Recurse -File -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty FullName -First 1
            } catch { $null }

            if ($rutaReal) {
                $workingDir = Split-Path $rutaReal
                try {
                    # Ajuste: Start-Process con validación de existencia de archivo
                    if (Test-Path $rutaReal) {
                        Start-Process -FilePath $rutaReal -WorkingDirectory $workingDir -WindowStyle Normal -ErrorAction Stop
                    } else { throw "Ruta no encontrada tras validación." }
                } catch {
                    Write-Typewriter -Texto " [!] Error al iniciar $friendlyName : $($_.Exception.Message)" -Color "Red"
                    Write-Log -Mensaje "Fallo al lanzar $friendlyName en $rutaReal" -Gravedad "ERROR"
                }
            } else {
                Write-Typewriter -Texto " [!] Archivo $fileName no encontrado en el paquete." -Color "Red"
            }
        } else {
            Write-Typewriter -Texto " [i] $friendlyName ya se encuentra abierto." -Color "Gray"
        }
    }
    Write-Typewriter -Texto " [+] Todas las herramientas han sido desplegadas." -Color "Green"
    Out-SmartWare ""
}

# --- CIERRE UNIFICADO DEL SHOW (Para todos los usuarios) ---
Write-Progreso " Cargando información del sistema..." 2.5
Write-Progreso " Compilando reporte final de diagnóstico..." 2.0
Write-Typewriter -Texto " [!] Todo listo. Abriendo Interfaz del Diagnóstico..." -Color "Yellow"

# Efecto visual de transición (Espera de 2 segundos con opción de salir)
try {
    for($k=0; $k -lt 20; $k++){
        if([console]::KeyAvailable) {
            $key = [console]::ReadKey($true)
            if($key.Key -eq "Escape"){ Exit-SmartWare }
        }
        Start-Sleep -Milliseconds 100
    }
} catch {
    Start-Sleep -Seconds 2
}

# [LIMPIEZA DE MEMORIA FINAL]
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# [4. RECOLECCIÓN TÉCNICA]
# 1. IDENTIFICACIÓN DEL CHASIS Y TIPO DE EQUIPO
# Determina si es un sistema fijo o portátil basado en el estándar SMBIOS
try {
    $chasis = Get-CimInstance Win32_SystemEnclosure -ErrorAction Stop | Select-Object -ExpandProperty ChassisTypes
    $esNotebook = ($chasis -in @(8, 9, 10, 11, 12, 14, 30, 31, 32))
} catch {
    $esNotebook = $false # Asumimos Fijo por defecto
    Write-Log -Mensaje "Fallo al detectar chasis: $($_.Exception.Message)"
}
$tipoEquipo = if($esNotebook){ "Portátil" } else { "Fijo" }

# 2. ANÁLISIS PROFUNDO DE MEMORIA RAM
# Extrae velocidad máxima, configurada y calcula generación por frecuencias
try {
    $ri = @(Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop)
    if ($ri.Count -gt 0) {
        $vMax = ($ri | Measure-Object Speed -Maximum).Maximum
        $vActual = ($ri | Measure-Object ConfiguredClockSpeed -Maximum).Maximum
        if ($vActual -le 0) { $vActual = $vMax }
        $totalRAM_GB = [math]::Round(($ri | Measure-Object Capacity -Sum).Sum / 1GB, 0)
        $modSize = [math]::Round($ri[0].Capacity / 1GB, 0)
    } else {
        throw "No se detectaron módulos"
    }
} catch {
    $totalRAM_GB = [math]::Round((Get-CimInstance Win32_OperatingSystem).TotalVisibleMemorySize / 1MB, 0)
    $vActual = 0
    $modSize = 0
    Write-Log -Mensaje "RAM Genérica: $($_.Exception.Message)"
}

# Detección de DDR5 basado en JEDEC (>4800MT/s) o enumeración SMBIOS (34)
$genRAM = try {
    # 1. Intentamos obtener el código de tipo (SMBIOS)
    $typeCode = if ($null -ne $ri -and $ri.Count -gt 0) { $ri[0].MemoryType } else { 0 }

    # 2. Escalera de coherencia (Frecuencia vs Código)
    if ($typeCode -eq 34 -or $vMax -ge 4700) {
        "DDR5"
    }
    elseif ($typeCode -eq 26 -or ($vMax -ge 2133 -and $vMax -lt 4700)) {
        "DDR4"
    }
    elseif ($typeCode -eq 24 -or ($vMax -gt 0 -and $vMax -lt 2133)) {
        "DDR3"
    }
    else {
        "DDR/Desconocida"
    }
}
catch {
    # 3. Si WMI falla catastróficamente, usamos la frecuencia como último recurso
    if ($vMax -ge 4700) { "DDR5" }
    elseif ($vMax -ge 2133) { "DDR4" }
    elseif ($vMax -gt 0) { "DDR3" }
    else { "DDR/Desconocida" }
}

$totalRAM_GB = [math]::Round(($ri | Measure-Object Capacity -Sum).Sum / 1GB, 0)
$modSize = if($ri.Count -gt 0){ [math]::Round($ri[0].Capacity / 1GB, 0) } else { 0 }
$canal = if($ri.Count -gt 1){ "Dual Channel" } else { "Single Channel" }

$formatRAM = "$($totalRAM_GB)GB ($genRAM/$($canal)/$($ri.Count)x$($modSize)GB/${vActual}MHz)"

# 3. ANÁLISIS DE PROCESADOR (Lógica de Generación y Vigencia v1.3.3)
$cpuInfo = Get-CimInstance Win32_Processor
$cpuModelo = $cpuInfo.Name.Trim()
$añoActual = (Get-Date).Year

# --- A. EXTRACCIÓN DE GENERACIÓN ---
$genNum = 0
try {
    if ($cpuModelo -match "(i\d|Ryzen \d)[- ](?<num>\d{4,5})") {
        $fullNum = $Matches['num']
        $genNum = if ($fullNum.Length -eq 5) { [int]$fullNum.Substring(0, 2) } else { [int]$fullNum.Substring(0, 1) }
    } elseif ($cpuModelo -match "Ultra \d") {
        if ($cpuModelo -match "Ultra (?<gen>\d)") { $genNum = [int]$Matches['gen'] }
    }
} catch {
    $genNum = 0
    Write-Log -Mensaje "Error al parsear generación de CPU ($cpuModelo)"
}

# --- B. CRITERIO DE VIGENCIA (Matriz C. Díaz) ---
$esVigente = ($cpuModelo -match "Intel" -and $genNum -ge 12) -or
             ($cpuModelo -match "Ryzen" -and $genNum -ge 5) -or
             ($cpuModelo -match "Core Ultra|Snapdragon|Apple")

# --- C. CLASIFICACIÓN DE GAMA ---
$gamaCPU = "Estándar PRE-$añoActual"
if ($cpuModelo -match "Celeron|Pentium|Athlon|Silver|Gold|N\d{2,3}") {
    $gamaCPU = "Entrada"
}
elseif ($esVigente -and ($cpuModelo -match "[579]-|Ryzen [579]|Ultra [579]")) {
    $gamaCPU = "Alta (Estándar $añoActual)"
}
elseif ($esVigente) {
    $gamaCPU = "Media (Estándar $añoActual)"
}
elseif ($numHilosLogicos -ge 8) {
    $gamaCPU = "Media (Clásica)"
}

# --- D. PERFIL DE TRABAJO ---
$perfilCPU = "Estándar"
if ($cpuModelo -match "[UYG]$|N\d{2,4}") {
    $perfilCPU = "Bajo Consumo (Eficiencia/Movilidad)"
}
elseif ($cpuModelo -match "[HKX]$") {
    $perfilCPU = "Alto Rendimiento (Gaming/Productividad)"
}

# --- E. EXTRACCIÓN DE CACHÉ L3 ---
$cacheL3 = 0
try {
    # Buscamos específicamente la caché de nivel 3 (Level 4 en el enum de WMI, que corresponde a L3)
    $cacheData = Get-CimInstance Win32_CacheMemory | Where-Object { $_.Level -eq 5 -or $_.Purpose -like "*L3*" }

    if ($cacheData) {
        # Sumamos en caso de que reporte múltiples bloques (común en arquitecturas multi-chiplet)
        $totalL3KB = ($cacheData | Measure-Object -Property MaxCacheSize -Sum).Sum
        $cacheL3 = [math]::Round($totalL3KB / 1024, 0) # Convertir KB a MB
    }
} catch {
    $cacheL3 = 0
    Write-Log -Mensaje "Error al obtener caché L3"
}

# --- REFINAMIENTO DE GAMA (Lógica de Respuesta L3) ---
if ($cacheL3 -ge 32) {
    $gamaCPU += " [Ultra-Rápida]"
}
elseif ($cacheL3 -ge 8) {
    # No añadimos etiqueta o usamos una sutil para no saturar el reporte
    $gamaCPU += " [Fluida]"
}
else {
    $gamaCPU += " [Limitada]"
}

# 4. ANÁLISIS DE PLACA BASE Y BUS PCIe
try {
    $baseBoard = Get-CimInstance Win32_BaseBoard -ErrorAction Stop
    $rawManufacturer = if ($baseBoard.Manufacturer) { $baseBoard.Manufacturer.Trim() } else { "Unknown" }
    $rawProduct = if ($baseBoard.Product) { $baseBoard.Product.Trim() } else { "Unknown" }
} catch {
    $rawManufacturer = "Genérica"
    $rawProduct = "Placa"
}

# Normalización de nombres de fabricantes (Traductor de Identidad)
$diccionarioOEM = @{
    "HEWLETT-PACKARD" = "HP"; "HP" = "HP"
    "DELL INC." = "DELL"; "DELL" = "DELL"
    "LENOVO" = "LENOVO"; "ACER" = "ACER"
    "ASUSTEK COMPUTER INC." = "ASUS"
}

$vendorLimpio = if ($diccionarioOEM.ContainsKey($rawManufacturer.ToUpper())) {
    $diccionarioOEM[$rawManufacturer.ToUpper()]
} else {
    $rawManufacturer
}

$placaModelo = "$vendorLimpio $rawProduct".ToUpper()

# --- NUEVO: LÓGICA DE GAMA OEM v1.3.4 ---
$gamaPlaca = "Estándar" # Valor por defecto
$esOEM = $false

# Si el fabricante es un integrador conocido, marcamos como OEM
if ($diccionarioOEM.Values -contains $vendorLimpio.ToUpper()) {
    $gamaPlaca = "Placa OEM ($vendorLimpio)"
    $esOEM = $true
}

# --- A. CLASIFICACIÓN POR CHIPSET (Solo si no es OEM o para precisar gama) ---
if ($placaModelo -match "(Z|X|TRX)\d{2,3}") {
    $gamaPlaca = "Alta"
    if ($placaModelo -match "Z(690|790)|X670") { $pcieGen = "PCIe 5.0" }
    elseif ($placaModelo -match "Z(390|490|590)|X570") { $pcieGen = "PCIe 3.0/4.0" }
}
elseif ($placaModelo -match "B\d{2,3}|H[67]\d{2}") {
    $gamaPlaca = "Media"
    if ($placaModelo -match "B(660|760)") { $pcieGen = "PCIe 4.0/5.0" }
    elseif ($placaModelo -match "B(250|360|450|460|560|550)") { $pcieGen = "PCIe 3.0/4.0" }
}
elseif ($placaModelo -match "(H|A)\d{2,3}") {
    $gamaPlaca = "Entrada/Oficina"
    if ($placaModelo -match "H(610|710)|A520") { $pcieGen = "PCIe 3.0/4.0" }
}

# --- B. VALIDACIÓN DE MODERNIDAD (Mantenemos tu lógica original) ---
$seriesIntelVigentes = "6|7"
$seriesAMDVigentes   = "5|6|7"
$regexIntel = "(B|H|Z)($seriesIntelVigentes)\d{2}"
$regexAMD   = "(A|B|X)($seriesAMDVigentes)\d{2}"
$esPlacaModerna = ($placaModelo -match $regexIntel) -or ($placaModelo -match $regexAMD)

# --- C. DETECCIÓN PCIe PNP (Blindaje para placas OEM sin chipset en nombre) ---
if ($pcieGen -eq "PCIe (Clásico/Estándar)" -or $esOEM) {
    try {
        $checkPnp = Get-CimInstance Win32_PnPEntity -ErrorAction Stop | Where-Object { $_.Caption -match "PCI Express" } | Select-Object -First 1
        if ($checkPnp.Caption -match "Gen 4|4.0") { $pcieGen = "PCIe 4.0" }
        elseif ($checkPnp.Caption -match "Gen 5|5.0") { $pcieGen = "PCIe 5.0" }
        elseif ($checkPnp.Caption -match "Gen 3|3.0") { $pcieGen = "PCIe 3.0" }
    } catch {
        $pcieGen = "PCIe 3.0" # Fallback conservador
    }
}

# --- D. VALIDACIÓN DE BUS COMBINADA (REVISADA) ---
$cpuSoportaGen4 = ($cpuModelo -match "Intel" -and $genNum -ge 11) -or
                  ($cpuModelo -match "Ryzen" -and $genNum -ge 3 -and $cpuModelo -notmatch "G$|4100|5500")

if ($pcieGen -match "4.0|5.0" -and -not $cpuSoportaGen4) {
    # Asignamos el texto directamente. PowerShell lo tratará como String.
    $pcieGenEfectivo = "3.0 (Limitado por Procesador)"
} else {
    try {
        # Intentamos extraer el número. Si falla, cae al catch.
        if ($pcieGen -match "(\d\.\d)") {
            $pcieGenEfectivo = [double]$Matches[1]
        } else {
            $pcieGenEfectivo = 3.0
        }
    } catch {
        # Si algo explota en la conversión, nos aseguramos de que la variable no sea nula
        $pcieGenEfectivo = 3.0
    }
}

# 5. INFORMACIÓN DE GRÁFICOS (v1.3.4 BLINDADA)
$gpus = Get-CimInstance Win32_VideoController
$gpuInfo = foreach($g in $gpus) {
    try {
        # 1. Obtención de VRAM con corrección de desbordamiento (Overflow)
        # Convertimos a Decimal ([decimal]) para evitar límites de Int32/Int64 de WMI
        $vramRaw = 0
        if ($g.AdapterRAM) {
            $vramRaw = [decimal]$g.AdapterRAM
            if ($vramRaw -lt 0) {
                # Si es negativo, le sumamos 4GB (2^32) que es el desfase común de WMI
                $vramRaw += 4294967296
            }
        }

        $vramGB = [math]::Round($vramRaw / 1GB, 0)
        $nombre = $g.Name

        # 2. Lógica de clasificación mejorada
        $esDedicada = ($vramGB -ge 1 -and $nombre -notmatch "Graphics|Integrated|Basic|UHD|Iris|Vega|Radeon\(TM\) Graphics")

        $gamaGpu = "Integrada"
        if ($esDedicada) {
            # Subimos el estándar para 2026
            if ($vramGB -ge 8 -and ($nombre -match "RTX [34]|RX [67]")) { $gamaGpu = "Alta" }
            elseif ($vramGB -ge 6) { $gamaGpu = "Media" }
            else { $gamaGpu = "Entrada (Gaming Clásico)" }
        }

        [PSCustomObject]@{
            Nombre    = $nombre
            VRAM_Raw  = $vramRaw
            # Si tiene menos de 1GB real (como las integradas), mostramos "Dinámica"
            VRAM      = if($vramGB -lt 1){ "Dinámica" } else { "${vramGB}GB" }
            Tipo      = if($esDedicada){ "Dedicada" } else { "Integrada" }
            Gama      = $gamaGpu
            EsGaming  = ($esDedicada -and $vramGB -ge 4)
        }
    } catch {
        [PSCustomObject]@{
            Nombre = $g.Name; VRAM = "Dinámica"; Tipo = "Integrada"; Gama = "Básica"; EsGaming = $false
        }
    }
}

# 6. ALMACENAMIENTO Y RENDIMIENTO DE DISCO
# Mapeo de letras de unidad a números físicos de disco
$mapaDiscos = @{}
try {
    $particiones = Get-Partition -ErrorAction Stop | Where-Object { $_.DriveLetter }
    foreach ($p in $particiones) {
        $num = $p.DiskNumber.ToString()
        if (-not $mapaDiscos.ContainsKey($num) -or $p.DriveLetter -eq 'C') {
            $mapaDiscos[$num] = $p.DriveLetter
        }
    }
} catch {
    $mapaDiscos = @{}
}

# --- AJUSTE DE ACTIVIDAD DE DISCO (MUESTREO DINÁMICO) ---
$diskIOStr = try {
    $rutaUniversal = "\234(*)\1150" # Contador de Disk Transfers/sec
    $diskCounters = Get-Counter -Counter $rutaUniversal -SampleInterval 1 -MaxSamples 2 -ErrorAction Stop
    $muestrasAgrupadas = $diskCounters.CounterSamples | Group-Object InstanceName

    $detalles = foreach ($grupo in $muestrasAgrupadas) {
        if ($grupo.Name -notlike "*total*") {
            $promedio = [int]($grupo.Group | Measure-Object -Property CookedValue -Average).Average
            if ($promedio -gt 0) {
                $numDisco = ($grupo.Name -split " ")[0].Trim()
                $letra = if($mapaDiscos.ContainsKey($numDisco)){ " ($($mapaDiscos[$numDisco]):)" } else { "" }
                "Disco $($numDisco)$($letra): $promedio ops/s"
            }
        }
    }
    if ($detalles) { $detalles -join "/" } else { "Inactivo" }
} catch {
    # Plan B: Detección estática mediante WMI si los contadores de rendimiento fallan
    try {
        $wmiDisks = Get-CimInstance Win32_PerfRawData_PerfDisk_PhysicalDisk -ErrorAction Stop | Where-Object { $_.Name -ne "_Total" }
        $detallesWMI = foreach ($d in $wmiDisks) {
            if ($d.DiskTransfersPerSec -gt 0) {
                $numWMI = if ($d.Name -match "\d+") { $Matches[0] } else { $d.Name.Trim() }
                $letraWMI = if($mapaDiscos.ContainsKey($numWMI)){ " ($($mapaDiscos[$numWMI]):)" } else { "" }
                "Disco $($numWMI)$($letraWMI): Activo"
            }
        }
        if ($detallesWMI) { $detallesWMI -join "/" } else { "Inactivo" }
    } catch { "Lectura no disp." }
}

# 7. SALUD DEL SISTEMA Y LICENCIAMIENTO
$os = Get-CimInstance Win32_OperatingSystem
$ramUsoPorc = [math]::Round(((($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize) * 100), 1)

# Predicción de falla S.M.A.R.T.
$smartStatusAll = Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue
$physicalDisks = Get-PhysicalDisk -ErrorAction SilentlyContinue
if ($null -eq $physicalDisks) {
    # Fallback: Intentar obtener al menos los discos básicos por CIM
    $physicalDisks = Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue
}
$unidadesCriticas = 0
$unidadesRiesgo = 0
$sistemaEnHDD = $false
$contieneHDD = $false

foreach ($disk in $physicalDisks) {
    # --- A. Lógica de Salud (SMART + HealthStatus) ---
    $smartFalla = $smartStatusAll | Where-Object { $_.InstanceName -match "PHYSICALDRIVE$($disk.DeviceNumber)" -and $_.PredictFailure -eq $true }
    if ($smartFalla -or $disk.HealthStatus -eq 'Unhealthy') { $unidadesCriticas++ }
    elseif ($disk.HealthStatus -eq 'AVISO') { $unidadesRiesgo++ }

    # --- B. Lógica de Tipo de Disco ---
    if ($disk.MediaType -eq 'HDD') {
        $contieneHDD = $true
        # Si el HDD que estamos recorriendo coincide con el disco del OS que detectamos en la Secc 1
        if ($disk.DeviceNumber -eq $numDiscoFisico) {
            $sistemaEnHDD = $true
        }
    }
}

# --- Veredicto de Salud General ---
$saludGral = if ($unidadesCriticas -gt 0) {
    "🚨 Crítico ($unidadesCriticas Unidad/es en Fallo Inminente)"
} elseif ($unidadesRiesgo -gt 0) {
    "⚠️ Riesgo ($unidadesRiesgo Unidad/es con Desgaste/Advertencia)"
} else {
    "✅ Saludable"
}

# Validación de activación de Windows
$licencia = try {
    Get-CimInstance SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL AND Name LIKE '%Windows%'" -ErrorAction Stop | Select-Object -First 1
} catch { $null }
$estadoAct = if($licencia -and $licencia.LicenseStatus -eq 1){ "✅ Activo" } else { "❌ No Activado" }

# Detección de Suite Office instalada
$office = try {
    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match "Microsoft (Office|365)" -and $_.DisplayName -notmatch "MUI|Language|Runtime" } |
    Select-Object -First 1
} catch { $null }
$estadoOffice = if($office){ $nombreLimpio = $office.DisplayName -replace " - [a-z]{2}-[a-z]{2}$", ""; " ✅ Detectado ($nombreLimpio)" } else { " ⚠️ No Detectado" }

# Forzar limpieza de memoria tras la recolección masiva
[System.GC]::Collect()

# [5. SEGURIDAD (AdwCleaner - Lectura de Resultados v1.3.3)]
$amenazasDetectadas = $false

if ($quiereSeguridad -eq $true) {
    $msgSeg = "⚠️ Error - No se pudo leer el reporte de seguridad"
    $rutasPosibles = @("C:\AdwCleaner\Logs", "C:\AdwCleaner", "$env:AppData\AdwCleaner\Logs")

    for ($i=0; $i -lt 12; $i++) {
        if ([console]::KeyAvailable) {
            if ([console]::ReadKey($true).Key -eq "Escape") { Exit-SmartWare }
        }

        $logFile = $null
        foreach ($ruta in $rutasPosibles) {
            if (Test-Path $ruta -ErrorAction SilentlyContinue) {
                $logFile = Get-ChildItem -Path $ruta -Filter "AdwCleaner*.txt" -ErrorAction SilentlyContinue |
                           Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($logFile) { break }
            }
        }

        if ($logFile) {
            $lecturaExitosa = $false
            try {
                $stream = New-Object System.IO.FileStream($logFile.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $reader = New-Object System.IO.StreamReader($stream)
                $content = $reader.ReadToEnd()
                $lecturaExitosa = $true
            } catch {
                Write-Log -Mensaje "Error de lectura física en Log: $($_.Exception.Message)"
            } finally {
                if ($null -ne $reader) { $reader.Close(); $reader.Dispose() }
                if ($null -ne $stream) { $stream.Close(); $stream.Dispose() }
            }

            # Procesamos el contenido solo si la lectura funcionó
            if ($lecturaExitosa -and $content -match "(?i)(?:Detected|Detectadas|Encontradas|Results|#)\s*:\s*(?<count>\d+)") {
                if ($Matches.ContainsKey('count')) {
                    $count = [int]$Matches['count']
                    if ($count -gt 0) {
                        $msgSeg = "⚠️ Alerta - Se detectaron $count amenaza(s)"
                        $amenazasDetectadas = $true
                    } else {
                        $msgSeg = "🛡️ Protegido - No se detectaron amenazas"
                    }
                    break # Salimos del for, ya tenemos el resultado
                }
            }
        }
        Start-Sleep -Seconds 2
    }
} else {
    $msgSeg = "Análisis no realizado (Omitido por el usuario)"
}

[System.GC]::Collect()

# [6.1 PREPARACIÓN DE DATOS PARA VEREDICTO]

# --- CÁLCULO DE TAMAÑO DE ARCHIVOS TEMPORALES ---
$tempPaths = @("$env:TEMP", "$env:SystemRoot\Temp")
$tempSize = 0
foreach ($p in $tempPaths) {
    if (Test-Path $p) {
        # Sumamos el tamaño de todos los archivos en las rutas temporales
        $tempSize += (Get-ChildItem $p -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    }
}
$tempMB = [math]::Round($tempSize / 1MB, 0)

# --- CÁLCULO DE UPTIME (TIEMPO ENCENDIDO) ---
$osInfo = Get-CimInstance Win32_OperatingSystem
$lastBoot = $osInfo.LastBootUpTime
$uptime = (Get-Date) - $lastBoot
$uptimeStr = "$($uptime.Days)d $($uptime.Hours)h $($uptime.Minutes)m"

# --- DETECCIÓN DE SALUD Y UNIDADES (DISEÑO FICHA TÉCNICA) ---
$infoSaludDiscos = Get-PhysicalDisk | ForEach-Object {
    $stats = $_ | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
    $tipoSimple = if($_.MediaType -eq "SSD" -or $_.Model -match "SSD") { "SSD" } else { "HDD" }

    # A. Identificación de Marca Refinada
    $rawModel = $_.FriendlyName.Trim().ToUpper()
    $marca = "Genérico"

    # Diccionario de patrones conocidos
    if ($rawModel -match "CT\d{3}") { $marca = "CRUCIAL" }
    elseif ($rawModel -match "WDC|WD") { $marca = "WESTERN DIGITAL" }
    elseif ($rawModel -match "ST\d{4}") { $marca = "SEAGATE" }
    elseif ($rawModel -match "MZ-") { $marca = "SAMSUNG" }
    elseif ($rawModel -match "SA\d{3}|SUV") { $marca = "KINGSTON" }
    else {
        $marcasConocidas = @('CRUCIAL','TOSHIBA','SAMSUNG','KINGSTON', 'WESTERN DIGITAL', 'HIKSEMI','SEAGATE','LEXAR','ADATA','PNY','SANDISK')
        foreach ($m in $marcasConocidas) { if ($rawModel -match $m) { $marca = $m; break } }
    }

    $nombreFinal = if ($marca -ne "Genérico") { $marca } else {
        $limpio = ($rawModel -replace '^(ATA|SCSI|NVMe|SATA)\s+', '').Trim()
        if ($limpio.Length -gt 15) { $limpio.Substring(0,12) + "..." } else { $limpio }
    }

    # B. Veredicto de Salud traducido
    $saludVeredicto = switch ($_.HealthStatus) {
        'Healthy'   { "Óptimo" }
        'Warning'   { "Precaución" }
        'Unhealthy' { "Crítico" }
        default     { "Indeterminado" }
    }

    # C. Lógica de Integridad de Datos
    $horasRaw = if ($stats.UpTime -gt 0) { [math]::Round($stats.UpTime / 3600, 0) } else { 0 }
    $datosOK = $true

    # Si es SSD y no reporta desgaste o reporta 0 tras muchas horas, la telemetría podría estar bloqueada
    if ($tipoSimple -eq "SSD" -and ($null -eq $stats.Wear -or ($stats.Wear -eq 0 -and $horasRaw -gt 100))) { $datosOK = $false }
    if ($horasRaw -eq 0) { $datosOK = $false }

    # D. Construcción del Bloque Visual
    $capacidad = if($_.Size){ [math]::Round($_.Size/1GB,0) } else { 0 }
    $cabecera = " UNIDAD: $nombreFinal ($tipoSimple $($capacidad)GB)"

    # Generamos la línea decorativa basada en el largo del texto
    $largoContorno = [math]::Max(5, ($cabecera.Length + 1))
    $lineaSólida = " " + ("─" * $largoContorno)

    $bloque = "$cabecera`n$lineaSólida`n"
    $bloque += "  > Estado Físico : $saludVeredicto`n"

    if (-not $datosOK) {
        $bloque += "  > Telemetría    : Requiere software especializado`n"
    } else {
        $usoVida = if ($tipoSimple -eq "SSD") { "$($stats.Wear)%" } else { "N/A (Mecánico)" }
        $bloque += "  > Desgaste      : $usoVida`n"
        $bloque += "  > Tiempo Uso    : $($horasRaw) hrs`n"
    }

    return $bloque
}

# [6.2 LÓGICA DE VEREDICTO INTEGRAL - REESTRUCTURADA v1.3.3]
$n = [System.Environment]::NewLine
$propuestas = @()

# --- 1. CÁLCULOS BASE Y AUDITORÍA DE HARDWARE ---
$cpuInfo = Get-CimInstance Win32_Processor
$nucleos = if($cpuInfo.NumberOfCores){ $cpuInfo.NumberOfCores } else { 1 }
$hilos = if($cpuInfo.NumberOfLogicalProcessors){ $cpuInfo.NumberOfLogicalProcessors } else { 1 }

$physicalDisks = @(Get-PhysicalDisk)

# 1. Detección robusta de SSD (Busca la palabra SSD o si es NVMe)
$tieneSSD = $false
foreach ($d in $physicalDisks) {
    if ($d.MediaType -eq "SSD" -or $d.BusType -eq "NVMe") {
        $tieneSSD = $true
        break
    }
}

# 2. El "Plan B": Si el fabricante no etiquetó el MediaType,
# pero la velocidad de rotación es 0, ES UN SSD.
if (-not $tieneSSD) {
    if ($physicalDisks.SpindleSpeed -eq 0 -and $physicalDisks.Count -gt 0) {
        $tieneSSD = $true
    }
}

$esNVMe = ($physicalDisks.BusType -contains "NVMe")
$capacidadTotal = [math]::Round(($physicalDisks | Measure-Object -Property Size -Sum).Sum / 1GB, 0)

# Conteo preciso de Slots de Memoria (Ajuste: Validación de Colección)
try {
    $memArray = Get-CimInstance Win32_MemoryArray -ErrorAction SilentlyContinue
    $slotsTotales = if ($memArray) { ($memArray | Measure-Object -Property MemoryDevices -Sum).Sum } else { 0 }

    $modulosDetectados = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
    $conteoModulos = if ($null -ne $modulosDetectados) { @($modulosDetectados).Count } else { 0 }

    if (!$slotsTotales -or $slotsTotales -eq 0) { $slotsTotales = $conteoModulos }

    if ($slotsTotales -eq 0) {
        $tipoRamMsg = "Memoria integrada (No ampliable)"
    }
    elseif ($slotsTotales -gt $conteoModulos -and $conteoModulos -ge 1) {
        $tipoRamMsg = "Híbrida (Integrada + $($slotsTotales - $conteoModulos) Slot libre)"
    }
    else {
        $tipoRamMsg = "$slotsTotales Slots Físicos"
    }
} catch {
    $tipoRamMsg = "Requiere software especializado"
}
$formatSlots = $tipoRamMsg

# Auditoría de Salud de Batería (Solo Notebooks)
$wearLevelMsg = "⚠️ Batería No Detectada"
$desgaste = 0
$porcSalud = 0
$whCapacidadActual = 0

if ($esNotebook) {
    try {
        # NIVEL 1: Deep Scan (root/wmi)
        $batStatic = Get-CimInstance -Namespace "root\wmi" -ClassName BatteryStaticData -ErrorAction SilentlyContinue
        $batFull   = Get-CimInstance -Namespace "root\wmi" -ClassName BatteryFullChargedCapacity -ErrorAction SilentlyContinue

        if ($batStatic -and $batFull.FullChargedCapacity -gt 0) {
            $whDiseño = $batStatic.DesignedCapacity
            $whCapacidadActual = $batFull.FullChargedCapacity

            $porcSalud = [math]::Round(($whCapacidadActual / $whDiseño) * 100, 0)
            $desgaste = 100 - $porcSalud

            $wearLevelMsg = switch ($desgaste) {
                { $_ -gt 45 } { "🚨 Batería Crítica ($desgaste% pérdida)" }
                { $_ -gt 25 } { "⚠️ Batería Degradada ($desgaste% pérdida)" }
                { $_ -gt 10 } { "⚠️ Desgaste Moderado ($desgaste% pérdida)" }
                Default       { "✅ Saludable ($porcSalud% vida)" }
            }
        }
        else {
            # NIVEL 2: Fallback Estándar (Win32_Battery)
            $batWin32 = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
            if ($batWin32) {
                $statusBat = $batWin32.Status
                $cargaActual = $batWin32.EstimatedChargeRemaining
                $wearLevelMsg = if ($null -ne $cargaActual) { "ℹ️ Telemetría Limitada (Carga: $cargaActual% | Estado: $statusBat)" } else { "⚠️ Lectura Limitada por Hardware" }
            } else {
                $wearLevelMsg = "Requiere software especializado"
            }
        }
    } catch { $wearLevelMsg = "Error en lectura de energía" }
}

# --- 2. DETERMINACIÓN DEL VEREDICTO DE ARQUITECTURA ---
$capacidadSuficiente = ($capacidadTotal -ge 450)
$cpuPotente = ($esVigente -and ($cpuModelo -match "[79]-|Ryzen [79]|Ultra [79]"))
# Ajuste: Validación de variables de arquitectura
$esArquitecturaSolida = ($esPlacaModerna -or $esVigente -or $cpuPotente)
$bateriaOk = ($esNotebook -and $whCapacidadActual -gt 0 -and $desgaste -lt 45)
$sufAut = if($esNotebook -and -not $bateriaOk -and $wearLevelMsg -notmatch "Saludable"){ " [Autonomía limitada]" } else { "" }

$veredictoFinalStr = ""
if ($esArquitecturaSolida) {
    $esPotenciaIdeal = ($totalRAM_GB -ge 16 -and $esNVMe -and $capacidadSuficiente -and ($gamaCPU -match "Alta|Media"))
    $esPotenciaMinima = ($totalRAM_GB -ge 8 -and $esNVMe -and $capacidadSuficiente)

    if ($esPotenciaIdeal) {
        $veredictoFinalStr = "✅ ESTÁNDAR IDEAL $añoActual" + $(if($tipoEquipo -eq "Fijo"){" (Alto Rendimiento)"} else {" [Rendimiento Óptimo]"})
    }
    elseif ($esPotenciaMinima) {
        $veredictoFinalStr = "🔷 ESTÁNDAR BASE MODERNO (Potencial de mejora)" + $sufAut
    }
    else {
        $veredictoFinalStr = "🔷 ESTÁNDAR FUNCIONAL $añoActual" + $sufAut
    }
} else {
    if ($hilos -ge 8 -and $tieneSSD -and $totalRAM_GB -ge 8) {
        $veredictoFinalStr = "🔷 EQUIPO CLÁSICO VIGENTE (Rendimiento estable)"
    }
    elseif ($nucleos -ge 4 -and $tieneSSD -and $totalRAM_GB -ge 8) {
        $veredictoFinalStr = "🔷 EQUIPO BÁSICO (Uso administrativo/estudio)"
    }
    else {
        $veredictoFinalStr = "🚨 EQUIPO OBSOLETO (No se recomienda inversión)"
    }
}

# --- 3. INICIALIZACIÓN DE RESULTADOS Y GAMING ---
[System.Collections.Generic.List[string]]$vHW = @($veredictoFinalStr)
[System.Collections.Generic.List[string]]$vSW = @()

$gpuPrincipal = $gpuInfo | Where-Object { $_.Tipo -eq "Dedicada" } | Select-Object -First 1

if ($gpuPrincipal -and ($gpuPrincipal.VRAM_Raw -ge 2.5GB)) {
    $vramGB = [math]::Round($gpuPrincipal.VRAM_Raw / 1GB, 0)
    $rangoPotencia = if ($vramGB -ge 8 -and $gpuPrincipal.Nombre -match "RTX [345]|RX [678]") { "Alto" }
                     elseif ($vramGB -ge 6) { "Medio" }
                     else { "Básico" }

    $vHW.Add("          🎮 RENDIMIENTO GAMING ${añoActual}: $rangoPotencia")

    # Detección de Cuellos de Botella PCIe (Ajuste: Normalización de tipos)
    $ToDouble = {
    param($val)
    if ($val -match "(\d\.\d)") { [double]$Matches[1] } else { [double]($val -as [double]) }
    }

    $gpuGenNativa = if ($gpuPrincipal.Nombre -match "RTX 50|RX 8") { 5.0 }
                    elseif ($gpuPrincipal.Nombre -match "RTX [34]|RX [67]") { 4.0 }
                    else { 3.0 }

    # Usamos la extracción limpia para las comparaciones
    $n_pcieEfectivo = &$ToDouble $pcieGenEfectivo
    $n_pciePlaca     = &$ToDouble $pcieGen

    if ($n_pcieEfectivo -lt $gpuGenNativa) {
        $culpable = if ($n_pciePlaca -lt $gpuGenNativa) { "la PLACA BASE" } else { "el PROCESADOR" }
        $vHW.Add("             - ⚠️ CUELLO DE BOTELLA: $culpable limita el ancho de banda de la GPU")
    }

    # --- LÓGICA SSD NVMe ---
    $ssdNVMe = $physicalDisks | Where-Object { $_.BusType -eq "NVMe" } | Select-Object -First 1
    if ($ssdNVMe) {
        $ssdGenPide = if ($ssdNVMe.Model -match "Gen5|T700|T500") { 5.0 }
                      elseif ($ssdNVMe.Size -ge 450GB -or $ssdNVMe.Model -match "980|990|SN850|FireCuda 5") { 4.0 }
                      else { 3.0 }

        if ($n_pcieEfectivo -lt $ssdGenPide) {
            $culpableSSD = if ($n_pciePlaca -lt $ssdGenPide) { "la PLACA BASE" } else { "el PROCESADOR" }
            $vHW.Add("             - ⚠️ VELOCIDAD SSD RESTRINGIDA: $culpableSSD no soporta la velocidad nativa de la unidad")
        }
    }

    if ($totalRAM_GB -lt 16) { $vHW.Add("              - ⚠️ RAM LIMITADA: Posibles tirones (stuttering) en juegos modernos") }
    if ($infoSaludDiscos -notmatch "SSD") { $vHW.Add("              - ⚠️ HDD DETECTADO: Tiempos de carga elevados") }
}

# --- 4. LÓGICA DE MEJORAS SUGERIDAS ---
if ($veredictoFinalStr -match "🔷") {
    if ($totalRAM_GB -lt 16) {
        $modulosActuales = (Get-CimInstance Win32_PhysicalMemory).Count
        $slotsLibres = $slotsTotales - $modulosActuales # Ajuste: Uso de variable auditada
        $calidadInversion = if (-not $esArquitecturaSolida) { "segunda mano/económicos" } else { "de alto rendimiento" }

        if ($slotsLibres -gt 0) {
            $propuestas += "Expandir RAM: Añadir módulo para Dual Channel y totalizar 16GB"
        } else {
            $propuestas += "Reemplazar RAM: Sustituir módulos actuales por kits de 16GB
              ($calidadInversion)"
        }
    }

    if (-not $capacidadSuficiente -or -not $tieneSSD) {
        $msgSSD = if (-not $tieneSSD) { "MIGRACIÓN OBLIGATORIA: Instalar SSD" } else { "Ampliación de capacidad SSD" }
        if ($esArquitecturaSolida -and $esNVMe) {
            $propuestas += "$msgSSD a 1TB (M.2 NVMe Gen$pcieGenEfectivo)"
        } else {
            $propuestas += "$msgSSD a 480GB/960GB (SATA III)"
        }
    }

    if ($propuestas) {
        $vHW.Add("          ⚙️ MEJORAS RECOMENDADAS PARA OPTIMIZAR:")
        foreach ($p in $propuestas) { $vHW.Add("              - $p") }
    }
}

# --- 5. EVALUACIÓN DE ESTADO DEL SOFTWARE ---
$discoC = Get-Volume -DriveLetter C -ErrorAction SilentlyContinue
if ($discoC) {
    $librePorc = [math]::Round(($discoC.SizeRemaining / $discoC.Size) * 100, 1)
    if ($librePorc -lt 12) { $vSW.Add("ALMACENAMIENTO: Espacio crítico en C: ($librePorc% libre)") }
    elseif ($librePorc -lt 20) { $vSW.Add("ALMACENAMIENTO: Espacio en C: bajo ($librePorc% libre)") }
}

[System.GC]::Collect()

# [7.1 ANÁLISIS DE PROCESOS E IMPACTO DE SOFTWARE]

# 1. Clasificación de temporales
$msgTemp = if($tempMB -gt 1024){ "Limpieza Necesaria" } elseif($tempMB -gt 250){ "Limpieza Recomendada" } else { "Estado Óptimo" }

# 2. ANÁLISIS AVANZADO DE INICIO (v1.3.3 - Soporte UWP y Registro)
$wmiInicio = Get-CimInstance Win32_StartupCommand -ErrorAction SilentlyContinue
$listaLimpia = @()

# 1. Mapeo de estados en el Registro
$rutasRegistro = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32",
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run",
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32"
)

$tablaImpacto = @{}
foreach ($ruta in $rutasRegistro) {
    if (Test-Path $ruta) {
        $propiedades = Get-ItemProperty $ruta -ErrorAction SilentlyContinue
        if ($null -ne $propiedades) {
            foreach ($nombreProp in $propiedades.PSObject.Properties.Name) {
                if ($nombreProp -notmatch "PSParentPath|PSChildName|PSPath") {
                    $tablaImpacto[$nombreProp] = $propiedades.$nombreProp
                }
            }
        }
    }
}

# 2. Procesamiento de Apps Estándar
foreach($app in $wmiInicio){
    $valorRegistro = $tablaImpacto[$app.Name]

    # Lógica de filtrado (Bit impar = Deshabilitado)
    $estaHabilitado = $true
    if ($null -ne $valorRegistro -and $valorRegistro.Length -gt 0) {
        if ($valorRegistro[0] % 2 -ne 0) { $estaHabilitado = $false }
    }
    if (-not $estaHabilitado) { continue }

    # Limpieza de ruta para obtener descripción
    $rutaLimpiada = $app.Command -replace '"', '' -replace ' /.*', ''

    $nombre = switch -regex ($app.Name) {
        "MicrosoftEdgeAutoLaunch" { "Microsoft Edge" }
        "OneNote" { "Microsoft OneNote" }
        "SecurityHealth" { continue }
        Default {
            try {
                # 1. Expandimos variables y limpiamos espacios
                $rutaExpandida = [System.Environment]::ExpandEnvironmentVariables($rutaLimpiada).Trim()

                # 2. Validación lógica preventiva
                if (-not [string]::IsNullOrWhiteSpace($rutaExpandida)) {
                    # 3. Intento de acceso a la ruta
                    if (Test-Path -LiteralPath $rutaExpandida -ErrorAction Stop) {
                        $info = (Get-Item -LiteralPath $rutaExpandida -ErrorAction Stop).VersionInfo
                        if ($info.FileDescription) { $info.FileDescription } else { $app.Caption }
                    }
                    else { $app.Caption }
                }
                else { $app.Caption }
            }
            catch {
                # 4. Si TODO falla (permisos, archivos corruptos, etc.), rescatamos el nombre básico
                # Usamos el Caption o el Name original para que el reporte no quede vacío
                if ($app.Caption) { $app.Caption } else { $app.Name }
            }
        }
    }
    if (!$nombre) { $nombre = $app.Name }

    $impactoStr = ""
    if ($null -ne $valorRegistro -and $valorRegistro.Length -gt 0) {
        if ($valorRegistro[0] -ge 0x06) { $impactoStr = " [Impacto: Alto]" }
        elseif ($valorRegistro[0] -ge 0x02) { $impactoStr = " [Impacto: Medio]" }
    }

    $ruidoOcultable = "RTHDVCPL|OneDrive.*Setup|_{|Teams.*|Lync|SysTray"
    if ($nombre -notmatch $ruidoOcultable -and $app.Command -notmatch "windows\\system32") {
        $listaLimpia += "$nombre$impactoStr"
    }
}

# 3. --- REFUERZO PARA HYPERX ---
try {
    $hyperX = Get-StartApps | Where-Object { $_.Name -match "NGENUITY" }
    if ($hyperX) {
        $listaStr = $listaLimpia -join " "
        if ($listaStr -notmatch "NGENUITY") {
            $listaLimpia += "HyperX NGENUITY [Impacto: Medio]"
        }
    }
} catch {}

# 4. Consolidación Final
$inicioFinal = $listaLimpia | Where-Object { $_ -and $_.Trim() -ne "" } | Select-Object -Unique | Sort-Object
$numAppsInicio = $inicioFinal.Count

if ($numAppsInicio -gt 0) {
    $inicioAppsTexto = ($inicioFinal | ForEach-Object { "- $_" }) -join "`n"
} else {
    $inicioAppsTexto = "- ✅ El arranque del sistema se encuentra optimizado."
}

# --- DICCIONARIO DE TRADUCCIÓN DE PROCESOS (v1.3.3) ---
$traductorProcesos = @{
    "MsMpEng"               = "Antivirus de Windows"
    "SearchHost"            = "Buscador de Windows"
    "ShellExperienceHost"   = "Interfaz de Windows"
    "explorer"              = "Explorador de Archivos"
    "svchost"               = "Servicios del Sistema"
    "RuntimeBroker"         = "Gestor de Apps de MS Store"
    "Taskmgr"               = "Administrador de Tareas"
    "dwm"                   = "Gestor de Ventanas (Escritorio)"
    "csrss"                 = "Proceso Crítico del Sistema (CSRSS)"
    "SecurityHealthService" = "Servicio de Seguridad de Windows"
    "SecurityHealthHost"    = "Interfaz de Seguridad de Windows"
    "sppsvc"                = "Servicio de Protección de Software"
    "SystemSettings"        = "Configuración del Sistema"
    "Widgets"               = "Widgets de Windows"
    "Memory Compression"    = "Compresión de Memoria (Sistema)"
    "WmiPrvSE"              = "Gestión de Sistema (WMI)"
    "MSPCManagerService"    = "PC Manager"
    "System"                = "Sistema"
    "Photos"                = "Fotos"
    "brave"                 = "Brave Browser"
    "chrome"                = "Google Chrome"
    "msedge"                = "Microsoft Edge"
    "firefox"               = "Firefox"
    "Steam"                 = "Steam"
    "EpicGamesLauncher"     = "Epic Games"
    "EADM"                  = "EA App"
    "Origin"                = "EA Origin"
    "vgtray"                = "Riot Vanguard"
    "RiotClientServices"    = "Riot Client"
    "MobalyticsHQ.DesktopApp" = "Mobalytics Overlay"
    "LGHUB"                 = "Logitech G HUB"
    "Razer Central"         = "Razer Synapse"
    "HP.SystemEventUtility" = "Teclas de Función HP"
    "SteelSeriesGG"         = "SteelSeries Engine"
    "Corsair.Service"       = "Corsair iCUE"
    "Discord"               = "Discord"
    "WhatsApp"              = "WhatsApp"
    "Spotify"               = "Spotify"
    "AdobeCollabSync"       = "Sincronizador Adobe Acrobat"
    "AdobeUpdateService"    = "Actualizador de Adobe"
    "Lightshot"             = "Capturador Lightshot"
}

# 3. TOP 5 CPU - INTEGRADO CON DICCIONARIO
$top5Data = (Get-Process | Where-Object { $_.MainWindowTitle -ne "" -or ($_.CPU -and $_.CPU -gt 1) } | Group-Object Name | ForEach-Object {
    $procOriginal = $_.Group[0]
    $cargaNum = [math]::Round(($_.Group | Measure-Object CPU -Sum).Sum, 1)

    $nombreBase = $procOriginal.Name
    $nombre = if($traductorProcesos.ContainsKey($nombreBase)){
        $traductorProcesos[$nombreBase]
    } elseif($procOriginal.Description){
        $procOriginal.Description
    } else {
        $nombreBase
    }

    $palabras = $nombre.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($palabras.Count -gt 3) { $nombre = "$($palabras[0]) $($palabras[1]) $($palabras[2])..." }
    if ($nombre.Length -gt 30) { $nombre = $nombre.Substring(0,27) + "..." }

    [PSCustomObject]@{
        Nombre = $nombre
        Carga  = $cargaNum
        Texto  = "{0,-35} {1,15}" -f $nombre, "$cargaNum s"
    }
} | Sort-Object Carga -Descending | Select-Object -First 5)

# 3.1 TOP 5 RAM - INTEGRADO CON DICCIONARIO Y %
$top5RAMData = (Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 5 | ForEach-Object {
    $ws = if($_.WorkingSet64){ $_.WorkingSet64 } else { $_.WorkingSet }
    $memPorc = [math]::Round(($ws / ($totalRAM_GB * 1GB)) * 100, 1)

    $nombreBase = $_.Name
    $nombre = if($traductorProcesos.ContainsKey($nombreBase)){
        $traductorProcesos[$nombreBase]
    } elseif($_.Description){
        $_.Description
    } else {
        $nombreBase
    }

    $palabras = $nombre.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($palabras.Count -gt 3) { $nombre = "$($palabras[0]) $($palabras[1]) $($palabras[2])..." }
    if ($nombre.Length -gt 30) { $nombre = $nombre.Substring(0,27) + "..." }

    [PSCustomObject]@{
        Nombre = $nombre
        Uso    = $memPorc
        Texto  = "{0,-35} {1,15}" -f $nombre, "$memPorc %"
    }
})

[System.GC]::Collect()

# [7.1.4 DETECTOR DE CAUSA RAÍZ (EL CEREBRO)]

# 1. ESTADÍSTICAS OPERATIVAS
$muestrasProcesos = @()
$muestrasHilos = @()

# Tomamos 3 muestras con un intervalo de 1 segundo para promediar carga real
for ($i = 0; $i -lt 3; $i++) {
    $procesosActuales = Get-Process -ErrorAction SilentlyContinue
    if ($procesosActuales) {
        $muestrasProcesos += $procesosActuales.Count
        $muestrasHilos += ($procesosActuales | Select-Object -ExpandProperty Threads -ErrorAction SilentlyContinue).Count
    }
    if ($i -lt 2) { Start-Sleep -Seconds 1 }
}

# Calculamos el promedio para evitar picos falsos (Ajuste: Validación de colección)
$numProcesos = 0
$numHilosActual = 0
if ($muestrasProcesos.Count -gt 0) {
    $numProcesos = [int]($muestrasProcesos | Measure-Object -Average | Select-Object -ExpandProperty Average)
    $numHilosActual = [int]($muestrasHilos | Measure-Object -Average | Select-Object -ExpandProperty Average)
}

# Validación preventiva de temperatura
if ([string]::IsNullOrWhiteSpace($tempDisplay)) { $tempDisplay = "Requiere software especializado" }

# Inicialización de listas de hallazgos
$hallazgosSW = New-Object System.Collections.Generic.List[string]
$hallazgosHW = New-Object System.Collections.Generic.List[string]

# --- 2. VEREDICTO TÉCNICO DE RAM (BIOS vs Límite Físico) ---
$frecuenciasBase = @(2133, 2400, 2666, 2933)

if ($vActual -gt 0) {
    if ($vActual -lt $vMax) {
        if ($esArquitecturaSolida) {
            # Optimizable por BIOS (XMP/DOCP)
            $vHW += "          ⚠️ RENDIMIENTO OPTIMIZABLE: Perfil XMP/DOCP (${vMax}MHz)"
        } else {
            # Limitación física real (Chipset o CPU antigua)
            $hallazgosHW.Add("RAM a ${vActual}MHz (Límite físico del hardware)")
        }
    }
    elseif ($vActual -in $frecuenciasBase) {
        if ($esArquitecturaSolida) {
            $vHW += "          ℹ️ INFO: RAM en frecuencia base. Revisar BIOS por margen de optimización"
        } else {
            $hallazgosHW.Add("RAM en frecuencia base (${vActual}MHz). Limitado por Placa/Procesador")
        }
    }
}

# --- 3. SENSORES DE HARDWARE (LIMITANTES FÍSICAS) ---
if (-not $tieneSSD) { $hallazgosHW.Add("Ausencia de unidad de estado sólido (SSD)") }
if ($totalRAM_GB -lt 8) { $hallazgosHW.Add("RAM Insuficiente para multitarea moderna ($totalRAM_GB GB)") }
if ($esNotebook -and $desgaste -gt 35) { $hallazgosHW.Add("Degradación crítica de Batería ($desgaste%)") }
if ($ssdNVMe -and ($pcieGenEfectivo -lt $ssdGenPide)) {
    $culpableSSD = if ($pcieGen -lt $ssdGenPide) { "la PLACA BASE" } else { "el PROCESADOR" }
    $hallazgosHW.Add("VELOCIDAD SSD RESTRINGIDA: $culpableSSD limita el ancho de banda")
}
if ($cacheL3 -lt 8) {
    if ($cacheL3 -le 4) {
        $hallazgosHW.Add("VELOCIDAD PROCESAMIENTO: Crítica (Caché L3 de solo ${cacheL3}MB)")
    } else {
        $hallazgosHW.Add("VELOCIDAD PROCESAMIENTO: Limitada (Caché L3 baja: ${cacheL3}MB)")
    }
}

# --- 4. SENSORES DE SOFTWARE (CARGA LÓGICA) ---
if ($ramUsoPorc -ge 85)       { $hallazgosSW.Add("Saturación de RAM ($ramUsoPorc%)") }
if ($numHilosActual -ge 3500) { $hallazgosSW.Add("Exceso de Subprocesos ($numHilosActual)") }
if ($tempMB -gt 1024)         { $hallazgosSW.Add("Basura Digital ($([math]::Round($tempMB/1024, 1)) GB)") }
if ($inicioFinal.Count -gt 7) { $hallazgosSW.Add("Muchos programas de inicio ($($inicioFinal.Count))") }
if ($preCPU -gt 70)           { $hallazgosSW.Add("CPU Bajo Carga Crítica ($preCPU%)") }
if ($amenazasDetectadas)      { $hallazgosSW.Add("Amenazas de Seguridad detectadas") }

# --- 5. CONEXIÓN Y FORMATEO DEL REPORTE (SOFTWARE) ---
$vSW_Final = [System.Collections.Generic.List[string]]@()
if ($hallazgosSW.Count -gt 0 -or $vSW.Count -gt 0) {
    $todasLasAlertasSW = $vSW + $hallazgosSW | Select-Object -Unique
    $vSW_Final.Add("⚠️ OPTIMIZACIÓN Y ESTADO LÓGICO:")
    foreach ($alerta in $todasLasAlertasSW) {
        $vSW_Final.Add("              - $($alerta.Trim())")
    }
} else {
    $vSW_Final.Add("✅ SISTEMA OPERATIVO SALUDABLE")
}
$vSW = $vSW_Final

# --- 6. CONEXIÓN Y FORMATEO DEL REPORTE (HARDWARE) ---
if ($hallazgosHW.Count -gt 0) {
    # Limpiar avisos genéricos previos para dar prioridad a los hallazgos del cerebro
    $vHW = $vHW | Where-Object { $_ -notmatch "ℹ️ INFO:|⚠️ AVISO:" }
    $vHW += "          ℹ️ LIMITACIONES TÉCNICAS:"
    foreach ($h in $hallazgosHW) {
        $vHW += "              - $($h.Trim())"
    }
}

# --- 7. LÓGICA DE SMARTSCORE v1.3.3 (1-100) ---
$score = 100
$esWindows11 = $preOS.Version -ge "10.0.22000"

# 1. PENALIZACIÓN POR HARDWARE BASE
if ($sistemaEnHDD) {
    $score -= 40 # Penalización severa
}
elseif ($contieneHDD) {
    $score -= 5  # Penalización leve: HDD como almacén secundario
}

# 2. GESTIÓN DE RAM POR ESCENARIO
if ($totalRAM_GB -lt 8) {
    $score -= 25
}
elseif ($totalRAM_GB -eq 8) {
    $score -= if ($esWindows11) { 15 } else { 5 }
}
elseif ($totalRAM_GB -lt 16) {
    # Penaliza más si es un equipo Gaming (donde 16GB es el piso)
    if($gpuInfo.EsGaming -contains $true){ $score -= 15 } else { $score -= 5 }
}

# 2.5 PENALIZACIÓN POR CACHÉ L3 (CUELLO DE BOTELLA LOGICO)
if ($cacheL3 -lt 8) {
    # Penalización severa para estándares de 2026
    $score -= 15
}
elseif ($cacheL3 -ge 32) {
    # Bonificación por agilidad (solo si no ha sido penalizado por otras cosas críticas)
    if ($score -lt 100 -and $unidadesRiesgo -eq 0) { $score += 5 }
}

# --- AJUSTE DE TUERCAS v1.3.3: PENALIZACIÓN POR RAM SOLDADA ---
if ($tipoRamMsg -match "Integrada" -and $totalRAM_GB -le 8) {
    $score -= 10 # No ampliable
}
elseif ($tipoRamMsg -match "Híbrida" -and $totalRAM_GB -le 8) {
    $score -= 5  # Requiere módulo adicional para Dual Channel
}

# 3. SALUD, SEGURIDAD Y BATERÍA
if ($esNotebook -and $desgaste -gt 35) { $score -= 10 }
if ($amenazasDetectadas) { $score -= 25 }

# 4. CARGA DE TRABAJO
if ($preCPU -gt 85) { $score -= 15 }
if ($ramUsoPorc -gt 90) { $score -= 15 }

# Uso de umbral dinámico (si no existe, se asume 3500)
$uHilos = if ($null -ne $umbralHilosDinamico) { $umbralHilosDinamico } else { 3500 }
if ($numHilosActual -gt $uHilos) { $score -= 10 }

# 5. MANTENIMIENTO Y ESPACIO
if ($tempMB -gt 4096) { $score -= 5 }
if ($inicioFinal.Count -gt 10) { $score -= 10 }

if ($porcLibre -lt 20) {
    if ($espacioLibreGB -lt 20) { $score -= 20 } else { $score -= 10 }
}

# 6. --- EL CANDADO DE SALUD (OVERRIDE CRÍTICO) ---
if ($unidadesCriticas -gt 0) {
    $score = 20
}
elseif ($unidadesRiesgo -gt 0) {
    if ($score -gt 40) { $score = 40 }
}
elseif ($cacheL3 -lt 4 -and $score -gt 70) {
    $score = 70 # Techo para CPUs con caché extremadamente baja (evita falsos excelentes)
}
elseif ($sistemaEnHDD -and $score -gt 55) {
    $score = 55
}

# Ajuste Final de Score
$globalScore = [math]::Max(1, [math]::Min(100, $score))

# 7. CALIFICACIÓN CUALITATIVA
$calificacion = switch ($globalScore) {
    { $_ -ge 85 }               { "EXCELENTE" }
    { $_ -ge 65 -and $_ -lt 85 }{ "BUENO" }
    { $_ -ge 40 -and $_ -lt 65 }{ "REGULAR" }
    { $unidadesCriticas -gt 0 } { "CRÍTICO - FALLA DE DISCO" }
    Default                     { "CRÍTICO" }
}

# [7.1.5 LÓGICA DE VEREDICTO DINÁMICO v1.3.3 - REPARADO]

# 1. Definición de Escenarios de Plataforma
# Se basa en variables de arquitectura definidas en secciones previas (Protección contra nulos)
$plataformaModerna = ($esPlacaModerna -and ($esGeneracionModerna -or $esVigente))
$plataformaVigente = ($esArquitecturaSolida -or ($hilos -ge 8 -and $totalRAM_GB -ge 8))

# 2. Preparación de fragmentos de texto descriptivo
# Ajuste: Validación de existencia de hallazgos para evitar concatenaciones vacías
$textoSW = if ($null -ne $hallazgosSW -and $hallazgosSW.Count -gt 0) {
    "se detecta: " + ($hallazgosSW -join ', ')
} else { "" }

$textoHW = if ($null -ne $hallazgosHW -and $hallazgosHW.Count -gt 0) {
    ". Las limitaciones físicas detectadas se detallan más abajo."
} else { "." }

# 3. Construcción del mensaje según el escenario de Hardware/Software
if ($plataformaModerna) {
    $base = "Equipo moderno de alto rendimiento."
    $msg = if ($textoSW -ne "") {
        "$base Los componentes son excelentes, pero $textoSW. Una optimización del sistema permitirá aprovechar al máximo el potencial del equipo$textoHW"
    } else {
        "$base Los componentes superan los estándares de $añoActual y el sistema opera en condiciones óptimas$textoHW"
    }
}
elseif ($plataformaVigente) {
    $base = "Equipo clásico vigente para $añoActual."
    $msg = if ($textoSW -ne "") {
        "$base Los componentes aún son capaces, pero $textoSW. Una optimización profunda del sistema podría mejorar la fluidez$textoHW"
    } else {
        "$base Mantiene un rendimiento decente y el sistema opera en condiciones óptimas$textoHW"
    }
}
else {
    $base = "Equipo sin posibilidad de mejora rentable para los estándares de $añoActual."
    $msg = if ($textoSW -ne "") {
        "$base Componentes al límite debido a que $textoSW. Se sugiere realizar un respaldo de la información urgentemente y priorizar una renovación de equipo$textoHW"
    } else {
        "$base Sistema optimizado, pero los componentes presentan limitaciones físicas de mejora (ver detalle más abajo). Se sugiere realizar un respaldo de la información urgentemente y priorizar una renovación de equipo."
    }
}

# 4. ALGORITMO DE ENVOLTURA (Word Wrap a 82 caracteres)
# Garantiza que el párrafo no se corte abruptamente en la consola (Ajuste: Manejo de nulos en split)
$anchoMax = 82
$lineas = @()
$palabras = if ($null -ne $msg) { $msg -split "\s+" } else { @("") }
$lineaActual = ""

foreach ($palabra in $palabras) {
    if (($lineaActual + $palabra).Length -gt $anchoMax) {
        $lineas += $lineaActual.Trim()
        $lineaActual = $palabra + " "
    } else {
        $lineaActual += $palabra + " "
    }
}
# Añadir la última línea residual
if ($lineaActual.Trim() -ne "") { $lineas += $lineaActual.Trim() }

# Variable final que se imprimirá en el reporte
$analisisHumano = $lineas -join "`r`n"

# [7.2 CONSTRUCCIÓN DEL CUERPO DEL REPORTE v1.3.3]

# 0. Preparación de variables de soporte
$n = "`r`n" # Salto de línea estándar
$top5CPU = if($top5Data){ ($top5Data.Texto -join $n).Trim() } else { "No se detectaron procesos de alta carga." }
$top5RAM = if($top5RAMData){ ($top5RAMData.Texto -join $n).Trim() } else { "No se detectaron procesos de alta carga." }
$strHW = "HARDWARE: " + ($vHW -join $n)
$strSW = "SOFTWARE: " + ($vSW -join $n)

# Ensamblaje de particiones lógicas con filtrado de unidades pequeñas (pendrives o recovery)
$estadoDiscos = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
    if ($_.Size -gt 2GB) {
        $libreGB = [math]::Round($_.FreeSpace/1GB, 1)
        $totalGB = [math]::Round($_.Size/1GB, 1)
        $porcLibre = [math]::Round(($_.FreeSpace/$_.Size)*100, 1)
        " - Unidad $($_.DeviceID) >> Libre: ${libreGB}GB de ${totalGB}GB ($porcLibre%)"
    }
}

# --- 1. IDENTIFICACIÓN Y ESPECIFICACIONES ---
$cuerpo =  "FECHA: $((Get-Date).ToString('dd-MM-yyyy'))$n$n"

$cuerpo += "SISTEMA: $($preOS.Caption)$n"
$cuerpo += "PLACA BASE: $placaModelo$n"
$cuerpo += "SLOTS RAM: $formatSlots$n"
$cuerpo += "PROCESADOR: $cpuModelo ($nucleos Núcleos/$hilos Hilos)$n"

# Manejo de múltiples GPUs si existen (v1.3.3)
foreach($gpu in $gpuInfo) {
    $vramFinal = if ($gpu.VRAM -match "GB") { $gpu.VRAM } else { "$([math]::Round($gpu.VRAM_Raw/1GB, 1)) GB" }
    $cuerpo += "TARJETA GRÁFICA [$($gpu.Tipo)]: $($gpu.Nombre) ($vramFinal)$n"
}

$cuerpo += "RAM FÍSICA: $formatRAM$n$n"

$cuerpo += "TIPO DE EQUIPO: $tipoEquipo$n"
$cuerpo += "GAMA PLACA: $gamaPlaca$n"
$cuerpo += "GAMA/PERFIL PROC.: $gamaCPU/$perfilCPU$n"
$cuerpo += "GAMA GRÁFICA: $gamaGpu$n"
$cuerpo += "BUS PRINCIPAL (PCIe): Gen $pcieGenEfectivo$n$n"

# --- 2. RESUMEN DIAGNÓSTICO (Inyección del análisis dinámico 7.1.5) ---
$cuerpo += "RESUMEN DIAGNÓSTICO$n"
$cuerpo += "PUNTAJE DE SALUD: $globalScore/100 [$calificacion]$n"

# Inyectamos el análisis humano dinámico (limpieza de marcadores técnicos)
$resumenBreve = $analisisHumano -replace "REPORTE:", ""
$cuerpo += "$($resumenBreve.Trim())$n$n"

# Insights Rápidos (Banderas preventivas / Lógica de "Luces de Advertencia")
$cuerpo += "INSIGHTS OPERATIVOS:$n"
if ($ramUsoPorc -gt 85) { $cuerpo += "👉 MEMORIA: Tu PC está al límite de su capacidad de trabajo actual.$n" }
if ($esArquitecturaSolida -and ($vSW -match "⚠️|🚨")) { $cuerpo += "👉 RENDIMIENTO: El hardware es potente, pero el software lo está frenando.$n" }
if ($saludGral -match "Pred Fail|Risk|Riesgo|Falla|Malo") { $cuerpo += "🚨 ALERTA: Se detectaron anomalías físicas en el almacenamiento. RESPALDO URGENTE.$n" }
if ($amenazasDetectadas) { $cuerpo += "👉 SEGURIDAD: Se recomienda limpieza profunda y escaneo de virus.$n" }
if ($esNotebook -and $desgaste -gt 40) { $cuerpo += "👉 BATERÍA: La autonomía real es limitada por desgaste físico de celdas.$n" }
if ($cuerpo -notmatch "👉|🚨") { $cuerpo += "✅ Sin advertencias relevantes inmediatas.$n" }
$cuerpo += "$n"

# --- 3. EVIDENCIA TÉCNICA DETALLADA ---
$cuerpo += "CARGA Y SALUD DEL SISTEMA$n"
$cuerpo += "USO RAM: $ramUsoPorc% | CPU: $preCPU% | DISCOS: $diskIOStr$n"
$cuerpo += "CARGA LÓGICA: $numProcesos Procesos / $numHilosActual Subprocesos$n"
$cuerpo += "TEMPERATURA PROCESADOR: $tempDisplay$n"
$cuerpo += "TIEMPO ENCENDIDO: $uptimeStr$n"

if ($esNotebook) {
    $cuerpo += "ENERGÍA: $whCapacidadActual Wh (Original: $whDiseño Wh) | $wearLevelMsg$n"
}
$cuerpo += "SEGURIDAD: $msgSeg$n$n"

$cuerpo += "DISCOS Y ALMACENAMIENTO$n"
$cuerpo += "ESTADO FÍSICO (S.M.A.R.T.): $saludGral$n$n"

# Detalle dinámico de unidades físicas
$labelDisco = if ($infoSaludDiscos.Count -gt 1) { "DETALLE DE UNIDADES FÍSICAS:" } else { "DETALLE DE UNIDAD FÍSICA:" }
$cuerpo += "$labelDisco$n$n"
$cuerpo += ($infoSaludDiscos -join $n) + "$n"

# Sección de particiones (Letras de unidad)
$cuerpo += "ESTADO DE PARTICIONES (LÓGICAS):$n"
if ($estadoDiscos) {
    $cuerpo += ($estadoDiscos -join $n) + "$n$n"
} else {
    $cuerpo += " - ⚠️ No se detectaron volúmenes lógicos con letra asignada.$n$n"
}

$cuerpo += "LICENCIAS$n"
$cuerpo += "WINDOWS: $estadoAct$n"
$cuerpo += "OFFICE: $estadoOffice$n$n"

$cuerpo += "TOP 5 PROCESOS (Uso CPU Acumulado)$n"
$cuerpo += "TOP 5 PROCESOS (Uso RAM Actual)$n"

$cuerpo += "ARCHIVOS TEMPORALES$n"
$cuerpo += "TOTAL ENCONTRADO: $tempMB MB ($msgTemp)$n$n"

$cuerpo += "PROGRAMAS DE INICIO$n"
$cuerpo += "$inicioAppsTexto$n$n"

# --- 4. CIERRE ---
$cuerpo += "RESULTADO DEL DIAGNÓSTICO$n"

# [8.1 AJUSTE ESTÉTICO v1.3.3]

# Definición de separadores visuales
$sepDoble  = "════════════════════════════════════════════════════════════════════════════════════"
$sepSimple = "────────────────────────────────────────────────────────────────────────────────────"

# 1. Preparación del bloque de CIERRE
# Limpieza de saltos de línea redundantes (Ajuste: Preservar estructura de diagnóstico)
$estadoFinalLimpio = $strHW + "`r`n`r`n" + $strSW

# 2. Limpieza inicial del cuerpo (quitar rastros de separadores previos si existen)
$cuerpoFinal = $cuerpo -replace "={10,}", "" -replace "-{10,}", ""

# 3. Formateo de Secciones (Bucle de inyección de encabezados elegantes)
$secciones = @(
    "CARGA Y SALUD DEL SISTEMA",
    "DISCOS Y ALMACENAMIENTO",
    "LICENCIAS",
    "ARCHIVOS TEMPORALES",
    "PROGRAMAS DE INICIO",
    "RESULTADO DEL DIAGNÓSTICO"
)

foreach ($sec in $secciones) {
    if ($cuerpoFinal -match [regex]::Escape($sec)) {
        if ($sec -eq "RESULTADO DEL DIAGNÓSTICO") {
             # El gran final con todo el detalle técnico de hardware y software
             $cuerpoFinal = $cuerpoFinal -replace [regex]::Escape($sec), "`r`n$sepDoble`r`n$sec`r`n$sepDoble`r`n$estadoFinalLimpio"
        } else {
             $cuerpoFinal = $cuerpoFinal -replace [regex]::Escape($sec), "`r`n$sepDoble`r`n$sec`r`n$sepDoble"
        }
    }
}

# 4. Formateo MANUAL del Resumen Diagnóstico
$targetResumen = "RESUMEN DIAGNÓSTICO"
$nuevoBloqueResumen = "`r`n$sepDoble`r`n$targetResumen`r`n$sepDoble"
$cuerpoFinal = $cuerpoFinal.Replace($targetResumen, $nuevoBloqueResumen)

# 5. MANEJO QUIRÚRGICO TOP 5 CPU
$barraProceso = "────────────────────"
$barraCarga   = "───────────────"
$headerTexto = "{0,-35} {1,15}" -f "Proceso", "Carga Acum. (s)"
$espacioMedio = " " * 16
$headerCompleto = "$headerTexto`r`n$barraProceso$espacioMedio$barraCarga"

$target = "TOP 5 PROCESOS (Uso CPU Acumulado)"
$nuevoBloqueTop = "`r`n$sepDoble`r`nTOP 5 PROCESOS (Uso CPU Acumulado)`r`n$sepDoble`r`n$headerCompleto`r`n$($top5CPU.Trim())`r`n"

$cuerpoFinal = $cuerpoFinal.Replace($target, $nuevoBloqueTop)

# 6. MANEJO QUIRÚRGICO TOP 5 RAM
$headerTextoRAM  = "{0,-35} {1,15}" -f "Proceso", "Uso Memoria (%)"
$headerCompletoRAM = "$headerTextoRAM`r`n$barraProceso$espacioMedio$barraCarga"

$targetRAM = "TOP 5 PROCESOS (Uso RAM Actual)"
$nuevoBloqueRAM = "`r`n$sepDoble`r`nTOP 5 PROCESOS (Uso RAM Actual)`r`n$sepDoble`r`n$headerCompletoRAM`r`n$($top5RAM.Trim())`r`n"

$cuerpoFinal = $cuerpoFinal.Replace($targetRAM, $nuevoBloqueRAM)

# 7. GUÍA DE REFERENCIAS (Anexo educativo)
$c = [char]0x00A9 # Símbolo Copyright
$guiaReferencia = @"
$sepDoble
ANEXO: GUÍA DE REFERENCIA RÁPIDA (SmartWare $añoActual)
$sepDoble

1. ¿CUÁNTO CALOR ES NORMAL?:
   - Uso Normal: 35°C a 55°C (Estado óptimo).
   - Bajo Carga: 65°C a 85°C (Normal en tareas exigentes o juegos).
   - Alerta Crítica: Si supera los 95°C, el equipo requiere mantenimiento físico
     (limpieza y pasta térmica) de forma inmediata.

    En Notebooks es normal observar hasta 10°C adicionales debido al
    diseño compacto del sistema de ventilación y menor espacio.

2. ESTADO FÍSICO DE LOS DISCOS:
   - Diagnóstico (SMART): Si el reporte indica "Riesgo" o "Falla", la integridad
     de sus archivos está comprometida. ¡Respalde su información de inmediato!
   - SSD vs HDD: Los SSD se desgastan por uso; los HDD por tiempo, vibración o
     golpes.

3. REGLAS DE ORO DEL ALMACENAMIENTO:
   - Margen de >20GB: Nunca llene su disco al 100%. Mantener un espacio libre
     permite que el sistema procese archivos temporales sin ralentizarse.
   - Limpieza Segura: Borre descargas antiguas y temporales para ganar espacio.

 Nota: Si alguna lectura indica "Requiere software especializado", es debido a
 restricciones de acceso al sensor vía Windows. Se recomienda usar herramientas
 de lectura directa para un monitoreo detallado.

$sepSimple
$c $añoActual C. Díaz. Todos los derechos reservados.
Diagnóstico Inteligente SmartWare v$version.
"@

# --- LIMPIEZA FINAL DE ESPACIOS ---
$cuerpoFinal = $cuerpoFinal -replace "(`r?`n){4,}", "`r`n`r`n`r`n"
if ($cuerpoFinal) { $cuerpoFinal = $cuerpoFinal.Trim() }

# Unión del cuerpo con la Guía de Referencia
$cuerpoFinal += "`r`n`r`n`r`n$guiaReferencia"

# --- FINALIZACIÓN: OCULTAR LA CONSOLA ---
try {
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    if ($hwnd -and $hwnd -ne 0) {
        if ("Win32Functions.Win32ShowWindowAsync" -as [type]) {
            [Win32Functions.Win32ShowWindowAsync]::ShowWindowAsync($hwnd, 0) | Out-Null
        }
    }
} catch {
    Write-Log -Mensaje "No se pudo ocultar la ventana de la consola." -Gravedad "LOW"
}

# ----------------------------------------------------------------------
# DISEÑO XAML (Interfaz Visual de Usuario)
# ----------------------------------------------------------------------
# Se ha configurado el TextBox con IsReadOnly y Consolas para mantener la alineación técnica
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Diagnóstico Inteligente SmartWare" Height="730" Width="900" Background="#0A0A0A"
        WindowStartupLocation="CenterScreen" ResizeMode="CanMinimize">
    <Grid Margin="25">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="DIAGNÓSTICO TÉCNICO COMPLETO" Foreground="#00A2FF" FontSize="22" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,0,0,20"/>

        <Border Grid.Row="1" Background="#161618" CornerRadius="12" BorderBrush="#333333" BorderThickness="2">
            <TextBox x:Name="TxtCuerpo" Foreground="#F0F0F0" FontFamily="Consolas" FontSize="16" TextWrapping="Wrap" AcceptsReturn="True" TextAlignment="Left"
              IsReadOnly="True" Background="Transparent" BorderThickness="0" Padding="30,20,35,30" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" MaxWidth="840"/>
        </Border>

        <Button x:Name="BtnCerrar" Grid.Row="2" Content="GENERAR REPORTE Y CERRAR" Margin="0,25,0,0" Height="60" Background="#0078D4" Foreground="White" FontSize="16" FontWeight="Bold" BorderThickness="0">
            <Button.Resources><Style TargetType="Border"><Setter Property="CornerRadius" Value="10"/></Style></Button.Resources>
        </Button>

        <Grid Grid.Row="3" Margin="0,15,0,0">
            <StackPanel HorizontalAlignment="Left" VerticalAlignment="Center">
                <TextBlock Foreground="#AAAAAA" HorizontalAlignment="Left" FontSize="11" Text="¿Necesitas ayuda con estos resultados?"/>
                <TextBlock HorizontalAlignment="Left" Margin="0,2,0,0" FontSize="13" FontWeight="Bold">
                    <Hyperlink x:Name="LnkWhatsapp" NavigateUri="https://wa.link/iqojld" Foreground="#00BBFF" TextDecorations="None">
                        wa.link/iqojld
                    </Hyperlink>
                </TextBlock>
            </StackPanel>

            <StackPanel HorizontalAlignment="Right">
                <TextBlock Text="SmartWare" Foreground="#00A2FF" FontSize="14" FontWeight="Bold" HorizontalAlignment="Right"/>
                <TextBlock Text="v$version © Por C. Díaz" Foreground="#2A2A2A" FontSize="10" Margin="0,5,0,0" HorizontalAlignment="Right" Opacity="0.7"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@

# [8.2 LANZAMIENTO Y LÓGICA DE CIERRE]

try {
    # Inicialización del lector XAML para la interfaz gráfica
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]($xaml))
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # 1. Cargar el reporte generado en el cuadro de texto de la UI
    $txt = $window.FindName("TxtCuerpo")
    $txt.Text = if ($cuerpoFinal) { $cuerpoFinal.Trim() } else { "No se generó reporte." }

    # 2. Lógica para el Hyperlink de Soporte (WhatsApp)
    $lnk = $window.FindName("LnkWhatsapp")
    if ($lnk) {
        $lnk.Add_RequestNavigate({
            param($s, $e)
            $null = $s
            try {
                Start-Process -FilePath $e.Uri.AbsoluteUri
            } catch {
                # Fallback: Abrir vía explorer si el protocolo no está asociado
                Start-Process "explorer.exe" $e.Uri.AbsoluteUri
            }
            $e.Handled = $true
        })
    }

    # 3. Lógica del Botón Principal de Cierre y Guardado
    $btn = $window.FindName("BtnCerrar")
    $btn.Add_Click({
        # --- A. DETECCIÓN INTELIGENTE DE ESCRITORIO ---
        $rutaEscritorio = $null
        $regPaths = @("HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
        foreach ($reg in $regPaths) {
            $val = (Get-ItemProperty $reg -ErrorAction SilentlyContinue).Desktop
            if ($val) {
                $rutaEscritorio = [System.Environment]::ExpandEnvironmentVariables($val)
                break
            }
        }
        if (!$rutaEscritorio) { $rutaEscritorio = Join-Path $env:USERPROFILE "Desktop" }

        # --- B. GENERACIÓN DEL ARCHIVO DE REPORTE ---
        $firmaTxt = "`r`n" + ("=" * 82) + "`r`nReporte generado por SmartWare - Soporte: https://wa.link/iqojld`r`n" + ("=" * 82)
        $reporteParaArchivo = $cuerpoFinal + $firmaTxt
        $nombreArchivo = "Reporte Diagnóstico SmartWare $((Get-Date).ToString('dd-MM-yyyy_HHmm')).txt"
        $rutaReporte = Join-Path $rutaEscritorio $nombreArchivo

        try {
            # Guardado en UTF8 con BOM
            $reporteParaArchivo | Out-File -FilePath $rutaReporte -Encoding UTF8 -Force
            Start-Sleep -Milliseconds 300
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Aviso: No se pudo escribir el reporte en el Escritorio.", "SmartWare")
        }

        # --- C. CIERRE FORZADO DE HERRAMIENTAS EXTERNAS ---
        $procesosHerramientas = @("HWMonitor", "CrystalDiskInfo", "TreeSizeFree", "AdwCleaner", "AdwCleaner7")
        foreach ($proc in $procesosHerramientas) {
            Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Milliseconds 500

        # Borrado de logs de AdwCleaner
        if (Test-Path "C:\AdwCleaner") {
            Remove-Item -Path "C:\AdwCleaner" -Recurse -Force -ErrorAction SilentlyContinue
        }

        # --- D. PROTOCOLO DE AUTODESTRUCCIÓN ---
        $miPID = $PID
        $rutaExe = (Get-Process -Id $PID).MainModule.FileName
        $rutaCarpeta = (Split-Path -Parent $rutaExe).ToLower()
        $nombreCarpeta = Split-Path $rutaCarpeta -Leaf
        $rutaTempSmartWare = (Join-Path $env:TEMP "Herramientas SmartWare").ToLower()
        $archivoBat = Join-Path $env:TEMP "Limpieza_SmartWare.bat"

        # Script Batch optimizado para esperar al proceso padre y limpiar todo
        $contenidoBat = @"
@echo off
:esperarProceso
tasklist /FI "PID eq $miPID" 2>nul | find "$miPID" >nul
if %ERRORLEVEL%==0 (timeout /t 1 /nobreak >nul & goto esperarProceso)

cd /d c:\
timeout /t 2 /nobreak >nul

:intentarBorrado
if exist "$rutaTempSmartWare" rd /s /q "$rutaTempSmartWare" 2>nul
rd /s /q "$rutaCarpeta" 2>nul

if exist "$rutaCarpeta" (
    del /f /q /s /a "$rutaCarpeta\*" 2>nul
    timeout /t 2 /nobreak >nul
    goto intentarBorrado
)
del /f /q "%~f0"
"@

        # Validación de seguridad: Solo borra si la carpeta es específica de SmartWare
        $descargas = (Join-Path $env:USERPROFILE "Downloads").ToLower()
        $rutasProhibidas = @($env:USERPROFILE.ToLower(), $rutaEscritorio.ToLower(), $descargas, "c:\", "c:\windows")

        if ($nombreCarpeta -like "*SmartWare*" -and ($rutasProhibidas -notcontains $rutaCarpeta)) {
            [System.IO.File]::WriteAllLines($archivoBat, $contenidoBat, [System.Text.Encoding]::GetEncoding(850))
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$archivoBat`"" -WindowStyle Hidden
        }

        $window.Close()
    })

    # Renderizado y foco
    $window.Topmost = $true
    $window.Add_ContentRendered({
        $window.Activate()
        Start-Sleep -Milliseconds 500
        $window.Topmost = $false
    })

    $window.ShowDialog() | Out-Null

} catch {
    $mensajeUI = $_.Exception.Message
    Write-Log -Mensaje "Fallo crítico al cargar interfaz XAML: $mensajeUI" -Gravedad "ALTA"
    [System.Windows.Forms.MessageBox]::Show("Error en la interfaz: $mensajeUI", "SmartWare UI")
}

# ======================================================================
# BLOQUE FINALLY GLOBAL
# ======================================================================
} catch {
    $lineaError = $_.InvocationInfo.ScriptLineNumber
    $mensaje = $_.Exception.Message
    Write-Log -Mensaje "FALLO TOTAL DEL MOTOR (Línea $lineaError): $mensaje" -Gravedad "CRITICA"
    [System.Windows.Forms.MessageBox]::Show("Error Crítico de SmartWare: $mensaje", "SmartWare Engine")
} finally {
    # 1. Aseguramos que existan las rutas básicas incluso si el script falló al inicio
    if (!$rutaExe) { $rutaExe = (Get-Process -Id $pid).MainModule.FileName }
    if (!$rutaEscritorio) {
        $rutaEscritorio = [System.Environment]::GetFolderPath("Desktop")
    }

    Write-Log -Mensaje "Finalizando SmartWare. Sistema limpio."

    # --- Lógica de Exportación de Log Crítico ---
    $rutaLogOriginal = Join-Path (Split-Path $rutaExe -Parent) "SmartWare_Debug.log"

    if (Test-Path $rutaLogOriginal) {
        # Pequeña espera para asegurar que el buffer del archivo se cerró
        Start-Sleep -Milliseconds 200

        $contenidoLog = Get-Content $rutaLogOriginal -ErrorAction SilentlyContinue
        $tieneErrores = $contenidoLog | Where-Object { $_ -match "GRAVEDAD: (ALTA|CRITICA|AVISO|ERROR|BAJA)" }

        if ($null -ne $tieneErrores) {
            $rutaLogEscritorio = Join-Path $rutaEscritorio "SmartWare_Debug.log"
            $cabeceraLog = "### ATENCIÓN: Se han detectado incidencias técnicas durante el análisis. ###`r`n"
            $cabeceraLog + ($tieneErrores -join "`r`n") | Out-File -FilePath $rutaLogEscritorio -Encoding UTF8 -Force
        }
    }

    # Limpieza final
    $tempSW = Join-Path $env:TEMP "Herramientas SmartWare"
    if (Test-Path $tempSW) { Remove-Item -Path $tempSW -Recurse -Force -ErrorAction SilentlyContinue }
    [console]::CursorVisible = $true
}