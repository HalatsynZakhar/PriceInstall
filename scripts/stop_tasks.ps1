[CmdletBinding()]
param(
    [switch]$KeepTask,
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"
$TaskName = "HoroshopPrices"
$AppDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ServerScript = Join-Path $AppDir "price_server.py"
$SupervisorScript = Join-Path $AppDir "scripts\supervisor.ps1"
$AutoUpdateBat = Join-Path $AppDir "scripts\auto_update.bat"
$PidFile = Join-Path $AppDir "logs\horoshop_prices.pid"
$stopped = 0

function Invoke-Schtasks {
    param([string[]]$Arguments)

    $process = Start-Process -FilePath "schtasks.exe" -ArgumentList $Arguments -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -notin @(0, 1)) {
        throw "schtasks failed with code $($process.ExitCode): $($Arguments -join ' ')"
    }
}

function Stop-ProjectProcesses {
    $paths = @($ServerScript, $SupervisorScript, $AutoUpdateBat)
    $processes = Get-CimInstance Win32_Process | Where-Object {
        $commandLine = [string]$_.CommandLine
        $commandLine -and ($paths | Where-Object {
            $commandLine.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0
        })
    }
    foreach ($process in $processes) {
        Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
        Write-Output "Stopped $($process.Name), PID $($process.ProcessId)."
        $script:stopped++
    }
}

try {
    Write-Output "Stopping scheduled task $TaskName..."
    Invoke-Schtasks -Arguments @("/End", "/TN", $TaskName)
    Stop-ProjectProcesses
    if (!$KeepTask) {
        Write-Output "Removing scheduled task $TaskName..."
        Invoke-Schtasks -Arguments @("/Delete", "/TN", $TaskName, "/F")
    }
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Stop-ProjectProcesses
    Write-Host "Horoshop Prices stopped. Processes stopped: $stopped." -ForegroundColor Green
}
catch {
    Write-Host "Stop failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    if (!$NoPause) { Read-Host "Press Enter to close" }
}
