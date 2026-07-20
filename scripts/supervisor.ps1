[CmdletBinding()]
param(
    [string]$Branch = "main",
    [ValidateRange(1, 1439)]
    [int]$CheckIntervalMinutes = 5
)

$ErrorActionPreference = "Stop"
$AppDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$VenvDir = Join-Path $AppDir ".venv"
$PythonExe = Join-Path $AppDir ".venv\Scripts\python.exe"
$ServerScript = Join-Path $AppDir "price_server.py"
$LogsDir = Join-Path $AppDir "logs"
$ConfigFile = Join-Path $AppDir "config.json"
$PidFile = Join-Path $LogsDir "horoshop_prices.pid"
$FallbackSupervisorLog = Join-Path $LogsDir "supervisor.log"
$WorkerOutputLog = Join-Path $LogsDir "server-output.log"
$WorkerErrorLog = Join-Path $LogsDir "server-error.log"
$Worker = $null
$NextUpdateCheck = [DateTime]::MinValue
$GitExe = $null

$machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
$env:Path = "$machinePath;$userPath"

function Get-PublicLogFile {
    $defaultLog = Join-Path $LogsDir "horoshop_prices.log"
    if (!(Test-Path $ConfigFile)) { return $defaultLog }
    try {
        $config = Get-Content -LiteralPath $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $pathValue = [string]$config.logging.public_log_path
        $name = [string]$config.logging.public_log_name
        if ([string]::IsNullOrWhiteSpace($pathValue)) { return $defaultLog }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = "horoshop_prices.log" }
        if ([System.IO.Path]::GetFileName($name) -ne $name) { return $defaultLog }
        $directory = if ([System.IO.Path]::IsPathRooted($pathValue)) {
            $pathValue
        }
        else {
            Join-Path $AppDir $pathValue
        }
        return Join-Path $directory $name
    }
    catch {
        return $defaultLog
    }
}

$SupervisorLog = Get-PublicLogFile

function Write-Log {
    param([string]$Message)

    $Line = "[{0:yyyy-MM-ddTHH:mm:ssK}] {1}" -f (Get-Date), $Message
    $activeLog = $SupervisorLog
    try {
        New-Item -ItemType Directory -Path (Split-Path $activeLog -Parent) -Force | Out-Null
        Add-Content -LiteralPath $activeLog -Value $Line -Encoding UTF8
    }
    catch {
        $activeLog = $FallbackSupervisorLog
        New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
        Add-Content -LiteralPath $activeLog -Value $Line -Encoding UTF8
    }
    Write-Output $Line

    if ((Test-Path $activeLog) -and (Get-Item $activeLog).Length -gt 2MB) {
        Get-Content -LiteralPath $activeLog -Tail 1000 |
            Set-Content -LiteralPath $activeLog -Encoding UTF8
    }
}

function Find-GitExecutable {
    $command = Get-Command git.exe -ErrorAction SilentlyContinue
    if ($command -and (Test-Path $command.Source)) {
        return $command.Source
    }

    $candidates = @(
        (Join-Path $env:ProgramFiles "Git\cmd\git.exe"),
        (Join-Path $env:ProgramFiles "Git\bin\git.exe"),
        "C:\Program Files\Git\cmd\git.exe",
        "C:\Program Files\Git\bin\git.exe",
        "C:\Program Files (x86)\Git\cmd\git.exe"
    ) | Select-Object -Unique
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }
    return $null
}

function Test-WorkerRunning {
    return $null -ne $script:Worker -and !$script:Worker.HasExited
}

function Ensure-PythonEnvironment {
    if (Test-Path $PythonExe) { return }

    Write-Log "Virtual environment is missing. Creating $VenvDir."
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -m venv $VenvDir
    }
    elseif (Get-Command python -ErrorAction SilentlyContinue) {
        & python -m venv $VenvDir
    }
    else {
        throw "Python 3 was not found. Install Python and rerun FULL_SERVER_INSTALL.ps1."
    }
    if ($LASTEXITCODE -ne 0 -or !(Test-Path $PythonExe)) {
        throw "Could not create virtual environment: $VenvDir"
    }

    & $PythonExe -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) {
        throw "Could not upgrade pip in $VenvDir"
    }
    & $PythonExe -m pip install -r (Join-Path $AppDir "requirements.txt")
    if ($LASTEXITCODE -ne 0) {
        throw "Could not install Python dependencies into $VenvDir"
    }
    Write-Log "Virtual environment and dependencies were created."
}

function Test-UpdatedProject {
    param([bool]$RequirementsChanged)

    Ensure-PythonEnvironment
    if ($RequirementsChanged) {
        Write-Log "requirements.txt changed. Updating packages in the existing virtual environment."
        & $PythonExe -m pip install --disable-pip-version-check -r (Join-Path $AppDir "requirements.txt")
        if ($LASTEXITCODE -ne 0) {
            throw "Could not update Python dependencies after updating the project."
        }
    }
    & $PythonExe -m py_compile $ServerScript (Join-Path $AppDir "horoshop_prices.py")
    if ($LASTEXITCODE -ne 0) {
        throw "Python syntax check failed after updating the project."
    }
    Write-Log "Updated application files were verified successfully."
}

function Backup-LocalData {
    $backupDir = Join-Path ([System.IO.Path]::GetTempPath()) ("HoroshopPrices-update-" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

    foreach ($relativePath in @("config.json")) {
        $source = Join-Path $AppDir $relativePath
        if (Test-Path $source) {
            $target = Join-Path $backupDir $relativePath
            New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
            Copy-Item -LiteralPath $source -Destination $target -Force
        }
    }
    return $backupDir
}

function Restore-LocalData {
    param([string]$BackupDir)

    foreach ($relativePath in @("config.json")) {
        $source = Join-Path $BackupDir $relativePath
        if (Test-Path $source) {
            $target = Join-Path $AppDir $relativePath
            New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
            Copy-Item -LiteralPath $source -Destination $target -Force
        }
    }
}

function Restore-WorkerFromPidFile {
    if (!(Test-Path $PidFile)) { return }

    $savedPid = 0
    if (![int]::TryParse((Get-Content $PidFile -Raw).Trim(), [ref]$savedPid)) {
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        return
    }
    $processInfo = Get-CimInstance Win32_Process -Filter "ProcessId = $savedPid" -ErrorAction SilentlyContinue
    if ($null -eq $processInfo -or [string]$processInfo.CommandLine -notlike "*$ServerScript*") {
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        return
    }
    $script:Worker = Get-Process -Id $savedPid -ErrorAction SilentlyContinue
    if (Test-WorkerRunning) { Write-Log "Found running web server. PID: $savedPid" }
}

function Start-Worker {
    if (Test-WorkerRunning) { return }
    Ensure-PythonEnvironment
    if (!(Test-Path $ServerScript)) { throw "Server file was not found: $ServerScript" }

    New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
    $script:Worker = Start-Process `
        -FilePath $PythonExe `
        -ArgumentList "`"$ServerScript`"" `
        -WorkingDirectory $AppDir `
        -WindowStyle Hidden `
        -RedirectStandardOutput $WorkerOutputLog `
        -RedirectStandardError $WorkerErrorLog `
        -PassThru
    Set-Content -LiteralPath $PidFile -Value $script:Worker.Id -Encoding ASCII
    Start-Sleep -Seconds 2
    $script:Worker.Refresh()
    if ($script:Worker.HasExited) {
        $exitCode = $script:Worker.ExitCode
        $script:Worker = $null
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        $details = if (Test-Path $WorkerErrorLog) {
            (Get-Content -LiteralPath $WorkerErrorLog -Tail 10) -join " | "
        } else { "No details in server-error.log." }
        throw "Web server stopped immediately (code $exitCode): $details"
    }
    Write-Log "Web server started. PID: $($script:Worker.Id)"
}

function Stop-Worker {
    if (!(Test-WorkerRunning)) { return }
    $workerId = $script:Worker.Id
    Stop-Process -Id $workerId -Force -ErrorAction SilentlyContinue
    $script:Worker.WaitForExit()
    $script:Worker = $null
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
    Write-Log "Web server stopped. PID: $workerId"
}

function Update-ProjectIfNeeded {
    if ((Get-Date) -lt $script:NextUpdateCheck) { return $false }
    $script:NextUpdateCheck = (Get-Date).AddMinutes($CheckIntervalMinutes)

    & $GitExe -C $AppDir fetch origin $Branch --quiet
    if ($LASTEXITCODE -ne 0) { throw "Could not fetch origin/$Branch." }

    $local = (& $GitExe -C $AppDir rev-parse HEAD).Trim()
    $remote = (& $GitExe -C $AppDir rev-parse "origin/$Branch").Trim()
    if ($local -eq $remote) { return $false }

    Write-Log "Update found. Stopping service and rebuilding deployment from origin/$Branch."
    $backupDir = $null
    $requirementsChanged = $false
    try {
        & $GitExe -C $AppDir diff --quiet $local "origin/$Branch" -- requirements.txt
        $requirementsChanged = $LASTEXITCODE -ne 0
        Stop-Worker
        $backupDir = Backup-LocalData
        & $GitExe -C $AppDir reset --hard "origin/$Branch"
        if ($LASTEXITCODE -ne 0) { throw "Git hard reset failed." }
        & $GitExe -C $AppDir clean -fd
        if ($LASTEXITCODE -ne 0) { throw "Could not clean untracked deployment files." }
        Restore-LocalData $backupDir
        Test-UpdatedProject $requirementsChanged
        Write-Log "Update completed. The web server will be started from the new revision."
        return $true
    }
    catch {
        $failure = $_.Exception.Message
        Write-Log "Update failed: $failure. Restoring revision $local."
        try {
            & $GitExe -C $AppDir reset --hard $local
            if ($LASTEXITCODE -ne 0) { throw "Git rollback failed." }
            & $GitExe -C $AppDir clean -fd
            if ($LASTEXITCODE -ne 0) { throw "Could not clean files during rollback." }
            if ($backupDir) { Restore-LocalData $backupDir }
            Ensure-PythonEnvironment
            & $PythonExe -m py_compile $ServerScript (Join-Path $AppDir "horoshop_prices.py")
            if ($LASTEXITCODE -ne 0) { throw "Python syntax check failed after rollback." }
            Write-Log "Previous revision was restored. The web server will be started again."
        }
        catch {
            throw "Update failed: $failure Rollback also failed: $($_.Exception.Message)"
        }
        throw "Update failed: $failure Previous revision was restored."
    }
    finally {
        if ($backupDir -and (Test-Path $backupDir)) {
            Remove-Item -LiteralPath $backupDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

$GitExe = Find-GitExecutable
if (!$GitExe) {
    Write-Log "Supervisor startup error: Git was not found for the SYSTEM account. Install Git for all users or add it to the system PATH."
    throw "Git was not found for the SYSTEM account."
}
if (!(Test-Path (Join-Path $AppDir ".git"))) { throw "$AppDir is not a Git repository." }

try {
    Restore-WorkerFromPidFile
    Write-Log "Supervisor started. Branch: $Branch; Git check every $CheckIntervalMinutes min."
    while ($true) {
        try {
            if (Update-ProjectIfNeeded) {
                Write-Log "Updated deployment is active. The supervisor stays running."
            }
            Start-Worker
        }
        catch {
            Write-Log "Supervisor error: $($_.Exception.Message)"
            try {
                if (!(Test-WorkerRunning)) {
                    Write-Log "Attempting to start the last verified web server after an error."
                    Start-Worker
                }
            }
            catch {
                Write-Log "Recovery start failed: $($_.Exception.Message)"
            }
        }
        Start-Sleep -Seconds 10
    }
}
finally {
    Stop-Worker
}
