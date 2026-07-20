[CmdletBinding()]
param(
    [string]$Repository = "https://github.com/REPLACE_WITH_YOUR_ACCOUNT/PriceInstall.git",
    [string]$InstallDir = "C:\PriceInstall",
    [string]$Branch = "main",
    [ValidateRange(1, 1439)]
    [int]$CheckIntervalMinutes = 5,
    [switch]$SkipConfigEdit,
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"
$TaskName = "HoroshopPrices"
$InstallLog = Join-Path $env:ProgramData "HoroshopPrices-install.log"
$TranscriptStarted = $false

[Net.ServicePointManager]::SecurityProtocol = `
    [Net.ServicePointManager]::SecurityProtocol -bor `
    [Net.SecurityProtocolType]::Tls12

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    if (!$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Run PowerShell as Administrator."
    }
}

function Update-CurrentPath {
    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machinePath;$userPath"
}

function Install-WingetPackage {
    param([string]$PackageId, [string]$DisplayName)
    Write-Output "Installing $DisplayName through winget..."
    winget install --id $PackageId --exact --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) { throw "winget could not install $DisplayName." }
}

function Assert-ValidSignature {
    param([string]$Path, [string]$DisplayName)
    if ((Get-AuthenticodeSignature -FilePath $Path).Status -ne "Valid") {
        throw "The installer signature for $DisplayName is invalid."
    }
}

function Install-PythonDirect {
    $version = "3.13.14"
    $installer = Join-Path $env:TEMP "python-$version-amd64.exe"
    Invoke-WebRequest -Uri "https://www.python.org/ftp/python/$version/python-$version-amd64.exe" -OutFile $installer -UseBasicParsing
    Assert-ValidSignature -Path $installer -DisplayName "Python"
    $process = Start-Process -FilePath $installer -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" -Wait -PassThru
    if ($process.ExitCode -ne 0) { throw "Python installer exited with code $($process.ExitCode)." }
}

function Install-GitDirect {
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/git-for-windows/git/releases/latest" -Headers @{ "User-Agent" = "HoroshopPrices-Installer" }
    $asset = $release.assets | Where-Object { $_.name -match "^Git-.+-64-bit\.exe$" } | Select-Object -First 1
    if ($null -eq $asset) { throw "Could not find a 64-bit Git for Windows installer." }
    $installer = Join-Path $env:TEMP $asset.name
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $installer -UseBasicParsing
    Assert-ValidSignature -Path $installer -DisplayName "Git"
    $process = Start-Process -FilePath $installer -ArgumentList "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /SP-" -Wait -PassThru
    if ($process.ExitCode -ne 0) { throw "Git installer exited with code $($process.ExitCode)." }
}

function Test-PythonInstalled {
    foreach ($command in @("py", "python")) {
        if (Get-Command $command -ErrorAction SilentlyContinue) {
            & $command --version 2>$null | Out-Null
            if ($LASTEXITCODE -eq 0) { return $true }
        }
    }
    return $false
}

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) { return @("py", "-3") }
    if (Get-Command python -ErrorAction SilentlyContinue) { return @("python") }
    throw "Python 3 was not found after installation."
}

function Invoke-TaskCommand {
    param([string[]]$Arguments)
    & schtasks.exe @Arguments
    if ($LASTEXITCODE -ne 0) { throw "schtasks failed: $($Arguments -join ' ')" }
}

function Stop-ExistingRuntime {
    Start-Process -FilePath "schtasks.exe" -ArgumentList @("/End", "/TN", $TaskName) -WindowStyle Hidden -Wait | Out-Null
    $paths = @(
        (Join-Path $InstallDir "price_server.py"),
        (Join-Path $InstallDir "scripts\supervisor.ps1"),
        (Join-Path $InstallDir "scripts\auto_update.bat")
    )
    Get-CimInstance Win32_Process | Where-Object {
        $commandLine = [string]$_.CommandLine
        $commandLine -and ($paths | Where-Object {
            $commandLine.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0
        })
    } | ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
        Write-Output "Stopped existing process: $($_.Name), PID $($_.ProcessId)."
    }
    Remove-Item -LiteralPath (Join-Path $InstallDir "logs\horoshop_prices.pid") -Force -ErrorAction SilentlyContinue
}

function Grant-ProjectAccess {
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    & icacls.exe $InstallDir /grant "*S-1-5-18:(OI)(CI)M" /T /C | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Could not grant SYSTEM access to $InstallDir." }
}

function Initialize-Configuration {
    $configPath = Join-Path $InstallDir "config.json"
    if (!(Test-Path $configPath)) {
        Copy-Item (Join-Path $InstallDir "config.example.json") $configPath
        Write-Host "Created local configuration: $configPath"
    }
    else {
        Write-Host "Using existing local configuration: $configPath"
    }

    if (!$SkipConfigEdit) {
        Write-Host "Set the real Horoshop domain in config.json, then close Notepad." -ForegroundColor Yellow
        Start-Process notepad.exe -ArgumentList "`"$configPath`"" -Wait
    }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "config.json is invalid: $($_.Exception.Message)"
    }
    if (![string]$config.horoshop.domain -or [string]$config.horoshop.domain -match "example\.com") {
        throw "Set the real horoshop.domain in $configPath before continuing."
    }
    return $config
}

function Get-PublicLogFile {
    param([object]$Config)

    $defaultLog = Join-Path $InstallDir "logs\horoshop_prices.log"
    $pathValue = [string]$Config.logging.public_log_path
    $name = [string]$Config.logging.public_log_name
    if ([string]::IsNullOrWhiteSpace($pathValue)) { return $defaultLog }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "horoshop_prices.log" }
    if ([System.IO.Path]::GetFileName($name) -ne $name) { return $defaultLog }
    $directory = if ([System.IO.Path]::IsPathRooted($pathValue)) {
        $pathValue
    }
    else {
        Join-Path $InstallDir $pathValue
    }
    return Join-Path $directory $name
}

function Initialize-PublicLog {
    param([string]$LogFile)

    $directory = Split-Path -Path $LogFile -Parent
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    & icacls.exe $directory /grant "*S-1-5-18:(OI)(CI)M" /T /C | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Could not grant SYSTEM write access to the public log directory: $directory"
    }
    New-Item -ItemType File -Path $LogFile -Force | Out-Null
}

function Start-AndVerifyService {
    param([string]$PublicLogFile)

    $pidFile = Join-Path $InstallDir "logs\horoshop_prices.pid"
    $supervisorLog = $PublicLogFile
    Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue
    Invoke-TaskCommand -Arguments @("/Run", "/TN", $TaskName)
    $deadline = (Get-Date).AddSeconds(90)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        if (!(Test-Path $pidFile)) { continue }
        $savedPid = 0
        if (![int]::TryParse((Get-Content $pidFile -Raw).Trim(), [ref]$savedPid)) { continue }
        $process = Get-CimInstance Win32_Process -Filter "ProcessId = $savedPid" -ErrorAction SilentlyContinue
        if ($null -ne $process -and [string]$process.CommandLine -like "*price_server.py*") {
            Write-Output "Web server verified. PID: $savedPid."
            return
        }
    }
    $fallbackSupervisorLog = Join-Path $InstallDir "logs\supervisor.log"
    $details = if (Test-Path $supervisorLog) {
        (Get-Content $supervisorLog -Tail 20) -join "`n"
    }
    elseif (Test-Path $fallbackSupervisorLog) {
        (Get-Content $fallbackSupervisorLog -Tail 20) -join "`n"
    }
    else {
        "Supervisor log is empty."
    }
    throw "The scheduled task did not keep the web server running. See $supervisorLog`n$details"
}

try {
    try {
        Start-Transcript -Path $InstallLog -Append -Force | Out-Null
        $TranscriptStarted = $true
    }
    catch {
        Write-Warning "Could not start installer transcript: $($_.Exception.Message)"
    }
    Assert-Administrator
    if (!(Get-Command git -ErrorAction SilentlyContinue)) {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Install-WingetPackage -PackageId "Git.Git" -DisplayName "Git"
        }
        else {
            Install-GitDirect
        }
    }
    if (!(Test-PythonInstalled)) {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Install-WingetPackage -PackageId "Python.Python.3.13" -DisplayName "Python 3.13"
        }
        else {
            Install-PythonDirect
        }
    }
    Update-CurrentPath

    $safeDirectory = $InstallDir.Replace("\", "/")
    $knownSafeDirectories = @(& git config --system --get-all safe.directory 2>$null)
    if ($knownSafeDirectories -notcontains $safeDirectory) {
        & git config --system --add safe.directory $safeDirectory
        if ($LASTEXITCODE -ne 0) { throw "Could not mark $InstallDir as a safe Git directory." }
    }

    if (Test-Path (Join-Path $InstallDir ".git")) {
        $configBackup = Join-Path $env:TEMP "HoroshopPrices-config-$PID.json"
        if (Test-Path (Join-Path $InstallDir "config.json")) {
            Copy-Item (Join-Path $InstallDir "config.json") $configBackup -Force
        }
        Stop-ExistingRuntime
        git -C $InstallDir fetch origin $Branch
        if ($LASTEXITCODE -ne 0) { throw "Could not fetch origin/$Branch." }
        git -C $InstallDir reset --hard "origin/$Branch"
        if ($LASTEXITCODE -ne 0) { throw "Could not reset the server copy to origin/$Branch." }
        git -C $InstallDir clean -fd
        if ($LASTEXITCODE -ne 0) { throw "Could not clean untracked deployment files." }
        if (Test-Path $configBackup) {
            Copy-Item $configBackup (Join-Path $InstallDir "config.json") -Force
            Remove-Item $configBackup -Force -ErrorAction SilentlyContinue
            Write-Output "Restored local config.json after code update."
        }
    }
    elseif (Test-Path $InstallDir) {
        throw "$InstallDir exists but is not a Git repository."
    }
    else {
        git clone --branch $Branch $Repository $InstallDir
        if ($LASTEXITCODE -ne 0) { throw "Could not clone the repository." }
    }

    Grant-ProjectAccess

    $config = Initialize-Configuration
    $publicLogFile = Get-PublicLogFile -Config $config
    Initialize-PublicLog -LogFile $publicLogFile

    Set-Location $InstallDir
    $venvDir = Join-Path $InstallDir ".venv"
    $venvPython = Join-Path $InstallDir ".venv\Scripts\python.exe"
    if (!(Test-Path $venvPython)) {
        $python = Get-PythonCommand
        if ($python.Count -eq 2) {
            & $python[0] $python[1] -m venv $venvDir
        }
        else {
            & $python[0] -m venv $venvDir
        }
        if ($LASTEXITCODE -ne 0) { throw "Could not create virtual environment." }
    }
    & $venvPython -m pip install --upgrade pip
    & $venvPython -m pip install -r (Join-Path $InstallDir "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "Could not install Python dependencies." }
    & $venvPython -m py_compile (Join-Path $InstallDir "horoshop_prices.py") (Join-Path $InstallDir "price_server.py")
    if ($LASTEXITCODE -ne 0) { throw "Python syntax check failed." }

    $port = if ($config.server.port) { [int]$config.server.port } else { 8095 }
    $ruleName = "HoroshopPrices-$port"
    Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue | Remove-NetFirewallRule
    New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port | Out-Null

    $launcher = Join-Path $InstallDir "scripts\auto_update.bat"
    $taskAction = "`"$launcher`" `"$Branch`" `"$CheckIntervalMinutes`""
    Invoke-TaskCommand -Arguments @("/Create", "/TN", $TaskName, "/SC", "ONSTART", "/TR", $taskAction, "/RU", "SYSTEM", "/RL", "HIGHEST", "/F")
    $task = Get-ScheduledTask -TaskName $TaskName
    $task.Settings.ExecutionTimeLimit = "PT0S"
    Set-ScheduledTask -InputObject $task | Out-Null
    if ((Get-ScheduledTask -TaskName $TaskName).Settings.ExecutionTimeLimit -ne "PT0S") {
        throw "Could not disable the scheduled task execution limit."
    }
    Start-AndVerifyService -PublicLogFile $publicLogFile

    Write-Host "Installation completed. Web panel: http://<server>:$port" -ForegroundColor Green
    Write-Host "Public service log: $publicLogFile" -ForegroundColor Green
    Write-Host "Installer log: $InstallLog" -ForegroundColor Green
}
catch {
    Write-Host "Installation failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Installer log: $InstallLog" -ForegroundColor Yellow
    exit 1
}
finally {
    if ($TranscriptStarted) { Stop-Transcript | Out-Null }
    if (!$NoPause) { Read-Host "Press Enter to close" }
}
