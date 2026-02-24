#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Full setup for the temperature-scraper stack:
      1. Downloads and installs OpenHardwareMonitor as a Windows service (via NSSM)
      2. Creates a Python virtual environment and installs pip requirements
      3. Registers the Prometheus exporter as a Windows Scheduled Task

.PARAMETER Uninstall
    Stops and removes both the OpenHardwareMonitor service and the scraper
    scheduled task, and cleans up downloaded files.

.PARAMETER SkipOHM
    Skip the OpenHardwareMonitor installation step (e.g. it is already installed).

.PARAMETER Start
    Start the scraper scheduled task immediately after setup.

.EXAMPLE
    .\setup.ps1                 # Full setup
    .\setup.ps1 -Start          # Full setup + start scraper immediately
    .\setup.ps1 -SkipOHM        # Skip OHM install, set up scraper only
    .\setup.ps1 -Uninstall      # Remove everything
#>

param(
    [switch]$Uninstall,
    [switch]$SkipOHM,
    [switch]$Start
)

$ErrorActionPreference = "Stop"

# --- Paths ---
$ScriptDir        = $PSScriptRoot
$VenvPath         = Join-Path $ScriptDir ".venv"
$PythonExe        = Join-Path $VenvPath "Scripts\python.exe"
$ScraperScript    = Join-Path $ScriptDir "scraper.py"
$RequirementsPath = Join-Path $ScriptDir "requirements.txt"

$OhmDir           = "C:\Program Files\OpenHardwareMonitor"
$OhmSetupExe      = "$env:TEMP\OpenHardwareMonitorSetup.exe"
$OhmExe           = "$OhmDir\OpenHardwareMonitor.exe"
$OhmSettingsPath  = "$OhmDir\OpenHardwareMonitor.settings"
$NssmZip          = "$env:TEMP\nssm.zip"
$NssmDir          = "$env:TEMP\nssm"
$NssmExe          = "$NssmDir\nssm-2.24\win64\nssm.exe"

$OhmServiceName   = "OpenHardwareMonitor"
$ScraperTaskName  = "TemperatureScraper"

# ============================================================================
# Uninstall
# ============================================================================
if ($Uninstall) {
    Write-Host "`nUninstalling temperature-scraper stack..." -ForegroundColor Yellow

    # Remove scraper scheduled task
    $task = Get-ScheduledTask -TaskName $ScraperTaskName -ErrorAction SilentlyContinue
    if ($task) {
        if ($task.State -eq "Running") { Stop-ScheduledTask -TaskName $ScraperTaskName }
        Unregister-ScheduledTask -TaskName $ScraperTaskName -Confirm:$false
        Write-Host "Removed scheduled task: $ScraperTaskName" -ForegroundColor Green
    } else {
        Write-Host "Scheduled task not found - skipping." -ForegroundColor Yellow
    }

    # Remove OHM service
    $svc = Get-Service -Name $OhmServiceName -ErrorAction SilentlyContinue
    if ($svc) {
        if ($svc.Status -ne "Stopped") { Stop-Service -Name $OhmServiceName -Force }
        sc.exe delete $OhmServiceName | Out-Null
        Write-Host "Removed service: $OhmServiceName" -ForegroundColor Green
    } else {
        Write-Host "Service not found - skipping." -ForegroundColor Yellow
    }

    # Remove OHM install directory
    if (Test-Path $OhmDir) {
        Remove-Item $OhmDir -Recurse -Force
        Write-Host "Removed install directory: $OhmDir" -ForegroundColor Green
    }

    # Clean up temp files
    foreach ($f in @($OhmSetupExe, $NssmZip)) {
        if (Test-Path $f) { Remove-Item $f -Force }
    }
    if (Test-Path $NssmDir) { Remove-Item $NssmDir -Recurse -Force }

    Write-Host "`nUninstall complete." -ForegroundColor Green
    exit 0
}

# ============================================================================
# Step 1 - OpenHardwareMonitor
# ============================================================================
if ($SkipOHM) {
    Write-Host "`nStep 1/3 - Skipping OpenHardwareMonitor installation (-SkipOHM)." -ForegroundColor Yellow
} else {
    Write-Host "`nStep 1/3 - Checking OpenHardwareMonitor..." -ForegroundColor Cyan

    $ohmSvc = Get-Service -Name $OhmServiceName -ErrorAction SilentlyContinue

    # -- Resolve latest release version and download URL from GitHub --
    Write-Host "Resolving latest OHM release from GitHub (hexagon-oss fork)..."
    $latestVersion   = "1.0.3.0"  # fallback
    $ohmInstallerUrl = "https://github.com/hexagon-oss/openhardwaremonitor/releases/download/v1.0.3.0/OpenHardwareMonitorSetup.exe"
    try {
        $release = Invoke-RestMethod -Uri "https://api.github.com/repos/hexagon-oss/openhardwaremonitor/releases/latest" -UseBasicParsing
        $latestVersion = $release.tag_name.TrimStart('v')
        $asset = $release.assets | Where-Object { $_.name -like '*.exe' } | Select-Object -First 1
        if ($asset) { $ohmInstallerUrl = $asset.browser_download_url }
        Write-Host "Latest release: v$latestVersion" -ForegroundColor Cyan
    } catch {
        Write-Host "Could not query GitHub API - using fallback v$latestVersion." -ForegroundColor Yellow
    }

    # -- Check which version is currently installed (via registry) --
    $installedVersion = $null
    $ohmRegEntry = Get-ChildItem `
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
        -ErrorAction SilentlyContinue |
        Get-ItemProperty -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like '*OpenHardwareMonitor*' } |
        Select-Object -First 1
    if ($ohmRegEntry) { $installedVersion = $ohmRegEntry.DisplayVersion }

    $versionOk = $installedVersion -and ($installedVersion -eq $latestVersion)

    if ($ohmSvc -and (Test-Path $OhmExe) -and $versionOk) {
        Write-Host "OpenHardwareMonitor v$installedVersion (hexagon-oss) is already installed and up to date." -ForegroundColor Green
        if ($ohmSvc.Status -ne "Running") {
            Write-Host "OHM service is not running - starting it..." -ForegroundColor Yellow
            Start-Service -Name $OhmServiceName
        }
    } else {
        if ($installedVersion -and -not $versionOk) {
            Write-Host "Version mismatch: installed=v$installedVersion, latest=v$latestVersion - reinstalling..." -ForegroundColor Yellow
        } elseif (-not $installedVersion -and $ohmSvc) {
            Write-Host "OHM service exists but is not from the hexagon-oss fork - reinstalling..." -ForegroundColor Yellow
        } else {
            Write-Host "Installing OpenHardwareMonitor v$latestVersion (hexagon-oss fork) as a service..." -ForegroundColor Cyan
        }
        $nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"

        # -- .NET 8 Desktop Runtime (required by OHM v1.x) --
        Write-Host "Checking .NET 8 Desktop Runtime..."
        $dotnetRuntimes = & dotnet --list-runtimes 2>$null
        $hasDotnet8 = $dotnetRuntimes | Where-Object { $_ -match 'Microsoft\.WindowsDesktop\.App 8\.' }
        if (-not $hasDotnet8) {
            Write-Host ".NET 8 Desktop Runtime not found - installing via winget..." -ForegroundColor Yellow
            $winget = Get-Command winget -ErrorAction SilentlyContinue
            if ($winget) {
                & winget install Microsoft.DotNet.DesktopRuntime.8 --silent --accept-package-agreements --accept-source-agreements
            } else {
                Write-Warning "winget not available. Please install .NET 8 Desktop Runtime manually from https://aka.ms/dotnet/8/desktop-runtime-win-x64.exe then re-run this script."
                exit 1
            }
        } else {
            Write-Host ".NET 8 Desktop Runtime is already installed." -ForegroundColor Green
        }

        Write-Host "Downloading OpenHardwareMonitor v$latestVersion installer..."
        Invoke-WebRequest -Uri $ohmInstallerUrl -OutFile $OhmSetupExe

        # Stop and remove existing OHM service before reinstalling
        if ($ohmSvc) {
            if ($ohmSvc.Status -ne "Stopped") { Stop-Service -Name $OhmServiceName -Force }
            sc.exe delete $OhmServiceName | Out-Null
            Write-Host "Removed existing OHM service." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
        }

        Write-Host "Running OpenHardwareMonitor installer silently..."
        Start-Process -FilePath $OhmSetupExe -ArgumentList '/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES' -Wait

        # Locate the installed exe (try registry first, then default path)
        if (-not (Test-Path $OhmExe)) {
            $regKey = Get-ChildItem `
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
                -ErrorAction SilentlyContinue |
                Get-ItemProperty -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like '*OpenHardwareMonitor*' } |
                Select-Object -First 1
            if ($regKey -and $regKey.InstallLocation) {
                $OhmExe = Join-Path $regKey.InstallLocation 'OpenHardwareMonitor.exe'
                $OhmDir = $regKey.InstallLocation.TrimEnd('\')
            }
        }
        if (-not (Test-Path $OhmExe)) {
            Write-Error "OpenHardwareMonitor.exe not found after installation. Expected: $OhmExe"
            exit 1
        }
        Write-Host "OHM installed at: $OhmDir" -ForegroundColor Green

        # -- Download NSSM (only if not already present) --
        if (Test-Path $NssmExe) {
            Write-Host "NSSM already present - skipping download." -ForegroundColor Yellow
        } else {
            Write-Host "Downloading NSSM..."
            Invoke-WebRequest -Uri $nssmUrl -OutFile $NssmZip
            if (Test-Path $NssmDir) { Remove-Item $NssmDir -Recurse -Force }
            Expand-Archive -Path $NssmZip -DestinationPath $NssmDir
        }

        # -- Register OHM as a Windows service via NSSM --
        Write-Host "Installing the OHM Windows service..."
        & $NssmExe install $OhmServiceName $OhmExe
        & $NssmExe set     $OhmServiceName AppDirectory $OhmDir
        & $NssmExe set     $OhmServiceName Start SERVICE_AUTO_START
        & $NssmExe set     $OhmServiceName AppNoConsole 1

        Write-Host "Starting the OHM service..."
        & $NssmExe start $OhmServiceName
        Start-Sleep -Seconds 3
        $ohmStatus = (Get-Service -Name $OhmServiceName -ErrorAction SilentlyContinue).Status
        if ($ohmStatus -ne "Running") {
            Write-Warning "OHM service status is '$ohmStatus' - it may still be starting. Check with: Get-Service $OhmServiceName"
        } else {
            Write-Host "OpenHardwareMonitor service installed and started." -ForegroundColor Green
        }
    }

    # -- Always ensure OHM settings are correct (web server, port, minimised) --
    $ohmSettings = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <add key="runWebServer" value="true" />
    <add key="allowWebServerRemoteAccess" value="true" />
    <add key="HttpServerPort" value="8086" />
    <add key="startMinMenuItem" value="true" />
    <add key="minTrayMenuItem" value="true" />
  </appSettings>
</configuration>
"@
    $currentSettings = if (Test-Path $OhmSettingsPath) { [System.IO.File]::ReadAllText($OhmSettingsPath).Trim() } else { '' }
    if ($currentSettings -ne $ohmSettings.Trim()) {
        Write-Host "Updating OHM settings at: $OhmSettingsPath" -ForegroundColor Cyan
        [System.IO.File]::WriteAllText($OhmSettingsPath, $ohmSettings, [System.Text.UTF8Encoding]::new($false))
        # Restart the service so the new settings take effect
        $liveSvc = Get-Service -Name $OhmServiceName -ErrorAction SilentlyContinue
        if ($liveSvc -and $liveSvc.Status -eq 'Running') {
            Write-Host "Restarting OHM service to apply updated settings..." -ForegroundColor Yellow
            Restart-Service -Name $OhmServiceName -Force
        }
    } else {
        Write-Host "OHM settings are already correct." -ForegroundColor Green
    }
}

# ============================================================================
# Step 2 - Python virtual environment
# ============================================================================
Write-Host "`nStep 2/3 - Setting up Python virtual environment..." -ForegroundColor Cyan

try {
    $pythonVersion = & python --version 2>&1
    Write-Host "Found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Error "Python was not found in PATH. Please install Python 3.9+ from https://python.org and retry."
    exit 1
}

if (Test-Path $VenvPath) {
    Write-Host "Virtual environment already exists at '.venv' - skipping creation." -ForegroundColor Yellow
} else {
    Write-Host "Creating virtual environment at '.venv'..."
    python -m venv $VenvPath
    Write-Host "Virtual environment created." -ForegroundColor Green
}

Write-Host "Upgrading pip..."
& $PythonExe -m pip install --upgrade pip --quiet

Write-Host "Installing packages from requirements.txt..."
& $PythonExe -m pip install -r $RequirementsPath

Write-Host "Dependencies installed." -ForegroundColor Green

# ============================================================================
# Step 3 - Scraper scheduled task
# ============================================================================
Write-Host "`nStep 3/3 - Checking scheduled task '$ScraperTaskName'..." -ForegroundColor Cyan

$existing = Get-ScheduledTask -TaskName $ScraperTaskName -ErrorAction SilentlyContinue
$needsRegister = $true

if ($existing) {
    $existingExe = $existing.Actions | Select-Object -First 1 -ExpandProperty Execute
    if ($existingExe -eq $PythonExe) {
        Write-Host "Scheduled task already registered with correct configuration - skipping." -ForegroundColor Green
        $needsRegister = $false
    } else {
        Write-Host "Scheduled task configuration has changed - re-registering." -ForegroundColor Yellow
        if ($existing.State -eq "Running") { Stop-ScheduledTask -TaskName $ScraperTaskName }
        Unregister-ScheduledTask -TaskName $ScraperTaskName -Confirm:$false
    }
}

if ($needsRegister) {
    $action = New-ScheduledTaskAction `
        -Execute          $PythonExe `
        -Argument         "`"$ScraperScript`"" `
        -WorkingDirectory $ScriptDir

    # Trigger: run at every system startup
    $trigger = New-ScheduledTaskTrigger -AtStartup

    # Settings: run indefinitely, restart up to 5 times on failure, allow on battery
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -ExecutionTimeLimit  ([TimeSpan]::Zero) `
        -RestartCount        5 `
        -RestartInterval     (New-TimeSpan -Minutes 1) `
        -MultipleInstances   IgnoreNew

    # Principal: run as SYSTEM with highest privileges (required for WMI/OHM access)
    $principal = New-ScheduledTaskPrincipal `
        -UserId    "SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel  Highest

    Register-ScheduledTask `
        -TaskName    $ScraperTaskName `
        -Action      $action `
        -Trigger     $trigger `
        -Settings    $settings `
        -Principal   $principal `
        -Description "Prometheus exporter for hardware sensors via OpenHardwareMonitor. Managed by setup.ps1." `
        -Force | Out-Null

    Write-Host "Scheduled task registered successfully." -ForegroundColor Green
}

if ($Start) {
    Write-Host "`nStarting scraper task now..." -ForegroundColor Cyan
    Start-ScheduledTask -TaskName $ScraperTaskName
    Start-Sleep -Seconds 2
    $state = (Get-ScheduledTask -TaskName $ScraperTaskName).State
    Write-Host "Task state: $state" -ForegroundColor Green
}

# ============================================================================
# Summary
# ============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Setup complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
if (-not $SkipOHM) {
    Write-Host "  OHM Service : $OhmServiceName (running)"
}
Write-Host "  Task Name   : $ScraperTaskName"
Write-Host "  Python      : $PythonExe"
Write-Host "  Script      : $ScraperScript"
Write-Host "  Metrics     : http://localhost:9877/metrics (default)"
Write-Host ""
Write-Host "The scraper will start automatically on the next system boot." -ForegroundColor Cyan
Write-Host ""
Write-Host "Useful commands:"
Write-Host "  Start now    : Start-ScheduledTask -TaskName $ScraperTaskName"
Write-Host "  Stop         : Stop-ScheduledTask  -TaskName $ScraperTaskName"
Write-Host "  Check status : Get-ScheduledTask   -TaskName $ScraperTaskName | Select-Object TaskName, State"
Write-Host "  Uninstall    : .\setup.ps1 -Uninstall"
Write-Host ""