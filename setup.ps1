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
$OhmZip           = "$env:TEMP\openhardwaremonitor.zip"
$OhmExe           = "$OhmDir\OpenHardwareMonitor.exe"
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
    foreach ($f in @($OhmZip, $NssmZip)) {
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

    if ($ohmSvc -and (Test-Path $OhmExe)) {
        # Already installed - just ensure it is running
        if ($ohmSvc.Status -ne "Running") {
            Write-Host "OHM service exists but is not running - starting it..." -ForegroundColor Yellow
            Start-Service -Name $OhmServiceName
        } else {
            Write-Host "OpenHardwareMonitor is already installed and running - skipping." -ForegroundColor Green
        }
    } else {
        Write-Host "Installing OpenHardwareMonitor as a service..." -ForegroundColor Cyan

        $fallbackOhmUrl = "https://openhardwaremonitor.org/files/openhardwaremonitor-v0.9.6.zip"
        $nssmUrl        = "https://nssm.cc/release/nssm-2.24.zip"

        # Try to resolve the latest release from the OHM homepage
        $ohmUrl = $fallbackOhmUrl
        try {
            $resp  = Invoke-WebRequest -Uri "https://openhardwaremonitor.org/" -UseBasicParsing
            $match = [regex]::Match(
                $resp.Content,
                "openhardwaremonitor-v\d+\.\d+\.\d+\.zip",
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
            if ($match.Success) { $ohmUrl = "https://openhardwaremonitor.org/files/$($match.Value)" }
        } catch {
            Write-Host "Could not fetch latest OHM version - using fallback URL." -ForegroundColor Yellow
        }

        Write-Host "Downloading OpenHardwareMonitor..."
        Invoke-WebRequest -Uri $ohmUrl -OutFile $OhmZip
        $ohmExtractTemp = "$env:TEMP\ohm-extract"
        if (Test-Path $ohmExtractTemp) { Remove-Item $ohmExtractTemp -Recurse -Force }
        Expand-Archive -Path $OhmZip -DestinationPath $ohmExtractTemp

        # The zip may contain a subdirectory - find the exe wherever it landed
        $ohmExeFound = Get-ChildItem -Path $ohmExtractTemp -Filter "OpenHardwareMonitor.exe" -Recurse | Select-Object -First 1
        if (-not $ohmExeFound) {
            Write-Error "Could not find OpenHardwareMonitor.exe in the downloaded archive."
            exit 1
        }
        if (Test-Path $OhmDir) { Remove-Item $OhmDir -Recurse -Force }
        Move-Item -Path $ohmExeFound.DirectoryName -Destination $OhmDir
        Remove-Item $ohmExtractTemp -Recurse -Force -ErrorAction SilentlyContinue
        $OhmExe = Join-Path $OhmDir "OpenHardwareMonitor.exe"

        # Download NSSM only if not already extracted
        if (Test-Path $NssmExe) {
            Write-Host "NSSM already present - skipping download." -ForegroundColor Yellow
        } else {
            Write-Host "Downloading NSSM..."
            Invoke-WebRequest -Uri $nssmUrl -OutFile $NssmZip
            if (Test-Path $NssmDir) { Remove-Item $NssmDir -Recurse -Force }
            Expand-Archive -Path $NssmZip -DestinationPath $NssmDir
        }

        # Remove existing (broken) service if present; wait for SCM to deregister
        if ($ohmSvc) {
            if ($ohmSvc.Status -ne "Stopped") { Stop-Service -Name $OhmServiceName -Force }
            sc.exe delete $OhmServiceName | Out-Null
            Write-Host "Removed existing OHM service." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
        }

        Write-Host "Installing the OHM service..."
        & $NssmExe install $OhmServiceName $OhmExe
        & $NssmExe set     $OhmServiceName AppDirectory $OhmDir
        & $NssmExe set     $OhmServiceName Start SERVICE_AUTO_START
        & $NssmExe set     $OhmServiceName AppNoConsole 1

        Write-Host "Starting the OHM service..."
        & $NssmExe start $OhmServiceName
        Start-Sleep -Seconds 2
        $ohmStatus = (Get-Service -Name $OhmServiceName -ErrorAction SilentlyContinue).Status
        if ($ohmStatus -ne "Running") {
            Write-Warning "OHM service status is '$ohmStatus' - it may still be starting. Check with: Get-Service $OhmServiceName"
        } else {
            Write-Host "OpenHardwareMonitor service installed and started." -ForegroundColor Green
        }
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