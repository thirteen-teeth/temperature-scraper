#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Thin wrapper — OHM installation is now handled by setup.ps1.
    This script is kept for backwards compatibility with the README.

    To install OHM only (without setting up the scraper), run:
        .\setup.ps1 -SkipOHM:$false

    To run the full setup (OHM + venv + scheduled task), run:
        .\setup.ps1

.PARAMETER Uninstall
    Forwards to: .\setup.ps1 -Uninstall
#>
param(
    [switch]$Uninstall
)

$ScriptDir = $PSScriptRoot

if ($Uninstall) {
    Write-Host "Forwarding to setup.ps1 -Uninstall..." -ForegroundColor Yellow
    & "$ScriptDir\setup.ps1" -Uninstall
} else {
    Write-Host "Forwarding to setup.ps1..." -ForegroundColor Yellow
    & "$ScriptDir\setup.ps1"
}

if ($Uninstall) {
	$svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
	if ($null -ne $svc) {
		if ($svc.Status -ne "Stopped") {
			Stop-Service -Name $serviceName -Force
		}
		sc.exe delete $serviceName | Out-Null
		Write-Host "Removed service: $serviceName"
	} else {
		Write-Host "Service not found: $serviceName"
	}

	if (Test-Path $ohmDir) {
		Remove-Item $ohmDir -Recurse -Force
		Write-Host "Removed install directory: $ohmDir"
	}

	if (Test-Path $ohmZip) { Remove-Item $ohmZip -Force }
	if (Test-Path $nssmZip) { Remove-Item $nssmZip -Force }
	if (Test-Path $nssmDir) { Remove-Item $nssmDir -Recurse -Force }
	return
}

$fallbackOhmUrl = "https://openhardwaremonitor.org/files/openhardwaremonitor-v0.9.6.zip"
$ohmUrl = $fallbackOhmUrl
$nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"

$ohmPage = "https://openhardwaremonitor.org/"
try {
	$resp = Invoke-WebRequest -Uri $ohmPage -UseBasicParsing
	$match = [regex]::Match(
		$resp.Content,
		"openhardwaremonitor-v\d+\.\d+\.\d+\.zip",
		[System.Text.RegularExpressions.RegexOptions]::IgnoreCase
	)
	if ($match.Success) {
		$ohmUrl = "https://openhardwaremonitor.org/files/$($match.Value)"
	}
} catch {
	$ohmUrl = $fallbackOhmUrl
}

# Download and extract OpenHardwareMonitor
Invoke-WebRequest -Uri $ohmUrl -OutFile $ohmZip
if (Test-Path $ohmDir) { Remove-Item $ohmDir -Recurse -Force }
Expand-Archive -Path $ohmZip -DestinationPath $ohmDir

# Download and extract NSSM
Invoke-WebRequest -Uri $nssmUrl -OutFile $nssmZip
if (Test-Path $nssmDir) { Remove-Item $nssmDir -Recurse -Force }
Expand-Archive -Path $nssmZip -DestinationPath $nssmDir

# Install service
& $nssmExe install $serviceName $ohmExe
& $nssmExe set $serviceName AppDirectory $ohmDir
& $nssmExe set $serviceName Start SERVICE_AUTO_START
& $nssmExe set $serviceName AppNoConsole 1

# Start service
Start-Service -Name $serviceName

Write-Host "Installed and started service: $serviceName"