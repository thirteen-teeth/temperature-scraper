# Requires admin
param(
	[switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$ohmDir = "C:\Program Files\OpenHardwareMonitor"
$ohmZip = "$env:TEMP\openhardwaremonitor.zip"
$ohmExe = "$ohmDir\OpenHardwareMonitor.exe"

$nssmZip = "$env:TEMP\nssm.zip"
$nssmDir = "$env:TEMP\nssm"
$nssmExe = "$nssmDir\nssm-2.24\win64\nssm.exe"

$serviceName = "OpenHardwareMonitor"

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