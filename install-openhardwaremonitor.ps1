# Requires admin
$ErrorActionPreference = "Stop"

$ohmUrl = "https://openhardwaremonitor.org/files/openhardwaremonitor-v0.9.6.zip"
$ohmDir = "C:\Program Files\OpenHardwareMonitor"
$ohmZip = "$env:TEMP\openhardwaremonitor.zip"
$ohmExe = "$ohmDir\OpenHardwareMonitor.exe"

$nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"
$nssmZip = "$env:TEMP\nssm.zip"
$nssmDir = "$env:TEMP\nssm"
$nssmExe = "$nssmDir\nssm-2.24\win64\nssm.exe"

# Download and extract OpenHardwareMonitor
Invoke-WebRequest -Uri $ohmUrl -OutFile $ohmZip
if (Test-Path $ohmDir) { Remove-Item $ohmDir -Recurse -Force }
Expand-Archive -Path $ohmZip -DestinationPath $ohmDir

# Download and extract NSSM
Invoke-WebRequest -Uri $nssmUrl -OutFile $nssmZip
if (Test-Path $nssmDir) { Remove-Item $nssmDir -Recurse -Force }
Expand-Archive -Path $nssmZip -DestinationPath $nssmDir

# Install service
$serviceName = "OpenHardwareMonitor"
& $nssmExe install $serviceName $ohmExe
& $nssmExe set $serviceName AppDirectory $ohmDir
& $nssmExe set $serviceName Start SERVICE_AUTO_START
& $nssmExe set $serviceName AppNoConsole 1

# Start service
Start-Service -Name $serviceName

Write-Host "Installed and started service: $serviceName"