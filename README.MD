# temperature-scraper

Collects Windows hardware sensor data (temperature, power, fan) from Open Hardware Monitor and exposes it as Prometheus metrics.

## Requirements

- Windows
- Python 3.9+
- PowerShell (run as Administrator for setup)

## Setup

`setup.ps1` handles everything in one shot:

1. Downloads and installs **OpenHardwareMonitor** as a Windows service (via NSSM)
2. Creates a **Python virtual environment** and installs dependencies
3. Registers the **Prometheus exporter** as a Windows Scheduled Task that starts automatically at boot

Open PowerShell as Administrator and run:

```powershell
.\setup.ps1
```

To also start the scraper immediately (without waiting for a reboot):

```powershell
.\setup.ps1 -Start
```

If OpenHardwareMonitor is already installed, skip that step:

```powershell
.\setup.ps1 -SkipOHM
```

### Uninstall

Removes the scheduled task, OHM service, and cleans up downloaded files:

```powershell
.\setup.ps1 -Uninstall
```

### Manage the scraper manually

```powershell
Start-ScheduledTask  -TaskName 'TemperatureScraper'
Stop-ScheduledTask   -TaskName 'TemperatureScraper'
Get-ScheduledTask    -TaskName 'TemperatureScraper' | Select-Object TaskName, State
```

## Configuration

The scraper reads the following environment variables:

| Variable | Default | Description |
| --- | --- | --- |
| `POLLING_INTERVAL_SECONDS` | `5` | How often to poll OHM sensors |
| `EXPORTER_PORT` | `9877` | Port the Prometheus HTTP server listens on |

## Metrics

The Prometheus endpoint is available at `http://localhost:9877/metrics` by default.

Exposed metrics:

- `temperature_degrees{temperature_sensor="..."}`
- `power_watts{power_sensor="..."}`
- `fan_rpm{fan_sensor="..."}`
