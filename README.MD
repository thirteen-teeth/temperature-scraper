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
Start-ScheduledTask -TaskName 'TemperatureScraper'
Stop-ScheduledTask -TaskName 'TemperatureScraper'
Get-ScheduledTask -TaskName 'TemperatureScraper' | Select-Object TaskName, State
```

## Manual Virtual Environment Setup

If you prefer to set up the Python environment without running `setup.ps1`, or want to work on the project manually:

```powershell
# Create the virtual environment
python -m venv .venv

# Activate it
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

To deactivate when you're done:

```powershell
deactivate
```

## Testing Sensor Collection

`test-collection.py` queries OpenHardwareMonitor directly via WMI and prints all detected sensor values. Use it to verify that OHM is running and sensors are being exposed correctly before running the full scraper.

Activate the virtual environment first, then run:

```powershell
.\.venv\Scripts\Activate.ps1
python test-collection.py
```

Example output:

```
{'fan': {'Chassis Fan #1': 0.0, 'CPU Fan': 1234.0},
 'power': {'CPU Package': 45.2, 'CPU Cores': 30.1},
 'temp': {'CPU Core #1': 52.0, 'CPU Core #2': 49.0, 'CPU Package': 53.0}}
```

> **Note:** OpenHardwareMonitor must be running (as a service or standalone) for WMI sensor data to be available.

## Configuration

The scraper reads the following environment variables:

| Variable | Default | Description |
| --- | --- | --- |
| `POLLING_INTERVAL_SECONDS` | `5` | How often to poll OHM sensors |
| `EXPORTER_PORT` | `9877` | Port the Prometheus HTTP server listens on |

## Metrics

The Prometheus endpoint is available at `http://localhost:9877/metrics` by default.

Metrics are discovered dynamically from OpenHardwareMonitor and exposed using the naming pattern `ohm_{sensor_type}{sensor="..."}`. All metrics use a single `sensor` label containing the sensor name.

| Metric | Unit | Example sensor labels |
| --- | --- | --- |
| `ohm_temperature_celsius{sensor="..."}` | °C | `CPU Package`, `CPU Core #1` |
| `ohm_power_watts{sensor="..."}` | Watts | `CPU Package`, `CPU Cores` |
| `ohm_fan_rpm{sensor="..."}` | RPM | `CPU Fan`, `Chassis Fan #1` |
| `ohm_load_percent{sensor="..."}` | % | `CPU Total`, `CPU Core #1` |
| `ohm_clock_megahertz{sensor="..."}` | MHz | `CPU Core #1`, `Bus Speed` |
| `ohm_voltage_volts{sensor="..."}` | V | `CPU VCore` |
| `ohm_data_gigabytes{sensor="..."}` | GB | `Used Memory`, `Available Memory` |

> Additional sensor types reported by OpenHardwareMonitor will be exposed automatically using the same `ohm_{sensor_type}` pattern.
