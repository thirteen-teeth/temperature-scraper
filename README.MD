# temperature-scraper

Collects Windows hardware sensor data (temperature, power, fan) from Open Hardware Monitor and exposes it as Prometheus metrics.

## Requirements

- Windows
- Python 3.9+
- Open Hardware Monitor installed and running (or installed as a service)

## Install dependencies

```bash
python -m pip install -r requirements.txt
```

## Install Open Hardware Monitor as a service (optional)

Run PowerShell as Administrator:

```powershell
./install-openhardwaremonitor.ps1
```

Uninstall:

```powershell
./install-openhardwaremonitor.ps1 -Uninstall
```

## Run the exporter

```bash
python scraper.py
```

The Prometheus endpoint will be available at http://localhost:9877/metrics by default.

## Configuration

Environment variables:

- `POLLING_INTERVAL_SECONDS` (default: `5`)
- `EXPORTER_PORT` (default: `9877`)

Example:

```bash
POLLING_INTERVAL_SECONDS=2 EXPORTER_PORT=9100 python scraper.py
```

## Metrics

- `temperature_degrees{temperature_sensor="..."}`
- `power_watts{power_sensor="..."}`
- `fan_rpm{fan_sensor="..."}`
