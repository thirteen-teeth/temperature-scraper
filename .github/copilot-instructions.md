# Copilot Instructions

## Documentation

`README.MD` is the primary documentation for this project. Keep it in sync whenever relevant changes are made:

- If a metric name, label, or unit changes in `scraper.py`, update the Metrics table in `README.MD`.
- If new environment variables are added or defaults change, update the Configuration table in `README.MD`.
- If setup steps or script flags change in `setup.ps1`, update the Setup section in `README.MD`.
- If `test-collection.py` output format changes, update the example output in the Testing Sensor Collection section of `README.MD`.

## Code style

- Sensor type metadata (units, descriptions) lives in `SENSOR_TYPE_UNITS` in `scraper.py`. Add new entries there rather than hardcoding strings elsewhere.
- Prometheus metric names follow the pattern `ohm_{sensor_type}_{unit}` (e.g. `ohm_temperature_celsius`).
- New sensor types not in `SENSOR_TYPE_UNITS` are still collected dynamically — only add an entry when the unit is known.
