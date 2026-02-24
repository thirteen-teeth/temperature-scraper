import re
import os
import time
import requests
from collections import defaultdict
from prometheus_client import start_http_server, Gauge

# Sensor type metadata: maps OHM SensorType -> Prometheus unit suffix and description.
# New sensor types not listed here are still collected dynamically without a unit suffix.
SENSOR_TYPE_UNITS = {
    'Temperature': {'suffix': 'celsius',     'description': 'Temperature in degrees Celsius'},
    'Power':       {'suffix': 'watts',        'description': 'Power usage in Watts'},
    'Fan':         {'suffix': 'rpm',          'description': 'Fan speed in RPM'},
    'Load':        {'suffix': 'percent',      'description': 'Load as a percentage'},
    'Clock':       {'suffix': 'megahertz',    'description': 'Clock speed in MHz'},
    'Voltage':     {'suffix': 'volts',        'description': 'Voltage in Volts'},
    'Data':        {'suffix': 'gigabytes',    'description': 'Data amount in Gigabytes'},
    'Throughput':  {'suffix': 'bytes',        'description': 'Throughput in Bytes/s'},
    'Level':       {'suffix': 'percent',      'description': 'Level as a percentage'},
    'Control':     {'suffix': 'percent',      'description': 'Control value as a percentage'},
    'Humidity':    {'suffix': 'percent',      'description': 'Relative humidity as a percentage'},
    'Noise':       {'suffix': 'decibels',     'description': 'Noise level in dB'},
}


def sensor_type_to_metric_name(sensor_type):
    """Convert an OHM SensorType string to a valid Prometheus metric name (ohm_{type}_{unit})."""
    name = re.sub(r'[^a-z0-9_]', '_', sensor_type.lower())
    unit = SENSOR_TYPE_UNITS.get(sensor_type, {}).get('suffix')
    return f"ohm_{name}_{unit}" if unit else f"ohm_{name}"


def _infer_hardware_type(hardware_id):
    """Infer hardware category from an OHM identifier (e.g. '/intelcpu/0' -> 'CPU')."""
    if not hardware_id:
        return 'Unknown'
    segment = hardware_id.strip('/').split('/')[0].lower()
    mapping = {
        'intelcpu': 'CPU',   'amdcpu': 'CPU',     'cpu': 'CPU',
        'nvidiagpu': 'GPU',  'amdgpu': 'GPU',      'intelgpu': 'GPU', 'gpu': 'GPU',
        'hdd': 'Storage',    'ssd': 'Storage',     'nvme': 'Storage', 'storage': 'Storage',
        'ram': 'RAM',        'memory': 'RAM',
        'lpc': 'Motherboard', 'superio': 'Motherboard', 'motherboard': 'Motherboard',
        'nic': 'Network',    'network': 'Network',
        'battery': 'Battery',
    }
    for key, value in mapping.items():
        if key in segment:
            return value
    return segment.upper() if segment else 'Unknown'


def _parse_value(value_str):
    """Parse a formatted OHM value string (e.g. '52.0 °C', '1 234 RPM') to float.

    Takes the first whitespace-delimited token (stripping thousands-separator commas)
    and attempts a float conversion, returning None on failure.
    """
    if not value_str:
        return None
    token = value_str.strip().split()[0].replace(',', '') if value_str.strip() else ''
    try:
        return float(token)
    except (ValueError, TypeError):
        return None


def _traverse_node(node, results, hardware_name, hardware_type):
    """Recursively walk the data.json node tree, accumulating sensor readings.

    Standard OpenHardwareMonitor /data.json uses NodeId path strings to encode
    the hardware/sensor hierarchy:
      - Hardware node : 2 segments  e.g. /intelcpu/0  or  /lpc/nct6798d
      - Sensor node   : 4 segments  e.g. /intelcpu/0/temperature/0
    Group/header nodes have non-path NodeIds (e.g. "Temperatures") and are skipped.
    """
    node_id = node.get('NodeId', '')

    if isinstance(node_id, str) and node_id.startswith('/'):
        parts = [p for p in node_id.strip('/').split('/') if p]

        # Hardware node – update context for descendant sensors.
        if len(parts) == 2:
            hardware_name = node.get('Text', hardware_name)
            hardware_type = _infer_hardware_type(node_id)

        # Sensor node – collect the reading.
        elif len(parts) == 4:
            sensor_type = parts[2].capitalize()   # e.g. 'temperature' -> 'Temperature'
            value = _parse_value(node.get('Value', ''))
            if value is not None:
                results[sensor_type].append({
                    'name':          node.get('Text', node_id),
                    'value':         value,
                    'hardware':      hardware_name or 'Unknown',
                    'hardware_type': hardware_type or 'Unknown',
                })

    for child in node.get('Children', []):
        if isinstance(child, dict):
            _traverse_node(child, results, hardware_name, hardware_type)


def collect_all_sensors(ohm_url):
    """
    Fetch sensor data from OpenHardwareMonitor's REST API and return a dict keyed
    by SensorType, each containing a list of sensor reading dicts:
      {
        'Temperature': [
          {'name': 'CPU Core #1', 'value': 52.0, 'hardware': 'Intel Core i9', 'hardware_type': 'CPU'},
          ...
        ],
        ...
      }
    """
    try:
        resp = requests.get(f"{ohm_url}/data.json", timeout=5)
        resp.raise_for_status()
        data = resp.json()
    except requests.RequestException as exc:
        print(f"Warning: failed to reach OHM at {ohm_url}: {exc}")
        return {}

    results = defaultdict(list)
    _traverse_node(data, results, hardware_name=None, hardware_type=None)
    return dict(results)


class AppMetrics:
    def __init__(self, polling_interval_seconds=5, ohm_url="http://localhost:8086"):
        self.polling_interval_seconds = polling_interval_seconds
        self.ohm_url = ohm_url
        self.gauges = {}  # Gauges are created dynamically as new sensor types are discovered.

    def _get_or_create_gauge(self, sensor_type):
        if sensor_type not in self.gauges:
            metric_name = sensor_type_to_metric_name(sensor_type)
            description = SENSOR_TYPE_UNITS.get(sensor_type, {}).get(
                'description', f"OpenHardwareMonitor {sensor_type} sensor readings"
            )
            self.gauges[sensor_type] = Gauge(
                metric_name, description, ["sensor", "hardware", "hardware_type"]
            )
        return self.gauges[sensor_type]

    def run_metrics_loop(self):
        while True:
            self.fetch()
            time.sleep(self.polling_interval_seconds)

    def fetch(self):
        data = collect_all_sensors(self.ohm_url)
        for sensor_type, readings in data.items():
            gauge = self._get_or_create_gauge(sensor_type)
            for r in readings:
                if r['value'] is not None:
                    gauge.labels(
                        sensor=r['name'],
                        hardware=r['hardware'],
                        hardware_type=r['hardware_type'],
                    ).set(r['value'])


def main():
    polling_interval_seconds = int(os.getenv("POLLING_INTERVAL_SECONDS", "5"))
    exporter_port = int(os.getenv("EXPORTER_PORT", "9877"))
    ohm_url = os.getenv("OHM_URL", "http://localhost:8086")

    app_metrics = AppMetrics(
        polling_interval_seconds=polling_interval_seconds,
        ohm_url=ohm_url,
    )
    start_http_server(exporter_port)
    app_metrics.run_metrics_loop()


if __name__ == "__main__":
    main()