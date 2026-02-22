import re
import wmi
import os
import time
from collections import defaultdict
from prometheus_client import start_http_server, Gauge

''' ensure open hardware monitor is installed to read values from WMI '''
''' https://openhardwaremonitor.org/ '''

def collect_all_sensors():
    """
    Dynamically discovers all sensor categories from OpenHardwareMonitor,
    enriched with hardware name and type by joining against the Hardware WMI class.

    Returns a dict keyed by SensorType, each containing a list of sensor reading dicts:
      {
        'Temperature': [
          {'name': 'CPU Core #1', 'value': 52.0, 'hardware': 'AMD Ryzen 9 5900X', 'hardware_type': 'CPU'},
          ...
        ],
        ...
      }
    """
    w = wmi.WMI(namespace="root\\OpenHardwareMonitor")

    # Build a lookup of hardware by Identifier so sensors can resolve their parent.
    hardware_by_id = {
        hw.Identifier: {'name': hw.Name, 'type': hw.HardwareType}
        for hw in w.Hardware()
    }

    values = defaultdict(list)
    for sensor in w.Sensor():
        hw = hardware_by_id.get(sensor.Parent, {})
        values[sensor.SensorType].append({
            'name':          sensor.Name,
            'value':         sensor.Value,
            'hardware':      hw.get('name', 'Unknown'),
            'hardware_type': hw.get('type', 'Unknown'),
        })

    return dict(values)

def sensor_type_to_metric_name(sensor_type):
    """Convert an OHM SensorType string to a valid Prometheus metric name."""
    name = sensor_type.lower()
    name = re.sub(r'[^a-z0-9_]', '_', name)
    unit = SENSOR_TYPE_UNITS.get(sensor_type, {}).get('suffix')
    return f"ohm_{name}_{unit}" if unit else f"ohm_{name}"

# Known OHM sensor types with their units and human-readable descriptions.
# Unknown types will still be exposed dynamically, without a unit suffix.
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
}

class AppMetrics:
    def __init__(self, polling_interval_seconds=5):
        self.polling_interval_seconds = polling_interval_seconds
        # Gauges are created dynamically as new sensor types are discovered
        self.gauges = {}

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
        data = collect_all_sensors()

        for sensor_type, readings in data.items():
            gauge = self._get_or_create_gauge(sensor_type)
            for reading in readings:
                if reading['value'] is not None:
                    gauge.labels(
                        sensor=reading['name'],
                        hardware=reading['hardware'],
                        hardware_type=reading['hardware_type'],
                    ).set(reading['value'])

def main():
    polling_interval_seconds = int(os.getenv("POLLING_INTERVAL_SECONDS", "5"))
    exporter_port = int(os.getenv("EXPORTER_PORT", "9877"))

    app_metrics = AppMetrics(
        polling_interval_seconds=polling_interval_seconds
    )
    start_http_server(exporter_port)
    app_metrics.run_metrics_loop()

if __name__ == "__main__":
    main()
