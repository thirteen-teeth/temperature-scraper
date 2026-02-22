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
    Dynamically discovers all sensor categories and values from OpenHardwareMonitor.
    Returns a dict keyed by SensorType, each containing a dict of {sensor_name: value}.
    """
    w = wmi.WMI(namespace="root\\OpenHardwareMonitor")
    sensors = w.Sensor()

    values = defaultdict(dict)
    for sensor in sensors:
        values[sensor.SensorType][sensor.Name] = sensor.Value

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
            self.gauges[sensor_type] = Gauge(metric_name, description, ["sensor"])
        return self.gauges[sensor_type]

    def run_metrics_loop(self):
        while True:
            self.fetch()
            time.sleep(self.polling_interval_seconds)

    def fetch(self):
        data = collect_all_sensors()

        for sensor_type, readings in data.items():
            gauge = self._get_or_create_gauge(sensor_type)
            for sensor_name, value in readings.items():
                if value is not None:
                    gauge.labels(sensor=sensor_name).set(value)

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
