import wmi
import os
import time
from prometheus_client import start_http_server, Gauge

''' ensure open hardware monitor is installed to read values from WMI '''
''' https://openhardwaremonitor.org/ '''

def collect_data():
    w = wmi.WMI(namespace="root\OpenHardwareMonitor")
    temperature_infos = w.Sensor()
    values = { 'temperature': {}, 'power': {}, 'fan': {} }
    for sensor in temperature_infos:
        if sensor.SensorType==u'Temperature':
            values['temperature'][sensor.Name] = sensor.Value
        if sensor.SensorType==u'Power':
            values['power'][sensor.Name] = sensor.Value
        if sensor.SensorType==u'Fan':
            values['fan'][sensor.Name] = sensor.Value
    return values

class AppMetrics:
    def __init__(self, polling_interval_seconds=5):
        self.polling_interval_seconds = polling_interval_seconds

        # Prometheus metrics to collect
        self.temperature = Gauge(
            "temperature_degrees",
            "Current temperature",
            ["temperature_sensor"]
        )
        self.power = Gauge(
            "power_watts",
            "Current power usage",
            ["power_sensor"]
        )
        self.fan = Gauge(
            "fan_rpm",
            "Current fan speed",
            ["fan_sensor"]
        )

    def run_metrics_loop(self):
        while True:
            self.fetch()
            time.sleep(self.polling_interval_seconds)

    def fetch(self):
        # Fetch raw data from the wmi
        data = collect_data()

        # Update Prometheus metrics with application metrics
        for key, value in data['temperature'].items():
            self.temperature.labels(key).set(value)
        for key, value in data['fan'].items():
            self.fan.labels(key).set(value)
        for key, value in data['power'].items():
            self.power.labels(key).set(value)

        #getattr(object, attrname)
        #setattr(object, attrname, value)
           
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
