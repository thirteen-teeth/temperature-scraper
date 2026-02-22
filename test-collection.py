import wmi
from collections import defaultdict
from pprint import pprint

def collect_all_sensors():
    """
    Dynamically discovers all sensor categories and values from OpenHardwareMonitor.
    Returns a dict keyed by SensorType, each containing a dict of {sensor_name: value}.
    """
    w = wmi.WMI(namespace="root\\OpenHardwareMonitor")
    sensors = w.Sensor()

    values = defaultdict(dict)
    for sensor in sensors:
        sensor_type = sensor.SensorType
        values[sensor_type][sensor.Name] = sensor.Value

    return dict(values)

def main():
    data = collect_all_sensors()

    print(f"Discovered {len(data)} sensor categories:\n")
    for sensor_type, readings in sorted(data.items()):
        print(f"[{sensor_type}]")
        for name, value in sorted(readings.items()):
            print(f"  {name}: {value}")
        print()

if __name__ == "__main__":
    main()
