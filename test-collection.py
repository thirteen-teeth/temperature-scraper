import wmi
from collections import defaultdict

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

def main():
    data = collect_all_sensors()

    print(f"Discovered {len(data)} sensor categories:\n")
    for sensor_type, readings in sorted(data.items()):
        print(f"[{sensor_type}]")
        for r in sorted(readings, key=lambda r: (r['hardware_type'], r['hardware'], r['name'])):
            print(f"  {r['hardware_type']} / {r['hardware']} / {r['name']}: {r['value']}")
        print()

if __name__ == "__main__":
    main()
