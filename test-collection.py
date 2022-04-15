import wmi
from pprint import pprint
import json

def main():
    w = wmi.WMI(namespace="root\OpenHardwareMonitor")
    temperature_infos = w.Sensor()
    values = { 'temp': {}, 'power': {}, 'fan': {} }
    for sensor in temperature_infos:
        if sensor.SensorType==u'Temperature':
            values['temp'][sensor.Name] = sensor.Value
        if sensor.SensorType==u'Power':
            values['power'][sensor.Name] = sensor.Value
        if sensor.SensorType==u'Fan':
            values['fan'][sensor.Name] = sensor.Value

    pprint(values)

if __name__ == "__main__":
    main()
