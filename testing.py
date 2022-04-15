import wmi
from pprint import pprint
import json

# ensure open hardware monitor is installed

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