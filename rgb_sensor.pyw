from pystray import Icon as icon, Menu as menu, MenuItem as item
from PIL import Image, ImageDraw
import time
import wmi
import requests
import json
import os
import pythoncom

# This script depends on OpenHardwareMonitor (running). It is free and can be configured
# to run at Windows start.
# The SteelSeries gamesense-sdk documents can be found at https://github.com/SteelSeries/gamesense-sdk.

# define the sensors to monitor and key mappings
#   The SENSORS_KEY_MAPPING dictionary uses the format "Sensor.Name|Sensor.SensorType". The examples
#   given should make the format clear. This defines what sensors should be read, and which keys
#   should be used to display the colour gradient.
#
#   min_value: required integer value indicating the minimum expected value from sensor
#   max_value: required integer value indicating the maximum expected value from sensor
#		The difference between the min and max values map the value from the sensor to a guage value
#		between 0 and 100.
#		keymap: required dictionary that defines event names for gamesense-sdk, using format:
#			"eventname": [key_weight, key_definition, keys]
#				The eventname must consist of only uppercase and numbers.
#				The key_weight is an integer defining how much of the 0-100 scale is represented by the
#				key. For example, if you are using 4 equally sized keys to show the sensor value, the key_weight
#				would be 25 for all for eventnames (25 x 4 = 100).
#				The key_definition and keys values can take one of two formats.
#				1) "custom-zone-keys": [keynum]
#				2) "zone":"named_key_zone"
#				The "custom-zone-keys" are drawn from https://www.usb.org/document-library/hid-usage-tables-112, but can be
#				found in the included keymapping_values.txt (more reliable than named key zones). In this form, [keynum]
#				is an array, so you could specify multiple keys [x, y, z] or just one key [x].
#				Named key zones are easier to "figure out" when reading the code, but I found they don't always match
#				up with the keys on the MSI Leopard GL65. Other models might work better. The named key zones can be
#				found at https://github.com/SteelSeries/gamesense-sdk/blob/master/doc/api/standard-zones.md.
#	low_colour: required value from "red", "blue", "green", and "white"
#	high_colour: required value from "red", "blue", "green", and "white"

class SSE_CPU_LOAD:
    running = True
    headers = {"Content-Type":"application/json"}
    coreProps = {}
    SENSORS_KEY_MAPPING = {
        "CPU Package|Temperature": {
            "min_value": 35,
            "max_value": 100,
            "keymap": {
                "CPUTEMP0": [5, "custom-zone-keys", [0x29]],
                "CPUTEMP1": [5, "custom-zone-keys", [0x3a]],
                "CPUTEMP2": [5, "custom-zone-keys", [0x3b]],
                "CPUTEMP3": [5, "custom-zone-keys", [0x3c]],
                "CPUTEMP4": [5, "custom-zone-keys", [0x3d]],
                "CPUTEMP5": [5, "custom-zone-keys", [0x3e]],
                "CPUTEMP6": [5, "custom-zone-keys", [0x3f]],
                "CPUTEMP7": [5, "custom-zone-keys", [0x40]],
                "CPUTEMP8": [5, "custom-zone-keys", [0x41]],
                "CPUTEMP9": [5, "custom-zone-keys", [0x42]],
                "CPUTEMP10": [5, "custom-zone-keys", [0x43]],
                "CPUTEMP11": [5, "custom-zone-keys", [0x44]],
                "CPUTEMP12": [5, "custom-zone-keys", [0x45]],
                "CPUTEMP13": [5, "custom-zone-keys", [0x46]],
                "CPUTEMP14": [5, "custom-zone-keys", [0x47]],
                "CPUTEMP15": [5, "custom-zone-keys", [0x48]],
                "CPUTEMP16": [5, "custom-zone-keys", [0x49]],
                "CPUTEMP17": [5, "custom-zone-keys", [0x4c]],
                "CPUTEMP18": [5, "custom-zone-keys", [0x4b]],
                "CPUTEMP19": [5, "custom-zone-keys", [0x4e]],
            },
            "low_colour":"blue",
            "high_colour":"red"
        },
        "CPU Total|Load": {
            "min_value": 0,
            "max_value": 100,
            "keymap": {
                "CPULOAD0": [5, "custom-zone-keys", [0x35]],
                "CPULOAD1": [6, "custom-zone-keys", [0x1e]],
                "CPULOAD2": [5, "custom-zone-keys", [0x1f]],
                "CPULOAD3": [6, "custom-zone-keys", [0x20]],
                "CPULOAD4": [5, "custom-zone-keys", [0x21]],
                "CPULOAD5": [6, "custom-zone-keys", [0x22]],
                "CPULOAD6": [5, "custom-zone-keys", [0x23]],
                "CPULOAD7": [6, "custom-zone-keys", [0x24]],
                "CPULOAD8": [5, "custom-zone-keys", [0x25]],
                "CPULOAD9": [6, "custom-zone-keys", [0x26]],
                "CPULOAD10": [5, "custom-zone-keys", [0x27]],
                "CPULOAD11": [6, "custom-zone-keys", [0x2d]],
                "CPULOAD12": [5, "custom-zone-keys", [0x2e]],
                "CPULOAD13": [6, "custom-zone-keys", [0x2a]],
                "CPULOAD14": [5, "custom-zone-keys", [0x53]],
                "CPULOAD15": [6, "custom-zone-keys", [0x54]],
                "CPULOAD16": [5, "custom-zone-keys", [0x55]],
                "CPULOAD17": [6, "custom-zone-keys", [0x56]],
            },
            "low_colour":"green",
            "high_colour":"red"
        },
        "GPU Core|Temperature": {
            "min_value": 35,
            "max_value": 100,
            "keymap": {
                "GPUTEMP0": [5, "custom-zone-keys", [0x2b]],
                "GPUTEMP1": [6, "custom-zone-keys", [0x14]],
                "GPUTEMP2": [5, "custom-zone-keys", [0x1a]],
                "GPUTEMP3": [6, "custom-zone-keys", [0x08]],
                "GPUTEMP4": [5, "custom-zone-keys", [0x15]],
                "GPUTEMP5": [6, "custom-zone-keys", [0x17]],
                "GPUTEMP6": [5, "custom-zone-keys", [0x1c]],
                "GPUTEMP7": [6, "custom-zone-keys", [0x18]],
                "GPUTEMP8": [5, "custom-zone-keys", [0x0c]],
                "GPUTEMP9": [6, "custom-zone-keys", [0x12]],
                "GPUTEMP10": [5, "custom-zone-keys", [0x13]],
                "GPUTEMP11": [6, "custom-zone-keys", [0x2f]],
                "GPUTEMP12": [5, "custom-zone-keys", [0x30]],
                "GPUTEMP13": [6, "custom-zone-keys", [0x31]],
                "GPUTEMP14": [5, "custom-zone-keys", [0x5f]],
                "GPUTEMP15": [6, "custom-zone-keys", [0x60]],
                "GPUTEMP16": [5, "custom-zone-keys", [0x61]],
                "GPUTEMP17": [6, "custom-zone-keys", [0x57]],
            },
            "low_colour":"blue",
            "high_colour":"red"
        },
        "GPU Core|Load": {
            "min_value": 0,
            "max_value": 100,
            "keymap": {
                "GPULOAD0": [6, "custom-zone-keys", [0x39]],
                "GPULOAD1": [6, "custom-zone-keys", [0x04]],
                "GPULOAD2": [6, "custom-zone-keys", [0x16]],
                "GPULOAD3": [7, "custom-zone-keys", [0x07]],
                "GPULOAD4": [6, "custom-zone-keys", [0x09]],
                "GPULOAD5": [6, "custom-zone-keys", [0x0a]],
                "GPULOAD6": [6, "custom-zone-keys", [0x0b]],
                "GPULOAD7": [7, "custom-zone-keys", [0x0d]],
                "GPULOAD8": [6, "custom-zone-keys", [0x0e]],
                "GPULOAD9": [6, "custom-zone-keys", [0x0f]],
                "GPULOAD10": [6, "custom-zone-keys", [0x33]],
                "GPULOAD11": [7, "custom-zone-keys", [0x34]],
                "GPULOAD12": [6, "custom-zone-keys", [0x28]],
                "GPULOAD13": [6, "custom-zone-keys", [0x5c]],
                "GPULOAD14": [6, "custom-zone-keys", [0x5d]],
                "GPULOAD15": [7, "custom-zone-keys", [0x5e]],
            },
            "low_colour":"green",
            "high_colour":"red"
        },
        "Memory|Load": {
            "min_value": 0,
            "max_value": 100,
            "keymap": {
                "MEMLOAD0": [6, "custom-zone-keys", [0xe1]],
                "MEMLOAD1": [6, "custom-zone-keys", [0x1d]],
                "MEMLOAD2": [6, "custom-zone-keys", [0x1b]],
                "MEMLOAD3": [7, "custom-zone-keys", [0x06]],
                "MEMLOAD4": [6, "custom-zone-keys", [0x19]],
                "MEMLOAD5": [6, "custom-zone-keys", [0x05]],
                "MEMLOAD6": [6, "custom-zone-keys", [0x11]],
                "MEMLOAD7": [7, "custom-zone-keys", [0x10]],
                "MEMLOAD8": [6, "custom-zone-keys", [0x36]],
                "MEMLOAD9": [6, "custom-zone-keys", [0x37]],
                "MEMLOAD10": [6, "custom-zone-keys", [0x38]],
                "MEMLOAD11": [7, "custom-zone-keys", [0xe5]],
                "MEMLOAD12": [6, "custom-zone-keys", [0x52]],
                "MEMLOAD13": [6, "custom-zone-keys", [0x59]],
                "MEMLOAD14": [6, "custom-zone-keys", [0x5a]],
                "MEMLOAD15": [7, "custom-zone-keys", [0x5b]],
            },
            "low_colour":"green",
            "high_colour":"red"
        },
    }


    def shutdown(self):
        self.running = False
        json_to_post = {"game": "HWMONITOR"}
        r = requests.post('http://' + self.coreProps['address'] + '/stop_game', json=json_to_post, headers=self.headers)
        self.icon.stop()

    def refreshAll(self):
        for sensor in self.SENSORS_KEY_MAPPING:
            for event_name in self.SENSORS_KEY_MAPPING[sensor]['keymap']:
                # reset the key colour
                json_to_post = {"game":"HWMONITOR",
                                "event":event_name,
                                "data": {"value":100}}
                r = requests.post('http://' + self.coreProps['address'] + '/game_event', json=json_to_post, headers=self.headers)
                json_to_post = {"game":"HWMONITOR",
                                "event":event_name,
                                "data": {"value":0}}
                r = requests.post('http://' + self.coreProps['address'] + '/game_event', json=json_to_post, headers=self.headers)

    def runApp(self, icon):
        self.icon.visible = True
        pythoncom.CoInitialize()
        
        # define sleep duration between sensor checks
        sleepy_time = 1
        # JSON colour configurations
        guage_colour = {
            "red": {"red":255,"green":0,"blue":0},
            "green": {"red":0,"green":255,"blue":0},
            "blue": {"red":0,"green":0,"blue":255},
            "white": {"red":255,"green":255,"blue":255}
        }

        # Need to discover address:port from %ProgramData%\SteelSeries\SteelSeries Engine 3\coreProps.json
        programdata = os.getenv('PROGRAMDATA')
        with open(programdata + '\SteelSeries\SteelSeries Engine 3\coreProps.json') as SSEcoreProps:
            self.coreProps = json.load(SSEcoreProps)

        # De-register the game to clear out all events
        json_to_post = {"game": "HWMONITOR"}
        r = requests.post('http://' + self.coreProps['address'] + '/remove_game', json=json_to_post, headers=self.headers)

        # Register the app with SteelSeries
        json_to_post = {"game": "HWMONITOR","game_display_name": "RGB HW Monitor","developer": "Kyle J Tuck"}
        r = requests.post('http://' + self.coreProps['address'] + '/game_metadata', json=json_to_post, headers=self.headers)

        # Register the event handlers
        # https://github.com/SteelSeries/gamesense-sdk/blob/master/doc/api/writing-handlers-in-json.md
        for sensor in self.SENSORS_KEY_MAPPING:
            for event_name in self.SENSORS_KEY_MAPPING[sensor]['keymap']:
                json_to_post = {"game": "HWMONITOR",
                    "event":event_name,
                    "handlers":[ {"device-type":"rgb-per-key-zones",
                        self.SENSORS_KEY_MAPPING[sensor]['keymap'][event_name][1]:self.SENSORS_KEY_MAPPING[sensor]['keymap'][event_name][2],
                        "mode":"color",
                        "color":{
                            "gradient":{
                                "zero": guage_colour[self.SENSORS_KEY_MAPPING[sensor]["low_colour"]],
                                "hundred": guage_colour[self.SENSORS_KEY_MAPPING[sensor]["high_colour"]]
                            }
                        }
                    }]
                }
                r = requests.post('http://' + self.coreProps['address'] + '/bind_game_event', json=json_to_post, headers=self.headers)
                # reset the key colour
                json_to_post = {"game":"HWMONITOR",
                                "event":event_name,
                                "data": {"value":100}
                                }
                r = requests.post('http://' + self.coreProps['address'] + '/bind_game_event', json=json_to_post, headers=self.headers)
                json_to_post = {"game":"HWMONITOR",
                                "event":event_name,
                                "data": {"value":0}
                                }
                r = requests.post('http://' + self.coreProps['address'] + '/bind_game_event', json=json_to_post, headers=self.headers)
        # read the sensors
        while self.running:
            w = wmi.WMI(namespace="root\OpenHardwareMonitor")
            temperature_data = w.Sensor()
            for sensor in temperature_data:
                full_sensor_name = sensor.Name + '|' + sensor.SensorType
                if full_sensor_name in self.SENSORS_KEY_MAPPING.keys():
                    sensor_value = float(sensor.Value) - float(self.SENSORS_KEY_MAPPING[full_sensor_name]['min_value'])
                    #if full_sensor_name == 'CPU Total|Load':
                    #	print(sensor_value)
                    min_max_delta = float(self.SENSORS_KEY_MAPPING[full_sensor_name]['max_value'] - self.SENSORS_KEY_MAPPING[full_sensor_name]['min_value'])
                    if sensor_value >= 0.0:
                        sensor_guage = int(sensor_value / min_max_delta * 100.0)
                        for event_name in self.SENSORS_KEY_MAPPING[full_sensor_name]['keymap']:
                            if sensor_guage <= 0:
                                key_value = 0
                            elif self.SENSORS_KEY_MAPPING[full_sensor_name]['keymap'][event_name][0] > sensor_guage:
                                key_value = int(sensor_guage*1.0/float(self.SENSORS_KEY_MAPPING[full_sensor_name]['keymap'][event_name][0])*100.0)
                            else:
                                key_value = 100
                            json_to_post = {"game":"HWMONITOR",
                                            "event":event_name,
                                            "data": {"value":key_value}
                                            }
                            sensor_guage -= self.SENSORS_KEY_MAPPING[full_sensor_name]['keymap'][event_name][0]
                            if self.running:
                                r = requests.post('http://' + self.coreProps['address'] + '/game_event', json=json_to_post, headers=self.headers)
            time.sleep(sleepy_time)

    def start(self):
        self.icon = icon('SSE CPU Load')
        image = Image.open('msi-282303.png')
        self.icon.icon = image
        self.icon.menu = menu(
                item('Refresh All', self.refreshAll),
                item('Exit', self.shutdown)
        )
        self.icon.run(self.runApp)

        
if __name__ == '__main__':
    mon = SSE_CPU_LOAD()
    mon.start()
