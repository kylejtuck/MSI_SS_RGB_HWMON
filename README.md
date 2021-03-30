# MSI_SS_RGB_HWMON
Python script used with OpenHardware Monitor to use the per-key RGB SteelSeries keyboards of MSI laptops to
display hardware sensor details.

Requirements:
 OpenHardware Monitor (running)
 Python
 - pip install urllib3
 - pip install requests
 - pip install WMI

sensor_examples.txt contains sample sensor names and values.

keymapping_values.txt is a reference to the supported keys to use for graphing out the sensor value.

This script should run in the background. There is no built-in mechanism to stop/restart the script, so you will
need to use Task Manager.
