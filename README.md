# MSI_SS_RGB_HWMON
Python script used with OpenHardware Monitor to use the per-key RGB SteelSeries keyboards of MSI laptops to
display hardware sensor details.

Requirements:
 OpenHardware Monitor (running)
 Python
 - `pip install requests wmi Pillow pystray`

sensor_examples.txt contains sample sensor names and values.

keymapping_values.txt is a reference to the supported keys to use for graphing out the sensor value.

This script should run in the background and will show up on the system tray. To shutdown the app, right click the icon and select exit. After changing backlight brightness, to force a refresh, right click and select refresh all
