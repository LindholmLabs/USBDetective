import win32com.client
import time

def get_pnp_devices():
    wmi = win32com.client.GetObject("winmgmts:")
    pnp_devices = wmi.InstancesOf("Win32_PnPEntity")
    return {device.DeviceID: device.Name for device in pnp_devices}

if __name__ == '__main__':
    devices_before = get_pnp_devices()
    print("Waiting for PnP device change...")
    while True:
        time.sleep(1)  # Adding a delay to reduce CPU usage
        devices_now = get_pnp_devices()

        new_devices = devices_now.keys() - devices_before.keys()
        removed_devices = devices_before.keys() - devices_now.keys()

        if new_devices:
            print("New PnP devices detected:")
            for device_id in new_devices:
                print(devices_now[device_id])

        if removed_devices:
            print("PnP devices removed:")
            for device_id in removed_devices:
                print(devices_before[device_id])

        devices_before = devices_now