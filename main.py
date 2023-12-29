import win32com.client
import pythoncom
import threading
import tkinter as tk
from tkinter import scrolledtext
import time


class USBDeviceDetector(threading.Thread):
    def __init__(self, update_callback):
        super().__init__()
        self.update_callback = update_callback
        self.devices_before = self.get_pnp_devices()
        self.running = True

    def get_pnp_devices(self):
        pythoncom.CoInitialize()  # Initialize COM in the thread
        wmi = win32com.client.GetObject("winmgmts:")
        pnp_devices = wmi.InstancesOf("Win32_PnPEntity")
        devices = {device.DeviceID: device.Name for device in pnp_devices}
        pythoncom.CoUninitialize()  # Uninitialize COM
        return devices

    def run(self):
        while self.running:
            time.sleep(1)
            devices_now = self.get_pnp_devices()
            new_devices = devices_now.keys() - self.devices_before.keys()
            removed_devices = self.devices_before.keys() - devices_now.keys()

            # Update the GUI if there are changes
            if new_devices or removed_devices:
                self.update_callback(new_devices, removed_devices)

            self.devices_before = devices_now

    def stop(self):
        self.running = False


class USBDeviceLoggerGUI:
    def __init__(self, root):
        self.root = root
        root.title("USB Device Logger")

        # Create a ScrolledText widget for logging device changes
        self.log = scrolledtext.ScrolledText(root, state='disabled', width=70, height=20)
        self.log.grid(row=0, column=0, padx=10, pady=10)

        # Start the USB device detection in a separate thread
        self.detector = USBDeviceDetector(self.log_device_changes)
        self.detector.start()

    def log_device_changes(self, new_devices, removed_devices):
        message = ""
        for device_id in new_devices:
            message += f"Connected: {device_id}\n"
        for device_id in removed_devices:
            message += f"Disconnected: {device_id}\n"

        # Update the log widget with the detected changes
        if message:
            self.log.configure(state='normal')
            self.log.insert(tk.END, message)
            self.log.configure(state='disabled')
            self.log.yview(tk.END)

    def on_close(self):
        # Ensure thread is properly stopped before closing the GUI
        self.detector.stop()
        self.detector.join()
        self.root.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    gui = USBDeviceLoggerGUI(root)
    root.protocol("WM_DELETE_WINDOW", gui.on_close)
    root.mainloop()