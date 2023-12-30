import win32com.client
import pythoncom
import threading
import tkinter as tk
from tkinter import ttk
import time


class USBDeviceDetector(threading.Thread):
    def __init__(self, update_callback):
        super().__init__()
        self.update_callback = update_callback
        self.running = True
        self.devices_before = {}

    def run(self):
        pythoncom.CoInitialize()  # Initialize COM in the thread

        # Initialize devices_before in the thread after COM initialization
        self.devices_before = self.get_pnp_devices()

        while self.running:
            time.sleep(1)
            devices_now = self.get_pnp_devices()
            new_devices = devices_now.keys() - self.devices_before.keys()
            removed_devices = self.devices_before.keys() - devices_now.keys()

            # Update the GUI if there are changes
            if new_devices or removed_devices:
                self.update_callback(new_devices, removed_devices)

            self.devices_before = devices_now

        pythoncom.CoUninitialize()  # Uninitialize COM

    def get_pnp_devices(self):
        wmi = win32com.client.GetObject("winmgmts:")
        pnp_devices = wmi.InstancesOf("Win32_PnPEntity")
        return {device.DeviceID: device.Name for device in pnp_devices}

    def stop(self):
        self.running = False


class USBDeviceLoggerGUI:
    def __init__(self, root):
        self.root = root
        root.title("USBDetective")

        # Configure the grid to expand the cell containing the Treeview widget
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)

        # Create a Treeview widget for logging device changes
        self.log = ttk.Treeview(root, columns=('DeviceID',))
        self.log.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        # Configure Treeview columns
        self.log.heading('#0', text='Status')
        self.log.column('#0', stretch=tk.YES)
        self.log.heading('DeviceID', text='Device ID')
        self.log.column('DeviceID', stretch=tk.YES)

        self.log.bind("<Double-1>", self.on_double_click)

        # Start the USB device detection in a separate thread
        self.detector = USBDeviceDetector(self.log_device_changes)
        self.detector.start()

    def log_device_changes(self, new_devices, removed_devices):
        for device_id in new_devices:
            self.log.insert('', 'end', text='Connected', values=(device_id,))
        for device_id in removed_devices:
            self.log.insert('', 'end', text='Disconnected', values=(device_id,))

    def on_double_click(self, event):
        item = self.log.selection()[0]
        device_id = self.log.item(item, 'values')[0]
        self.open_device_manager(device_id)

    def open_device_manager(self, device_id):
        pass

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