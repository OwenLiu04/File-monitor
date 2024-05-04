# -*- coding: utf-8 -*-
"""
Author: OwenLiu04
Version: 0.1.0
License: GNU General Public License v3 (GPLv3)
Email: 69543538@qq.com

Description:
    This script monitors the modified date of files at user-defined time interval.
    If any file is modified, it opens the file automatically using PowerPoint.

Requirements:
    1. Operating system: Windows
    2. Microsoft PowerPoint (POWERPNT.EXE)
    3. Python Environment:
        - Python 3.11.7 (packaged by Anaconda, Inc.)

Initial Configuration:
    Users must specify the following at the start of the script:
    1. Paths and names of the files to be monitored.
    2. Monitoring interval, specified in hours, minutes, and seconds.
    3. Path to the PowerPoint executable (POWERPNT.EXE).
"""

import os
import time
import subprocess
import tkinter as tk
from threading import Thread

# Files to monitor
files_to_monitor = [
    r'full file path and file name',
    
    # Add more file paths as needed. Seperate them using ','.
]

# User-defined monitor interval in terms of hours, minutes, and seconds
monitor_interval_hours = 0
monitor_interval_minutes = 0
monitor_interval_seconds = 5

# Convert monitor interval to seconds
monitor_interval = (monitor_interval_hours * 3600) + (monitor_interval_minutes * 60) + monitor_interval_seconds

# Path to PowerPoint
powerpoint_path = r'path\POWERPNT.EXE'
# Input the file path of your POWERPNT.EXE

# Store the last modification times of the files
last_mod_times = {file: None for file in files_to_monitor}

def open_file_with_powerpoint(file_path, text_widget):
    """Open a file using Microsoft PowerPoint."""
    try:
        subprocess.Popen([powerpoint_path, '/R', file_path], shell=True)
        file_name = os.path.basename(file_path)
        info = f"{file_path} is modified and opened.\n"
        text_widget.insert(tk.END, info)
        text_widget.see(tk.END)  # Automatically scroll to the bottom
    except Exception as e:
        print(f"Failed to open {file_path}: {str(e)}")

def monitor_files(text_widget):
    """Monitor the files for any changes and print the modification time."""
    # Add information about the monitored files to the text window
    for file in files_to_monitor:
        file_info = f"Monitoring: {file}\n"
        text_widget.insert(tk.END, file_info)

    while True:
        for file in files_to_monitor:
            file_name = os.path.basename(file)
            info = f"{file_name}: Date modified: {time.ctime(os.path.getmtime(file))}"
            text_widget.insert(tk.END, info + "\n")
            text_widget.see(tk.END)  # Automatically scroll to the bottom
            if os.path.exists(file):
                current_mod_time = os.path.getmtime(file)
                if last_mod_times[file] is None:
                    last_mod_times[file] = current_mod_time
                elif current_mod_time != last_mod_times[file]:
                    # File has been modified
                    open_file_with_powerpoint(file, text_widget)
                    last_mod_times[file] = current_mod_time
            else:
                print(f"Warning: {file} not found.")
        time.sleep(monitor_interval)

def close_script_on_window_close(window):
    """Close the script when the main application window is closed."""
    def on_close():
        window.destroy()
        os._exit(0)  # Exit the script
    window.protocol("WM_DELETE_WINDOW", on_close)

def main():
    # Create the main application window
    root = tk.Tk()
    root.title("File Monitor")
    
    # Text window to display information
    text_widget = tk.Text(root, wrap=tk.WORD, height=20, width=50)
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # Scrollbars for the text window
    scrollbar = tk.Scrollbar(root, command=text_widget.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_widget.config(yscrollcommand=scrollbar.set)
    
    # Run the script in a separate thread
    script_thread = Thread(target=monitor_files, args=(text_widget,))
    script_thread.daemon = True  # Close the thread when the main thread exits
    script_thread.start()

    # Close the script when the main application window is closed
    close_script_on_window_close(root)

    root.mainloop()

if __name__ == "__main__":
    print("Starting file monitoring...")
    main()
