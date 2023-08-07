import psutil
import time
import pygetwindow as gw
import ctypes
from tqdm import tqdm
from pynput.keyboard import Key, Controller

# Get all windows
windows = gw.getAllWindows()

# Print a list of windows for the user to select
for i, window in enumerate(windows):
    print(f"{i+1}: {window.title}")

# Ask the user to select a window
selection = int(input("Select a window by its number: ")) - 1
if 0 <= selection < len(windows):
    selected_window = windows[selection]
    print(f"You selected: {selected_window.title}")
else:
    print("Invalid selection.")
    exit()

# Get the process ID of the selected window
thread_id, process_id = ctypes.wintypes.DWORD(), ctypes.wintypes.DWORD()
ctypes.windll.user32.GetWindowThreadProcessId(selected_window._hWnd, ctypes.pointer(process_id))
print(f"Process ID of the selected window: {process_id.value}")

# Ask the user to input the autosave interval in minutes
interval = int(input("Autosave interval in min: "))
interval_in_sec = interval * 60

# Calculate number of 5 seconds intervals for the progress bar
interval_5_sec = interval_in_sec // 5
remainder = interval_in_sec % 5

keyboard = Controller()

# Check every interval if the selected process is still running and the window is active, if so press F5
while True:
    for _ in tqdm(range(interval_5_sec), desc="Waiting...", bar_format="{l_bar}{bar}| {remaining}"):
        time.sleep(5)
    try:
        process = psutil.Process(process_id.value)
        if process.status() == 'running':
            active_hwnd = gw.getActiveWindow()._hWnd
            _, active_pid = ctypes.wintypes.DWORD(), ctypes.wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(active_hwnd, ctypes.pointer(active_pid))
            if process_id.value == active_pid.value:
                keyboard.press(Key.f5)
                keyboard.release(Key.f5)
                print(f"Auto save {time.strftime('%H:%M',time.localtime())} Uhr")
            else:
                print("Selected window is not active.")
        else:
            print("Selected process is stopped")
            break
    except psutil.NoSuchProcess:
        print(f"Process {process_id.value} no longer exists")
        exit(0)
    time.sleep(remainder)