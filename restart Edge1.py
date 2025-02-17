import psutil
import time
import subprocess
import os
import ctypes

# Assign custom AppUserModelID
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("RestartEdge.02012025")

# Configuration
EDGE_PROCESS_NAME = "msedge.exe"
IDLE_THRESHOLD = 300  # Idle time in seconds
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# Get the active window's process name
def get_active_process_name():
    try:
        import win32gui
        import win32process

        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['pid'] == pid:
                return proc.info['name']
    except Exception as e:
        print(f"Error getting active process: {e}")
    return None

# Check if Edge is running
def is_edge_running():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == EDGE_PROCESS_NAME:
            return True
    return False

# Kill all instances of Edge
def kill_edge_processes():
    print("Restarting Microsoft Edge...")
    for proc in psutil.process_iter(['name', 'pid']):
        if proc.info['name'] == EDGE_PROCESS_NAME:
            try:
                proc.terminate()
                proc.wait(timeout=5)  # Wait for the process to terminate
                print(f"SUCCESS: Process {proc.info['name']} with PID {proc.info['pid']} has been terminated.")
            except Exception as e:
                print(f"Error terminating process {proc.info['pid']}: {e}")

# Restart Edge
def restart_edge():
    kill_edge_processes()
    time.sleep(2)  # Allow time for processes to close
    try:
        subprocess.Popen([EDGE_PATH, "--restore-last-session"])
        print("Microsoft Edge restarted with the previous session.")
    except Exception as e:
        print(f"Error restarting Edge: {e}")

# Main loop
last_interaction_time = time.time()
print(last_interaction_time)
while True:
    try:
        if is_edge_running():
            active_process = get_active_process_name()
            if active_process == EDGE_PROCESS_NAME:
                last_interaction_time = time.time()  # Reset idle timer
            else:
                idle_time = time.time() - last_interaction_time
                if idle_time > IDLE_THRESHOLD:
                    restart_edge()
                    last_interaction_time = time.time()
        else:
            print("Microsoft Edge is not running.")
            time.sleep(10)  # Check again after 10 seconds
    except Exception as e:
        print(f"Error: {e}")
    time.sleep(5)  # Check every 5 seconds
