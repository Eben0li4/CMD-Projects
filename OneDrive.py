import os
import time
import subprocess
from multiprocessing import Process
import signal
import sys

# Global variable to manage the running state
running = True

def pause_onedrive():
    subprocess.run([r"C:\Program Files\Microsoft OneDrive\OneDrive.exe", "/shutdown"], check=True)
    print("OneDrive sync paused.")

def resume_onedrive():
    subprocess.run([r"C:\Program Files\Microsoft OneDrive\OneDrive.exe", "/background"], check=True)
    print("OneDrive sync resumed.")

def change_sync_frequency(pause_duration, sync_duration):
    global running
    try:
        while running:
            print("Pausing OneDrive sync.")
            pause_onedrive()
            time.sleep(pause_duration)
            
            print("Resuming OneDrive sync.")
            resume_onedrive()
            time.sleep(sync_duration)
    except KeyboardInterrupt:
        print("Sync frequency adjustment stopped by user.")

def run_in_background(pause_duration, sync_duration, timeout=None):
    def signal_handler(sig, frame):
        global running
        running = False
        print('Stopping the sync frequency adjustment...')
        sys.exit(0)

    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    process = Process(target=change_sync_frequency, args=(pause_duration, sync_duration))
    process.start()

    if timeout:
        time.sleep(timeout)
        running = False
        process.terminate()

    return process

# Example usage: Pause sync for 20 minutes, sync for 10 minutes
pause_duration = 0.1 * 60  # 20 minutes in seconds
sync_duration = 1 * 60   # 10 minutes in seconds
timeout = 60 * 60  # Run for 1 hour (optional)

if __name__ == "__main__":
    process = run_in_background(pause_duration, sync_duration, timeout)
    print("OneDrive sync frequency adjustment running in background.")
    process.join()
