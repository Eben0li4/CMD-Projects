import subprocess
import time

def pause_onedrive():
    # Pause OneDrive
    subprocess.run(["powershell", "-Command", "& {Start-Process -NoNewWindow -FilePath 'C:\\Program Files\\Microsoft OneDrive\\OneDrive.exe' -ArgumentList '/pause'}"])

def resume_onedrive():
    # Resume OneDrive
    subprocess.run(["powershell", "-Command", "& {Start-Process -NoNewWindow -FilePath 'C:\\Program Files\\Microsoft OneDrive\\OneDrive.exe' -ArgumentList '/resume'}"])

if __name__ == "__main__":
    # Pause OneDrive
    print("Pausing OneDrive...")
    pause_onedrive()

    # Wait for a specific time (e.g., 10 seconds)
    time.sleep(10)

    # Resume OneDrive
    print("Resuming OneDrive...")
    resume_onedrive()
