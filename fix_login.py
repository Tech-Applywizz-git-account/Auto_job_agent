import os
import subprocess
from pathlib import Path
import time

def open_original_profile():
    local_app_data = Path(os.environ["LOCALAPPDATA"])
    user_data_root = local_app_data / "Google" / "Chrome" / "User Data"
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    
    # We open the ORIGINAL Profile 6 (applywizzportfolios@gmail.com)
    # This allows you to log in ONCE and have it stay logged in for the automation.
    profile_name = "Profile 6"
    
    print(f"--- LOGIN MAINTENANCE ---")
    print(f"Opening original Chrome profile: {profile_name}")
    print(f"1. Click 'Verify it's you' in the top right.")
    print(f"2. Complete the login process.")
    print(f"3. Close the browser when done.")
    print(f"--------------------------")

    chrome_cmd = [
        chrome_path,
        f"--profile-directory={profile_name}",
        f"--user-data-dir={user_data_root}",
        "--no-first-run"
    ]
    
    try:
        subprocess.Popen(chrome_cmd)
        print("\nChrome opened. Please log in now...")
    except Exception as e:
        print(f"Error opening Chrome: {e}")

if __name__ == "__main__":
    open_original_profile()
