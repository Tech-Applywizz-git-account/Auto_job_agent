import os
import shutil
import time
import subprocess
from pathlib import Path
from playwright.sync_api import sync_playwright

def test_profile_launch():
    local_app_data = Path(os.environ["LOCALAPPDATA"])
    src_user_data = local_app_data / "Google" / "Chrome" / "User Data"
    profile_name = "Profile 9"
    src_profile = src_user_data / profile_name
    
    # Kill Chrome
    subprocess.run("taskkill /F /IM chrome.exe /T", shell=True, stderr=subprocess.DEVNULL)
    time.sleep(2)
    
    base_dir = Path(os.getcwd())
    temp_user_data = base_dir / "automation_context"
    if temp_user_data.exists():
        shutil.rmtree(temp_user_data, ignore_errors=True)
    temp_user_data.mkdir(parents=True, exist_ok=True)
    
    # To TRULY mimic the profile:
    # 1. Copy 'Local State' to temp_user_data root
    if (src_user_data / "Local State").exists():
        shutil.copy2(src_user_data / "Local State", temp_user_data / "Local State")
    
    # 2. Copy 'Profile 9' folder content to 'temp_user_data/Default'
    dest_default = temp_user_data / "Default"
    dest_default.mkdir(parents=True, exist_ok=True)
    
    print(f"Copying {profile_name} to temp context...")
    # Using more aggressive copy - everything in the profile
    for item in src_profile.iterdir():
        if item.name == "SingletonLock" or item.name == "SingletonSocket" or item.name == "SingletonCookie":
            continue
        try:
            if item.is_dir():
                shutil.copytree(item, dest_default / item.name, dirs_exist_ok=True)
            else:
                shutil.copy2(item, dest_default / item.name)
        except:
            pass
            
    print("Launch browser with temp context...")
    with sync_playwright() as p:
        # We don't specify --profile-directory because we renamed Profile 9 to Default
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(temp_user_data),
            executable_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            channel="chrome",
            headless=False,
            args=["--no-sandbox", "--disable-infobars"]
        )
        page = context.pages[0]
        page.goto("https://mail.google.com/")
        print("Browser is open. Please check if you are logged in.")
        time.sleep(20) # Stay open long enough for user to see
        context.close()

if __name__ == "__main__":
    test_profile_launch()
