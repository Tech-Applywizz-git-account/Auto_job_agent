import os
import shutil
import time
import subprocess
from pathlib import Path
from playwright.sync_api import sync_playwright

def test_profile_launch_v2():
    local_app_data = Path(os.environ["LOCALAPPDATA"])
    src_user_data = local_app_data / "Google" / "Chrome" / "User Data"
    profile_name = "Profile 9"
    src_profile = src_user_data / profile_name
    
    # Kill Chrome
    subprocess.run("taskkill /F /IM chrome.exe /T", shell=True, stderr=subprocess.DEVNULL)
    time.sleep(2)
    
    base_dir = Path(os.getcwd())
    temp_user_data = base_dir / "automation_context_v2"
    if temp_user_data.exists():
        shutil.rmtree(temp_user_data, ignore_errors=True)
    temp_user_data.mkdir(parents=True, exist_ok=True)
    
    # Root: Copy Local State
    if (src_user_data / "Local State").exists():
        shutil.copy2(src_user_data / "Local State", temp_user_data / "Local State")
    
    # Folder: Create Profile 9 and copy content
    dest_profile = temp_user_data / "Profile 9"
    dest_profile.mkdir(parents=True, exist_ok=True)
    
    print(f"Syncing Profile 9 structure...")
    # List of essential files for a profile to be recognized correctly
    essentials = ["Cookies", "Preferences", "Login Data", "Web Data", "Extension Cookies", "Secure Preferences"]
    folders = ["Local Storage", "Extension State", "Extensions", "IndexedDB", "Network", "Sync Data"]
    
    for f in essentials:
        if (src_profile / f).exists(): shutil.copy2(src_profile / f, dest_profile / f)
    for fol in folders:
        if (src_profile / fol).exists(): shutil.copytree(src_profile / fol, dest_profile / fol, dirs_exist_ok=True)
            
    print("Launching Chrome with structural profile...")
    with sync_playwright() as p:
        # We point to the parent of Profile 9 and use --profile-directory
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(temp_user_data),
            executable_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            channel="chrome",
            headless=False,
            args=[
                f"--profile-directory={profile_name}",
                "--no-sandbox",
                "--disable-infobars"
            ],
            ignore_default_args=["--disable-extensions"]
        )
        page = context.pages[0]
        page.goto("https://mail.google.com/")
        print("Waiting for check...")
        time.sleep(15) 
        context.close()

if __name__ == "__main__":
    test_profile_launch_v2()
