import os
import shutil
import time
import csv
from pathlib import Path
from playwright.sync_api import sync_playwright

def run():
    base_dir = Path(r"c:\Users\DELL\Desktop\DINESH-AGENTIC")
    local_app_data = Path(os.environ["LOCALAPPDATA"])
    user_data_path = local_app_data / "Google" / "Chrome" / "User Data"
    profile_name = "Profile 9"
    
    # Target profile source
    src_profile = user_data_path / profile_name
    
    # We copy essentials to a temp folder to avoid profile locks if possible
    # though using it directly is better if Chrome is closed.
    temp_profile_dir = base_dir / "temp_profile"
    if not temp_profile_dir.exists():
        temp_profile_dir.mkdir(parents=True, exist_ok=True)

    # Let's try to just use it directly, but kill Chrome first.
    import subprocess
    subprocess.run("taskkill /F /IM chrome.exe /T", shell=True, stderr=subprocess.DEVNULL)
    time.sleep(2)

    with sync_playwright() as p:
        # Launching with user profile
        try:
            print(f"Launching Chrome with profile: {profile_name}")
            # Note: Playwright needs the path to the ROOT of User Data, and then you specify the profile directory name.
            # However, launch_persistent_context 'user_data_dir' must BE the profile dir itself?
            # Actually, standard chromium launch doesn't support 'profile-directory' well in launch_persistent_context easily.
            # Best way: Copy Profile 9 to a temp dir and use that as the user_data_dir.
            
            # Simple copy logic (just essentials to speed up)
            # Actually, the user says "Keep the session active", so we need the full profile or at least session files.
            
            # Let's try pointing to the actual profile if no one is using it.
            # But Playwright launch_persistent_context expects a clean dir or it might corrupt things.
            # Better: use the user_data_dir as the PARENT of the profiles and specify the profile-directory flag.
            
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(user_data_path),
                executable_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                channel="chrome",
                headless=False,
                args=[
                    f"--profile-directory={profile_name}",
                    "--no-sandbox",
                    "--disable-infobars",
                ],
                ignore_default_args=["--disable-extensions"]
            )
            
            page = context.pages[0] if context.pages else context.new_page()
            
            # Load links
            links = []
            with open(base_dir / "links.csv", mode='r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    url = row.get('url') or row.get('link') or list(row.values())[0]
                    if url and url.startswith('http'):
                        links.append(url.strip())
            
            if not links:
                print("No links found in CSV.")
                return

            for url in links:
                print(f"Navigating to {url}")
                page.goto(url, wait_until="load", timeout=60000)
                
                # Step 4: Wait until page fully loads (3-5 seconds as requested)
                time.sleep(4)
                
                # Step 5: Identify and click "Apply" button
                # Step 8: Add delay before click (2-3 seconds)
                time.sleep(2.5)
                
                # Try common Apply selectors
                apply_selectors = [
                    "button:has-text('Apply')", 
                    "a:has-text('Apply')", 
                    "role=button[name='Apply']",
                    "id=apply_button",
                    ".apply-button"
                ]
                
                found_apply = False
                for sel in apply_selectors:
                    elem = page.locator(sel).first
                    if elem.is_visible(timeout=5000):
                        print(f"Found Apply button: {sel}")
                        
                        # Step 6: Move cursor manually (simulate real behavior)
                        box = elem.bounding_box()
                        if box:
                            target_x = box['x'] + box['width'] / 2
                            target_y = box['y'] + box['height'] / 2
                            page.mouse.move(target_x, target_y, steps=20)
                            time.sleep(0.5)
                            page.mouse.click(target_x, target_y)
                            found_apply = True
                            print("Clicked Apply.")
                            break
                            
                if not found_apply:
                    print("Apply button not found, skipping...")
                    continue

                # Step 6: a. Wait for next page or popup
                time.sleep(5)
                
                # Step 7 & 6c: Click extension icon
                # The prompt implies the extension icon is VISIBLE and we should move cursor to it.
                # If it's a floating 'AW' icon from ApplyWizz:
                print("Looking for extension icon/button...")
                ext_selectors = [
                    "div[class*='applywizz']", 
                    "div[id*='applywizz']",
                    "button[class*='applywizz']",
                    "shadow-root", # Some extensions use shadow DOM
                    "canvas",
                    "img[src*='logo']",
                    "img[src*='icon']",
                    "[aria-label*='ApplyWizz']"
                ]
                
                # Let's wait a bit more for the extension to inject
                time.sleep(3)
                
                found_ext = False
                for sel in ext_selectors:
                    # Some selectors might be in Shadow DOM?
                    # Playwright handles shadow roots with >> syntax if needed, but often automatically finds it.
                    # We'll try to find any visible element with these selectors.
                    elems = page.locator(sel).all()
                    for e in elems:
                        if e.is_visible(timeout=1000):
                            box = e.bounding_box()
                            if box and box['width'] > 0:
                                print(f"Found something likely extension related: {sel}")
                                # Move cursor manually
                                page.mouse.move(box['x'] + box['width'] / 2, box['y'] + box['height'] / 2, steps=25)
                                time.sleep(2) # 2 seconds before clicking extension as requested
                                page.mouse.click(box['x'] + box['width'] / 2, box['y'] + box['height'] / 2)
                                found_ext = True
                                print("Clicked Extension Icon.")
                                break
                    if found_ext: break
                
                if not found_ext:
                    print("Extension icon not found.")
                
                # Pause to see results for the user (optional but helpful if debugging)
                # input("Press Enter to continue to next link...")
                time.sleep(3)

            context.close()
        except Exception as e:
            print(f"An error occurred: {e}")

if __name__ == "__main__":
    run()
