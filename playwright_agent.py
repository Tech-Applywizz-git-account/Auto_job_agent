import csv
import os
import subprocess
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

def force_kill_chrome():
    print("Closing Chrome completely...")
    subprocess.run("taskkill /F /IM chrome.exe /T", shell=True, stderr=subprocess.DEVNULL)
    time.sleep(2)

def run_agent():
    # Configuration
    base_dir = Path(r"e:\AUTO-JOB-AGENTIC")
    csv_file = base_dir / "links.csv"
    
    local_app_data = Path(os.environ["LOCALAPPDATA"])
    user_data_root = local_app_data / "Google" / "Chrome" / "User Data"
    
    profile_name = "Profile 6" # applywizzportfolios@gmail.com
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    debug_port = 9222

    if not csv_file.exists():
        return

    # Read Links
    targets = []
    with open(csv_file, mode='r', newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            link = row.get('url') or row.get('link') or list(row.values())[0]
            if link and link.startswith('http'):
                targets.append(link.strip())

    if not targets:
        return

    force_kill_chrome()

    for index, url in enumerate(targets, 1):
        # 1. NATIVE LAUNCH: This is the ONLY way to guarantee the correct profile + link
        # We pass the URL in the command line itself.
        print(f"\n[{index}/{len(targets)}] NATIVE LAUNCH: Profile 9 + {url}")
        
        chrome_cmd = [
            chrome_path,
            f"--profile-directory={profile_name}",
            f"--user-data-dir={user_data_root}",
            f"--remote-debugging-port={debug_port}",
            "--remote-allow-origins=*",
            url
        ]
        
        try:
            # Launch natively
            subprocess.Popen(chrome_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print("Chrome opened with the link. Waiting for initialization...")
            time.sleep(15) # Give it plenty of time to load the profile and link
            
            # 2. CONNECT FOR CLICKING
            with sync_playwright() as p:
                print("Connecting Agent for automation...")
                try:
                    browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{debug_port}")
                    context = browser.contexts[0]
                    # Since we opened the link in the command line, it's already there!
                    # Find the page that has the job URL
                    page = None
                    for p_iter in context.pages:
                        if url in p_iter.url or "greenhouse" in p_iter.url or "lyft" in p_iter.url:
                            page = p_iter
                            break
                    
                    if not page:
                        page = context.pages[0] if context.pages else context.new_page()

                    # CLICK AUTOMATION
                    print("Automating 3-step interaction...")
                    time.sleep(5)
                    
                    # 1. Page Apply Button
                    try:
                        print("Clicking webpage 'Apply'...")
                        page.get_by_role("button", name="Apply").first.click(timeout=8000)
                    except:
                        try:
                            page.locator("button:has-text('Apply'), a:has-text('Apply')").first.click(timeout=5000)
                        except: pass
                    
                    # 2. AW Extension Logo
                    time.sleep(7)
                    print("Clicking ApplyWizz logo...")
                    try:
                        aw_btn = page.locator("[class*='applywizz'], [id*='applywizz'], canvas").first
                        aw_btn.click(force=True, timeout=10000)
                        
                        # 3. Scan Application
                        time.sleep(3)
                        print("Clicking 'Scan Application'...")
                        page.get_by_text("Scan Application", exact=False).first.click(timeout=8000)
                        print("==> SCAN TRIGGERED.")
                    except:
                        print("Extension interaction failed. You can click it manually.")

                    input("\nJOB PROCESSED. Press Enter to proceed or finish...")
                    browser.close()
                    
                except Exception as e:
                    print(f"CDP Connection failed: {e}")
                    print("However, the browser is open with the link and profile.")
                    input("\nLink is OPEN with your account. Press Enter here when finished with this job...")
            
            # Close Chrome before next link
            force_kill_chrome()
            
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    run_agent()
