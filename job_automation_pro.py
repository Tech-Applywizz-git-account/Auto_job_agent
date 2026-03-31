import csv
import os
import subprocess
import shutil
import time
import random
import json
from pathlib import Path
from playwright.sync_api import sync_playwright
from datetime import datetime
import requests

def send_teams_notification(url, job_data):
    """Sends job application status to Microsoft Teams via Power Automate."""
    if not url:
        return
    try:
        # Construct a clean HTML table for Teams (more reliable for rendering)
        html_table = "<table border='1' style='border-collapse: collapse; width: 100%; font-family: Calibri, sans-serif;'>"
        html_table += "<tr style='background-color: #4CAF50; color: white;'><th style='padding: 8px; text-align: left;'>Field</th><th style='padding: 8px; text-align: left;'>Value</th></tr>"
        
        for key, value in job_data.items():
            # Skip the 'details' object to keep the table clean
            if key == "details": continue
            
            clean_key = key.replace('_', ' ').title()
            html_table += f"<tr><td style='padding: 8px; border: 1px solid #ddd; font-weight: bold;'>{clean_key}</td><td style='padding: 8px; border: 1px solid #ddd;'>{value}</td></tr>"
        
        html_table += "</table>"
        
        # Prepare payload and headers to match the new schema and authorized sender
        payload = {
            "project": "job",
            "type": "job",
            "message": html_table
        }
        
        headers = {
            "X-Authorized-Sender": "nikhil@applywizz.com",
            "Content-Type": "application/json"
        }
        
        response = requests.post(url, json=payload, headers=headers, timeout=10)
        
        if response.status_code in [200, 202]:
            print(f"Teams Notification sent for {job_data.get('url')} (Status: {response.status_code})")
        else:
            print(f"Teams Notification failed with status: {response.status_code}")
    except Exception as e:
        print(f"Error sending Teams notification: {e}")

def human_move_and_click(page, locator, delay_before=2):
    """Moves the mouse in a human-like way to an element and clicks it."""
    try:
        if locator.is_visible(timeout=5000):
            print(f"Moving cursor to button...")
            time.sleep(delay_before)
            
            # Using locator.hover() instead of manual page.mouse.move() 
            # because Playwright handles iframe offsets automatically with locators.
            locator.hover()
            time.sleep(random.uniform(0.2, 0.5))
            
            # Use locator.click() with a human-like delay between mouse-down and mouse-up
            locator.click(delay=random.randint(50, 150))
            return True
    except Exception as e:
        print(f"Human interaction failed: {e}")
        # Final fallback: Force click
        try:
            locator.click(force=True)
            return True
        except:
            return False
    return False

def sync_profile(src_profile_path, target_root, profile_name):
    """Prepares a FULL copy of the profile to maintain login/verified status."""
    print(f"Syncing FULL profile from {src_profile_path} to {target_root}...")
    
    if target_root.exists():
        # Try to remove old files, but don't fail if some are locked
        try: shutil.rmtree(target_root, ignore_errors=True)
        except: pass
    target_root.mkdir(parents=True, exist_ok=True)
    
    # Root: Copy Local State (CRITICAL for profile recognition)
    local_state = src_profile_path.parent / "Local State"
    if local_state.exists():
        shutil.copy2(local_state, target_root / "Local State")
    
    # Copy the ENTIRE profile folder
    dest_profile = target_root / profile_name
    try:
        shutil.copytree(src_profile_path, dest_profile, dirs_exist_ok=True)
    except Exception as e:
        print(f"Warning during full copy: {e}")
    
    print("Full profile sync complete.")

def run_application_loop():
    # --- CONFIGURATION ---
    SEND_TEAMS_NOTIFICATIONS = True  # Set to False to skip Teams messages
    # ---------------------
    
    base_dir = Path(os.getcwd())
    csv_file = base_dir / "links.csv"
    temp_profile_dir = base_dir / "automation_profile"
    
    local_app_data = Path(os.environ["LOCALAPPDATA"])
    chrome_data_root = local_app_data / "Google" / "Chrome" / "User Data"
    source_profile = chrome_data_root / "Profile 6"
    chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    
    # Teams Integration URL
    teams_url = "https://defaultdd60b0661b78451584fba565c251cb.5a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/793b96ec34ab4d49899c8d479a84bc01/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-5L-m5tZA2Yd6lKGHryMzJA7cNEQhXKZqKOPSZti7IE"

    if not csv_file.exists():
        print(f"Link CSV not found at {csv_file}")
        return

    # 1. Kill Chrome to release locks
    print("Closing Chrome...")
    subprocess.run("taskkill /F /IM chrome.exe /T", shell=True, stderr=subprocess.DEVNULL)
    time.sleep(2)

    # 2. Sync FULL profile to maintain verified status
    sync_profile(source_profile, temp_profile_dir, "Profile 6")

    # 3. Read links line-by-line (robust against header/no-header)
    links = []
    with open(csv_file, mode='r', encoding='utf-8') as f:
        for line in f:
            url = line.strip()
            # Skip common headers
            if url.lower() in ["url", "link", "job link"]: continue
            if url and url.startswith('http'):
                links.append(url)
    
    if not links:
        print("No valid job links found in CSV.")
        return

    print(f"Found {len(links)} links to process.")

    all_analytics = []
    overall_start_time = datetime.now()
    
    with sync_playwright() as p:
        try:
            print(f"Launching Browser with FULL profile copy: {temp_profile_dir}")
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(temp_profile_dir),
                executable_path=chrome_exe,
                channel="chrome",
                headless=False,
                args=[
                    f"--profile-directory=Profile 6",
                    "--no-sandbox",
                    "--disable-infobars",
                    "--start-maximized",
                    "--no-first-run",
                    "--no-default-browser-check"
                ],
                ignore_default_args=["--disable-extensions"],
                timeout=60000
            )
            
            # Robust page connection
            print("Connecting to page...")
            page = None
            start_time = time.time()
            while time.time() - start_time < 30:
                if context.pages:
                    page = context.pages[0]
                    break
                time.sleep(1)
            
            if not page:
                print("No initial page found, creating new page...")
                page = context.new_page()
            
            page.set_default_timeout(60000)
            print("Browser ready. Starting job applications...")
            page.bring_to_front()
            
            for index, url in enumerate(links):
                # Open each link in a NEW TAB as requested
                page = context.new_page()
                page.set_default_timeout(60000)
                
                print(f"\n[{index+1}/{len(links)}] Processing: {url}")
                
                link_analytics = {
                    "automation_start_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "url": url,
                    "link_open_time": None,
                    "extension_icon_clicked_time": None,
                    "scan_application_clicked_time": None,
                    "go_to_panel_clicked_time": None,
                    "panel_closed_time": None,
                    "screenshot_before_time": None,
                    "submit_clicked_time": None,
                    "screenshot_after_time": None,
                    "total_link_processing_time": None
                }
                
                link_start_time = datetime.now()

                try:
                    # Open Job page
                    page.goto(url, wait_until="domcontentloaded", timeout=60000)
                    print("Page loaded.")
                    link_analytics["link_open_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # 4. Wait until page fully loads (3-5 seconds)
                    delay = random.uniform(3, 5)
                    print(f"Waiting {delay:.2f}s for load...")
                    time.sleep(delay)
                    
                    # 5. Click Apply Button
                    print("Searching for Apply button...")
                    apply_selectors = [
                        "button:has-text('Apply')", 
                        "a:has-text('Apply')", 
                        "role=button[name='Apply']",
                        "text='Apply for this job'",
                        "input[value='Apply']",
                        "[data-testid*='apply']",
                        "[aria-label*='Apply']",
                        ".apply-button"
                    ]
                    
                    clicked_apply = False
                    for selector in apply_selectors:
                        btn = page.locator(selector).first
                        if btn.is_visible(timeout=5000):
                            print(f"Found Apply button: {selector}")
                            # Delay before clicking Apply (2-3 seconds)
                            clicked_apply = human_move_and_click(page, btn, delay_before=random.uniform(2, 3))
                            if clicked_apply:
                                print("Apply button clicked.")
                                break
                    
                    if not clicked_apply:
                        print("Apply button not found or already on form. Proceeding to extension check...")
                    
                    # 6a. Wait for next page or popup
                    print("Waiting for extension widget/next page (5s)...")
                    time.sleep(5)
                    
                    # 6c. Click on browser extension icon (ApplyWizz)
                    print("Searching for ApplyWizz extension widget icon (including frames)...")
                    ext_selectors = [
                        "img[src*='applywizz']",
                        "img[alt*='ApplyWizz']",
                        "img[alt*='Autofill']",
                        "img[src*='icon']",
                        "div[class*='applywizz']",
                        "div[id*='applywizz']",
                        "button[class*='applywizz']",
                        "[aria-label*='ApplyWizz']",
                        "[data-testid*='applywizz']",
                        "div[style*='z-index']",
                        "img[src*='logo']", # Low priority fallback
                        "canvas"
                    ]
                    
                    clicked_scan_success = False
                    
                    # Get all candidate elements first
                    candidates = []
                    
                    def collect_candidates(root_or_frame, frame_ref=None):
                        for sel in ext_selectors:
                            try:
                                elems = root_or_frame.locator(sel).all()
                                for e in elems:
                                    if e.is_visible(timeout=500):
                                        # Exclude Greenhouse
                                        html = e.evaluate("el => el.outerHTML").lower()
                                        if "greenhouse" in html or "job-board" in html:
                                            continue
                                            
                                        box = e.bounding_box()
                                        if box and 20 < box['width'] < 120 and 20 < box['height'] < 120:
                                            candidates.append((e, sel, frame_ref or page))
                            except: continue

                    collect_candidates(page)
                    for f in page.frames:
                        collect_candidates(f, f)
                    
                    print(f"Found {len(candidates)} candidate extension icons.")
                    
                    for i, (ext_elem, sel, root) in enumerate(candidates):
                        print(f"Trying candidate {i+1} with selector: {sel}")
                        if human_move_and_click(page, ext_elem, delay_before=2):
                            link_analytics["extension_icon_clicked_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            print("Clicked candidate. Waiting to see if Scan Application button appears...")
                            time.sleep(4)
                            
                            scan_selectors = [
                                "button:has-text('Scan Application')",
                                "text='Scan Application'",
                                "div:has-text('Scan Application')",
                                "[aria-label*='Scan Application']",
                                ".green-btn",
                                "button.Scan_Application"
                            ]
                            
                            # Check all frames for scan button
                            found_scan = False
                            for frame in page.frames:
                                for scan_sel in scan_selectors:
                                    try:
                                        scan_btn = frame.locator(scan_sel).first
                                        if scan_btn.is_visible(timeout=1000):
                                            print(f"SUCCESS: Found Scan Application button in frame '{frame.name or 'main'}' after clicking icon.")
                                            if human_move_and_click(page, scan_btn, delay_before=2):
                                                link_analytics["scan_application_clicked_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                print("Clicked Scan Application button successfully.")
                                                found_scan = True
                                                break
                                    except: continue
                                if found_scan: break
                            
                            if found_scan:
                                clicked_scan_success = True
                                break
                            else:
                                print("Scan button not found after this icon. Trying next candidate...")
                                # Maybe the widget needs to be closed? Usually not.
                        
                    if clicked_scan_success:
                        print("Extension is processing application data...")
                        # Dynamic polling for "Go to panel" button
                        found_go_to_panel = False
                        polling_start = time.time()
                        timeout = 180 # 3 minutes timeout
                        
                        while time.time() - polling_start < timeout:
                            # Check all frames for "Go to panel"
                            for frame in page.frames:
                                panel_selectors = [
                                    "button:has-text('Go to panel')",
                                    "text='Go to panel'",
                                    "div:has-text('Go to panel')",
                                    "[aria-label*='Go to panel']"
                                ]
                                for p_sel in panel_selectors:
                                    try:
                                        panel_btn = frame.locator(p_sel).first
                                        if panel_btn.is_visible(timeout=500):
                                            print(f"SUCCESS: Found 'Go to panel' button in frame '{frame.name or 'main'}'.")
                                            if human_move_and_click(page, panel_btn, delay_before=1):
                                                link_analytics["go_to_panel_clicked_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                print("Clicked 'Go to panel' button. Extension work complete.")
                                                found_go_to_panel = True
                                                
                                                # Close extension panel by clicking 'x'
                                                print("Searching for extension panel close ('x') button...")
                                                close_selectors = [
                                                     "button[aria-label='Close']",
                                                     "button:has-text('Close')",
                                                     ".close-icon",
                                                     "button .icon-close",
                                                     "span.close",
                                                     "a.close",
                                                     "button[class*='close']",
                                                     "button[aria-label*='close']",
                                                     "button[aria-label='Minimise']",
                                                     "button:has-text('-')",
                                                     "[data-testid*='close']",
                                                     "svg[class*='close']"
                                                 ]
                                                 
                                                closed_panel = False
                                                for frame_c in page.frames:
                                                    for c_sel in close_selectors:
                                                        try:
                                                            close_btn = frame_c.locator(c_sel).first
                                                            if close_btn.is_visible(timeout=500):
                                                                print(f"Found close button: {c_sel}")
                                                                # Try human click first, then force click if it fails
                                                                if human_move_and_click(page, close_btn, delay_before=1):
                                                                    # Wait for panel to disappear
                                                                    time.sleep(1)
                                                                else:
                                                                    print("Standard close click failed. Force-clicking close button.")
                                                                    close_btn.click(force=True)
                                                                    
                                                                link_analytics["panel_closed_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                                print("Extension panel closed.")
                                                                closed_panel = True
                                                                break
                                                        except: continue
                                                    if closed_panel: break
                                                
                                                if not closed_panel:
                                                    print("Could not find extension close button. DEBUGGING: Listing visible buttons in frames:")
                                                    for frame in page.frames:
                                                        try:
                                                            btns = frame.locator("button, [role='button']").all()
                                                            for b in btns:
                                                                if b.is_visible(timeout=100):
                                                                    print(f"  Frame '{frame.name}': Text='{b.inner_text()}', ID='{b.get_attribute('id')}', Class='{b.get_attribute('class')}'")
                                                        except: pass
                                                    print("Proceeding anyway...")
                                                
                                                time.sleep(2)
                                                
                                                # Take screenshot after clicking "Go to panel"
                                                print("Waiting for panel to load for screenshot...")
                                                time.sleep(5)
                                                
                                                timestamp = time.strftime("%Y%m%d_%H%M%S")
                                                os.makedirs("screenshot_before", exist_ok=True)
                                                filename = f"screenshot_before/screenshot_{timestamp}.png"
                                                
                                                try:
                                                    page.screenshot(path=filename, full_page=True)
                                                    link_analytics["screenshot_before_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                    print(f"Screenshot (Before) saved to {filename}")
                                                except Exception as ss_err:
                                                    print(f"Failed to take screenshot: {ss_err}")
                                                
                                                # Final step: Click "Submit Application" button
                                                print("Searching for Submit Application button...")
                                                submit_selectors = [
                                                    "button#submit_app",
                                                    "input[type='submit']",
                                                    "button:has-text('Submit Application')",
                                                    "button:has-text('Submit')",
                                                    "text='Submit Application'",
                                                    "text='Submit'",
                                                    "[data-testid*='submit']",
                                                    "[aria-label*='Submit']"
                                                ]
                                                
                                                clicked_submit = False
                                                # Check main page for submit button
                                                for sub_sel in submit_selectors:
                                                     try:
                                                         sub_btn = page.locator(sub_sel).first
                                                         if sub_btn.is_visible(timeout=5000):
                                                             print(f"Found Submit button on main page: {sub_sel}")
                                                             # Centering the element to avoid sticky headers/footers
                                                             sub_btn.scroll_into_view_if_needed()
                                                             time.sleep(1)
                                                             
                                                             if human_move_and_click(page, sub_btn, delay_before=2):
                                                                 link_analytics["submit_clicked_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                                 print("SUCCESS: Submit Application button clicked!")
                                                                 clicked_submit = True
                                                                 break
                                                             else:
                                                                 print("Human click failed. Attempting force-click...")
                                                                 sub_btn.click(force=True)
                                                                 # Ensure the click event is dispatched if physical click is obscured
                                                                 sub_btn.dispatch_event('click')
                                                                 link_analytics["submit_clicked_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                                 print("SUCCESS: Submit Application button clicked (Force + Event Dispatched)!")
                                                                 clicked_submit = True
                                                                 break
                                                     except: continue
                                                
                                                # If not found on main page, check frames as a fallback
                                                if not clicked_submit:
                                                    for frame in page.frames:
                                                        for sub_sel in submit_selectors:
                                                            try:
                                                                sub_btn = frame.locator(sub_sel).first
                                                                if sub_btn.is_visible(timeout=1000):
                                                                    print(f"Found Submit button in frame '{frame.name or 'iframe'}': {sub_sel}")
                                                                    sub_btn.scroll_into_view_if_needed()
                                                                    time.sleep(1)
                                                                    
                                                                    if human_move_and_click(page, sub_btn, delay_before=2):
                                                                        link_analytics["submit_clicked_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                                        print("SUCCESS: Submit Application button clicked in frame!")
                                                                        clicked_submit = True
                                                                        break
                                                                    else:
                                                                        print("Attempting force click in frame...")
                                                                        sub_btn.click(force=True)
                                                                        link_analytics["submit_clicked_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                                        print("SUCCESS: Submit Application button force-clicked in frame!")
                                                                        clicked_submit = True
                                                                        break
                                                            except: continue
                                                        if clicked_submit: break
                                                
                                                if clicked_submit:
                                                    print("Waiting 10s for submission to process and checking results...")
                                                    time.sleep(10)
                                                    
                                                    # POST-SUBMISSION VERIFICATION
                                                    status_detected = "Processed"
                                                    error_msg = "None"
                                                    
                                                    try:
                                                        # Case 1: Success Detection
                                                        success_keywords = ["Thanks", "Thank", "Thank you", "Submitted", "Success", "Received", "Complete"]
                                                        page_text = page.content()
                                                        found_success = any(kw.lower() in page_text.lower() for kw in success_keywords)
                                                        
                                                        # Case 2: Captcha / Security Detection
                                                        captcha_keywords = ["Captcha", "Security Code", "Verify", "Human", "Verification"]
                                                        found_captcha = any(kw.lower() in page_text.lower() for kw in captcha_keywords)
                                                        
                                                        # Case 3: Error Detection
                                                        error_keywords = ["Error", "Failed", "Invalid", "Required"]
                                                        found_error = any(kw.lower() in page_text.lower() for kw in error_keywords)
                                                        
                                                        if found_success:
                                                            status_detected = "success"
                                                            error_msg = "None"
                                                            print("✅ SUCCESS: Confirmation message detected on page.")
                                                        elif found_captcha:
                                                            status_detected = "human assistance required"
                                                            error_msg = "Captcha or Security Code detected on page."
                                                            print("⚠️ CAPTCHA DETECTED: Human assistance required.")
                                                        elif found_error:
                                                            status_detected = "Failure"
                                                            error_msg = "Explicit error message detected on page after submission."
                                                            print("❌ FAILURE: Error message found on page.")
                                                        else:
                                                            # Fallback: if URL changed, consider it a success if not otherwise errored
                                                            if url not in page.url:
                                                                status_detected = "success"
                                                                print("✅ SUCCESS: URL changed (presumed submission).")
                                                            else:
                                                                status_detected = "Failure"
                                                                error_msg = "Application stuck on form page (no success message or URL change after 10s wait)."
                                                                print("❌ FAILURE: Page looks unchanged after submission.")
                                                    except Exception as check_err:
                                                        print(f"Warning: Results verification failed: {check_err}")
                                                    
                                                    link_analytics["status"] = status_detected
                                                    link_analytics["error_message"] = error_msg
                                                    
                                                    # Take post-submission screenshot
                                                    print("Taking post-submission screenshot...")
                                                    screenshot_after_dir = base_dir / "screenshot_after"
                                                    screenshot_after_dir.mkdir(exist_ok=True)
                                                    
                                                    timestamp_after = time.strftime("%Y%m%d_%H%M%S")
                                                    screenshot_after_path = screenshot_after_dir / f"screenshot_{timestamp_after}.png"
                                                    
                                                    try:
                                                        page.screenshot(path=str(screenshot_after_path), full_page=True)
                                                        link_analytics["screenshot_after_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                        print(f"Post-submission screenshot saved to {screenshot_after_path} (Full Page)")
                                                    except Exception as ss_err:
                                                        print(f"Failed to take post-submission screenshot: {ss_err}")
                                                break
                                            else:
                                                print("Could not find Submit button. Manual intervention may be required.")
                                                
                                            break # Break out of panel_selectors loop
                                    except: continue
                                if found_go_to_panel: break
                            
                            if found_go_to_panel:
                                break
                            
                            time.sleep(5)
                            elapsed = int(time.time() - polling_start)
                            # Log every 20 seconds to show it's still working
                            if elapsed % 20 == 0:
                                print(f"Still waiting for extension to complete... ({elapsed}s elapsed)")
                        
                        if not found_go_to_panel:
                            print(f"Timeout: 'Go to panel' button did not appear within {timeout}s.")
                    else:
                        print("Failed to click Scan Application button after trying all icon candidates.")
                        
                        print("DEBUGGING: Listing all visible images for manual review:")
                        images = page.locator("img").all()
                        for i, img in enumerate(images):
                            try:
                                if img.is_visible(timeout=500):
                                    print(f"  Img {i}: src='{img.get_attribute('src')}', alt='{img.get_attribute('alt')}'")
                            except: pass
                    time.sleep(2)
                    
                except Exception as link_err:
                    print(f"Error processing link {url}: {link_err}")
                
                # Finalize link analytics for this attempt
                try:
                    start_dt = datetime.strptime(link_analytics["automation_start_time"], "%Y-%m-%d %H:%M:%S")
                    end_dt = datetime.now()
                    link_analytics["total_link_processing_time"] = str(end_dt - start_dt)
                    
                    # Accumulate and Save analytics incrementally
                    all_analytics.append(link_analytics)
                    with open(csv_file.parent / "analytics.json", "w") as f:
                        json.dump({
                            "overall_session_start": overall_start_time.strftime("%Y-%m-%d %H:%M:%S"),
                            "jobs": all_analytics,
                            "overall_total_time": str(datetime.now() - overall_start_time)
                        }, f, indent=4)
                    print(f"Analytics updated for {url}")
                    
                    # Send to Teams only if enabled
                    if SEND_TEAMS_NOTIFICATIONS:
                        send_teams_notification(teams_url, link_analytics)
                    else:
                        print(f"Skipping Teams notification for {url} (Notifications Disabled)")
                except Exception as ana_err:
                    print(f"Warning: Failed to save analytics: {ana_err}")

                continue

            print("\n" + "="*50)
            print("ALL JOBS PROCESSED!")
            print("="*50)
            print("Every tab has been LEFT OPEN for your review.")
            print(">>> PRESS CTRL+C IN THIS TERMINAL TO STOP THE SCRIPT BUT KEEP ALL TABS OPEN <<<")
            
            try:
                while True:
                    time.sleep(10)
            except KeyboardInterrupt:
                print("\n" + "!"*50)
                print("STOPPING SCRIPT - BROWSER STAYING OPEN")
                print("!"*50)
                print("The browser and all tabs will remain open for you.")
                os._exit(0)
            
        except Exception as e:
            print(f"Fatal Error: {e}")

if __name__ == "__main__":
    run_application_loop()
