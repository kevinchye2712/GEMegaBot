import csv
import os
import time
import base64
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- GLOBAL CONFIG ---
SOURCE_FILE = 'data.csv'  # Used by Bots 1 & 2
OUTPUT_PHONE_FILE = 'phone_numbers_found.csv'
OUTPUT_POLICY_FOLDER = 'Chan_Kwang_Policies'
OUTPUT_BATCH_ROOT = os.path.join(os.getcwd(), 'Chan_Kwang_Downloads')

# ==========================================
#        SHARED HELPER FUNCTIONS
# ==========================================

def clean_text(text):
    """Standardizes text formatting."""
    if not text: return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def sanitize_filename(name):
    """Removes illegal characters for filenames."""
    return re.sub(r'[<>:"/\\|?*]', '', str(name)).strip()

def setup_driver(mode="standard", download_dir=None):
    """
    Creates a Chrome Driver with specific settings based on the bot mode.
    Modes: 'standard', 'printing', 'downloading'
    """
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)

    if mode == "downloading" and download_dir:
        prefs = {
            "plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}],
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)
    
    # Initialize Driver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    # Specific CDP command for headless downloading if needed
    if mode == "downloading" and download_dir:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": download_dir
        })
        
    return driver

# ==========================================
#        BOT 1: PHONE SCRAPER
# ==========================================

def run_phone_scraper():
    print("\n" + "="*40)
    print("📱 RUNNING: PHONE NUMBER SCRAPER")
    print("="*40)

    def fix_malaysian_number(phone_str):
        digits = re.sub(r'\D', '', str(phone_str))
        if not digits: return None
        if digits.startswith('1'): digits = '0' + digits
        return digits

    def find_search_box(driver):
        search_xpaths = ["//td[contains(text(), 'Policy No')]/following-sibling::td/input", "//input[@name='policyNo']"]
        
        # Check Main
        driver.switch_to.default_content()
        for xpath in search_xpaths:
            try:
                el = driver.find_element(By.XPATH, xpath)
                if el.is_displayed(): return el, "MAIN"
            except: pass

        # Check Frames
        frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
        for i, frame in enumerate(frames):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(frame)
                for xpath in search_xpaths:
                    try:
                        el = driver.find_element(By.XPATH, xpath)
                        if el.is_displayed(): return el, i
                    except: pass
            except: continue
        return None, None

    def check_if_details_loaded(driver):
        driver.switch_to.default_content()
        if "Mobile No" in driver.page_source: return "MAIN"
        frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
        for i, frame in enumerate(frames):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(frame)
                if "Mobile No" in driver.page_source: return i
            except: continue
        return None

    def extract_mobile(driver):
        try:
            mobile_cell = driver.find_element(By.XPATH, "//td[contains(., 'Mobile No')]/following-sibling::td")
            return fix_malaysian_number(clean_text(mobile_cell.text))
        except: return None

    # --- EXECUTION ---
    driver = setup_driver("standard")
    driver.get("https://epartner-my.greateasternlife.com/en.html")
    input("🔴 LOG IN manually, stay on the page, and press ENTER here...")

    if not os.path.exists(SOURCE_FILE):
        print(f"❌ Error: {SOURCE_FILE} not found!")
        return

    # Prepare Output
    if not os.path.exists(OUTPUT_PHONE_FILE):
        with open(OUTPUT_PHONE_FILE, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(["Name", "Policy Number", "Mobile Number"])

    with open(SOURCE_FILE, mode='r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        next(reader, None)
        records = list(reader)

    SEARCH_URL = "https://epartner-my.greateasternlife.com/ePartner3_fpms/fpms/customerService/policyEnquiry.do"

    for i, row in enumerate(records):
        if not row: continue
        name = row[0]
        policy_num = row[1].strip().zfill(10)

        print(f"[{i+1}/{len(records)}] {policy_num}...", end=" ", flush=True)

        try:
            driver.get(SEARCH_URL)
            box, frame_id = None, None
            
            # Find Box
            for _ in range(3):
                box, frame_id = find_search_box(driver)
                if box: break
                time.sleep(1)

            if box:
                try:
                    box.click()
                    box.send_keys(Keys.CONTROL + "a")
                    box.send_keys(Keys.DELETE)
                    box.send_keys(policy_num)
                    time.sleep(0.5)
                    box.send_keys(Keys.RETURN)
                    time.sleep(3)
                except:
                    print("❌ Error typing")
                    continue

                # Check Result
                details_frame = check_if_details_loaded(driver)
                
                # Logic: Direct Hit vs Link Hunt
                if details_frame is not None:
                    driver.switch_to.default_content()
                    if details_frame != "MAIN":
                        frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
                        driver.switch_to.frame(frames[details_frame])
                    
                    mobile = extract_mobile(driver)
                    print(f"📞 {mobile} (Direct Hit)")
                    with open(OUTPUT_PHONE_FILE, "a", newline="", encoding="utf-8") as f:
                        csv.writer(f).writerow([name, policy_num, mobile])
                else:
                    # Link Hunt Logic
                    link_found = False
                    driver.switch_to.default_content()
                    frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
                    search_targets = ["MAIN"] + list(range(len(frames)))
                    
                    for target in search_targets:
                        try:
                            driver.switch_to.default_content()
                            if target != "MAIN": driver.switch_to.frame(frames[target])
                            link = driver.find_element(By.PARTIAL_LINK_TEXT, policy_num)
                            link.click()
                            time.sleep(3)
                            link_found = True
                            break
                        except: pass
                    
                    if link_found:
                        details_frame = check_if_details_loaded(driver)
                        if details_frame is not None:
                            driver.switch_to.default_content()
                            if details_frame != "MAIN":
                                frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
                                driver.switch_to.frame(frames[details_frame])
                            mobile = extract_mobile(driver)
                            print(f"📞 {mobile} (Via Link)")
                            with open(OUTPUT_PHONE_FILE, "a", newline="", encoding="utf-8") as f:
                                csv.writer(f).writerow([name, policy_num, mobile])
                        else:
                            print("⚠️ Details page empty?")
                    else:
                        print("⚠️ Not found")
                        with open(OUTPUT_PHONE_FILE, "a", newline="", encoding="utf-8") as f:
                            csv.writer(f).writerow([name, policy_num, "Not Found"])
            else:
                print("❌ Search box missing")
        except Exception as e:
            print(f"❌ Error: {e}")

    print("✅ Phone Scraper Complete.")

# ==========================================
#        BOT 2: POLICY PRINTER
# ==========================================

def run_policy_printer():
    print("\n" + "="*40)
    print("🖨️  RUNNING: POLICY PRINTER (PDF)")
    print("="*40)

    def print_page_to_pdf(driver, output_path):
        try:
            print_options = {
                'landscape': False, 'displayHeaderFooter': False,
                'printBackground': True, 'preferCSSPageSize': True,
            }
            result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
            with open(output_path, 'wb') as f:
                f.write(base64.b64decode(result['data']))
            return True
        except Exception as e:
            print(f"   ❌ PDF Error: {e}")
            return False

    def hunt_for_box_globally(driver):
        print("   Scanning tabs...")
        for tab in driver.window_handles:
            try:
                driver.switch_to.window(tab)
                # Check Main
                try:
                    el = driver.find_element(By.NAME, "policyNo")
                    if el.is_displayed(): return el, tab, None
                except: pass
                
                # Check Frames
                frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
                for i, frame in enumerate(frames):
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame(frame)
                        el = driver.find_element(By.NAME, "policyNo")
                        if el.is_displayed(): return el, tab, i
                    except: pass
            except: continue
        return None, None, None

    # --- EXECUTION ---
    driver = setup_driver("standard")
    driver.get("https://epartner-my.greateasternlife.com/en.html")
    
    print("🔴 1. Log in.")
    print("🔴 2. Open 'Customer Service' page.")
    input("Press ENTER to Calibrate...")

    box, correct_tab, correct_frame = hunt_for_box_globally(driver)
    if not box:
        print("❌ Could not find search box.")
        return

    # Visual Confirmation
    driver.switch_to.window(correct_tab)
    if correct_frame is not None:
        frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
        driver.switch_to.frame(frames[correct_frame])
    
    driver.execute_script("arguments[0].style.border='5px solid red';", box)
    if input("👉 Box found (Red Border)? (y/n): ").lower() != 'y': return

    if not os.path.exists(OUTPUT_POLICY_FOLDER):
        os.makedirs(OUTPUT_POLICY_FOLDER)

    with open(SOURCE_FILE, mode='r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        next(reader, None)
        records = list(reader)

    for i, row in enumerate(records):
        if not row: continue
        name = row[0]
        policy_num = row[1].strip().zfill(10)
        
        print(f"[{i+1}/{len(records)}] {policy_num}...", end=" ", flush=True)

        try:
            driver.switch_to.window(correct_tab)
            driver.switch_to.default_content()
            if correct_frame is not None:
                frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
                driver.switch_to.frame(frames[correct_frame])
            
            target_box = driver.find_element(By.NAME, "policyNo")
            target_box.click()
            target_box.send_keys(Keys.CONTROL + "a")
            target_box.send_keys(Keys.DELETE)
            target_box.send_keys(policy_num)
            target_box.send_keys(Keys.RETURN)
            
            time.sleep(4)

            # Click Link if exists
            try:
                driver.find_element(By.PARTIAL_LINK_TEXT, policy_num).click()
                time.sleep(3)
            except: pass

            filename = f"{sanitize_filename(name)} - {sanitize_filename(policy_num)}.pdf"
            save_path = os.path.join(OUTPUT_POLICY_FOLDER, filename)
            
            if print_page_to_pdf(driver, save_path):
                print("✅ Saved")
            else:
                print("⚠️ Print Failed")

            driver.back()
            time.sleep(2)
        except Exception as e:
            print(f"❌ Error: {e}")
            try: driver.back()
            except: pass

    print("✅ Policy Printer Complete.")

# ==========================================
#        BOT 3: BATCH DOWNLOADER
# ==========================================

def run_batch_downloader():
    print("\n" + "="*40)
    print("⬇️  RUNNING: BATCH DOWNLOADER")
    print("="*40)
    
    # Helper to clean .tmp files
    def clean_temp_files(download_dir):
        if not os.path.exists(download_dir): return
        for f in os.listdir(download_dir):
            if f.endswith(".crdownload") or f.endswith(".tmp"):
                try: os.remove(os.path.join(download_dir, f))
                except: pass

    def wait_and_rename(download_dir, new_filename, files_before, timeout=60):
        end_time = time.time() + timeout
        while time.time() < end_time:
            files_now = set(os.listdir(download_dir))
            new_files = files_now - files_before
            valid_new = [f for f in new_files if not f.endswith('.crdownload') and not f.endswith('.tmp')]
            
            if not valid_new:
                time.sleep(0.5)
                continue
            
            target = valid_new[0]
            try:
                time.sleep(0.5)
                os.rename(os.path.join(download_dir, target), os.path.join(download_dir, new_filename))
                print(f"      ↳ Renamed: {new_filename}")
                return True
            except:
                time.sleep(1)
                continue
        return False

    def process_table(driver, wait, save_folder):
        clean_temp_files(save_folder)
        # Update download path dynamically for this batch
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": save_folder})
        
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "tr")))
            rows = driver.find_elements(By.XPATH, "//tr[.//a[text()='Download']]")
            print(f"   ⬇️ Found {len(rows)} files.")
            
            for index in range(len(rows)):
                try:
                    current_row = driver.find_elements(By.XPATH, "//tr[.//a[text()='Download']]")[index]
                    holder_name = current_row.find_element(By.XPATH, "./td[2]").text.strip()
                    policy_no = current_row.find_element(By.XPATH, "./td[4]").text.strip()
                    fname = f"{sanitize_filename(holder_name)} - {sanitize_filename(policy_no)}_Repricing.pdf"
                    
                    files_before = set(os.listdir(save_folder))
                    btn = current_row.find_element(By.XPATH, ".//a[text()='Download']")
                    driver.execute_script("arguments[0].scrollIntoView();", btn)
                    btn.click()
                    
                    wait_and_rename(save_folder, fname, files_before)
                except Exception as e:
                    print(f"      ⚠️ Row Error: {e}")
        except Exception as e:
            print(f"   ❌ Table Error: {e}")

    # --- EXECUTION ---
    # Setup driver with initial default path
    driver = setup_driver("downloading", OUTPUT_BATCH_ROOT)
    wait = WebDriverWait(driver, 15)
    
    driver.get("https://epartner-my.greateasternlife.com/en.html")
    print("Please Log In.")

    while True:
        print("\nOPTIONS:")
        print(" A) Auto-process Batch List")
        print(" B) Process Single Folder (Manual)")
        print(" Q) Return to Main Menu")
        choice = input("👉 Command: ").lower()
        
        if choice == 'q': 
            driver.quit()
            break

        # Hunt Logic
        batch_tab, batch_frame = None, None
        for tab in driver.window_handles:
            driver.switch_to.window(tab)
            if driver.find_elements(By.PARTIAL_LINK_TEXT, "Policy Anniversary"): 
                batch_tab = tab; break
            # Check frames... (simplified for brevity, assumes standard frame hunt)
            
        if choice == 'a':
            # Simplified Batch Logic
            cat_name = input("Category Name: ") or "Uncategorized"
            links = driver.find_elements(By.PARTIAL_LINK_TEXT, "Policy Anniversary")
            batches = [(el.text, el.get_attribute('href')) for el in links if el.get_attribute('href')]
            
            for b_name, b_url in batches:
                folder = os.path.join(OUTPUT_BATCH_ROOT, sanitize_filename(cat_name), sanitize_filename(b_name))
                if not os.path.exists(folder): os.makedirs(folder)
                driver.get(b_url)
                process_table(driver, wait, folder)
        
        elif choice == 'b':
            # Single Folder Logic
            f_name = input("Folder Name: ") or "Manual_DL"
            folder = os.path.join(OUTPUT_BATCH_ROOT, sanitize_filename(f_name))
            if not os.path.exists(folder): os.makedirs(folder)
            process_table(driver, wait, folder)

# ==========================================
#        MAIN MENU
# ==========================================

def main():
    while True:
        print("\n" + "█"*40)
        print("      🤖 CHAN KWANG MEGA BOT 🤖")
        print("█"*40)
        print("1. Phone Number Scraper")
        print("2. Policy Printer (PDF)")
        print("3. Batch Downloader")
        print("Q. Quit")
        
        choice = input("\n👉 Select a bot (1-3): ").lower()
        
        if choice == '1':
            run_phone_scraper()
        elif choice == '2':
            run_policy_printer()
        elif choice == '3':
            run_batch_downloader()
        elif choice == 'q':
            print("Goodbye!")
            break
        else:
            print("❌ Invalid selection.")

if __name__ == "__main__":
    main()