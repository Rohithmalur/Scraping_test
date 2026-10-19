import random # Add this at the top of your file!
import fitz  # PyMuPDF
import pandas as pd
import re
import os
import glob
import shutil
import time
import base64
from datetime import datetime
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==============================================================================
# PHASE 1: CACHE & DISK WORKSPACE SETUP (CLOUD ADAPTED)
# ==============================================================================
print("Initializing cloud script workspace...")
path = os.path.expanduser("~/.undetected_chromedriver")
if os.path.exists(path):
    try:
        shutil.rmtree(path)
        print("Cache cleared successfully.")
    except Exception as e:
        print(f"Notice: Could not clear cache directory: {e}")

today_date = datetime.today().strftime("%Y%m%d")

# Cloud File Path Adaptations (Replaces Tkinter filedialog)
base_folder = os.getcwd()
input_folder = os.path.join(base_folder, "input_data")
save_folder = os.path.join(base_folder, "output_reports", f"Maersk_pdf_reports_{today_date}")

# Ensure necessary workspaces exist
os.makedirs(input_folder, exist_ok=True)
os.makedirs(save_folder, exist_ok=True)

# Look for an excel file inside the 'input_data' folder
excel_files = glob.glob(os.path.join(input_folder, "*.xlsx")) + glob.glob(os.path.join(input_folder, "*.xls"))

if not excel_files:
    print(f"\n❌ Execution halted: No input Excel file found in '{input_folder}'.")
    print("Please place your target tracking Excel file named 'input.xlsx' there.")
    exit()

input_file = excel_files[0]
print(f"Using input tracking data file: {os.path.basename(input_file)}")

# Parse container numbers out of Excel input data source
df_input = pd.read_excel(input_file)
if "Container number" not in df_input.columns:
    print("❌ Error: The selected file does not contain a 'Container number' column.")
    exit()

container_numbers = df_input["Container number"].dropna().tolist()
print(f"Loaded {len(container_numbers)} container numbers successfully.")
print(f"Target workspace report folder created at: {save_folder}")

# ==============================================================================
# PHASE 2: SELENIUM WEB SCRAPING & PDF CAPTURE ENGINE (FORTIFIED)
# ==============================================================================
#import random # Add this at the top of your file!
base_url = "https://www.maersk.com/tracking/"

print("\nStarting cloud headless chrome driver session container...")
options = uc.ChromeOptions()
options.add_argument("--headless=new")           
options.add_argument("--no-sandbox")              
options.add_argument("--disable-dev-shm-usage")   

# 💥 ANTI-DETECTION TRICK 1: Spoof a real Windows Desktop identity
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# 💥 ANTI-DETECTION TRICK 2: Pretend you have a real physical monitor display attached
options.add_argument("--window-size=1920,1080")
options.add_argument("--start-maximized")

# Hide automation variables completely from structural canvas checks
options.add_argument("--disable-blink-features=AutomationControlled")

# Read the Chrome binary path provided by GitHub Actions
chrome_browser_path = os.environ.get("CHROME_BIN_PATH")

try:
    if chrome_browser_path:
        print(f"Directing driver to Chrome binary at: {chrome_browser_path}")
        driver = uc.Chrome(options=options, browser_executable_path=chrome_browser_path)
    else:
        driver = uc.Chrome(options=options)
except Exception as browser_error:
    print(f"❌ Initialization failure starting Chromium: {browser_error}")
    exit(1)

# 💥 ANTI-DETECTION TRICK 3: Modify execution property fields via Chrome DevTools Protocol
# This destroys remaining hidden automation parameters inside the execution memory layer
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": "Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});"
})

cookie_accepted = False

for container_number in container_numbers:
    url = f"{base_url}{container_number}"
    print(f"\nNavigating to tracking target: {url}")
    driver.get(url)
    
    # 💥 ANTI-DETECTION TRICK 4: Mimic human loading delays
    time.sleep(random.uniform(3.5, 6.2)) 

    # Accept cookie banner modal if present
    if not cookie_accepted:
        try:
            allow_button = WebDriverWait(driver, 7).until(
                EC.element_to_be_clickable((By.橫向XPATH || By.XPATH, "//button[contains(., 'Allow all')]"))
            )
            time.sleep(random.uniform(1.2, 2.5)) # Wait a brief moment before clicking like a human
            allow_button.click()
            cookie_accepted = True
            print("Tracking site cookie preferences accepted.")
        except Exception:
            pass

    # Extended buffer window to give the dashboard widgets ample time to securely pull API payloads
    print("Waiting for tracking content data to settle...")
    time.sleep(random.uniform(8.0, 12.0))  

    # Generate print snapshot via DevTools Protocol target
    pdf_filename = os.path.join(save_folder, f"{container_number}_{today_date}.pdf")
    try:
        pdf = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True,
            "landscape": False,
            "preferCSSPageSize": True
        })

        with open(pdf_filename, "wb") as f:
            f.write(base64.b64decode(pdf['data']))
        print(f"✅ PDF saved: {os.path.basename(pdf_filename)}")
    except Exception as e:
        print(f"❌ Failed to extract PDF print frame for {container_number}: {e}")

print("\nClosing scraping driver browser instance...")
driver.quit()

# ==============================================================================
# PHASE 3: DATA EXTRACTION ENGINE & PARSING FUNCTIONS
# ==============================================================================
def clean_extracted_text(text):
    """Fix common PDF layout text fragmentation issues."""
    text = re.sub(r'departu\s+re', 'departure', text, flags=re.IGNORECASE)
    text = re.sub(r'arriv\s+al', 'arrival', text, flags=re.IGNORECASE)
    text = re.sub(r'arriv\s+es', 'arrives', text, flags=re.IGNORECASE)
    text = re.sub(r'GATEW\s+AY', 'GATEWAY', text, flags=re.IGNORECASE)
    text = re.sub(r'upda\s+ted', 'updated', text, flags=re.IGNORECASE)
    return text

def extract_data_from_single_pdf(pdf_path):
    """Extract and isolate metrics, dates, and port strings from the PDF target."""
    try:
        current_extraction_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        doc = fitz.open(pdf_path)
        raw_text = ""
        for page in doc:
            raw_text += page.get_text() + "\n"
        
        clean_text = clean_extracted_text(raw_text)
        lines = [line.strip() for line in clean_text.split("\n") if line.strip()]
        
        # Search layout for either "BL number" or "Container number" variants
        identifier_index = None
        for index, line in enumerate(lines):
            if "bl number" in line.lower() or "container number" in line.lower():
                identifier_index = index
                break
                
        if identifier_index is None:
            print(f"Skipping {os.path.basename(pdf_path)}: Reference metadata header missing.")
            return None
            
        identifier_value = lines[identifier_index + 1]
        from_port = lines[identifier_index + 2]
        to_port = lines[identifier_index + 3]
        no_of_cnt = lines[identifier_index + 4]
        
        last_updated_value = "Not Found"
        for line in lines:
            if "last updated:" in line.lower():
                parts = line.split(":", 1)
                if len(parts) > 1:
                    last_updated_value = parts[1].strip()
                break
        
        date_pattern = r'(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})(?:\s+\d{2}:\d{2})?'
        timeline_events = []
        current_port = "Unknown Port"
        
        for i, line in enumerate(lines):
            if line.isupper() and len(line) > 3 and "TERMINAL" not in line and "GATEWAY" not in line:
                current_port = line
                
            match = re.search(date_pattern, line)
            if match:
                date_str = match.group(1)
                date_str_clean = re.sub(r'\s+', ' ', date_str)
                date_obj = datetime.strptime(date_str_clean, "%d %b %Y")
                
                desc = lines[i+1] if i+1 < len(lines) else ""
                if any(k in desc.lower() for k in ["departure", "arrival", "gate"]):
                    event_desc = f"{current_port} ({desc})"
                else:
                    event_desc = current_port
                    
                timeline_events.append({
                    "date": date_obj, 
                    "date_str": date_str_clean, 
                    "desc": event_desc
                })
        
        if not timeline_events:
            return None
            
        timeline_events.sort(key=lambda x: x["date"])
        
        departure_date = timeline_events[0]["date_str"]
        arrival_date = timeline_events[-1]["date_str"]
        
        transshipment_segments = []
        for ev in timeline_events[1:-1]:
            transshipment_segments.append(f"{ev['date_str']}: {ev['desc']}")
        transshipment_details = " | ".join(transshipment_segments) if transshipment_segments else "Direct Route"
        
        # Return structured mapping data
        return {
            "Container/BL ID": identifier_value,
            "Origin Port": from_port,
            "Destination Port": to_port,
            "Count Detail": no_of_cnt,
            "Departure Date": departure_date,
            "Arrival Date": arrival_date,
            "Transshipment Log": transshipment_details,
            "Last Updated": last_updated_value,
            "Processed At": current_extraction_timestamp
        }
    except Exception as parse_error:
        print(f"Error parsing {os.path.basename(pdf_path)}: {parse_error}")
        return None

# ==============================================================================
# PHASE 4: EXECUTE PARSING ENGINE
# ==============================================================================
generated_pdfs = glob.glob(os.path.join(save_folder, "*.pdf"))
if generated_pdfs:
    print(f"\nParsing data from {len(generated_pdfs)} generated PDFs...")
    extracted_records = []
    
    for pdf_path in generated_pdfs:
        record = extract_data_from_single_pdf(pdf_path)
        if record:
            extracted_records.append(record)
            
    if extracted_records:
        df_output = pd.DataFrame(extracted_records)
        output_csv_path = os.path.join(os.path.dirname(save_folder), f"Scraped_Tracking_Data_{today_date}.csv")
        df_output.to_csv(output_csv_path, index=False)
        print(f"📊 Tracking data summary successfully exported to:\n{output_csv_path}")
else:
    print("\nNo newly generated PDFs found to extract data from.")
