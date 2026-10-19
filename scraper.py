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
# PHASE 2: SELENIUM WEB SCRAPING & PDF CAPTURE ENGINE
# ==============================================================================
base_url = "https://www.maersk.com/tracking/"

print("\nStarting cloud headless chrome driver session container...")
options = uc.ChromeOptions()
options.add_argument("--headless=new")           # Required for cloud server deployments
options.add_argument("--no-sandbox")              # Bypass OS security boundaries inside Docker
options.add_argument("--disable-dev-shm-usage")   # Overcome shared memory resource restrictions
options.add_argument("--window-size=1920,1080")

try:
    # Let UC handle version tracking automatically for stability
    driver = uc.Chrome(options=options)
except Exception as browser_error:
    print(f"❌ Initialization failure starting Chromium: {browser_error}")
    exit()

cookie_accepted = False

for container_number in container_numbers:
    url = f"{base_url}{container_number}"
    print(f"\nNavigating to tracking target: {url}")
    driver.get(url)
    time.sleep(2)

    # Accept cookie banner modal if present
    if not cookie_accepted:
        try:
            allow_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Allow all')]"))
            )
            allow_button.click()
            cookie_accepted = True
            print("Tracking site cookie preferences accepted.")
        except Exception:
            pass

    time.sleep(6)  # Standard loading buffer pause

    # Generate print snapshot via DevTools Protocol target
    pdf_filename = os.path.join(save_folder, f"{container_number}_{today_date}.pdf")
    try:
        pdf = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True,
            "landscape": False
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
