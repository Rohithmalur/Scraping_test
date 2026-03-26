import os
import shutil
import time
import base64
import pandas as pd
from datetime import datetime
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium import webdriver

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# ================== Clear Chrome Cache ==================
path = os.path.expanduser("~/.undetected_chromedriver")
if os.path.exists(path):
   shutil.rmtree(path)
print("Chrome cache cleared")
# ================== Setup Paths ==================
base_dir = os.path.dirname(os.path.abspath(__file__))
data_folder = os.path.join(base_dir, "..", "Data_Folder")
input_file = os.path.join(data_folder, "maersk_cntr_input_file.xlsx")
today_date = datetime.today().strftime("%Y%m%d")
pdf_folder = os.path.join(data_folder, "Maersk_pdf_reports")
os.makedirs(pdf_folder, exist_ok=True)
output_file = os.path.join(
   data_folder,
   f"Maersk_container_tracking_{today_date}.xlsx"
)
# ================== Load Containers ==================
df_input = pd.read_excel(input_file)
container_numbers = df_input["Container number"].dropna().tolist()
print(f"Total Containers: {len(container_numbers)}")

# ================== Start Browser ==================
options = webdriver.ChromeOptions()
#options.add_argument("--headless")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
driver = webdriver.Chrome(options=options)

"""options = uc.ChromeOptions()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--start-maximized")
options.add_argument(
"user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
)

driver = uc.Chrome(version_main=145, options=options)"""
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

"""options = uc.ChromeOptions()
#options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")


driver = uc.Chrome(version_main=145, options=options)"""

wait = WebDriverWait(driver, 30)


base_url = "https://www.maersk.com/tracking/"
time.sleep(5)
all_results = []
# ================== Loop Containers ==================
for container_number in container_numbers:
   print(f"Processing: {container_number}")
   try:
       url = f"{base_url}{container_number}"
       driver.get(url)
       # Accept cookies
       try:
           allow_button = WebDriverWait(driver, 5).until(
               EC.element_to_be_clickable(
                   (By.XPATH, "//button[contains(., 'Allow all')]")
               )
           )
           allow_button.click()
       except:
           pass
       # Wait for page
       wait.until(
           EC.presence_of_element_located(
               (By.ID, "transport-plan__container__0")
           )
       )
       
       # ================== Save PDF ==================
       pdf_filename = os.path.join(
           pdf_folder,
           f"{container_number}_{today_date}.pdf"
       )
       pdf = driver.execute_cdp_cmd(
           "Page.printToPDF",
           {"printBackground": True}
       )
       with open(pdf_filename, "wb") as f:
           f.write(base64.b64decode(pdf["data"]))
       print(f"PDF saved: {pdf_filename}")
       # ================== Extract Data ==================
       container_text = driver.find_element(
           By.ID,
           "transport-plan__container__0"
       ).text
       Shipment_From = driver.find_element(
           By.CLASS_NAME,
           "search-summary-ocean__locations"
       ).text
       Container_header = driver.find_element(
           By.CLASS_NAME,
           "container__header"
       ).text
   except Exception as e:
       print(f"Error fetching {container_number}: {e}")
       continue
   # ================== Ship From / To ==================
   extracted_ship_from = "Not Available"
   extracted_ship_to = "Not Available"
   ship_from = Shipment_From.split("From", 1)
   if len(ship_from) > 1:
       extracted_ship_from = ship_from[1].split("To", 1)[0].strip()
   ship_to = Shipment_From.split("To", 1)
   if len(ship_to) > 1:
       extracted_ship_to = ship_to[1].strip()
   # ================== Equipment ==================
   equi_type = "Not Available"
   equipment_type = Container_header.split("|")
   if len(equipment_type) > 1:
       equi_type = equipment_type[1].split("Last", 1)[0].strip()
   # ================== Last Updated ==================
   last_update = "Not Available"
   last_updated = Container_header.split("Last updated:", 1)
   if len(last_updated) > 1:
       last_update = last_updated[1].strip()
   # ================== Timeline ==================
   lines = [line.strip() for line in container_text.splitlines() if line.strip()]
   location = ""
   terminal = ""
   i = 0
   while i < len(lines):
       line = lines[i]
       if line.isupper() or "terminal" in line.lower():
           if "terminal" in line.lower():
               terminal = line
           else:
               location = line
           i += 1
           continue
       if (
           i + 1 < len(lines)
           and any(month in lines[i+1] for month in
           ["Jan","Feb","Mar","Apr","May","Jun",
            "Jul","Aug","Sep","Oct","Nov","Dec"])
       ):
           event = line
           datetime_text = lines[i+1]
           parts = datetime_text.rsplit(" ", 1)
           date = parts[0] if len(parts) > 0 else ""
           time_val = parts[1] if len(parts) > 1 else ""
           row = {
               "Extract Date": datetime.today(),
               "Container No": container_number,
               "Ship From": extracted_ship_from,
               "Ship To": extracted_ship_to,
               "Equipment Type": equi_type,
               "Last Updated": last_update,
               "Event": event,
               "Date": date,
               "Time": time_val,
               "Location": location,
               "Terminal": terminal
           }
           all_results.append(row)
           i += 2
       else:
           i += 1
   print(f"Completed: {container_number}")
# ================== Save Excel ==================
try:
    driver.quit()
except Exception:
    pass
#driver.quit()
df_out = pd.DataFrame(all_results)
df_out.to_excel(output_file, index=False)
print(f"Excel saved: {output_file}")
