import time
import re
import pandas as pd
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ================= CONFIG =================
FILE_PATH = r"C:\Users\pradsush\Downloads\CPT.xlsx"

URL_COLUMN = "URL"
CPT_COLUMN = "CPT"
COMP_COLUMN = "Competitor Name"

BATCH_SIZE = 20
TAB_LOAD_DELAY = 1   # Small delay between opening tabs
BATCH_WAIT = 10          # Buffer for tabs to start loading (reduced from 10)
ELEMENT_WAIT = 5        # Max seconds to wait for XPath to appear per tab

DEBUG_PORT = "127.0.0.1:9222"

# ================= CHECK DEBUG CHROME =================
print("🔍 Checking Chrome debug port...")
try:
    res = requests.get(f"http://{DEBUG_PORT}/json", timeout=3)
    if res.status_code != 200:
        raise Exception("Debug port not responding")
    print("✅ Chrome debug mode detected")
except Exception:
    print("❌ Chrome NOT running in debug mode!")
    print('👉 Start: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\\selenium\\chrome-profile"')
    exit()

# ================= LOAD EXCEL =================
print("📂 Loading Excel...")
df = pd.read_excel(FILE_PATH)

# Validate URL column
if URL_COLUMN not in df.columns:
    print(f"❌ Column '{URL_COLUMN}' not found in Excel.")
    print(f"   Available columns: {list(df.columns)}")
    exit()

# Ensure CPT and Competitor columns exist AND accept string values
for col in [CPT_COLUMN, COMP_COLUMN]:
    if col not in df.columns:
        df[col] = ""
    df[col] = df[col].astype(object)

valid_rows = df[df[URL_COLUMN].notna()]
print(f"📊 Total rows: {len(df)}")
print(f"✅ Valid URL rows: {len(valid_rows)}")

if valid_rows.empty:
    print("❌ No valid URLs found.")
    exit()

# ================= CONNECT SELENIUM TO EXISTING CHROME =================
print("🔗 Connecting to existing Chrome...")
chrome_options = Options()
chrome_options.debugger_address = DEBUG_PORT
driver = webdriver.Chrome(options=chrome_options)

try:
    driver.maximize_window()
except Exception:
    pass
print("✅ Connected successfully (using existing logged-in session)")

# Anchor handle = the tab that was already open in your existing Chrome
try:
    anchor_handle = driver.current_window_handle
except Exception:
    anchor_handle = driver.window_handles[0]

# ================= EXTRACTION FUNCTION =================
def extract_data():
    """Extract CPT and Competitor Name from the current tab using smart waits."""
    cpt_xpath = '(//tr[@style="grid-template-columns: repeat(2, 1fr);"])[16]'
    comp_xpath = '(//tr[@style="grid-template-columns: repeat(2, 1fr);"])[18]'

    # ===== CPT =====
    try:
        cpt_element = WebDriverWait(driver, ELEMENT_WAIT).until(
            EC.presence_of_element_located((By.XPATH, cpt_xpath))
        )
        cpt_text = cpt_element.text.strip()
        cpt_value = re.sub(r'^CPT\s*', '', cpt_text).strip()
        if not cpt_value:
            cpt_value = "Not Found"
    except Exception:
        cpt_value = "Not Found"

    # ===== Competitor =====
    try:
        comp_element = WebDriverWait(driver, ELEMENT_WAIT).until(
            EC.presence_of_element_located((By.XPATH, comp_xpath))
        )
        comp_text = comp_element.text.strip()

        match = re.search(
            r'Competitor Name(.*?)Legal Constraints',
            comp_text,
            re.DOTALL
        )

        if match:
            competitor_name = match.group(1).strip()
        else:
            competitor_name = comp_text.replace("Competitor Name", "").split("Legal Constraints")[0].strip()

        if not competitor_name:
            competitor_name = "Not Found"
    except Exception:
        competitor_name = "Not Found"

    return str(cpt_value), str(competitor_name)

# ================= BATCH PROCESSALL ROWS =================
rows_list = valid_rows.index.tolist()

for batch_start in range(0, len(rows_list), BATCH_SIZE):
    batch_indices = rows_list[batch_start: batch_start + BATCH_SIZE]
    print(f"🚀 Processing Batch {batch_start} to {batch_start + len(batch_indices)}")

    # Verify session
    try:
        handles = driver.window_handles
        if anchor_handle not in handles:
            anchor_handle = handles[0]
        driver.switch_to.window(anchor_handle)
    except Exception as e:
        print(f"❌ Browser session lost: {e}")
        try:
            df.to_excel(FILE_PATH, index=False)
            print("💾 Partial progress saved.")
        except Exception:
            pass
        break

    tab_map = []  # (handle, row_index)

    # ===== OPEN TABS WITH MAPPING =====
    print(f"📂 Opening {len(batch_indices)} tabs...")
    for idx in batch_indices:
        url = df.at[idx, URL_COLUMN]
        try:
            driver.switch_to.window(anchor_handle)
            driver.execute_script(f"window.open('{url}', '_blank');")
            new_handle = driver.window_handles[-1]
            tab_map.append((new_handle, idx))
            time.sleep(TAB_LOAD_DELAY)
        except Exception as e:
            print(f"⚠️ Could not open tab for row {idx}: {e}")

    if not tab_map:
        print("ℹ️  No tabs opened in this batch — skipping.")
        continue

    print(f"⏳ Buffer wait {BATCH_WAIT} sec before extracting...")
    time.sleep(BATCH_WAIT)

    # ===== PROCESS EACH TAB =====
    for handle, row_index in tab_map:
        try:
            if handle not in driver.window_handles:
                print(f"⚠️ Tab missing for row {row_index}, skipping.")
                continue

            driver.switch_to.window(handle)

            cpt_value, competitor_name = extract_data()

            df.at[row_index, CPT_COLUMN] = cpt_value
            df.at[row_index, COMP_COLUMN] = competitor_name

            print(f"✅ Row {row_index} | CPT: {cpt_value} | Comp: {competitor_name}")

        except Exception as e:
            print(f"❌ Error on row {row_index}: {e}")
            df.at[row_index, CPT_COLUMN] = "Error"
            df.at[row_index, COMP_COLUMN] = "Error"

        # Close tab (never close anchor)
        try:
            if handle != anchor_handle and handle in driver.window_handles:
                driver.switch_to.window(handle)
                driver.close()
        except Exception:
            pass

    # Switch back to anchor
    try:
        if anchor_handle in driver.window_handles:
            driver.switch_to.window(anchor_handle)
        else:
            anchor_handle = driver.window_handles[0]
            driver.switch_to.window(anchor_handle)
    except Exception as e:
        print(f"⚠️ Could not switch to anchor: {e}")

    # Save progress after each batch
    try:
        df.to_excel(FILE_PATH, index=False)
        print(f"💾 Progress saved after batch {batch_start}")
    except Exception as e:
        print(f"⚠️ Could not save: {e}")

# ================= FINAL SAVE =================
try:
    df.to_excel(FILE_PATH, index=False)
    print("🎯 Completed and saved Excel")
except Exception as e:
    print(f"⚠️ Final save failed: {e}")

input("Press ENTER to exit...")
