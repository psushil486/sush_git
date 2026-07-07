import time
import urllib.parse
import pandas as pd
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#================= CONFIG =================
FILE_PATH = r"C:\Users\pradsush\Downloads\PI.xlsx"

# ----- Input columns -----
INPUT_URL_COLUMN = "PI URL"
COMP_COLUMN= "Competitor Name"

# ----- Output columns -----
PRICE_COLUMN= "Price"
AVAIL_COLUMN   = "Availability"
REBATE_COLUMN  = "Rebate"
SHIP_COLUMN    = "Shipping"
SELLER_COLUMN  = "Seller Name"
URL_COLUMN     = "URL"

BATCH_SIZE = 20
BATCH_WAIT = 5
TAB_RENDER_WAIT = .5

DEBUG_PORT = "127.0.0.1:9222"

# Sentinel values per business rules
NOT_AVAILABLE_FILL = "0"
NOT_FOUND_FILL     = "Check manually"
DEFAULT_SELLER     = "Retail"

# ================= CHECK DEBUG CHROME =================
print("🔍 Checking Chrome debug port...")
try:
    res = requests.get(f"http://{DEBUG_PORT}/json", timeout=3)
    if res.status_code != 200:
        raise Exception("Debug port not responding")
    print("✅ Chrome debug mode detected")
except:
    print("❌ Chrome NOT running in debug mode!")
    print('👉 Start: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\\selenium\\chrome-profile"')
    exit()

# ================= LOAD EXCEL =================
print("📂 Loading Excel...")
try:
    df = pd.read_excel(FILE_PATH, dtype={
        PRICE_COLUMN: str, AVAIL_COLUMN: str,
        REBATE_COLUMN: str, SHIP_COLUMN: str,
        SELLER_COLUMN: str, URL_COLUMN: str,
    })
except Exception:
    df = pd.read_excel(FILE_PATH)

# Validate input column exists
if INPUT_URL_COLUMN not in df.columns:
    print(f"❌ Column '{INPUT_URL_COLUMN}' not found in Excel.")
    print(f"   Available columns: {list(df.columns)}")
    exit()

if COMP_COLUMN not in df.columns:
    print(f"❌ Column '{COMP_COLUMN}' not found in Excel.")
    print(f"   Available columns: {list(df.columns)}")
    exit()

# Force string dtype on all result columns
for col in [PRICE_COLUMN, AVAIL_COLUMN, REBATE_COLUMN, SHIP_COLUMN, SELLER_COLUMN, URL_COLUMN]:
    if col not in df.columns:
        df[col] = ""
    df[col] = df[col].astype(object)

valid_rows = df[df[INPUT_URL_COLUMN].notna()]
print(f"📊 Total rows: {len(df)}")
print(f"✅ Valid input URL rows: {len(valid_rows)}")

if valid_rows.empty:
    print("❌ No valid input URLs found.")
    exit()

# ================= CONNECT SELENIUM =================
print("🔗 Connecting to Chrome...")
chrome_options = Options()
chrome_options.debugger_address = DEBUG_PORT
driver = webdriver.Chrome(options=chrome_options)

try:
    driver.maximize_window()
except:
    pass
print("✅ Connected successfully")

# ================= HELPER: clean competitor URL =================
def clean_comp_url(raw_href):
    if not raw_href:
        return ""
    try:
        parsed = urllib.parse.urlparse(raw_href)
        qs = urllib.parse.parse_qs(parsed.query)
        if "URL" in qs:
            return urllib.parse.unquote(qs["URL"][0])
        return raw_href
    except Exception:
        return raw_href

# ================= HELPER: detect "Not Available" =================
def is_not_available(value):
    """Check if a value represents 'Not Available' (case-insensitive)."""
    if value is None:
        return False
    s = str(value).strip().lower()
    if s == "" or s == "nan":
        return False
    return s in ("not available", "n/a", "na", "not avilable", "notavailable")

# ================= EXTRACTION FUNCTION =================
def extract_from_page(comp_name):
    """
    Returns: (price, availability, rebate, shipping, seller, comp_url, status)
    status:
      "ok"→ data extracted normally
      "not_available" → competitor row found but availability = Not Available
      "not_found"     → competitor TD not found on page (probably on page 2)
    """
    price = ""
    availability = ""
    rebate = ""
    shipping = ""
    seller = ""        # ← NOT defaulted to "Retail" anymore
    comp_url = ""

    # ===== Step 1: Locate competitor TD =====
    comp_td = None
    comp_td_xpath = f'//td[@data-search="{comp_name}"]'
    try:
        comp_td = driver.find_element(By.XPATH, comp_td_xpath)
    except Exception:
        comp_td_xpath_ci = (
            f'//td[contains(translate(@data-search,'
            f'"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),'
            f'"{comp_name.lower()}")]'
        )
        try:
            comp_td = driver.find_element(By.XPATH, comp_td_xpath_ci)
        except Exception:
            comp_td = None

    # If competitor not found at all → likely on page 2 → "not_found"
    if comp_td is None:
        print(f"   ⚠️ '{comp_name}' not found on page 1 (likely on page 2)")
        return price, availability, rebate, shipping, seller, comp_url, "not_found"

    # ===== Step 2: Check Availability FIRST =====
    try:
        avail_td = comp_td.find_element(By.XPATH, './following-sibling::td[2]')
        avail_attr = avail_td.get_attribute("data-order")
        if avail_attr and avail_attr.strip() and "{{" not in avail_attr:
            availability = avail_attr.strip()
        else:
            avail_span = avail_td.find_element(
                By.XPATH, './/span[contains(@class,"label")]'
            )
            availability = avail_span.text.strip()
    except Exception:
        pass

    # If availability says "Not Available" → return "not_available" status
    if availability.strip().lower() == "not available":
        return price, availability, rebate, shipping, seller, comp_url, "not_available"

    # ===== Step 3: Seller Name (default "Retail"ONLY here — valid + available comp) =====
    try:
        seller_p = comp_td.find_element(By.XPATH, './/p[contains(., "Seller Name")]')
        seller_text = seller_p.text.replace("Seller Name:", "").strip()
        seller = seller_text if seller_text else DEFAULT_SELLER
    except Exception:
        seller = DEFAULT_SELLER  # No seller name found → "Retail"

    # ===== Step 4: Price + Comp URL =====
    try:
        price_td = comp_td.find_element(By.XPATH, './following-sibling::td[1]')

        price_attr = price_td.get_attribute("data-order")
        if price_attr and price_attr.strip() and "{{" not in price_attr:
            price = price_attr.strip()
        else:
            price_span = price_td.find_element(
                By.XPATH, './/span[contains(@class,"ng-binding")]'
            )
            price = price_span.text.strip()

        try:
            anchor = price_td.find_element(By.XPATH, './/a[@href]')
            raw_href = anchor.get_attribute("href") or ""
            comp_url = clean_comp_url(raw_href)
        except Exception:
            comp_url = ""
    except Exception as e:
        print(f"   ⚠️ Price/URL extraction failed: {e}")

    # ===== Step 5: Rebate =====
    try:
        rebate_td = comp_td.find_element(By.XPATH, './following-sibling::td[3]')
        rebate_span = rebate_td.find_element(
            By.XPATH, './/span[contains(@class,"ng-binding")]'
        )
        rebate = rebate_span.text.strip() or "0"
    except Exception:
        rebate = "0"

    # ===== Step 6: Shipping =====
    try:
        ship_td = comp_td.find_element(By.XPATH, './following-sibling::td[4]')
        ship_span = ship_td.find_element(By.XPATH, './/span[contains(@class,"ng-binding")]')
        shipping = ship_span.text.strip()
    except Exception:
        shipping = ""

    return (price, availability, rebate, shipping, seller, comp_url, "ok")

# ================= APPLY BUSINESS RULES =================
def apply_rules(row_idx, result):
    """Apply business rules to fill DataFrame based on extraction status."""
    price, availability, rebate, shipping, seller, comp_url, status = result

    if status == "not_available":
        # Competitor found but availability = Not Available → all "0"
        df.at[row_idx, PRICE_COLUMN]  = NOT_AVAILABLE_FILL
        df.at[row_idx, AVAIL_COLUMN]  = NOT_AVAILABLE_FILL
        df.at[row_idx, REBATE_COLUMN] = NOT_AVAILABLE_FILL
        df.at[row_idx, SHIP_COLUMN]   = NOT_AVAILABLE_FILL
        df.at[row_idx, SELLER_COLUMN] = NOT_AVAILABLE_FILL
        df.at[row_idx, URL_COLUMN]    = NOT_AVAILABLE_FILL
    elif status == "not_found":
        # Valid competitor name but not found on page 1 → "Check manually"
        df.at[row_idx, PRICE_COLUMN]  = NOT_FOUND_FILL
        df.at[row_idx, AVAIL_COLUMN]  = NOT_FOUND_FILL
        df.at[row_idx, REBATE_COLUMN] = NOT_FOUND_FILL
        df.at[row_idx, SHIP_COLUMN]   = NOT_FOUND_FILL
        df.at[row_idx, SELLER_COLUMN] = NOT_FOUND_FILL
        df.at[row_idx, URL_COLUMN]    = NOT_FOUND_FILL
    else:
        # Normal extraction (Retail default already handled in extract_from_page)
        df.at[row_idx, PRICE_COLUMN]  = str(price)
        df.at[row_idx, AVAIL_COLUMN]  = str(availability)
        df.at[row_idx, REBATE_COLUMN] = str(rebate)
        df.at[row_idx, SHIP_COLUMN]   = str(shipping)
        df.at[row_idx, SELLER_COLUMN] = str(seller)
        df.at[row_idx, URL_COLUMN]    = str(comp_url)

# ================= FILL "Not Available" ROWS (skip page entirely) =================
def fill_not_available(row_idx):
    """When Competitor Name itself is 'Not Available', fill all output columns with '0'."""
    df.at[row_idx, PRICE_COLUMN]  = NOT_AVAILABLE_FILL
    df.at[row_idx, AVAIL_COLUMN]  = NOT_AVAILABLE_FILL
    df.at[row_idx, REBATE_COLUMN] = NOT_AVAILABLE_FILL
    df.at[row_idx, SHIP_COLUMN]   = NOT_AVAILABLE_FILL
    df.at[row_idx, SELLER_COLUMN] = NOT_AVAILABLE_FILL
    df.at[row_idx, URL_COLUMN]    = NOT_AVAILABLE_FILL

# ================= BATCH PROCESS =================
rows_list = valid_rows.index.tolist()

try:
    anchor_handle = driver.current_window_handle
except Exception:
    anchor_handle = driver.window_handles[0]

for batch_start in range(0, len(rows_list), BATCH_SIZE):
    batch_indices = rows_list[batch_start: batch_start + BATCH_SIZE]
    print(f"🚀 Batch {batch_start} → {batch_start + len(batch_indices)}")

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
        except:
            pass
        break

    tab_map = []

    # ===== OPEN TABS (skip rows where Competitor Name = "Not Available") =====
    print(f"📂 Opening tabs for batch (skipping rows where Competitor Name = 'Not Available')...")
    for idx in batch_indices:
        comp_name_raw = df.at[idx, COMP_COLUMN]

        # ✅ KEY CHECK: If COMPETITOR NAME is "Not Available" → skip page entirely
        if is_not_available(comp_name_raw):
            fill_not_available(idx)
            print(f"⏭️  Row {idx} | Competitor Name = 'Not Available' → all fields set to 0 (page skipped)")
            continue

        url = df.at[idx, INPUT_URL_COLUMN]
        try:
            driver.switch_to.window(anchor_handle)
            driver.execute_script(f"window.open('{url}', '_blank');")
            new_handle = driver.window_handles[-1]
            tab_map.append((new_handle, idx))
        except Exception as e:
            print(f"⚠️ Could not open tab for row {idx}: {e}")

    # If no tabs were opened in this batch, skip the wait
    if not tab_map:
        print("ℹ️  All rows in this batch were 'Not Available' — no tabs to process.")
        try:
            df.to_excel(FILE_PATH, index=False)
            print(f"💾 Progress saved after batch {batch_start}")
        except Exception as e:
            print(f"⚠️ Could not save: {e}")
        continue

    print(f"⏳ Waiting {BATCH_WAIT} sec for parallel tab loads...")
    time.sleep(BATCH_WAIT)

    # ===== PROCESS EACH TAB =====
    for handle, row_index in tab_map:
        try:
            if handle not in driver.window_handles:
                print(f"⚠️ Tab missing for row {row_index}, skipping.")
                continue

            driver.switch_to.window(handle)

            # Wait for table data to render
            try:
                WebDriverWait(driver, TAB_RENDER_WAIT * 5).until(
                    EC.presence_of_element_located(
                        (By.XPATH, '//td[@data-search and @data-search!=""]')
                    )
                )
            except Exception:
                time.sleep(TAB_RENDER_WAIT)

            time.sleep(1)

            comp_name = str(df.at[row_index, COMP_COLUMN]).strip()
            result = extract_from_page(comp_name)
            apply_rules(row_index, result)

            status = result[-1]
            price = result[0]
            avail = result[1]
            seller = result[4]

            if status == "not_available":
                print(f"⚪ Row {row_index} | {comp_name} | NOT AVAILABLE on page → all fields set to 0")
            elif status == "not_found":
                print(f"🔴 Row {row_index} | {comp_name} | NOT FOUND on page 1 → 'Check manually'")
            else:
                print(f"✅ Row {row_index} | {comp_name} | Price: {price} | Avail: {avail} | Seller: {seller}")
        except Exception as e:
            print(f"❌ Error row {row_index}: {e}")# Close tab (never close anchor) — properly inside the for-loop
        try:
            if handle != anchor_handle and handle in driver.window_handles:
                driver.switch_to.window(handle)
                driver.close()
        except:
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

    # Save progress
    try:
        df.to_excel(FILE_PATH, index=False)
        print(f"💾 Progress saved after batch {batch_start}")
    except Exception as e:
        print(f"⚠️ Could not save: {e}")

# ================= FINAL SAVE =================
try:
    df.to_excel(FILE_PATH, index=False)
    print("🎯 Completed Successfully")
except Exception as e:
    print(f"⚠️ Final save failed: {e}")

input("Press ENTER to exit...")
