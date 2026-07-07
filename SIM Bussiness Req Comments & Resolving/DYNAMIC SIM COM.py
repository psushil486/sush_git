import time
import re
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ================= CONFIG =================

CSV_PATH = r"C:\Users\pradsush\PycharmProjects\PythonProject\learn\SIM COMMENTING TEST.csv"

COMMENT_TEXT = """Hi Team, Can we get an update on this ?

Regards,
Sushil"""

ACTIVITY_COMMENT_XPATH = "//section[@class='document-stream']//div[@class='activity-body rich-text']"
TIME_XPATH = "//span[@class='activity-actor-action']//time[@data-format='preferred']"

COMMENT_BOX_XPATH = "//textarea[@placeholder='Compose a new comment...']"
POST_BUTTON_XPATH = "//form[@id='issue-stream-form']//button[@data-csm-counter='issueStreamFormComment']"

RESOLVE_BUTTON_XPATH = "//div[@class='flex-item']//button[contains(@class,'resolve-issue')]"
FINAL_DEFECT_DROPDOWN_XPATH = "//div[@class='custom-field input-wrapper']//select"
FINAL_RESOLVE_XPATH = "//input[contains(@class,'resolve')]"

WAIT = 20
BATCH_SIZE = 20
BATCH_LOAD_WAIT = 50   # 🔹 NEW: wait after opening 20 SIMs

# ================= PAGE READY =================

def wait_for_jira_ready(driver):
    WebDriverWait(driver, WAIT).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    WebDriverWait(driver, WAIT).until(
        EC.presence_of_element_located((By.XPATH, COMMENT_BOX_XPATH))
    )

# ================= BUSINESS DAYS =================

def business_days_ago(days):
    count = 0
    cur = datetime.today()
    while count < days:
        cur -= timedelta(days=1)
        if cur.weekday() < 5:
            count += 1
    return count

def get_business_days(driver):
    try:
        text = driver.find_element(By.XPATH, TIME_XPATH).text.lower()
        m = re.search(r"(\d+)(h|d|mo|y)", text)
        if not m:
            return 0

        val, unit = int(m.group(1)), m.group(2)

        if unit == "d":
            return business_days_ago(val)
        if unit in ("mo", "y"):
            return 999
        return 0
    except Exception:
        return 0

# ================= RESOLVE GATE =================

def get_clickable_resolve_button(driver):
    try:
        return WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, RESOLVE_BUTTON_XPATH))
        )
    except Exception:
        return None

# ================= DRIVER =================

driver = webdriver.Chrome()
wait = WebDriverWait(driver, WAIT)

df = pd.read_csv(CSV_PATH)
urls = df["IssueUrl"].dropna().tolist()

# ================= LOGIN =================

driver.get(urls[0])
input("Login manually, then press ENTER to continue...")
main_tab = driver.current_window_handle

# ================= BATCH PROCESS =================

work_urls = urls[1:]
batches = [work_urls[i:i + BATCH_SIZE] for i in range(0, len(work_urls), BATCH_SIZE)]

for batch_no, batch_urls in enumerate(batches, start=1):
    print(f"\n🚀 Opening Batch {batch_no} ({len(batch_urls)} SIMs)")

    opened_tabs = []

    # -------------------------------------------------
    # STEP 1: OPEN 20 SIMs SIMULTANEOUSLY
    # -------------------------------------------------
    for url in batch_urls:
        driver.execute_script("window.open(arguments[0])", url)
        opened_tabs.append(driver.window_handles[-1])

    # 🔹 NEW: WAIT AFTER BATCH LOAD
    print(f"⏳ Waiting {BATCH_LOAD_WAIT} seconds for batch to fully load...")
    time.sleep(BATCH_LOAD_WAIT)

    # -------------------------------------------------
    # STEP 2: PROCESS EACH TAB
    # -------------------------------------------------
    for tab in opened_tabs:
        driver.switch_to.window(tab)
        print(f"\nProcessing: {driver.current_url}")

        try:
            wait_for_jira_ready(driver)
        except TimeoutException:
            print("❌ Page not ready → closing tab")
            driver.close()
            driver.switch_to.window(main_tab)
            continue

        comment_blocks = driver.find_elements(By.XPATH, ACTIVITY_COMMENT_XPATH)
        comment_text = " ".join(c.text.lower() for c in comment_blocks)

        # ================= CASE 1 =================
        if "accepting the error" in comment_text:
            print("✅ Case 1: Accepting the error")

            resolve_btn = get_clickable_resolve_button(driver)
            if not resolve_btn:
                driver.close()
                driver.switch_to.window(main_tab)
                continue

            resolve_btn.click()

            final_defect = wait.until(
                EC.element_to_be_clickable((By.XPATH, FINAL_DEFECT_DROPDOWN_XPATH))
            )
            Select(final_defect).select_by_visible_text("Mapper Oversight")

            final_resolve = wait.until(
                EC.element_to_be_clickable((By.XPATH, FINAL_RESOLVE_XPATH))
            )
            final_resolve.click()

            print("🟢 Resolved → tab kept open")
            driver.switch_to.window(main_tab)
            continue

        # ================= CASE 2 =================
        if comment_text.strip():
            print("ℹ️ Existing comment → closing tab")
            driver.close()
            driver.switch_to.window(main_tab)
            continue

        business_days = get_business_days(driver)
        print(f"No comments | Business days = {business_days}")

        if 3 <= business_days <= 5:
            comment_box = wait.until(
                EC.element_to_be_clickable((By.XPATH, COMMENT_BOX_XPATH))
            )
            comment_box.send_keys(COMMENT_TEXT)

            time.sleep(1)

            post_btn = wait.until(
                EC.element_to_be_clickable((By.XPATH, POST_BUTTON_XPATH))
            )
            post_btn.click()

            print("🟢 Follow-up posted → tab kept open")
            driver.switch_to.window(main_tab)
            continue

        # ================= NO MATCH =================
        print("❌ No case matched → closing tab")
        driver.close()
        driver.switch_to.window(main_tab)

    print(f"\n✅ Batch {batch_no} completed.")

# ================= MANUAL EXIT =================

cmd = input(
    "\nAll batches completed.\n"
    "Press ENTER to keep browser open,\n"
    "or type 'close' and press ENTER: "
)

if cmd.strip().lower() == "close":
    driver.quit()
    print("🛑 Browser closed.")
else:
    print("🟢 Browser left open.")
