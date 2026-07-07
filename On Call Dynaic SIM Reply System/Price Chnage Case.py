import os
import glob
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# =====================================================
# CONFIG
# =====================================================

EXCEL_FOLDER = r"C:\Users\pradsush\Documents\Custom Office Templates\IBC"
LOGIN_URL = "https://issues.amazon.com/"
BASE_URL = "https://issues.amazon.com/issues/"

BATCH_SIZE = 30
WAIT = 20
SIM_PAUSE = 0.50
BATCH_LOAD_WAIT = 50   # 🔴 NEW

COMMENT_BOX_XPATH = "//textarea[@placeholder='Compose a new comment...']"
POST_BUTTON_XPATH = "//form[@id='issue-stream-form']//button[@data-csm-counter='issueStreamFormComment']"

RESOLVE_BUTTON_XPATH = "//div[@class='flex-item']//button[contains(@class,'resolve-issue')]"
ROOT_CAUSE_INPUT_XPATH = "//input[@placeholder='Find your root cause']"

CLOSURE_CODE_SELECT_XPATH = (
    "//div[@class='controls controls-non-input']"
    "[normalize-space(text())='Closure Code']"
    "//select[@data-index='4']"
)

FINAL_RESOLVE_XPATH = "//input[@class='resolve btn-primary btn']"

RESOLVED_STATUS_XPATH = (
    "//span[contains(text(),'Resolved') or contains(text(),'Closed')]"
)

ACTIVITY_STREAM_XPATH = "//section[@class='document-stream']"

# =====================================================
# HELPERS
# =====================================================

def wait_for_issue_ready(driver):
    WebDriverWait(driver, WAIT).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    WebDriverWait(driver, WAIT).until(
        EC.presence_of_element_located((By.XPATH, ACTIVITY_STREAM_XPATH))
    )

def safe_js_click(driver, element):
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", element
    )
    time.sleep(1)
    driver.execute_script("arguments[0].click();", element)

# =====================================================
# LOAD EXCEL
# =====================================================

excel_files = glob.glob(os.path.join(EXCEL_FOLDER, "*.xls*"))
latest_excel = max(excel_files, key=os.path.getmtime)

df = pd.read_excel(latest_excel)
df["issue_url"] = BASE_URL + df["sim_issue_alias"].astype(str)
records = df.to_dict("records")

print(f"Total issues: {len(records)}")

# =====================================================
# DRIVER
# =====================================================

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, WAIT)

# =====================================================
# LOGIN
# =====================================================

driver.get(LOGIN_URL)
input("Login completed. Press ENTER to continue...")

main_tab = driver.current_window_handle
time.sleep(5)

# =====================================================
# CONSTANT COMMENT & ROOT CAUSE
# =====================================================

ROOT_CAUSE_TEXT = "Error-Competitor PARS Update-In Between Scheduled Crawls"

COMMENT_TEXT = (
    "Hi,\n\n"
    "Thanks for reaching out to us!\n\n"
    "On analysis we see that the competitor has updated the PARS. "
    "This was a competitor PARS update that occurred in between scheduled crawls. "
    "Correct PARS values have been reported post the subsequent crawl.\n\n"
    "Resolving the TT. Please feel free to re-open the ticket if you find any issue.\n\n"
    "Regards,\n"
    "Oncall Audits."
)

# =====================================================
# PROCESS
# =====================================================

for start in range(0, len(records), BATCH_SIZE):
    batch = records[start:start + BATCH_SIZE]

    tabs = []
    for row in batch:
        driver.execute_script("window.open(arguments[0]);", row["issue_url"])
        time.sleep(0.5)
        tabs.append(driver.window_handles[-1])

    # 🔹 BATCH LOAD WAIT
    print(f"⏳ Waiting {BATCH_LOAD_WAIT} seconds for batch to fully load...")
    time.sleep(BATCH_LOAD_WAIT)

    for tab in tabs:
        driver.switch_to.window(tab)

        try:
            wait_for_issue_ready(driver)

            # ---------- CHECK RESOLVE ----------
            resolve_btn = driver.find_element(By.XPATH, RESOLVE_BUTTON_XPATH)
            if resolve_btn.text.strip().lower() != "resolve":
                print("Skip: not resolvable")
                driver.close()
                time.sleep(SIM_PAUSE)
                continue

            # ---------- COMMENT ----------
            comment_box = wait.until(
                EC.element_to_be_clickable((By.XPATH, COMMENT_BOX_XPATH))
            )
            comment_box.clear()
            comment_box.send_keys(COMMENT_TEXT)

            wait.until(
                EC.element_to_be_clickable((By.XPATH, POST_BUTTON_XPATH))
            ).click()

            print("Comment posted")
            time.sleep(3)

            # ---------- RESOLVE ----------
            safe_js_click(driver, resolve_btn)
            time.sleep(2)

            root_input = wait.until(
                EC.element_to_be_clickable((By.XPATH, ROOT_CAUSE_INPUT_XPATH))
            )
            root_input.send_keys(ROOT_CAUSE_TEXT)
            time.sleep(1)
            root_input.send_keys("\n")

            closure_select = Select(
                wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, CLOSURE_CODE_SELECT_XPATH)
                    )
                )
            )
            closure_select.select_by_visible_text("Successful")

            final_resolve = wait.until(
                EC.element_to_be_clickable((By.XPATH, FINAL_RESOLVE_XPATH))
            )
            final_resolve.click()

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, RESOLVED_STATUS_XPATH)
                )
            )

            print("Resolved (confirmed)")
            time.sleep(2)
            driver.close()
            time.sleep(SIM_PAUSE)

        except TimeoutException:
            print("Failed / timed out → leaving tab open for review")
            driver.switch_to.window(main_tab)
            time.sleep(SIM_PAUSE)
            continue

    driver.switch_to.window(main_tab)

# =====================================================
# END
# =====================================================

input("\n✅ All issues processed.\nPress ENTER to close the browser...")
driver.quit()
