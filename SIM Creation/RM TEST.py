from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import logging
import os
import shutil

# ---------------- LOGGING ----------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ---------------- FILE HANDLING ----------------

def get_input_file(folder_path):
    user_input = input(
        "PressENTER to use today's file OR type full file name\n (e.g. RMINF_04_28_pradsush.xlsx): "
    ).strip()

    if user_input:
        path = os.path.join(folder_path, user_input)
        if not os.path.exists(path):
            raise Exception(f"File not found: {user_input}")
        logger.info(f"Using user provided file: {user_input}")
        return path, user_input

    today = pd.Timestamp.today().strftime("%m_%d")

    files = [
        f for f in os.listdir(folder_path)
        if f.startswith(f"RMINF_{today}") and f.endswith(".xlsx")
    ]

    if not files:
        raise Exception(f"No file found for today: RMINF_{today}_*.xlsx")

    files.sort(key=lambda x: os.path.getmtime(os.path.join(folder_path, x)), reverse=True)

    selected = files[0]
    logger.info(f"Using today's file: {selected}")
    return os.path.join(folder_path, selected), selected

def copy_output_file(source_path, file_name):
    dest_folder = r"W:\Shared With Me\INF Audit Output - 2026\RMINF\June\Audit Output Files"

    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)

    dest_path = os.path.join(dest_folder, file_name)
    shutil.copy2(source_path, dest_path)

    logger.info(f"File copied to: {dest_path}")

# ---------------- DATA PROCESSING ----------------

def normalize_date_columns(df):
    for col in ['Mapped Date', 'Audit Date']:
        df[col] = pd.to_datetime(df[col], errors='coerce')
        df[col] = df[col].dt.strftime('%m/%d/%Y')
        df[col] = df[col].str.lstrip('0').str.replace('/0', '/', regex=False)
    return df

def process_excel_file(path):
    df = pd.read_excel(path, sheet_name="Sheet2")

    if 'SIM Link' not in df.columns:
        df['SIM Link'] = ''

    #✅ BUG FIX: Force SIM Link column to object dtype
    # (prevents dtype crash when column is empty/all NaN)
    df['SIM Link'] = df['SIM Link'].astype('object')

    df = normalize_date_columns(df)

    # SAFE WRITE (preserve Sheet1)
    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Sheet2', index=False)

    return df

# ---------------- SIM CREATOR ----------------

class SIMCreator:
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--start-maximized')

        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 90)
        self.tabs = []

    #---------- FORMAT ----------
    def format_sim_data(self, row):

        description = f"""Hi Team,

You have received a request for RCA from Business Quality-INF Audits. Please refer to attachment for more details.

Regards,
{row['Auditor ID']}

Final Search Technique --- {row['Final Search Technique']}"""

        return {
            'title': f"{row['MKPL']} | {row['Competitor']} | RMINF Error",
            'description': description,
            'competitor': str(row['Competitor']),
            'competitor_url': str(row['URL']),
            'reason_code': str(row['Reason code']),
            'mapper_id': str(row['Mapped By']),
            'mapped_date': str(row['Mapped Date']),
            'manager_id': str(row['Manager ID']),
            'auditor_id': str(row['Auditor ID']),
            'audit_date': str(row['Audit Date']),
            'audit_status': str(row['Audit Status (Error/No Error)']),
            'asin': str(row['ASIN']),
            'final_search_technique': "Navigate To Issue Description Area",
            'POC': str(row['Poc'])
        }

    # ---------- OPEN ----------
    def open_all_sims(self, df):
        logger.info("Opening first SIM page for login...")
        self.driver.get("https://issues.amazon.com/issues/create?assignedFolder=e63c6076-7336-40c3-bd62-15094adc4fb5")
        self.tabs.append({'handle': self.driver.window_handles[0], 'index': 0})

        input("Please login and press Enter once completed...")
        logger.info("Login completed, opening remaining tabs...")

        for i in range(1, len(df)):
            self.driver.execute_script("window.open('');")
            self.driver.switch_to.window(self.driver.window_handles[i])
            self.driver.get("https://issues.amazon.com/issues/create?assignedFolder=e63c6076-7336-40c3-bd62-15094adc4fb5")
            self.tabs.append({'handle': self.driver.window_handles[i], 'index': i})
            logger.info(f"Opened SIM page {i+1}")

        return True

    # ---------- FILL ----------
    def fill_form(self, data):
        try:
            time.sleep(1)

            # title
            title_input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-field='title']//input")))
            self.driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input',{bubbles:true}));
                arguments[0].dispatchEvent(new Event('change',{bubbles:true}));arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));
            """, title_input, data['title'])

            # description
            desc_input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-field='description']//textarea")))
            self.driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input',{bubbles:true}));
                arguments[0].dispatchEvent(new Event('change',{bubbles:true}));arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));
            """, desc_input, data['description'])

            # severity
            severity = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-field='impact']//select")))
            self.driver.execute_script("""
                arguments[0].value='3';
                arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
            """, severity)

            # custom fields (RMINF-specific mapping)
            custom_fields = {
                ".custom-field input[data-index='0']": data['competitor'],
                "textarea[data-index='1']": data['competitor_url'],
                "textarea[data-index='2']": data['reason_code'],
                "input[data-index='3']": data['mapper_id'],
                "input[data-index='4']": data['mapped_date'],
                "input[data-index='5']": data['manager_id'],
                "input[data-index='6']": data['auditor_id'],
                "input[data-index='7']": data['audit_date'],
                "input[data-index='8']": data['audit_status'],
                "input[data-index='9']": data['asin'],
                "textarea[data-index='10']": data['final_search_technique']
            }

            for sel, val in custom_fields.items():
                el = self.driver.find_element(By.CSS_SELECTOR, sel)
                self.driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input',{bubbles:true}));
                    arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
                """, el, val)

            # WATCHERS
            watchers_link = self.driver.find_element(By.XPATH, "//a[@data-show-field='watchers']")
            self.driver.execute_script("arguments[0].click();", watchers_link)

            watchers_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "issue-watchers")))
            watchers_text = f"{data['mapper_id']}, {data['manager_id']}, {data['POC']}"
            self.driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input',{bubbles:true}));
                arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
            """, watchers_input, watchers_text)

            # ASSIGNEE
            assignee_link = self.driver.find_element(By.XPATH, "//a[@data-show-field='assigneeIdentity']")
            self.driver.execute_script("arguments[0].click();", assignee_link)

            selector = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[@data-csm-counter='createViewAssigneeSelector']")))
            self.driver.execute_script("arguments[0].click();", selector)

            assignee_input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@class='assignee-input input-medium']")))
            assignee_input.clear()
            assignee_input.send_keys(data['manager_id'])

            option = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//li[contains(@data-value,'{data['manager_id']}')]")))
            self.driver.execute_script("arguments[0].click();", option)

            return True

        except Exception as e:
            logger.error(f"Error in fill_form: {e}")
            return False

    def fill_all_forms(self, df):
        logger.info("Filling all forms...")
        for i, row in df.iterrows():
            self.driver.switch_to.window(self.tabs[i]['handle'])
            self.fill_form(self.format_sim_data(row))# ---------- CAPTURE ----------
    def capture_all_sim_urls(self, df, path):
        logger.info("Capturing SIM URLs...")
        created, pending = 0, 0

        for i in range(len(self.tabs)):
            self.driver.switch_to.window(self.tabs[i]['handle'])
            url = self.driver.current_url
            df.at[i, 'SIM Link'] = str(url)   # ✅ BUG FIX: explicit str() cast for dtype safety

            if "/issues/" in url and "/create" not in url:
                created += 1
            else:
                pending += 1
        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Sheet2', index=False)

        logger.info(f"Saved URLs -> Created: {created}, Pending: {pending}")

# ---------------- MAIN ----------------

def main():
    folder = r"C:\Users\pradsush\Documents\BQ_INF_FM"
    file_path, file_name = get_input_file(folder)

    df = process_excel_file(file_path)

    sim = SIMCreator()

    if sim.open_all_sims(df):
        sim.fill_all_forms(df)

        while True:
            cmd = input("Enter tab number / u / q: ")

            if cmd.lower() == 'q':
                break

            elif cmd.lower() == 'u':
                confirm = input("Have you created all the issues? (y/n): ")

                if confirm.lower() == 'y':
                    sim.capture_all_sim_urls(df, file_path)
                    copy_output_file(file_path, file_name)
                else:
                    logger.info("Waiting for user to create SIMs...")

            else:
                try:
                    sim.driver.switch_to.window(sim.tabs[int(cmd)-1]['handle'])
                    logger.info(f"Switched to tab {cmd}")
                except:
                    logger.warning("Invalid tab")

    sim.driver.quit()

if __name__ == "__main__":
    main()
