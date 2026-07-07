from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
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
        "\nPress ENTER to use today's file OR type full file name\n (e.g. FMINF_04_28_pradsush.xlsx): "
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
        if f.startswith(f"FMINF_{today}") and f.endswith(".xlsx")
    ]

    if not files:
        raise Exception(f"No file found for today: FMINF_{today}_*.xlsx")

    files.sort(key=lambda x: os.path.getmtime(os.path.join(folder_path, x)), reverse=True)

    selected = files[0]
    logger.info(f"Using today's file: {selected}")
    return os.path.join(folder_path, selected), selected


def copy_output_file(source_path, file_name):
    dest_folder = r"W:\Shared With Me\INF Audit Output - 2026\FMINF\June\Audit output files"

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


def clean_dataframe(df):
    df['IS_NA_CASE'] = False

    for i in range(len(df)):
        if pd.isna(df.at[i, 'GL']) or str(df.at[i, 'GL']).strip() == '#N/A':

            df.at[i, 'IS_NA_CASE'] = True

            for col in ['GL', 'PL', 'Mapper Reason code', 'Mapped By']:
                df.at[i, col] = '-'

            df.at[i, 'Manager ID'] = 'nobody'
            df.at[i, 'POC'] = 'nobody'

            # next valid mapped date
            next_date = None
            for j in range(i + 1, len(df)):
                if pd.notna(df.at[j, 'Mapped Date']):
                    next_date = df.at[j, 'Mapped Date']
                    break

            if next_date:
                df.at[i, 'Mapped Date'] = next_date

    return df


def process_excel_file(path):
    df = pd.read_excel(path, sheet_name="Sheet2")

    if 'SIM Url' not in df.columns:
        df['SIM Url'] = ''

    df = normalize_date_columns(df)
    df = clean_dataframe(df)

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

    # ---------- FORMAT ----------
    def format_sim_data(self, row):

        description = f"""Hi Team,

You have received a request for RCA from Business Quality-INF Audits. Please refer to attachment for more details.

Regards,
{row['Auditor ID']}


Final Search Technique --- {row['Final Search Technique']}"""

        return {
            'title': f"{row['MKPL']} | {row['Competitor']} | FMINF Error-RCA",
            'description': description,
            'MKPL': str(row['MKPL']),
            'ASIN': str(row['ASIN']),
            'Competitor': str(row['Competitor']),
            'GL': str(row['GL']),
            'PL': str(row['PL']),
            'Mapper Reason code': str(row['Mapper Reason code']),
            'Mapper ID': str(row['Mapped By']),
            'mapped_date': str(row['Mapped Date']),
            'Manager ID': str(row['Manager ID']),
            'Auditor ID': str(row['Auditor ID']),
            'audit_date': str(row['Audit Date']),
            'Audit Status': str(row['Audit Status (Error/No Error)']),
            'Comp URL': str(row['Comp URL']),
            'Search technique': "Navigate To Issue Description Area",
            'POC': str(row['POC'])
        }

    # ---------- OPEN ----------
    def open_all_sims(self, df):
        logger.info("Opening first SIM page for login...")
        self.driver.get("https://issues.amazon.com/issues/create?assignedFolder=c7bde71a-32ab-48cd-af7f-9ba11ba841b9")
        self.tabs.append({'handle': self.driver.window_handles[0], 'index': 0})

        input("Please login and press Enter once completed...")
        logger.info("Login completed, opening remaining tabs...")

        for i in range(1, len(df)):
            self.driver.execute_script("window.open('');")
            self.driver.switch_to.window(self.driver.window_handles[i])
            self.driver.get("https://issues.amazon.com/issues/create?assignedFolder=c7bde71a-32ab-48cd-af7f-9ba11ba841b9")
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
                arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
                arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));
            """, title_input, data['title'])

            # description
            desc_input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-field='description']//textarea")))
            self.driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input',{bubbles:true}));
                arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
                arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));
            """, desc_input, data['description'])

            # severity
            severity = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-field='impact']//select")))
            self.driver.execute_script("""
                arguments[0].value='3';
                arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
            """, severity)

            # custom fields
            custom_fields = {
                "input[data-index='0']": data['MKPL'],
                "input[data-index='1']": data['ASIN'],
                "input[data-index='2']": data['Competitor'],
                "input[data-index='3']": data['GL'],
                "input[data-index='4']": data['PL'],
                "textarea[data-index='5']": data['Mapper Reason code'],
                "input[data-index='6']": data['Mapper ID'],
                "input[data-index='7']": data['mapped_date'],
                "input[data-index='8']": data['Manager ID'],
                "input[data-index='9']": data['Auditor ID'],
                "input[data-index='10']": data['audit_date'],
                "input[data-index='11']": data['Audit Status'],
                "textarea[data-index='12']": data['Comp URL'],
                "textarea[data-index='13']": data['Search technique']
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
            watchers_text = f"{data['Mapper ID']}, {data['Manager ID']}, {data['POC']}"
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
            assignee_input.send_keys(data['Manager ID'])

            option = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//li[contains(@data-value,'{data['Manager ID']}')]")))
            self.driver.execute_script("arguments[0].click();", option)

            return True

        except Exception as e:
            logger.error(f"Error in fill_form: {e}")
            return False

    def fill_all_forms(self, df):
        logger.info("Filling all forms...")
        for i, row in df.iterrows():
            self.driver.switch_to.window(self.tabs[i]['handle'])
            self.fill_form(self.format_sim_data(row))

    # ---------- RESOLVE ----------
    def resolve_na_sims(self, df):
        logger.info("Resolving NA SIMs → Sibling Variant")

        for i, row in df.iterrows():
            if not row['IS_NA_CASE']:
                continue

            self.driver.switch_to.window(self.tabs[i]['handle'])

            if "/issues/" not in self.driver.current_url:
                logger.warning(f"Tab {i+1} not created yet → skipping")
                continue

            try:
                self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(@class,'resolve-issue')]"))).click()

                dropdown = self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//div[@class='custom-field input-wrapper']//select")))
                Select(dropdown).select_by_visible_text("Sibling Variant")

                self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[contains(@class,'resolve')]"))).click()

                logger.info(f"Resolved tab {i+1}")

            except Exception as e:
                logger.error(f"Resolve failed tab {i+1}: {e}")

    # ---------- CAPTURE ----------
    def capture_all_sim_urls(self, df, path):
        logger.info("Capturing SIM URLs...")
        created, pending = 0, 0

        for i in range(len(self.tabs)):
            self.driver.switch_to.window(self.tabs[i]['handle'])
            url = self.driver.current_url
            df.at[i, 'SIM Url'] = url

            if "/issues/" in url and "/create" not in url:
                created += 1
            else:
                pending += 1

        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Sheet2', index=False)

        logger.info(f"Saved URLs → Created: {created}, Pending: {pending}")

# ---------------- MAIN ----------------

def main():
    folder = r"C:\Users\pradsush\Documents\BQ_INF_FM"
    file_path, file_name = get_input_file(folder)

    df = process_excel_file(file_path)

    sim = SIMCreator()

    if sim.open_all_sims(df):
        sim.fill_all_forms(df)

        while True:
            cmd = input("\nEnter tab number / u / q: ")

            if cmd.lower() == 'q':
                break

            elif cmd.lower() == 'u':
                confirm = input("Have you created all the issues? (y/n): ")

                if confirm.lower() == 'y':
                    sim.resolve_na_sims(df)
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