from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class IFAuditSIMCreator:
    def __init__(self):
        self.setup_driver()
        self.wait = WebDriverWait(self.driver, 90)
        self.tabs = []

    def setup_driver(self):
        chrome_options = Options()
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--start-maximized')
        self.driver = webdriver.Chrome(options=chrome_options)

    def format_sim_data(self, row):
        return {
            'title': f"IF Audit | {row['GL']} | {row['Mapping Channel']}",
            'description': """Hi Team,

You have received a request for RCA from Business Quality-IF Audits. Please refer to Information tab for more details.

Regards,
BQ IF Audit""",
            'process': 'IF Audit',
            'auditor_id': str(row['Auditor ID']),
            'mapper_id': str(row['Mapper']),
            'manager_id': str(row['Manager']),
            'asin': str(row['ASIN']),
            'competitor': str(row['Competitor']),
            'url': str(row['New URL']),
            'audit_date': row['Audit date'].strftime('%m/%d/%Y'),
            'mapped_date': row['Mapped Date'].strftime('%m/%d/%Y'),
            'source_mapper_id': str(row['Source mapper']),
            'source_asin': str(row['Source ASIN']),
            'gl': str(row['GL']),
            'reason_code': str(row['Reason code']),
            'remarks': str(row['Remarks']),
            'watchers': str(row['Watcherlist'] +',' + row['Manager']),
            'assignee': str(row['Assignee'])
        }

    def open_all_sims(self, df):
        try:
            logger.info("Opening first SIM page for login...")
            self.driver.get(
                "https://issues.amazon.com/issues/create?assignedFolder=06ab5dcb-4db6-4cd4-a820-5b741bcc2e0d")
            self.tabs.append({
                'handle': self.driver.window_handles[0],
                'index': 0
            })

            input("Please login and press Enter once completed...")
            logger.info("Login completed, proceeding to open other tabs...")

            # Add a delay after login to ensure session is properly established
            time.sleep(3)

            for index in range(1, len(df)):
                self.driver.execute_script("window.open('');")
                new_handle = self.driver.window_handles[index]
                self.driver.switch_to.window(new_handle)

                # Add explicit wait for page load
                self.driver.get(
                    "https://issues.amazon.com/issues/create?assignedFolder=06ab5dcb-4db6-4cd4-a820-5b741bcc2e0d")

                # Wait for a key element to be present to ensure page is loaded
                try:
                    self.wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-field='title']//input")))
                except:
                    logger.warning(f"Page {index + 1} might not have loaded completely. Retrying...")
                    self.driver.refresh()
                    self.wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-field='title']//input")))

                self.tabs.append({
                    'handle': new_handle,
                    'index': index
                })

                # Add delay between opening tabs
                time.sleep(2)

                logger.info(f"Opened SIM page {index + 1}")

            return True
        except Exception as e:
            logger.error(f"Error in open_all_sims: {e}")
            return False

    def fill_form(self, form_data):
        try:
            time.sleep(1)

            # Fill title
            title_input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-field='title']//input")))
            self.driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
            """, title_input, form_data['title'])

            # Fill description
            description_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-field='description']//textarea")))
            self.driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
            """, description_field, form_data['description'])

            # Handle severity
            severity_dropdown = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-field='impact']//select")))
            self.driver.execute_script("""
                arguments[0].value = '3';
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, severity_dropdown)

            # Custom fields mapping
            custom_fields = {
                "input[data-index='0']": form_data['process'],
                "input[data-index='1']": form_data['auditor_id'],
                "input[data-index='2']": form_data['mapper_id'],
                "input[data-index='3']": form_data['manager_id'],
                "input[data-index='4']": form_data['asin'],
                "input[data-index='5']": form_data['competitor'],
                "input[data-index='6']": form_data['url'],
                "input[data-index='7']": form_data['audit_date'],
                "input[data-index='8']": form_data['mapped_date'],
                "input[data-index='9']": form_data['source_mapper_id'],
                "input[data-index='10']": form_data['source_asin'],
                "input[data-index='11']": form_data['gl'],
                "input[data-index='12']": form_data['reason_code'],
                "input[data-index='13']": form_data['remarks']
            }

            for selector, value in custom_fields.items():
                element = self.driver.find_element(By.CSS_SELECTOR, selector)
                self.driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """, element, str(value))

            # Handle watchers
            watchers_link = self.driver.find_element(By.XPATH, "//a[@data-show-field='watchers']")
            self.driver.execute_script("arguments[0].click();", watchers_link)
            watchers_input = self.wait.until(EC.presence_of_element_located((By.XPATH, "//input[@id='issue-watchers']")))
            self.driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, watchers_input, form_data['watchers'])

            # Handle assignee
            assignee_link = self.driver.find_element(By.XPATH, "//a[@data-show-field='assigneeIdentity']")
            self.driver.execute_script("arguments[0].click();", assignee_link)
            time.sleep(0.5)

            assignee_selector = self.wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[@data-csm-counter='createViewAssigneeSelector']")))
            self.driver.execute_script("arguments[0].click();", assignee_selector)
            time.sleep(0.5)

            assignee_input = self.wait.until(EC.presence_of_element_located((
                By.XPATH, "//input[@class='assignee-input input-medium']")))
            assignee_input.clear()
            assignee_input.send_keys(form_data['assignee'])
            time.sleep(0.5)

            poc_xpath = f"//li[@data-value='kerberos:{form_data['assignee']}@ANT.AMAZON.COM']"
            poc_option = self.wait.until(EC.element_to_be_clickable((By.XPATH, poc_xpath)))
            self.driver.execute_script("arguments[0].click();", poc_option)

            return True

        except Exception as e:
            logger.error(f"Error in fill_form: {e}")
            return False

    def capture_all_sim_urls(self, df):
        logger.info("Capturing all URLs...")
        created_sims = 0
        pending_sims = 0

        for index in range(len(self.tabs)):
            try:
                self.driver.switch_to.window(self.tabs[index]['handle'])
                current_url = self.driver.current_url
                df.at[index, 'SIM Link'] = current_url

                if '/issues/' in current_url and '/create' not in current_url:
                    created_sims += 1
                    logger.info(f"Tab {index + 1}: Created SIM - {current_url}")
                else:
                    pending_sims += 1
                    logger.info(f"Tab {index + 1}: Pending Creation - {current_url}")

            except Exception as e:
                logger.error(f"Error with tab {index + 1}: {e}")
                df.at[index, 'SIM Link'] = f"Error: {str(e)}"

        try:
            df.to_excel('IF_AUDIT_SIM.xlsx', index=False)
            logger.info(f"Saved all URLs to Excel: {created_sims} created SIMs, {pending_sims} pending")
        except Exception as e:
            logger.error(f"Error saving to Excel: {e}")

        return True

    def fill_all_forms(self, df):
        logger.info("Starting to fill all forms...")
        for index, row in df.iterrows():
            logger.info(f"Filling form {index + 1}")
            self.driver.switch_to.window(self.tabs[index]['handle'])
            form_data = self.format_sim_data(row)
            success = self.fill_form(form_data)
            if not success:
                logger.warning(f"Failed to fill form {index + 1}")
        logger.info("All forms filled. You can now create the issues.")
        return True

def process_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        if 'SIM Link' not in df.columns:
            df['SIM Link'] = ''
        df.to_excel(file_path, index=False)
        logger.info(f"Loaded {len(df)} rows from Excel file")
        return df
    except Exception as e:
        logger.error(f"Error loading Excel file: {e}")
        raise

def main():
    sim_creator = None
    try:
        df = process_excel_file('IF_AUDIT_SIM.xlsx')
        sim_creator = IFAuditSIMCreator()

        if sim_creator.open_all_sims(df):
            logger.info("All SIM pages loaded successfully")
            time.sleep(4)
            sim_creator.fill_all_forms(df)

            print("\nOptions available:")
            print("1. Enter tab number to switch between tabs")
            print("2. Enter 'u' to collect URLs after creating issues")
            print("3. Enter 'q' to quit")

            while True:
                command = input("\nEnter command (tab number/u/q): ")

                if command.lower() == 'q':
                    break
                elif command.lower() == 'u':
                    # Add confirmation before collecting URLs
                    confirm = input("Have you created all the issues? (y/n): ")
                    if confirm.lower() == 'y':
                        sim_creator.capture_all_sim_urls(df)
                else:
                    try:
                        tab_num = int(command) - 1
                        if 0 <= tab_num < len(sim_creator.tabs):
                            sim_creator.driver.switch_to.window(sim_creator.tabs[tab_num]['handle'])
                            logger.info(f"Switched to tab {tab_num + 1}")
                        else:
                            logger.warning("Invalid tab number")
                    except ValueError:
                        logger.warning("Invalid input")

    except Exception as e:
        logger.error(f"An error occurred: {e}")
        raise

    finally:
        if sim_creator and hasattr(sim_creator, 'driver'):
            sim_creator.driver.quit()

if __name__ == "__main__":
    main()
