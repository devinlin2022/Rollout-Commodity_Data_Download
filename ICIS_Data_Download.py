import base64
import io
import time
import pandas as pd
import requests
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
import pymupdf
import pygsheets
from gspread_dataframe import set_with_dataframe
import gspread
import re
import json
import os
from selenium.webdriver.chrome.service import Service
from retrying import retry # Import the retrying library

try:
    gc_pygsheets = pygsheets.authorize(service_file='service_account_key.json')
    gc_gspread = gspread.service_account(filename='service_account_key.json')
except Exception as e:
    raise

CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH", "/usr/local/bin/chromedriver")
DOWNLOAD_DIR = "/tmp/downloads"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def get_chrome_driver():
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option('prefs', prefs)

    service = Service(executable_path=CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)
    return driver

@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException), stop_max_attempt_number=5, wait_fixed=2000)
def click_element_with_retry(driver, by_locator):
    wait = WebDriverWait(driver, 10)
    element = wait.until(EC.element_to_be_clickable(by_locator))
    element.click()

@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException), stop_max_attempt_number=5, wait_fixed=2000)
def js_click_element_with_retry(driver, css_selector):
    wait = WebDriverWait(driver, 10)
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_selector)))
    driver.execute_script("arguments[0].click();", element)


def save_pdf(driver, path, img_path):
    settings = {
        "landscape": False,
        "displayHeaderFooter": False,
        "printBackground": True,
        "preferCSSPageSize": True
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", settings)
    pdf_data = base64.b64decode(result['data'])
    with open(path, 'wb') as f:
        f.write(pdf_data)
    time.sleep(3)
    pdf_to_img(path, img_path)

def pdf_to_img(pdf_path, img_path):
    pdf_document = pymupdf.open(pdf_path)
    page = pdf_document[0]
    pixmap = page.get_pixmap(dpi=300)
    pixmap.save(img_path)

def fetch_data(link):
    driver = get_chrome_driver()
    driver.implicitly_wait(10)
    driver.get(link)

    wait = WebDriverWait(driver, 30)
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#login-button')))

    icis_username = os.getenv("ICIS_USERNAME")
    icis_password = os.getenv("ICIS_PASSWORD")

    if not icis_username or not icis_password:
        driver.quit()
        raise ValueError("ICIS_USERNAME or ICIS_PASSWORD environment variables not set.")

    driver.execute_script(f'document.querySelector("#username-input").value = "{icis_username}"')
    driver.execute_script(f'document.querySelector("#password-input").value = "{icis_password}"')
    
    try:
        click_element_with_retry(driver, (By.CSS_SELECTOR, '#login-button'))
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click ICIS login button: {e}")

    try:
        js_click_element_with_retry(driver, '#continue-login-button')
        time.sleep(2)
    except Exception as e:
        pass

    temp_dir = '/tmp'
    os.makedirs(temp_dir, exist_ok=True)
    pdf_output_path = os.path.join(temp_dir, "webpage_icis.pdf")
    img_output_path = os.path.join(temp_dir, "icis_screenshot.jpg")

    try:
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,
            '#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Largestyle__DisplayItem-vzpFY.fbUftf > div > div:nth-child(2) > div > div > div.PriceDeltastyle__DeltaContainer-jdFEoE.dtfcmD > div.Textstyles__Heading1Blue-gtxuIB.dzShK'
        )))
        save_pdf(driver, pdf_output_path, img_output_path)
    except Exception as e:
        pass

    element_scripts = [
        """return document.querySelector("#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Largestyle__DisplayItem-vzpFY.fbUftf > div > div:nth-child(1) > div > div > div.PriceDeltastyle__DeltaContainer-jdFEoE.dtfcmD > div.Textstyles__Heading1Blue-gtxuIB.dzShK").textContent;""",
        """return document.querySelector("#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Largestyle__DisplayItem-vzpFY.fbUftf > div > div:nth-child(2) > div > div > div.PriceDeltastyle__DeltaContainer-jdFEoE.dtfcmD > div.Textstyles__Heading1Blue-gtxuIB.dzShK").textContent;""",
        """return document.querySelector("#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Largestyle__DisplayItem-vzpFY.fbUftf > div > div:nth-child(3) > div > div > div.PriceDeltastyle__DeltaContainer-jdFEoE.dtfcmD > div.Textstyles__Heading1Blue-gtxuIB.dzShK").textContent;"""
    ]

    elements = ['Element Not Found'] * len(element_scripts)
    for i, script in enumerate(element_scripts):
        try:
            elements[i] = driver.execute_script(script)
        except Exception as e:
            elements[i] = 'Element Not Found'

    date = 'Date Not Found'
    try:
        date_element = driver.find_element(By.CSS_SELECTOR,
            '#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Mainstyle__Group-ciNpsy.fYvNPb > div > div > div:nth-child(2) > div')
        date = date_element.text
    except Exception as e:
        date = 'Date Not Found'

    driver.quit()
    return elements, date

def upload_to_google_sheet(data, sheet_key, worksheet_name, row):
    try:
        wb_key = gc_pygsheets.open_by_key(sheet_key)
        sheet = wb_key.worksheet_by_title(worksheet_name)

        if len(row) > 1 and row[1].strip() == "ICIS":
            original_date = data['Date'][0]
            try:
                formatted_date = datetime.strptime(original_date, '%d-%b-%y').strftime('%Y-%m-%d')
            except ValueError as e:
                return

            new_row = [formatted_date] + data['Commodity'][0]
        else:
            return

        all_values = sheet.get_all_values(returnas='matrix')
        last_non_empty_row = next((i for i, existing_row in reversed(list(enumerate(all_values))) if any(cell.strip() for cell in existing_row)), -1)

        new_row_index = max(last_non_empty_row + 1, 1)
        if new_row_index >= sheet.rows:
            sheet.add_rows(new_row_index - sheet.rows + 100)

        sheet.update_row(new_row_index + 1, new_row)
    except Exception as e:
        pass

@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException) or isinstance(e, ValueError), stop_max_attempt_number=5, wait_fixed=2000)
def fetch_RISI_data(link, search_term, pdf_output_path, img_output_path):
    driver = get_chrome_driver()
    driver.implicitly_wait(10)
    driver.get(link)

    wait = WebDriverWait(driver, 100)
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#userEmail')))

    risi_username = os.getenv("RISI_USERNAME")
    risi_password = os.getenv("RISI_PASSWORD")

    if not risi_username or not risi_password:
        driver.quit()
        raise ValueError("RISI_USERNAME or RISI_PASSWORD environment variables not set.")

    driver.execute_script(f'document.querySelector("#userEmail").value = "{risi_username}"')
    driver.execute_script(f'document.querySelector("#password").value = "{risi_password}"')
    
    try:
        click_element_with_retry(driver, (By.CSS_SELECTOR, '#login-button'))
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click RISI login button: {e}")
        
    time.sleep(5)

    try:
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#mat-mdc-chip-list-input-0')))
        input_element = driver.execute_script("return document.querySelector('#mat-mdc-chip-list-input-0')")

        input_element.clear()
        input_element.send_keys(search_term, Keys.ENTER)
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed during RISI search input: {e}")

    time.sleep(5)
    try:
        js_click_element_with_retry(driver, 'body > fui-root-container > fui-root > fui-dashboard-container > fui-dashboard > main > header > fui-workspace-header-container > fui-dashboard-header-container > fui-dashboard-header > div.header-section.search.draggable > fui-search-bar > fui-dropdown-container > div > div.content > div > fui-global-search-dropdown > div > section.results-view.ng-star-inserted > div > fui-instrument-search-result-item > fui-search-result-item')
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click RISI search result: {e}")

    time.sleep(5)
    try:
        click_element_with_retry(driver, (By.CSS_SELECTOR, "#mat-tab-label-1-3"))
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click RISI tab: {e}")

    time.sleep(5)
    save_pdf(driver, pdf_output_path, img_output_path)
    driver.close()
    return True

def update_google_sheet_from_csv(csv_path, sheet_id, sheet_name):
    df = pd.read_csv(csv_path)

    df.columns = df.iloc[0]
    df = df[1:]
    df.reset_index(drop=True, inplace=True)

    df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], errors='coerce')
    df = df.dropna(subset=[df.columns[0]])

    df[df.columns[0]] = df[df.columns[0]].dt.strftime('%Y-%m-%d')

    df = df.sort_values(by=df.columns[0])

    sh = gc_gspread.open_by_key(sheet_id)
    worksheet = sh.worksheet(sheet_name)

    existing_data = pd.DataFrame(worksheet.get_all_values())

    if not existing_data.empty:
        if existing_data.shape[0] > 0:
            existing_data.columns = existing_data.iloc[0]
            existing_data = existing_data[1:]
            existing_data.reset_index(drop=True, inplace=True)
        else:
            existing_data = pd.DataFrame(columns=df.columns)

    new_rows = df

    if not existing_data.empty and df.columns[0] in existing_data.columns:
        records_to_add = new_rows[
            ~new_rows[df.columns[0]].isin(existing_data[existing_data.columns[0]].astype('str'))]
    else:
        records_to_add = new_rows

    if not records_to_add.empty:
        set_with_dataframe(worksheet, records_to_add, include_index=False,
                            include_column_header=False, row=len(existing_data) + 2)
    else:
        pass

class ImageService:
    def __init__(self, api_key):
        self.api_key = api_key

    def get_image_base64_string(self, image_path):
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')

    def send_post_request(self, url, json_content):
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        response = requests.post(url, data=json_content, headers=headers)
        if 200 <= response.status_code < 300:
            return response.text
        else:
            raise Exception(f"Request failed: {response.status_code}, {response.text}")

    def get_results(self, image_path):
        url = "https://dashscope.aliyuncs.com/compatible-model/v1/chat/completions"
        encoded_image = self.get_image_base64_string(image_path)
        json_content = {
            "model": "qwen-vl-plus",
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": "抓取图片中表格。"},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{encoded_image}"}}
                    ]
                }
            ]
        }
        json_content_str = json.dumps(json_content)
        result = self.send_post_request(url, json_content_str)
        return result

    def parse_table_and_save_to_csv(self, response_body, csv_path):
        data = json.loads(response_body)
        if 'choices' in data and len(data['choices']) > 0:
            message = data['choices'][0]['message']['content']
            table_lines = re.findall(r"\|.*\|", message)

            if table_lines:
                rows = []
                start_row_index = 0
                if len(table_lines) > 1 and all(c in '|- ' for c in table_lines[1]):
                    start_row_index = 2
                elif len(table_lines) > 0:
                    start_row_index = 1

                for line in table_lines[start_row_index:]:
                    columns = [col.strip() for col in line.split('|')[1:-1]]
                    rows.append(columns)

                header = [col.strip() for col in table_lines[0].split('|')[1:-1]]
                table_df = pd.DataFrame(rows, columns=header)
                table_df.to_csv(csv_path, index=False)
            else:
                pass
        else:
            pass

def get_dynamic_parameters(sheet_id):
    sh = gc_gspread.open_by_key(sheet_id)
    worksheet = sh.worksheet("ALL_EU_AEA")

    data = worksheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])

    filtered_df = df[df['Tab Name'].str.startswith("RISI")]

    return filtered_df[['Address', 'Tab Name', 'Series code']].to_dict('records')

def risi_main():
    sheet_id = '1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE'

    rows = get_dynamic_parameters(sheet_id)

    api_key = os.getenv("DASHSCOPE_API_KEY")
    if not api_key:
        raise ValueError("DASHSCOPE_API_KEY environment variable not set.")

    temp_dir = '/tmp'
    os.makedirs(temp_dir, exist_ok=True)
    pdf_path = os.path.join(temp_dir, 'webpage_risi.pdf')
    image_path = os.path.join(temp_dir, '01.jpg')
    csv_path = os.path.join(temp_dir, 'extracted_table.csv')

    service = ImageService(api_key)

    for row in rows:
        link = row['Address']
        sheet_name = row['Tab Name']
        search_term = row['Series code']

        try:
            fetch_RISI_data(link, search_term, pdf_path, image_path)
        except Exception as e:
            continue

        try:
            result = service.get_results(image_path)
            service.parse_table_and_save_to_csv(result, csv_path)
        except Exception as e:
            pass

        update_google_sheet_from_csv(csv_path, sheet_id, sheet_name)

@retry(stop_max_attempt_number=5, wait_fixed=5000) # Retry the main function 5 times with 5-second delay
def main(id, tab_name):
    sh = gc_pygsheets.open_by_key(id)
    wks = sh.worksheet_by_title(tab_name)

    data = wks.get_all_values()

    sheet_key = '1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE'
    worksheet_name_col_idx = 8
    link_col_idx = 5

    for row_num, row in enumerate(data[1:], start=2):
        if len(row) <= max(worksheet_name_col_idx, link_col_idx):
            continue

        worksheet_name = row[worksheet_name_col_idx]
        link = row[link_col_idx]

        if not worksheet_name or not link:
            continue

        if len(row) > 1 and row[1].strip() == "ICIS":
            try:
                price, date = fetch_data(link)
            except Exception as e:
                # If fetch_data (ICIS) fails after its own retries,
                # then proceed to call risi_main as a fallback for this specific row.
                risi_main() # risi_main also has its own retries inside fetch_RISI_data
                continue

            if any('Element Not Found' in p for p in price) or date == 'Date Not Found':
                risi_main()
                continue

            price_data = pd.DataFrame([[date, price]], columns=['Date', 'Commodity'])

            upload_to_google_sheet(price_data, sheet_key, worksheet_name, row)
        else:
            pass

if __name__ == "__main__":
    main('1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE', 'ALL_EU_AEA')
