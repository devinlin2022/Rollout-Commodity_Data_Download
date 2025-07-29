import os
import time
import pandas as pd
import io
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pygsheets
from collections import Counter

# --- Configuration ---
RISI_USERNAME = os.getenv('RISI_USERNAME')
RISI_PASSWORD = os.getenv('RISI_PASSWORD')
SERVICE_ACCOUNT_FILE = 'service_account_key.json'
GSHEET_ID = '1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE'
GSHEET_TITLE = 'Sheet12' # The error log showed 'Sheet11'
CHROMEDRIVER_PATH = os.getenv('CHROMEDRIVER_PATH', '/usr/bin/chromedriver')
DOWNLOAD_DIR = "/tmp/downloads" # For error screenshots

def scrape_table_data(link):
    """
    抓取原始表格数据，此函数保持不变。
    """
    options = Options()
    options.binary_location = '/usr/bin/chromium-browser'
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument("--start-maximized")

    service = webdriver.chrome.service.Service(executable_path=CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(link)
        wait = WebDriverWait(driver, 60)

        print("Waiting for login page...")
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#userEmail'))).send_keys(RISI_USERNAME)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#password'))).send_keys(RISI_PASSWORD)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#login-button'))).click()
        print("Login successful.")

        print("Waiting for the data grid to load...")
        grid_selector = (By.CSS_SELECTOR, 'div[role="treegrid"]')
        grid_container = wait.until(EC.visibility_of_element_located(grid_selector))
        print("AG-Grid container found. Extracting all raw cell data...")

        fixed_headers = [
            'Month', 'Crude Palm Oil Malaysia', 'RBD Palm Stearin MY',
            'RBD Palm Kernel MY', 'Coconut Oil', 'Crude CNO', 'Tallow',
            'Soybean Oil 1st', 'Soybean Oil 2nd', 'Soybean Oil 3rd'
        ]
        num_expected_columns = len(fixed_headers)
        
        # 这里的逻辑保持不变，它会抓取到分离的行
        all_rows_elements = grid_container.find_elements(By.CSS_SELECTOR, '[role="row"]')
        data_rows = []
        for r in all_rows_elements:
            cells = r.find_elements(By.CSS_SELECTOR, '[role="gridcell"]')
            # 过滤掉空的header行或不完整的行
            if cells and len(cells) <= num_expected_columns:
                data_rows.append([c.text for c in cells])

        if not data_rows:
            raise ValueError("Scraping failed: No raw data rows were found.")
        
        # 创建原始的、未处理的DataFrame
        # 注意：这里的列名只是一个临时的占位符，因为数据是错位的
        temp_cols = [f'col_{i}' for i in range(num_expected_columns)]
        raw_df = pd.DataFrame(data_rows, columns=temp_cols[:len(data_rows[0])])
        
        print("Successfully created raw DataFrame. It will be processed next.")
        print(raw_df.head())
        
        return raw_df

    except Exception as e:
        print(f"An error occurred during scraping: {e}")
        error_screenshot_path = os.path.join(DOWNLOAD_DIR, 'error_screenshot.png')
        if not os.path.exists(DOWNLOAD_DIR):
            os.makedirs(DOWNLOAD_DIR)
        driver.save_screenshot(error_screenshot_path)
        print(f"Error screenshot saved to: {error_screenshot_path}")
        raise
    finally:
        print("Closing the browser.")
        driver.quit()

def process_and_clean_data(raw_df):
    """
    🔧 新增函数：清理和重组DataFrame。
    将日期行和数据行合并，并格式化日期。
    """
    print("Processing and cleaning the raw data...")
    if raw_df is None or raw_df.empty:
        print("Raw DataFrame is empty, skipping processing.")
        return pd.DataFrame()

    # 1. 计算分割点（总行数的一半）
    num_rows = len(raw_df)
    if num_rows % 2 != 0:
        print(f"Warning: The number of rows ({num_rows}) is odd. Data might be incomplete.")
        return pd.DataFrame() # 返回空表以避免后续错误
        
    half_point = num_rows // 2
    
    # 2. 提取日期部分和数据部分
    # 日期在前半部分的第1列
    dates = raw_df.iloc[:half_point, 0].reset_index(drop=True)
    # 数据在后半部分，从第1列开始的所有列
    # 因为原网站结构问题，数据行的第一列是空的，所以我们从第二列开始取
    numeric_data = raw_df.iloc[half_point:].reset_index(drop=True)
    
    # 3. 将日期和数据水平合并成一个新的DataFrame
    clean_df = pd.concat([dates, numeric_data], axis=1)

    # 4. 设置正确的列名
    clean_df.columns = [
        'Month', 'Crude Palm Oil Malaysia', 'RBD Palm Stearin MY',
        'RBD Palm Kernel MY', 'Coconut Oil', 'Crude CNO', 'Tallow',
        'Soybean Oil 1st', 'Soybean Oil 2nd', 'Soybean Oil 3rd'
    ]

    # 5. 转换日期格式从 "25 Jul 2025" 到 "2025-07-25"
    # 使用 errors='coerce' 会将任何无法转换的日期变为NaT(Not a Time)，避免程序中断
    clean_df['Month'] = pd.to_datetime(clean_df['Month'], format='%d %b %Y', errors='coerce').dt.strftime('%Y-%m-%d')

    # 删除任何因为日期转换失败而产生的空行
    clean_df.dropna(subset=['Month'], inplace=True)
    
    print("Data processing complete. Final DataFrame is ready:")
    print(clean_df.head())
    
    return clean_df

def append_to_gsheet(dataframe, gsheet_id, sheet_title):
    """
    将DataFrame附加到Google Sheet，此函数保持不变。
    """
    if dataframe is None or dataframe.empty:
        print("DataFrame is empty. Skipping Google Sheet update.")
        return
    try:
        print(f"Connecting to Google Sheet '{sheet_title}' to append data...")
        gc = pygsheets.authorize(service_file=SERVICE_ACCOUNT_FILE)
        sh = gc.open_by_key(gsheet_id)
        wks = sh.worksheet_by_title(sheet_title)
        
        all_values = wks.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False)
        last_data_row = len(all_values)
        
        num_new_rows = len(dataframe)
        if (last_data_row + num_new_rows) > wks.rows:
            rows_to_add = (last_data_row + num_new_rows) - wks.rows + 500
            print(f"Sheet is full. Adding {rows_to_add} more rows...")
            wks.add_rows(rows_to_add)
            print("Successfully added more rows.")

        next_empty_row = last_data_row + 1
        print(f"Appending new data starting at row {next_empty_row}...")
        wks.set_dataframe(dataframe, start=(next_empty_row, 1), copy_head=False, nan='')
        
        print(f"Successfully appended data to Google Sheet '{sheet_title}'.")

    except Exception as e:
        print(f"An error occurred during Google Sheet sync: {e}")
        raise

def main():
    """主执行函数"""
    print("Automation task started...")
    
    # 步骤 1: 抓取原始的、未处理的数据
    raw_dataframe = scrape_table_data('https://dashboard.fastmarkets.com/sw/x2TtMTTianBBefSdGCeZXc/palm-oil-global-prices')
    
    # 步骤 2: 清理和重组数据
    clean_dataframe = process_and_clean_data(raw_dataframe)
    
    # 步骤 3: 将干净的数据上传到 Google Sheet
    append_to_gsheet(
        dataframe=clean_dataframe,
        gsheet_id=GSHEET_ID,
        sheet_title=GSHEET_TITLE
    )
    print("Automation task completed successfully! ✅")

if __name__ == "__main__":
    main()
