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
GSHEET_TITLE = 'Palm Oil Price'
CHROMEDRIVER_PATH = os.getenv('CHROMEDRIVER_PATH', '/usr/bin/chromedriver')
DOWNLOAD_DIR = "/tmp/downloads" # For error screenshots

def scrape_table_data(link):
    # This function is correct.
    options = Options()
    options.binary_location = '/usr/bin/chromium-browser'
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument("window-size=1920,1080")

    service = webdriver.chrome.service.Service(executable_path=CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(link)
        wait = WebDriverWait(driver, 60)

        # 1. Login
        print("Waiting for login page...")
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#userEmail'))).send_keys(RISI_USERNAME)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#password'))).send_keys(RISI_PASSWORD)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#login-button'))).click()
        print("Login successful.")

        # 2. Scrape the AG-Grid table
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
        print(f"Using {num_expected_columns} fixed headers.")

        rows = grid_container.find_elements(By.CSS_SELECTOR, '[role="row"]')
        all_rows_raw = []
        for row in rows:
            cells = row.find_elements(By.CSS_SELECTOR, '[role="gridcell"]')
            row_data = [cell.text for cell in cells]
            all_rows_raw.append(row_data)

        data_rows = [row for row in all_rows_raw if len(row) == num_expected_columns]

        if not data_rows:
            raise ValueError(f"Scraping failed: No rows found with the expected {num_expected_columns} columns.")
        
        price_df = pd.DataFrame(data_rows, columns=fixed_headers)
        
        print("Successfully created final DataFrame with fixed headers:")
        print(price_df.head())
        
        return price_df

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

def append_to_gsheet(dataframe, gsheet_id, sheet_title):
    """
    Appends a DataFrame to a Google Sheet by finding the next empty row
    and using set_dataframe for a direct paste. This is the most reliable method.
    """
    if dataframe is None or dataframe.empty:
        print("DataFrame is empty. Skipping Google Sheet update.")
        return
    try:
        print(f"Connecting to Google Sheet '{sheet_title}' to append data...")
        gc = pygsheets.authorize(service_file=SERVICE_ACCOUNT_FILE)
        sh = gc.open_by_key(gsheet_id)
        wks = sh.worksheet_by_title(sheet_title)

        last_row = len(wks.get_all_records(head=1, empty_values=False))
        next_empty_row = last_row + 2 # +1 for header, +1 for next row
        
        print(f"Found {last_row} existing records. Appending new data at row {next_empty_row}...")
        
        # 2. Use set_dataframe to paste data at the specific empty row.
        # copy_head=False ensures we don't write the header row again.
        wks.set_dataframe(dataframe, start=(next_empty_row, 1), copy_head=False, nan='')
        
        print(f"Successfully appended data to Google Sheet '{sheet_title}'.")

    except Exception as e:
        print(f"An error occurred during Google Sheet sync: {e}")
        raise

def main():
    """Main execution function."""
    print("Automation task started...")
    price_dataframe = scrape_table_data('https://dashboard.fastmarkets.com/sw/x2TtMTTianBBefSdGCeZXc/palm-oil-global-prices')
    append_to_gsheet(
        dataframe=price_dataframe,
        gsheet_id=GSHEET_ID,
        sheet_title=GSHEET_TITLE
    )
    print("Automation task completed successfully! ✅")

if __name__ == "__main__":
    main()
