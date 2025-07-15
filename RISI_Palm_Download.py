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

# --- Configuration ---
RISI_USERNAME = os.getenv('RISI_USERNAME')
RISI_PASSWORD = os.getenv('RISI_PASSWORD')
SERVICE_ACCOUNT_FILE = 'service_account_key.json'
GSHEET_ID = '1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE'
GSHEET_TITLE = 'Palm Oil Price'
CHROMEDRIVER_PATH = os.getenv('CHROMEDRIVER_PATH', '/usr/bin/chromedriver')
DOWNLOAD_DIR = "/tmp/downloads" # For error screenshots

def scrape_table_data(link):
    """
    Logs in and scrapes the AG-Grid table. It now uses the first row of
    data as the header, which is a robust method.
    """
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

        # 2. Scrape all rows from the AG-Grid table
        print("Waiting for the data grid to load...")
        time.sleep(10)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#cells-container > fui-grid-cell > fui-widget")))
        print("Find Table!")
        grid_selector = (By.CSS_SELECTOR, 'div[role="treegrid"]')
        grid_container = wait.until(EC.visibility_of_element_located(grid_selector))
        print("AG-Grid container found. Extracting all row and cell data...")

        # Find all row elements
        rows = grid_container.find_elements(By.CSS_SELECTOR, '[role="row"]')
        
        # Extract the text from all cells in all rows
        all_rows_raw = []
        for row in rows:
            cells = row.find_elements(By.CSS_SELECTOR, '[role="gridcell"]')
            row_data = [cell.text for cell in cells]
            # Add the row only if it contains some actual text
            if any(cell_text.strip() for cell_text in row_data):
                all_rows_raw.append(row_data)

        if not all_rows_raw or len(all_rows_raw) < 2:
            raise ValueError("Not enough rows found to create a header and data. Found {} rows.".format(len(all_rows_raw)))

        # --- FINAL LOGIC: Use the first row as the header ---
        print("Successfully extracted raw data. Using the first row as the header.")
        
        # The first item in our list is the header row
        headers = all_rows_raw[0]
        
        # The rest of the items are the data rows
        data_rows = all_rows_raw[1:]
        
        # Create the final DataFrame
        price_df = pd.DataFrame(data_rows, columns=headers)
        
        print("Successfully created DataFrame:")
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

def update_gsheet(dataframe, gsheet_id, sheet_title):
    # This function remains the same and is ready to use
    if dataframe is None or dataframe.empty:
        print("DataFrame is empty. Skipping Google Sheet update.")
        return
    try:
        print(f"Connecting to Google Sheet '{sheet_title}'...")
        gc = pygsheets.authorize(service_file=SERVICE_ACCOUNT_FILE)
        sh = gc.open_by_key(gsheet_id)
        wks = sh.worksheet_by_title(sheet_title)
        print("Clearing the worksheet...")
        wks.clear()
        print("Writing new data to the worksheet...")
        wks.set_dataframe(dataframe, (1, 1), nan='')
        print(f"Successfully overwrote data in Google Sheet '{sheet_title}'.")
    except Exception as e:
        print(f"An error occurred during Google Sheet sync: {e}")
        raise

def main():
    # This function remains the same and is ready to use
    print("Automation task started...")
    price_dataframe = scrape_table_data('https://dashboard.fastmarkets.com/sw/x2TtMTTianBBefSdGCeZXc/palm-oil-global-prices')
    update_gsheet(
        dataframe=price_dataframe,
        gsheet_id=GSHEET_ID,
        sheet_title=GSHEET_TITLE
    )
    print("Automation task completed successfully! ✅")

if __name__ == "__main__":
    main()
