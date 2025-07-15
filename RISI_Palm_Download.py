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
    Logs in and scrapes the AG-Grid table directly from the page HTML.
    This method is specifically tailored to the provided HTML structure.
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

        # 2. Precisely Scrape the AG-Grid Table
        print("Waiting for the data grid to load...")
        time.sleep(10)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#cells-container > fui-grid-cell > fui-widget")))
        print("找到表格")
        # Locate the main grid container using its role. This is the key.
        grid_selector = (By.CSS_SELECTOR, 'div[role="treegrid"]')
        grid_container = wait.until(EC.visibility_of_element_located(grid_selector))
        print("AG-Grid container found. Parsing headers and rows...")

        # Extract headers from elements with role="columnheader"
        # We also filter out any empty headers (like control columns)
        header_elements = grid_container.find_elements(By.CSS_SELECTOR, '[role="columnheader"] .ag-header-cell-text')
        headers = [h.text for h in header_elements if h.text.strip()]
        print(f"Found Headers: {headers}")

        # Extract data from each row
        all_rows_data = []
        # Find all row elements. AG-Grid marks them with role="row".
        rows = grid_container.find_elements(By.CSS_SELECTOR, '[role="row"]')
        print(f"Found {len(rows)} data rows. Extracting cell data...")

        for row in rows:
            # Find all cells within that specific row
            cells = row.find_elements(By.CSS_SELECTOR, '[role="gridcell"]')
            # Extract text from each cell
            row_data = [cell.text for cell in cells]

            # The first few cells might be empty control columns, let's align data with headers
            # We assume the meaningful data starts where the description is.
            # A simple heuristic: find the first non-empty cell and start from there.
            if row_data:
                # This part may need slight adjustment based on the exact number of control columns
                # Let's find the first cell that seems to have real data
                first_data_index = -1
                for i, cell_text in enumerate(row_data):
                    if cell_text.strip(): # Find first non-empty cell
                        first_data_index = i
                        break
                
                # Slice the row data to align with headers
                if first_data_index != -1:
                    # The number of "real" data columns should match the number of headers
                    # This slicing assumes control columns are at the start.
                    aligned_row_data = row_data[first_data_index : first_data_index + len(headers)]
                    if len(aligned_row_data) == len(headers):
                         all_rows_data.append(aligned_row_data)

        if not all_rows_data:
             raise ValueError("Could not extract any valid data rows from the grid.")

        price_df = pd.DataFrame(all_rows_data, columns=headers)
        
        print("Successfully parsed table data.")
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
    # This function remains the same
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
    # Main function remains the same
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
