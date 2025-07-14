import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pygsheets
from gspread_dataframe import set_with_dataframe
import gspread
import os
from selenium.webdriver.chrome.service import Service
from retrying import retry

# --- Google Sheets Authentication ---
try:
    gc_pygsheets = pygsheets.authorize(service_file='service_account_key.json')
    gc_gspread = gspread.service_account(filename='service_account_key.json')
except Exception as e:
    raise Exception(f"Google Sheets authentication failed: {e}")

CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH", "/usr/local/bin/chromedriver")

def get_chrome_driver():
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    
    # Removed download preferences and CDP download commands as we are no longer downloading files
    
    # === FIX START ===
    service = Service(executable_path=CHROMEDRIVER_PATH) 
    # === FIX END ===
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# Helper function for clicking elements with retries
@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException), stop_max_attempt_number=5, wait_fixed=2000)
def click_element_with_retry(driver, by_locator):
    print(f"Attempting to click element: {by_locator}")
    wait = WebDriverWait(driver, 15) # Increased wait time for clickability
    element = wait.until(EC.element_to_be_clickable(by_locator))
    element.click()
    print(f"Clicked element: {by_locator}")

# Helper function for JavaScript clicks with retries
@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException), stop_max_attempt_number=5, wait_fixed=2000)
def js_click_element_with_retry(driver, css_selector):
    print(f"Attempting to JS click element: {css_selector}")
    wait = WebDriverWait(driver, 15) # Increased wait time for presence
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_selector)))
    driver.execute_script("arguments[0].click();", element)
    print(f"JS Clicked element: {css_selector}")

def fetch_RISI_data(link):
    """
    Fetches data from RISI by logging in, navigating, and directly extracting
    table data into a pandas DataFrame.
    """
    print(f"Attempting to fetch RISI data from: {link}")
    driver = get_chrome_driver()
    driver.implicitly_wait(10)
    driver.get(link)

    wait = WebDriverWait(driver, 100) # Increased initial wait
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#userEmail')))
    print("Login page elements visible.")

    risi_username = os.getenv("RISI_USERNAME")
    risi_password = os.getenv("RISI_PASSWORD")

    if not risi_username or not risi_password:
        driver.quit()
        raise ValueError("RISI_USERNAME or RISI_PASSWORD environment variables not set.")

    driver.execute_script(f'document.querySelector("#userEmail").value = "{risi_username}"')
    driver.execute_script(f'document.querySelector("#password").value = "{risi_password}"')
    print("Username and password entered.")
    
    try:
        click_element_with_retry(driver, (By.CSS_SELECTOR, '#login-button'))
        print("Login button clicked.")
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click RISI login button: {e}")
        
    time.sleep(7) # Wait for page load after login

    # Check for potential 2FA or redirection
    current_url = driver.current_url
    if "login" in current_url.lower() and "success" not in current_url.lower():
        print(f"Still on login page after attempt, current URL: {current_url}")
        try:
            continue_button_selector = '#continue-login-button'
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, continue_button_selector)))
            js_click_element_with_retry(driver, continue_button_selector)
            time.sleep(5) # Wait for 2FA/continue
            print("Clicked continue login button.")
        except:
            pass # No continue button found, proceed
    
    print(f"Current URL after login attempts: {driver.current_url}")

    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, table_selector)))
        print(f"Table element found with selector: {table_selector}. Attempting to extract data.")
    except Exception as e:
        driver.quit()
        raise Exception(f"Table element not found after login/navigation: {e}. Please inspect the page and update 'table_selector'.")

    # Extract table headers
    headers = []
    header_elements_selector = f'{table_selector} thead th'
    try:
        header_elements = driver.find_elements(By.CSS_SELECTOR, header_elements_selector)
        headers = [header.text.strip() for header in header_elements if header.text.strip()]
        
        if not headers: # Fallback if headers are not in thead th
            print(f"Warning: No headers found using selector {header_elements_selector}. Trying first row of tbody.")
            # COMMON FALLBACK - Adjust if your table headers are in the first <tr> of the <tbody>
            first_row_cells = driver.find_elements(By.CSS_SELECTOR, f'{table_selector} tbody tr:first-child td')
            headers = [cell.text.strip() for cell in first_row_cells if cell.text.strip()]
            if headers:
                print("Headers found in tbody first row.")

        if not headers:
            driver.quit()
            raise Exception("Could not extract table headers. Selector might be wrong or table structure is unusual.")
    except Exception as e:
        driver.quit()
        raise Exception(f"Error extracting table headers: {e}")

    # Extract table rows
    data_rows = []
    # COMMON PLACEHOLDER - Adjust this if your table rows have a different selector
    row_elements_selector = f'{table_selector} tbody tr' 
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, row_elements_selector)
        if not rows:
            driver.quit()
            raise Exception(f"No table rows found using selector {row_elements_selector}. Table might be empty or selector is wrong.")

        # If headers were extracted from the first tbody row, start data extraction from the second row
        start_row_index = 1 if (f'{table_selector} tbody tr:first-child td' == header_elements_selector and headers) else 0

        for i, row_element in enumerate(rows):
            if i < start_row_index:
                continue # Skip the header row if it was taken from tbody

            cells = row_element.find_elements(By.TAG_NAME, 'td')
            row_data = [cell.text.strip() for cell in cells]
            if row_data: # Only add non-empty rows
                data_rows.append(row_data)

        if not data_rows:
            driver.quit()
            raise Exception("Extracted rows are empty after processing. Table might have no data or incorrect row/cell selectors.")
    except Exception as e:
        driver.quit()
        raise Exception(f"Error extracting table rows: {e}")

    driver.quit()
    
    # Create DataFrame
    # Handle cases where collected headers and row data don't perfectly match in count
    if headers and data_rows and len(headers) == len(data_rows[0]):
        df = pd.DataFrame(data_rows, columns=headers)
    elif headers and data_rows and len(headers) < len(data_rows[0]):
        print("Warning: Header count mismatch (fewer headers than data columns). Using integer index for extra columns.")
        df = pd.DataFrame(data_rows, columns=headers + list(range(len(headers), len(data_rows[0]))))
    elif headers and data_rows:
        print("Warning: Header count mismatch (more headers than data columns). Truncating headers.")
        df = pd.DataFrame(data_rows, columns=headers[:len(data_rows[0])])
    else:
        # Fallback if no valid headers or data, return empty DataFrame or raise
        print("Warning: No valid headers or data rows found to create DataFrame.")
        df = pd.DataFrame() # Return empty DataFrame or raise error

    # Post-processing the extracted DataFrame (similar to clean_palm_csv but now integrated)
    # This assumes the first column is the date and needs formatting/filling
    if not df.empty:
        # Fill NaN values with the last valid observation in each column
        for column in df.columns:
            last_valid_value = df[column].dropna().iloc[-1] if not df[column].dropna().empty else None
            if last_valid_value is not None:
                df[column] = df[column].fillna(last_valid_value)

        # Convert the first column to datetime and then 'YYYY-MM-DD' format
        if df.columns[0]:
            try:
                df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], errors='coerce')
                df = df.dropna(subset=[df.columns[0]]) # Drop rows where date conversion failed
                df[df.columns[0]] = df[df.columns[0]].dt.strftime('%Y-%m-%d')
            except Exception as e:
                print(f"Warning: Could not format date column '{df.columns[0]}': {e}")

    print(f"Successfully extracted DataFrame: {len(df)} rows, {len(df.columns) if not df.empty else 0} columns.")
    return df

# Renamed to clarify it works with a DataFrame directly
def sync_and_dedup_df_to_gsheet(df_new, gsheet_id, sheet_title):
    print(f"Syncing DataFrame to Google Sheet '{sheet_title}' (ID: {gsheet_id})...")
    
    sh = gc_pygsheets.open_by_key(gsheet_id)
    wks = sh.worksheet_by_title(sheet_title)

    try:
        df_old = wks.get_as_df(has_header=True, include_tailing_empty=False)
        if not df_old.empty and df_old.columns[0]:
            try:
                df_old[df_old.columns[0]] = pd.to_datetime(df_old[df_old.columns[0]], errors='coerce')
                df_old = df_old.dropna(subset=[df_old.columns[0]])
                df_old[df_old.columns[0]] = df_old[df_old.columns[0]].dt.strftime('%Y-%m-%d')
            except Exception as e:
                print(f"Warning: Could not format existing sheet's date column '{df_old.columns[0]}': {e}")

        # Align columns before concatenation
        if not df_old.empty and not df_new.empty:
            all_cols = list(pd.Index(df_old.columns).union(df_new.columns))
            df_new_aligned = df_new.reindex(columns=all_cols)
            df_old_aligned = df_old.reindex(columns=all_cols)
            df_all = pd.concat([df_old_aligned, df_new_aligned], ignore_index=True).dropna(how='all', subset=all_cols)
        else:
            df_all = pd.concat([df_old, df_new], ignore_index=True).dropna(how='all')

    except Exception as e:
        print(f"Could not retrieve existing data or concatenate: {e}. Proceeding with new data only.")
        df_all = df_new
        if not df_all.empty and df_all.columns[0]:
            try:
                df_all[df_all.columns[0]] = pd.to_datetime(df_all[df_all.columns[0]], errors='coerce')
                df_all = df_all.dropna(subset=[df_all.columns[0]])
                df_all[df_all.columns[0]] = df_all[df_all.columns[0]].dt.strftime('%Y-%m-%d')
            except Exception as e_new:
                print(f"Warning: Could not format new data's date column '{df_all.columns[0]}' after initial error: {e_new}")


    # Drop duplicates based on the first column (assumed to be date/key)
    if not df_all.empty:
        df_all = df_all.drop_duplicates(subset=[df_all.columns[0]]).sort_values(by=df_all.columns[0])
    
    # Clear and update Google Sheet
    wks.clear()
    wks.set_dataframe(df_all, (1,1), copy_head=True) # copy_head=True ensures headers are written

    print(f"Data synced and deduped to Google Sheet '{sheet_title}' successfully. Total rows: {len(df_all)}")

def main_workflow():
    risi_palm_oil_link = 'https://dashboard.fastmarkets.com/sw/x2TtMTTianBBefSdGCeZXc/palm-oil-global-prices'

    print("Starting fetch_RISI_data (direct extraction)...")
    extracted_df = fetch_RISI_data(risi_palm_oil_link)
    print("fetch_RISI_data completed.")

    if extracted_df.empty:
        print("No data extracted. Skipping Google Sheet update.")
        return # Exit if no data was extracted

    gsheet_id = '1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE'
    sheet_title = 'Palm Oil Price'

    sync_and_dedup_df_to_gsheet(
        df_new=extracted_df,
        gsheet_id=gsheet_id,
        sheet_title=sheet_title
    )
    print("RISI Palm Oil workflow completed successfully.")


if __name__ == "__main__":
    main_workflow()
