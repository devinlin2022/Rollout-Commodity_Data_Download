You've hit a known compatibility issue\! The `driver.service.get_session()._ws` approach to access the WebDriver's underlying WebSocket (for CDP communication) is an internal implementation detail and can change across Selenium versions, or might not be exposed directly in all driver implementations (like `Service` objects).

This `AttributeError: 'Service' object has no attribute 'get_session'` indicates that the `Service` object (which manages the ChromeDriver executable process) doesn't directly provide a way to get the session needed for CDP. Instead, CDP commands are usually executed directly on the `driver` object itself.

My apologies for including that problematic line. It's a remnant of some older or more advanced CDP integration patterns.

Let's fix this by removing the problematic `get_session()` call and relying on the standard `driver.execute_cdp_cmd` method which is the correct and supported way to send CDP commands directly via the `WebDriver` instance.

The `wait_for_download_completion` function needs to correctly use `driver.execute_cdp_cmd('Browser.getDownloads', {})` without trying to access the internal session details.

Here's the corrected and simplified `wait_for_download_completion` function and `get_chrome_driver` function (removing the `driver.service.get_session()` part) in your `RISI_data_download.py` script:

-----

### **Corrected Python Code (`main/RISI_data_download.py`)**

The main change is removing the problematic `driver.service.get_session()._ws` line and ensuring `wait_for_download_completion` directly uses `driver.execute_cdp_cmd`.

```python
import time
import pandas as pd
import requests
import re
import json
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from retrying import retry

# --- Google Sheets Authentication ---
try:
    gc_pygsheets = pygsheets.authorize(service_file='service_account_key.json')
    gc_gspread = gspread.service_account(filename='service_account_key.json')
    print("Google Sheets authenticated using service account.")
except Exception as e:
    print(f"Error authenticating with Google Sheets service account: {e}")
    raise Exception("Google Sheets authentication failed. Check service account key.")

CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH", "/usr/local/bin/chromedriver")
DOWNLOAD_DIR = "/tmp/downloads"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# A dictionary to store download details for CDP monitoring (this is managed internally by wait_for_download_completion)
# DOWNLOAD_TRACKER = {} # This global is no longer strictly needed with the revised CDP polling

def get_chrome_driver():
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument('--log-level=3')

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safeBrowse.enabled": True
    }
    options.add_experimental_option('prefs', prefs)

    service = Service(executable_path=CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    # Enable downloads in headless mode via Chrome DevTools Protocol (CDP)
    # This is crucial for reliable headless downloads and is called directly on the driver
    driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': DOWNLOAD_DIR})

    # The problematic line 'client = driver.service.get_session()._ws' has been removed.
    # CDP commands like Browser.getDownloads are executed directly on the driver object.

    return driver

# Helper function for clicking elements with retries
@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException), stop_max_attempt_number=5, wait_fixed=2000)
def click_element_with_retry(driver, by_locator):
    print(f"Attempting to click element: {by_locator}")
    wait = WebDriverWait(driver, 15)
    element = wait.until(EC.element_to_be_clickable(by_locator))
    element.click()
    print(f"Clicked element: {by_locator}")

# Helper function for JavaScript clicks with retries
@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException), stop_max_attempt_number=5, wait_fixed=2000)
def js_click_element_with_retry(driver, css_selector):
    print(f"Attempting to JS click element: {css_selector}")
    wait = WebDriverWait(driver, 15)
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_selector)))
    driver.execute_script("arguments[0].click();", element)
    print(f"JS Clicked element: {css_selector}")


@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException) or isinstance(e, ValueError), stop_max_attempt_number=5, wait_fixed=2000)
def fetch_RISI_data(link):
    """
    Fetches data from RISI by logging in, simulating clicks to trigger CSV download.
    Returns the driver instance to monitor download completion.
    """
    print(f"Attempting to fetch RISI data from: {link}")
    driver = get_chrome_driver()
    driver.implicitly_wait(10)
    driver.get(link)

    wait = WebDriverWait(driver, 100)
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

    current_url = driver.current_url
    if "login" in current_url.lower() and "success" not in current_url.lower():
        print(f"Still on login page after initial login attempt, current URL: {current_url}")
        try:
            continue_button_selector = '#continue-login-button'
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, continue_button_selector)))
            js_click_element_with_retry(driver, continue_button_selector)
            time.sleep(5)
            print("Clicked continue login button.")
        except:
            pass
    
    print(f"Current URL after login attempts: {driver.current_url}")

    export_button_selector = '#cells-container > fui-grid-cell > fui-widget > header > fui-widget-actions > div:nth-child(1) > button > span.mat-mdc-button-touch-target'
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, export_button_selector)))
        js_click_element_with_retry(driver, export_button_selector)
        time.sleep(3)
        print("Export button clicked.")
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click RISI export button: {e}")

    csv_export_selector = "#mat-menu-panel-3 > div > div > div:nth-child(2) > fui-export-dropdown-item:nth-child(3) > button > span"
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, csv_export_selector)))
        js_click_element_with_retry(driver, csv_export_selector)
        print("CSV export option clicked. Now monitoring download status.")
        # Do NOT close driver here, it's returned for download monitoring
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click RISI CSV export option: {e}")

    # Return the driver instance to allow external monitoring of downloads
    return driver

def wait_for_download_completion(driver, timeout=180):
    """
    Monitors Chrome DevTools Protocol to wait for download completion.
    Returns the final path to the completed download file.
    """
    print(f"Monitoring download completion for {timeout} seconds...")
    start_time = time.time()
    download_id_to_track = None

    # First, list existing downloads to find a new 'inProgress' item
    # Poll for a new download starting
    while time.time() - start_time < timeout / 3: # Spend up to 1/3 of timeout just waiting for download to show up
        downloads_info = driver.execute_cdp_cmd('Browser.getDownloads', {})
        
        # Look for any 'inProgress' item that is new (not seen in a previous state, if we were tracking globally)
        # For simplicity, we assume one new download per call to fetch_RISI_data.
        for item in downloads_info['items']:
            if item['state'] == 'inProgress':
                download_id_to_track = item['id']
                print(f"New download detected, ID: {download_id_to_track}, Initial file path: {item['filePath']}")
                break
        if download_id_to_track:
            break
        time.sleep(1)
    
    if not download_id_to_track:
        # As a fallback, check if a download already completed very quickly
        downloads_info = driver.execute_cdp_cmd('Browser.getDownloads', {})
        for item in downloads_info['items']:
            if item['state'] == 'completed':
                print(f"Download completed very quickly (ID: {item['id']}). Path: {item['filePath']}")
                return item['filePath']
        raise Exception(f"No new download detected as 'inProgress' or 'completed' within {int(timeout/3)} seconds.")


    # Now, monitor the specific download item
    while time.time() - start_time < timeout:
        downloads_info = driver.execute_cdp_cmd('Browser.getDownloads', {})
        found_item = None
        for item in downloads_info['items']:
            if item['id'] == download_id_to_track:
                found_item = item
                break
        
        if found_item:
            print(f"Download ID: {found_item['id']}, Status: {found_item['state']}, Progress: {found_item['receivedBytes']}/{found_item['totalBytes']} Bytes, Current path: {found_item['filePath']}")
            if found_item['state'] == 'completed':
                final_path = found_item['filePath']
                print(f"Download COMPLETED. Final Path: {final_path}")
                # Optional: Sanity check if the file exists and is not empty before returning
                if os.path.exists(final_path) and os.path.getsize(final_path) > 0:
                    return final_path
                else:
                    raise Exception(f"Downloaded file '{final_path}' is empty or does not exist after completion reported.")
            elif found_item['state'] == 'interrupted':
                raise Exception(f"Download INTERRUPTED for ID {found_item['id']}: {found_item.get('dangerType', 'Unknown')}. Reason: {found_item.get('lastError', 'No error message')}")
        else:
            print(f"Download ID {download_id_to_track} not found in current downloads list (might have vanished or completed).")

        time.sleep(1) # Check status every second
    
    # If timeout reached and download not completed
    downloads_info = driver.execute_cdp_cmd('Browser.getDownloads', {})
    final_state_info = next((item for item in downloads_info['items'] if item['id'] == download_id_to_track), None)
    current_state = final_state_info.get('state', 'Unknown') if final_state_info else 'Not found'
    raise Exception(f"Download of '{download_id_to_track}' did not complete within {timeout} seconds. Current state: {current_state}")


def clean_palm_csv(input_path, output_path):
    """
    Cleans the downloaded Palm CSV file by skipping header rows,
    removing footer rows, and filling NaN values.
    """
    print(f"Cleaning CSV from {input_path}...")
    try:
        # Assuming the structure requires skipping 2 initial rows
        df = pd.read_csv(input_path, skiprows=2) 
        
        # Promote the current first row to header, then remove it from data
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)

        # Drop the next two rows as specified in your original logic
        df = df.drop(0).reset_index(drop=True)
        df = df.drop(0).reset_index(drop=True)
        
        # Remove the last 5 rows as specified
        df = df.iloc[:-5]

        # Fill NaN values with the last valid observation in each column
        for column in df.columns:
            last_valid_value = df[column].dropna().iloc[-1] if not df[column].dropna().empty else None
            if last_valid_value is not None:
                df[column] = df[column].fillna(last_valid_value)

        # Convert the first column (assumed to be date) to 'YYYY-MM-DD' format
        if not df.empty and df.columns[0]:
            try:
                df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], errors='coerce')
                df = df.dropna(subset=[df.columns[0]]) # Drop rows where date conversion failed
                df[df.columns[0]] = df[df.columns[0]].dt.strftime('%Y-%m-%d')
            except Exception as e:
                print(f"Warning: Could not format date column '{df.columns[0]}': {e}")
                pass # Continue even if date formatting fails

        df.to_csv(output_path, index=False)
        print(f"Cleaned CSV saved to: {output_path}")
    except Exception as e:
        print(f"Error cleaning CSV {input_path}: {e}")
        raise # Re-raise to fail the workflow if cleaning fails

def sync_and_dedup_csv_to_gsheet(csv_path, gsheet_id, sheet_title):
    """
    Reads data from a CSV file, syncs and deduplicates it with a Google Sheet.
    """
    print(f"Syncing {csv_path} to Google Sheet '{sheet_title}' (ID: {gsheet_id})...")
    df_new = pd.read_csv(csv_path)

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
                pass

        # Align columns before concatenation for robust merging
        if not df_old.empty and not df_new.empty:
            all_cols = list(pd.Index(df_old.columns).union(df_new.columns))
            df_new_aligned = df_new.reindex(columns=all_cols)
            df_old_aligned = df_old.reindex(columns=all_cols)
            df_all = pd.concat([df_old_aligned, df_new_aligned], ignore_index=True).dropna(how='all', subset=all_cols)
        else: # Either old_df or new_df is empty, just use concat on what's available
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

# Main workflow function that orchestrates the RISI Palm Oil data processing
def main_workflow():
    risi_palm_oil_link = 'https://dashboard.fastmarkets.com/sw/x2TtMTTianBBefSdGCeZXc/palm-oil-global-prices'

    print("Starting fetch_RISI_data (CSV download initiation)...")
    # fetch_RISI_data now returns the driver object, which is needed for download monitoring
    driver_instance = None # Initialize driver_instance outside try block
    try:
        driver_instance = fetch_RISI_data(risi_palm_oil_link)
        
        download_dir_path = DOWNLOAD_DIR
        new_filename = "Palm_original.csv"
        
        # Use the new CDP-based download monitoring
        downloaded_file_path = wait_for_download_completion(driver_instance, timeout=180) # Increased timeout
        print(f"Download completed: {downloaded_file_path}")
        
        # Rename the downloaded file to the desired new_filename
        final_file_path = os.path.join(download_dir_path, new_filename)
        if downloaded_file_path != final_file_path:
            try:
                os.rename(downloaded_file_path, final_file_path)
                print(f"Renamed downloaded file to: {final_file_path}")
                downloaded_file_path = final_file_path # Update path for cleaning
            except OSError as e:
                print(f"Could not rename downloaded file {downloaded_file_path} to {final_file_path}: {e}")
                # Proceed with original path if rename fails, cleaning will handle it
        
        driver_instance.quit() # Close the driver AFTER download is confirmed and file is ready
        print("Driver closed after successful download.")

    except Exception as e:
        print(f"Download process failed: {e}")
        if driver_instance:
            driver_instance.quit() # Ensure driver is closed on failure
            print("Driver closed due to download failure.")
        raise # Re-raise the exception to fail the workflow

    cleaned_filename = "Palm_cleaned.csv"
    cleaned_file_path = os.path.join(download_dir_path, cleaned_filename)
    clean_palm_csv(downloaded_file_path, cleaned_file_path)

    gsheet_id = '1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE'
    sheet_title = 'Palm Oil Price'

    sync_and_dedup_csv_to_gsheet(
        csv_path=cleaned_file_path,
        gsheet_id=gsheet_id,
        sheet_title=sheet_title
    )
    print("RISI Palm Oil workflow completed successfully.")

if __name__ == "__main__":
    @retry(stop_max_attempt_number=5, wait_fixed=5000)
    def run_main_workflow_with_retries():
        main_workflow()

    run_main_workflow_with_retries()
```
