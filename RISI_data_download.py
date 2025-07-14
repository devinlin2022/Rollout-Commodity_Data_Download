import time
import pandas as pd
import requests # Keeping this as it's general purpose
import re # Keeping this as it's general purpose
import json # Keeping this as it's general purpose
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from retrying import retry
import pygsheets

# --- Google Sheets Authentication ---
try:
    gc_pygsheets = pygsheets.authorize(service_file='service_account_key.json')
    gc_gspread = gspread.service_account(filename='service_account_key.json')
    print("Google Sheets authenticated using service account.")
except Exception as e:
    print(f"Error authenticating with Google Sheets service account: {e}")
    raise Exception("Google Sheets authentication failed. Check service account key.")

CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH", "/usr/local/bin/chromedriver")
DOWNLOAD_DIR = "/tmp/downloads" # Standard writable temp directory for GitHub Actions
os.makedirs(DOWNLOAD_DIR, exist_ok=True) # Create the directory if it doesn't exist

# A dictionary to store download details for CDP monitoring
# It will map download ID to its state (e.g., "in_progress", "completed", "failed") and filename
DOWNLOAD_TRACKER = {}

def get_chrome_driver():
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_experimental_option("excludeSwitches", ["enable-logging"]) # Suppress verbose logging
    options.add_argument('--log-level=3') # Suppress informational messages

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safeBrowse.enabled": True # Corrected preference name
    }
    options.add_experimental_option('prefs', prefs)

    service = Service(executable_path=CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    # Enable downloads in headless mode via Chrome DevTools Protocol (CDP)
    # This is crucial for reliable headless downloads
    driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': DOWNLOAD_DIR})

    # Listen to download progress events for robust monitoring
    # This part sets up a listener for download events.
    # We will reset the DOWNLOAD_TRACKER before each download attempt.
    client = driver.service.get_session()._ws # Access the websocket client
    # The listener needs to be outside get_chrome_driver if you want to track across driver instances
    # But for a single download per fetch_RISI_data, it's fine here.
    
    # We don't have direct access to add_listener in headless chrome, so we poll the status
    # Instead of an event listener, we will poll the download status from CDP after initiating download.

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

    # Check for potential 2FA or redirection
    current_url = driver.current_url
    if "login" in current_url.lower() and "success" not in current_url.lower():
        print(f"Still on login page after initial login attempt, current URL: {current_url}")
        try:
            continue_button_selector = '#continue-login-button'
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, continue_button_selector)))
            js_click_element_with_retry(driver, continue_button_selector)
            time.sleep(5) # Wait for 2FA/continue process
            print("Clicked continue login button.")
        except:
            pass # No continue button found or it wasn't clickable, proceed anyway
    
    print(f"Current URL after login attempts: {driver.current_url}")

    # Click the export button
    export_button_selector = '#cells-container > fui-grid-cell > fui-widget > header > fui-widget-actions > div:nth-child(1) > button > span.mat-mdc-button-touch-target'
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, export_button_selector)))
        js_click_element_with_retry(driver, export_button_selector) # Use JS click for robustness
        time.sleep(3) # Wait for the export menu to appear
        print("Export button clicked.")
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click RISI export button: {e}")

    # Reset download tracker for this specific download
    global DOWNLOAD_TRACKER
    DOWNLOAD_TRACKER = {} # Clear previous download state

    # Click the "Export CSV" option
    csv_export_selector = "#mat-menu-panel-3 > div > div > div:nth-child(2) > fui-export-dropdown-item:nth-child(3) > button > span"
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, csv_export_selector)))
        js_click_element_with_retry(driver, csv_export_selector) # Use JS click for robustness
        # No time.sleep here, as wait_for_download_completion will handle timing
        print("CSV export option clicked. Now monitoring download status.")
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click RISI CSV export option: {e}")

    # Return the driver for the download monitoring part to pick up
    return driver

def wait_for_download_completion(driver, timeout=120):
    """
    Monitors Chrome DevTools Protocol to wait for download completion.
    Returns the path to the completed download file.
    """
    start_time = time.time()
    last_known_file = None
    
    # Enable Fetch domain to listen for network requests (optional, but can provide more info)
    # driver.execute_cdp_cmd('Fetch.enable', {}) 

    # List files in download directory to find initial .crdownload file
    print(f"Monitoring download directory: {DOWNLOAD_DIR}")
    
    # We will poll the download status via CDP
    # Get the current list of downloads
    downloads_info = driver.execute_cdp_cmd('Browser.getDownloads', {})
    initial_downloads_count = len(downloads_info['items'])
    
    # Wait for a new download item to appear
    current_downloads_count = initial_downloads_count
    while current_downloads_count == initial_downloads_count and (time.time() - start_time < timeout / 2):
        time.sleep(1)
        downloads_info = driver.execute_cdp_cmd('Browser.getDownloads', {})
        current_downloads_count = len(downloads_info['items'])
        print(f"Waiting for new download: {current_downloads_count} items found (initial: {initial_downloads_count})")

    if current_downloads_count == initial_downloads_count:
        raise Exception(f"Download did not start in {timeout / 2} seconds.")

    # Now, try to find the *new* download item and monitor it
    new_download_item = None
    for item in downloads_info['items']:
        if item['state'] == 'inProgress' and not item['id'] in DOWNLOAD_TRACKER: # Check if this is a newly found in-progress download
             new_download_item = item
             DOWNLOAD_TRACKER[item['id']] = 'inProgress' # Mark as tracked
             break

    if not new_download_item:
        # If no new inProgress item found, check for a 'completed' one that might have finished quickly
        for item in downloads_info['items']:
            if item['state'] == 'completed' and not item['id'] in DOWNLOAD_TRACKER:
                print(f"Download completed very quickly: {item['filePath']}")
                DOWNLOAD_TRACKER[item['id']] = 'completed'
                return item['filePath']
        raise Exception("New download item not identified or already completed before tracking.")


    print(f"Found new download item with ID: {new_download_item['id']}")

    while (time.time() - start_time < timeout):
        downloads_info = driver.execute_cdp_cmd('Browser.getDownloads', {})
        
        found_item = None
        for item in downloads_info['items']:
            if item['id'] == new_download_item['id']: # Find the specific download we're tracking
                found_item = item
                break
        
        if found_item:
            print(f"Download status: {found_item['state']}, Progress: {found_item['receivedBytes']}/{found_item['totalBytes']} Bytes, File: {found_item['filePath']}")
            if found_item['state'] == 'completed':
                print(f"Download with ID {found_item['id']} COMPLETED. Path: {found_item['filePath']}")
                DOWNLOAD_TRACKER[found_item['id']] = 'completed'
                return found_item['filePath']
            elif found_item['state'] == 'interrupted':
                DOWNLOAD_TRACKER[found_item['id']] = 'failed'
                raise Exception(f"Download with ID {found_item['id']} was INTERRUPTED: {found_item.get('dangerType', 'Unknown')}")
            
            # Update last_known_file if it's progressing
            last_known_file = found_item['filePath']

        time.sleep(1) # Check status every second
    
    # If timeout reached
    if last_known_file:
        raise Exception(f"Download of '{last_known_file}' did not complete within {timeout} seconds. Current state: {found_item.get('state', 'Unknown')}")
    else:
        raise Exception(f"No active download found or completed within {timeout} seconds.")

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

    # Use the globally authenticated pygsheets client
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
# This function is the primary entry point for this specific script
def main_workflow():
    risi_palm_oil_link = 'https://dashboard.fastmarkets.com/sw/x2TtMTTianBBefSdGCeZXc/palm-oil-global-prices'

    print("Starting fetch_RISI_data (CSV download initiation)...")
    # fetch_RISI_data now returns the driver object, which is needed for download monitoring
    driver_after_clicks = fetch_RISI_data(risi_palm_oil_link)
    
    download_dir_path = DOWNLOAD_DIR
    new_filename = "Palm_original.csv"
    
    # Use the new CDP-based download monitoring
    try:
        downloaded_file_path = wait_for_download_completion(driver_after_clicks, timeout=180) # Increased timeout
        print(f"Download completed: {downloaded_file_path}")
        
        # Rename the downloaded file to the desired new_filename
        final_file_path = os.path.join(download_dir_path, new_filename)
        # Ensure the downloaded file is moved/renamed to the clean name
        # If the downloaded_file_path is already the clean name (Chrome often names it directly),
        # this rename will just confirm. If it's a temp name, it will rename it.
        if downloaded_file_path != final_file_path:
            try:
                os.rename(downloaded_file_path, final_file_path)
                print(f"Renamed downloaded file to: {final_file_path}")
                downloaded_file_path = final_file_path # Update path for cleaning
            except OSError as e:
                print(f"Could not rename downloaded file {downloaded_file_path} to {final_file_path}: {e}")
                # Proceed with original path if rename fails, cleaning will handle it
                downloaded_file_path = downloaded_file_path # Use the path returned by CDP

    except Exception as e:
        print(f"Download process failed: {e}")
        driver_after_clicks.quit() # Ensure driver is closed on failure
        raise # Re-raise the exception to fail the workflow

    driver_after_clicks.quit() # Close the driver after download is confirmed and file is ready
    print("Driver closed after download.")

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
    # Apply retry to the main workflow execution
    @retry(stop_max_attempt_number=5, wait_fixed=5000) # Retry 5 times, 5 seconds between retries
    def run_main_workflow_with_retries():
        main_workflow()

    run_main_workflow_with_retries()
