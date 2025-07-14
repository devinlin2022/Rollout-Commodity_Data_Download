import time
import pandas as pd
import requests # Used for ImageService in other contexts, but keeping it for consistency if it's shared.
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from PIL import Image # Not used in this specific script's flow
# import pymupdf # Not used in this specific script's flow
import pygsheets
from gspread_dataframe import set_with_dataframe
import gspread
import re # Not strictly used in this specific script's flow, but often handy
import json # Not strictly used in this specific script's flow, but often handy
import os
from selenium.webdriver.chrome.service import Service
from retrying import retry # Import the retrying library

# --- Google Sheets Authentication (using service account from file) ---
try:
    # 'service_account_key.json' is created by GitHub Actions from a secret
    gc_pygsheets = pygsheets.authorize(service_file='service_account_key.json')
    gc_gspread = gspread.service_account(filename='service_account_key.json')
    print("Google Sheets authenticated using service account.")
except Exception as e:
    print(f"Error authenticating with Google Sheets service account: {e}")
    raise Exception("Google Sheets authentication failed. Check service account key.")

# --- Define ChromeDriver Path and Download Directory ---
CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH", "/usr/local/bin/chromedriver")
DOWNLOAD_DIR = "/tmp/downloads" # Standard writable temp directory for GitHub Actions
os.makedirs(DOWNLOAD_DIR, exist_ok=True) # Ensure the directory exists

def get_chrome_driver():
    """Initializes and returns a Selenium Chrome WebDriver configured for headless download."""
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu') # Essential for headless on Linux
    options.add_experimental_option("excludeSwitches", ["enable-logging"]) # Suppress verbose logging
    options.add_argument('--log-level=3') # Suppress informational messages

    # Configure download preferences
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
    driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': DOWNLOAD_DIR})

    return driver

# Helper function for clicking elements with retries
@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException), stop_max_attempt_number=5, wait_fixed=2000)
def click_element_with_retry(driver, by_locator):
    print(f"Attempting to click element: {by_locator}")
    wait = WebDriverWait(driver, 15) # Increased wait time for clickability
    element = wait.until(EC.element_to_be_clickable(by_locator))
    element.click()
    print(f"Clicked element: {by_locator}")

# Helper function for JavaScript clicks with retries (useful if native click fails)
@retry(retry_on_exception=lambda e: isinstance(e, EC.WebDriverException), stop_max_attempt_number=5, wait_fixed=2000)
def js_click_element_with_retry(driver, css_selector):
    print(f"Attempting to JS click element: {css_selector}")
    wait = WebDriverWait(driver, 15) # Increased wait time for presence
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_selector)))
    driver.execute_script("arguments[0].click();", element)
    print(f"JS Clicked element: {css_selector}")

# Removed save_pdf and pdf_to_img functions as they are not used in this script's flow.

def fetch_RISI_data(link):
    """
    Fetches data from a RISI link by logging in, simulating clicks to download a CSV,
    and then closing the driver.
    """
    print(f"Attempting to fetch RISI data from: {link}")
    driver = get_chrome_driver()
    driver.implicitly_wait(10)
    driver.get(link)

    wait = WebDriverWait(driver, 100)
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#userEmail')))
    print("Login page elements visible.")

    # Retrieve RISI credentials from environment variables (GitHub Secrets)
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
        
    time.sleep(7) # Increased wait for page load after login

    # Check for potential 2FA or redirection (as seen in previous tracebacks)
    current_url = driver.current_url
    if "login" in current_url.lower() and "success" not in current_url.lower():
        print(f"Still on login page after initial login attempt, current URL: {current_url}")
        try:
            # Attempt to click a 'continue login' button if it appears
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

    # Click the "Export CSV" option
    csv_export_selector = "#mat-menu-panel-3 > div > div > div:nth-child(2) > fui-export-dropdown-item:nth-child(3) > button > span"
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, csv_export_selector)))
        js_click_element_with_retry(driver, csv_export_selector) # Use JS click for robustness
        time.sleep(10) # Increased wait for the download to initiate and start writing
        print("CSV export option clicked. Waiting for download to start.")
    except Exception as e:
        driver.quit()
        raise Exception(f"Failed to click RISI CSV export option: {e}")

    driver.quit()
    print("Driver closed.")
    return True # Indicates that the download process was initiated

def wait_file_and_rename(download_dir, old_ext, new_filename, timeout=120):
    """
    Waits for a file with a specific extension to appear in the download directory,
    then renames it. Includes retry logic for renaming if file is still busy.
    """
    print(f"Waiting for file with extension '{old_ext}' in '{download_dir}' for {timeout} seconds...")
    time_counter = 0
    while time_counter < timeout:
        files = [f for f in os.listdir(download_dir) if f.endswith(old_ext) and not f.endswith(".crdownload")]
        if files:
            # Sort by modification time to get the newest file (useful if multiple downloads occur)
            files.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)
            old_file = os.path.join(download_dir, files[0])
            new_file = os.path.join(download_dir, new_filename)
            try:
                os.rename(old_file, new_file)
                print(f"File '{old_file}' renamed to '{new_file}'.")
                return new_file
            except OSError as e:
                # This can happen if the file is still being written to by Chrome, retry
                print(f"Error renaming file {old_file}: {e}. Retrying in 1 second.")
                time.sleep(1)
        time.sleep(1)
        time_counter += 1
        if time_counter % 10 == 0:
            print(f"Still waiting... {time_counter} seconds elapsed. Files in dir: {os.listdir(download_dir)}")
    raise Exception(f"Download file with extension '{old_ext}' not found in '{download_dir}' within {timeout} seconds!")

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
    fetch_RISI_data(risi_palm_oil_link)
    print("fetch_RISI_data completed (download initiated).")

    download_dir_path = DOWNLOAD_DIR
    new_filename = "Palm_original.csv"
    
    # This is the line that failed previously. Increased timeout to 120s and added robustness.
    downloaded_file_path = wait_file_and_rename(download_dir_path, ".csv", new_filename)
    print(f"Download and renamed as: {downloaded_file_path}")

    cleaned_filename = "Palm_cleaned.csv"
    cleaned_file_path = os.path.join(download_dir_path, cleaned_filename)
    clean_palm_csv(downloaded_file_path, cleaned_file_path)

    gsheet_id = '1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE'
    sheet_title = 'Palm Oil Price'

    sync_and_dedup_csv_to_gsheet(
        csv_path=cleaned_file_path,
        gsheet_id=gsheet_id,
        sheet_title=sheet_title
        # service_account_json parameter is no longer needed as authentication is global
    )
    print("RISI Palm Oil workflow completed successfully.")

if __name__ == "__main__":
    def run_main_workflow_with_retries():
        main_workflow()

    run_main_workflow_with_retries()
