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
        "safeBrowse.enabled": True
    }
    options.add_experimental_option('prefs', prefs)

    service = Service(executable_path=CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def fetch_RISI_data(link):
    driver = get_chrome_driver()
    driver.implicitly_wait(10)
    driver.get(link)

    wait = WebDriverWait(driver, 100)
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#userEmail')))

    risi_username = os.getenv("RISI_USERNAME")
    risi_password = os.getenv("RISI_PASSWORD")

    if not risi_username or not risi_password:
        raise ValueError("RISI_USERNAME or RISI_PASSWORD environment variables not set.")

    driver.execute_script(f'document.querySelector("#userEmail").value = "{risi_username}"')
    driver.execute_script(f'document.querySelector("#password").value = "{risi_password}"')
    driver.execute_script(f'document.querySelector("#login-button").click()')
    time.sleep(5)

    try:
        export_button = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#cells-container > fui-grid-cell > fui-widget > header > fui-widget-actions > div:nth-child(1) > button > span.mat-mdc-button-touch-target')))
        export_button.click()
        time.sleep(2)
    except Exception as e:
        driver.quit()
        raise

    try:
        csv_export_option = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#mat-menu-panel-3 > div > div > div:nth-child(2) > fui-export-dropdown-item:nth-child(3) > button > span"))
        )
        csv_export_option.click()
        time.sleep(5)
    except Exception as e:
        driver.quit()
        raise

    driver.quit()
    return True

def wait_file_and_rename(download_dir, old_ext, new_filename, timeout=60):
    time_counter = 0
    while time_counter < timeout:
        files = [f for f in os.listdir(download_dir) if f.endswith(old_ext) and not f.endswith(".crdownload")]
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)
            old_file = os.path.join(download_dir, files[0])
            new_file = os.path.join(download_dir, new_filename)
            os.rename(old_file, new_file)
            return new_file
        time.sleep(1)
        time_counter += 1
    raise Exception(f"Download file with extension '{old_ext}' not found in '{download_dir}' within {timeout} seconds!")

def clean_palm_csv(input_path, output_path):
    df = pd.read_csv(input_path, skiprows=2)
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    df = df.drop(0).reset_index(drop=True)
    df = df.drop(0).reset_index(drop=True)
    df = df.iloc[:-5]

    for column in df.columns:
        last_valid_value = df[column].dropna().iloc[-1] if not df[column].dropna().empty else None
        if last_valid_value is not None:
            df[column] = df[column].fillna(last_valid_value)

    if not df.empty and df.columns[0]:
        try:
            df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], errors='coerce')
            df = df.dropna(subset=[df.columns[0]])
            df[df.columns[0]] = df[df.columns[0]].dt.strftime('%Y-%m-%d')
        except Exception as e:
            pass

    df.to_csv(output_path, index=False)

def sync_and_dedup_csv_to_gsheet(csv_path, gsheet_id, sheet_title):
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
                pass

        df_all = pd.concat([df_old, df_new], ignore_index=True).dropna(how='all')
    except Exception as e:
        df_all = df_new

    df_all = df_all.drop_duplicates(subset=[df_all.columns[0]]).sort_values(by=df_all.columns[0])

    wks.clear()
    wks.set_dataframe(df_all, (1,1))

def main_workflow():
    risi_palm_oil_link = 'https://dashboard.fastmarkets.com/sw/x2TtMTTianBBefSdGCeZXc/palm-oil-global-prices'

    fetch_RISI_data(risi_palm_oil_link)

    download_dir_path = DOWNLOAD_DIR
    new_filename = "Palm_original.csv"
    downloaded_file_path = wait_file_and_rename(download_dir_path, ".csv", new_filename)

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

if __name__ == "__main__":
    main_workflow()
