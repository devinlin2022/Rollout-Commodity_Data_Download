import os
import time
import base64
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pygsheets

# --- 配置区域 ---
# 从环境变量获取敏感信息
RISI_USERNAME = os.getenv('RISI_USERNAME')
RISI_PASSWORD = os.getenv('RISI_PASSWORD')

# 服务账户密钥路径 (由 GitHub Actions 工作流创建)
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

# 下载目录 (与 GitHub Actions 工作流中的设置保持一致)
DOWNLOAD_DIR = "/tmp/downloads"

# Chromedriver 路径 (由 GitHub Actions 工作流设置)
CHROMEDRIVER_PATH = os.getenv('CHROMEDRIVER_PATH', '/usr/bin/chromedriver')

def save_pdf(driver, path):
    """使用 CDP 命令将当前页面保存为 PDF"""
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

def wait_for_download_complete(directory, timeout=60):
    """
    等待下载完成的辅助函数。
    它会检查下载目录，直到不再有 .crdownload (Chrome临时下载文件) 文件为止。
    """
    start_time = time.time()
    print("正在等待文件下载完成...")
    while time.time() - start_time < timeout:
        # 检查目录中是否有任何以 .crdownload 结尾的文件
        if not any(file.endswith('.crdownload') for file in os.listdir(directory)):
            # 找到最新的一个CSV文件作为下载成功的文件
            csv_files = [f for f in os.listdir(directory) if f.endswith('.csv')]
            if csv_files:
                latest_file = max([os.path.join(directory, f) for f in csv_files], key=os.path.getmtime)
                print(f"文件下载成功: {latest_file}")
                return latest_file
        time.sleep(1) # 每隔1秒检查一次
    raise Exception(f"错误：文件在 {timeout} 秒内未下载完成。")
    
def fetch_RISI_data(link):
    """启动 Selenium, 登录并下载 CSV 文件"""
    options = Options()
    options.binary_location = '/usr/bin/chromium-browser'
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage') # 在 Docker/GitHub Actions 环境中推荐使用
    options.add_argument("window-size=1920,1080") # 设置窗口大小以确保元素可见

    # 设置下载首选项
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safeBrowse.enabled": True
    }
    options.add_experimental_option('prefs', prefs)
    
    # 在 GitHub Actions 环境中使用特定的 Chromedriver
    service = webdriver.chrome.service.Service(executable_path=CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(link)
        wait = WebDriverWait(driver, 60) # 增加等待时间以应对网络延迟

        # 1. 登录 (为所有交互元素添加显式等待)
        print("等待登录页面加载...")
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#userEmail'))).send_keys(RISI_USERNAME)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#password'))).send_keys(RISI_PASSWORD)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#login-button'))).click()
        print("登录成功...")

        # 2. 点击导出菜单 (您的选择器未变)
        print("等待并点击第一个导出相关按钮...")
        # 等待第一个按钮出现并变为可点击状态
        time.sleep(10)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#cells-container > fui-grid-cell > fui-widget")))
        print("找到表格")
        first_button = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#cells-container > fui-grid-cell > fui-widget > header > fui-widget-actions > div:nth-child(1) > button')))        
        first_button.click()
        time.sleep(2)
        
        # 3. 点击下载 CSV 选项 (您的选择器未变)
        print("等待并点击下载CSV选项...")
        # 等待下载菜单项出现并变为可点击状态
        # download_option = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#mat-menu-panel-3 > div > div > div:nth-child(2) > fui-export-dropdown-item:nth-child(3) > button > span")))
        download_option = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#mat-menu-panel-3 > div > div > div:nth-child(2) > fui-export-dropdown-item:nth-child(3) > button > span")))
        download_option.click()
        print("下载命令已点击!")
        
        # 4. 等待文件下载完成 (移除了 time.sleep(10))
        # 使用新的、更可靠的函数来等待下载结束
        wait_for_download_complete(DOWNLOAD_DIR, timeout=120) # 等待最多2分钟

    except Exception as e:
        print(f"在获取数据时发生错误: {e}")
        # 发生错误时截图，帮助调试
        error_screenshot_path = os.path.join(DOWNLOAD_DIR, 'error_screenshot.png')
        driver.save_screenshot(error_screenshot_path)
        print(f"错误截图已保存至: {error_screenshot_path}")
        raise
    finally:
        print("任务完成，关闭浏览器。")
        driver.quit()

def wait_for_file_and_rename(old_ext=".csv", new_filename="Palm.csv", timeout=60):
    """在指定目录中等待文件下载完成，并将其重命名"""
    time_counter = 0
    while time_counter < timeout:
        files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith(old_ext) and not f.startswith('.')]
        if files:
            old_file = os.path.join(DOWNLOAD_DIR, files[0])
            new_file = os.path.join(DOWNLOAD_DIR, new_filename)
            os.rename(old_file, new_file)
            print(f"文件已找到并重命名为: {new_file}")
            return new_file
        time.sleep(1)
        time_counter += 1
    raise FileNotFoundError(f"错误：在 {timeout} 秒内未在目录 {DOWNLOAD_DIR} 中找到 {old_ext} 文件。")


def clean_palm_csv(input_path, output_path):
    """清理下载的 CSV 文件"""
    df = pd.read_csv(input_path, skiprows=2)
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    # 根据实际文件内容调整需要跳过的行数
    df = df.drop(index=df.index[:2]).reset_index(drop=True)
    df = df.iloc[:-5]

    # 向前填充缺失值
    df = df.fillna(method='ffill')
    df.to_csv(output_path, index=False)
    print(f"CSV 数据已清理并保存到: {output_path}")


def sync_to_gsheet(csv_path, gsheet_id, sheet_title):
    """将清理后的 CSV 同步到 Google Sheet 并去重"""
    try:
        df_new = pd.read_csv(csv_path)
        
        # 使用服务账户进行授权
        gc = pygsheets.authorize(service_file=SERVICE_ACCOUNT_FILE)
        sh = gc.open_by_key(gsheet_id)
        wks = sh.worksheet_by_title(sheet_title)
        
        print("成功连接到 Google Sheet。")

        try:
            df_old = wks.get_as_df(has_header=True, include_tailing_empty=False)
            df_all = pd.concat([df_old, df_new], ignore_index=True)
        except Exception:
            df_all = df_new

        # 转换日期列以正确去重和排序
        date_column_name = df_all.columns[0]
        df_all[date_column_name] = pd.to_datetime(df_all[date_column_name])
        
        # 按日期降序排序并删除重复项，保留最新的记录
        df_all = df_all.sort_values(by=date_column_name, ascending=False).drop_duplicates(subset=df_all.columns.drop(date_column_name), keep='first')
        
        wks.clear()
        wks.set_dataframe(df_all, (1,1), nan='')
        print("数据已成功同步到 Google Sheet 并完成去重。")

    except Exception as e:
        print(f"同步到 Google Sheet 时出错: {e}")
        raise

def main():
    """主执行函数"""
    print("自动化任务开始...")
    # 确保下载目录存在
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)
        print(f"已创建下载目录: {DOWNLOAD_DIR}")
    
    # 1. 获取数据
    fetch_RISI_data('https://dashboard.fastmarkets.com/sw/x2TtMTTianBBefSdGCeZXc/palm-oil-global-prices')
    
    # 2. 验证下载并重命名
    downloaded_file = wait_for_file_and_rename()
    
    # 3. 清理数据
    cleaned_file_path = os.path.join(DOWNLOAD_DIR, "Palm_cleaned.csv")
    clean_palm_csv(downloaded_file, cleaned_file_path)
    
    # 4. 同步到 Google Sheet
    sync_to_gsheet(
        csv_path=cleaned_file_path,
        gsheet_id='1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE', # 请确认这是正确的ID
        sheet_title='Palm Oil Price'
    )
    print("自动化任务成功完成！✅")

if __name__ == "__main__":
    main()
