import os
import time
import base64
import pandas as pd
import io # 用于将字符串转为文件对象
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pygsheets
import pyperclip # <--- 新增库，用于访问剪贴板

# --- 配置区域 ---
# 从环境变量获取敏感信息
RISI_USERNAME = os.getenv('RISI_USERNAME')
RISI_PASSWORD = os.getenv('RISI_PASSWORD')

# 服务账户密钥路径
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

# Google Sheet 配置
GSHEET_ID = '1Qonj5yKwHVrxApUi7_N2CJtxj61rPfULXALrY4f8lPE'
GSHEET_TITLE = 'Palm Oil Price'

# Chromedriver 路径
CHROMEDRIVER_PATH = os.getenv('CHROMEDRIVER_PATH', '/usr/bin/chromedriver')
DOWNLOAD_DIR = "/tmp/downloads" # 保留此项用于截图

def get_data_from_clipboard(link):
    """
    启动 Selenium, 登录, 点击复制按钮, 并返回从系统剪贴板获取的数据。
    【注意】此函数严格按照您的要求，使用您原始的定位器。
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

        # 1. 登录 (使用您原始的定位器)
        print("等待登录页面加载...")
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#userEmail'))).send_keys(RISI_USERNAME)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#password'))).send_keys(RISI_PASSWORD)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#login-button'))).click()
        print("登录成功...")

        # 2. 点击导出菜单 (使用您原始的定位器)
        print("等待并点击第一个导出相关按钮...")
        time.sleep(10)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#cells-container > fui-grid-cell > fui-widget")))
        print("找到表格")
        first_button = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#cells-container > fui-grid-cell > fui-widget > header > fui-widget-actions > div:nth-child(1) > button')))
        first_button.click()
        time.sleep(3)

        # 3. 点击复制选项 (使用您原始的定位器)
        print("等待并点击'复制'选项...")
        # 此操作会将数据复制到系统的剪贴板
        driver.execute_script(f'document.querySelector("#mat-menu-panel-4 > div > div > div:nth-child(1) > fui-export-dropdown-item:nth-child(3) > button").click()')
        print("'复制'命令已点击!")
        
        # 等待一小段时间确保数据已进入剪贴板
        time.sleep(2)

        # 4. 从系统剪贴板读取数据
        clipboard_text = pyperclip.paste()
        print("成功从剪贴板获取到数据。")
        return clipboard_text

    except Exception as e:
        print(f"在获取数据时发生错误: {e}")
        error_screenshot_path = os.path.join(DOWNLOAD_DIR, 'error_screenshot.png')
        if not os.path.exists(DOWNLOAD_DIR):
            os.makedirs(DOWNLOAD_DIR)
        driver.save_screenshot(error_screenshot_path)
        print(f"错误截图已保存至: {error_screenshot_path}")
        raise
    finally:
        print("任务完成，关闭浏览器。")
        driver.quit()

def update_gsheet(clipboard_text, gsheet_id, sheet_title):
    """
    将剪贴板文本处理后，清空并覆盖到指定的 Google Sheet。
    """
    if not clipboard_text or not clipboard_text.strip():
        print("剪贴板数据为空，跳过 Google Sheet 更新。")
        return

    try:
        # 将剪贴板中的文本（通常是制表符分隔）读入 Pandas DataFrame
        # 使用 io.StringIO 可以让 pandas 直接读取字符串变量
        data_io = io.StringIO(clipboard_text)
        df = pd.read_csv(data_io, sep='\t') # "复制"功能通常是 tab 分隔
        print("已将剪贴板文本解析为数据帧 (DataFrame)。")
        print("数据预览:\n", df.head())
        
        # 授权并连接 Google Sheet
        print(f"准备同步数据到 Google Sheet '{sheet_title}'...")
        gc = pygsheets.authorize(service_file=SERVICE_ACCOUNT_FILE)
        sh = gc.open_by_key(gsheet_id)
        wks = sh.worksheet_by_title(sheet_title)
        print("成功连接到 Google Sheet。")

        # 核心操作: 1. 清空
        print("正在清空工作表...")
        wks.clear()

        # 核心操作: 2. 写入
        print("正在将新数据写入工作表...")
        wks.set_dataframe(df, (1, 1), nan='') # 从 A1 单元格开始写入
        
        print(f"数据已成功覆盖到 Google Sheet '{sheet_title}'。")

    except Exception as e:
        print(f"同步到 Google Sheet 时出错: {e}")
        raise

def main():
    """主执行函数"""
    print("自动化任务开始...")
    
    # 1. 从网站操作并将数据复制到剪贴板
    clipboard_content = get_data_from_clipboard('https://dashboard.fastmarkets.com/sw/x2TtMTTianBBefSdGCeZXc/palm-oil-global-prices')
    
    # 2. 将剪贴板数据同步到 Google Sheet
    update_gsheet(
        clipboard_text=clipboard_content,
        gsheet_id=GSHEET_ID,
        sheet_title=GSHEET_TITLE
    )
    
    print("自动化任务成功完成！✅")

if __name__ == "__main__":
    main()
