from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime, timedelta
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os
import shutil

# 1. 更改网页显示语言为英语
def change_language_to_english():
    try:
        # 打开卖家中心首页
        driver.get("https://sellercentral.amazon.com/home")

        # 等待页面加载完成
        wait = WebDriverWait(driver, 10)
        language_dropdown = wait.until(
            EC.presence_of_element_located((By.ID, "sc-lang-switcher"))
        )

        # 点击选择语言的下拉菜单
        language_dropdown.click()

        # 选择英语（English）作为显示语言
        english_option = wait.until(
            EC.presence_of_element_located((By.XPATH, '//option[@value="en_US"]'))
        )
        english_option.click()

        # 确认并保存设置
        save_button = wait.until(
            EC.presence_of_element_located((By.XPATH, '//input[@value="Save"]'))
        )
        save_button.click()

        print("Language changed to English successfully.")
    except Exception as e:
        print(f"Error occurred while changing language to English: {e}")

change_language_to_english()



# 指定浏览器驱动和登录URL
driver = webdriver.Chrome('D:/运营/chromedriver')  # 更改为您的实际chromedriver路径
login_url = 'https://sellercentral.amazon.com/home'  # 更改为您需要登录的Amazon卖家中心网址

# 打开Amazon卖家中心登录页面
driver.get(login_url)

# 等待页面加载完成
time.sleep(2)

# 定位用户名输入框并输入用户名
username_input = driver.find_element_by_id('ap_email')
username_input.send_keys('3381680581@qq.com')  # 替换为您的实际用户名

# 定位密码输入框并输入密码
password_input = driver.find_element_by_id('ap_password')
password_input.send_keys('sailingstar2022')  # 替换为您的实际密码

# 定位并点击登录按钮
signin_button = driver.find_element_by_id('signInSubmit')
signin_button.click()


# 执行自动下载过程
change_language_to_english()
open_reports()
open_business_reports()
open_detail_page_sales_and_traffic()
set_date_range()
download_csv_report()

# 指定下载文件夹路径
download_directory = 'path/to/your/download/directory'  # 更改为您的实际下载目录

# 计算距离今天最近的过去一个星期六的日期
today = datetime.today()
last_saturday = today - timedelta(days=(today.weekday() + 2) % 7)

# 监控下载文件夹并更改下载的文件名
monitor_downloads_folder(download_directory, last_saturday)
