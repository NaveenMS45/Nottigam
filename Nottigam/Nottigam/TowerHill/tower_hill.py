import datetime
import traceback
import pdfplumber
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import time
import pandas as pd
import re
import os

user = os.environ['USERPROFILE']
# download_directory = r"{}\Documents\Nottigam\InputPdf".format(user)
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    # 'download.default_directory': download_directory,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    "plugins.always_open_pdf_externally": True,
    'safebrowsing.enabled': True,
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False
})
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_argument("--remote-debugging-port=9222")
options.add_argument('disable-infobars')
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')
options.add_argument('start-maximized')
options.add_argument('disable-infobars')
options.add_argument("useAutomationExtension")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--browser-scripts")
options.add_argument('--ignore-certificate-errors')
options.add_argument("--disable-notifications")

path = 'https://auth.thig.com'
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 120)
driver.get("chrome://settings/")
driver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.8);')
driver.get(path)
driver.maximize_window()


def site_login():
    wait.until(EC.element_to_be_clickable((By.ID, 'inputGroupID'))).send_keys('FL6113')
    driver.find_element(By.ID, 'inputUserName').send_keys('KLERCH')
    driver.find_element(By.ID, 'inputPassword').send_keys('DarrSchackow5200b!!')
    driver.find_element(By.ID, 'authSubmit').click()
    time.sleep(3)


site_login()
time.sleep(10)