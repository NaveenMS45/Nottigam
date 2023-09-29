import datetime
import traceback
import pdfplumber
from selenium.webdriver import ActionChains
from selenium.webdriver.support.select import Select
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

path = 'https://sfiprod.hostedinsurance.com/AgentPortal/login'
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 120)
driver.get("chrome://settings/")
driver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.8);')
driver.get(path)
driver.maximize_window()


def site_login():
    wait.until(EC.element_to_be_clickable((By.ID, 'j_username'))).send_keys('A061657SFI005')
    driver.find_element(By.XPATH, '//input[@type="password"]').send_keys('February723!', Keys.ENTER)
    # driver.find_element(By.XPATH, '//input[@type="submit"]').click()
    time.sleep(3)


site_login()


def site_action():
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    ref_ele = wait.until(EC.presence_of_element_located((By.XPATH, "//*[text()=' Reference Number']")))
    driver.execute_script("arguments[0].previousElementSibling.click()", ref_ele)
    time.sleep(6)
    driver.find_element(By.XPATH, '//input[@type="text" and @class="qtb"]').send_keys('p000009880', Keys.ENTER)
    time.sleep(5)
    driver.find_elements(By.XPATH, '//a[@class = "listColumnLink"]')[1].click()
    time.sleep(5)
    # Choose "DropDown"
    drop_down = driver.find_elements(By.XPATH, '//select[@class="cb2"]')[0]
    drp_value = Select(drop_down)
    drp_value.select_by_value('10/07/2023 - 10/07/2024')
    time.sleep(3)
    # Click "Go to Payment Screen"
    payment_ref = wait.until(EC.presence_of_element_located((By.XPATH, '//span[text()="Go to Payment Screen"]')))
    driver.execute_script("arguments[0].parentElement.click()", payment_ref)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    time.sleep(5)
    # Scroll down the page
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    # Click ---> "Payments"
    wait.until(EC.element_to_be_clickable((By.XPATH, '//td[text()="Payments"]'))).click()
    time.sleep(5)
    img_ref = driver.find_element(By.XPATH, '//img[@alt="View Details"]')
    driver.execute_script("arguments[0].parentElement.click()", img_ref)
    time.sleep(5)
    # Scroll down the page
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    # Applied Amount & Receipt Date
    # applied_amt = driver.find_element(By.XPATH, '//input[@id="f39:j_id1662:0:c178_1048"]').get_attribute('value')
    # print(applied_amt)
    # receipt_date = driver.find_element(By.XPATH, '//input[@class="tbro1" and @name="f39:j_id1662:0:c731_571"]').get_attribute('value')
    # print(receipt_date)


site_action()
time.sleep(10)