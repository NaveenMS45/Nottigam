import datetime
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
print(user)

download_directory = r"{}\Desktop\Nottigam\PCIP\DownloadedPdf".format(user)
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    'download.default_directory': download_directory,
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

path = 'https://thekey.contributionship.com/innovation'
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 120)
driver.get("chrome://settings/")
driver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.8);')
driver.get(path)
driver.maximize_window()


def login():
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    driver.find_element(By.ID, 'j_username').send_keys('elliesmith')
    driver.find_element(By.ID, 'j_password').send_keys('Welcome2023!', Keys.ENTER)
    # driver.find_element(By.ID, 'SignIn').click()
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    login_successful = True
    return login_successful


login()


def do_filter():
    # ---- HOVER action using ActionChains ----
    action = ActionChains(driver)
    sub_ele = driver.find_element(By.ID, 'Menu_AgencyReports_AgentTransactionReport')
    action.move_to_element(driver.find_element(By.ID, 'Menu_AgencyReports')).move_to_element(sub_ele).click()
    action.perform()
    # ---- Select DATE ----
    driver.find_element(By.ID, 'FromTransDate').send_keys('09/01/2023')
    current_date = datetime.date.today()
    formatted_date = current_date.strftime("%m/%d/%Y")
    print(formatted_date)
    driver.find_element(By.ID, 'EndTransDate').send_keys(formatted_date)
    # ---- Select Renewal ----
    driver.find_element(By.ID, 'Renewal').click()
    # ----Click RUN REPORT ----
    driver.find_element(By.ID, 'RunReport').click()


def get_details():
    do_filter()


get_details()
wait.until(EC.presence_of_element_located((By.XPATH, '//a[text()="Export"]'))).click()
tableLink = wait.until(EC.presence_of_element_located((By.ID, 'downloadTransReportLink'))).find_element(By.TAG_NAME,'a').get_attribute('href')
driver.execute_script(f"""
    var file_url = '{tableLink}';
    var a = document.createElement('a');
    a.href = file_url;
    a.download = 'table.txt';
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
""")
tableFileName = str(tableLink).split('/')[10].split('&')[0]
print()


# # def get_details():
# #     for i in range(18):
# #         print(i)
#         # # ---- HOVER action using ActionChains ----
#         # action = ActionChains(driver)
#         # sub_ele = driver.find_element(By.ID, 'Menu_AgencyReports_AgentTransactionReport')
#         # action.move_to_element(driver.find_element(By.ID, 'Menu_AgencyReports')).move_to_element(sub_ele).click()
#         # action.perform()
#         # # ---- Select DATE ----
#         # driver.find_element(By.ID, 'FromTransDate').send_keys('09/01/2023')
#         # current_date = datetime.date.today()
#         # formatted_date = current_date.strftime("%m/%d/%Y")
#         # print(formatted_date)
#         # driver.find_element(By.ID, 'EndTransDate').send_keys(formatted_date)
#         # # ---- Select Renewal ----
#         # driver.find_element(By.ID, 'Renewal').click()
#         # # ----Click RUN REPORT ----
#         # driver.find_element(By.ID, 'RunReport').click()
#         # ---- Iterate Table ----
#         # total_rows = driver.find_elements(By.XPATH, '//table[@id="transactionTable"]/tbody/tr')
#         # rows = len(total_rows)
#         if i == 0:
#             pn = driver.find_elements(By.XPATH, '//table[@id="transactionTable"]/tbody/tr')[1].find_element(By.TAG_NAME,
#                                                                                                             'a')
#             pn.click()
#             # ---- Click POLICY FILE ----
#             driver.find_element(By.ID, 'Tab_Documents').click()
#             time.sleep(5)
#             # ---- Click RENEWAL PACKAGE ----
#             a_btn = driver.find_elements(By.XPATH, '//*[@id="rowItemContainer0"]')[0].find_elements(By.TAG_NAME, 'a')
#             for k in range(len(a_btn)):
#                 if a_btn[k].text == "Renewal Package":
#                     print(a_btn[k].text)
#                     a_btn[k].click()
#             time.sleep(7)
#             # ---- Click Return Home ----
#             driver.find_element(By.ID, 'Return').click()
#         else:
#             # ---- Next Iteration ----
#             # ---- Iterate Table ----
#             total_rows = driver.find_elements(By.XPATH, '//table[@id="transactionTable"]/tbody/tr')
#             rows = len(total_rows)
#             for j in range(2, rows - 1):
#                 print(j)
#                 print(i)
#                 print(i + j)
#                 pn = total_rows[i + j].find_element(By.TAG_NAME, 'a')
#                 print(pn.text)
#                 pn.click()
#                 # ---- Click POLICY FILE ----
#                 driver.find_element(By.ID, 'Tab_Documents').click()
#                 time.sleep(5)
#                 # ---- Click RENEWAL PACKAGE ----
#                 a_btn = driver.find_elements(By.XPATH, '//*[@id="rowItemContainer0"]')[0].find_elements(By.TAG_NAME, 'a')
#                 for k in range(len(a_btn)):
#                     if a_btn[k].text == "Renewal Package":
#                         print(a_btn[k].text)
#                         a_btn[k].click()
#                 time.sleep(5)
#                 # ---- Click Return Home ----
#                 driver.find_element(By.ID, 'Return').click()
#                 break


# get_details()
# driver.quit()


# def extract_text_pdfplumber(source):
#     text = ""
#     file = source
#     print(file)
#     pdf = pdfplumber.open(file)
#     # for page in pdf.pages:
#     #     text += page.extract_text() + "\n"
#     page = pdf.pages[4]
#     text += page.extract_text() + "\n"
#     pdf.close()
#     return text


def extraction_regex(text):
    policy_number, eff_date = [], []
    # dates = re.findall(r"(?P<EffDate>\d{2}\/\d{2}\/\d{2,4})", text)
    # for date in dates:
    #     print(date)
    #     eff_date.append(date)
    #     break
    # policy_pattern = re.compile(r"Policy\sNumber\:\s(?P<PolicyNum>[A-Z\d]+)")
    # policy = re.search(policy_pattern, text)
    # print(policy.group())
    # policy_number.append(policy.group())
    # pattern = r"(?P<EffDate>[\d/]+)\s.*Policy\sNumber\:\s(?P<PolicyNum>[A-Z\d]+)"
    # result = re.finditer(pattern, text)
    # for i in result:
    #     print(i['EffDate'])
    #     print(i['PolicyNum'])
    #     print("hi")
    #     policy_number.append(i['PolicyNum'])
    #     eff_date.append(i['EffDate'])
    #
    # return policy_number, eff_date


# if login():
#     directory_path = r"C:\Users\naveen.sampath\Desktop\Nottigam\PCIP\DownloadedPdf"
#     contents = os.listdir(directory_path)
#     for item in contents:
#         source = os.path.join(directory_path, item)
#         print(source)
#         if os.path.isfile(os.path.join(directory_path, item)):
#             print(f'File: {item}')
#             extracted_text = extract_text_pdfplumber(source)
#             print(extracted_text)
#             policy_number, eff_date = extraction_regex(extracted_text)
#             print('Policy Number ---> ' + str(policy_number) + 'Eff date ---> ' + str(eff_date))
# pdf_path = r"C:\Users\naveen.sampath\Desktop\Nottigam\PCIP\DownloadedPdf\output_gwops_0vYZvD8u8xa7xcF.pdf"
# extracted_text = extract_text_pdfplumber(pdf_path)
# print(extracted_text)
# extraction_regex(extracted_text)
time.sleep(10)


