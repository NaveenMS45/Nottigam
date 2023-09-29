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
    # ---- Find No.of Iteration ----
    total_rows = driver.find_elements(By.XPATH, '//table[@id="transactionTable"]/tbody/tr')
    rows = len(total_rows)
    print(rows)
    time.sleep(5)
    driver.find_element(By.ID, 'Menu_Workflow').click()
    do_filter()
    time.sleep(5)
    try:
        for d in range(rows - 1):
            if d == 0:
                pn = driver.find_elements(By.XPATH, '//table[@id="transactionTable"]/tbody/tr')[1].find_element(
                    By.TAG_NAME,
                    'a')
                print(pn.text)
                pn.click()
                # ---- Click POLICY FILE ----
                driver.find_element(By.ID, 'Tab_Documents').click()
                time.sleep(5)
                # ---- Click RENEWAL PACKAGE ----
                a_btn = driver.find_elements(By.XPATH, '//*[@id="rowItemContainer0"]')[0].find_elements(By.TAG_NAME,'a')
                for k in range(len(a_btn)):
                    if a_btn[k].text == "Renewal Package":
                        print(a_btn[k].text)
                        a_btn[k].click()
                time.sleep(7)
                # ---- Click Return Home ----
                driver.find_element(By.ID, 'Return').click()
            else:
                # ---- Next Iteration ----
                do_filter()
                # ---- Iterate Table ----
                total_rows = driver.find_elements(By.XPATH, '//table[@id="transactionTable"]/tbody/tr')
                rows = len(total_rows)
                try:
                    # ---- Click POLICY FILE ----
                    time.sleep(7)
                    driver.find_element(By.ID, 'Tab_Documents').click()
                    time.sleep(5)
                    # ---- Click RENEWAL PACKAGE ----
                    a_btn = driver.find_elements(By.XPATH, '//*[@id="rowItemContainer0"]')[0].find_elements(
                        By.TAG_NAME,
                        'a')
                    for k in range(len(a_btn)):
                        if a_btn[k].text == "Renewal Package":
                            print(a_btn[k].text)
                            a_btn[k].click()
                    time.sleep(5)
                    # ---- Click Return Home ----
                    driver.find_element(By.ID, 'Return').click()
                    break
                except IndexError as e:
                    print("List Index out of range : " + str(e))
                    driver.close()
        driver.close()
    except IndexError as e:
        print("List Index out of range : " + str(e))
        driver.close()


get_details()
driver.quit()


def extract_text_pdfplumber(source):
    text = ""
    file = source
    print(file)
    pdf = pdfplumber.open(file)
    for pageNo in range(3, 7):
        text += pdf.pages[pageNo].extract_text() + "\n"
    pdf.close()
    # print(text)
    return text


def extraction_regex(text):
    policy_number, eff_date = [], []
    pattern = r"Effective\s*Date:(?P<EffDate>\d{1,2}\/\d{1,2}\/\d{2,4})[\S\s]+Policy\sNumber\:\s(?P<PolicyNum>[A-Z\d]+)"
    result = re.finditer(pattern, text)
    for m in result:
        # print(m['EffDate'])
        # print(m['PolicyNum'])
        policy_number.append(m['PolicyNum'])
        eff_date.append(m['EffDate'])

    return policy_number, eff_date


key_list, value_list = [], []
directory_path = r"C:\Users\naveen.sampath\Desktop\Nottigam\PCIP\DownloadedPdf"
contents = os.listdir(directory_path)
# print(contents)
policyData = dict()
# print(contents)
i = 1
for item in contents:
    print(i)
    source_path = os.path.join(directory_path, item)
    if os.path.isfile(source_path):
        extracted_text = extract_text_pdfplumber(source_path)
        extracted_results = extraction_regex(extracted_text)
        key = extracted_results[0][0]
        value = extracted_results[1][0]
        key_list.append(key)
        value_list.append(value)
        policyData = {
            'Policy Number': key_list,
            'Effective Date': value_list
        }
    i += 1
print(policyData)
dataset = policyData
df = pd.DataFrame(dataset)
output_path = r'{}\Desktop\Nottigam\PCIP\Input\inputData.xlsx'.format(user)
print("Before Excel Write")
df.to_excel(output_path, index=False)
# with pd.ExcelWriter(output_path) as writer:
#     df.to_excel(writer, index=False)
print("After excel Write")
