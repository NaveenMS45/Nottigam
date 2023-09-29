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
download_directory = r"{}\Documents\Nottigam\InputPdf".format(user)
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
# driver.get("chrome://settings/")
# driver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.8);')
driver.get(path)
driver.maximize_window()


def site_login():
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    driver.find_element(By.ID, 'j_username').send_keys('elliesmith')
    driver.find_element(By.ID, 'j_password').send_keys('Welcome2023!', Keys.ENTER)
    # driver.find_element(By.ID, 'SignIn').click()
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    login_successful = True
    return login_successful


site_login()


def do_filter():
    # ---- HOVER action using ActionChains ----
    time.sleep(5)
    action = ActionChains(driver)
    sub_ele = driver.find_element(By.ID, 'Menu_AgencyReports_AgentTransactionReport')
    action.move_to_element(driver.find_element(By.ID, 'Menu_AgencyReports')).move_to_element(sub_ele).click()
    action.perform()
    # ---- Select DATE ----
    driver.find_element(By.ID, 'FromTransDate').send_keys('09/01/2023')
    current_date = datetime.date.today()
    formatted_date = current_date.strftime("%m/%d/%Y")
    driver.find_element(By.ID, 'EndTransDate').send_keys(formatted_date)
    # ---- Cancellation Notice ----
    driver.find_element(By.ID, 'CancellationNotice').click()
    # ----Click RUN REPORT ----
    driver.find_element(By.ID, 'RunReport').click()
    time.sleep(2)


def get_details():
    eff_list = []
    exp_list = []
    do_filter()
    # ---- Find No.of Iteration ----
    total_rows = driver.find_elements(By.XPATH, '//table[@id="transactionTable"]/tbody/tr')
    rows = len(total_rows)
    print("Rows - 1 --> " + str(rows - 1))
    time.sleep(5)
    driver.find_element(By.ID, 'Menu_Workflow').click()
    time.sleep(5)
    try:
        for p in range(rows - 1):
            if p == 0 or p == rows - 1:
                continue
            do_filter()
            pn = driver.find_elements(By.XPATH, '//table[@id="transactionTable"]/tbody/tr')[p].find_element(By.TAG_NAME, 'a')
            print(pn.text)
            pn.click()
            # ---- Click POLICY FILE ----
            driver.find_element(By.ID, 'Tab_Documents').click()
            time.sleep(5)
            # ---- Click RENEWAL PACKAGE ----
            a_btn = driver.find_elements(By.XPATH, '//*[@id="rowItemContainer0"]')[0].find_elements(By.TAG_NAME, 'a')
            for k in range(len(a_btn)):
                if a_btn[k].text == "Cancellation Notice":
                    # ---- Extract Eff and Exp Date ----
                    text = driver.find_elements(By.XPATH, '//*[@id="ContainerHeader0"]')[0].find_elements(By.TAG_NAME, 'th')[0].text
                    print(text)
                    pattern = r"\((?P<EffDate>[\d\/\s]+)to\s(?P<ExpDate>[\d\/]+)"
                    result = re.finditer(pattern, text)
                    for m in result:
                        ef = m['EffDate']
                        day, month, year = ef.split('/')
                        day = day.lstrip('0')
                        month = month.lstrip('0')
                        formatted_ef = f"{day}/{month}/{year}"
                        eff_list.append(formatted_ef)
                        ex = m['ExpDate']
                        d, m, y = ex.split('/')
                        d = d.lstrip('0')
                        m = m.lstrip('0')
                        formatted_ex = f"{d}/{m}/{y}"
                        exp_list.append(formatted_ex)
                    a_btn[k].click()
            time.sleep(7)
            # ---- Click Return Home ----
            driver.find_element(By.ID, 'Return').click()
    except Exception as e:
        print(traceback.format_exc())
        driver.close()
        return eff_list, exp_list
    return eff_list, exp_list



def extract_text_pdfplumber(source):
    text = ""
    file = source
    print(file)
    pdf = pdfplumber.open(file)
    for pageNo in range(0, 3):
        text += pdf.pages[pageNo].extract_text() + "\n"
    pdf.close()
    # print(text)
    return text


def extraction_regex(text):
    policy_number, min_due_amount, due_date = [], [], []
    pattern = r"(?i)Payment\sDue\sDate\:\s(?P<DueDate>\d{2}\/\d{2}\/\d{2,4})\nPayment\sAmount\sDue\:(?P<DueAmt>[\s]\$[\d,\.]+).*(?:\n.*)*\s*Policy Number: (?P<PolNum>[A-Z\d]+)"
    result = re.finditer(pattern, text)
    for m in result:
        # print(m['EffDate'])
        # print(m['PolicyNum'])
        policy_number.append(m['PolNum'])
        due_date.append(m['DueDate'])
        min_due_amount.append(m['DueAmt'])
        break

    return policy_number, due_date, min_due_amount


eff_list, exp_list = get_details()
policyNumList = []
dueDateList = []
dueAmtList = []
sourceList = []
directory_path = r"{}\Documents\Nottigam\InputPdf".format(user)
contents = os.listdir(directory_path)
print(contents)
policyData = dict()
print(contents)
for item in contents:
    source_path = os.path.join(directory_path, item)
    if os.path.isfile(source_path):
        extracted_text = extract_text_pdfplumber(source_path)
        extracted_results = extraction_regex(extracted_text)
        pol_no = extracted_results[0][0]
        # print(source_path)
        # source_path = source_path.split("\\")
        # fn = source_path[8]
        new_name = r"CANCEL_{}.pdf".format(pol_no)
        new_path = os.path.join(directory_path, new_name)
        os.rename(source_path, new_path)
        sourceList.append(new_path)
        print(sourceList)
        policyNumList.append(pol_no)
        dueDateList.append(extracted_results[1][0])
        dueAmtList.append(extracted_results[2][0])

policyData = {
    "policy_number": policyNumList,
    "minimum_due": dueAmtList,
    "due_date": dueDateList,
    "eff_date": eff_list,
    "exp_date": exp_list,
    "download_path": sourceList

}
print(policyData)
dataset = policyData
df = pd.DataFrame(dataset)
output_path = r"{}\Documents\Nottigam\Extracted Input\extracted_data_NonPayment.xlsx".format(user)
print("Before Excel Write")
df.to_excel(output_path, index=False)
# with pd.ExcelWriter(output_path) as writer:
#     df.to_excel(writer, index=False)
print("After excel Write")
