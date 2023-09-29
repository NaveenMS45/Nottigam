import os
import re
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import pandas as pd
import time
from selenium.webdriver.support.select import Select
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import xmltodict
import json
import os
import cv2
import re
from pdf2image import convert_from_path
import pytesseract
from PIL import ImageEnhance
from PIL import Image
from pytesseract import Output


user = os.environ['USERPROFILE']
data_list = []
path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
options = webdriver.ChromeOptions()
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches", ["enable-automation"])
download_dir = os.path.join(os.environ['USERPROFILE'], r"{}\Desktop\Nottigam\NationWide\DownloadedPDF".format(user))
print(download_dir)
prefs = {"credentials_enable_service": False, "profile.password_manager_enabled": False,
         "download.default_directory": download_dir,
         "download.prompt_for_download": False,
         "download.directory_upgrade": True,
         "plugins.always_open_pdf_externally": True,
         }
options.add_experimental_option("prefs", prefs)

path = 'https://agentcenter.nationwide.com/WorkspaceAC/login'

# chromedriver_path = r"C:\Users\pravin.marasamy\Downloads\chromedriver-win64\chromedriver.exe"
# driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
driver = webdriver.Chrome(options=options)

wait = WebDriverWait(driver, 120)
driver.get(path)
driver.maximize_window()

all_data = pd.DataFrame(columns=["Policy Number", "Minimum Due", "Payment Acceptance Date"])


def clear_input_field(ele):
    ele.click()
    ele.send_keys(Keys.CONTROL + "a")
    ele.send_keys(Keys.DELETE)


def filterName(selectMode):
    wait.until(EC.element_to_be_clickable((By.ID, "bolt-select-root")))
    select1 = Select(driver.find_element(By.ID, "bolt-select-root"))
    for i in select1.options:
        if i.text in selectMode or selectMode in i.text:
            i.click()
            break

def extractDataFromDict(d):
    df = pd.DataFrame(d)
    df1 = df[(df.conf != '-1') & (df.text != ' ') & (df.text != '')]
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    sorted_blocks = df1.groupby('block_num').first().sort_values('top').index.tolist()
    text = ''
    for block in sorted_blocks:
        curr = df1[df1['block_num'] == block]
        sel = curr[curr.text.str.len() > 3]
        char_w = (sel.width / sel.text.str.len()).mean()
        prev_par, prev_line, prev_left = 0, 0, 0
        text = ''
        for ix, ln in curr.iterrows():
            if prev_par != ln['par_num']:
                text += '\n'
                prev_par = ln['par_num']
                prev_line = ln['line_num']
                prev_left = 0
            elif prev_line != ln['line_num']:
                text += '\n'
                prev_line = ln['line_num']
                prev_left = 0
            added = 0
            if ln['left'] / char_w > prev_left + 1:
                added = int((ln['left']) / char_w) - prev_left + 1
                text += ' ' * added
            text += ln['text'] + ' '
            prev_left += len(ln['text']) + added + 1
        text += '\n'
    return text
def linked_policy():
    global Policy_Number, Minimum_Due, Due_Date
    pdf_filename = None
    new_pdf_filename = None

    container = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "printContainer")))
    bolt_rows = container.find_elements(By.CLASS_NAME, 'bolt-row')

    for row in bolt_rows:
        text = row.text
        if 'minimum due' in text.lower():
            Minimum_Due = row.find_elements(By.TAG_NAME, 'div')[1].text
        if 'payment acceptance date' in text.lower():
            Due_Date = row.find_elements(By.TAG_NAME, 'div')[1].text
        if 'policies' in text.lower():
            Policy_Number = row.find_elements(By.TAG_NAME, 'div')[1].text

    for row in bolt_rows:
        text = row.text
        if "billing account number" in text.lower():
            wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "mockLinkForAccessibility"))).click()
            time.sleep(4)
            if wait.until(EC.number_of_windows_to_be(2)):
                driver.switch_to.window(driver.window_handles[1])
                try:
                    ele = WebDriverWait(driver, 12).until(EC.element_to_be_clickable((By.ID, 'accountHistoryRow')))
                    row_elements = ele.find_elements(By.TAG_NAME, 'tr')

                    for row in row_elements:
                        if 'warning notice' in row.text.lower():
                            time.sleep(2)
                            row.find_element(By.TAG_NAME, 'a').click()
                            time.sleep(2)
                            print("Warning Notice Clicked")
                            pdf_files = [f for f in os.listdir(download_dir) if f.endswith('.pdf')]
                            print(pdf_files)
                            if pdf_files:
                                pdf_file_path = os.path.join(download_dir, pdf_files[0])
                                textValue = ""
                                pages = convert_from_path(
                                    pdf_file_path,
                                    poppler_path=os.environ["USERPROFILE"]
                                                 + r"\Documents\EleviantRPA\Tools\poppler-0.68.0\bin",
                                )
                                for i, page in enumerate(pages):
                                    imagePath = f"temp{i}.jpg"
                                    page.save(imagePath)
                                    # preprocessImage(imagePath)
                                    image = cv2.imread(imagePath)
                                    cv2.imwrite(imagePath, image)
                                    custom_config = r'-l eng --oem 1 --psm 6'
                                    pytesseract.pytesseract.tesseract_cmd = path_to_tesseract
                                    d = pytesseract.image_to_data(imagePath, config=custom_config, output_type=Output.DICT)
                                    textValue += extractDataFromDict(d)
                                    print(textValue)
                                    print()


                            driver.switch_to.window(driver.window_handles[0])
                            time.sleep(7)
                            break
                except Exception as e:
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    time.sleep(2)
                    break


def NationWide():
    wait.until(EC.element_to_be_clickable((By.ID, 'UID'))).send_keys("AEB91G")
    wait.until(EC.element_to_be_clickable((By.ID, 'PWD'))).send_keys("Welcome1", Keys.ENTER)
    action = ActionChains(driver)
    action.send_keys(Keys.TAB)
    action.perform()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.TAG_NAME, "bolt-button")))
    ele_Continue = driver.find_elements(By.TAG_NAME, 'bolt-button')
    for ele in ele_Continue:
        if 'continue to site' in ele.text.lower():
            ele.click()
            break
    billingEle = wait.until(EC.element_to_be_clickable((By.ID, "billing")))
    billingEle.click()
    notificationFilterElement = driver.execute_script("return document.querySelector('bolt-tabpanel > app-billing-notices > app-notification-filter').children[0].getElementsByTagName('div')[2].getElementsByTagName('bolt-select')[0].shadowRoot.querySelector('#bolt-select-wrapper')")
    action.click(notificationFilterElement).perform()
    action.send_keys(Keys.DOWN)
    action.send_keys(Keys.DOWN)
    action.click().perform()
    time.sleep(2)
    select_element = driver.find_element(By.ID, "mat-select-2")
    select_element.click()

    dropdown_options = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//mat-option/span")))
    desired_text = "20"
    for option in dropdown_options:
        if option.text == desired_text:
            option.click()
            break
    time.sleep(5)
    count = driver.execute_script("return document.querySelector('bolt-tabpanel > app-billing-notices > div.table-scroll-container > table').children[1].getElementsByTagName('tr').length")
    for i in range(count):
        driver.execute_script(f"document.querySelector('bolt-tabpanel > app-billing-notices > div.table-scroll-container > table').children[1].getElementsByTagName('tr')[{i}].querySelector('td').scrollIntoView()")
        driver.execute_script(f"document.querySelector('bolt-tabpanel > app-billing-notices > div.table-scroll-container > table').children[1].getElementsByTagName('tr')[{i}].querySelector('td').click();")
        # print("Click")
        time.sleep(2)
        linked_policy()
    return True


try:
    Flag = NationWide()
except Exception as e:
    print(str(e))


all_data = pd.DataFrame(data_list)
print(all_data)
print()

