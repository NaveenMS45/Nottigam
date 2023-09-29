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

options = webdriver.ChromeOptions()
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
prefs = {"credentials_enable_service": False, "profile.password_manager_enabled": False}
options.add_experimental_option("prefs", prefs)

path = 'https://lmidp.libertymutual.com/as/authorization.oauth2?client_id=uscm_iaportalpl_1&response_type=code&scope=openid%20profile&state=OpenIdConnect.AuthenticationProperties%3DL8VJvkteJtJcxV4sSEnErjmcgYovTSsL7LQgthzuEWHR5Y5EnRKwt551LKOID368Nbbt5dhD8utATGUwx8tCn134p-YENl64J7kO6vL9GKs-V36EL_WsJWuuts3bHKM8gcogqhqM0KFpHTzFYnKIH2DvSR8vuT8zWp2GssVMqV1KZ12_pt8ZIsbKa5Kj3JtmyU-znixzRTgxTYn2YVrBvg&response_mode=form_post&nonce=638266849546790907.MjBhZWM5OTctMzVjMy00NWE1LThmNWEtYTdjMGY1YmNhNDI5ZmUzOGUwZTctYmU0ZC00OGU1LTljYzQtNmZhMDZjODUyODEz&audience=uscm_iaportalpl_1&redirect_uri=https%3A%2F%2Fnow.agent.safeco.com%2Fidentity%2Fexternallogincallback%3FReturnUrl%3D%2Fstart&x-client-SKU=ID_NET461&x-client-ver=5.3.0.0'
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 120)
driver.get("chrome://settings/")
driver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.8);')
driver.get(path)
driver.maximize_window()

user = os.environ['USERPROFILE']
print(user)


def safco_login():
    wait.until(EC.element_to_be_clickable((By.ID, 'username'))).send_keys("EBOT00")
    wait.until(EC.element_to_be_clickable((By.ID, 'password'))).send_keys("Welcome2023$*")
    wait.until(EC.element_to_be_clickable((By.ID, 'submit1'))).click()
    time.sleep(60)
    login_successful = True
    return login_successful


def get_safe():
    policy_number, due_date, min_amount_due = [], [], []
    account = []
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(5)
    # Handle POP UP
    try:
        driver.find_element(By.XPATH, '//*[@id="lmig-Modal-1"]/section/button').click()
    except Exception as e:
        print(str(e))
    # Click Recent Activities
    driver.find_element(By.XPATH, '//*[text() = "Recent Activities"]').click()
    # No.of Handles ---> 2
    wait.until(EC.number_of_windows_to_be(2))
    handles = driver.window_handles
    driver.switch_to.window(handles[1])
    # SWITCH TO HANDLE ---> 1
    time.sleep(12)
    # wait.until(EC.presence_of_element_located((By.TAG_NAME,'body')))
    # CLICK BILLING ACTIVITIES
    driver.find_elements(By.XPATH, '//*[contains(text(), "Billing activities")]')[1].click()
    time.sleep(5)
    # SELECT DROPDOWN
    drp = driver.find_element(By.ID, 'date-selector-select-select')
    drp_days = Select(drp)
    drp_days.select_by_visible_text('Last 7 days')
    time.sleep(10)
    # CHOOSE PENDING NON-PAY CANCEL
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//*[@class='lm-FieldGroup-wrapper lmig-Popover-trigger']"))).find_elements(
        By.TAG_NAME, "button")[0].click()
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "lmig-Popover-overlay-content"))).find_elements(
        By.TAG_NAME, "label")[2].click()
    time.sleep(2)
    total_rows = driver.find_elements(By.XPATH, '//*[@class="lmig-Table-body"]/tr')
    print("Length of rows : " + str(len(total_rows)))
    try:
        for row in range((len(total_rows))):
            print(row)
            # driver.execute_script("return arguments[0].getElementsByTagName('a')[1].click()", total_rows[row])
            pn = total_rows[row].find_elements(By.TAG_NAME, 'a')[1]
            account_text = pn.text
            print(account_text)
            account.append(account_text)
            pn.click()
            # No.Of Handles --> 3
            wait.until(EC.number_of_windows_to_be(3))
            handles = driver.window_handles
            # SWITCH TO HANDLE ---> 2
            driver.switch_to.window(handles[2])
            driver.maximize_window()
            time.sleep(3)
            # selected_label = driver.find_elements(By.TAG_NAME, "tbody")[10].find_elements(By.TAG_NAME, "label")[3]
            selected_label = driver.find_element(By.ID, 'BillingAccountCustomerAction3')
            print(selected_label.text)
            selected_label.click()
            # Click GO Button
            driver.find_element(By.ID, 'btnGo').click()
            # driver.find_elements(By.TAG_NAME, "table")[11].find_element(By.TAG_NAME, "input").click()
            time.sleep(5)
            # TRY/CATCH
            try:
                view_btn = driver.find_element(By.XPATH, '//a[@id="BillingAccountNoticeShowView_link"]')
                print(view_btn.text)
                time.sleep(4)
                view_btn.click()
                time.sleep(8)
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                # ---- FIND DUE DATE ----
                due_text = driver.find_elements(By.TAG_NAME, 'tbody')[9].find_elements(By.TAG_NAME, 'td')[0].text
                print(due_text)
                # Pattern
                date_pattern = r'\d{1,2}/\d{1,2}/\d{4}'
                dates = re.findall(date_pattern, due_text)
                print(dates)
                if dates:
                    # extracted dates
                    for date in dates:
                        print(date)
                        due_date.append(date)
                    # ---- FIND MIN DUE AMOUNT ----
                    due_amount = driver.find_elements(By.TAG_NAME, 'tbody')[15].find_elements(By.TAG_NAME, 'tr')[1].find_elements(By.TAG_NAME, 'td')[3].text
                    print(due_amount)
                    min_amount_due.append(due_amount)
                    # ---- FIND POLICY NUMBER ----
                    try:
                        policy = driver.find_elements(By.TAG_NAME, 'tbody')[18].find_elements(By.TAG_NAME, 'tr')[1].find_elements(By.TAG_NAME, 'td')[1].text
                        print(str(policy))
                        if policy != ' ':
                            policy_number.append(policy)
                    except Exception as e:
                        print("Policy Number Not found --> " + str(e))
                    try:
                        policy = driver.find_elements(By.TAG_NAME, 'tbody')[18].find_elements(By.TAG_NAME, 'tr')[2].find_elements(By.TAG_NAME, 'td')[1].text
                        print(str(policy))
                        if policy != ' ':
                            policy_number.append(policy)
                    except Exception as e:
                        print("Policy Number Not found --> " + str(e))
                    try:
                        policy = driver.find_elements(By.TAG_NAME, 'tbody')[18].find_elements(By.TAG_NAME, 'tr')[3].find_elements(By.TAG_NAME, 'td')[1].text
                        print(str(policy))
                        if policy != ' ':
                            policy_number.append(policy)
                    except Exception as e:
                        print("Policy Number Not found --> " + str(e))
                    # END OF SWITCH TO HANDLE ---> 2
                    driver.close()
                    driver.switch_to.window(handles[1])
                else:
                    print("Due Date Not Found")
                    # due_date.append("Not Found")
                # # ---- FIND MIN DUE AMOUNT ----
                # due_amount = driver.find_elements(By.TAG_NAME, 'tbody')[15].find_elements(By.TAG_NAME, 'tr')[1].find_elements(
                #     By.TAG_NAME, 'td')[3].text
                # print(due_amount)
                # min_amount_due.append(due_amount)
                # # ---- FIND POLICY NUMBER ----
                # try:
                #     policy = driver.find_elements(By.TAG_NAME, 'tbody')[18].find_elements(By.TAG_NAME, 'tr')[1].find_elements(
                #         By.TAG_NAME, 'td')[1].text
                #     print(str(policy))
                #     if policy != ' ':
                #         policy_number.append(policy)
                # except Exception as e:
                #     print("Policy Number Not found --> " + str(e))
                # try:
                #     policy = driver.find_elements(By.TAG_NAME, 'tbody')[18].find_elements(By.TAG_NAME, 'tr')[2].find_elements(
                #         By.TAG_NAME, 'td')[1].text
                #     print(str(policy))
                #     if policy != ' ':
                #         policy_number.append(policy)
                # except Exception as e:
                #     print("Policy Number Not found --> " + str(e))
                # try:
                #     policy = driver.find_elements(By.TAG_NAME, 'tbody')[18].find_elements(By.TAG_NAME, 'tr')[3].find_elements(
                #         By.TAG_NAME, 'td')[1].text
                #     print(str(policy))
                #     if policy != ' ':
                #         policy_number.append(policy)
                # except Exception as e:
                #     print("Policy Number Not found --> " + str(e))
                # # END OF SWITCH TO HANDLE ---> 2
                # driver.close()
                # driver.switch_to.window(handles[1])
            except Exception as e:
                print("Can't fetch required data -->" + str(e))
                # policy_number.append("Not Found")
                # min_amount_due.append("Not Found")
                # due_date.append("Not Found")
                # END OF SWITCH TO HANDLE ---> 2
                driver.close()
                driver.switch_to.window(handles[1])

        time.sleep(10)
        # END OF SWITCH TO HANDLE --> 1
        driver.close()
        driver.switch_to.window(handles[0])
        return policy_number, due_date, min_amount_due, account
    except IndexError as e:
        print("List Index Out of Range : " + str(e))
        time.sleep(10)
        # END OF SWITCH TO HANDLE --> 1
        driver.close()
        driver.switch_to.window(handles[0])
        return policy_number, due_date, min_amount_due, account



if safco_login():
    policy_number, due_date, min_amount_due, account = get_safe()
    print(" Due Date --> " + str(due_date) + "Due Amount --> " + str(min_amount_due))
    print("Policy Number --> {}".format(policy_number))
    print("Excel before write")
    Input = pd.DataFrame({'Account/Policy': account, 'Policy Number': policy_number, 'Due Date': due_date, 'Minimum Amount Due': min_amount_due},
                         columns=['Account/Policy', 'Policy Number', 'Due Date', 'Minimum Amount Due'])
    print(Input)

    outputFilePath = r"{}\Desktop\Nottigam\safeco\Input\inputData.xlsx".format(user)
    print(outputFilePath)
    Input.to_excel(outputFilePath, index=False)
    # with pd.ExcelWriter(outputFilePath) as writer:
    #     Input.to_excel(writer, index=False)
    print("after excel write")

time.sleep(10)
driver.close()
driver.quit()
