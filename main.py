#!/usr/bin/env python
import win32con
import json
from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import pandas as pd
from datetime import date
import datetime
from openpyxl import load_workbook
import json


options = Options()
options.add_argument("--user-data-dir=C:\\Users\\Moshe\\AppData\\Local\\Google\\Chrome\\User Data")
options.add_argument("--profile-directory=Profile 3")
options.add_argument("--no-sandbox")
chromedriver_autoinstaller.install()
browser = webdriver.Chrome(options=options)


# browser = webdriver.Chrome(r'C:\Users\Moshe\Downloads\chromedriver.exe', options=options)
# Initialize driver
# driver_path = r"C:\Users\Moshe\Downloads\chromedriver.exe"  # replace with the path to your ChromeDriver
# browser = webdriver.Chrome(executable_path=driver_path)

def get_formatted_date():
    if is_yesterday:
        yesterday = datetime.date.today() - datetime.timedelta(days=1)
        formatted_date = yesterday.strftime("%m_%d_%Y")
    else:
        today = datetime.date.today()
        formatted_date = today.strftime("%m_%d_%Y")

    return formatted_date


def login_to_actum(login_url, login_username, login_password):

    #time.sleep(1)
    browser.maximize_window()
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.NAME, 'username')))
    user_field = browser.find_element(By.NAME, "username")
    pass_field = browser.find_element(By.NAME, "password")
    signin_button = browser.find_element(By.XPATH, "//input[@type='submit']")
    # Clear the password field using keyboard actions
    ActionChains(browser).click(user_field).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(
        Keys.BACKSPACE).perform()
    ActionChains(browser).send_keys(login_username).perform()
    ActionChains(browser).click(pass_field).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(
        Keys.BACKSPACE).perform()
    ActionChains(browser).send_keys(login_password).perform()

    signin_button.click()
    # Wait for the next page to load and request validation code from the user
    time.sleep(0.5)
    if browser.find_elements(By.CSS_SELECTOR, "td.control[align='right']"):
        # Request validation code from the user
        val_code = input("Please enter your validation code: ")
        # Find the validation code field and submit button and perform actions
        val_field = browser.find_element(By.NAME, "valCode")
        submit_button = browser.find_element(By.XPATH, "//input[@type='submit']")
        val_field.send_keys(val_code)
        submit_button.click()
        time.sleep(0.5)
        print(" Step 1: login successful")
    else:
        # Check if there is an error message indicating incorrect credentials
        try:
            error_element = browser.find_element(By.ID, "APEX_ERROR_MESSAGE")
            # Find the error message within the error element
            error_message = "user or password is incorrect"
            # Print the error message
            print("Login failed. Error message:", error_message)
        except NoSuchElementException:
            # If no error message is found, assume login was successful
            print("Step 1: Login successful!")


def navigate_to_reports():
    # time.sleep(0.5)
    reports_tab = browser.find_element(By.LINK_TEXT, "Reporting")
    merchant_activity_report = browser.find_element(By.ID, "menuItem20")
    action = ActionChains(browser)
    action.move_to_element(reports_tab).click(merchant_activity_report).perform()


# Generate settlements and returns reports
def generate_reports(task):
    if is_yesterday:
        yesterday_button = WebDriverWait(browser, 20).until(
            EC.presence_of_element_located((By.ID, "rp_ydy"))
        )
        # click on 'Yesterday' button and then click on 'Submit' button
        yesterday_button.click()
        time.sleep(0.5)
    #   submit_button = browser.find_element(By.CLASS_NAME, "button1")
    #   time.sleep(1)
    #   browser.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", submit_button)
    #   time.sleep(1)
    #   browser.execute_script("arguments[0].click();", submit_button)

    # click submit by Moving the mouse to XY coordinates.
    pyautogui.moveTo(530, 530)  # for 1080p screen
    # pyautogui.moveTo(530, 514)
    # pyautogui.moveTo(409, 387) # for macbook
    pyautogui.click()

    # Wait for the Settlements / Returns pages to load
    time.sleep(1)
    # find Settlements/returns button
    task_button = browser.find_element(By.LINK_TEXT, task)
    time.sleep(1.5)
    # click on 'Settlements' or 'returns; button
    task_button.click()
    # Wait for the page to load
    time.sleep(1.5)

def navigate_to_utilities():
    # Wait for page to load
    time.sleep(1)
    # find elements
    reports_tab = browser.find_element(By.LINK_TEXT, "Utilities")
    log_out = browser.find_element(By.ID, "menuItem29")
    # create action chain object
    action = ActionChains(browser)
    # move to element on web page and then click on it
    action.move_to_element(reports_tab).click(log_out).perform()
def download_settlements(company_name):
    html = browser.page_source
    formatted_date = get_formatted_date()
    table = pd.read_html(html)[13]
    filename = company_name + ' Settlements_' + formatted_date + '.xlsx'
    sheet_name = 'Sheet_' + formatted_date
    if os.path.exists(filename):
        os.remove(filename)
    ''' 
    # Check if file exists
    if os.path.exists(filename):
        # Ask for confirmation to overwrite
        overwrite = input("The file already exists. Do you want to overwrite it? (yes/no): ")
        if overwrite.lower() == "yes":
            # Delete the existing file
            os.remove(filename)
        else:
            print("Okay; continuing.")
            return
    '''
    # Write DataFrame to Excel file
    table.to_excel(filename, sheet_name=sheet_name, index=False)
    print("Step 2: Settlements downloaded to excel sheet")


def download_returns(company_name):
    html = browser.page_source
    formatted_date = get_formatted_date()
    table = pd.read_html(html)[13]
    # today = date.today()
    # formatted_date = today.strftime("%m_%d_%Y")
    filename = company_name + ' Returns_' + formatted_date + '.xlsx'
    sheet_name = 'Sheet_' + formatted_date
    if os.path.exists(filename):
        os.remove(filename)
    ''' 
    # Check if file exists
    if os.path.exists(filename):
        # Ask for confirmation to overwrite
        overwrite = input("The file already exists. Do you want to overwrite it? (yes/no): ")
        if overwrite.lower() == "yes":
            # Delete the existing file
            os.remove(filename)
        else:
            print("Okay; continuing.")
            return
    '''
    # Write DataFrame to Excel file
    table.to_excel(filename, sheet_name=sheet_name, index=False)
    print("Step 3 :Returns downloaded to excel sheet")


url = "https://reports.actumprocessing.com/cgi-bin/makepage.cgi?page=tranrep"
with open('config.json') as f:
  config = json.load(f)

usernameNovus = config['usernameNovus']
passwordNovus = config['passwordNovus']
usernameRobin = config['usernameRobin']
passwordRobin = config['passwordRobin']

is_yesterday = True  # Global variable to track if we're processing yesterdays reports
browser.get(url)
login_to_actum(url, usernameRobin, passwordRobin)  # log in to robin
navigate_to_reports()  # go to reporting tab and select merchant reports
generate_reports('Settlements', )
download_settlements("Robin")  # downloads to .xlsx file
browser.maximize_window()  # take back control of browser
browser.back()
generate_reports('Returns', )
download_returns("Robin")
print("Step 4: Downloading tasks for Robin complete")
##################
browser.maximize_window()
time.sleep(0.5)
navigate_to_utilities()
time.sleep(2)
login_to_actum(url, usernameNovus, passwordNovus)  # log in to Novus
navigate_to_reports()  # go to reporting tab and select merchant reports
generate_reports('Settlements', )
download_settlements("Novus")  # downloads to .xlsx file
browser.maximize_window()  # take back control of browser
browser.back()
generate_reports('Returns')
download_returns("Novus")
print("Step 5: Downloading tasks for Novus complete")

import excel_processing

try:
    excel_processing.create_excel_sheets()
    excel_processing.process_excel_files()
    print("Step 6: Files have been processed ")
except Exception as e:
    print(e)
time.sleep(2)

import CombineFiles
CombineFiles.combine_excel_files()

import UploadToOrg
try:
    UploadToOrg.driver_code()
except Exception as e:
    print(e)

browser.quit()
exit(0)