from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import datetime
import pyautogui
from selenium.webdriver.chrome.options import Options
import pygame
import json


def check_file_existence(file_path):
    if os.path.exists(file_path):
        return True
    else:
        print(f"Error: File '{file_path}' doesn't exist.")
        return False

def get_resolution():
  pygame.init()
  screen = pygame.display.set_mode((0, 0), pygame.FULLSCREEN)
  width, height = screen.get_size()
  pygame.quit()

  return (width, height)


def convert_xy(x, y, orig_res=(1920, 1080), new_res=(2560, 1600)):
    ow, oh = orig_res
    nw, nh = new_res

    x_scale = nw / ow
    y_scale = nh / oh

    new_x = x * x_scale
    new_y = y * y_scale

    return int(new_x), int(new_y)

def upload_files(browser):
    try:
        # Assume files are in the same directory as the script
        file_path = os.getcwd()

        # Get yesterday's date
        yesterday = datetime.date.today() - datetime.timedelta(days=1)
        formatted_date = yesterday.strftime("%m_%d_%Y")

        # Find the file with "CombinedData" and formatted date in the name
        filenames = [filename for filename in os.listdir(file_path) if f"CombinedData_{formatted_date}" in filename]

        if len(filenames) == 0:
            print(f"No files with 'CombinedData_{formatted_date}' in the name found.")
            return

        filename = os.path.join(file_path, filenames[0])  # Use the first found filename

        if os.path.exists(filename):
            print("Step 7: Beginning upload process.")

            # Un-hide the file input element so it can be interacted with
            browser.execute_script("""
                var element = document.getElementById('payment_upload_uploadData_file');
                element.style.position = 'static';
                element.style.height = '1px';
                element.style.width = '1px';
                element.style.opacity = 1;
            """)

            # Wait for the file input field to be visible
            file_input = WebDriverWait(browser, 20).until(
                EC.element_to_be_clickable((By.ID, 'payment_upload_uploadData_file')))
            # Send the file path to the file input field
            file_input.send_keys(filename)
            time.sleep(1)
            # Wait for the upload button to be clickable
            upload_submit = WebDriverWait(browser, 20).until(
                EC.element_to_be_clickable((By.ID, 'payment_upload_upload')))
            # Click the upload button
            upload_submit.click()
            time.sleep(3)
            # Check for invalid rows
            invalid_rows_count_element = WebDriverWait(browser, 20).until(
                EC.presence_of_element_located((By.ID, 'upload_data_type_invalidRowsCount')))
            invalid_rows_count = int(invalid_rows_count_element.text.strip())

            if invalid_rows_count > 0:
                # Export invalid rows file
                export_button = WebDriverWait(browser, 20).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.export-invalid-button')))
                export_button.click()
                print(f"{invalid_rows_count} invalid rows found. Exported to file.")
                # Click upload valid
                approve_button = WebDriverWait(browser, 20).until(
                    EC.element_to_be_clickable((By.ID, 'upload_data_type_approveValid1')))
                approve_button.click()
                time.sleep(0.5)

                width, height = get_resolution()
                if width == 1920 and height == 1080:
                    pyautogui.moveTo(1080, 450)
                    pyautogui.click()

                else:
                    print("Converting instructions to work for your resolution")
                    new_x, new_y = convert_xy(1080, 450, (1920, 1080), (width, height))
                    pyautogui.moveTo(new_x, new_y)
                    pyautogui.click()
                time.sleep(5)
            else:
                # Click upload valid
                approve_button = WebDriverWait(browser, 20).until(
                    EC.element_to_be_clickable((By.ID, 'upload_data_type_approveValid1')))
                approve_button.click()
            time.sleep(5)
        else:
            print(f"File {filename} doesn't exist.")
    except Exception as e:
        print(f"An error occurred: {e}")


def upload_files_to_crm(company_name, filename_prefix, username, password):
    try:
        # Get the current directory of the Python script
        script_directory = os.path.dirname(os.path.abspath(__file__))
        # Configure Chrome options
        chrome_options = Options()
        chrome_options.add_argument(f"--download.default_directory={script_directory}")
        # Initialize Chrome WebDriver with the options
        browser = webdriver.Chrome(options=chrome_options)
        # Go to orgmeter.com
        browser.get('https://app.orgmeter.com/login')
        browser.maximize_window()
        # Fill company name
        input_field = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.ID, 'login_company')))
        input_field.send_keys(company_name)
        # Fill username
        username_field = WebDriverWait(browser, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="login_username"]')))
        username_field.send_keys(username)
        # Fill password
        password_field = browser.find_element(By.ID, 'login_password')
        password_field.send_keys(password)
        # Click login
        login_submit = browser.find_element(By.ID, 'login_submit')
        login_submit.click()
        # Click terminate button
        time.sleep(3)
        terminate_button = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.NAME, 'session_id')))
        terminate_button.click()
        # Wait for page to load after login
        time.sleep(0.7)
        # Navigate to the upload page
        payments_toggle = WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Payments")))
        payments_toggle.click()
        time.sleep(0.5)
        add_new = browser.find_element(By.XPATH, '//a[contains(@data-title, "Add New")]')
        add_new.click()
        time.sleep(0.5)
        upload_toggle_tab = browser.find_element(By.XPATH, '//a[contains(@href, "#tab-upload")]')
        upload_toggle_tab.click()
        time.sleep(0.5)
        # locate the select element
        select_element = browser.find_element(By.ID, 'payment_upload_paymentType')
        # create a Select object
        select_object = Select(select_element)
        # select by visible text
        select_object.select_by_visible_text('Advance Payback')
        time.sleep(0.7)
        # Upload the files
        upload_files(browser)
        '''
        # Logout
        browser.maximize_window()
        dropdown_menu = WebDriverWait(browser, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//a[@data-title='Katerina Marciante']"))
        )
        # dropdown_menu = browser.find_element(By.XPATH, "//a[@data-title='Katerina Marciante']")
        dropdown_menu.click()
        print("dropdown clicked")
        # Wait for the logout button to be visible
        logout_button = WebDriverWait(browser, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//li[@class='last']//a[@data-title='Log out']"))
        )
        # Click the logout button
        logout_button.click()
        print("Step 8: Files successfully uploaded.")
        '''
        browser.quit()
        print("Finished uploading to CRM")
    except Exception as e:
        print(f"An error occurred: {e}")

def driver_code():
    username_orgmeter = config['usernameOrgmeter']
    password_orgmeter = config['passwordOrgmeter']
    upload_files_to_crm("Novus", "Processed", username_orgmeter, password_orgmeter)

# driver_code()