import sys
import os
import time
import pandas as pd
import pytest  # Import pytest
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../')))

from pages.base_page import BasePage

@pytest.fixture(scope="module")
def setup_driver():
    # Path to the Edge driver
    driver_path = 'Add your path to driver'

    # Creating an instance of Service and setting browser options
    service = EdgeService(driver_path)
    options = EdgeOptions()

    # Edge browser options
    options.add_argument("--inprivate")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--remote-debugging-port=9222")

    # Launch the browser
    driver = webdriver.Edge(service=service, options=options)
    driver.maximize_window()

    yield driver
    driver.quit()

def read_data_from_excel(excel_path, lp_value):
    # Read data from Excel
    df = pd.read_excel(excel_path)

    # Find the row corresponding to the given LP
    user_data = df[df['LP'] == lp_value]

    if not user_data.empty:
        username = user_data['username'].values[0]
        password = user_data['password'].values[0]
        return username, password
    else:
        raise ValueError(f"Record with LP = {lp_value} not found.")


# Path to the Excel file
excel_path = 'Add path to users.xlsx file'

# Test for login and interaction with social media icons
def test_social_media_interactions(setup_driver):
    lp_value = 1
    username, password = read_data_from_excel(excel_path, lp_value)

    driver = setup_driver
    driver.get("https://www.saucedemo.com/")

    wait = WebDriverWait(driver, 10)

    # Initializing the BasePage object
    base_page = BasePage(driver)

    # STEP1 - Enter the username
    login_input_xpath = '//*[@id="user-name"]'
    login_input = wait.until(EC.element_to_be_clickable((By.XPATH, login_input_xpath)))
    base_page.fill_text_field(By.XPATH, login_input_xpath, username)

    # STEP2 - Enter the password
    passwd_input_xpath = '//*[@id="password"]'
    passwd_input = wait.until(EC.element_to_be_clickable((By.XPATH, passwd_input_xpath)))
    base_page.fill_text_field(By.XPATH, passwd_input_xpath, password)

    # STEP3 - Press the login button
    login_btn_xpath = '//*[@id="login-button"]'
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, login_btn_xpath)))
    base_page.click_element(By.XPATH, login_btn_xpath)

    # STEP4 - Click the Twitter icon
    twitter_link_xpath = '//a[@data-test="social-twitter"]'
    twitter_link = wait.until(EC.element_to_be_clickable((By.XPATH, twitter_link_xpath)))
    driver.execute_script("arguments[0].removeAttribute('target');", twitter_link)  # Open in the same tab
    base_page.click_element(By.XPATH, twitter_link_xpath)
    time.sleep(4)

    # STEP5 - Go back to the previous page
    base_page.load_page('https://www.saucedemo.com/inventory.html')

    # STEP6 - Click the Facebook icon
    facebook_link_xpath = '//a[@data-test="social-facebook"]'
    facebook_link = wait.until(EC.element_to_be_clickable((By.XPATH, facebook_link_xpath)))
    driver.execute_script("arguments[0].removeAttribute('target');", facebook_link)  # Open in the same tab
    base_page.click_element(By.XPATH, facebook_link_xpath)
    time.sleep(4)

    # STEP7 - Accept cookies on Facebook
    cookies_xpath = '//*[@id="facebook"]/body/div[3]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div/span/span'
    cookies_element = wait.until(EC.element_to_be_clickable((By.XPATH, cookies_xpath)))
    base_page.click_element(By.XPATH, cookies_xpath)
    time.sleep(4)

    # STEP8 - Go back to the previous page
    base_page.load_page('https://www.saucedemo.com/inventory.html')

    # STEP9 - Click the LinkedIn icon
    linkedin_link_xpath = '//a[@data-test="social-linkedin"]'
    linkedin_link = wait.until(EC.element_to_be_clickable((By.XPATH, linkedin_link_xpath)))
    driver.execute_script("arguments[0].removeAttribute('target');", linkedin_link)  # Open in the same tab
    base_page.click_element(By.XPATH, linkedin_link_xpath)
    time.sleep(4)

    # STEP10 - Go back to the previous page
    base_page.load_page('https://www.saucedemo.com/inventory.html')

    print('-----------------------------------------------------------------------------------')
    print('Result-Positive')
