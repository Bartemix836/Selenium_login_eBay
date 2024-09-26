import sys
import os
import time
import pandas as pd
import pytest
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
    # Path to the Edge WebDriver
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
    """Read user credentials from the Excel file based on the given LP value."""
    # Read data from Excel
    df = pd.read_excel(excel_path)

    # Find the row corresponding to the given LP value
    user_data = df[df['LP'] == lp_value]

    if not user_data.empty:
        username = user_data['username'].values[0]
        password = user_data['password'].values[0]
        return username, password
    else:
        raise ValueError(f"Record with LP = {lp_value} not found.")


def validate_login(driver):
    """Check if login was successful by looking for an error message."""
    try:
        error_msg_xpath = '//*[@data-test="error"]'
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, error_msg_xpath)))
        return False  # Login failed
    except:
        return True  # Login successful


@pytest.mark.parametrize("lp_value", [1, 2, 3, 4, 5, 6, 7, 8])  # LP range from Excel
def test_login_with_excel_data(setup_driver, lp_value):
    """Test login using data from Excel for multiple users."""
    # Path to the Excel file
    excel_path = 'C:/Users/barte/PycharmProjects/selenium_kurs/test_login_with_exceldata/data/users.xlsx'

    # Get user credentials from Excel
    username, password = read_data_from_excel(excel_path, lp_value)

    driver = setup_driver
    driver.get("https://www.saucedemo.com/")

    wait = WebDriverWait(driver, 10)

    # Initialize the BasePage object
    base_page = BasePage(driver)

    # STEP1 - Enter the username
    login_input_xpath = '//*[@id="user-name"]'
    login_input = wait.until(EC.element_to_be_clickable((By.XPATH, login_input_xpath)))
    base_page.fill_text_field(By.XPATH, login_input_xpath, username)

    # STEP2 - Enter the password
    passwd_input_xpath = '//*[@id="password"]'
    passwd_input = wait.until(EC.element_to_be_clickable((By.XPATH, passwd_input_xpath)))
    base_page.fill_text_field(By.XPATH, passwd_input_xpath, password)

    # STEP3 - Click the login button
    login_btn_xpath = '//*[@id="login-button"]'
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, login_btn_xpath)))
    base_page.click_element(By.XPATH, login_btn_xpath)

    # STEP4 - Verify the result of the login
    time.sleep(2)  # Small delay for the browser to react
    assert validate_login(driver), f"Failed login for LP = {lp_value} with username: {username}"

    # STEP5 - Return to the login page
    url_page = 'https://www.saucedemo.com/'
    base_page.load_page(url_page)
    time.sleep(4)

