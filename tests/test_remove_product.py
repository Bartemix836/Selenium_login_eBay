import sys
import os
import time
import pandas as pd
import pytest  # Importing pytest for tests
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../')))

from pages.base_page import BasePage

# Fixture to set up the browser
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

    # Launching the browser
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

# Record number (LP) from Excel
lp_value = 1  # Dynamically change the LP value

# Test using pytest
def test_user_login_checkout(setup_driver):
    # Get user data from the Excel sheet
    username, password = read_data_from_excel(excel_path, lp_value)

    driver = setup_driver
    driver.get("https://www.saucedemo.com/")

    # Waiting for elements on the page
    wait = WebDriverWait(driver, 10)

    # Initializing the BasePage object
    base_page = BasePage(driver)

    # STEP1 - Enter the username:
    login_input_xpath = '//*[@id="user-name"]'
    login_input = wait.until(EC.element_to_be_clickable((By.XPATH, login_input_xpath)))
    base_page.fill_text_field(By.XPATH, login_input_xpath, username)

    # STEP2 - Enter the password:
    passwd_input_xpath = '//*[@id="password"]'
    passwd_input = wait.until(EC.element_to_be_clickable((By.XPATH, passwd_input_xpath)))
    base_page.fill_text_field(By.XPATH, passwd_input_xpath, password)

    # STEP3 - Press the login button:
    login_btn_xpath = '//*[@id="login-button"]'
    wait.until(EC.element_to_be_clickable((By.XPATH, login_btn_xpath)))
    base_page.click_element(By.XPATH, login_btn_xpath)

    # STEP4 - Select the sorting criterion "Price (high to low)":
    list_criterions_xpath = '//*[@id="header_container"]/div[2]/div/span/select'
    base_page.select_dropdown_option_by_index(list_criterions_xpath, 3)

    # STEP5 - Add products to the cart:
    add_btn1_xpath = '//*[@id="add-to-cart-sauce-labs-fleece-jacket"]'
    add_product1 = wait.until(EC.element_to_be_clickable((By.XPATH, add_btn1_xpath)))
    base_page.click_element(By.XPATH, add_btn1_xpath)

    add_btn2_xpath = '//*[@id="add-to-cart-sauce-labs-backpack"]'
    add_product2 = wait.until(EC.element_to_be_clickable((By.XPATH, add_btn2_xpath)))
    base_page.click_element(By.XPATH, add_btn2_xpath)

    # STEP6 - Go to the cart:
    cart_xpath = '//*[@id="shopping_cart_container"]/a'
    go_cart = wait.until(EC.element_to_be_clickable((By.XPATH, cart_xpath)))
    base_page.click_element(By.XPATH, cart_xpath)

    # STEP6A - Remove the selected product from the cart:
    removebtn_xpath = '//*[@id="remove-sauce-labs-fleece-jacket"]'
    go_rembtn = wait.until(EC.element_to_be_clickable((By.XPATH, removebtn_xpath)))
    base_page.click_element(By.XPATH, removebtn_xpath)

    # STEP7 - Press the "Checkout" button:
    checkout_btn_xpath = '/html/body/div/div/div/div[2]/div/div[2]/button[2]'
    click_checkoutbtn = wait.until(EC.element_to_be_clickable((By.XPATH, checkout_btn_xpath)))
    base_page.click_element(By.XPATH, checkout_btn_xpath)

    # STEP8 - Fill in the First Name, Last Name, and ZIP fields
    firstname_xpath = '//*[@id="first-name"]'
    wait.until(EC.element_to_be_clickable((By.XPATH, firstname_xpath)))
    base_page.fill_text_field(By.XPATH, firstname_xpath, 'test_name01')

    lastname_xpath = '//*[@id="last-name"]'
    wait.until(EC.element_to_be_clickable((By.XPATH, lastname_xpath)))
    base_page.fill_text_field(By.XPATH, lastname_xpath, 'test_lastname01')

    zip_xpath = '//*[@id="postal-code"]'
    wait.until(EC.element_to_be_clickable((By.XPATH, zip_xpath)))
    base_page.fill_text_field(By.XPATH, zip_xpath, '50-431')

    # STEP9 - Press the "Continue" button:
    continue_btn_xpath = '//*[@id="continue"]'
    click_continuebtn = wait.until(EC.element_to_be_clickable((By.XPATH, continue_btn_xpath)))
    base_page.click_element(By.XPATH, continue_btn_xpath)

    # STEP10 - Press the "Finish" button:
    finish_btn_xpath = '//*[@id="finish"]'
    click_finishbtn = wait.until(EC.element_to_be_clickable((By.XPATH, finish_btn_xpath)))
    base_page.click_element(By.XPATH, finish_btn_xpath)

    # STEP11 - Return to the homepage:
    backhome_btn_xpath = '//*[@id="back-to-products"]'
    click_homebtn = wait.until(EC.element_to_be_clickable((By.XPATH, backhome_btn_xpath)))
    base_page.click_element(By.XPATH, backhome_btn_xpath)

    print('Result-Positive')
