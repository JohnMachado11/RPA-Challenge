from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

import pandas as pd


# Initialize WebDriver
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
driver.maximize_window()

# Load excel and convert to list
df = pd.read_excel("data/challenge.xlsx")
data_list = df.values.tolist()

# Open RPA website
rpa_website = driver.get("https://www.rpachallenge.com/")

# Locate and click the start button 
start_button = driver.find_element(By.CLASS_NAME, "btn-large.uiColorButton")
start_button.click()

for data in data_list:
    # Unpack row data into individual variables
    first_name, last_name, company, role, address, email, phone = data

    # Locate input fields on the web page
    first_name_form = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelFirstName"]')
    last_name_form = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelLastName"]')
    cmpny_name_form = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelCompanyName"]')
    role_in_cmpny_form = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelRole"]')
    address_form = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelAddress"]')
    email_form = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelEmail"]')
    phone_num_form = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelPhone"]')
    submit_button = driver.find_element(By.CLASS_NAME, 'btn.uiColorButton')

    # Fill out form fields with the corresponding data
    first_name_form.send_keys(first_name) 
    last_name_form.send_keys(last_name)
    cmpny_name_form.send_keys(company)
    role_in_cmpny_form.send_keys(role)
    address_form.send_keys(address)
    email_form.send_keys(email)
    phone_num_form.send_keys(phone)

    # Submit the form
    submit_button.click()


input("Press Enter to close the browser...")
driver.quit()