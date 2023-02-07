import os
from dotenv import dotenv_values
config = dotenv_values(".env")
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from chromedriver_py import binary_path

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
options = webdriver.ChromeOptions()
#options.add_argument('--headless')
options.add_argument('--no-sandbox')
browser = webdriver.Chrome(options=options)

def main():
    base_url = "https://brightspace.carleton.ca/d2l/home"
    grading_page_url = "https://brightspace.carleton.ca/d2l/lms/grades/admin/enter/grade_item_edit.d2l?objectId=551527&ou=131240"

    browser.add_cookie({"name": "IDMSESSID", "value": config["IDMSESSID"]})
    browser.add_cookie({"name": "TS012103f9", "value": config["TS012103f9"]})
    browser.add_cookie({"name": "TS0186ecd1", "value": config["TS0186ecd1"]})
    browser.add_cookie({"name": "d2lSecureSessionVal", "value": config["d2lSecureSessionVal"]})
    browser.add_cookie({"name": "d2lSessionVal", "value": config["d2lSessionVal"]})

    email_field = (By.ID, 'userNameInput')
    password_field = (By.ID, 'passwordInput')
    login_button = (By.ID, 'submitButton')
    browser.get(base_url)
    WebDriverWait(browser,10).until(EC.element_to_be_clickable(email_field)).send_keys(config["USERNAME"])
    WebDriverWait(browser,10).until(EC.element_to_be_clickable(password_field)).send_keys(config["PASSWORD"])
    WebDriverWait(browser,10).until(EC.element_to_be_clickable(login_button)).click()
    browser.get(grading_page_url)


if __name__ == "__main__":
    main()
