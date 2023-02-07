from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from chromedriver_py import binary_path
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
import os
import xlwings
from openpyxl import load_workbook
from dotenv import dotenv_values
config = dotenv_values(".env")

options = webdriver.ChromeOptions()
# options.add_argument('--headless')
options.add_argument('--no-sandbox')
browser = webdriver.Chrome(options=options)


def main():
    BASE_URL = "https://brightspace.carleton.ca/d2l/home"
    GRADING_PAGE_URL = "https://brightspace.carleton.ca/d2l/lms/grades/admin/enter/grade_item_edit.d2l?objectId=551527&ou=131240"
    PATH_TO_FEEDBACK_SHEETS = '/Users/aguo/Dev/2022-2023/Winter/2401/1/final/a1_andrew_guo_sheets/'

    files = []

    email_field = (By.ID, 'userNameInput')
    password_field = (By.ID, 'passwordInput')
    login_button = (By.ID, 'submitButton')
    browser.get(BASE_URL)
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
        email_field)).send_keys(config["USERNAME"])
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
        password_field)).send_keys(config["PASSWORD"])
    WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable(login_button)).click()
    browser.get(GRADING_PAGE_URL)

    # getting all files in directory
    for (dirpath, dirnames, filenames) in os.walk(PATH_TO_FEEDBACK_SHEETS):
        files.extend(filenames)

    # opening every .xlsx file and updating feedback
    for f in files:
        if f.endswith('.xlsx'):
            fpath = PATH_TO_FEEDBACK_SHEETS + f

            # hack to cache excel so that formulas are evaulated
            excel_app = xlwings.App(visible=False)
            excel_book = excel_app.books.open(fpath)
            excel_book.save()
            excel_book.close()
            excel_app.quit()

            workbook = load_workbook(
                filename=fpath, data_only=True, read_only=True)
            sheet = workbook.active

            assignment = {"feedback": sheet['B7'].value, "max_grade": sheet['C5'].value,
                          "actual_grade": sheet['B5'].value, "sname": sheet['B2'].value, "sid": sheet['B3'].value}

            workbook.close()


if __name__ == "__main__":
    main()
