from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
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
options.add_experimental_option("detach", True)
options.add_argument('--no-sandbox')
browser = webdriver.Chrome(options=options)


def main():
    BASE_URL = "https://brightspace.carleton.ca/d2l/home"
    GRADING_PAGE_URL = "https://brightspace.carleton.ca/d2l/lms/grades/admin/enter/grade_item_edit.d2l?objectId=551527&ou=131240"
    PATH_TO_FEEDBACK_SHEETS = '/Users/aguo/Dev/2022-2023/Winter/2401/1/test/'

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

    sid_search_bar = (
        By.XPATH,  "//*[contains(@placeholder,'Search For…')]")
    grade_input = (
        By.XPATH, "//*[starts-with(@title,'Grade for ')]")
    save_all_button = (By.ID, 'z_b')
    confirm_button = (
        By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[1]/button[1]')
    save_feedback_button = (
        By.XPATH, '/html/body/div[2]/div/div[3]/div/div/d2l-floating-buttons/button[2]')
    feedback_box = (By.CSS_SELECTOR, "[data-id='d2l-uid-2']")

    # getting all files in directory
    for (dirpath, dirnames, filenames) in os.walk(PATH_TO_FEEDBACK_SHEETS):
        files.extend(filenames)

    # opening every .xlsx file and updating feedback
    for f in files:
        if f.endswith('.xlsx'):
            file_path = PATH_TO_FEEDBACK_SHEETS + f

            # hack to cache excel so that formulas are evaulated
            excel_app = xlwings.App(visible=False)
            excel_book = excel_app.books.open(file_path)
            excel_book.save()
            excel_book.close()
            excel_app.quit()

            workbook = load_workbook(
                filename=file_path, data_only=True, read_only=True)
            sheet = workbook.active

            assignment = {"feedback": sheet['B7'].value, "max_grade": sheet['C5'].value,
                          "actual_grade": sheet['B5'].value, "sname": sheet['B2'].value, "sid": sheet['B3'].value}

            workbook.close()

            grade_percentage = assignment['actual_grade'] / assignment['max_grade'] * 100

            WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
                sid_search_bar)).send_keys(assignment["sid"])
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
                sid_search_bar)).send_keys(Keys.ENTER)
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
                grade_input)).clear()
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
                grade_input)).send_keys(grade_percentage)


    # WebDriverWait(browser,10).until(EC.element_to_be_clickable(login_button)).click().send_keys


if __name__ == "__main__":
    main()
