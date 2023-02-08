from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from chromedriver_py import binary_path
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
import os
import time
import xlwings
from openpyxl import load_workbook
from dotenv import dotenv_values
config = dotenv_values(".env")

options = webdriver.ChromeOptions()
# options.add_argument('--headless') #- Headless doesn't work, perhap increasing the sleep timers would do it 
options.add_experimental_option("detach", True)
options.add_argument('--no-sandbox')
options.add_argument("--start-maximized")
browser = webdriver.Chrome(options=options)


def main():
    BASE_URL = "https://brightspace.carleton.ca/d2l/home"
    GRADING_PAGE_URL = "https://brightspace.carleton.ca/d2l/lms/grades/admin/enter/grade_item_edit.d2l?objectId=551527&ou=131240"
    PATH_TO_FEEDBACK_SHEETS = '/Users/aguo/Dev/2022-2023/Winter/2401/1/extensions/' #This needs to end with a slash

    files = []

    #These will probably break and need to updated before use
    email_field = (By.ID, 'userNameInput')
    password_field = (By.ID, 'passwordInput')
    login_button = (By.ID, 'submitButton')
    sid_search_bar = (
        By.XPATH,  "//*[contains(@placeholder,'Search For…')]")
    grade_input = (
        By.XPATH, "//*[starts-with(@title,'Grade for ')]")
    open_feedback_dialog = (By.XPATH, "//a[starts-with(@title,'Edit comments for')] | //a[starts-with(@title,'Enter comments for')]")
    feedback_iframe = (By.CLASS_NAME, 'd2l-dialog-frame')
    fullscreen_button = 'return document.querySelector("#publicComments").shadowRoot.querySelector("div.d2l-htmleditor-label-flex-container > div > div.d2l-htmleditor-flex-container > div.d2l-htmleditor-toolbar-container > d2l-htmleditor-toolbar-full").shadowRoot.querySelector("div.d2l-htmleditor-toolbar-container.d2l-htmleditor-toolbar-overflowing.d2l-htmleditor-toolbar-chomping.d2l-htmleditor-toolbar-measured > div.d2l-htmleditor-toolbar-pinned-actions > d2l-htmleditor-button-toggle:nth-child(2)").shadowRoot.querySelector("button")'
    save_feedback_button = (
        By.XPATH, '//*[@id="d_content"]/div[4]/div/button[1]')

    save_all_button = (By.ID, 'z_b')
    confirm_button = (
        By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[1]/button[1]')

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
            file_path = PATH_TO_FEEDBACK_SHEETS + f

            try:
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
            except:
                print(f'Invalid file_path:{file_path}')
                continue

            grade_percentage = assignment['actual_grade'] / \
                assignment['max_grade'] * 100
            
            print(f"{assignment['sname']} - {grade_percentage}%")

            try:
                #Search for student id
                search = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable(sid_search_bar))
                search.send_keys(Keys.COMMAND, 'a')
                search.send_keys(Keys.DELETE)
                search.send_keys(assignment['sid'])
                search.send_keys(Keys.ENTER)

                #Input Grade
                grade = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
                    grade_input))
                grade.send_keys(Keys.COMMAND, 'a')
                grade.send_keys(Keys.DELETE)
                grade.send_keys(grade_percentage)

                #Update Feedback
                edit = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
                    open_feedback_dialog)).click()

                feedback = WebDriverWait(browser, 10).until(
                    EC.frame_to_be_available_and_switch_to_it(feedback_iframe))

                time.sleep(1)
                fs_button = browser.execute_script(fullscreen_button)
                fs_button.click()

                feedback_text = browser.switch_to.active_element
                feedback_text.send_keys(Keys.COMMAND, 'a')
                feedback_text.send_keys(Keys.DELETE)
                feedback_text.send_keys(assignment['feedback'])

                fs_button.click()

                save_feedback = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
                    save_feedback_button)).click()
                
                time.sleep(1)
                #Save grade
                save_grade = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
                    save_all_button)).click()

                confirm = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
                     confirm_button)).click()
                
                print("Done.")
            except Exception as e:
                print(f"Failed to input feedback for student({assignment['sid']}): {assignment['sname']}")
                print(f"{file_path}")
                #print(e)
                continue


if __name__ == "__main__":
    main()
