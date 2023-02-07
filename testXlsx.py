import os
import xlwings
from openpyxl import load_workbook

PATH_TO_FEEDBACK_SHEETS = '/Users/aguo/Dev/2022-2023/Winter/2401/1/fixed/'

files = []

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

        grade_percentage = assignment['actual_grade'] / \
            assignment['max_grade'] * 100
        
        print(f"{assignment['sname']} - {grade_percentage}%")
