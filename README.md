# Brightspace Feedback Entry Scraper 🔍

### By: Andrew Guo

## Description

A selenium python program to input grades and feedback. Works on MacOS with Chrome.

### Dependencies
Most of these can be installed with pip (Python3):
 - Selenium
 - chromedriver-py
 - xlwings
 - openpyxl
 - dotenv
 - Excel (Native Installation)

### Start-Up
1. Make a .env file with the following variables filled in with your Brightspace login credentials:
```
USERNAME=""
PASSWORD=""
```
2. In main.py set the constants found at the top of main() accordingly.
3. Then run main.py

Note: 
- On MacOS, Excel will ask for you to grant permission to access each file. This has to be manually accepted each time afaik.
  - https://stackoverflow.com/questions/39604876/using-xlwings-to-open-an-excel-file-on-mac-os-x-el-capitan-requires-grant-access
- GRADING_PAGE_URL is the one found at "Homepage"->"Progress"->"Grades"->"{Downwards_arrow_beside_assignment}"->"Enter Grades"
