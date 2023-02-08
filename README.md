# Brightspace Feedback Entry Scraper 🔍

### By: Andrew Guo

## Description

A selenium python program to input grades and feedback. Works on MacOS with Chrome.

### Dependencies
Most of these can be installed with pip:
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
2. In main.py set the constants accordingly.
3. Then run main.py

Note: 
- Excel will ask for you to grant permission to access each file. This has to be manually accepted each time afaik.
- GRADING_PAGE_URL is the one found at "Homepage"->"Progress"->"Grades"->"{Downwards_arrow_beside_assignment}"->"Enter Grades"
