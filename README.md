# Brightspace Feedback Entry Scraper 

### By: Andrew Guo

## Description

A selenium program to input grades and feedback.

### Dependencies
Most of these can be installed with pip:
 - Selenium
 - chromedriver-py
 - xlwings
 - openpyxl
 - dotenv

### Start-Up
1. Make a .env file with the following variables filled in with your Brightspace login credentials:
```
USERNAME=""
PASSWORD=""
```
2. In main.py set the constants accordingly. Mainly PATH_TO_FEEDBACK_SHEETS and GRADING_PAGE_URL.

3. Then run main.py