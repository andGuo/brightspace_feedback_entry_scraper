import os
from dotenv import load_dotenv
load_dotenv() 
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
import chromedriver_binary  # Adds chromedriver binary to path

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
browser = webdriver.Chrome(options=options)