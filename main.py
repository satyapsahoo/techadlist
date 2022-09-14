from openpyxl import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

# Selenium driver to scrape the TechAD website
ser = Service("/Users/satyaprakashsahoo/Documents/Chrome Driver/chromedriver")
driver = webdriver.Chrome(service=ser)
driver.get("https://www.autonomous-driving-berlin.com/speaker")
# Build list of speakers - position -company
advisor_list_element = driver.find_elements(By.CSS_SELECTOR, ".member-info-headline h3 a")
advisor_list = [element.text for element in advisor_list_element]

position_list_element = driver.find_elements(By.CSS_SELECTOR, ".position a")
position_list = [element.text for element in position_list_element]

company_list_element = driver.find_elements(By.CSS_SELECTOR, ".company-info")
company_list = [element.get_attribute("title") for element in company_list_element]

rows = (())
for i in range(len(company_list)):
    rows = rows + ((advisor_list[i], position_list[i], company_list[i]),)

# Printing in excel
workbook = load_workbook('peoplelist.xlsx')
sheet = workbook.active
for row in rows:
    sheet.append(row)
workbook.save("peoplelist.xlsx")
driver.quit()
