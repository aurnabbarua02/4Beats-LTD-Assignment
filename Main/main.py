from openpyxl import load_workbook
from selenium import webdriver
import time
import datetime

file = load_workbook('Excel.xlsx')
sheetName = datetime.datetime.today().strftime('%A')
print('Today is', sheetName,'. So, sheet name \'', sheetName,'\' will be modified.')
sheet = file[sheetName]

driver=webdriver.Edge()
driver.maximize_window()
driver.get("https://www.google.com/")

j = sheet.min_row + 1
while j < sheet.max_row + 1:
    searchKey = driver.find_element('name', "q")
    searchKey.clear()
    searchKey.send_keys(sheet.cell(row=j, column=3).value)
    time.sleep(3)
    suggestions = driver.find_elements('css selector', 'li.sbct')
    longest_suggestion = max(suggestions, key=lambda suggestion: len(suggestion.text))
    sheet.cell(row=j, column=4).value = longest_suggestion.text
    sheet.cell(row=j, column=5).value = sheet.cell(row=j, column=3).value
    j += 1

file.save('Excel.xlsx')
print('Completed')