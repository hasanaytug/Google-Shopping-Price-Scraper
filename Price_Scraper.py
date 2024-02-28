import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
import re

# Get the search term from the user
url = input("Enter your Niche: ")
file_name = url

# Start web driver
driver = webdriver.Chrome()

# Search the niche on google shopping
driver.get("https://www.google.com/search?tbm=shop&hl=en&psb=1&ved=2ahUKEwjc9ruT-syEAxWUXggEHdiFCFQQu-kFegQIABAN&q="+url)

# Wait for the full load
time.sleep(5)

# Find all of the strings starting with $
elements = driver.find_elements(By.XPATH, "//*[starts-with(text(),'$')]")
data = []  # For storing the numbers
total = 0
for element in elements:
    text = element.text.strip()
    if re.match(r'\$\s*([0-9,.]+)', text):
        # Convert strings to numbers
        number_str = re.sub(r'[^\d.]', '', text)
        if number_str.count('.') <= 1:
            number = float(number_str)
            data.append(number)
            total += number

# Get rid of the first 5 numbers
data_without_first_five = data[5:]

# Find the average
if data_without_first_five:
    average = sum(data_without_first_five) / len(data_without_first_five)
else:
    average = 0

# Write the data to Excel
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Scraped Data"
sheet['A1'] = "Scraped Numbers"

for i, number in enumerate(data_without_first_five, start=2):
    sheet[f'A{i}'] = number

sheet['E3'] = "Average"
sheet['F3'] = average

wb.save(f"{file_name}.xlsx")

# Stop the web driver
driver.quit()
